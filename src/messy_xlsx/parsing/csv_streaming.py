"""Bounded-prefix CSV inspection and native pandas-chunk Arrow streaming."""

from __future__ import annotations

import io
import warnings
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, replace
from pathlib import Path
from types import TracebackType
from typing import Any, BinaryIO, Self, cast

import pandas as pd
import pyarrow as pa
from pandas.io.parsers.base_parser import ParserBase
from pandas.io.parsers.readers import TextFileReader

from messy_xlsx._fallback_signals import _contains_process_failure, _exception_traceback
from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FormatError, StreamingTypeError
from messy_xlsx.normalization import compile_normalization_plan
from messy_xlsx.normalization.plan import (
    MAX_SAMPLE_BYTES,
    MAX_SAMPLE_CELLS,
    MAX_SAMPLE_VALUES,
)
from messy_xlsx.parsing.base_handler import ParseOptions
from messy_xlsx.parsing.contracts import ParseMetrics, StreamingBatchReader
from messy_xlsx.parsing.csv_handler import DEFAULT_NA_VALUES, CSVHandler
from messy_xlsx.parsing.csv_io import (
    FooterTrimmingReader,
    LogicalRecordBudgetReader,
    MalformedRecordFilteringReader,
    RecordLimitingReader,
)
from messy_xlsx.parsing.materialized_streaming import (
    PreparedStreamingReader,
    _canonical_materialized_value,
    _CloseOnceReader,
    _public_dataframe_display_names,
    normalization_sample_from_dataframe,
    wrap_normalized_streaming_reader,
)
from messy_xlsx.parsing.parse_plan import ParsePlan
from messy_xlsx.parsing.physical_values import (
    UnsupportedPhysicalValueError,
    encode_physical_value,
    physical_label_description,
    physical_value_description,
)
from messy_xlsx.parsing.streams import _close_if_present, _run_cleanups
from messy_xlsx.parsing.xlsx_streaming import (
    _enter_prearmed_context,
    _record_batch_with_row_count,
    _RetryableSourceContext,
)


@contextmanager
def _open_bounded_sample_target(
    source: SourceHandle,
    inspection: CSVInspection,
    max_records: int,
    ignored_prefix_records: int,
    footer_lookahead: int,
) -> Iterator[io.RawIOBase]:
    """Open one sample borrow and expose a non-owning record-budget proxy."""

    def guarded(stream: BinaryIO) -> io.RawIOBase:
        limited = RecordLimitingReader(
            stream,
            inspection.encoding,
            inspection.delimiter,
            source.description,
            max_records,
        )
        budgeted = LogicalRecordBudgetReader(
            limited,
            inspection.encoding,
            inspection.delimiter,
            source.description,
            enforce_total=True,
            ignored_prefix_records=ignored_prefix_records,
        )
        if footer_lookahead == 0:
            return budgeted
        return FooterTrimmingReader(
            budgeted,
            inspection.encoding,
            inspection.delimiter,
            source.description,
            footer_lookahead,
            ignored_prefix_records,
            max_record_bytes=MAX_SAMPLE_BYTES,
            max_pending_bytes=MAX_SAMPLE_BYTES,
        )

    with source.open_backend() as target:
        if isinstance(target, Path):
            with target.open("rb") as stream:
                yield guarded(stream)
            return
        yield guarded(target)


@dataclass(frozen=True, slots=True)
class CSVInspection:
    """Immutable dialect/header evidence derived from one bounded prefix."""

    encoding: str
    delimiter: str
    skip_rows: int
    encoding_errors: str = "strict"


def inspect_csv_source(source: SourceHandle, options: ParseOptions) -> CSVInspection:
    """Inspect a path or borrowed binary source without retaining complete input."""
    handler = CSVHandler()
    path = source.path
    if path is not None:
        prefix = handler._read_path_prefix(path)
    else:
        with source.open_backend() as backend:
            if isinstance(backend, Path):
                raise ValueError("Stream-backed CSV source unexpectedly produced a path")
            prefix = handler._read_stream_prefix(backend)
    encoding = handler._detect_encoding_from_bytes(prefix, options.encoding)
    delimiter = handler._detect_delimiter_from_bytes(prefix, encoding)
    encoding_errors = "ignore" if source.is_stream else "strict"
    if path is not None:
        skip_rows = handler._resolved_skip_rows(
            prefix,
            encoding,
            delimiter,
            options,
            target=path,
            file_desc=source.description,
            is_stream=False,
        )
    else:
        with source.open_backend() as backend:
            if isinstance(backend, Path):
                raise ValueError("Stream-backed CSV source unexpectedly produced a path")
            skip_rows = handler._resolved_skip_rows(
                prefix,
                encoding,
                delimiter,
                options,
                target=backend,
                file_desc=source.description,
                is_stream=True,
            )
    return CSVInspection(
        encoding=encoding,
        delimiter=delimiter,
        skip_rows=skip_rows,
        encoding_errors=encoding_errors,
    )


def prepare_csv_streaming_reader(  # noqa: C901
    source: SourceHandle,
    plan: ParsePlan,
    metrics: ParseMetrics,
    *,
    construction_owner: Any | None = None,
) -> PreparedStreamingReader:
    """Compile bounded CSV evidence and return an inert native row reader."""
    options = plan.to_parse_options()
    try:
        inspection = inspect_csv_source(source, options)
        try:
            sample_frame = _read_bounded_sample(source, options, inspection)
        except BaseException as error:
            if _contains_process_failure(error):
                raise
            if source.path is None or not _caused_by_unicode_decode_error(error):
                raise
            inspection = _latin1_fallback_inspection(source, options)
            sample_frame = _read_bounded_sample(source, options, inspection)
    except BaseException:
        metrics.failed_attempts += 1
        raise
    metrics.sample_reads += 1
    sample = normalization_sample_from_dataframe(
        sample_frame,
        date_system="1900",
        preserve_native=not plan.normalize,
    )
    if not plan.normalize and len(sample_frame):
        columns = list(sample.columns)
        changed = False
        for ordinal, column in enumerate(columns):
            if not pa.types.is_null(column.type):
                continue
            inferred = pa.array(sample_frame.iloc[:, ordinal], from_pandas=True)
            if pa.types.is_null(inferred.type):
                continue
            columns[ordinal] = inferred
            changed = True
        if changed:
            sample = replace(
                sample,
                schema=pa.schema(
                    [pa.field(str(ordinal), column.type) for ordinal, column in enumerate(columns)]
                ),
                columns=tuple(columns),
            )
    normalization_plan = compile_normalization_plan(sample, plan)
    display_names = _public_dataframe_display_names(
        tuple(sample_frame.columns),
        normalization_plan.final_display_names,
        plan,
    )
    raw_schema = pa.schema(
        [pa.field(str(ordinal), pa.string()) for ordinal in range(len(sample_frame.columns))]
    )
    if construction_owner is None:
        raw_reader: StreamingBatchReader = CSVStreamingReader.prepare(
            source,
            options,
            inspection,
            tuple(sample_frame.columns),
            raw_schema,
            normalization_plan.input_schema,
            cast(int, plan.batch_size),
            metrics,
        )
        return wrap_normalized_streaming_reader(
            raw_reader,
            normalization_plan,
            normalize=plan.normalize,
            pandas_display_names=display_names,
        )

    owned_reader = _CloseOnceReader()
    construction_owner.attach(owned_reader)
    raw_reader = CSVStreamingReader.prepare(
        source,
        options,
        inspection,
        tuple(sample_frame.columns),
        raw_schema,
        normalization_plan.input_schema,
        cast(int, plan.batch_size),
        metrics,
    )
    owned_reader.attach(raw_reader)
    prepared = wrap_normalized_streaming_reader(
        owned_reader,
        normalization_plan,
        normalize=plan.normalize,
        pandas_display_names=display_names,
        rollback_on_error=False,
    )
    construction_owner.replace(owned_reader, prepared.reader)
    return prepared


def _read_bounded_sample(
    source: SourceHandle,
    options: ParseOptions,
    inspection: CSVInspection,
) -> pd.DataFrame:
    """Read only bounded rows needed to compile the stable stream schema."""
    _validate_row_limits(options, source.description)
    # Probe one logical record first so the retained sample respects the cell
    # budget even for very wide CSV files.
    probe_limit = 1
    if options.header_rows > 1:
        probe_limit += options.skip_rows + options.header_rows
    probe_footer = _sample_footer_lookahead(options, probe_limit)
    try:
        with (
            _open_bounded_sample_target(
                source,
                inspection,
                _sample_record_limit(
                    options,
                    inspection,
                    probe_limit,
                    probe_footer,
                ),
                _sample_ignored_prefix_records(options, inspection),
                probe_footer,
            ) as target,
            warnings.catch_warnings(),
        ):
            warnings.simplefilter("ignore", pd.errors.ParserWarning)
            probe = pd.read_csv(
                target,
                **_materialized_read_kwargs(
                    options,
                    inspection,
                    nrows=probe_limit,
                    skip_footer=0,
                ),
            )
    except BaseException as error:
        _raise_csv_error(error, source.description)
    width = len(probe.columns)
    if width > MAX_SAMPLE_CELLS:
        raise ValueError(f"sample raw window may retain at most {MAX_SAMPLE_CELLS} cells")
    sample_rows = min(
        MAX_SAMPLE_VALUES,
        MAX_SAMPLE_CELLS // max(1, width),
    )
    if options.max_rows is not None:
        sample_rows = min(sample_rows, cast(int, options.max_rows))
    footer_candidate = _sample_footer_candidate(options)
    if footer_candidate:
        sample_rows = min(sample_rows, MAX_SAMPLE_VALUES - footer_candidate)
    template = CSVHandler()._finish_frame(probe, options).iloc[:0]
    del probe

    if sample_rows == 0:
        return template.reset_index(drop=True)
    read_limit = sample_rows
    if options.header_rows > 1:
        read_limit += options.skip_rows + options.header_rows
    footer_lookahead = _sample_footer_lookahead(options, read_limit)
    frame = _read_sample_frame(
        source,
        options,
        inspection,
        read_limit,
        footer_lookahead,
    )
    if frame.empty and footer_lookahead:
        frame = _read_sample_frame(
            source,
            options,
            inspection,
            read_limit,
            0,
        )
    return frame.iloc[:sample_rows].reset_index(drop=True)


def _read_sample_frame(
    source: SourceHandle,
    options: ParseOptions,
    inspection: CSVInspection,
    read_limit: int,
    footer_lookahead: int,
) -> pd.DataFrame:
    try:
        with (
            _open_bounded_sample_target(
                source,
                inspection,
                _sample_record_limit(
                    options,
                    inspection,
                    read_limit,
                    footer_lookahead,
                ),
                _sample_ignored_prefix_records(options, inspection),
                footer_lookahead,
            ) as target,
            warnings.catch_warnings(),
        ):
            warnings.simplefilter("ignore", pd.errors.ParserWarning)
            frame = pd.read_csv(
                target,
                **_materialized_read_kwargs(
                    options,
                    inspection,
                    nrows=read_limit,
                    skip_footer=0,
                ),
            )
    except BaseException as error:
        _raise_csv_error(error, source.description)
    return CSVHandler()._finish_frame(frame, options)


def _sample_record_limit(
    options: ParseOptions,
    inspection: CSVInspection,
    data_rows: int,
    footer_lookahead: int,
) -> int:
    """Translate pandas data-row limits into bounded physical record limits."""
    header_records = 1 if options.header_rows > 0 else 0
    return max(
        1,
        cast(int, inspection.skip_rows) + header_records + data_rows + footer_lookahead,
    )


def _sample_footer_candidate(options: ParseOptions) -> int:
    """Admit only conservative footer lookahead within the row budget."""
    footer = cast(int, options.skip_footer or 0)
    if footer <= 0 or footer > MAX_SAMPLE_VALUES // 2:
        return 0
    return footer


def _sample_footer_lookahead(options: ParseOptions, data_rows: int) -> int:
    footer = _sample_footer_candidate(options)
    if data_rows + footer > MAX_SAMPLE_VALUES:
        return 0
    return footer


def _caused_by_unicode_decode_error(error: BaseException) -> bool:
    """Return whether a bounded evidence failure wraps strict decoding."""
    seen: set[int] = set()
    current: BaseException | None = error
    while current is not None and id(current) not in seen:
        if isinstance(current, UnicodeDecodeError):
            return True
        seen.add(id(current))
        current = current.__cause__ or current.__context__
    return False


def _latin1_fallback_inspection(
    source: SourceHandle,
    options: ParseOptions,
) -> CSVInspection:
    """Rebuild bounded path evidence using the legacy first fallback."""
    path = source.path
    assert path is not None
    handler = CSVHandler()
    prefix = handler._read_path_prefix(path)
    encoding = "latin-1"
    return CSVInspection(
        encoding=encoding,
        delimiter=handler._detect_delimiter_from_bytes(prefix, encoding),
        skip_rows=cast(int, options.skip_rows),
        encoding_errors="strict",
    )


def _sample_ignored_prefix_records(
    options: ParseOptions,
    inspection: CSVInspection,
) -> int:
    protected = cast(int, inspection.skip_rows)
    if options.header_rows > 1:
        return protected + 1 + cast(int, options.header_rows)
    return protected + (1 if options.header_rows > 0 else 0)


def _materialized_read_kwargs(
    options: ParseOptions,
    inspection: CSVInspection,
    *,
    nrows: int | None,
    skip_footer: int,
) -> dict[str, object]:
    return {
        "encoding": inspection.encoding,
        "encoding_errors": inspection.encoding_errors,
        "delimiter": inspection.delimiter,
        "skiprows": inspection.skip_rows if options.header_rows <= 1 else 0,
        "skipfooter": skip_footer,
        "nrows": nrows,
        "na_values": options.na_values or DEFAULT_NA_VALUES,
        "header": 0 if options.header_rows > 0 else None,
        "engine": "python" if skip_footer else "c",
        "on_bad_lines": "warn",
    }


def _stream_read_kwargs(
    options: ParseOptions,
    inspection: CSVInspection,
    column_names: tuple[object, ...],
    batch_size: int,
) -> dict[str, object]:
    kwargs = _materialized_read_kwargs(
        options,
        inspection,
        nrows=options.max_rows,
        skip_footer=0,
    )
    # Small C-parser chunks make malformed-row decisions before enough
    # neighboring rows are available.  Use one stable bounded compatibility
    # window for small public batches, then rechunk at our API boundary.
    kwargs["chunksize"] = (
        max(MAX_SAMPLE_VALUES, batch_size) if not options.skip_footer else batch_size
    )
    kwargs["dtype"] = object
    # pandas' C parser changes malformed-row decisions with small chunks. The
    # Python parser honors the complete-record field count independently of
    # batch_size while still yielding bounded chunks.
    kwargs["engine"] = "python" if options.skip_footer else "c"
    if options.header_rows > 1:
        kwargs["skiprows"] = range(
            1,
            1 + options.skip_rows + options.header_rows,
        )
        kwargs["names"] = list(column_names)
    elif options.header_rows == 0:
        kwargs["names"] = list(column_names)
    return kwargs


def _validate_row_limits(options: ParseOptions, description: str) -> None:
    """Preserve pandas' established rejection of footer trimming with nrows."""
    if options.max_rows is None or not options.skip_footer:
        return
    raise FormatError(
        "Cannot parse CSV file: 'skipfooter' not supported with 'nrows'",
        file_path=description,
        detected_format="csv",
    )


def _initialize_text_file_reader(
    reader: TextFileReader,
    target: object,
    kwargs: dict[str, object],
) -> None:
    """Initialize one already-owned pandas reader without a factory return gap."""
    refined = dict(kwargs)
    on_bad_lines = refined.get("on_bad_lines")
    if isinstance(on_bad_lines, str):
        refined["on_bad_lines"] = ParserBase.BadLineHandleMethod[on_bad_lines.upper()]
    reader.__init__(target, **refined)


def _initialize_file_io(file: io.FileIO, path: Path) -> None:
    """Initialize one already-owned path file without a factory return gap."""
    io.FileIO.__init__(file, path, "rb")


class _PreownedPandasReader:
    """Own a pandas reader before its resource-acquiring initializer begins."""

    def __init__(self) -> None:
        self._reader: TextFileReader | None = TextFileReader.__new__(TextFileReader)
        self._input_owner: io.IOBase | None = None

    def initialize(
        self,
        target: object,
        kwargs: dict[str, object],
        *,
        input_owner: io.IOBase | None = None,
    ) -> None:
        reader = self._reader
        if reader is None:
            raise RuntimeError("pandas CSV reader owner is closed")
        if input_owner is not None:
            self.attach_input_owner(input_owner)
        _initialize_text_file_reader(reader, target, kwargs)

    def attach_input_owner(self, input_owner: io.IOBase) -> None:
        if self._input_owner is not None and self._input_owner is not input_owner:
            raise RuntimeError("pandas input owner is already attached")
        self._input_owner = input_owner

    def __iter__(self) -> _PreownedPandasReader:
        return self

    def __next__(self) -> pd.DataFrame:
        reader = self._reader
        if reader is None:
            raise StopIteration
        return next(reader)

    def close(self) -> None:
        cleanups: list[tuple[str, Any]] = []
        if self._reader is not None:
            cleanups.append(("pandas text reader cleanup", self._close_reader))
        if self._input_owner is not None:
            cleanups.append(("pandas input proxy cleanup", self._close_input_owner))
        _run_cleanups(cleanups)

    def _close_reader(self) -> None:
        reader = self._reader
        if reader is None:
            return
        try:
            if hasattr(reader, "_engine") and hasattr(reader, "handles"):
                reader.close()
            else:
                handles = getattr(reader, "handles", None)
                if handles is not None:
                    handles.close()
                engine = getattr(reader, "_engine", None)
                if engine is not None:
                    engine.close()
        except BaseException as error:
            if not _contains_process_failure(error):
                self._reader = None
            raise
        self._reader = None

    def _close_input_owner(self) -> None:
        input_owner = self._input_owner
        if input_owner is None:
            return
        try:
            input_owner.close()
        except BaseException as error:
            if not _contains_process_failure(error):
                self._input_owner = None
            raise
        self._input_owner = None


class CSVStreamingReader:
    """Lazily yield schema-stable encoded Arrow batches from pandas chunks."""

    def __init__(
        self,
        source: SourceHandle,
        options: ParseOptions,
        inspection: CSVInspection,
        column_names: tuple[object, ...],
        schema: pa.Schema,
        physical_schema: pa.Schema,
        batch_size: int,
        metrics: ParseMetrics | None = None,
    ) -> None:
        self._initialize(
            source,
            options,
            inspection,
            column_names,
            schema,
            physical_schema,
            batch_size,
            metrics,
        )

    @classmethod
    def prepare(
        cls,
        source: SourceHandle,
        options: ParseOptions,
        inspection: CSVInspection,
        column_names: tuple[object, ...],
        schema: pa.Schema,
        physical_schema: pa.Schema,
        batch_size: int,
        metrics: ParseMetrics | None = None,
    ) -> Self:
        """Return an inert closeable reader before the full row pass starts."""
        reader = cls.__new__(cls)
        reader._initialize(
            source,
            options,
            inspection,
            column_names,
            schema,
            physical_schema,
            batch_size,
            metrics,
        )
        return reader

    def _initialize(
        self,
        source: SourceHandle,
        options: ParseOptions,
        inspection: CSVInspection,
        column_names: tuple[object, ...],
        schema: pa.Schema,
        physical_schema: pa.Schema,
        batch_size: int,
        metrics: ParseMetrics | None,
    ) -> None:
        if not isinstance(source, SourceHandle):
            raise TypeError("source must be a SourceHandle")
        if batch_size < 1:
            raise ValueError("batch_size must be >= 1")
        if len(column_names) != len(schema):
            raise ValueError("CSV sampled columns do not match the compiled schema")
        if len(column_names) != len(physical_schema):
            raise ValueError("CSV sampled columns do not match the physical schema")
        self._source = source
        self._options = options
        self._inspection = inspection
        self._column_names = column_names
        self._schema = schema
        self._physical_schema = physical_schema
        self._batch_size = batch_size
        self._metrics = metrics
        self._source_context: _RetryableSourceContext | None = _RetryableSourceContext(source)
        self._chunks: Any | None = None
        self._pending: pd.DataFrame | None = None
        self._row_offset = 0
        self._started = False
        self._pass_attempted = False
        self._terminal = False
        self._closed = False
        self._exhausted = False
        self._metrics_recorded = False

    @property
    def schema(self) -> pa.Schema:
        return self._schema

    @property
    def pending_row_count(self) -> int:
        """Expose bounded footer state for deterministic architecture tests."""
        return 0 if self._pending is None else len(self._pending)

    def read_next_batch(self) -> pa.RecordBatch | None:
        if self._closed or self._terminal:
            return None
        try:
            if not self._started:
                self._start_row_pass()
            frame = self._next_output_frame()
            if frame is None:
                self._terminal = True
                self._exhausted = True
                return None
            return self._encode_frame(frame)
        except BaseException as error:
            self._terminal = True
            self._record_failed_pass()
            self._close_resources(
                primary_error=error,
                primary_traceback=_exception_traceback(error),
            )
            raise

    def _start_row_pass(self) -> None:
        self._pass_attempted = True
        _validate_row_limits(self._options, self._source.description)
        source_context = self._source_context
        assert source_context is not None
        context = self._source.open_backend()
        target = _enter_prearmed_context(source_context, context)
        chunks = _PreownedPandasReader()
        self._chunks = chunks
        path_file: io.FileIO | None = None
        if isinstance(target, Path):
            path_file = io.FileIO.__new__(io.FileIO)
            pipeline_target: Any = path_file
            pipeline_owns_target = True
        else:
            pipeline_target = target
            pipeline_owns_target = False
        if not self._options.skip_footer:
            reader_target = pipeline_target
            if path_file is not None:
                chunks.attach_input_owner(path_file)
        else:
            protected = self._inspection.skip_rows
            if self._options.header_rows > 1:
                protected += 1 + self._options.header_rows
            elif self._options.header_rows > 0:
                protected += 1
            pipeline_target = FooterTrimmingReader(
                pipeline_target,
                self._inspection.encoding,
                self._inspection.delimiter,
                self._source.description,
                cast(int, self._options.skip_footer),
                protected,
                owns_stream=pipeline_owns_target,
            )
            pipeline_owns_target = True
            input_owner = MalformedRecordFilteringReader(
                pipeline_target,
                self._inspection.encoding,
                self._inspection.encoding_errors,
                self._inspection.delimiter,
                self._source.description,
                len(self._column_names),
                protected,
                owns_stream=pipeline_owns_target,
            )
            reader_target = input_owner
            chunks.attach_input_owner(input_owner)
        try:
            if path_file is not None:
                _initialize_file_io(path_file, target)
            chunks.initialize(
                reader_target,
                _stream_read_kwargs(
                    self._options,
                    self._inspection,
                    self._column_names,
                    self._batch_size,
                ),
            )
        except BaseException as error:
            _raise_csv_error(error, self._source.description)
        self._started = True

    def _next_output_frame(self) -> pd.DataFrame | None:
        while True:
            if self._pending is not None:
                pending = self._pending
                if len(pending) > self._batch_size:
                    self._pending = pending.iloc[self._batch_size :]
                    return pending.iloc[: self._batch_size]
                self._pending = None
                if len(pending):
                    return pending
            chunk = self._read_chunk()
            if chunk is None:
                self._pending = None
                return None
            self._pending = chunk

    def _read_chunk(self) -> pd.DataFrame | None:
        chunks = self._chunks
        assert chunks is not None
        try:
            chunk = next(chunks)
        except StopIteration:
            return None
        except BaseException as error:
            _raise_csv_error(error, self._source.description)
        if len(chunk.columns) != len(self._column_names):
            raise FormatError(
                "Cannot parse CSV file: column count changed during streaming",
                file_path=self._source.description,
                detected_format="csv",
            )
        chunk.columns = list(self._column_names)
        return chunk

    def _encode_frame(self, frame: pd.DataFrame) -> pa.RecordBatch:
        arrays: list[pa.Array] = []
        for ordinal, target in enumerate(self._physical_schema.types):
            encoded: list[str | None] = []
            for relative, value in enumerate(frame.iloc[:, ordinal].tolist()):
                canonical = _coerce_sampled_value(value, target)
                try:
                    encoded.append(encode_physical_value(canonical))
                except UnsupportedPhysicalValueError:
                    raise StreamingTypeError(
                        "streamed value is incompatible with the fixed schema",
                        ordinal=ordinal,
                        display_label=physical_label_description(self._column_names[ordinal]),
                        row_offset=self._row_offset + relative,
                        value_description=physical_value_description(canonical),
                        expected_type="supported Arrow scalar",
                    ) from None
            arrays.append(pa.array(encoded, type=pa.string()))
        self._row_offset += len(frame)
        return _record_batch_with_row_count(arrays, self._schema, len(frame))

    def close(self) -> None:
        if self._closed and self._chunks is None and self._source_context is None:
            return
        self._closed = True
        self._terminal = True
        self._pending = None
        self._close_resources()

    def _close_resources(
        self,
        *,
        primary_error: BaseException | None = None,
        primary_traceback: TracebackType | None = None,
    ) -> None:
        cleanups: list[tuple[str, Any]] = []
        if self._chunks is not None:
            cleanups.append(("pandas CSV chunk reader cleanup", self._close_chunks))
        if self._source_context is not None:
            cleanups.append(("CSV source borrow cleanup", self._close_source_context))
        try:
            _run_cleanups(
                cleanups,
                primary_error=primary_error,
                primary_traceback=primary_traceback,
            )
        except BaseException:
            self._record_failed_pass()
            raise
        if primary_error is None and self._exhausted:
            self._record_successful_pass()

    def _close_chunks(self) -> None:
        chunks = self._chunks
        if chunks is None:
            return
        try:
            self._chunks = _close_if_present(chunks)
        except BaseException as error:
            if not _contains_process_failure(error):
                self._chunks = None
            raise

    def _close_source_context(self) -> None:
        source_context = self._source_context
        if source_context is None:
            return
        try:
            self._source_context = _close_if_present(source_context)
        except BaseException as error:
            if not _contains_process_failure(error):
                self._source_context = None
            raise

    def _record_successful_pass(self) -> None:
        if self._metrics_recorded or not self._started:
            return
        self._metrics_recorded = True
        if self._metrics is not None:
            self._metrics.streaming_passes += 1

    def _record_failed_pass(self) -> None:
        if self._metrics_recorded or not self._pass_attempted:
            return
        self._metrics_recorded = True
        if self._metrics is not None:
            self._metrics.failed_attempts += 1


def _raise_csv_error(error: BaseException, description: str) -> None:
    if _contains_process_failure(error):
        raise error
    if not isinstance(error, Exception):
        raise error
    if isinstance(error, FormatError):
        raise error
    raise FormatError(
        f"Cannot parse CSV file: {error}",
        file_path=description,
        detected_format="csv",
    ) from error


def _coerce_sampled_value(value: object, target: pa.DataType) -> object | None:
    """Coerce only lexically compatible CSV values to the sampled physical type."""
    canonical = cast(object | None, _canonical_materialized_value(value))
    if canonical is None:
        return None
    if pa.types.is_string(target) or pa.types.is_large_string(target):
        return str(canonical)
    if pa.types.is_integer(target):
        return _coerce_sampled_integer(canonical)
    if pa.types.is_floating(target):
        return _coerce_sampled_float(canonical)
    if pa.types.is_boolean(target):
        return _coerce_sampled_boolean(canonical)
    return canonical


def _coerce_sampled_integer(value: object) -> object:
    try:
        numeric = pd.to_numeric(value, errors="raise")
    except (TypeError, ValueError):
        return value
    if isinstance(numeric, bool):
        return value
    try:
        integer = int(numeric)
    except (OverflowError, TypeError, ValueError):
        return value
    return integer if numeric == integer else value


def _coerce_sampled_float(value: object) -> object:
    try:
        return float(pd.to_numeric(value, errors="raise"))
    except (OverflowError, TypeError, ValueError):
        return value


def _coerce_sampled_boolean(value: object) -> object:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.casefold()
        if lowered == "true":
            return True
        if lowered == "false":
            return False
    return value


__all__ = [
    "CSVInspection",
    "CSVStreamingReader",
    "inspect_csv_source",
    "prepare_csv_streaming_reader",
]
