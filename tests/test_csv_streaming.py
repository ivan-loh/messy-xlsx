"""Native CSV/TSV bounded-input and streaming contracts."""

from __future__ import annotations

import io
import sys
import warnings
from pathlib import Path
from typing import Any

import pandas as pd
import pyarrow as pa
import pytest
from pandas.io.parsers.readers import TextFileReader

import messy_xlsx.parsing.csv_io as csv_io_module
import messy_xlsx.parsing.csv_streaming as csv_streaming_module
import messy_xlsx.workbook as workbook_module
from messy_xlsx import MessyWorkbook, SheetConfig, StreamingTypeError
from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FormatError
from messy_xlsx.normalization.plan import MAX_SAMPLE_BYTES
from messy_xlsx.parsing import CSVHandler, ParseOptions
from messy_xlsx.parsing.contracts import ParseMetrics
from messy_xlsx.parsing.csv_io import (
    FooterTrimmingReader,
    LogicalRecordBudgetReader,
    MalformedRecordFilteringReader,
    RecordLimitingReader,
)
from messy_xlsx.parsing.csv_streaming import (
    CSVInspection,
    CSVStreamingReader,
    _read_bounded_sample,
)
from messy_xlsx.parsing.handler_registry import HandlerRegistry
from messy_xlsx.warnings import LegacyAPIWarning


@pytest.fixture(autouse=True)
def _enable_native_csv_prototype(monkeypatch: pytest.MonkeyPatch) -> None:
    """Keep the Task 14 source-checkout prototype under characterization."""
    from messy_xlsx.parsing import csv_native

    monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", True)


def _expected_native_metrics(**counters: int) -> ParseMetrics:
    """Build an explicit expected snapshot for one native CSV operation."""
    from messy_xlsx.parsing.csv_contracts import (
        CSVExecutionDecision,
        CSVExecutionKind,
        CSVExecutionReason,
    )

    decision = CSVExecutionDecision(
        operation_id=1,
        kind=CSVExecutionKind.NATIVE,
        reason=CSVExecutionReason.NATIVE_SELECTED,
    )
    return ParseMetrics(
        **counters,
        csv_operation_sequence=1,
        last_csv_execution=decision,
        csv_execution_counts={(decision.kind, decision.reason): 1},
    )


class NoUnboundedReadBytesIO(io.BytesIO):
    """Seekable caller stream that rejects accidental complete reads."""

    def __init__(self, content: bytes) -> None:
        super().__init__(content)
        self.read_sizes: list[int] = []

    def read(self, size: int = -1) -> bytes:
        self.read_sizes.append(size)
        if size is None or size < 0:
            raise AssertionError("CSV parsing must never request read(-1)")
        return super().read(size)


class NonSeekableBytes:
    """Caller-owned upload stream consumed once by SourceHandle's replay spool."""

    def __init__(self, content: bytes, name: str = "data.csv") -> None:
        self._content = content
        self._offset = 0
        self.name = name
        self.closed = False
        self.read_sizes: list[int] = []

    def seekable(self) -> bool:
        return False

    def read(self, size: int = -1) -> bytes:
        self.read_sizes.append(size)
        if size is None or size < 0:
            raise AssertionError("caller stream must never receive read(-1)")
        start = self._offset
        stop = min(len(self._content), start + size)
        self._offset = stop
        return self._content[start:stop]


def _table(stream: Any) -> pa.Table:
    return pa.Table.from_batches(list(stream), schema=stream.schema)


def test_materialized_seekable_csv_is_borrowed_directly_without_unbounded_read() -> None:
    source = NoUnboundedReadBytesIO(b"name,value\nalice,1\nbob,2\n")
    source.seek(7)

    frame = CSVHandler().parse(
        source,
        "Sheet1",
        ParseOptions(auto_detect_header=False),
    )

    assert frame.to_dict("list") == {
        "name": ["alice", "bob"],
        "value": [1, 2],
    }
    assert source.tell() == 7
    assert source.closed is False
    assert source.read_sizes
    assert all(size >= 0 for size in source.read_sizes)


def test_materialized_seekable_csv_passes_binary_stream_directly_to_pandas(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"name,value\nalice,1\n")
    source.seek(4)
    observed: dict[str, object] = {}

    def fake_read_csv(target: object, **kwargs: object) -> pd.DataFrame:
        observed["target"] = target
        observed["encoding"] = kwargs.get("encoding")
        return pd.DataFrame({"name": ["alice"], "value": [1]})

    monkeypatch.setattr(pd, "read_csv", fake_read_csv)

    frame = CSVHandler().parse(
        source,
        "Sheet1",
        ParseOptions(auto_detect_header=False),
    )

    assert frame.to_dict("list") == {"name": ["alice"], "value": [1]}
    assert observed == {"target": source, "encoding": "utf-8"}
    assert source.tell() == 4


def test_builtin_csv_batch_route_is_native_and_never_materializes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"name,value\nalice,1\nbob,2\n")

    def fail_materialization(*_args: object, **_kwargs: object) -> pd.DataFrame:
        raise AssertionError("built-in CSV streaming must not materialize a DataFrame")

    monkeypatch.setattr(
        workbook_module.MessyWorkbook,
        "_materialize_raw_frame",
        fail_materialization,
    )

    with (
        MessyWorkbook(source, filename="data.csv") as workbook,
        workbook.iter_batches(
            batch_size=1,
            config=SheetConfig(auto_detect=False, sanitize_column_names=False),
        ) as stream,
    ):
        table = _table(stream)
        metrics = workbook.parse_metrics

    assert table.to_pydict() == {"name": ["alice", "bob"], "value": [1, 2]}
    assert metrics.full_materializations == 0
    assert metrics.streaming_passes == 1


def test_native_csv_footer_buffer_respects_post_footer_batch_bound(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = tmp_path / "footer.csv"
    path.write_text("value\n1\n2\n3\n4\n5\n", encoding="utf-8")

    def fail_materialization(*_args: object, **_kwargs: object) -> pd.DataFrame:
        raise AssertionError("skip_footer must use a bounded streaming tail")

    monkeypatch.setattr(
        workbook_module.MessyWorkbook,
        "_materialize_raw_frame",
        fail_materialization,
    )
    config = SheetConfig(
        auto_detect=False,
        header_rows=1,
        skip_footer=2,
        normalize=False,
        sanitize_column_names=False,
    )

    with (
        MessyWorkbook(path) as workbook,
        workbook.iter_batches(batch_size=2, config=config) as stream,
    ):
        schema = stream.schema
        batches = list(stream)

    assert schema.names == ["value"]
    assert batches
    assert all(batch.schema == schema for batch in batches)
    assert all(0 < batch.num_rows <= 2 for batch in batches)
    assert pa.Table.from_batches(batches, schema=schema).to_pydict() == {"value": [1, 2, 3]}


@pytest.mark.parametrize("consume", ["early_close", "exhaust"])
def test_native_csv_stream_restores_exact_caller_cursor(
    consume: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = NoUnboundedReadBytesIO(b"value\n1\n2\n3\n")
    source.seek(5)

    def fail_materialization(*_args: object, **_kwargs: object) -> pd.DataFrame:
        raise AssertionError("built-in CSV streaming must not materialize")

    monkeypatch.setattr(
        workbook_module.MessyWorkbook,
        "_materialize_raw_frame",
        fail_materialization,
    )
    workbook = MessyWorkbook(source, filename="data.csv")
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(
            auto_detect=False,
            normalize=False,
            sanitize_column_names=False,
        ),
    )
    try:
        if consume == "early_close":
            assert next(stream).num_rows == 1
            stream.close()
        else:
            assert sum(batch.num_rows for batch in stream) == 3
        assert source.tell() == 5
        assert source.closed is False
        assert source.read_sizes
        assert all(size >= 0 for size in source.read_sizes)
    finally:
        stream.close()
        workbook.close()


@pytest.mark.parametrize("source_kind", ["path", "seekable", "nonseekable"])
def test_native_csv_source_kinds_match_materialized_values(
    source_kind: str,
    tmp_path: Path,
) -> None:
    content = b"name,value\nalice,1\nbob,2\n"
    path = tmp_path / "parity.csv"
    path.write_bytes(content)
    source: object
    if source_kind == "path":
        source = path
    elif source_kind == "seekable":
        source = NoUnboundedReadBytesIO(content)
    else:
        source = NonSeekableBytes(content)
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )

    with MessyWorkbook(source, filename="parity.csv") as workbook:  # type: ignore[arg-type]
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", LegacyAPIWarning)
            expected = workbook.to_dataframe(config=config)
        with workbook.iter_batches(batch_size=1, config=config) as stream:
            actual = _table(stream).to_pandas()

    pd.testing.assert_frame_equal(actual, expected)
    if source_kind != "path":
        assert source.closed is False  # type: ignore[union-attr]
        assert all(size >= 0 for size in source.read_sizes)  # type: ignore[union-attr]


@pytest.mark.parametrize(
    ("skip_footer", "expected"),
    [
        (0, [1, 2, 3, 4, 5]),
        (1, [1, 2, 3, 4]),
        (2, [1, 2, 3]),
        (5, []),
        (8, []),
    ],
)
def test_footer_matrix_handles_chunk_boundaries_and_footer_at_least_rows(
    skip_footer: int,
    expected: list[int],
) -> None:
    source = io.BytesIO(b"value\n1\n2\n3\n4\n5\n")
    config = SheetConfig(
        auto_detect=False,
        header_rows=1,
        skip_footer=skip_footer,
        normalize=False,
        sanitize_column_names=False,
    )

    with (
        MessyWorkbook(source, filename="footer.csv") as workbook,
        workbook.iter_batches(batch_size=2, config=config) as stream,
    ):
        table = _table(stream)

    assert table.column_names == ["value"]
    assert table.column("value").to_pylist() == expected


def test_bounded_footer_sampling_excludes_text_footer_from_numeric_schema() -> None:
    source = io.BytesIO(b"value\n1\n2\n3\nEND OF REPORT\n")
    config = SheetConfig(
        auto_detect=False,
        skip_footer=1,
        normalize=False,
        sanitize_column_names=False,
    )

    with (
        MessyWorkbook(source, filename="footer.csv") as workbook,
        workbook.iter_batches(batch_size=2, config=config) as stream,
    ):
        table = _table(stream)

    assert table.schema.types == [pa.int64()]
    assert table.column(0).to_pylist() == [1, 2, 3]


@pytest.mark.parametrize(
    ("content", "config", "columns", "values"),
    [
        (
            b"first,second\nA,B\nunit,unit\nalice,1\nbob,2\n",
            SheetConfig(
                auto_detect=False,
                header_rows=2,
                normalize=False,
                sanitize_column_names=False,
            ),
            ["A__unit", "B__unit"],
            [["alice", "bob"], ["1", "2"]],
        ),
        (
            b"name,value\nalice,1\nbob,2\n",
            SheetConfig(
                auto_detect=False,
                header_rows=0,
                normalize=False,
                sanitize_column_names=False,
            ),
            ["col_0", "col_1"],
            [["name", "alice", "bob"], ["value", "1", "2"]],
        ),
    ],
    ids=["multi-row-header", "no-header"],
)
def test_native_header_modes_preserve_legacy_column_contract(
    content: bytes,
    config: SheetConfig,
    columns: list[str],
    values: list[list[object]],
) -> None:
    with (
        MessyWorkbook(io.BytesIO(content), filename="headers.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        table = _table(stream)

    assert table.column_names == columns
    assert [column.to_pylist() for column in table.columns] == values


def test_metadata_header_detection_uses_only_the_bounded_prefix() -> None:
    source = NoUnboundedReadBytesIO(
        b"Report: Sales,\nGenerated: 2024-01-01,\nName,Amount\nAlice,10\nBob,20\n"
    )
    config = SheetConfig(
        auto_detect=True,
        normalize=False,
        sanitize_column_names=False,
    )

    with (
        MessyWorkbook(source, filename="report.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        table = _table(stream)

    assert table.to_pydict() == {"Name": ["Alice", "Bob"], "Amount": [10, 20]}
    assert all(size >= 0 for size in source.read_sizes)


@pytest.mark.parametrize(
    ("filename", "content"),
    [
        ("semicolon.csv", b'name;note\nalice;"one;two"\n'),
        ("tabs.tsv", b"name\tvalue\nalice\t1\n"),
        ("bom.csv", b"\xef\xbb\xbfname,value\nalice,1\n"),
        ("latin.csv", "name,note\nalice,café\n".encode("latin-1")),
        ("quoted.csv", b'name,note\nalice,"line one\nline two"\n'),
        ("malformed.csv", b"a,b\n1,2\n3\n4,5,6\n"),
    ],
)
def test_native_dialect_encoding_quoting_and_malformed_rows_match_materialized(
    filename: str,
    content: bytes,
) -> None:
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with MessyWorkbook(io.BytesIO(content), filename=filename) as workbook:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", LegacyAPIWarning)
            expected = workbook.to_dataframe(config=config)
        with workbook.iter_batches(batch_size=1, config=config) as stream:
            actual = _table(stream).to_pandas()

    pd.testing.assert_frame_equal(actual, expected)


def test_materialized_csv_max_rows_remains_enforced_without_a_complete_copy() -> None:
    source = NoUnboundedReadBytesIO(b"value\n1\n2\n3\n4\n")
    frame = CSVHandler().parse(
        source,
        None,
        ParseOptions(max_rows=2, auto_detect_header=False),
    )

    assert frame.to_dict("list") == {"value": [1, 2]}
    assert all(size >= 0 for size in source.read_sizes)


def test_bounded_sample_preserves_legacy_max_rows_footer_error() -> None:
    options = ParseOptions(
        max_rows=4,
        skip_footer=1,
        auto_detect_header=False,
    )
    with pytest.raises(FormatError, match="'skipfooter' not supported with 'nrows'"):
        _read_bounded_sample(
            workbook_module.SourceHandle(
                io.BytesIO(b"value\n1\n2\n3\nEND\nignored\n"),
                filename="max-footer.csv",
            ),
            options,
            CSVInspection("utf-8", ",", 0),
        )


def test_csv_reader_preserves_legacy_max_rows_footer_error() -> None:
    content = b"value\n1\n2\n3\nEND\nignored\n"
    source = SourceHandle(io.BytesIO(content), filename="max-footer.csv")
    reader = CSVStreamingReader.prepare(
        source,
        ParseOptions(
            auto_detect_header=False,
            max_rows=4,
            skip_footer=1,
        ),
        CSVInspection("utf-8", ",", 0),
        ("value",),
        pa.schema([pa.field("0", pa.string())]),
        pa.schema([pa.field("0", pa.string())]),
        2,
    )
    try:
        with pytest.raises(
            FormatError,
            match="'skipfooter' not supported with 'nrows'",
        ):
            reader.read_next_batch()
    finally:
        reader.close()
        source.close()


def test_bounded_sample_rejects_width_beyond_cell_budget(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(csv_streaming_module, "MAX_SAMPLE_CELLS", 2)
    source = workbook_module.SourceHandle(
        io.BytesIO(b"a,b,c\n1,2,3\n"),
        filename="too-wide.csv",
    )
    try:
        with pytest.raises(
            ValueError,
            match="sample raw window may retain at most 2 cells",
        ):
            _read_bounded_sample(
                source,
                ParseOptions(auto_detect_header=False),
                CSVInspection("utf-8", ",", 0),
            )
    finally:
        source.close()


def test_csv_metrics_count_one_sample_and_one_lazy_streaming_pass() -> None:
    source = io.BytesIO(b"value\n1\n2\n")
    with MessyWorkbook(source, filename="metrics.csv") as workbook:
        stream = workbook.iter_batches(
            batch_size=1,
            config=SheetConfig(auto_detect=False),
        )
        assert workbook.parse_metrics == _expected_native_metrics(sample_reads=1)
        assert next(stream).num_rows == 1
        assert workbook.parse_metrics == _expected_native_metrics(sample_reads=1)
        assert sum(batch.num_rows for batch in stream) == 1
        assert workbook.parse_metrics == _expected_native_metrics(
            sample_reads=1,
            streaming_passes=1,
        )


def test_csv_early_close_is_not_counted_as_success_or_failure() -> None:
    with MessyWorkbook(
        io.BytesIO(b"value\n1\n2\n"),
        filename="early-close.csv",
    ) as workbook:
        stream = workbook.iter_batches(
            batch_size=1,
            config=SheetConfig(auto_detect=False),
        )
        assert next(stream).num_rows == 1
        stream.close()

        assert workbook.parse_metrics == _expected_native_metrics(sample_reads=1)


def test_native_all_null_column_retains_sampled_pandas_physical_type() -> None:
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(
            io.BytesIO(b"value,blank\n1,\n2,\n"),
            filename="all-null.csv",
        ) as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        table = _table(stream)

    assert table.schema.types == [pa.int64(), pa.float64()]
    assert table.to_pydict() == {"value": [1, 2], "blank": [None, None]}


def test_footer_pending_state_never_exceeds_declared_footer() -> None:
    reader = FooterTrimmingReader(
        io.BytesIO(b"value\n1\n2\n3\n4\n5\n6\n"),
        "utf-8",
        ",",
        "pending.csv",
        skip_footer=3,
        protected_prefix_records=1,
    )
    try:
        output = bytearray()
        while chunk := reader.read(2):
            output.extend(chunk)
            assert reader.pending_record_count <= 3
        assert bytes(output) == b"value\n1\n2\n3\n"
    finally:
        reader.close()


def test_custom_csv_registry_remains_materialized_authority(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class CustomRegistry(HandlerRegistry):
        parse_calls = 0

        def parse(self, *_args: object, **_kwargs: object) -> pd.DataFrame:
            self.parse_calls += 1
            return pd.DataFrame({"source": ["custom"]})

    registry = CustomRegistry()

    def fail_native(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("custom CSV registries must not use native streaming")

    monkeypatch.setattr(workbook_module, "prepare_csv_streaming_reader", fail_native)
    with (
        MessyWorkbook(
            io.BytesIO(b"value\nbuiltin\n"),
            filename="custom.csv",
            registry=registry,
        ) as workbook,
        workbook.iter_batches(batch_size=1) as stream,
    ):
        table = _table(stream)

    assert registry.parse_calls == 1
    assert table.to_pydict() == {"source": ["custom"]}


def test_late_csv_value_raises_contextual_streaming_type_error_and_restores_cursor() -> None:
    rows = ["value", *(str(value) for value in range(2_100)), "not-an-integer"]
    source = io.BytesIO(("\n".join(rows) + "\n").encode())
    source.seek(11)
    config = SheetConfig(
        auto_detect=False,
        normalize=True,
        sanitize_column_names=False,
    )
    workbook = MessyWorkbook(source, filename="late.csv")
    stream = workbook.iter_batches(batch_size=500, config=config)

    with pytest.raises(StreamingTypeError) as captured:
        list(stream)

    assert captured.value.context["ordinal"] == 0
    assert captured.value.context["row_offset"] == 2_100
    assert captured.value.context["expected_type"] == "int64"
    assert source.tell() == 11
    assert source.closed is False
    workbook.close()


@pytest.mark.parametrize(
    ("failure", "expected_type"),
    [
        (ValueError("row failure"), FormatError),
        (MemoryError("process failure"), MemoryError),
    ],
)
def test_csv_iteration_failure_preserves_error_class_and_restores_cursor(
    failure: BaseException,
    expected_type: type[BaseException],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"value\n1\n2\n")
    source.seek(4)
    original_next = csv_streaming_module._PreownedPandasReader.__next__
    original_close_reader = csv_streaming_module._PreownedPandasReader._close_reader
    next_calls = 0
    close_calls = 0

    def fail_next(reader: object) -> pd.DataFrame:
        nonlocal next_calls
        next_calls += 1
        if next_calls == 1:
            return original_next(reader)
        raise failure

    def track_close(reader: object) -> None:
        nonlocal close_calls
        close_calls += 1
        original_close_reader(reader)

    monkeypatch.setattr(
        csv_streaming_module._PreownedPandasReader,
        "__next__",
        fail_next,
    )
    monkeypatch.setattr(
        csv_streaming_module._PreownedPandasReader,
        "_close_reader",
        track_close,
    )
    workbook = MessyWorkbook(source, filename="failure.csv")
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    )
    assert next(stream).num_rows == 1
    assert next(stream).num_rows == 1

    with pytest.raises(expected_type, match=str(failure)):
        next(stream)

    assert close_calls == 1
    assert workbook.parse_metrics == _expected_native_metrics(
        sample_reads=1,
        failed_attempts=1,
    )
    assert source.tell() == 4
    assert source.closed is False
    stream.close()
    workbook.close()


def test_csv_public_return_interruption_releases_operation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"value\n1\n")
    source.seek(3)
    workbook = MessyWorkbook(source, filename="return-gap.csv")
    target_code = workbook.iter_batches.__func__.__code__
    interrupted = False

    def interrupt(frame: Any, event: str, _arg: object) -> Any:
        nonlocal interrupted
        if frame.f_code is target_code and event == "return" and not interrupted:
            interrupted = True
            sys.settrace(None)
            frame.f_trace = None
            raise MemoryError("CSV stream return interrupted")
        return interrupt

    sys.settrace(interrupt)
    try:
        with pytest.raises(MemoryError, match="CSV stream return interrupted"):
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(auto_detect=False),
            )
    finally:
        sys.settrace(None)

    assert workbook._active_operation_token is None
    assert workbook._active_stream is None
    assert source.tell() == 3
    with workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    ) as stream:
        assert _table(stream).num_rows == 1
    workbook.close()


def test_csv_prepare_return_interruption_closes_preowned_reader() -> None:
    source = io.BytesIO(b"value\n1\n")
    source.seek(2)
    workbook = MessyWorkbook(source, filename="prepare-gap.csv")
    target_code = csv_streaming_module.prepare_csv_streaming_reader.__code__
    interrupted = False

    def interrupt(frame: Any, event: str, _arg: object) -> Any:
        nonlocal interrupted
        if frame.f_code is target_code and event == "return" and not interrupted:
            interrupted = True
            sys.settrace(None)
            frame.f_trace = None
            raise MemoryError("CSV prepare return interrupted")
        return interrupt

    sys.settrace(interrupt)
    try:
        with pytest.raises(MemoryError, match="CSV prepare return interrupted"):
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(auto_detect=False),
            )
    finally:
        sys.settrace(None)

    assert workbook._active_operation_token is None
    assert workbook._active_stream is None
    assert source.tell() == 2
    assert workbook.parse_metrics == _expected_native_metrics(sample_reads=1)
    with workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    ) as stream:
        assert _table(stream).num_rows == 1
    workbook.close()


def test_schema_sampling_suppresses_duplicate_malformed_row_warnings() -> None:
    source = io.BytesIO(b"a,b\n1,x\n2,y,z\n3,q\n")
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        with (
            MessyWorkbook(source, filename="malformed.csv") as workbook,
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(
                    auto_detect=False,
                    normalize=False,
                    sanitize_column_names=False,
                ),
            ) as stream,
        ):
            table = _table(stream)

    parser_warnings = [
        warning for warning in captured if issubclass(warning.category, pd.errors.ParserWarning)
    ]
    assert len(parser_warnings) == 1
    assert table.to_pydict() == {"a": [1, 3], "b": ["x", "q"]}


@pytest.mark.parametrize(
    ("source_kind", "late_value"),
    [
        ("seekable", "late"),
        ("nonseekable", "late"),
    ],
)
def test_late_invalid_utf8_preserves_legacy_source_specific_decoding(
    source_kind: str,
    late_value: str,
    tmp_path: Path,
) -> None:
    prefix_rows = b"".join(f"row-{index},ascii\n".encode() for index in range(5_000))
    content = b"name,note\n" + prefix_rows + b"last,\x93late\n"
    path = tmp_path / "late-encoding.csv"
    path.write_bytes(content)

    def make_source() -> object:
        if source_kind == "path":
            return path
        if source_kind == "seekable":
            return NoUnboundedReadBytesIO(content)
        return NonSeekableBytes(content)

    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(
            make_source(),  # type: ignore[arg-type]
            filename="late-encoding.csv",
        ) as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", LegacyAPIWarning)
        expected = workbook.to_dataframe(config=config)
    with (
        MessyWorkbook(
            make_source(),  # type: ignore[arg-type]
            filename="late-encoding.csv",
        ) as workbook,
        workbook.iter_batches(batch_size=127, config=config) as stream,
    ):
        actual = _table(stream).to_pandas()

    pd.testing.assert_frame_equal(actual, expected)
    assert actual.iloc[-1].tolist() == ["last", late_value]


def test_materialized_csv_process_failure_propagates_and_restores_cursor(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"value\n1\n")
    source.seek(3)

    def fail_read(*_args: object, **_kwargs: object) -> pd.DataFrame:
        raise MemoryError("materialized CSV process failure")

    monkeypatch.setattr(pd, "read_csv", fail_read)

    with pytest.raises(MemoryError, match="materialized CSV process failure"):
        CSVHandler().parse(
            source,
            None,
            ParseOptions(auto_detect_header=False),
        )

    assert source.tell() == 3
    assert source.closed is False


def test_large_footer_sampling_does_not_scan_to_eof_before_return(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = pd.read_csv
    calls: list[tuple[object, object]] = []

    def observed_read_csv(target: object, **kwargs: object) -> object:
        calls.append((kwargs.get("nrows"), kwargs.get("chunksize")))
        return original(target, **kwargs)

    monkeypatch.setattr(pd, "read_csv", observed_read_csv)
    content = b"value\n" + (b"1234567890\n" * 1_000_000)
    source = NoUnboundedReadBytesIO(content)
    with MessyWorkbook(source, filename="large-footer.csv") as workbook:
        stream = workbook.iter_batches(
            batch_size=2,
            config=SheetConfig(
                auto_detect=False,
                skip_footer=100_000,
                normalize=False,
                sanitize_column_names=False,
            ),
        )
        try:
            assert stream.schema.types == [pa.int64()]
            assert sum(source.read_sizes) < len(content)
        finally:
            stream.close()

    bounded_nrows = [nrows for nrows, _chunksize in calls if isinstance(nrows, int)]
    assert bounded_nrows
    assert max(bounded_nrows) <= csv_streaming_module.MAX_SAMPLE_VALUES + 1
    assert all(chunksize is None for _nrows, chunksize in calls)


def test_zero_retained_footer_rows_keep_bounded_inferred_schema() -> None:
    source = io.BytesIO(b"value\n1\n2\n")
    with (
        MessyWorkbook(source, filename="zero-retained.csv") as workbook,
        workbook.iter_batches(
            batch_size=1,
            config=SheetConfig(
                auto_detect=False,
                skip_footer=2,
                normalize=False,
                sanitize_column_names=False,
            ),
        ) as stream,
    ):
        table = _table(stream)

    assert table.schema.types == [pa.int64()]
    assert table.num_rows == 0


def test_csv_sample_failure_is_counted_once_and_restores_cursor(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"value\n1\n")
    source.seek(3)

    def fail_sample(*_args: object, **_kwargs: object) -> pd.DataFrame:
        raise ValueError("CSV sample failure")

    monkeypatch.setattr(pd, "read_csv", fail_sample)
    workbook = MessyWorkbook(source, filename="sample-failure.csv")
    try:
        with pytest.raises(FormatError, match="CSV sample failure"):
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(auto_detect=False),
            )
        assert workbook.parse_metrics == _expected_native_metrics(failed_attempts=1)
        assert source.tell() == 3
    finally:
        workbook.close()


def test_csv_reader_start_failure_is_counted_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail_start(*_args: object, **_kwargs: object) -> None:
        raise ValueError("CSV reader start failure")

    monkeypatch.setattr(
        csv_streaming_module,
        "_initialize_text_file_reader",
        fail_start,
    )
    workbook = MessyWorkbook(io.BytesIO(b"value\n1\n"), filename="start-failure.csv")
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    )
    try:
        with pytest.raises(FormatError, match="CSV reader start failure"):
            next(stream)
        assert workbook.parse_metrics == _expected_native_metrics(
            sample_reads=1,
            failed_attempts=1,
        )
    finally:
        stream.close()
        workbook.close()


def test_csv_reader_close_failure_is_counted_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_close = csv_streaming_module._PreownedPandasReader._close_reader
    close_calls = 0

    def fail_close(reader: object) -> None:
        nonlocal close_calls
        close_calls += 1
        if close_calls == 1:
            raise ValueError("CSV reader close failure")
        original_close(reader)

    monkeypatch.setattr(
        csv_streaming_module._PreownedPandasReader,
        "_close_reader",
        fail_close,
    )
    workbook = MessyWorkbook(io.BytesIO(b"value\n1\n"), filename="close-failure.csv")
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    )
    assert next(stream).num_rows == 1

    with pytest.raises(ValueError, match="CSV reader close failure"):
        stream.close()

    assert workbook.parse_metrics == _expected_native_metrics(
        sample_reads=1,
        failed_attempts=1,
    )
    workbook.close()


def test_csv_cursor_restore_failure_is_counted_once() -> None:
    class RestoreFailureBytesIO(io.BytesIO):
        fail_restore = False
        restore_position = 0

        def seek(self, position: int, whence: int = 0) -> int:
            if self.fail_restore and whence == 0 and position == self.restore_position:
                raise OSError("CSV cursor restore failure")
            return super().seek(position, whence)

    source = RestoreFailureBytesIO(b"value\n1\n2\n")
    source.seek(4)
    source.restore_position = 4
    workbook = MessyWorkbook(source, filename="restore-failure.csv")
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    )
    assert next(stream).num_rows == 1
    source.fail_restore = True

    with pytest.raises(OSError, match="CSV cursor restore failure"):
        stream.close()

    assert workbook.parse_metrics == _expected_native_metrics(
        sample_reads=1,
        failed_attempts=1,
    )
    source.fail_restore = False
    stream.close()
    workbook.close()
    assert source.tell() == 4


def test_nonseekable_csv_acquires_one_replay_spool_for_every_pass() -> None:
    source = NonSeekableBytes(b"value\n1\n2\n")
    workbook = MessyWorkbook(source, filename="spooled.csv")
    acquisition_reads = tuple(source.read_sizes)
    assert workbook._source_handle.was_snapshotted is True

    try:
        with workbook.iter_batches(
            batch_size=1,
            config=SheetConfig(auto_detect=False),
        ) as stream:
            assert _table(stream).column(0).to_pylist() == [1, 2]
        assert tuple(source.read_sizes) == acquisition_reads
        assert source.closed is False
    finally:
        workbook.close()


def test_csv_dataframe_chunks_preserve_values_index_and_cursor() -> None:
    source = NoUnboundedReadBytesIO(b"name,value\na,1\nb,2\nc,3\n")
    source.seek(6)
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(source, filename="chunks.csv") as workbook,
        workbook.iter_dataframe_chunks(batch_size=2, config=config) as stream,
    ):
        chunks = list(stream)

    actual = pd.concat(chunks)
    expected = pd.DataFrame({"name": ["a", "b", "c"], "value": [1, 2, 3]}).convert_dtypes(
        dtype_backend="pyarrow"
    )
    expected.columns = pd.Index(["name", "value"], dtype=object)
    pd.testing.assert_frame_equal(actual, expected)
    assert [chunk.index.tolist() for chunk in chunks] == [[0, 1], [2]]
    assert source.tell() == 6
    assert source.closed is False


def test_csv_full_reader_init_return_gap_closes_reader_borrow_and_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = io.BytesIO(b"value\n1\n2\n")
    source.seek(4)
    workbook = MessyWorkbook(source, filename="full-init-gap.csv")
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    )
    target_code = TextFileReader.__init__.__code__
    real_close = TextFileReader.close
    interrupted_reader: TextFileReader | None = None
    closed_reader_ids: list[int] = []

    def track_close(reader: TextFileReader) -> None:
        closed_reader_ids.append(id(reader))
        real_close(reader)

    def interrupt(frame: Any, event: str, _arg: object) -> Any:
        nonlocal interrupted_reader
        if frame.f_code is target_code and event == "return" and interrupted_reader is None:
            interrupted_reader = frame.f_locals["self"]
            sys.settrace(None)
            frame.f_trace = None
            raise MemoryError("full CSV reader init return interrupted")
        return interrupt

    monkeypatch.setattr(TextFileReader, "close", track_close)
    sys.settrace(interrupt)
    try:
        with pytest.raises(MemoryError, match="full CSV reader init return interrupted"):
            next(stream)
    finally:
        sys.settrace(None)

    assert interrupted_reader is not None
    assert id(interrupted_reader) in closed_reader_ids
    assert source.tell() == 4
    assert source.closed is False
    assert workbook._active_operation_token is None
    assert workbook._active_stream is None
    stream.close()
    workbook.close()


def test_csv_path_file_init_return_gap_closes_file_and_token(
    tmp_path: Path,
) -> None:
    path = tmp_path / "path-init-gap.csv"
    path.write_bytes(b"value\n1\n2\n")
    workbook = MessyWorkbook(path)
    stream = workbook.iter_batches(
        batch_size=1,
        config=SheetConfig(auto_detect=False),
    )
    target_code = csv_streaming_module._initialize_file_io.__code__
    interrupted_file: io.FileIO | None = None

    def interrupt(frame: Any, event: str, _arg: object) -> Any:
        nonlocal interrupted_file
        if frame.f_code is target_code and event == "return" and interrupted_file is None:
            interrupted_file = frame.f_locals["file"]
            sys.settrace(None)
            frame.f_trace = None
            raise MemoryError("CSV path file init return interrupted")
        return interrupt

    sys.settrace(interrupt)
    try:
        with pytest.raises(MemoryError, match="path file init return interrupted"):
            next(stream)
    finally:
        sys.settrace(None)

    assert interrupted_file is not None
    assert interrupted_file.closed is True
    assert workbook._active_operation_token is None
    assert workbook._active_stream is None
    stream.close()
    workbook.close()


@pytest.mark.parametrize("encoding", ["utf-8", "utf-16"])
def test_csv_sample_rejects_oversized_quoted_logical_record_before_parsing(
    encoding: str,
) -> None:
    character_count = MAX_SAMPLE_BYTES + 1 if encoding == "utf-8" else MAX_SAMPLE_BYTES // 2 + 1
    content = f'value\n"{"x" * character_count}"\n'.encode(encoding)
    source = NoUnboundedReadBytesIO(content)
    source.seek(7)
    workbook = MessyWorkbook(source, filename=f"oversized-{encoding}.csv")
    try:
        with pytest.raises(
            FormatError,
            match=rf"logical CSV record exceeds.*{MAX_SAMPLE_BYTES}",
        ):
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(auto_detect=False),
            )
        assert source.tell() == 7
        assert source.closed is False
        assert workbook._active_operation_token is None
        assert workbook._active_stream is None
    finally:
        workbook.close()


@pytest.mark.parametrize(
    ("content", "expected"),
    [
        (b"a\n1\n2\n\n", [1, 2]),
        (b"a\n1\n2\nbad,extra\n", [1, 2]),
        (b'a\n"line-1\nline-2"\nlast\n', ["line-1\nline-2"]),
        (b"a\r\n1\r\n2\r\n\r\n", [1, 2]),
    ],
)
def test_csv_streaming_footer_uses_physical_logical_record_semantics(
    content: bytes,
    expected: list[object],
) -> None:
    config = SheetConfig(
        auto_detect=False,
        skip_footer=1,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="physical-footer.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        table = _table(stream)

    assert table.column(0).to_pylist() == expected


@pytest.mark.parametrize(
    "content",
    [
        b"a,b\n1,x\x00junk\n2,y\n",
        b'a,b\n1,"x"junk\n2,y\n',
        b"a\r\n1\r\n2\r\n",
    ],
)
def test_csv_no_footer_streaming_matches_materialized_c_engine(content: bytes) -> None:
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="c-engine-parity.csv") as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", LegacyAPIWarning)
        expected = workbook.to_dataframe(config=config)
    with (
        MessyWorkbook(io.BytesIO(content), filename="c-engine-parity.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        actual = _table(stream).to_pandas()

    pd.testing.assert_frame_equal(actual, expected)


def test_csv_path_late_decode_failure_is_lazy_for_streaming(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    prefix_rows = b"".join(f"row-{index},ascii\n".encode() for index in range(5_000))
    path = tmp_path / "lazy-late-encoding.csv"
    path.write_bytes(b"name,note\n" + prefix_rows + b"last,\x93late\n")
    invalid_offset = path.read_bytes().index(b"\x93")
    real_open = Path.open
    maximum_read_position = 0

    class ObservedPathFile:
        def __init__(self, raw: Any) -> None:
            self._raw = raw

        def read(self, size: int = -1) -> object:
            nonlocal maximum_read_position
            result = self._raw.read(size)
            maximum_read_position = max(maximum_read_position, self._raw.tell())
            return result

        def __enter__(self) -> ObservedPathFile:
            return self

        def __exit__(self, *_args: object) -> None:
            self._raw.close()

        def __getattr__(self, name: str) -> object:
            return getattr(self._raw, name)

    def observed_open(target: Path, *args: object, **kwargs: object) -> object:
        raw = real_open(target, *args, **kwargs)
        if target == path and "b" in str(args[0] if args else kwargs.get("mode", "r")):
            return ObservedPathFile(raw)
        return raw

    monkeypatch.setattr(Path, "open", observed_open)
    workbook = MessyWorkbook(path)
    stream = workbook.iter_batches(
        batch_size=127,
        config=SheetConfig(
            auto_detect=False,
            normalize=False,
            sanitize_column_names=False,
        ),
    )
    assert maximum_read_position < invalid_offset
    try:
        with pytest.raises(FormatError, match="codec can't decode byte"):
            list(stream)
    finally:
        stream.close()
        workbook.close()


@pytest.mark.parametrize("source_kind", ["path", "seekable"])
def test_materialized_metadata_detection_reads_complete_bounded_rows(
    source_kind: str,
    tmp_path: Path,
) -> None:
    content = b'"' + (b"x" * 70_000) + b'",\nname,value\nalice,1\nbob,2\n'
    path = tmp_path / "long-metadata.csv"
    path.write_bytes(content)
    source: object = path if source_kind == "path" else NoUnboundedReadBytesIO(content)
    with (
        MessyWorkbook(source, filename="long-metadata.csv") as workbook,  # type: ignore[arg-type]
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", LegacyAPIWarning)
        frame = workbook.to_dataframe(
            config=SheetConfig(
                normalize=False,
                sanitize_column_names=False,
            )
        )

    assert frame.columns.tolist() == ["name", "value"]
    assert frame.to_dict("list") == {
        "name": ["alice", "bob"],
        "value": [1, 2],
    }


def test_csv_sample_budget_excludes_explicitly_skipped_framing_rows() -> None:
    content = ((b"x" * (1024 * 1024) + b"\n") * 9) + b"value\n1\n2\n"
    config = SheetConfig(
        auto_detect=False,
        skip_rows=9,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="large-skips.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        table = _table(stream)

    assert table.to_pydict() == {"value": [1, 2]}


@pytest.mark.parametrize("line_ending", [b"\r", b"\r\n"])
def test_csv_sample_budget_handles_ignored_prefix_record_endings(
    line_ending: bytes,
) -> None:
    data_record = b"d" * (MAX_SAMPLE_BYTES // 2 + 1)
    content = (
        (b"x" * (1024 * 1024)) + line_ending + data_record + line_ending + data_record + line_ending
    )
    reader = LogicalRecordBudgetReader(
        io.BytesIO(content),
        "utf-8",
        ",",
        "record-endings.csv",
        enforce_total=True,
        ignored_prefix_records=1,
    )
    try:
        with pytest.raises(FormatError, match="sample window exceeds"):
            while reader.read(64 * 1024):
                pass
    finally:
        reader.close()


def test_materialized_encoding_fallback_does_not_mask_process_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    path = tmp_path / "fallback-process.csv"
    path.write_bytes(b"value\n1\n")
    calls = 0

    def fail_read(*_args: object, **_kwargs: object) -> pd.DataFrame:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise UnicodeDecodeError("utf-8", b"\xff", 0, 1, "invalid")
        raise MemoryError("fallback process failure")

    monkeypatch.setattr(pd, "read_csv", fail_read)
    with pytest.raises(MemoryError, match="fallback process failure"):
        CSVHandler().parse(
            path,
            None,
            ParseOptions(auto_detect_header=False),
        )


def test_csv_c_engine_malformed_filter_is_batch_boundary_independent() -> None:
    content = b"a,b\n1,x\n2,y,z\n3,q\n4,r\n"
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="chunk-boundary.csv") as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore")
        expected = workbook.to_dataframe(config=config)

    for batch_size in (1, 2, 3):
        with (
            MessyWorkbook(
                io.BytesIO(content),
                filename="chunk-boundary.csv",
            ) as workbook,
            workbook.iter_batches(batch_size=batch_size, config=config) as stream,
            warnings.catch_warnings(),
        ):
            warnings.simplefilter("ignore")
            actual = _table(stream).to_pandas()
        pd.testing.assert_frame_equal(actual, expected.reset_index(drop=True))


def test_csv_c_engine_preserves_first_row_implicit_index_inference() -> None:
    content = b"a,b\n1,x,y\n2,z\n3,q\n"
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="implicit-index.csv") as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore")
        expected = workbook.to_dataframe(config=config)

    for batch_size in (1, 2, 3):
        with (
            MessyWorkbook(
                io.BytesIO(content),
                filename="implicit-index.csv",
            ) as workbook,
            workbook.iter_batches(batch_size=batch_size, config=config) as stream,
            warnings.catch_warnings(),
        ):
            warnings.simplefilter("ignore")
            actual = _table(stream).to_pandas()
        pd.testing.assert_frame_equal(actual, expected.reset_index(drop=True))


def test_csv_c_engine_ignores_blank_before_implicit_index_inference() -> None:
    content = b"a,b\n\n1,x,y\n2,z\n3,q\n"
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="blank-index.csv") as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore")
        expected = workbook.to_dataframe(config=config).reset_index(drop=True)
    with (
        MessyWorkbook(io.BytesIO(content), filename="blank-index.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore")
        actual = _table(stream).to_pandas()

    pd.testing.assert_frame_equal(actual, expected)


def test_csv_multi_header_footer_protects_all_non_data_records() -> None:
    content = b"group-a,group-b\nname,value\nalice,1\nbob,2\n"
    config = SheetConfig(
        auto_detect=False,
        header_rows=2,
        skip_footer=2,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="multi-header-footer.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        table = _table(stream)

    assert table.num_rows == 0
    assert table.column_names == ["name__alice", "value__1"]


@pytest.mark.parametrize(
    ("encoding", "content", "expected"),
    [
        (
            "utf-8-sig",
            b'\xef\xbb\xbf"line-1\nline-2"\ntail\n',
            b'\xef\xbb\xbf"line-1\nline-2"\n',
        ),
        (
            "utf-16-le",
            '\ufeff"line-1\r\nline-2"\r\ntail\r\n'.encode("utf-16-le"),
            '\ufeff"line-1\r\nline-2"\r\n'.encode("utf-16-le"),
        ),
    ],
)
def test_csv_record_framing_handles_bom_multiline_and_crlf(
    encoding: str,
    content: bytes,
    expected: bytes,
) -> None:
    class OneByteReads(io.BytesIO):
        def read(self, size: int = -1) -> bytes:
            return super().read(1 if size != 0 else 0)

    reader = FooterTrimmingReader(
        OneByteReads(content),
        encoding,
        ",",
        "framing.csv",
        skip_footer=1,
        protected_prefix_records=0,
    )
    try:
        chunks = iter(lambda: reader.read(2), b"")
        assert b"".join(chunks) == expected
    finally:
        reader.close()


def test_csv_record_limit_handles_chunk_split_escaped_quote() -> None:
    content = b'a\n"left ""quoted""\nright"\ntail\n'

    class OneByteReads(io.BytesIO):
        def read(self, size: int = -1) -> bytes:
            return super().read(1 if size != 0 else 0)

    reader = RecordLimitingReader(
        OneByteReads(content),
        "utf-8",
        ",",
        "limit.csv",
        max_records=2,
    )
    try:
        chunks = iter(lambda: reader.read(3), b"")
        assert b"".join(chunks) == b'a\n"left ""quoted""\nright"\n'
    finally:
        reader.close()


def test_csv_record_limit_treats_utf8_bom_first_quote_as_field_start() -> None:
    content = b'\xef\xbb\xbf"line-1\nline-2"\ntail\n'

    class OneByteReads(io.BytesIO):
        def read(self, size: int = -1) -> bytes:
            return super().read(1 if size != 0 else 0)

    reader = RecordLimitingReader(
        OneByteReads(content),
        "utf-8-sig",
        ",",
        "bom-limit.csv",
        max_records=1,
    )
    try:
        chunks = iter(lambda: reader.read(2), b"")
        assert b"".join(chunks) == b'\xef\xbb\xbf"line-1\nline-2"\n'
    finally:
        reader.close()


def test_csv_nested_input_proxy_process_close_failure_is_retryable() -> None:
    class FailOnceClose(io.BytesIO):
        close_calls = 0

        def close(self) -> None:
            self.close_calls += 1
            if self.close_calls == 1:
                raise MemoryError("nested CSV input close failure")
            super().close()

    raw = FailOnceClose(b"a\n1\n")
    limited = RecordLimitingReader(
        raw,
        "utf-8",
        ",",
        "nested-close.csv",
        max_records=2,
        owns_stream=True,
    )
    footer = FooterTrimmingReader(
        limited,
        "utf-8",
        ",",
        "nested-close.csv",
        skip_footer=0,
        protected_prefix_records=1,
        owns_stream=True,
    )
    outer = MalformedRecordFilteringReader(
        footer,
        "utf-8",
        "strict",
        ",",
        "nested-close.csv",
        expected_fields=1,
        protected_prefix_records=1,
        owns_stream=True,
    )

    with pytest.raises(MemoryError, match="nested CSV input close failure"):
        outer.close()
    assert outer.closed is False
    assert footer.closed is False
    assert limited.closed is False
    assert raw.closed is False

    outer.close()
    assert outer.closed is True
    assert footer.closed is True
    assert limited.closed is True
    assert raw.closed is True


def test_csv_no_footer_max_rows_counts_only_accepted_rows() -> None:
    content = b"a,b\n1,10\n2,20,EXTRA\n3,30\n4,40\n"
    options = ParseOptions(auto_detect_header=False, max_rows=2)
    expected = CSVHandler().parse(io.BytesIO(content), None, options)
    source = SourceHandle(io.BytesIO(content), filename="max-rows-malformed.csv")
    reader = CSVStreamingReader.prepare(
        source,
        options,
        CSVInspection("utf-8", ",", 0, "ignore"),
        ("a", "b"),
        pa.schema([pa.field("0", pa.string()), pa.field("1", pa.string())]),
        pa.schema([pa.field("0", pa.int64()), pa.field("1", pa.int64())]),
        1,
    )
    try:
        streamed_rows = 0
        while batch := reader.read_next_batch():
            streamed_rows += batch.num_rows
    finally:
        reader.close()
        source.close()

    assert expected.to_dict("list") == {"a": [1, 3], "b": [10, 30]}
    assert streamed_rows == len(expected)


def test_csv_path_uses_legacy_fallback_when_invalid_byte_is_in_sample(
    tmp_path: Path,
) -> None:
    rows = [f"row-{index:04d},ascii-value\n".encode() for index in range(650)]
    rows[550] = b"row-0550,\x93sample\n"
    path = tmp_path / "sample-encoding-fallback.csv"
    path.write_bytes(b"name,note\n" + b"".join(rows))
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )

    with (
        MessyWorkbook(path) as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", LegacyAPIWarning)
        expected = workbook.to_dataframe(config=config)
    with (
        MessyWorkbook(path) as workbook,
        workbook.iter_batches(batch_size=127, config=config) as stream,
    ):
        actual = _table(stream).to_pandas()

    assert actual.columns.tolist() == expected.columns.tolist()
    assert actual.shape == expected.shape
    assert actual.iloc[550].tolist() == expected.iloc[550].tolist()


def test_csv_arbitrary_sniffer_delimiter_matches_legacy_materialized_output() -> None:
    content = b"a^b\n1^2\n3^4\n"
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="caret.csv") as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", LegacyAPIWarning)
        expected = workbook.to_dataframe(config=config)
    with (
        MessyWorkbook(io.BytesIO(content), filename="caret.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
    ):
        actual = _table(stream).to_pandas()

    assert expected.columns.tolist() == ["a", "b"]
    pd.testing.assert_frame_equal(actual, expected)


def test_csv_full_pass_accepts_record_larger_than_sample_budget_without_filter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    large_value = "x" * (MAX_SAMPLE_BYTES + 1)
    content = (
        b"value\n"
        + b"".join(f"small-{index}\n".encode() for index in range(1_500))
        + large_value.encode()
        + b"\n"
    )
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )

    def fail_filter(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("ordinary no-footer input must use native pandas C")

    monkeypatch.setattr(
        csv_streaming_module,
        "MalformedRecordFilteringReader",
        fail_filter,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="large-record.csv") as workbook,
        workbook.iter_batches(batch_size=256, config=config) as stream,
    ):
        table = _table(stream)

    assert table.num_rows == 1_501
    assert len(table.column(0)[-1].as_py()) == MAX_SAMPLE_BYTES + 1


def test_unquoted_csv_filter_uses_raw_fast_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    content = b"a,b\n1,2\n3,4\n"

    def unexpected_csv_reader(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("clean unquoted records must not use csv.reader")

    monkeypatch.setattr(csv_io_module.csv, "reader", unexpected_csv_reader)
    reader = MalformedRecordFilteringReader(
        io.BytesIO(content),
        "utf-8",
        "strict",
        ",",
        "fast.csv",
        expected_fields=2,
        protected_prefix_records=1,
        owns_stream=False,
    )
    try:
        assert b"".join(iter(lambda: reader.read(4), b"")) == content
    finally:
        reader.close()


def test_csv_sample_process_failure_is_not_retried_as_encoding_fallback(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    path = tmp_path / "process-fallback.csv"
    path.write_bytes(b"value\n1\n")
    calls = 0

    def fail_then_mask(
        *_args: object,
        **_kwargs: object,
    ) -> pd.DataFrame:
        nonlocal calls
        calls += 1
        if calls == 1:
            decode = UnicodeDecodeError("utf-8", b"\xff", 0, 1, "invalid")
            raise MemoryError("primary sample process failure") from decode
        return pd.DataFrame({"value": [1]})

    monkeypatch.setattr(csv_streaming_module, "_read_bounded_sample", fail_then_mask)
    workbook = MessyWorkbook(path)
    try:
        with pytest.raises(MemoryError, match="primary sample process failure"):
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(auto_detect=False),
            )
        assert calls == 1
        assert workbook.parse_metrics == _expected_native_metrics(failed_attempts=1)
    finally:
        workbook.close()


def test_csv_metadata_restore_failure_does_not_replace_process_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class RestoreFailure(io.BytesIO):
        zero_seeks = 0

        def seek(self, offset: int, whence: int = 0) -> int:
            if offset == 0 and whence == 0:
                self.zero_seeks += 1
                if self.zero_seeks == 2:
                    raise OSError("metadata cursor restore failure")
            return super().seek(offset, whence)

    handler = CSVHandler()
    source = RestoreFailure(b"value\n1\n")

    def fail_detection(*_args: object, **_kwargs: object) -> int:
        raise MemoryError("metadata process failure")

    monkeypatch.setattr(
        handler._detector,
        "detect_skip_rows_from_binary",
        fail_detection,
    )
    with pytest.raises(MemoryError, match="metadata process failure"):
        handler._detect_skip_rows_from_target(
            source,
            "utf-8",
            ",",
            "restore.csv",
            is_stream=True,
        )
    assert source.zero_seeks == 2


@pytest.mark.parametrize(
    "content",
    [
        b'a,b\n1,"left,right"\n2,"up,down"\n',
        b"a,b\r\n1,left\r\n2,right\r\n",
    ],
    ids=["quoted-lf", "crlf"],
)
@pytest.mark.parametrize("batch_size", [1, 2, 3])
def test_no_footer_native_route_avoids_python_record_filter(
    content: bytes,
    batch_size: int,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail_transform(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("ordinary no-footer input must use native pandas C directly")

    monkeypatch.setattr(
        csv_streaming_module,
        "MalformedRecordFilteringReader",
        fail_transform,
    )
    monkeypatch.setattr(
        csv_streaming_module,
        "FooterTrimmingReader",
        fail_transform,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="native.csv") as workbook,
        workbook.iter_batches(
            batch_size=batch_size,
            config=SheetConfig(
                auto_detect=False,
                normalize=False,
                sanitize_column_names=False,
            ),
        ) as stream,
    ):
        assert _table(stream).num_rows == 2


def test_csv_cr_only_malformed_row_matches_materialized_c_engine() -> None:
    content = b'a,b\r x ,q"z,q"z\r'
    config = SheetConfig(
        auto_detect=False,
        normalize=False,
        sanitize_column_names=False,
    )
    with (
        MessyWorkbook(io.BytesIO(content), filename="cr-malformed.csv") as workbook,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", LegacyAPIWarning)
        expected = workbook.to_dataframe(config=config)
    with (
        MessyWorkbook(io.BytesIO(content), filename="cr-malformed.csv") as workbook,
        workbook.iter_batches(batch_size=1, config=config) as stream,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore")
        actual = _table(stream).to_pandas()

    pd.testing.assert_frame_equal(actual, expected)
