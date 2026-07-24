"""Task-12 adapters from bounded evidence or materialized frames to Arrow streams."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import cast

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.compute as pc

from messy_xlsx._fallback_signals import _contains_process_failure, _exception_traceback
from messy_xlsx.exceptions import StreamingTypeError
from messy_xlsx.normalization import (
    NormalizationPlan,
    NormalizationSample,
    NormalizedStreamingReader,
    compile_normalization_plan,
)
from messy_xlsx.normalization.plan import (
    MAX_SAMPLE_BYTES,
    MAX_SAMPLE_CELLS,
    MAX_SAMPLE_VALUES,
    _display_label_token,
    _LabelResolutionIndex,
    _safe_name_text,
    _timezone_label_projection,
)
from messy_xlsx.normalization.plan import (
    _label_tokens_match as _plan_label_tokens_match,
)
from messy_xlsx.parsing.contracts import StreamingBatchReader
from messy_xlsx.parsing.coordinates import ColumnIdentity, CoordinateBatch
from messy_xlsx.parsing.parse_plan import ParsePlan
from messy_xlsx.parsing.physical_values import (
    PandasTemporalPayload,
    UnsupportedPhysicalValueError,
    arrow_temporal_array,
    common_temporal_arrow_type,
    decode_physical_value,
    encode_physical_value,
    pandas_temporal_payload,
    physical_label_description,
    physical_normalization_value,
    physical_value_description,
    physical_value_family,
    physical_value_matches_arrow_type,
    temporal_payload,
)
from messy_xlsx.parsing.streams import _close_if_present, _run_cleanups
from messy_xlsx.parsing.xlsx_streaming import _record_batch_with_row_count

_ARROW_ARRAY_ERRORS = (
    pa.ArrowInvalid,
    pa.ArrowNotImplementedError,
    pa.ArrowTypeError,
    OverflowError,
    TypeError,
    ValueError,
)
_MAX_PUBLIC_LABEL_DEPTH = 16
_MAX_PUBLIC_LABEL_MEMBERS = 256
_label_tokens_match = _plan_label_tokens_match


def _public_dataframe_display_names(
    source_names: tuple[object, ...],
    safe_names: tuple[object, ...],
    plan: ParsePlan,
) -> tuple[object, ...]:
    """Retain only inert immutable labels in the pandas-only positional sidecar."""
    if plan.sanitize_column_names:
        return safe_names
    renamed_tokens = tuple(
        _display_label_token(label) for label, _value in plan.thaw_column_rename_items()
    )
    renamed_index = _LabelResolutionIndex(renamed_tokens)
    display_names: list[object] = []
    for source_name, safe_name in zip(source_names, safe_names, strict=True):
        source_token = _display_label_token(source_name)
        was_renamed = bool(renamed_index.matching(source_token))
        if was_renamed:
            display_names.append(safe_name)
            continue
        is_safe, exact_name = _safe_public_dataframe_label(
            source_name,
            depth=0,
            budget=[_MAX_PUBLIC_LABEL_MEMBERS],
        )
        display_names.append(exact_name if is_safe else safe_name)
    return tuple(display_names)


def _safe_public_dataframe_label(
    value: object,
    *,
    depth: int,
    budget: list[int],
) -> tuple[bool, object]:
    if depth > _MAX_PUBLIC_LABEL_DEPTH or budget[0] <= 0:
        return False, value
    budget[0] -= 1
    value_type = type(value)
    if value is pd.NaT:
        return True, value
    if value is None or value_type in {
        str,
        int,
        float,
        complex,
        bool,
        bytes,
        Decimal,
    }:
        return True, value
    if value_type is pd.Timedelta:
        return True, value
    if value_type is pd.Timestamp:
        timestamp = cast(pd.Timestamp, value)
        return _timezone_label_projection(timestamp.tzinfo)[1], value
    if value_type is tuple:
        for member in cast(tuple[object, ...], value):
            is_safe, _exact_member = _safe_public_dataframe_label(
                member,
                depth=depth + 1,
                budget=budget,
            )
            if not is_safe:
                return False, value
        return True, value
    return False, value


@dataclass(frozen=True, slots=True)
class PreparedStreamingReader:
    """One fully wrapped reader and its positional display-name sidecar."""

    reader: StreamingBatchReader
    display_names: tuple[object, ...]


class _CloseOnceReader:
    """Keep constructor and wrapper rollback on one close-once boundary."""

    def __init__(self, reader: StreamingBatchReader | None = None) -> None:
        self._reader: StreamingBatchReader | None = reader

    def attach(self, reader: StreamingBatchReader) -> None:
        if self._reader is not None:
            raise RuntimeError("reader owner is already attached")
        self._reader = reader

    @property
    def schema(self) -> pa.Schema:
        reader = self._reader
        if reader is None:
            raise RuntimeError("stream reader is closed")
        return reader.schema

    def read_next_batch(self) -> pa.RecordBatch | None:
        reader = self._reader
        if reader is None:
            return None
        return reader.read_next_batch()

    def close(self) -> None:
        reader = self._reader
        if reader is None:
            return
        try:
            self._reader = reader.close()
        except BaseException as error:
            if not _contains_process_failure(error):
                self._reader = None
            raise


def prepare_materialized_streaming_reader(
    frame: pd.DataFrame,
    plan: ParsePlan,
    batch_size: int,
    *,
    date_system: str,
) -> PreparedStreamingReader:
    """Apply the Task-11 streaming contract to one already materialized raw frame."""
    source_display_names = tuple(frame.columns)
    sample = normalization_sample_from_dataframe(
        frame,
        date_system=date_system,
        preserve_native=not plan.normalize,
    )
    normalization_plan = compile_normalization_plan(sample, plan)
    pandas_display_names = _public_dataframe_display_names(
        source_display_names,
        normalization_plan.final_display_names,
        plan,
    )
    raw_reader: StreamingBatchReader | None = None
    owned_reader: _CloseOnceReader | None = None
    try:
        raw_reader = _EncodedDataFrameReader(frame, batch_size)
        owned_reader = _CloseOnceReader(raw_reader)
        return wrap_normalized_streaming_reader(
            owned_reader,
            normalization_plan,
            normalize=plan.normalize,
            pandas_display_names=pandas_display_names,
        )
    except BaseException as error:
        rollback_reader = owned_reader if owned_reader is not None else raw_reader
        cleanups = (
            []
            if rollback_reader is None
            else [("materialized raw reader rollback", lambda: _close_if_present(rollback_reader))]
        )
        _run_cleanups(
            cleanups,
            primary_error=error,
            primary_traceback=_exception_traceback(error),
        )
        raise


def wrap_normalized_streaming_reader(
    raw_reader: StreamingBatchReader,
    plan: NormalizationPlan,
    *,
    normalize: bool,
    pandas_display_names: tuple[object, ...] | None = None,
    rollback_on_error: bool = True,
) -> PreparedStreamingReader:
    """Transactionally install physical, normalization, and public-schema adapters."""
    rollback_reader = raw_reader
    try:
        physical_reader = PhysicalTypeReader(
            raw_reader,
            plan.input_schema,
            plan.source_display_names,
            normalize=normalize,
        )
        rollback_reader = physical_reader
        normalized = NormalizedStreamingReader(physical_reader, plan)
        rollback_reader = normalized
        public = PublicSchemaReader(normalized, plan.final_display_names)
        rollback_reader = public
        return PreparedStreamingReader(
            reader=public,
            display_names=(
                plan.final_display_names if pandas_display_names is None else pandas_display_names
            ),
        )
    except BaseException as error:
        if not rollback_on_error:
            raise
        _run_cleanups(
            [("stream wrapper rollback", rollback_reader.close)],
            primary_error=error,
            primary_traceback=_exception_traceback(error),
        )
        raise


class _NormalizationSampleAccumulator:
    """Own decoded sample evidence without retaining transformed Arrow batches."""

    def __init__(
        self,
        probe_schema: pa.Schema,
        *,
        date_system: str,
        preserve_native: bool,
        max_rows: int,
        max_cells: int,
        max_bytes: int,
    ) -> None:
        self._probe_schema = probe_schema
        self._date_system = date_system
        self._preserve_native = preserve_native
        self._max_rows = max_rows
        self._max_cells = max_cells
        self._max_bytes = max_bytes
        self._identities: tuple[ColumnIdentity, ...] = ()
        self._column_values: list[list[object | None]] = [[] for _ in probe_schema]
        self._row_numbers: list[int] = []
        self._retained_cells = 0
        self._retained_bytes = 0
        self._full = False

    @property
    def full(self) -> bool:
        return self._full

    def consume(self, batches: tuple[CoordinateBatch, ...]) -> None:
        for coordinate_batch in batches:
            if coordinate_batch.batch.num_columns != len(self._probe_schema):
                raise ValueError("normalization sample width is unstable")
            if coordinate_batch.column_identities:
                self._identities = coordinate_batch.column_identities
            for offset in range(coordinate_batch.batch.num_rows):
                width = coordinate_batch.batch.num_columns
                if (
                    len(self._row_numbers) >= self._max_rows
                    or self._retained_cells + width > self._max_cells
                ):
                    self._full = True
                    return
                encoded_row = [
                    coordinate_batch.batch.column(ordinal)[offset].as_py()
                    for ordinal in range(width)
                ]
                row_bytes = sum(_sample_scalar_byte_cost(value) for value in encoded_row)
                if row_bytes > self._max_bytes:
                    raise ValueError("sample row exceeds the Arrow byte budget")
                if self._retained_bytes + row_bytes > self._max_bytes:
                    self._full = True
                    return
                for values, encoded in zip(
                    self._column_values,
                    encoded_row,
                    strict=True,
                ):
                    values.append(decode_physical_value(encoded))
                self._row_numbers.append(int(coordinate_batch.row_numbers[offset].as_py()))
                self._retained_cells += width
                self._retained_bytes += row_bytes

    def finish(
        self,
        identity_snapshot: tuple[ColumnIdentity, ...] = (),
    ) -> NormalizationSample:
        identities = self._identities or identity_snapshot
        if not identities:
            identities = tuple(
                ColumnIdentity(ordinal, f"col_{ordinal}")
                for ordinal in range(len(self._probe_schema))
            )
        return _normalization_sample(
            identities,
            self._column_values,
            self._row_numbers,
            date_system=self._date_system,
            preserve_native=self._preserve_native,
        )


def normalization_sample_from_coordinate_batches(
    batches: tuple[CoordinateBatch, ...],
    probe_schema: pa.Schema,
    *,
    date_system: str,
    identity_snapshot: tuple[ColumnIdentity, ...] = (),
    preserve_native: bool = False,
) -> NormalizationSample:
    """Build bounded physical evidence from post-coordinate tagged batches."""
    identities: tuple[ColumnIdentity, ...] = ()
    column_values: list[list[object | None]] = [[] for _ in probe_schema]
    row_numbers: list[int] = []
    remaining = min(
        MAX_SAMPLE_VALUES,
        MAX_SAMPLE_CELLS // max(1, len(probe_schema)),
    )
    retained_bytes = 0
    for coordinate_batch in batches:
        if coordinate_batch.column_identities:
            identities = coordinate_batch.column_identities
        take = min(remaining, coordinate_batch.batch.num_rows)
        for offset in range(take):
            encoded_row = [
                coordinate_batch.batch.column(ordinal)[offset].as_py()
                for ordinal in range(len(probe_schema))
            ]
            row_bytes = sum(_sample_scalar_byte_cost(value) for value in encoded_row)
            if retained_bytes + row_bytes > MAX_SAMPLE_BYTES:
                remaining = 0
                break
            for values, encoded in zip(column_values, encoded_row, strict=True):
                values.append(decode_physical_value(encoded))
            row_numbers.append(int(coordinate_batch.row_numbers[offset].as_py()))
            retained_bytes += row_bytes
            remaining -= 1
        if remaining == 0:
            break
    if not identities and identity_snapshot:
        identities = identity_snapshot
    if not identities:
        identities = tuple(
            ColumnIdentity(ordinal, f"col_{ordinal}") for ordinal in range(len(probe_schema))
        )
    return _normalization_sample(
        identities,
        column_values,
        row_numbers,
        date_system=date_system,
        preserve_native=preserve_native,
    )


def normalization_sample_from_dataframe(
    frame: pd.DataFrame,
    *,
    date_system: str,
    preserve_native: bool = False,
) -> NormalizationSample:
    """Sample a raw handler frame positionally without dropping declared columns."""
    width = len(frame.columns)
    remaining = min(
        MAX_SAMPLE_VALUES,
        MAX_SAMPLE_CELLS // max(1, width),
    )
    column_values: list[list[object | None]] = [[] for _ in range(width)]
    row_numbers: list[int] = []
    retained_bytes = 0
    for position in range(min(len(frame), remaining)):
        values = [
            _canonical_materialized_value(frame.iloc[position, ordinal]) for ordinal in range(width)
        ]
        encoded: list[str | None] = []
        for ordinal, value in enumerate(values):
            try:
                encoded.append(encode_physical_value(value))
            except UnsupportedPhysicalValueError:
                raise StreamingTypeError(
                    "streamed value is incompatible with the fixed schema",
                    ordinal=ordinal,
                    display_label=physical_label_description(frame.columns[ordinal]),
                    row_offset=position,
                    value_description=physical_value_description(value),
                    expected_type="supported Arrow scalar",
                ) from None
        row_bytes = sum(_sample_scalar_byte_cost(value) for value in encoded)
        if retained_bytes + row_bytes > MAX_SAMPLE_BYTES:
            break
        for column, value in zip(column_values, values, strict=True):
            column.append(value)
        row_numbers.append(position + 1)
        retained_bytes += row_bytes
    identities = tuple(
        ColumnIdentity(ordinal, label) for ordinal, label in enumerate(frame.columns)
    )
    return _normalization_sample(
        identities,
        column_values,
        row_numbers,
        date_system=date_system,
        preserve_native=preserve_native,
    )


def _normalization_sample(
    identities: tuple[ColumnIdentity, ...],
    column_values: list[list[object | None]],
    row_numbers: list[int],
    *,
    date_system: str,
    preserve_native: bool,
) -> NormalizationSample:
    columns = tuple(
        _infer_physical_sample_array(
            values,
            preserve_native=preserve_native,
            ordinal=ordinal,
            display_label=identities[ordinal].display_name,
        )
        for ordinal, values in enumerate(column_values)
    )
    schema = pa.schema(
        [pa.field(str(ordinal), column.type) for ordinal, column in enumerate(columns)]
    )
    return NormalizationSample(
        schema=schema,
        column_identities=identities,
        columns=columns,
        row_numbers=pa.array(row_numbers, type=pa.int64()),
        date_system=date_system,
    )


def _infer_physical_sample_array(
    values: list[object | None],
    *,
    preserve_native: bool,
    ordinal: int,
    display_label: object,
) -> pa.Array:
    """Infer only a type that tagged string batches can reconstruct exactly."""
    if preserve_native:
        return _preserved_physical_sample_array(
            values,
            ordinal=ordinal,
            display_label=display_label,
        )
    try:
        inferred = pa.array(values)
    except _ARROW_ARRAY_ERRORS:
        inferred = None
    if inferred is not None and (
        pa.types.is_null(inferred.type)
        or pa.types.is_string(inferred.type)
        or all(
            value is None or physical_value_matches_arrow_type(value, inferred.type)
            for value in values
        )
    ):
        return inferred
    lexical = pa.array(
        [None if value is None else physical_normalization_value(value) for value in values],
        type=pa.string(),
    )
    if inferred is None:
        return lexical
    try:
        pc.cast(lexical, inferred.type, safe=True)
    except _ARROW_ARRAY_ERRORS:
        return lexical
    return inferred


def _preserved_physical_sample_array(
    values: list[object | None],
    *,
    ordinal: int,
    display_label: object,
) -> pa.Array:
    target_family: tuple[str, str | None] | None = None
    compatible: list[object | None] = []
    for value in values:
        if value is None:
            compatible.append(None)
            continue
        family = physical_value_family(value)
        if family is None:
            raise TypeError("unsupported physical scalar type")
        if target_family is None:
            target_family = family
        compatible.append(value if family == target_family else None)
    if target_family is None:
        return pa.nulls(len(values))
    if target_family[0] in {"timestamp", "duration"}:
        payloads = [None if value is None else temporal_payload(value) for value in compatible]
        if all(
            value is None or payload is not None
            for value, payload in zip(compatible, payloads, strict=True)
        ):
            present = [
                cast(PandasTemporalPayload, payload) for payload in payloads if payload is not None
            ]
            target = common_temporal_arrow_type(present)
            if target is None:
                raise StreamingTypeError(
                    "sampled pandas temporal values have no lossless common Arrow unit",
                    ordinal=ordinal,
                    display_label=physical_label_description(display_label),
                    row_offset=0,
                    value_description="datetime",
                    expected_type=target_family[0],
                )
            return arrow_temporal_array(
                cast("list[PandasTemporalPayload | None]", payloads),
                target,
            )
    return pa.array(compatible, from_pandas=True)


def _sample_scalar_byte_cost(value: object) -> int:
    """Conservatively bound retained Arrow value and offset buffers."""
    if isinstance(value, str):
        return len(value.encode("utf-8", errors="surrogatepass")) + 16
    return 16


def _canonical_materialized_value(value: object) -> object | None:
    """Convert pandas/numpy scalar wrappers to the physical Python scalar."""
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    try:
        temporal = pandas_temporal_payload(value)
    except UnsupportedPhysicalValueError:
        temporal = None
    if temporal is not None:
        return temporal
    try:
        missing = pd.isna(value)
    except (TypeError, ValueError):
        missing = False
    if isinstance(missing, (bool, np.bool_)) and bool(missing):
        return None
    if isinstance(value, np.generic):
        return cast(object, value.item())
    return value


class _EncodedDataFrameReader:
    """Yield one materialized frame as tagged ordinal string batches."""

    def __init__(self, frame: pd.DataFrame, batch_size: int) -> None:
        self._frame: pd.DataFrame | None = frame
        self._batch_size = batch_size
        self._offset = 0
        self._display_names = tuple(frame.columns)
        self._schema = pa.schema(
            [pa.field(str(ordinal), pa.string()) for ordinal in range(len(frame.columns))]
        )

    @property
    def schema(self) -> pa.Schema:
        return self._schema

    def read_next_batch(self) -> pa.RecordBatch | None:
        frame = self._frame
        if frame is None or self._offset >= len(frame):
            return None
        stop = min(len(frame), self._offset + self._batch_size)
        arrays: list[pa.Array] = []
        for ordinal in range(len(frame.columns)):
            encoded: list[str | None] = []
            for relative, value in enumerate(frame.iloc[self._offset : stop, ordinal].tolist()):
                try:
                    encoded.append(encode_physical_value(_canonical_materialized_value(value)))
                except UnsupportedPhysicalValueError:
                    raise StreamingTypeError(
                        "streamed value is incompatible with the fixed schema",
                        ordinal=ordinal,
                        display_label=physical_label_description(self._display_names[ordinal]),
                        row_offset=self._offset + relative,
                        value_description=physical_value_description(
                            _canonical_materialized_value(value)
                        ),
                        expected_type="supported Arrow scalar",
                    ) from None
            arrays.append(pa.array(encoded, type=pa.string()))
        row_count = stop - self._offset
        self._offset = stop
        return _record_batch_with_row_count(arrays, self._schema, row_count)

    def close(self) -> None:
        self._frame = None


class PhysicalTypeReader:
    """Restore sampled physical types from collision-safe tagged string batches."""

    def __init__(
        self,
        reader: StreamingBatchReader,
        schema: pa.Schema,
        display_names: tuple[object, ...],
        *,
        normalize: bool,
    ) -> None:
        if len(reader.schema) != len(schema):
            raise ValueError("sampled physical schema width does not match the reader")
        self._reader: StreamingBatchReader | None = reader
        self._schema = schema
        self._display_names = display_names
        self._normalize = normalize
        self._row_offset = 0

    @property
    def schema(self) -> pa.Schema:
        return self._schema

    def read_next_batch(self) -> pa.RecordBatch | None:
        reader = self._reader
        if reader is None:
            return None
        batch = reader.read_next_batch()
        if batch is None:
            return None
        arrays: list[pa.Array] = []
        for ordinal, target in enumerate(self._schema.types):
            source = batch.column(ordinal)
            try:
                converted = self._physical_candidate(source, target, ordinal)
                converted = self._cast_candidate(converted, target)
            except _ARROW_ARRAY_ERRORS:
                raw_values = source.to_pylist()
                offset = self._first_incompatible_offset(source, target, ordinal)
                self._raise_incompatible(
                    ordinal,
                    offset,
                    raw_values[offset],
                    target,
                )
            arrays.append(converted)
        self._row_offset += batch.num_rows
        return _record_batch_with_row_count(arrays, self._schema, batch.num_rows)

    def _physical_candidate(
        self,
        source: pa.Array,
        target: pa.DataType,
        ordinal: int,
    ) -> pa.Array:
        raw_values = source.to_pylist()
        if not (pa.types.is_string(source.type) or pa.types.is_large_string(source.type)):
            return source
        if not self._normalize:
            decoded_values = [
                None if value is None else decode_physical_value(value) for value in raw_values
            ]
            for offset, decoded in enumerate(decoded_values):
                if decoded is None:
                    continue
                if not physical_value_matches_arrow_type(decoded, target):
                    self._raise_incompatible(
                        ordinal,
                        offset,
                        raw_values[offset],
                        target,
                    )
            temporal = _decoded_pandas_temporal_array(decoded_values, target)
            if temporal is not None:
                return temporal
            return pa.array(decoded_values, type=target, from_pandas=True)
        decoded_values = [
            None if value is None else decode_physical_value(value) for value in raw_values
        ]
        if all(
            value is None or physical_value_matches_arrow_type(value, target)
            for value in decoded_values
        ):
            temporal = _decoded_pandas_temporal_array(decoded_values, target)
            if temporal is not None:
                return temporal
            return pa.array(decoded_values, type=target, from_pandas=True)
        lexical = [
            None if value is None else physical_normalization_value(value) for value in raw_values
        ]
        return pa.array(lexical, type=source.type)

    @staticmethod
    def _cast_candidate(source: pa.Array, target: pa.DataType) -> pa.Array:
        if source.type == target:
            return source
        if pa.types.is_null(target):
            if source.null_count != len(source):
                raise pa.ArrowInvalid("non-null value cannot fit a null schema")
            return pa.nulls(len(source))
        return pc.cast(source, target, safe=True)

    def _first_incompatible_offset(  # noqa: C901
        self,
        source: pa.Array,
        target: pa.DataType,
        ordinal: int,
    ) -> int:
        if not self._normalize:
            for offset, raw_value in enumerate(source.to_pylist()):
                if raw_value is None:
                    continue
                decoded = decode_physical_value(raw_value)
                if not physical_value_matches_arrow_type(decoded, target):
                    return offset
                try:
                    temporal = pandas_temporal_payload(decoded)
                    if temporal is not None and (
                        pa.types.is_timestamp(target) or pa.types.is_duration(target)
                    ):
                        arrow_temporal_array([temporal], target)
                    else:
                        pa.scalar(decoded, type=target)
                except _ARROW_ARRAY_ERRORS:
                    return offset
            return 0
        for offset in range(len(source)):
            raw_value = source[offset].as_py()
            if raw_value is None:
                continue
            try:
                candidate = self._physical_candidate(
                    source.slice(offset, 1),
                    target,
                    ordinal,
                )
                self._cast_candidate(candidate, target)
            except StreamingTypeError:
                return offset
            except _ARROW_ARRAY_ERRORS:
                return offset
        return 0

    def _raise_incompatible(
        self,
        ordinal: int,
        offset: int,
        value: object,
        target: pa.DataType,
    ) -> None:
        label = self._display_names[ordinal]
        raise StreamingTypeError(
            "streamed value is incompatible with the fixed schema",
            ordinal=ordinal,
            display_label=physical_label_description(label),
            row_offset=self._row_offset + offset,
            value_description=physical_value_description(value),
            expected_type=str(target),
        ) from None

    def close(self) -> None:
        reader = self._reader
        if reader is None:
            return
        try:
            self._reader = reader.close()
        except BaseException as error:
            if not _contains_process_failure(error):
                self._reader = None
            raise


def _decoded_pandas_temporal_array(
    values: list[object | None],
    target: pa.DataType,
) -> pa.Array | None:
    if not (pa.types.is_timestamp(target) or pa.types.is_duration(target)):
        return None
    payloads = [None if value is None else temporal_payload(value) for value in values]
    if not all(
        value is None or payload is not None
        for value, payload in zip(values, payloads, strict=True)
    ):
        return None
    return arrow_temporal_array(
        cast("list[PandasTemporalPayload | None]", payloads),
        target,
    )


class PublicSchemaReader:
    """Rename ordinal fields while preserving the compiled positional sidecar."""

    def __init__(
        self,
        reader: StreamingBatchReader,
        display_names: tuple[object, ...],
    ) -> None:
        self._reader: StreamingBatchReader | None = reader
        self.display_names = display_names
        self._schema = pa.schema(
            [
                pa.field(_safe_name_text(display_name), field.type)
                for display_name, field in zip(
                    display_names,
                    reader.schema,
                    strict=True,
                )
            ]
        )

    @property
    def schema(self) -> pa.Schema:
        return self._schema

    def read_next_batch(self) -> pa.RecordBatch | None:
        reader = self._reader
        if reader is None:
            return None
        batch = reader.read_next_batch()
        if batch is None:
            return None
        return _record_batch_with_row_count(
            list(batch.columns),
            self._schema,
            batch.num_rows,
        )

    def close(self) -> None:
        reader = self._reader
        if reader is None:
            return
        try:
            self._reader = reader.close()
        except BaseException as error:
            if not _contains_process_failure(error):
                self._reader = None
            raise
