"""Stateful, schema-stable Arrow normalization for bounded batches."""

from __future__ import annotations

import math
import re
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from types import TracebackType
from typing import Final, cast

import pandas as pd
import pyarrow as pa
import pyarrow.compute as pc

from messy_xlsx._fallback_signals import (
    _contains_process_failure,
    _exception_traceback,
    _fallback_block_reason,
    _FallbackBlockReason,
    _mark_fallback_blocked,
)
from messy_xlsx.exceptions import StreamingTypeError
from messy_xlsx.normalization.plan import (
    ColumnNormalization,
    ConditionMode,
    NormalizationPlan,
    RowCondition,
    SemanticOperation,
)
from messy_xlsx.parsing.contracts import StreamingBatchReader
from messy_xlsx.parsing.streams import _run_cleanups

_HORIZONTAL_SPACE: Final = re.compile(r"[ \t\xa0]+")
_ALL_SPACE: Final = re.compile(r"[\s\xa0]+")
_ARROW_UNICODE_SPACE_CHARS: Final = (
    r"\x09-\x0d\x1c-\x20\x{0085}\x{00a0}\x{1680}"
    r"\x{2000}-\x{200a}\x{2028}-\x{2029}\x{202f}\x{205f}\x{3000}"
)
_ARROW_ALL_SPACE: Final = rf"[{_ARROW_UNICODE_SPACE_CHARS}]+"
_ARROW_BLANK: Final = rf"^[{_ARROW_UNICODE_SPACE_CHARS}]*$"
_ARROW_NORMALIZED_WHITESPACE: Final = (
    rf"^(?:[^{_ARROW_UNICODE_SPACE_CHARS}]+"
    rf"(?: [^{_ARROW_UNICODE_SPACE_CHARS}]+)*)?$"
)
_CURRENCY: Final = re.compile(r"(?:[$€£¥₹]|CHF|kr|zł)")
_ACCOUNTING: Final = re.compile(r"^\((.*)\)$")
_INTEGER_TEXT: Final = re.compile(r"^[+-]?\d+$")
_ARROW_CONVERSION_ERRORS: Final = (
    pa.ArrowInvalid,
    pa.ArrowTypeError,
    pa.ArrowNotImplementedError,
    ValueError,
    OverflowError,
)


class ArrowNormalizationOperation:
    """Normalize one stream while tracking stable pre-filter row offsets."""

    def __init__(self, plan: NormalizationPlan) -> None:
        if not isinstance(plan, NormalizationPlan):
            raise TypeError("plan must be a NormalizationPlan")
        self._plan = plan
        self._row_offset = 0
        self._terminal = False
        self._missing_marker_arrays: dict[pa.DataType, pa.Array] = {}

    @property
    def schema(self) -> pa.Schema:
        return self._plan.schema

    @property
    def input_schema(self) -> pa.Schema:
        return self._plan.input_schema

    @property
    def row_offset(self) -> int:
        return self._row_offset

    def normalize(self, batch: pa.RecordBatch) -> pa.RecordBatch:
        """Transform one batch or make this operation terminal on failure."""
        if self._terminal:
            raise RuntimeError("normalization operation is terminal")
        try:
            if not isinstance(batch, pa.RecordBatch):
                raise TypeError("normalization input must be a RecordBatch")
            if not batch.schema.equals(self._plan.input_schema, check_metadata=True):
                raise _blocked_schema_error(
                    "normalization input schema does not match the fixed input schema"
                )
            if batch.num_rows > self._plan.max_input_rows:
                raise _blocked_schema_error("normalization input exceeds the compiled batch_size")
            input_rows = batch.num_rows
            row_base = self._row_offset
            result = self._normalize_batch(batch, row_base)
        except BaseException as error:
            self._terminal = True
            _mark_semantic_failure(error)
            raise
        self._row_offset += input_rows
        return result

    def _normalize_batch(self, batch: pa.RecordBatch, row_base: int) -> pa.RecordBatch:
        if not self._plan.normalize:
            if not batch.schema.equals(self._plan.schema, check_metadata=True):
                raise ValueError("raw normalization schema is unstable")
            return batch

        arrays: list[pa.Array] = []
        unchanged = batch.schema.equals(self._plan.schema, check_metadata=True)
        for rule in self._plan.columns:
            source = batch.column(rule.ordinal)
            normalized = self._normalize_column(source, rule, row_base)
            arrays.append(normalized)
            unchanged = unchanged and normalized is source
        result = batch if unchanged else _record_batch(arrays, self._plan.schema, batch.num_rows)
        result = _drop_all_null_rows(result)
        result = _apply_regex_filter(result, self._plan.drop_regex)
        result = _apply_conditions(result, self._plan)
        if not result.schema.equals(self._plan.schema, check_metadata=True):
            raise RuntimeError("normalization output schema is unstable")
        return result

    def _normalize_column(
        self,
        array: pa.Array,
        rule: ColumnNormalization,
        row_base: int,
    ) -> pa.Array:
        array = _decode_encoded_strings(array)
        if pa.types.is_null(rule.output_type):
            return self._normalize_null_column(array, rule, row_base)
        if (
            rule.semantic is SemanticOperation.DATE
            and rule.timezone is not None
            and (pa.types.is_string(array.type) or pa.types.is_large_string(array.type))
        ):
            mismatch = _first_fixed_timezone_mismatch(array, rule.timezone)
            if mismatch is not None:
                offset, value = mismatch
                self._raise_incompatible(rule, row_base + offset, value)
        fast = _normalize_arrow_column(array, rule, self._missing_marker_arrays)
        if fast is not None:
            return fast
        if rule.semantic is SemanticOperation.DATE and (
            pa.types.is_string(array.type) or pa.types.is_large_string(array.type)
        ):
            return self._normalize_date_strings(array, rule, row_base)
        return self._normalize_scalar_column(array, rule, row_base)

    def _normalize_null_column(
        self,
        array: pa.Array,
        rule: ColumnNormalization,
        row_base: int,
    ) -> pa.Array:
        for offset, scalar in enumerate(array):
            if not scalar.is_valid:
                continue
            value = scalar.as_py()
            normalized = _normalize_scalar(value, rule)
            if normalized is not None:
                self._raise_incompatible(rule, row_base + offset, value)
        return pa.nulls(len(array))

    def _normalize_scalar_column(
        self,
        array: pa.Array,
        rule: ColumnNormalization,
        row_base: int,
    ) -> pa.Array:
        converted: list[object | None] = []
        for offset, scalar in enumerate(array):
            if not scalar.is_valid:
                converted.append(None)
                continue
            value = scalar.as_py()
            try:
                normalized = _normalize_scalar(value, rule)
            except _ARROW_CONVERSION_ERRORS:
                self._raise_incompatible(rule, row_base + offset, value)
            converted.append(normalized)
        try:
            return pa.array(converted, type=rule.output_type, safe=True)
        except _ARROW_CONVERSION_ERRORS:
            for offset, value in enumerate(converted):
                if value is None:
                    continue
                try:
                    pa.array([value], type=rule.output_type, safe=True)
                except _ARROW_CONVERSION_ERRORS:
                    self._raise_incompatible(rule, row_base + offset, value)
            raise

    def _normalize_date_strings(
        self,
        array: pa.Array,
        rule: ColumnNormalization,
        row_base: int,
    ) -> pa.Array:
        cleaned = _normalize_arrow_strings(array, rule, self._missing_marker_arrays)
        series = cleaned.to_pandas()
        parsed = pd.to_datetime(
            series,
            errors="coerce",
            format=rule.date_format or "mixed",
        )
        invalid = parsed.isna() & series.notna()
        if bool(invalid.any()):
            offset = int(invalid.to_numpy().nonzero()[0][0])
            self._raise_incompatible(rule, row_base + offset, series.iloc[offset])
        try:
            return pa.Array.from_pandas(parsed, type=rule.output_type, safe=True)
        except _ARROW_CONVERSION_ERRORS:
            return self._normalize_scalar_column(array, rule, row_base)

    @staticmethod
    def _raise_incompatible(
        rule: ColumnNormalization,
        row_offset: int,
        value: object,
    ) -> None:
        raise StreamingTypeError(
            "streamed value is incompatible with the fixed schema",
            ordinal=rule.ordinal,
            display_label=_safe_display_label(rule.source_display_name),
            row_offset=row_offset,
            value_description=_safe_value_description(value),
            expected_type=str(rule.output_type),
        ) from None


class NormalizedStreamingReader:
    """Own a backend reader and expose only schema-stable normalized batches."""

    def __init__(
        self,
        reader: StreamingBatchReader,
        plan: NormalizationPlan,
    ) -> None:
        self._reader: StreamingBatchReader | None = None
        self._schema: pa.Schema | None = None
        self._operation: ArrowNormalizationOperation | None = None
        self._closed = False
        self._terminal = False
        self._reader = reader
        try:
            self._schema = pa.schema([])
            try:
                operation = ArrowNormalizationOperation(plan)
            except BaseException as error:
                _mark_semantic_failure(error)
                raise
            self._operation = operation
            self._schema = operation.schema
            raw_schema = reader.schema
            if not isinstance(raw_schema, pa.Schema) or not raw_schema.equals(
                operation.input_schema,
                check_metadata=True,
            ):
                raise _blocked_schema_error(
                    "normalized reader input schema does not match its compiled plan"
                )
        except BaseException as error:
            self._terminate(
                primary_error=error,
                primary_traceback=_exception_traceback(error),
            )
            raise

    @property
    def schema(self) -> pa.Schema:
        schema = self._schema
        assert schema is not None
        return schema

    def read_next_batch(self) -> pa.RecordBatch | None:
        """Return the next non-empty normalized batch or sticky EOF."""
        if self._closed or self._terminal:
            return None
        while True:
            reader = self._reader
            operation = self._operation
            assert reader is not None
            assert operation is not None
            try:
                batch = reader.read_next_batch()
            except BaseException as error:
                self._terminate(
                    primary_error=error,
                    primary_traceback=_exception_traceback(error),
                )
                raise
            if batch is None:
                self._terminate()
                return None
            try:
                result = operation.normalize(batch)
            except BaseException as error:
                _mark_semantic_failure(error)
                self._terminate(
                    primary_error=error,
                    primary_traceback=_exception_traceback(error),
                )
                raise
            if result.num_rows:
                return result

    def close(self) -> None:
        """Close the owned reader exactly once."""
        self._terminate()

    def _terminate(
        self,
        *,
        primary_error: BaseException | None = None,
        primary_traceback: TracebackType | None = None,
        cleanup_overrides: bool = False,
    ) -> None:
        if self._closed:
            return
        self._closed = True
        self._terminal = True
        reader = self._reader
        self._reader = None
        self._operation = None
        cleanups = (
            []
            if reader is None
            else [("normalized reader source cleanup", lambda: _close_owned_reader(reader))]
        )
        cleanup_primary = None if cleanup_overrides else primary_error
        cleanup_traceback = None if cleanup_primary is None else primary_traceback
        try:
            cleanup_failed = _run_cleanups(
                cleanups,
                primary_error=cleanup_primary,
                primary_traceback=cleanup_traceback,
            )
        except BaseException as cleanup_error:
            if primary_error is None:
                _mark_source_cleanup_failure(cleanup_error)
            raise
        if (
            cleanup_failed
            and primary_error is not None
            and _fallback_block_reason(primary_error) is None
        ):
            _mark_fallback_blocked(primary_error, _FallbackBlockReason.SOURCE_OWNERSHIP)

    def __enter__(self) -> NormalizedStreamingReader:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        del exc_type
        if exc_value is None:
            self.close()
            return
        self._terminate(
            primary_error=exc_value,
            primary_traceback=traceback,
            cleanup_overrides=isinstance(exc_value, GeneratorExit),
        )


def _normalize_arrow_column(
    array: pa.Array,
    rule: ColumnNormalization,
    missing_marker_arrays: dict[pa.DataType, pa.Array],
) -> pa.Array | None:
    if rule.semantic is SemanticOperation.PASSTHROUGH and array.type == rule.output_type:
        return _normalize_arrow_passthrough(array, rule)
    if rule.semantic in {SemanticOperation.TEXT, SemanticOperation.IDENTIFIER}:
        if pa.types.is_string(array.type) or pa.types.is_large_string(array.type):
            normalized = _normalize_arrow_strings(array, rule, missing_marker_arrays)
            return (
                normalized
                if normalized.type == rule.output_type
                else pc.cast(normalized, rule.output_type, safe=True)
            )
        if array.type == rule.output_type:
            return _normalize_arrow_passthrough(array, rule)
    if rule.semantic is SemanticOperation.NUMBER and (
        pa.types.is_string(array.type) or pa.types.is_large_string(array.type)
    ):
        return _normalize_arrow_number_strings(array, rule, missing_marker_arrays)
    if rule.semantic is SemanticOperation.DATE:
        if pa.types.is_integer(array.type) or pa.types.is_floating(array.type):
            return _normalize_arrow_serial_dates(array, rule)
        if rule.date_format not in {None, "mixed"} and (
            pa.types.is_string(array.type) or pa.types.is_large_string(array.type)
        ):
            return _normalize_arrow_fixed_dates(array, rule, missing_marker_arrays)
    return None


def _decode_encoded_strings(array: pa.Array) -> pa.Array:
    if isinstance(array, pa.RunEndEncodedArray):
        value_type = cast("pa.RunEndEncodedType", array.type).value_type
        if pa.types.is_dictionary(value_type):
            value_type = cast("pa.DictionaryType", value_type).value_type
        if not (pa.types.is_string(value_type) or pa.types.is_large_string(value_type)):
            return array
        array = _decode_run_end_encoded(array)
    if isinstance(array, pa.DictionaryArray):
        value_type = cast("pa.DictionaryType", array.type).value_type
        if pa.types.is_string(value_type) or pa.types.is_large_string(value_type):
            return cast("pa.Array", pc.dictionary_decode(array))
    return array


def _close_owned_reader(reader: object) -> None:
    """Close the required source-reader capability without optional semantics."""
    close = getattr(reader, "close", None)
    if not callable(close):
        raise TypeError("owned streaming reader must provide a callable close")
    close()


def _normalize_arrow_passthrough(
    array: pa.Array,
    rule: ColumnNormalization,
) -> pa.Array:
    if "missing" not in rule.enabled_stages or not pa.types.is_floating(array.type):
        return array
    nan_mask = pc.fill_null(pc.is_nan(array), False)
    if pc.any(nan_mask).as_py() is not True:
        return array
    return pc.if_else(nan_mask, pa.scalar(None, type=array.type), array)


def _normalize_arrow_strings(
    array: pa.Array,
    rule: ColumnNormalization,
    missing_marker_arrays: dict[pa.DataType, pa.Array],
) -> pa.Array:
    normalized = array
    if "whitespace" in rule.enabled_stages:
        already_normalized = pc.fill_null(
            pc.match_substring_regex(
                normalized,
                pattern=_ARROW_NORMALIZED_WHITESPACE,
            ),
            True,
        )
        if pc.all(already_normalized).as_py() is not True:
            normalized = pc.replace_substring_regex(
                normalized,
                pattern=_ARROW_ALL_SPACE,
                replacement=" ",
            )
            normalized = pc.utf8_trim_whitespace(normalized)
    if "missing" not in rule.enabled_stages:
        return normalized
    markers = missing_marker_arrays.get(normalized.type)
    if markers is None:
        markers = pa.array(sorted(rule.missing_values), type=normalized.type)
        missing_marker_arrays[normalized.type] = markers
    marker_mask = pc.fill_null(pc.is_in(normalized, value_set=markers), False)
    blank_mask = pc.fill_null(
        pc.match_substring_regex(normalized, pattern=_ARROW_BLANK),
        False,
    )
    missing_mask = pc.or_(marker_mask, blank_mask)
    if pc.any(missing_mask).as_py() is not True:
        return normalized
    return pc.if_else(
        missing_mask,
        pa.scalar(None, type=normalized.type),
        normalized,
    )


def _normalize_arrow_number_strings(
    array: pa.Array,
    rule: ColumnNormalization,
    missing_marker_arrays: dict[pa.DataType, pa.Array],
) -> pa.Array | None:
    normalized = _normalize_arrow_strings(array, rule, missing_marker_arrays)
    try:
        text = pc.replace_substring_regex(
            normalized,
            pattern=r"(?:[$€£¥₹]|CHF|kr|zł)",
            replacement="",
        )
        text = pc.utf8_trim_whitespace(text)
        text = pc.replace_substring_regex(
            text,
            pattern=r"^\((.*)\)$",
            replacement=r"-\1",
        )
        text = pc.replace_substring_regex(
            text,
            pattern=r"[ \t\x{00a0}]+",
            replacement="",
        )
        if rule.numeric_mode == "mixed_locale":
            comma_decimal = pc.fill_null(
                pc.match_substring_regex(text, pattern=r"\d,\d{2}$"),
                False,
            )
            comma_branch = pc.replace_substring(text, pattern=".", replacement="")
            comma_branch = pc.replace_substring(
                comma_branch,
                pattern=",",
                replacement=".",
            )
            dot_branch = pc.replace_substring(text, pattern=",", replacement="")
            text = pc.if_else(comma_decimal, comma_branch, dot_branch)
        else:
            if rule.thousands_separator:
                text = pc.replace_substring(
                    text,
                    pattern=rule.thousands_separator,
                    replacement="",
                )
            if rule.decimal_separator and rule.decimal_separator != ".":
                text = pc.replace_substring(
                    text,
                    pattern=rule.decimal_separator,
                    replacement=".",
                )
        converted = pc.cast(text, rule.output_type, safe=True)
        if pa.types.is_floating(rule.output_type):
            invalid = pc.and_(pc.is_valid(converted), pc.invert(pc.is_finite(converted)))
            if pc.any(invalid).as_py() is True:
                return None
        return converted
    except _ARROW_CONVERSION_ERRORS:
        return None


def _normalize_arrow_serial_dates(
    array: pa.Array,
    rule: ColumnNormalization,
) -> pa.Array | None:
    try:
        values = pc.cast(array, pa.float64(), safe=True)
        valid_serial = pc.and_kleene(
            pc.is_finite(values),
            pc.and_kleene(
                pc.and_kleene(pc.greater_equal(values, 1), pc.less_equal(values, 60_000)),
                pc.equal(values, pc.floor(values)),
            ),
        )
        invalid = pc.and_(pc.is_valid(values), pc.invert(pc.fill_null(valid_serial, False)))
        if pc.any(invalid).as_py() is True:
            return None
        seconds = pc.multiply(pc.cast(values, pa.int64()), 86_400)
        epoch_seconds = int(
            (datetime.combine(rule.date_epoch, time()) - datetime(1970, 1, 1)).total_seconds()
        )
        timestamps = pc.cast(pc.add(seconds, epoch_seconds), pa.timestamp("s"))
        return (
            timestamps
            if timestamps.type == rule.output_type
            else pc.cast(timestamps, rule.output_type, safe=True)
        )
    except _ARROW_CONVERSION_ERRORS:
        return None


def _normalize_arrow_fixed_dates(
    array: pa.Array,
    rule: ColumnNormalization,
    missing_marker_arrays: dict[pa.DataType, pa.Array],
) -> pa.Array | None:
    assert rule.date_format is not None
    try:
        cleaned = _normalize_arrow_strings(array, rule, missing_marker_arrays)
        parsed = pc.strptime(
            cleaned,
            format=rule.date_format,
            unit="us",
            error_is_null=True,
        )
        invalid = pc.and_(pc.is_valid(cleaned), pc.is_null(parsed))
        if pc.any(invalid).as_py() is True:
            return None
        return (
            parsed
            if parsed.type == rule.output_type
            else pc.cast(parsed, rule.output_type, safe=True)
        )
    except _ARROW_CONVERSION_ERRORS:
        return None


def _first_fixed_timezone_mismatch(
    array: pa.Array,
    expected_timezone: str,
) -> tuple[int, str] | None:
    match = re.fullmatch(r"([+-])(\d{2}):(\d{2})", expected_timezone)
    if match is None:
        raise RuntimeError("compiled text timezone must be a fixed offset")
    direction = 1 if match.group(1) == "+" else -1
    expected_offset = direction * (int(match.group(2)) * 3_600 + int(match.group(3)) * 60)
    for offset, scalar in enumerate(array):
        if not scalar.is_valid:
            continue
        value = cast("str", scalar.as_py())
        try:
            parsed_offset = pd.Timestamp(value).utcoffset()
        except (OverflowError, TypeError, ValueError):
            continue
        if parsed_offset is None or parsed_offset.total_seconds() != expected_offset:
            return offset, value
    return None


def _normalize_scalar(value: object, rule: ColumnNormalization) -> object | None:
    normalized = value
    if "whitespace" in rule.enabled_stages and isinstance(normalized, str):
        normalized = _normalize_whitespace(normalized)
    if "missing" in rule.enabled_stages:
        if isinstance(normalized, float) and math.isnan(normalized):
            return None
        if isinstance(normalized, str) and (
            normalized in rule.missing_values or _ALL_SPACE.fullmatch(normalized)
        ):
            return None
    if rule.semantic is SemanticOperation.NUMBER:
        return _normalize_number(normalized, rule)
    if rule.semantic is SemanticOperation.DATE:
        return _normalize_date(normalized, rule)
    if rule.semantic in {SemanticOperation.TEXT, SemanticOperation.IDENTIFIER}:
        return normalized if isinstance(normalized, str) else str(normalized)
    return normalized


def _normalize_whitespace(value: str) -> str:
    return _ALL_SPACE.sub(" ", value).strip()


def _normalize_number(value: object, rule: ColumnNormalization) -> int | float:
    if isinstance(value, bool):
        raise ValueError("booleans are not numeric normalization candidates")
    if isinstance(value, (int, float)):
        number = value
    elif isinstance(value, str):
        number = _parse_number_text(value, rule)
    else:
        raise ValueError("unsupported numeric scalar")
    return _coerce_number_to_output(number, rule.output_type)


def _parse_number_text(value: str, rule: ColumnNormalization) -> int | float:
    text = _CURRENCY.sub("", value).strip()
    accounting = _ACCOUNTING.fullmatch(text)
    if accounting is not None:
        text = f"-{accounting.group(1).strip()}"
    text = _HORIZONTAL_SPACE.sub("", text)
    if rule.numeric_mode == "mixed_locale":
        text = _normalize_mixed_number_text(text)
    else:
        text = _normalize_fixed_number_text(text, rule)
    if not pa.types.is_integer(rule.output_type):
        return float(text)
    if _INTEGER_TEXT.fullmatch(text):
        return int(text)
    parsed = float(text)
    if not parsed.is_integer():
        raise ValueError("fractional value cannot fit integer schema")
    return int(parsed)


def _normalize_mixed_number_text(text: str) -> str:
    comma_decimal = bool(re.search(r"\d,\d{2}$", text))
    dot_decimal = bool(re.search(r"\d\.\d{2}$", text))
    if comma_decimal and dot_decimal:
        comma_decimal = text.rfind(",") > text.rfind(".")
    if comma_decimal:
        return text.replace(".", "").replace(",", ".")
    return text.replace(",", "")


def _normalize_fixed_number_text(text: str, rule: ColumnNormalization) -> str:
    decimal = rule.decimal_separator or "."
    thousands = rule.thousands_separator or ","
    if thousands:
        text = text.replace(thousands, "")
    return text if decimal == "." else text.replace(decimal, ".")


def _coerce_number_to_output(
    number: int | float,
    output_type: pa.DataType,
) -> int | float:
    if pa.types.is_integer(output_type):
        if isinstance(number, float) and (not math.isfinite(number) or not number.is_integer()):
            raise ValueError("fractional value cannot fit integer schema")
        return int(number)
    floating = float(number)
    if not math.isfinite(floating):
        raise ValueError("non-finite value cannot fit numeric schema")
    return floating


def _normalize_date(value: object, rule: ColumnNormalization) -> date | datetime:
    parsed = _parse_date_value(value, rule)
    if pa.types.is_date(rule.output_type):
        return parsed.date() if isinstance(parsed, datetime) else parsed
    if isinstance(parsed, datetime):
        return parsed
    return datetime.combine(parsed, time())


def _parse_date_value(value: object, rule: ColumnNormalization) -> date | datetime:
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, bool):
        raise ValueError("booleans are not date normalization candidates")
    if isinstance(value, (int, float)):
        serial = float(value)
        if not math.isfinite(serial) or not serial.is_integer() or not 1 <= serial <= 60_000:
            raise ValueError("numeric date is outside the compiled Excel serial range")
        return datetime.combine(rule.date_epoch, time()) + timedelta(days=serial)
    if isinstance(value, str):
        converted = pd.to_datetime(
            value,
            errors="raise",
            format=rule.date_format if rule.date_format not in {None, "mixed"} else None,
        )
        if isinstance(converted, pd.Timestamp):
            return cast("datetime", converted.to_pydatetime())
        if isinstance(converted, datetime):
            return converted
        raise ValueError("unsupported parsed date")
    raise ValueError("unsupported date scalar")


def _drop_all_null_rows(batch: pa.RecordBatch) -> pa.RecordBatch:
    if batch.num_rows == 0:
        return batch
    if batch.num_columns == 0:
        return _record_batch([], batch.schema, 0)
    if any(_has_no_logical_nulls(column) for column in batch.columns):
        return batch
    keep = _logical_validity(batch.column(0))
    for ordinal in range(1, batch.num_columns):
        keep = pc.or_(keep, _logical_validity(batch.column(ordinal)))
    if pc.all(keep).as_py() is True:
        return batch
    if not any(isinstance(column, pa.RunEndEncodedArray) for column in batch.columns):
        return batch.filter(keep)
    arrays: list[pa.Array] = []
    for column in batch.columns:
        if not isinstance(column, pa.RunEndEncodedArray):
            arrays.append(cast("pa.Array", pc.filter(column, keep)))
            continue
        arrays.append(_filter_run_end_encoded(column, keep))
    return _record_batch(arrays, batch.schema, len(arrays[0]))


def _has_no_logical_nulls(array: pa.Array) -> bool:
    if isinstance(array, pa.RunEndEncodedArray):
        return _has_no_logical_nulls(array.values)
    if isinstance(array, pa.DictionaryArray):
        if array.null_count:
            return False
        if array.dictionary.null_count == 0:
            return True
        return pc.all(_dictionary_logical_validity(array)).as_py() is True
    return bool(array.null_count == 0)


def _logical_validity(array: pa.Array) -> pa.BooleanArray:
    if isinstance(array, pa.RunEndEncodedArray):
        run_validity = _logical_validity(array.values)
        encoded = pa.RunEndEncodedArray.from_arrays(array.run_ends, run_validity)
        decoded = cast("pa.BooleanArray", pc.run_end_decode(encoded))
        return decoded.slice(array.offset, len(array))
    if isinstance(array, pa.DictionaryArray):
        return _dictionary_logical_validity(array)
    return cast("pa.BooleanArray", pc.is_valid(array))


def _dictionary_logical_validity(array: pa.DictionaryArray) -> pa.BooleanArray:
    dictionary_validity = pc.is_valid(array.dictionary)
    referenced_validity = pc.take(dictionary_validity, array.indices)
    return cast("pa.BooleanArray", pc.fill_null(referenced_validity, False))


def _filter_run_end_encoded(
    array: pa.RunEndEncodedArray,
    keep: pa.BooleanArray,
) -> pa.RunEndEncodedArray:
    decoded = _decode_run_end_encoded(array)
    filtered = cast("pa.Array", pc.filter(decoded, keep))
    run_end_type = cast("pa.RunEndEncodedType", array.type).run_end_type
    if not isinstance(filtered, pa.DictionaryArray):
        return pc.run_end_encode(filtered, run_end_type=run_end_type)
    encoded_indices = pc.run_end_encode(filtered.indices, run_end_type=run_end_type)
    run_values = pa.DictionaryArray.from_arrays(
        encoded_indices.values,
        filtered.dictionary,
        ordered=cast("pa.DictionaryType", filtered.type).ordered,
    )
    return pa.RunEndEncodedArray.from_arrays(encoded_indices.run_ends, run_values)


def _decode_run_end_encoded(array: pa.RunEndEncodedArray) -> pa.Array:
    if not isinstance(array.values, pa.DictionaryArray):
        return cast("pa.Array", pc.run_end_decode(array))
    index_runs = pa.RunEndEncodedArray.from_arrays(
        array.run_ends,
        array.values.indices,
    )
    decoded_indices = cast("pa.Array", pc.run_end_decode(index_runs)).slice(
        array.offset,
        len(array),
    )
    return pa.DictionaryArray.from_arrays(
        decoded_indices,
        array.values.dictionary,
        ordered=cast("pa.DictionaryType", array.values.type).ordered,
    )


def _apply_regex_filter(
    batch: pa.RecordBatch,
    pattern: re.Pattern[str] | None,
) -> pa.RecordBatch:
    if pattern is None or batch.num_rows == 0:
        return batch
    drop = pa.repeat(pa.scalar(False), batch.num_rows)
    for column in batch.columns:
        column_drop = _arrow_regex_mask(column, pattern)
        if column_drop is None:
            column_drop = _scalar_regex_mask(column, pattern)
        drop = pc.or_(drop, column_drop)
    if pc.any(drop).as_py() is not True:
        return batch
    return batch.filter(pc.invert(drop))


def _arrow_regex_mask(
    column: pa.Array,
    pattern: re.Pattern[str],
) -> pa.BooleanArray | None:
    if not _is_arrow_safe_regex(pattern):
        return None
    try:
        text = _arrow_regex_text(column)
        if text is None:
            return None
        matches = pc.fill_null(
            pc.match_substring_regex(text, pattern=pattern.pattern),
            False,
        )
        return pc.and_(matches, pc.is_valid(column))
    except _ARROW_CONVERSION_ERRORS:
        return None


def _arrow_regex_text(column: pa.Array) -> pa.Array | None:
    value_type = column.type
    if pa.types.is_string(value_type) or pa.types.is_large_string(value_type):
        return column
    if pa.types.is_integer(value_type) or pa.types.is_decimal(value_type):
        return pc.cast(column, pa.string(), safe=True)
    return None


def _scalar_regex_mask(
    column: pa.Array,
    pattern: re.Pattern[str],
) -> pa.BooleanArray:
    if not _supports_scalar_regex_text(column.type):
        return cast("pa.BooleanArray", pa.repeat(pa.scalar(False), len(column)))
    drop: list[bool] = []
    for scalar in column:
        if not scalar.is_valid:
            drop.append(False)
            continue
        value = scalar.as_py()
        drop.append(
            not (isinstance(value, float) and math.isnan(value))
            and pattern.search(_regex_text(value)) is not None
        )
    return pa.array(drop, type=pa.bool_())


def _supports_scalar_regex_text(value_type: pa.DataType) -> bool:
    while pa.types.is_dictionary(value_type) or pa.types.is_run_end_encoded(value_type):
        value_type = value_type.value_type
    return bool(
        pa.types.is_string(value_type)
        or pa.types.is_large_string(value_type)
        or pa.types.is_binary(value_type)
        or pa.types.is_large_binary(value_type)
        or pa.types.is_fixed_size_binary(value_type)
        or pa.types.is_boolean(value_type)
        or pa.types.is_integer(value_type)
        or pa.types.is_floating(value_type)
        or pa.types.is_decimal(value_type)
        or pa.types.is_date(value_type)
        or pa.types.is_timestamp(value_type)
        or pa.types.is_time(value_type)
    )


def _is_arrow_safe_regex(pattern: re.Pattern[str]) -> bool:
    source = pattern.pattern
    if pattern.flags & ~re.UNICODE or not source.isascii() or "(?" in source or "$" in source:
        return False
    unsafe_tokens = (
        r"\d",
        r"\D",
        r"\s",
        r"\S",
        r"\w",
        r"\W",
        r"\b",
        r"\B",
        "(?<",
        "(?=",
        "(?!",
        "(?P",
        "(?#",
    )
    return not any(token in source for token in unsafe_tokens) and not any(
        f"\\{index}" in source for index in range(1, 10)
    )


def _apply_conditions(
    batch: pa.RecordBatch,
    plan: NormalizationPlan,
) -> pa.RecordBatch:
    result = batch
    for condition in plan.drop_conditions:
        result = _apply_condition(result, condition)
    return result


def _apply_condition(
    batch: pa.RecordBatch,
    condition: RowCondition,
) -> pa.RecordBatch:
    if condition.mode is ConditionMode.DUPLICATE_SUBSET_ERROR:
        raise ValueError("cannot reindex on an axis with duplicate labels")
    if condition.mode is ConditionMode.IGNORE or batch.num_rows == 0:
        return batch
    if condition.mode is ConditionMode.DROP_ROWS:
        ordinal = condition.ordinals[0]
        operand = condition.operands[0]
        if operand is None:
            return batch
        equal = pc.fill_null(pc.equal(batch.column(ordinal), operand), False)
        return batch if pc.any(equal).as_py() is not True else batch.filter(pc.invert(equal))

    arrays = list(batch.columns)
    changed = False
    for ordinal, operand in zip(condition.ordinals, condition.operands, strict=True):
        if operand is None:
            continue
        equal = pc.fill_null(pc.equal(arrays[ordinal], operand), False)
        if pc.any(equal).as_py() is not True:
            continue
        arrays[ordinal] = pc.if_else(
            equal,
            pa.nulls(batch.num_rows, type=arrays[ordinal].type),
            arrays[ordinal],
        )
        changed = True
    return _record_batch(arrays, batch.schema, batch.num_rows) if changed else batch


def _record_batch(
    arrays: list[pa.Array],
    schema: pa.Schema,
    row_count: int,
) -> pa.RecordBatch:
    if arrays:
        return pa.record_batch(arrays, schema=schema)
    return pa.record_batch([pa.nulls(row_count)], names=["_row_count"]).select([])


def _safe_display_label(value: object) -> str:
    value_type = type(value)
    if value_type is str:
        return f"str label(length={len(cast('str', value))})"
    if value_type is int or value_type is float or value_type is bool:
        return f"{value_type.__name__} label"
    return "non-string label"


def _safe_value_description(value: object) -> str:
    value_type = type(value)
    if value_type is str:
        return f"str(length={len(cast('str', value))})"
    if value_type is bytes:
        return f"bytes(length={len(cast('bytes', value))})"
    if (
        value_type is int
        or value_type is float
        or value_type is bool
        or value_type is date
        or value_type is datetime
        or value_type is time
    ):
        return value_type.__name__
    return "unsupported value"


def _regex_text(value: object) -> str:
    value_type = type(value)
    if value_type is str:
        return str.__str__(value)
    if value_type is bytes:
        return bytes.__str__(value)
    if (
        value_type is int
        or value_type is float
        or value_type is bool
        or value_type is Decimal
        or value_type is date
        or value_type is datetime
        or value_type is time
    ):
        return str(value)
    return ""


def _blocked_schema_error(message: str) -> ValueError:
    return cast(
        "ValueError",
        _mark_fallback_blocked(
            ValueError(message),
            _FallbackBlockReason.CONFIGURATION,
        ),
    )


def _mark_semantic_failure(error: BaseException) -> None:
    if not _contains_process_failure(error) and _fallback_block_reason(error) is None:
        _mark_fallback_blocked(error, _FallbackBlockReason.CONFIGURATION)


def _mark_source_cleanup_failure(error: BaseException) -> None:
    if not _contains_process_failure(error) and _fallback_block_reason(error) is None:
        _mark_fallback_blocked(error, _FallbackBlockReason.SOURCE_OWNERSHIP)


__all__ = ["ArrowNormalizationOperation", "NormalizedStreamingReader"]
