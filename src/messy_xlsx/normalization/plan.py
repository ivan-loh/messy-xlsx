"""Pure compilation of bounded evidence into immutable streaming rules."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timezone
from decimal import Decimal
from enum import StrEnum
from itertools import pairwise
from typing import Final, cast

import pandas as pd
import pyarrow as pa

from messy_xlsx.normalization.missing_values import (
    DEFAULT_MISSING_VALUES,
    EXTENDED_MISSING_VALUES,
)
from messy_xlsx.normalization.type_inference import SemanticTypeInference
from messy_xlsx.parsing.contracts import OutputMode
from messy_xlsx.parsing.coordinates import ColumnIdentity
from messy_xlsx.parsing.parse_plan import ParsePlan
from messy_xlsx.utils import sanitize_column_name

MAX_SAMPLE_VALUES: Final = 1_000
MAX_SAMPLE_CELLS: Final = 1_000_000
MAX_SAMPLE_BYTES: Final = 8 * 1024 * 1024
_DATE_SYSTEMS: Final = frozenset({"1900", "1904"})
_NUMBER_PATTERN: Final = re.compile(
    r"^[+-]?[\d,.\s\xa0]+$|^\([0-9,.\s\xa0]+\)$|^[$€£¥₹][0-9,.\s\xa0]+$"
)
_COMMA_DECIMAL_PATTERN: Final = re.compile(r"\d,\d{2}$")
_DOT_DECIMAL_PATTERN: Final = re.compile(r"\d\.\d{2}$")
_DOT_THOUSANDS_PATTERN: Final = re.compile(r"\d\.\d{3}")
_COMMA_THOUSANDS_PATTERN: Final = re.compile(r"\d,\d{3}")
_KNOWN_DATE_FORMATS: Final = (
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%Y/%m/%d",
    "%d-%m-%Y",
    "%m-%d-%Y",
    "%d.%m.%Y",
    "%Y-%m-%d %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%B %d, %Y",
    "%b %d, %Y",
    "%d %B %Y",
    "%d %b %Y",
)


class SemanticOperation(StrEnum):
    """One fixed semantic treatment for a physical ordinal column."""

    PASSTHROUGH = "passthrough"
    TEXT = "text"
    NUMBER = "number"
    DATE = "date"
    IDENTIFIER = "identifier"


class ConditionMode(StrEnum):
    """Characterized positional behavior for one drop condition."""

    IGNORE = "ignore"
    DROP_ROWS = "drop_rows"
    MASK_ALL_DUPLICATES = "mask_all_duplicates"
    DUPLICATE_SUBSET_ERROR = "duplicate_subset_error"


@dataclass(frozen=True, slots=True)
class _LabelToken:
    """Type-tagged, hash-stable identity for an untrusted display label."""

    kind: str
    value: object


@dataclass(frozen=True, slots=True)
class NormalizationSample:
    """Bounded post-coordinate evidence used only during pure compilation."""

    schema: pa.Schema
    column_identities: tuple[ColumnIdentity, ...]
    columns: tuple[pa.Array, ...]
    row_numbers: pa.Int64Array
    date_system: str = "1900"

    def __post_init__(self) -> None:
        _validate_sample_schema(self.schema)
        _validate_sample_shape(self)
        _validate_sample_coordinates(self.row_numbers)
        _validate_sample_columns(self)
        _validate_sample_budget(self)
        _validate_date_system(self.date_system)

    @property
    def row_count(self) -> int:
        """Return retained evidence rows, including zero-column rows."""
        return len(self.row_numbers)


def _validate_sample_schema(schema: pa.Schema) -> None:
    if not isinstance(schema, pa.Schema):
        raise TypeError("sample schema must be a pyarrow.Schema")
    if schema.metadata or any(field.metadata for field in schema):
        raise ValueError("sample schema metadata is not supported")
    expected_names = [str(ordinal) for ordinal in range(len(schema))]
    if schema.names != expected_names:
        raise ValueError("sample schema fields must use ordinal names")


def _validate_sample_shape(sample: NormalizationSample) -> None:
    if not isinstance(sample.column_identities, tuple) or not isinstance(sample.columns, tuple):
        raise TypeError("sample columns and identities must be immutable tuples")
    if len(sample.schema) != len(sample.columns) or len(sample.schema) != len(
        sample.column_identities
    ):
        raise ValueError("sample schema, identities, and columns must have equal width")
    for expected, identity in enumerate(sample.column_identities):
        if type(identity) is not ColumnIdentity:
            raise TypeError("sample identities must be exact ColumnIdentity values")
        if type(identity.ordinal) is not int:
            raise TypeError("sample identity ordinal must be an exact int")
        if identity.ordinal != expected:
            raise ValueError("sample identity ordinals must match physical positions")


def _validate_sample_coordinates(row_numbers: pa.Int64Array) -> None:
    if not isinstance(row_numbers, pa.Int64Array):
        raise TypeError("sample row coordinates must be an int64 Arrow array")
    if row_numbers.null_count:
        raise ValueError("sample row coordinates cannot contain nulls")
    if len(row_numbers) > MAX_SAMPLE_VALUES:
        raise ValueError(f"sample may retain at most {MAX_SAMPLE_VALUES} rows")
    coordinates = tuple(row_numbers)
    if any(int(value.as_py()) < 1 for value in coordinates):
        raise ValueError("sample row coordinates must be positive")
    if any(int(left.as_py()) >= int(right.as_py()) for left, right in pairwise(coordinates)):
        raise ValueError("sample row coordinates must be strictly increasing")


def _validate_sample_columns(sample: NormalizationSample) -> None:
    for ordinal, (schema_field, column) in enumerate(
        zip(sample.schema, sample.columns, strict=True)
    ):
        if not isinstance(column, pa.Array):
            raise TypeError(f"sample column {ordinal} must be an Arrow array")
        if len(column) != sample.row_count:
            raise ValueError("sample column lengths must match row coordinates")
        if column.type != schema_field.type:
            raise ValueError("sample column types must match the sample schema")


def _validate_sample_budget(sample: NormalizationSample) -> None:
    cells = len(sample.columns) * sample.row_count
    if cells > MAX_SAMPLE_CELLS:
        raise ValueError(f"sample may retain at most {MAX_SAMPLE_CELLS} cells")
    seen_buffers: set[tuple[int, int]] = set()
    buffer_bytes = sum(
        _unique_physical_buffer_bytes(column, seen_buffers) for column in sample.columns
    )
    if buffer_bytes > MAX_SAMPLE_BYTES:
        raise ValueError(f"sample may retain at most {MAX_SAMPLE_BYTES} Arrow bytes")


def _unique_physical_buffer_bytes(
    array: pa.Array,
    seen: set[tuple[int, int]],
) -> int:
    total = 0
    for buffer in array.buffers():
        if buffer is None:
            continue
        while buffer.parent is not None:
            buffer = buffer.parent
        identity = (buffer.address, buffer.size)
        if identity not in seen:
            seen.add(identity)
            total += buffer.size
    if isinstance(array, pa.DictionaryArray):
        total += _unique_physical_buffer_bytes(array.dictionary, seen)
    elif isinstance(array, pa.ExtensionArray):
        total += _unique_physical_buffer_bytes(array.storage, seen)
    elif (
        pa.types.is_list(array.type)
        or pa.types.is_large_list(array.type)
        or pa.types.is_fixed_size_list(array.type)
        or pa.types.is_list_view(array.type)
        or pa.types.is_large_list_view(array.type)
        or pa.types.is_map(array.type)
    ):
        total += _unique_physical_buffer_bytes(array.values, seen)
    elif pa.types.is_struct(array.type) or pa.types.is_union(array.type):
        total += sum(
            _unique_physical_buffer_bytes(array.field(index), seen)
            for index in range(array.type.num_fields)
        )
    elif pa.types.is_run_end_encoded(array.type):
        total += _unique_physical_buffer_bytes(array.run_ends, seen)
        total += _unique_physical_buffer_bytes(array.values, seen)
    return total


def _validate_date_system(date_system: str) -> None:
    if type(date_system) is not str or date_system not in _DATE_SYSTEMS:
        raise ValueError("sample date_system must be '1900' or '1904'")


@dataclass(frozen=True, slots=True)
class ColumnNormalization:
    """All immutable normalization decisions for one ordinal."""

    ordinal: int
    input_type: pa.DataType
    output_type: pa.DataType
    source_display_name: object = field(compare=False, hash=False)
    final_display_name: object = field(compare=False, hash=False)
    source_label_token: _LabelToken
    final_label_token: _LabelToken
    semantic: SemanticOperation
    explicit_hint: str | None
    enabled_stages: tuple[str, ...]
    decimal_separator: str | None
    thousands_separator: str | None
    missing_values: frozenset[str]
    preserve_types: bool
    date_system: str
    date_epoch: date
    numeric_mode: str | None = None
    date_format: str | None = None
    timezone: str | None = None


@dataclass(frozen=True, slots=True)
class RowCondition:
    """One condition resolved to final-label ordinals at compile time."""

    mode: ConditionMode
    ordinals: tuple[int, ...]
    operands: tuple[pa.Scalar | None, ...]


@dataclass(frozen=True, slots=True)
class NormalizationPlan:
    """Stable schemas and ordinal rules for one streaming operation."""

    input_schema: pa.Schema
    schema: pa.Schema
    source_display_names: tuple[object, ...] = field(compare=False, hash=False)
    final_display_names: tuple[object, ...] = field(compare=False, hash=False)
    source_label_tokens: tuple[_LabelToken, ...] = field(repr=False)
    final_label_tokens: tuple[_LabelToken, ...] = field(repr=False)
    columns: tuple[ColumnNormalization, ...]
    normalize: bool
    max_input_rows: int
    drop_regex: re.Pattern[str] | None
    drop_conditions: tuple[RowCondition, ...]


@dataclass(frozen=True, slots=True)
class _ColumnDecision:
    semantic: SemanticOperation
    output_type: pa.DataType
    decimal_separator: str | None
    thousands_separator: str | None
    numeric_mode: str | None = None
    date_format: str | None = None
    timezone: str | None = None


def compile_normalization_plan(
    sample: NormalizationSample,
    plan: ParsePlan,
) -> NormalizationPlan:
    """Compile bounded evidence without I/O or retained sample values."""
    if not isinstance(sample, NormalizationSample):
        raise TypeError("sample must be a NormalizationSample")
    if not isinstance(plan, ParsePlan):
        raise TypeError("plan must be a ParsePlan")
    if plan.output_mode is not OutputMode.STREAMING:
        raise ValueError("normalization plans require streaming output")
    if plan.batch_size is None or plan.batch_size < 1:
        raise ValueError("streaming normalization requires a positive batch_size")

    source_label_tokens = tuple(
        _display_label_token(identity.display_name) for identity in sample.column_identities
    )
    source_names = tuple(
        _snapshot_display_name(identity.display_name) for identity in sample.column_identities
    )
    final_names, final_label_tokens = _compile_final_names(
        source_names,
        source_label_tokens,
        plan,
    )
    hints = plan.thaw_type_hints()
    enabled_stages = tuple(
        stage
        for stage in ("whitespace", "numbers", "dates", "missing", "type_coercion")
        if stage not in plan.skip_normalization_steps
    )
    missing_values = frozenset(
        (
            *DEFAULT_MISSING_VALUES,
            *(EXTENDED_MISSING_VALUES if plan.use_extended_missing_list else ()),
        )
    )
    rules: list[ColumnNormalization] = []
    fields: list[pa.Field] = []
    for ordinal, (schema_field, values, source_name, final_name) in enumerate(
        zip(
            sample.schema,
            sample.columns,
            source_names,
            final_names,
            strict=True,
        )
    ):
        hint_value = _mapping_get(hints, source_label_tokens[ordinal])
        explicit_hint = _validated_hint(hint_value)
        decision = _compile_column_decision(
            schema_field.type,
            values,
            source_name,
            explicit_hint,
            normalize=plan.normalize,
            enabled_stages=enabled_stages,
            missing_values=missing_values,
            decimal_separator=plan.decimal_separator,
            thousands_separator=plan.thousands_separator,
        )
        fields.append(pa.field(str(ordinal), decision.output_type))
        rules.append(
            ColumnNormalization(
                ordinal=ordinal,
                input_type=schema_field.type,
                output_type=decision.output_type,
                source_display_name=source_name,
                final_display_name=final_name,
                source_label_token=source_label_tokens[ordinal],
                final_label_token=final_label_tokens[ordinal],
                semantic=decision.semantic,
                explicit_hint=explicit_hint,
                enabled_stages=enabled_stages if plan.normalize else (),
                decimal_separator=decision.decimal_separator,
                thousands_separator=decision.thousands_separator,
                missing_values=missing_values,
                preserve_types=plan.preserve_types,
                date_system=sample.date_system,
                date_epoch=date(1899, 12, 30),
                numeric_mode=decision.numeric_mode,
                date_format=decision.date_format,
                timezone=decision.timezone,
            )
        )
    drop_regex = (
        re.compile(plan.drop_regex) if plan.normalize and plan.drop_regex is not None else None
    )
    return NormalizationPlan(
        input_schema=sample.schema,
        schema=sample.schema if not plan.normalize else pa.schema(fields),
        source_display_names=source_names,
        final_display_names=final_names,
        source_label_tokens=source_label_tokens,
        final_label_tokens=final_label_tokens,
        columns=tuple(rules),
        normalize=plan.normalize,
        max_input_rows=plan.batch_size or 0,
        drop_regex=drop_regex,
        drop_conditions=(
            _compile_conditions(final_label_tokens, tuple(rules), plan) if plan.normalize else ()
        ),
    )


def _compile_final_names(
    source_names: tuple[object, ...],
    source_label_tokens: tuple[_LabelToken, ...],
    plan: ParsePlan,
) -> tuple[tuple[object, ...], tuple[_LabelToken, ...]]:
    if plan.sanitize_column_names:
        seen: dict[str, int] = {}
        sanitized: list[object] = []
        for source_name in source_names:
            safe_source = (
                source_name if _is_sanitizable_display_name(source_name) else "unsafe_label"
            )
            name = sanitize_column_name(safe_source)
            occurrence = seen.get(name, 0)
            seen[name] = occurrence + 1
            sanitized.append(name if occurrence == 0 else f"{name}_{occurrence}")
        names = tuple(sanitized)
    else:
        names = source_names
    renames = plan.thaw_column_renames()
    name_tokens = (
        tuple(_display_label_token(name) for name in names)
        if plan.sanitize_column_names
        else source_label_tokens
    )
    renamed = tuple(
        _mapping_get(renames, token, name) for name, token in zip(names, name_tokens, strict=True)
    )
    return (
        tuple(_snapshot_display_name(name) for name in renamed),
        tuple(_display_label_token(name) for name in renamed),
    )


def _compile_conditions(
    final_label_tokens: tuple[_LabelToken, ...],
    rules: tuple[ColumnNormalization, ...],
    plan: ParsePlan,
) -> tuple[RowCondition, ...]:
    conditions: list[RowCondition] = []
    for raw_label, raw_value in plan.thaw_drop_conditions():
        value = _snapshot_condition_value(raw_value)
        label_token = _display_label_token(raw_label)
        ordinals = (
            ()
            if raw_label is None
            else tuple(
                ordinal
                for ordinal, final_token in enumerate(final_label_tokens)
                if final_token == label_token
            )
        )
        if not ordinals:
            mode = ConditionMode.IGNORE
        elif len(ordinals) == 1:
            mode = ConditionMode.DROP_ROWS
        elif len(ordinals) == len(final_label_tokens):
            mode = ConditionMode.MASK_ALL_DUPLICATES
        else:
            mode = ConditionMode.DUPLICATE_SUBSET_ERROR
        operands = tuple(
            _compile_condition_operand(value, rules[ordinal].output_type) for ordinal in ordinals
        )
        conditions.append(RowCondition(mode=mode, ordinals=ordinals, operands=operands))
    return tuple(conditions)


def _compile_condition_operand(
    value: object,
    output_type: pa.DataType,
) -> pa.Scalar | None:
    if value is None or pa.types.is_null(output_type):
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, tuple) and value[:1] == ("unsupported",):
        return None
    if not _condition_value_matches_type(value, output_type):
        return None
    try:
        return pa.scalar(value, type=output_type)
    except (pa.ArrowInvalid, pa.ArrowTypeError, OverflowError, TypeError, ValueError):
        return None


def _condition_value_matches_type(value: object, output_type: pa.DataType) -> bool:
    if pa.types.is_string(output_type) or pa.types.is_large_string(output_type):
        return type(value) is str
    if pa.types.is_binary(output_type) or pa.types.is_large_binary(output_type):
        return type(value) is bytes
    if pa.types.is_boolean(output_type):
        return type(value) is bool
    if pa.types.is_integer(output_type):
        return type(value) is int or (
            type(value) is float and math.isfinite(value) and value.is_integer()
        )
    if pa.types.is_floating(output_type):
        return type(value) in {int, float, Decimal}
    if pa.types.is_decimal(output_type):
        return type(value) in {int, Decimal}
    if pa.types.is_date(output_type):
        return type(value) is date
    if pa.types.is_timestamp(output_type):
        return type(value) is datetime
    if pa.types.is_time(output_type):
        return type(value) is time
    return False


def _compile_column_decision(
    observed_type: pa.DataType,
    values: pa.Array,
    source_name: object,
    explicit_hint: str | None,
    *,
    normalize: bool,
    enabled_stages: tuple[str, ...],
    missing_values: frozenset[str],
    decimal_separator: str | None,
    thousands_separator: str | None,
) -> _ColumnDecision:
    if not normalize:
        return _decision(SemanticOperation.PASSTHROUGH, observed_type)
    if explicit_hint is not None:
        return _hint_decision(
            explicit_hint,
            observed_type,
            values,
            decimal_separator,
            thousands_separator,
        )
    evidence = _normalized_sample_values(values, enabled_stages)
    non_null = tuple(value for value in evidence if value is not None)
    present = tuple(
        value
        for value in non_null
        if not (
            "missing" in enabled_stages
            and isinstance(value, str)
            and (value in missing_values or value == "")
        )
    )
    if not present:
        return _decision(SemanticOperation.PASSTHROUGH, pa.null())
    return _decision_for_observed_values(
        observed_type,
        non_null,
        present,
        source_name,
        enabled_stages,
        decimal_separator,
        thousands_separator,
    )


def _decision_for_observed_values(
    observed_type: pa.DataType,
    non_null: tuple[object, ...],
    present: tuple[object, ...],
    source_name: object,
    enabled_stages: tuple[str, ...],
    decimal_separator: str | None,
    thousands_separator: str | None,
) -> _ColumnDecision:
    inferred = SemanticTypeInference()._infer_from_name(_safe_name_text(source_name))
    if inferred == "VARCHAR":
        return _decision(SemanticOperation.IDENTIFIER, observed_type)
    if pa.types.is_timestamp(observed_type) or pa.types.is_date(observed_type):
        timezone = observed_type.tz if pa.types.is_timestamp(observed_type) else None
        return _decision(
            SemanticOperation.PASSTHROUGH,
            observed_type,
            timezone=timezone,
        )
    if pa.types.is_boolean(observed_type):
        return _decision(SemanticOperation.PASSTHROUGH, observed_type)
    if pa.types.is_integer(observed_type) or pa.types.is_floating(observed_type):
        return _numeric_observed_decision(
            observed_type,
            present,
            inferred,
            enabled_stages,
        )
    if pa.types.is_string(observed_type) or pa.types.is_large_string(observed_type):
        return _string_observed_decision(
            observed_type,
            present,
            inferred,
            enabled_stages,
            decimal_separator,
            thousands_separator,
        )
    return _decision(SemanticOperation.PASSTHROUGH, observed_type)


def _numeric_observed_decision(
    observed_type: pa.DataType,
    present: tuple[object, ...],
    inferred: str | None,
    enabled_stages: tuple[str, ...],
) -> _ColumnDecision:
    if inferred == "TIMESTAMP" and "dates" in enabled_stages:
        numeric_values = tuple(value for value in present if isinstance(value, (int, float)))
        candidates = tuple(
            value
            for value in numeric_values
            if 1 <= float(value) <= 60_000 and float(value).is_integer()
        )
        if numeric_values and len(candidates) == len(numeric_values):
            return _decision(SemanticOperation.DATE, pa.timestamp("s"))
    return _decision(SemanticOperation.PASSTHROUGH, observed_type)


def _string_observed_decision(
    observed_type: pa.DataType,
    non_null: tuple[object, ...],
    inferred: str | None,
    enabled_stages: tuple[str, ...],
    decimal_separator: str | None,
    thousands_separator: str | None,
) -> _ColumnDecision:
    strings = tuple(value for value in non_null if isinstance(value, str))
    if "numbers" in enabled_stages:
        number_decision = _numeric_decision(
            strings,
            decimal_separator,
            thousands_separator,
        )
        if number_decision is not None:
            return number_decision
    if "dates" in enabled_stages:
        date_format = _date_format(strings, inferred == "TIMESTAMP")
        if date_format is not None:
            fixed_timezone = _fixed_text_timezone(strings)
            return _decision(
                SemanticOperation.DATE,
                pa.timestamp("us", tz=fixed_timezone),
                date_format=date_format,
                timezone=fixed_timezone,
            )
    return _decision(SemanticOperation.TEXT, observed_type)


def _validated_hint(value: object | None) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError("streaming type hints must be strings")
    return value if type(value) is str else str.__str__(value)


def _hint_decision(
    hint: str,
    observed_type: pa.DataType,
    values: pa.Array,
    decimal_separator: str | None,
    thousands_separator: str | None,
) -> _ColumnDecision:
    normalized = hint.upper()
    if any(name in normalized for name in ("VARCHAR", "TEXT", "STRING", "CHAR")):
        if not (
            pa.types.is_string(observed_type)
            or pa.types.is_large_string(observed_type)
            or pa.types.is_null(observed_type)
        ) and values.null_count != len(values):
            return _decision(SemanticOperation.PASSTHROUGH, observed_type)
        return _decision(SemanticOperation.IDENTIFIER, pa.string())
    if any(name in normalized for name in ("INTEGER", "INT64", "INT32", "BIGINT")):
        return _decision(
            SemanticOperation.NUMBER,
            pa.int64(),
            decimal_separator=decimal_separator,
            thousands_separator=thousands_separator,
            numeric_mode="integer",
        )
    if any(name in normalized for name in ("DECIMAL", "NUMERIC", "FLOAT", "DOUBLE")):
        return _decision(
            SemanticOperation.NUMBER,
            pa.float64(),
            decimal_separator=decimal_separator,
            thousands_separator=thousands_separator,
            numeric_mode="float",
        )
    if "TIMESTAMP" in normalized:
        if pa.types.is_timestamp(observed_type):
            return _decision(
                SemanticOperation.PASSTHROUGH,
                observed_type,
                timezone=observed_type.tz,
            )
        fixed_timezone = _fixed_sample_timezone(values)
        return _decision(
            SemanticOperation.DATE,
            pa.timestamp("ns", tz=fixed_timezone),
            timezone=fixed_timezone,
        )
    if normalized == "DATE":
        return _decision(SemanticOperation.DATE, pa.date32())
    if normalized in {"BOOL", "BOOLEAN"}:
        return _decision(SemanticOperation.PASSTHROUGH, pa.bool_())
    raise ValueError(f"Unsupported type hint: {hint}")


def _fixed_sample_timezone(values: pa.Array) -> str | None:
    strings = tuple(
        value
        for scalar in values
        if scalar.is_valid
        for value in (scalar.as_py(),)
        if type(value) is str
    )
    return _fixed_text_timezone(strings)


def _fixed_text_timezone(strings: tuple[str, ...]) -> str | None:
    offsets: set[int] = set()
    saw_naive = False
    saw_aware = False
    for value in strings:
        try:
            parsed = pd.Timestamp(value)
        except (TypeError, ValueError):
            continue
        offset = parsed.utcoffset()
        if offset is None:
            saw_naive = True
        else:
            saw_aware = True
            offsets.add(int(offset.total_seconds()))
    if (saw_naive and saw_aware) or len(offsets) > 1:
        raise ValueError("mixed or varying timestamp timezones are not supported")
    if not offsets:
        return None
    offset_seconds = next(iter(offsets))
    if offset_seconds % 60:
        raise ValueError("sub-minute timestamp timezones are not supported")
    sign = "+" if offset_seconds >= 0 else "-"
    absolute_minutes = abs(offset_seconds) // 60
    hours, minutes = divmod(absolute_minutes, 60)
    return f"{sign}{hours:02d}:{minutes:02d}"


def _decision(
    semantic: SemanticOperation,
    output_type: pa.DataType,
    *,
    decimal_separator: str | None = None,
    thousands_separator: str | None = None,
    numeric_mode: str | None = None,
    date_format: str | None = None,
    timezone: str | None = None,
) -> _ColumnDecision:
    return _ColumnDecision(
        semantic=semantic,
        output_type=output_type,
        decimal_separator=decimal_separator,
        thousands_separator=thousands_separator,
        numeric_mode=numeric_mode,
        date_format=date_format,
        timezone=timezone,
    )


def _normalized_sample_values(
    values: pa.Array,
    enabled_stages: tuple[str, ...],
) -> tuple[object | None, ...]:
    normalized: list[object | None] = []
    for scalar in values:
        if not scalar.is_valid:
            normalized.append(None)
            continue
        value = scalar.as_py()
        if "whitespace" in enabled_stages and isinstance(value, str):
            value = re.sub(r"[\s\xa0]+", " ", value).strip()
        normalized.append(value)
    return tuple(normalized)


def _numeric_decision(
    strings: tuple[str, ...],
    configured_decimal: str | None,
    configured_thousands: str | None,
) -> _ColumnDecision | None:
    if not strings:
        return None
    matches = sum(bool(_NUMBER_PATTERN.fullmatch(value.strip())) for value in strings)
    if matches <= len(strings) * 0.5:
        return None
    if configured_decimal is None and configured_thousands is None:
        mixed = _mixed_numeric_decision(strings)
        if mixed is not None:
            return mixed
    decimal, thousands = _numeric_separators(
        strings,
        configured_decimal,
        configured_thousands,
    )
    parsed: list[float] = []
    is_float = False
    for value in strings:
        try:
            number, fractional = _parse_sample_number(value, decimal, thousands)
        except ValueError:
            return None
        parsed.append(number)
        is_float = is_float or fractional
    if not parsed:
        return None
    output_type = pa.float64() if is_float else pa.int64()
    return _decision(
        SemanticOperation.NUMBER,
        output_type,
        decimal_separator=decimal,
        thousands_separator=thousands,
        numeric_mode="float" if is_float else "integer",
    )


def _mixed_numeric_decision(strings: tuple[str, ...]) -> _ColumnDecision | None:
    has_comma_decimal = any(_COMMA_DECIMAL_PATTERN.search(value) for value in strings)
    has_dot_decimal = any(_DOT_DECIMAL_PATTERN.search(value) for value in strings)
    if not (has_comma_decimal and has_dot_decimal):
        return None
    try:
        for value in strings:
            _parse_mixed_sample_number(value)
    except ValueError:
        return None
    return _decision(
        SemanticOperation.NUMBER,
        pa.float64(),
        numeric_mode="mixed_locale",
    )


def _parse_mixed_sample_number(value: str) -> float:
    text = re.sub(r"(?:[$€£¥₹]|CHF|kr|zł)", "", value).strip()
    accounting = re.fullmatch(r"\((.*)\)", text)
    if accounting is not None:
        text = f"-{accounting.group(1).strip()}"
    text = re.sub(r"[ \t\xa0]+", "", text)
    comma_decimal = bool(_COMMA_DECIMAL_PATTERN.search(text))
    dot_decimal = bool(_DOT_DECIMAL_PATTERN.search(text))
    if comma_decimal and dot_decimal:
        comma_decimal = text.rfind(",") > text.rfind(".")
    if comma_decimal:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")
    if not text or text in {"-", "+"}:
        raise ValueError("empty number")
    return float(text)


def _numeric_separators(
    strings: tuple[str, ...],
    configured_decimal: str | None,
    configured_thousands: str | None,
) -> tuple[str, str]:
    if configured_decimal is not None or configured_thousands is not None:
        return configured_decimal or ".", configured_thousands or ""
    samples = strings[:50]
    comma_decimal = sum(bool(_COMMA_DECIMAL_PATTERN.search(value)) for value in samples)
    dot_decimal = sum(bool(_DOT_DECIMAL_PATTERN.search(value)) for value in samples)
    dot_thousands = sum(bool(_DOT_THOUSANDS_PATTERN.search(value)) for value in samples)
    comma_thousands = sum(bool(_COMMA_THOUSANDS_PATTERN.search(value)) for value in samples)
    if comma_decimal > dot_decimal and dot_thousands >= comma_thousands:
        return ",", "."
    if dot_decimal > comma_decimal and comma_thousands >= dot_thousands:
        return ".", ","
    if comma_decimal > dot_decimal and comma_thousands == 0:
        return ",", "."
    return ".", ","


def _parse_sample_number(
    value: str,
    decimal_separator: str,
    thousands_separator: str,
) -> tuple[float, bool]:
    text = re.sub(r"(?:[$€£¥₹]|CHF|kr|zł)", "", value).strip()
    accounting = re.fullmatch(r"\((.*)\)", text)
    if accounting is not None:
        text = f"-{accounting.group(1).strip()}"
    text = re.sub(r"[ \t\xa0]+", "", text)
    if thousands_separator:
        text = text.replace(thousands_separator, "")
    fractional = bool(decimal_separator and decimal_separator in text)
    if decimal_separator != ".":
        text = text.replace(decimal_separator, ".")
    if not text:
        raise ValueError("empty number")
    return float(text), fractional or "e" in text.lower()


def _date_format(strings: tuple[str, ...], name_suggests_date: bool) -> str | None:
    if not strings:
        return None
    for date_format in _KNOWN_DATE_FORMATS:
        successes = 0
        for value in strings:
            try:
                datetime.strptime(value, date_format)
            except ValueError:
                continue
            successes += 1
        if successes == len(strings):
            return date_format
    if not name_suggests_date:
        return None
    parsed = pd.to_datetime(pd.Series(strings), errors="coerce", format="mixed")
    return "mixed" if int(parsed.notna().sum()) == len(strings) else None


def _mapping_get(
    mapping: dict[object, object],
    key_token: _LabelToken,
    default: object | None = None,
) -> object | None:
    for candidate, value in mapping.items():
        if _display_label_token(candidate) == key_token:
            return value
    return default


def _snapshot_display_name(value: object) -> object:
    if value is None or type(value) in {str, int, float, bool, bytes, date}:
        return value
    if type(value) in {datetime, time}:
        temporal = cast("datetime | time", value)
        tz = temporal.tzinfo
        if tz is None or type(tz) is timezone:
            return value
        value_type = type(value)
        return f"<{value_type.__module__}.{value_type.__qualname__} label>"
    value_type = type(value)
    module = type.__getattribute__(value_type, "__module__")
    qualname = type.__getattribute__(value_type, "__qualname__")
    return f"<{module}.{qualname}>"


def _snapshot_condition_value(value: object) -> object:
    if value is None or type(value) in {
        str,
        int,
        float,
        bool,
        bytes,
        Decimal,
        date,
        datetime,
        time,
    }:
        return value
    value_type = type(value)
    module = type.__getattribute__(value_type, "__module__")
    qualname = type.__getattribute__(value_type, "__qualname__")
    return ("unsupported", module, qualname)


def _is_sanitizable_display_name(value: object) -> bool:
    if value is None or type(value) in {str, int, float, bool, bytes, date}:
        return True
    if type(value) in {datetime, time}:
        temporal = cast("datetime | time", value)
        return temporal.tzinfo is None or type(temporal.tzinfo) is timezone
    return False


def _display_label_token(value: object) -> _LabelToken:  # noqa: C901
    """Snapshot exact built-ins without invoking user equality, hash, or text hooks."""
    value_type = type(value)
    if value is None:
        return _LabelToken("none", None)
    if value_type is str:
        return _LabelToken("str", value)
    if value_type is bytes:
        return _LabelToken("bytes", value)
    if value_type is bool:
        return _LabelToken("bool", value)
    if value_type is int:
        return _LabelToken("int", value)
    if value_type is float:
        float_value = cast("float", value)
        if math.isnan(float_value):
            return _LabelToken("float", "nan")
        return _LabelToken("float", float_value.hex())
    if value_type is datetime:
        datetime_value = cast("datetime", value)
        return _LabelToken(
            "datetime",
            (
                datetime_value.year,
                datetime_value.month,
                datetime_value.day,
                datetime_value.hour,
                datetime_value.minute,
                datetime_value.second,
                datetime_value.microsecond,
                datetime_value.fold,
                _timezone_label_token(datetime_value.tzinfo),
            ),
        )
    if value_type is date:
        date_value = cast("date", value)
        return _LabelToken("date", (date_value.year, date_value.month, date_value.day))
    if value_type is time:
        time_value = cast("time", value)
        return _LabelToken(
            "time",
            (
                time_value.hour,
                time_value.minute,
                time_value.second,
                time_value.microsecond,
                time_value.fold,
                _timezone_label_token(time_value.tzinfo),
            ),
        )
    module = type.__getattribute__(value_type, "__module__")
    qualname = type.__getattribute__(value_type, "__qualname__")
    return _LabelToken("unsupported", (module, qualname))


def _timezone_label_token(value: object) -> object:
    if value is None:
        return ("naive",)
    value_type = type(value)
    if value_type is timezone:
        timezone_value = cast("timezone", value)
        offset = timezone_value.utcoffset(None)
        name = timezone_value.tzname(None)
        return (
            "timezone",
            offset.days if offset is not None else None,
            offset.seconds if offset is not None else None,
            offset.microseconds if offset is not None else None,
            name,
        )
    module = type.__getattribute__(value_type, "__module__")
    qualname = type.__getattribute__(value_type, "__qualname__")
    return ("untrusted_timezone", module, qualname)


def _safe_name_text(value: object) -> str:
    if type(value) is str:
        return value
    if type(value) in {int, float, bool}:
        return str(value)
    return ""


__all__ = [
    "MAX_SAMPLE_BYTES",
    "MAX_SAMPLE_CELLS",
    "MAX_SAMPLE_VALUES",
    "ColumnNormalization",
    "ConditionMode",
    "NormalizationPlan",
    "NormalizationSample",
    "RowCondition",
    "SemanticOperation",
    "compile_normalization_plan",
]
