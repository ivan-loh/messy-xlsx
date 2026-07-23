"""Pure compilation of bounded evidence into immutable streaming rules."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal
from enum import StrEnum
from itertools import pairwise
from typing import Final, cast

import pandas as pd
import pyarrow as pa

from messy_xlsx.normalization.encoded import encoded_logical_type
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
_MAX_LABEL_DEPTH: Final = 32
_MAX_LABEL_NODES: Final = 1_024
_INT64_MIN: Final = -(1 << 63)
_INT64_MAX: Final = (1 << 63) - 1
_UINT64_MAX: Final = (1 << 64) - 1
_DATE_SYSTEMS: Final = frozenset({"1900", "1904"})
_NUMBER_PATTERN: Final = re.compile(
    r"^[+-]?[\d,.\s\xa0]+$|^\([0-9,.\s\xa0]+\)$|^[$€£¥₹][0-9,.\s\xa0]+$"
)
_NUMERIC_CHARS_PATTERN: Final = re.compile(r"[\d,.\s\xa0]+")
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
    _reject_ambiguous_opaque_tuple_tokens(source_label_tokens)
    source_names = tuple(
        _snapshot_display_name(identity.display_name) for identity in sample.column_identities
    )
    source_name_texts = tuple(
        _safe_name_text(identity.display_name) for identity in sample.column_identities
    )
    source_names_sanitizable = tuple(
        _is_sanitizable_display_name(identity.display_name) for identity in sample.column_identities
    )
    final_names, final_label_tokens = _compile_final_names(
        source_names,
        source_names_sanitizable,
        source_label_tokens,
        plan,
    )
    _reject_ambiguous_opaque_tuple_tokens(final_label_tokens)
    hints = _tokenize_mapping(plan.thaw_type_hint_items())
    enabled_stages = tuple(
        stage
        for stage in ("whitespace", "numbers", "dates", "missing", "type_coercion")
        if stage not in plan.skip_normalization_steps
    )
    decimal_separator: str | None
    thousands_separator: str | None
    detect_mixed_locale = plan.decimal_separator is None and plan.thousands_separator is None
    if detect_mixed_locale:
        decimal_separator, thousands_separator = _detect_sample_numeric_locale(
            sample,
            enabled_stages,
        )
    else:
        decimal_separator = plan.decimal_separator
        thousands_separator = plan.thousands_separator
    missing_values = frozenset(
        (
            *DEFAULT_MISSING_VALUES,
            *(EXTENDED_MISSING_VALUES if plan.use_extended_missing_list else ()),
        )
    )
    rules: list[ColumnNormalization] = []
    fields: list[pa.Field] = []
    for ordinal, (schema_field, values, source_name, source_name_text, final_name) in enumerate(
        zip(
            sample.schema,
            sample.columns,
            source_names,
            source_name_texts,
            final_names,
            strict=True,
        )
    ):
        hint_value = hints.get(source_label_tokens[ordinal])
        explicit_hint = _validated_hint(hint_value)
        decision = _compile_column_decision(
            schema_field.type,
            values,
            source_name_text,
            explicit_hint,
            normalize=plan.normalize,
            enabled_stages=enabled_stages,
            missing_values=missing_values,
            decimal_separator=decimal_separator,
            thousands_separator=thousands_separator,
            detect_mixed_locale=detect_mixed_locale,
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
    source_names_sanitizable: tuple[bool, ...],
    source_label_tokens: tuple[_LabelToken, ...],
    plan: ParsePlan,
) -> tuple[tuple[object, ...], tuple[_LabelToken, ...]]:
    if plan.sanitize_column_names:
        seen: dict[str, int] = {}
        sanitized: list[object] = []
        for source_name, is_sanitizable in zip(
            source_names,
            source_names_sanitizable,
            strict=True,
        ):
            safe_source = source_name if is_sanitizable else "unsafe_label"
            name = sanitize_column_name(safe_source)
            occurrence = seen.get(name, 0)
            seen[name] = occurrence + 1
            sanitized.append(name if occurrence == 0 else f"{name}_{occurrence}")
        names = tuple(sanitized)
    else:
        names = source_names
    renames = _tokenize_mapping(plan.thaw_column_rename_items())
    name_tokens = (
        tuple(_display_label_token(name) for name in names)
        if plan.sanitize_column_names
        else source_label_tokens
    )
    renamed = tuple(
        renames.get(token, name) for name, token in zip(names, name_tokens, strict=True)
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
    ordinals_by_token: dict[_LabelToken, list[int]] = {}
    for ordinal, token in enumerate(final_label_tokens):
        ordinals_by_token.setdefault(token, []).append(ordinal)
    for raw_label, raw_value in plan.thaw_drop_conditions():
        value = _snapshot_condition_value(raw_value)
        label_token = _display_label_token(raw_label)
        _reject_opaque_tuple_configuration_token(label_token)
        ordinals = () if raw_label is None else tuple(ordinals_by_token.get(label_token, ()))
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
    logical_type = encoded_logical_type(output_type)
    if value is None or pa.types.is_null(logical_type):
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, tuple) and value[:1] == ("unsupported",):
        return None
    if not _condition_value_matches_type(value, logical_type):
        return None
    scalar_value = _coerce_condition_value(value, logical_type)
    try:
        return pa.scalar(scalar_value, type=logical_type)
    except (pa.ArrowInvalid, pa.ArrowTypeError, OverflowError, TypeError, ValueError):
        return None


def _condition_value_matches_type(  # noqa: C901
    value: object,
    output_type: pa.DataType,
) -> bool:
    value_type = type(value)
    if pa.types.is_string(output_type) or pa.types.is_large_string(output_type):
        return value_type is str
    if pa.types.is_binary(output_type) or pa.types.is_large_binary(output_type):
        return value_type is bytes
    if pa.types.is_boolean(output_type):
        if value_type is bool:
            return True
        if value_type is int:
            return value == 0 or value == 1
        if value_type is float:
            float_value = cast("float", value)
            return math.isfinite(float_value) and (float_value == 0.0 or float_value == 1.0)
        if value_type is Decimal:
            decimal_value = cast("Decimal", value)
            return decimal_value.is_finite() and (
                decimal_value == Decimal(0) or decimal_value == Decimal(1)
            )
        return False
    if pa.types.is_integer(output_type):
        if value_type is bool or value_type is int:
            return True
        if value_type is float:
            float_value = cast("float", value)
            return math.isfinite(float_value) and float_value.is_integer()
        if value_type is Decimal:
            decimal_value = cast("Decimal", value)
            return decimal_value.is_finite() and decimal_value == decimal_value.to_integral_value()
        return False
    if pa.types.is_floating(output_type):
        return _has_exact_finite_float_projection(value, output_type)
    if pa.types.is_decimal(output_type):
        if value_type is bool or value_type is int:
            return True
        if value_type is float:
            return math.isfinite(cast("float", value))
        if value_type is Decimal:
            return cast("Decimal", value).is_finite()
        return False
    if pa.types.is_date(output_type):
        return value_type is date
    if pa.types.is_timestamp(output_type):
        return value_type is datetime
    if pa.types.is_time(output_type):
        return value_type is time
    return False


def _coerce_condition_value(value: object, output_type: pa.DataType) -> object:
    """Project characterized pandas boolean/numeric equality into Arrow scalars."""
    value_type = type(value)
    if pa.types.is_boolean(output_type) and (
        value_type is int or value_type is float or value_type is Decimal
    ):
        return value == 1
    if pa.types.is_integer(output_type) and (
        value_type is bool or value_type is float or value_type is Decimal
    ):
        return int(cast("bool | float | Decimal", value))
    if pa.types.is_floating(output_type):
        return float(cast("bool | int | float | Decimal", value))
    if pa.types.is_decimal(output_type):
        if value_type is bool or value_type is int:
            return Decimal(int(cast("bool | int", value)))
        if value_type is float:
            return Decimal.from_float(cast("float", value))
    return value


def _has_exact_finite_float_projection(
    value: object,
    output_type: pa.DataType,
) -> bool:
    value_type = type(value)
    if value_type is bool:
        return True
    if value_type is float:
        return math.isfinite(cast("float", value))
    if value_type is int:
        try:
            projected = float(cast("int", value))
        except OverflowError:
            return False
        return math.isfinite(projected) and int(projected) == value
    if value_type is Decimal:
        decimal_value = cast("Decimal", value)
        if not decimal_value.is_finite():
            return False
        try:
            projected = float(decimal_value)
        except (OverflowError, ValueError):
            return False
        if not math.isfinite(projected):
            return False
        try:
            narrowed = pa.scalar(projected, type=output_type).as_py()
        except (pa.ArrowInvalid, pa.ArrowTypeError, OverflowError, TypeError, ValueError):
            return False
        return (
            type(narrowed) is float
            and math.isfinite(narrowed)
            and Decimal.from_float(narrowed) == decimal_value
        )
    return False


def _compile_column_decision(
    observed_type: pa.DataType,
    values: pa.Array,
    source_name_text: str,
    explicit_hint: str | None,
    *,
    normalize: bool,
    enabled_stages: tuple[str, ...],
    missing_values: frozenset[str],
    decimal_separator: str | None,
    thousands_separator: str | None,
    detect_mixed_locale: bool,
) -> _ColumnDecision:
    if not normalize:
        return _decision(SemanticOperation.PASSTHROUGH, observed_type)
    logical_observed_type = _encoded_string_value_type(observed_type)
    if explicit_hint is not None:
        return _hint_decision(
            explicit_hint,
            logical_observed_type,
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
        logical_observed_type,
        non_null,
        present,
        source_name_text,
        enabled_stages,
        decimal_separator,
        thousands_separator,
        detect_mixed_locale,
    )


def _decision_for_observed_values(
    observed_type: pa.DataType,
    non_null: tuple[object, ...],
    present: tuple[object, ...],
    source_name_text: str,
    enabled_stages: tuple[str, ...],
    decimal_separator: str | None,
    thousands_separator: str | None,
    detect_mixed_locale: bool,
) -> _ColumnDecision:
    inferred = SemanticTypeInference()._infer_from_name(source_name_text)
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
            non_null,
            inferred,
            enabled_stages,
            decimal_separator,
            thousands_separator,
            detect_mixed_locale,
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
    detect_mixed_locale: bool,
) -> _ColumnDecision:
    strings = tuple(value for value in non_null if isinstance(value, str))
    if "numbers" in enabled_stages:
        number_decision = _numeric_decision(
            strings,
            decimal_separator,
            thousands_separator,
            detect_mixed_locale=detect_mixed_locale,
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
    *,
    detect_mixed_locale: bool,
) -> _ColumnDecision | None:
    if not strings:
        return None
    matches = sum(bool(_NUMBER_PATTERN.fullmatch(value.strip())) for value in strings)
    if matches <= len(strings) * 0.5:
        return None
    if detect_mixed_locale:
        mixed = _mixed_numeric_decision(strings)
        if mixed is not None:
            return mixed
    decimal, thousands = _numeric_separators(
        strings,
        configured_decimal,
        configured_thousands,
    )
    parsed: list[int | float] = []
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
    output_type = _sampled_number_type(tuple(parsed), fractional=is_float)
    if output_type is None:
        return None
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
            if not math.isfinite(_parse_mixed_sample_number(value)):
                return None
    except (OverflowError, ValueError):
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
    return _detected_numeric_separators(strings[:50])


def _detect_sample_numeric_locale(
    sample: NormalizationSample,
    enabled_stages: tuple[str, ...],
) -> tuple[str, str]:
    samples: list[str] = []
    for column in sample.columns:
        logical_type = _encoded_string_value_type(column.type)
        if not (pa.types.is_string(logical_type) or pa.types.is_large_string(logical_type)):
            continue
        strings = tuple(
            value
            for value in _normalized_sample_values(column, enabled_stages)
            if type(value) is str
        )[:50]
        samples.extend(value for value in strings if _NUMERIC_CHARS_PATTERN.match(value))
    return _detected_numeric_separators(tuple(samples))


def _encoded_string_value_type(observed_type: pa.DataType) -> pa.DataType:
    value_type = encoded_logical_type(observed_type)
    if pa.types.is_string(value_type) or pa.types.is_large_string(value_type):
        return value_type
    return observed_type


def _detected_numeric_separators(samples: tuple[str, ...]) -> tuple[str, str]:
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
) -> tuple[int | float, bool]:
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
    fractional = fractional or "e" in text.lower()
    return (float(text) if fractional else int(text)), fractional


def _sampled_number_type(
    values: tuple[int | float, ...],
    *,
    fractional: bool,
) -> pa.DataType | None:
    if fractional:
        return pa.float64() if _all_finite_floats(values) else None
    integers = cast("tuple[int, ...]", values)
    minimum = min(integers)
    maximum = max(integers)
    if minimum >= _INT64_MIN and maximum <= _INT64_MAX:
        return pa.int64()
    if minimum >= 0 and maximum <= _UINT64_MAX:
        return pa.uint64()
    return pa.float64() if _all_finite_floats(values) else None


def _all_finite_floats(values: tuple[int | float, ...]) -> bool:
    try:
        return all(math.isfinite(float(value)) for value in values)
    except OverflowError:
        return False


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


def _tokenize_mapping(items: list[tuple[object, object]]) -> dict[_LabelToken, object]:
    """Index label-keyed configuration once while preserving its first match."""
    tokenized: dict[_LabelToken, object] = {}
    for candidate, value in items:
        candidate_token = _display_label_token(candidate)
        _reject_opaque_tuple_configuration_token(candidate_token)
        tokenized.setdefault(candidate_token, value)
    return tokenized


def _reject_ambiguous_opaque_tuple_tokens(
    tokens: tuple[_LabelToken, ...],
) -> None:
    """Reject opaque tuple identity before label-keyed configuration can alias it."""
    opaque_tuple_tokens = tuple(
        token for token in tokens if _tuple_token_contains_opaque_member(token)
    )
    if len(set(opaque_tuple_tokens)) != len(opaque_tuple_tokens):
        raise ValueError(
            "ambiguous labels contain opaque tuple members (unsupported tuple members)"
        )


def _reject_opaque_tuple_configuration_token(token: _LabelToken) -> None:
    if _tuple_token_contains_opaque_member(token):
        raise ValueError(
            "label resolution cannot target opaque tuple members (unsupported tuple members)"
        )


def _tuple_token_contains_opaque_member(token: _LabelToken) -> bool:
    if token.kind != "tuple":
        return False
    members = cast("tuple[_LabelToken, ...]", token.value)
    return any(
        member.kind == "unsupported"
        or _tuple_token_contains_opaque_member(member)
        or _temporal_token_has_unsafe_timezone(member)
        for member in members
    )


def _temporal_token_has_unsafe_timezone(token: _LabelToken) -> bool:
    if token.kind not in {"datetime", "time"}:
        return False
    temporal_parts = cast("tuple[object, ...]", token.value)
    timezone_token = cast("tuple[object, ...]", temporal_parts[-1])
    if timezone_token[:1] == ("untrusted_timezone",):
        return True
    if timezone_token[:1] != ("timezone",):
        return False
    offset_token = timezone_token[1]
    name_token = timezone_token[2]
    return (type(offset_token) is tuple and offset_token[:1] == ("unsafe_offset",)) or (
        type(name_token) is tuple and name_token[:1] == ("unsafe_name",)
    )


def _snapshot_display_name(
    value: object,
    *,
    _depth: int = 0,
    _budget: list[int] | None = None,
) -> object:
    budget = _consume_label_budget(_depth, _budget)
    value_type = type(value)
    if (
        value is None
        or value_type is str
        or value_type is int
        or value_type is float
        or value_type is bool
        or value_type is bytes
        or value_type is date
    ):
        return value
    if value_type is tuple:
        tuple_value = cast("tuple[object, ...]", value)
        return tuple(
            _snapshot_display_name(item, _depth=_depth + 1, _budget=budget) for item in tuple_value
        )
    if value_type is datetime or value_type is time:
        temporal = cast("datetime | time", value)
        tz = temporal.tzinfo
        if _timezone_label_projection(tz)[1]:
            return value
        module, qualname = _safe_type_description(value_type)
        return f"<{module}.{qualname} label>"
    module, qualname = _safe_type_description(value_type)
    return f"<{module}.{qualname}>"


def _snapshot_condition_value(value: object) -> object:
    value_type = type(value)
    if (
        value is None
        or value_type is str
        or value_type is int
        or value_type is float
        or value_type is bool
        or value_type is bytes
        or value_type is Decimal
        or value_type is date
        or value_type is datetime
        or value_type is time
    ):
        return value
    module, qualname = _safe_type_description(value_type)
    return ("unsupported", module, qualname)


def _is_sanitizable_display_name(value: object) -> bool:
    value_type = type(value)
    if (
        value is None
        or value_type is str
        or value_type is int
        or value_type is float
        or value_type is bool
        or value_type is bytes
        or value_type is date
    ):
        return True
    if value_type is datetime or value_type is time:
        temporal = cast("datetime | time", value)
        return _timezone_label_projection(temporal.tzinfo)[1]
    return False


def _display_label_token(  # noqa: C901
    value: object,
    *,
    _depth: int = 0,
    _budget: list[int] | None = None,
) -> _LabelToken:
    """Snapshot exact built-ins without invoking user equality, hash, or text hooks."""
    budget = _consume_label_budget(_depth, _budget)
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
    if value_type is tuple:
        tuple_value = cast("tuple[object, ...]", value)
        return _LabelToken(
            "tuple",
            tuple(
                _display_label_token(item, _depth=_depth + 1, _budget=budget)
                for item in tuple_value
            ),
        )
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
    module, qualname = _safe_type_description(value_type)
    return _LabelToken("unsupported", (module, qualname))


def _timezone_label_token(value: object) -> object:
    return _timezone_label_projection(value)[0]


def _timezone_label_projection(value: object) -> tuple[object, bool]:
    if value is None:
        return ("naive",), True
    value_type = type(value)
    if value_type is timezone:
        timezone_value = cast("timezone", value)
        offset = timezone.utcoffset(timezone_value, None)
        name = timezone.tzname(timezone_value, None)
        if type(offset) is timedelta:
            offset_token: object = (
                "offset",
                offset.days,
                offset.seconds,
                offset.microseconds,
            )
            offset_is_safe = True
        else:
            offset_module, offset_qualname = _safe_type_description(type(offset))
            offset_token = ("unsafe_offset", offset_module, offset_qualname)
            offset_is_safe = False
        if type(name) is str:
            name_token: object = name
            name_is_safe = True
        else:
            name_module, name_qualname = _safe_type_description(type(name))
            name_token = ("unsafe_name", name_module, name_qualname)
            name_is_safe = False
        return (
            ("timezone", offset_token, name_token),
            offset_is_safe and name_is_safe,
        )
    module, qualname = _safe_type_description(value_type)
    return ("untrusted_timezone", module, qualname), False


def _safe_type_description(value_type: type[object]) -> tuple[str, str]:
    """Copy only exact string metadata from an untrusted label type."""
    module = type.__getattribute__(value_type, "__module__")
    qualname = type.__getattribute__(value_type, "__qualname__")
    return (
        module if type(module) is str else "<unknown-module>",
        qualname if type(qualname) is str else "<unknown-type>",
    )


def _consume_label_budget(depth: int, budget: list[int] | None) -> list[int]:
    """Bound recursive tuple projection before walking caller-owned labels."""
    if budget is None:
        budget = [_MAX_LABEL_NODES]
    if depth > _MAX_LABEL_DEPTH or budget[0] < 1:
        raise ValueError("tuple label exceeds structural limits")
    budget[0] -= 1
    return budget


def _safe_name_text(value: object) -> str:  # noqa: C901
    value_type = type(value)
    if value_type is str:
        return cast("str", value)
    if value_type is bytes:
        return bytes.__str__(cast("bytes", value))
    if value_type is bool:
        return "True" if value else "False"
    if value_type is int:
        return int.__str__(cast("int", value))
    if value_type is float:
        return float.__str__(cast("float", value))
    if value_type is date:
        return date.__str__(cast("date", value))
    if value_type is datetime:
        datetime_value = cast("datetime", value)
        if not _timezone_label_projection(datetime_value.tzinfo)[1]:
            return ""
        return datetime.__str__(datetime_value)
    if value_type is time:
        time_value = cast("time", value)
        if not _timezone_label_projection(time_value.tzinfo)[1]:
            return ""
        return time.__str__(time_value)
    if value_type is tuple:
        projected = _safe_name_repr(value, _depth=0, _budget=None)
        return projected or ""
    return ""


def _safe_name_repr(  # noqa: C901
    value: object,
    *,
    _depth: int,
    _budget: list[int] | None,
) -> str | None:
    budget = _consume_label_budget(_depth, _budget)
    value_type = type(value)
    if value is None:
        return "None"
    if value_type is str:
        return str.__repr__(cast("str", value))
    if value_type is bytes:
        return bytes.__repr__(cast("bytes", value))
    if value_type is bool:
        return "True" if value else "False"
    if value_type is int:
        return int.__repr__(cast("int", value))
    if value_type is float:
        return float.__repr__(cast("float", value))
    if value_type is date:
        return date.__repr__(cast("date", value))
    if value_type is datetime:
        datetime_value = cast("datetime", value)
        if not _timezone_label_projection(datetime_value.tzinfo)[1]:
            return None
        return datetime.__repr__(datetime_value)
    if value_type is time:
        time_value = cast("time", value)
        if not _timezone_label_projection(time_value.tzinfo)[1]:
            return None
        return time.__repr__(time_value)
    if value_type is not tuple:
        return None
    members: list[str] = []
    for member in cast("tuple[object, ...]", value):
        projected = _safe_name_repr(
            member,
            _depth=_depth + 1,
            _budget=budget,
        )
        if projected is None:
            return None
        members.append(projected)
    trailing_comma = "," if len(members) == 1 else ""
    return f"({', '.join(members)}{trailing_comma})"


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
