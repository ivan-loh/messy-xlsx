"""Contracts for bounded streaming normalization and stable Arrow schemas."""

from __future__ import annotations

import gc
import inspect
import math
import re
import weakref
from dataclasses import FrozenInstanceError, dataclass
from datetime import date, datetime, time, timedelta, timezone, tzinfo
from decimal import Decimal
from typing import Any

import pyarrow as pa
import pytest

import messy_xlsx.normalization.arrow_pipeline as arrow_pipeline
import messy_xlsx.normalization.encoded as encoded
import messy_xlsx.normalization.physical_buffers as physical_buffers
import messy_xlsx.normalization.plan as normalization_plan
from messy_xlsx import SheetConfig
from messy_xlsx._fallback_signals import _fallback_block_reason, _FallbackBlockReason
from messy_xlsx.exceptions import StreamingTypeError
from messy_xlsx.normalization.arrow_pipeline import (
    ArrowNormalizationOperation,
    NormalizedStreamingReader,
)
from messy_xlsx.normalization.plan import (
    MAX_SAMPLE_BYTES,
    MAX_SAMPLE_CELLS,
    MAX_SAMPLE_VALUES,
    NormalizationSample,
    compile_normalization_plan,
)
from messy_xlsx.parsing.contracts import OutputMode
from messy_xlsx.parsing.coordinates import ColumnIdentity, CoordinateCompatibilityError
from messy_xlsx.parsing.fallback import FallbackCoordinator
from messy_xlsx.parsing.parse_plan import ParsePlan, compile_parse_plan


def _parse_plan(**overrides: Any) -> ParsePlan:
    batch_size = overrides.pop("batch_size", 2)
    values: dict[str, Any] = {
        "auto_detect": False,
        "header_rows": 0,
        "sanitize_column_names": False,
    }
    values.update(overrides)
    return compile_parse_plan(
        SheetConfig(**values),
        structure=None,
        format_type="xlsx",
        output_mode=OutputMode.STREAMING,
        batch_size=batch_size,
    )


def _sample(
    arrays: tuple[pa.Array, ...],
    labels: tuple[object, ...],
    *,
    row_numbers: pa.Int64Array | None = None,
    date_system: str = "1900",
) -> NormalizationSample:
    if row_numbers is None:
        length = len(arrays[0]) if arrays else 0
        row_numbers = pa.array(range(1, length + 1), type=pa.int64())
    schema = pa.schema([pa.field(str(ordinal), array.type) for ordinal, array in enumerate(arrays)])
    return NormalizationSample(
        schema=schema,
        column_identities=tuple(
            ColumnIdentity(ordinal, label) for ordinal, label in enumerate(labels)
        ),
        columns=arrays,
        row_numbers=row_numbers,
        date_system=date_system,
    )


def _zero_column_batch(row_count: int) -> pa.RecordBatch:
    return pa.record_batch([pa.nulls(row_count)], names=["_row_count"]).select([])


def _deep_parent_buffer_array(parent_links: int = 20_000) -> pa.Array:
    buffer = pa.allocate_buffer(8)
    for _ in range(parent_links):
        buffer = buffer.slice(0, 8)
    return pa.Array.from_buffers(pa.int64(), 1, [None, buffer])


class _Reader:
    def __init__(
        self,
        schema: pa.Schema,
        events: list[pa.RecordBatch | BaseException | None] | None = None,
        *,
        close_error: BaseException | None = None,
    ) -> None:
        self.schema = schema
        self.events = list(events or [None])
        self.close_error = close_error
        self.close_calls = 0

    def read_next_batch(self) -> pa.RecordBatch | None:
        event = self.events.pop(0) if self.events else None
        if isinstance(event, BaseException):
            raise event
        return event

    def close(self) -> None:
        self.close_calls += 1
        if self.close_error is not None:
            raise self.close_error


def test_sample_validates_ordinal_identity_width_and_bound() -> None:
    with pytest.raises(ValueError, match="ordinal names"):
        NormalizationSample(
            schema=pa.schema([("amount", pa.string())]),
            column_identities=(ColumnIdentity(0, "amount"),),
            columns=(pa.array(["1"]),),
            row_numbers=pa.array([1], type=pa.int64()),
        )

    with pytest.raises(ValueError, match="identity ordinals"):
        NormalizationSample(
            schema=pa.schema([("0", pa.string())]),
            column_identities=(ColumnIdentity(1, "amount"),),
            columns=(pa.array(["1"]),),
            row_numbers=pa.array([1], type=pa.int64()),
        )

    values = pa.nulls(MAX_SAMPLE_VALUES + 1)
    with pytest.raises(ValueError, match="at most"):
        _sample((values,), ("amount",))


def test_sample_enforces_total_cell_and_arrow_buffer_budgets() -> None:
    wide_columns = tuple(pa.nulls(MAX_SAMPLE_VALUES) for _ in range(1_001))
    wide_labels = tuple(f"column-{index}" for index in range(len(wide_columns)))
    assert len(wide_columns) * MAX_SAMPLE_VALUES > MAX_SAMPLE_CELLS
    with pytest.raises(ValueError, match="cells"):
        _sample(wide_columns, wide_labels)

    huge = pa.array([b"x" * (MAX_SAMPLE_BYTES + 1)], type=pa.binary())
    with pytest.raises(ValueError, match="bytes"):
        _sample((huge,), ("payload",))


def test_sample_requires_exact_column_identity_instances_and_integer_ordinals() -> None:
    with pytest.raises(TypeError, match="ColumnIdentity"):
        NormalizationSample(
            schema=pa.schema([("0", pa.string())]),
            column_identities=(object(),),  # type: ignore[arg-type]
            columns=(pa.array(["x"]),),
            row_numbers=pa.array([1], type=pa.int64()),
        )
    with pytest.raises(TypeError, match="ordinal"):
        NormalizationSample(
            schema=pa.schema([("0", pa.string())]),
            column_identities=(ColumnIdentity(True, "x"),),
            columns=(pa.array(["x"]),),
            row_numbers=pa.array([1], type=pa.int64()),
        )


def test_sample_accepts_zero_rows_and_zero_columns_with_explicit_coordinates() -> None:
    empty = _sample((pa.array([], type=pa.string()),), ("name",))
    zero_columns = _sample(
        (),
        (),
        row_numbers=pa.array([4, 5, 6], type=pa.int64()),
    )

    assert empty.row_count == 0
    assert zero_columns.row_count == 3


def test_compiler_is_pure_and_does_not_retain_sample_arrays() -> None:
    values = pa.array(["1", "2"])
    rows = pa.array([2, 3], type=pa.int64())
    values_ref = weakref.ref(values)
    rows_ref = weakref.ref(rows)
    sample = _sample((values,), ("amount",), row_numbers=rows)

    compiled = compile_normalization_plan(sample, _parse_plan())

    del sample, values, rows
    gc.collect()
    assert values_ref() is None
    assert rows_ref() is None
    assert compiled.input_schema == pa.schema([("0", pa.string())])
    with pytest.raises(FrozenInstanceError):
        compiled.normalize = False  # type: ignore[misc]


def test_compiler_preserves_exact_nonnullable_input_schema() -> None:
    schema = pa.schema([pa.field("0", pa.int64(), nullable=False)])
    sample = NormalizationSample(
        schema=schema,
        column_identities=(ColumnIdentity(0, "id"),),
        columns=(pa.array([1, 2], type=pa.int64()),),
        row_numbers=pa.array([1, 2], type=pa.int64()),
    )

    compiled = compile_normalization_plan(sample, _parse_plan())

    assert compiled.input_schema.equals(schema, check_metadata=True)


def test_compiler_requires_streaming_parse_plan() -> None:
    materialized = compile_parse_plan(
        SheetConfig(auto_detect=False),
        structure=None,
        format_type="xlsx",
    )
    with pytest.raises(ValueError, match="streaming"):
        compile_normalization_plan(_sample((pa.array(["1"]),), ("amount",)), materialized)


def test_compiler_resolves_duplicate_hints_and_final_names_by_ordinal() -> None:
    sample = _sample(
        (pa.array(["1", "2"]), pa.array(["x", "y"])),
        ("A B", "A-B"),
    )
    plan = _parse_plan(
        type_hints={"A B": "INTEGER"},
        sanitize_column_names=True,
        column_renames={"a_b": "primary"},
    )

    compiled = compile_normalization_plan(sample, plan)

    assert compiled.source_display_names == ("A B", "A-B")
    assert compiled.final_display_names == ("primary", "a_b_1")
    assert compiled.schema.names == ["0", "1"]
    assert compiled.schema.types == [pa.int64(), pa.string()]
    assert compiled.columns[0].explicit_hint == "INTEGER"
    assert compiled.columns[1].explicit_hint is None


def test_non_string_labels_and_sanitizer_collisions_remain_positional() -> None:
    sample = _sample(
        tuple(pa.array(["x"]) for _ in range(4)),
        (1, "1", "A B", "A-B"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(
            sanitize_column_names=True,
            column_renames={"col_1": "number"},
        ),
    )

    assert compiled.source_display_names == (1, "1", "A B", "A-B")
    assert compiled.final_display_names == ("number", "col_1_1", "a_b", "a_b_1")
    assert tuple(rule.ordinal for rule in compiled.columns) == (0, 1, 2, 3)


def test_same_label_hint_applies_to_every_duplicate_ordinal() -> None:
    sample = _sample(
        (pa.array(["1", "2"]), pa.array(["3", "4"])),
        ("Amount", "Amount"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(type_hints={"Amount": "INTEGER"}),
    )

    assert compiled.schema.types == [pa.int64(), pa.int64()]
    assert tuple(rule.explicit_hint for rule in compiled.columns) == (
        "INTEGER",
        "INTEGER",
    )


@pytest.mark.parametrize(
    ("values", "label", "expected_type"),
    [
        (("001", "002"), "value", pa.int64()),
        (("foo", "bar"), "event_date", pa.string()),
        (("2024-01-01", "2024-01-02"), "event", pa.timestamp("us")),
        (("NA", "N/A"), "value", pa.null()),
        (("x", "1"), "value", pa.string()),
    ],
)
def test_compiler_uses_bounded_values_not_only_names(
    values: tuple[str, ...],
    label: str,
    expected_type: pa.DataType,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(values),), (label,)),
        _parse_plan(),
    )

    assert compiled.schema.types == [expected_type]


@pytest.mark.parametrize("sanitize", [False, True])
@pytest.mark.parametrize(
    ("values", "label", "expected_type"),
    [
        (["Jan 1 2024", "Jan 2 2024"], "event_date", pa.timestamp("us")),
        (["Jan 1 2024", "Jan 2 2024"], b"event_date", pa.timestamp("us")),
        (
            ["Jan 1 2024", "Jan 2 2024"],
            (b"event_date", ("kind", 1, True, 1.0)),
            pa.timestamp("us"),
        ),
        (["001", "002"], b"phone", pa.string()),
        (["001", "002"], (b"phone",), pa.string()),
    ],
)
def test_safe_non_string_labels_participate_in_legacy_name_inference(
    sanitize: bool,
    values: list[str],
    label: object,
    expected_type: pa.DataType,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(values),), (label,)),
        _parse_plan(sanitize_column_names=sanitize),
    )

    assert normalization_plan._safe_name_text(label) == str(label)
    assert compiled.schema.types == [expected_type]
    assert compiled.source_label_tokens == (normalization_plan._display_label_token(label),)


def test_type_distinct_safe_name_projections_keep_distinct_label_tokens() -> None:
    values = pa.array(["Jan 1 2024", "Jan 2 2024"])
    labels = (b"event_date", "event_date", (b"event_date",))
    compiled = compile_normalization_plan(
        _sample((values, values, values), labels),
        _parse_plan(),
    )

    assert compiled.schema.types == [pa.timestamp("us")] * 3
    assert len(set(compiled.source_label_tokens)) == 3
    assert compiled.source_display_names == labels


@pytest.mark.parametrize("sanitize", [False, True])
def test_unsafe_tuple_member_is_excluded_from_name_inference_without_callbacks(
    sanitize: bool,
) -> None:
    callbacks: list[str] = []

    class UnsafeMember:
        def __repr__(self) -> str:
            callbacks.append("repr")
            raise AssertionError("unsafe tuple member must not be rendered")

        def __str__(self) -> str:
            callbacks.append("str")
            raise AssertionError("unsafe tuple member must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("format")
            raise AssertionError("unsafe tuple member must not be rendered")

    member = UnsafeMember()
    member_ref = weakref.ref(member)
    sample = _sample(
        (pa.array(["Jan 1 2024", "Jan 2 2024"]),),
        ((b"event_date", member),),
    )

    compiled = compile_normalization_plan(
        sample,
        _parse_plan(sanitize_column_names=sanitize),
    )

    assert compiled.schema.types == [pa.string()]
    assert callbacks == []
    del sample, member
    gc.collect()
    assert member_ref() is None


def test_unsafe_timezone_tuple_member_is_excluded_from_name_inference() -> None:
    callbacks: list[str] = []

    class UnsafeTimezone(tzinfo):
        def utcoffset(self, _value: datetime | None) -> timedelta:
            callbacks.append("utcoffset")
            raise AssertionError("unsafe timezone must not execute")

        def dst(self, _value: datetime | None) -> timedelta:
            callbacks.append("dst")
            raise AssertionError("unsafe timezone must not execute")

        def tzname(self, _value: datetime | None) -> str:
            callbacks.append("tzname")
            raise AssertionError("unsafe timezone must not execute")

    unsafe_timezone = UnsafeTimezone()
    unsafe_timezone_ref = weakref.ref(unsafe_timezone)
    label = (
        b"event_date",
        datetime(2024, 1, 1, tzinfo=unsafe_timezone),
    )
    sample = _sample((pa.array(["Jan 1 2024", "Jan 2 2024"]),), (label,))

    compiled = compile_normalization_plan(sample, _parse_plan())

    assert compiled.schema.types == [pa.string()]
    assert callbacks == []
    del sample, label, unsafe_timezone
    gc.collect()
    assert unsafe_timezone_ref() is None


def test_compiler_preserves_native_numeric_for_text_hint_and_native_bool() -> None:
    sample = _sample(
        (pa.array([1, 2], type=pa.int64()), pa.array([True, False])),
        ("identifier", "enabled"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(type_hints={"identifier": "VARCHAR"}),
    )

    assert compiled.schema.types == [pa.int64(), pa.bool_()]


def test_compiler_preserves_large_string_and_timezone_timestamp_types() -> None:
    timestamp_type = pa.timestamp("us", tz="Asia/Kuala_Lumpur")
    sample = _sample(
        (
            pa.array(["large"], type=pa.large_string()),
            pa.array([datetime(2024, 1, 1)], type=timestamp_type),
        ),
        ("text", "created_at"),
    )
    compiled = compile_normalization_plan(sample, _parse_plan())

    assert compiled.schema.types == [pa.large_string(), timestamp_type]


def test_compiler_fixes_numeric_locale_mode_and_date_format() -> None:
    sample = _sample(
        (
            pa.array(["1.234,56", "2.345,67"]),
            pa.array(["2024-01-01", "2024-01-02"]),
        ),
        ("amount", "event"),
    )
    compiled = compile_normalization_plan(sample, _parse_plan())

    assert compiled.columns[0].decimal_separator == ","
    assert compiled.columns[0].thousands_separator == "."
    assert compiled.columns[0].numeric_mode == "float"
    assert compiled.columns[1].date_format == "%Y-%m-%d"


def test_auto_numeric_locale_is_inferred_once_across_the_bounded_sheet_sample() -> None:
    comma_decimal = pa.array(["1.234,56", "2.345,67"])
    dot_decimal = pa.array(["1,234.56", "2,345.67"])
    compiled = compile_normalization_plan(
        _sample((comma_decimal, dot_decimal), ("comma", "dot")),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [comma_decimal, dot_decimal],
            schema=compiled.input_schema,
        )
    )

    assert tuple(rule.decimal_separator for rule in compiled.columns) == (".", ".")
    assert tuple(rule.thousands_separator for rule in compiled.columns) == (",", ",")
    assert result.column(0).to_pylist() == pytest.approx([1.23456, 2.34567])
    assert result.column(1).to_pylist() == pytest.approx([1234.56, 2345.67])


@pytest.mark.parametrize(
    ("value", "expected_type", "expected_value"),
    [
        ("9223372036854775807", pa.int64(), 9_223_372_036_854_775_807),
        ("-9223372036854775808", pa.int64(), -9_223_372_036_854_775_808),
        ("9223372036854775808", pa.uint64(), 9_223_372_036_854_775_808),
        ("18446744073709551615", pa.uint64(), 18_446_744_073_709_551_615),
        ("18446744073709551616", pa.float64(), float("18446744073709551616")),
        ("-9223372036854775809", pa.float64(), float("-9223372036854775809")),
        ("9" * 400, pa.string(), "9" * 400),
    ],
)
def test_sampled_integer_text_compiles_to_a_schema_that_can_hold_its_evidence(
    value: str,
    expected_type: pa.DataType,
    expected_value: object,
) -> None:
    values = pa.array([value])
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [expected_type]
    if pa.types.is_floating(expected_type):
        assert result.column(0).to_pylist() == pytest.approx([expected_value])
    else:
        assert result.column(0).to_pylist() == [expected_value]


def test_invalid_regex_is_rejected_during_pure_compilation() -> None:
    with pytest.raises(re.error):
        compile_normalization_plan(
            _sample((pa.array(["x"]),), ("note",)),
            _parse_plan(drop_regex="("),
        )


def test_streaming_type_error_context_is_primitive_and_fallback_blocked() -> None:
    error = StreamingTypeError(
        "streamed value is incompatible with the fixed schema",
        ordinal=2,
        display_label="amount",
        row_offset=7,
        value_description="str(length=9)",
        expected_type="double",
    )

    assert error.context == {
        "ordinal": 2,
        "display_label": "str label(length=6)",
        "row_offset": 7,
        "value_description": "str(length=9)",
        "expected_type": "double",
    }
    assert _fallback_block_reason(error) is _FallbackBlockReason.CONFIGURATION


def test_streaming_type_error_redacts_arbitrary_message_and_context_text() -> None:
    secret = "highly-sensitive-cell-content"
    error = StreamingTypeError(
        secret,
        ordinal=0,
        display_label=secret,
        row_offset=0,
        value_description=secret,
        expected_type="string",
    )

    rendered = str(error)
    assert secret not in rendered
    assert secret not in str(error.to_dict())


def test_normalize_false_preserves_values_filters_and_zero_column_rows() -> None:
    raw_sample = _sample(
        (pa.array([" 1,00 ", "NA"]),),
        ("Raw Value",),
    )
    raw_plan = compile_normalization_plan(
        raw_sample,
        _parse_plan(
            normalize=False,
            sanitize_column_names=True,
            column_renames={"raw_value": "raw"},
            drop_regex="NA",
            drop_conditions=[{"column": "raw", "value": " 1,00 "}],
        ),
    )
    zero_plan = compile_normalization_plan(
        _sample(
            (),
            (),
            row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        ),
        _parse_plan(normalize=False, batch_size=3),
    )

    raw = ArrowNormalizationOperation(raw_plan).normalize(
        pa.record_batch([pa.array([" 1,00 ", "NA"])], schema=raw_plan.input_schema)
    )
    zero = ArrowNormalizationOperation(zero_plan).normalize(_zero_column_batch(3))

    assert raw.to_pydict() == {"0": [" 1,00 ", "NA"]}
    assert raw_plan.final_display_names == ("raw",)
    assert zero.num_columns == 0
    assert zero.num_rows == 3


def test_number_whitespace_missing_and_empty_row_rules_are_fixed_by_sample() -> None:
    sample = _sample(
        (
            pa.array(["  $1,234.50 ", "(2.50)"]),
            pa.array(["a", "b"]),
            pa.nulls(2),
        ),
        ("amount", "name", "empty"),
    )
    compiled = compile_normalization_plan(sample, _parse_plan(batch_size=3))
    operation = ArrowNormalizationOperation(compiled)

    result = operation.normalize(
        pa.record_batch(
            [
                pa.array([" $1,234.50 ", "NA", "3.00"]),
                pa.array([" keep ", None, None]),
                pa.nulls(3),
            ],
            schema=compiled.input_schema,
        )
    )

    assert result.schema.equals(compiled.schema, check_metadata=True)
    assert result.to_pydict() == {
        "0": [1234.5, 3.0],
        "1": ["keep", None],
        "2": [None, None],
    }


@pytest.mark.parametrize(
    ("normalize_whitespace", "expected_text"),
    [(True, "a b"), (False, "a\u2003\u2003b")],
)
def test_unicode_whitespace_and_blank_detection_match_legacy_behavior(
    normalize_whitespace: bool,
    expected_text: str,
) -> None:
    values = pa.array(["a\u2003\u2003b", "\u2003"])
    keep = pa.array(["first", "second"])
    compiled = compile_normalization_plan(
        _sample((values, keep), ("text", "keep")),
        _parse_plan(normalize_whitespace=normalize_whitespace),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values, keep], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == [expected_text, None]
    assert result.column(1).to_pylist() == ["first", "second"]


def test_zero_column_normalized_batch_drops_every_empty_row() -> None:
    sample = _sample(
        (),
        (),
        row_numbers=pa.array([1, 2], type=pa.int64()),
    )
    compiled = compile_normalization_plan(sample, _parse_plan(normalize=True))

    result = ArrowNormalizationOperation(compiled).normalize(_zero_column_batch(2))

    assert result.num_columns == 0
    assert result.num_rows == 0
    assert result.schema.equals(compiled.schema, check_metadata=True)


def test_dense_column_skips_all_null_row_validity_mask_allocation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    batch = pa.record_batch(
        [pa.array([1, 2]), pa.array([None, 3])],
        names=["dense", "sparse"],
    )

    def unexpected_validity_mask(_array: pa.Array) -> pa.Array:
        raise AssertionError("dense metadata must bypass validity masks")

    monkeypatch.setattr(arrow_pipeline.pc, "is_valid", unexpected_validity_mask)

    assert arrow_pipeline._drop_all_null_rows(batch) is batch


def test_run_end_encoded_logical_nulls_do_not_take_the_dense_shortcut() -> None:
    mixed = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array([None, "x"]),
    )
    all_null = pa.RunEndEncodedArray.from_arrays(
        pa.array([2], type=pa.int16()),
        pa.array([None], type=pa.string()),
    )
    dense = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array(["x", "y"]),
    )
    mixed_batch = pa.record_batch([mixed], names=["value"])
    all_null_batch = pa.record_batch([all_null], names=["value"])
    dense_batch = pa.record_batch([dense], names=["value"])

    mixed_result = arrow_pipeline._drop_all_null_rows(mixed_batch)
    all_null_result = arrow_pipeline._drop_all_null_rows(all_null_batch)

    assert mixed_result.column(0).to_pylist() == ["x"]
    assert mixed_result.schema.equals(mixed_batch.schema, check_metadata=True)
    assert all_null_result.num_rows == 0
    assert all_null_result.schema.equals(all_null_batch.schema, check_metadata=True)
    assert arrow_pipeline._drop_all_null_rows(dense_batch) is dense_batch


def test_dense_million_row_encoded_columns_use_metadata_validity_fast_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    dense_dictionary = pa.DictionaryArray.from_arrays(
        pa.repeat(pa.scalar(0, type=pa.int8()), 1_000_000),
        pa.array([7], type=pa.int64()),
    )
    run_dictionary = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        pa.array([7], type=pa.int64()),
    )
    dictionary_runs = pa.RunEndEncodedArray.from_arrays(
        pa.array([1_000_000], type=pa.int32()),
        pa.array([7], type=pa.int64()),
    )
    arrays = (
        pa.RunEndEncodedArray.from_arrays(
            pa.array([1_000_000], type=pa.int32()),
            pa.array([7], type=pa.int64()),
        ),
        dense_dictionary,
        pa.RunEndEncodedArray.from_arrays(
            pa.array([1_000_000], type=pa.int32()),
            run_dictionary,
        ),
        pa.DictionaryArray.from_arrays(
            pa.repeat(pa.scalar(0, type=pa.int8()), 1_000_000),
            dictionary_runs,
        ),
    )
    validity_lengths: list[int] = []
    validity_buffer_bytes: list[int] = []
    real_validity = arrow_pipeline.encoded_logical_validity

    def counted_validity(array: pa.Array) -> pa.BooleanArray:
        validity = real_validity(array)
        validity_lengths.append(len(validity))
        validity_buffer_bytes.append(
            sum(buffer.size for buffer in validity.buffers() if buffer is not None)
        )
        return validity

    monkeypatch.setattr(arrow_pipeline, "encoded_logical_validity", counted_validity)

    for values in arrays:
        batch = pa.record_batch([values], names=["value"])

        assert arrow_pipeline._drop_all_null_rows(batch) is batch

    assert max(validity_lengths, default=0) <= 1
    assert max(validity_buffer_bytes, default=0) <= 64


@pytest.mark.parametrize(
    "operation",
    [
        "logical-view",
        "logical-validity",
        "logical-density",
        "end-to-end",
        "ree-filter",
        "ree-direct-filter",
        "ree-mask",
        "dictionary-filter",
        "dictionary-mask",
    ],
)
def test_encoded_operations_reject_oversized_retained_root_buffers(
    operation: str,
) -> None:
    root = pa.array(range(2_000_000), type=pa.int64())
    tiny_slice = root.slice(0, 1)
    run_values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        tiny_slice,
    )
    dictionary_values = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        tiny_slice,
    )
    keep = pa.array([True])
    mask = pa.array([False])
    batch = pa.record_batch([run_values], names=["value"])
    calls: dict[str, Any] = {
        "logical-view": lambda: encoded.encoded_logical_view(run_values),
        "logical-validity": lambda: encoded.encoded_logical_validity(run_values),
        "logical-density": lambda: encoded.encoded_has_no_logical_nulls(run_values),
        "end-to-end": lambda: arrow_pipeline._drop_all_null_rows(batch),
        "ree-filter": lambda: encoded.filter_encoded(run_values, keep),
        "ree-direct-filter": lambda: encoded.filter_run_end_encoded(run_values, keep),
        "ree-mask": lambda: encoded.mask_encoded(run_values, mask),
        "dictionary-filter": lambda: encoded.filter_encoded(dictionary_values, keep),
        "dictionary-mask": lambda: encoded.mask_encoded(dictionary_values, mask),
    }

    assert tiny_slice.nbytes == 8
    assert run_values.nbytes < 100
    assert dictionary_values.nbytes < 100
    with pytest.raises(ValueError, match="logical-view byte limit"):
        calls[operation]()


def test_encoded_buffer_budget_deduplicates_shared_roots_and_preserves_identity() -> None:
    shared = pa.repeat(
        pa.scalar(0, type=pa.int8()),
        encoded.MAX_ENCODED_VIEW_BYTES // 2 + 1,
    )
    values = pa.DictionaryArray.from_arrays(
        shared.slice(0, 1),
        shared.slice(1, 1),
    )
    keep = pa.array([True])
    mask = pa.array([False])

    assert encoded.encoded_has_no_logical_nulls(values)
    assert encoded.filter_encoded(values, keep) is values
    assert encoded.mask_encoded(values, mask) is values


def test_encoded_buffer_budget_counts_alternating_children_root_buffers() -> None:
    root = pa.array(range(2_000_000), type=pa.int64())
    leaf_dictionary = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        root.slice(0, 1),
    )
    nested_runs = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        leaf_dictionary,
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        nested_runs,
    )

    with pytest.raises(ValueError, match="logical-view byte limit"):
        encoded.encoded_has_no_logical_nulls(values)


def test_encoded_buffer_budget_applies_to_root_size_at_the_exact_boundary() -> None:
    accepted_root = pa.array(
        range((encoded.MAX_ENCODED_VIEW_BYTES - 8) // 8),
        type=pa.int64(),
    )
    accepted = pa.DictionaryArray.from_arrays(
        pa.repeat(pa.scalar(0, type=pa.int8()), 8),
        accepted_root,
    )
    rejected_root = pa.array(
        range(encoded.MAX_ENCODED_VIEW_BYTES // 8),
        type=pa.int64(),
    )
    rejected_slice = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        rejected_root.slice(0, 1),
    )

    assert accepted.nbytes == encoded.MAX_ENCODED_VIEW_BYTES
    assert rejected_slice.nbytes < 100
    assert encoded.encoded_logical_view(accepted).to_pylist() == [0] * 8
    with pytest.raises(ValueError, match="logical-view byte limit"):
        encoded.encoded_logical_view(rejected_slice)


def test_encoded_buffer_accounting_has_depth_and_node_traversal_guards() -> None:
    deep: pa.Array = pa.array([1], type=pa.int8())
    for _ in range(encoded.MAX_ENCODED_DEPTH + 1):
        deep = pa.DictionaryArray.from_arrays(
            pa.array([0], type=pa.int8()),
            deep,
        )
    wide = pa.StructArray.from_arrays(
        [pa.array([index], type=pa.int16()) for index in range(encoded.MAX_ENCODED_NODES + 1)],
        names=[f"field_{index}" for index in range(encoded.MAX_ENCODED_NODES + 1)],
    )
    wide_encoded = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        wide,
    )

    with pytest.raises(ValueError, match="logical-view structural limits"):
        encoded.encoded_has_no_logical_nulls(deep)
    with pytest.raises(ValueError, match="logical-view structural limits"):
        encoded.encoded_has_no_logical_nulls(wide_encoded)


def test_encoded_structural_preflight_precedes_flattened_and_bulk_child_walks(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    wide = pa.StructArray.from_arrays(
        [pa.array([index], type=pa.int16()) for index in range(4_096)],
        names=[f"field_{index}" for index in range(4_096)],
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        wide,
    )
    flattened_buffer_counts: list[int] = []
    expanded_child_counts: list[int] = []
    real_root_buffers = physical_buffers._unique_flattened_buffer_bytes
    real_children = physical_buffers._physical_child_arrays

    def counted_root_buffers(
        array: pa.Array,
        seen_buffers: set[tuple[int, int]],
        max_parent_nodes: int,
    ) -> int:
        if isinstance(array, pa.StructArray) and array.type.num_fields == 4_096:
            flattened_buffer_counts.append(len(array.buffers()))
        return real_root_buffers(array, seen_buffers, max_parent_nodes)

    def counted_children(array: pa.Array) -> tuple[pa.Array, ...]:
        children = tuple(real_children(array))
        if isinstance(array, pa.StructArray) and array.type.num_fields == 4_096:
            expanded_child_counts.append(len(children))
        return children

    monkeypatch.setattr(
        physical_buffers,
        "_unique_flattened_buffer_bytes",
        counted_root_buffers,
    )
    monkeypatch.setattr(physical_buffers, "_physical_child_arrays", counted_children)

    assert wide.type.num_buffers == 1
    with pytest.raises(ValueError, match="logical-view structural limits"):
        encoded.encoded_has_no_logical_nulls(values)
    assert (flattened_buffer_counts, expanded_child_counts) == ([], [])


@pytest.mark.parametrize("consumer", ["sample", "runtime"])
def test_base_extension_storage_preflight_precedes_flattened_buffer_enumeration(
    monkeypatch: pytest.MonkeyPatch,
    consumer: str,
) -> None:
    storage = pa.StructArray.from_arrays(
        [pa.array([index], type=pa.int16()) for index in range(1_025)],
        names=[f"field_{index}" for index in range(1_025)],
    )
    opaque_type = pa.opaque(storage.type, "messy-xlsx", "wide-struct")
    values = pa.ExtensionArray.from_storage(opaque_type, storage)
    flattened_types: list[pa.DataType] = []
    real_flattened = physical_buffers._unique_flattened_buffer_bytes

    def counted_flattened(
        array: pa.Array,
        seen_buffers: set[tuple[int, int]],
        max_parent_nodes: int,
    ) -> int:
        flattened_types.append(array.type)
        return real_flattened(array, seen_buffers, max_parent_nodes)

    monkeypatch.setattr(
        physical_buffers,
        "_unique_flattened_buffer_bytes",
        counted_flattened,
    )

    assert isinstance(opaque_type, pa.BaseExtensionType)
    assert not isinstance(opaque_type, pa.ExtensionType)
    if consumer == "sample":
        with pytest.raises(ValueError, match="physical-buffer structural limits"):
            _sample((values,), ("value",))
    else:
        with pytest.raises(ValueError, match="logical-view structural limits"):
            encoded.encoded_logical_view(values)
    assert flattened_types == []


@pytest.mark.parametrize("consumer", ["sample", "runtime"])
def test_declared_buffer_entry_budget_precedes_flattened_buffer_enumeration(
    monkeypatch: pytest.MonkeyPatch,
    consumer: str,
) -> None:
    values = pa.StructArray.from_arrays(
        [pa.array([index], type=pa.int16()) for index in range(512)],
        names=[f"field_{index}" for index in range(512)],
    )
    flattened_types: list[pa.DataType] = []
    real_flattened = physical_buffers._unique_flattened_buffer_bytes

    def counted_flattened(
        array: pa.Array,
        seen_buffers: set[tuple[int, int]],
        max_parent_nodes: int,
    ) -> int:
        flattened_types.append(array.type)
        return real_flattened(array, seen_buffers, max_parent_nodes)

    monkeypatch.setattr(
        physical_buffers,
        "_unique_flattened_buffer_bytes",
        counted_flattened,
    )

    assert values.type.num_buffers + sum(field.type.num_buffers for field in values.type) == 1_025
    if consumer == "sample":
        with pytest.raises(ValueError, match="physical-buffer structural limits"):
            _sample((values,), ("value",))
    else:
        with pytest.raises(ValueError, match="logical-view structural limits"):
            encoded.encoded_logical_view(values)
    assert flattened_types == []


@pytest.mark.parametrize(
    ("value_type", "value"),
    [(pa.string_view(), "value"), (pa.binary_view(), b"value")],
)
def test_variadic_buffer_sample_preflight_maps_structural_error_before_enumeration(
    monkeypatch: pytest.MonkeyPatch,
    value_type: pa.DataType,
    value: str | bytes,
) -> None:
    values = pa.array([value], type=value_type)
    flattened_types: list[pa.DataType] = []
    real_flattened = physical_buffers._unique_flattened_buffer_bytes

    def counted_flattened(
        array: pa.Array,
        seen_buffers: set[tuple[int, int]],
        max_parent_nodes: int,
    ) -> int:
        flattened_types.append(array.type)
        return real_flattened(array, seen_buffers, max_parent_nodes)

    monkeypatch.setattr(
        physical_buffers,
        "_unique_flattened_buffer_bytes",
        counted_flattened,
    )

    assert value_type.has_variadic_buffers
    with pytest.raises(ValueError, match="physical-buffer structural limits"):
        _sample((values,), ("value",))
    assert flattened_types == []


@pytest.mark.parametrize(
    ("value_type", "value"),
    [(pa.string_view(), "value"), (pa.binary_view(), b"value")],
)
@pytest.mark.parametrize(
    "operation",
    [
        "logical-type",
        "logical-view",
        "logical-validity",
        "logical-density",
        "generic-filter",
        "generic-mask",
        "ree-filter",
        "ree-trim",
        "normalization-runtime",
    ],
)
def test_variadic_buffer_runtime_operations_map_structural_error_before_enumeration(
    monkeypatch: pytest.MonkeyPatch,
    value_type: pa.DataType,
    value: str | bytes,
    operation: str,
) -> None:
    values = pa.array([value], type=value_type)
    runs = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        values,
    )
    keep = pa.array([True])
    mask = pa.array([False])
    batch = pa.record_batch([runs], names=["value"])
    flattened_types: list[pa.DataType] = []
    real_flattened = physical_buffers._unique_flattened_buffer_bytes
    calls: dict[str, Any] = {
        "logical-type": lambda: encoded.encoded_logical_type(value_type),
        "logical-view": lambda: encoded.encoded_logical_view(values),
        "logical-validity": lambda: encoded.encoded_logical_validity(values),
        "logical-density": lambda: encoded.encoded_has_no_logical_nulls(values),
        "generic-filter": lambda: encoded.filter_encoded(values, keep),
        "generic-mask": lambda: encoded.mask_encoded(values, mask),
        "ree-filter": lambda: encoded.filter_run_end_encoded(runs, keep),
        "ree-trim": lambda: encoded.trim_run_end_encoded(runs),
        "normalization-runtime": lambda: arrow_pipeline._drop_all_null_rows(batch),
    }

    def counted_flattened(
        array: pa.Array,
        seen_buffers: set[tuple[int, int]],
        max_parent_nodes: int,
    ) -> int:
        flattened_types.append(array.type)
        return real_flattened(array, seen_buffers, max_parent_nodes)

    monkeypatch.setattr(
        physical_buffers,
        "_unique_flattened_buffer_bytes",
        counted_flattened,
    )

    with pytest.raises(ValueError, match="logical-view structural limits"):
        calls[operation]()
    assert flattened_types == []


def test_nonvariadic_physical_buffer_operations_preserve_ordinary_behavior() -> None:
    values = pa.array([1, 2], type=pa.int64())
    keep = pa.array([True, False])
    mask = pa.array([False, True])

    assert encoded.encoded_logical_type(values.type) == pa.int64()
    assert encoded.encoded_logical_view(values) is values
    assert encoded.encoded_logical_validity(values).to_pylist() == [True, True]
    assert encoded.encoded_has_no_logical_nulls(values)
    assert encoded.filter_encoded(values, keep).to_pylist() == [1]
    assert encoded.mask_encoded(values, mask).to_pylist() == [1, None]


@pytest.mark.parametrize("consumer", ["sample", "encoded"])
def test_physical_buffer_preflight_supports_map_children(consumer: str) -> None:
    values = pa.array(
        [[("key", 1)]],
        type=pa.map_(pa.string(), pa.int64()),
    )

    if consumer == "sample":
        assert _sample((values,), ("value",)).row_count == 1
    else:
        wrapped = pa.DictionaryArray.from_arrays(
            pa.array([0], type=pa.int8()),
            values,
        )
        assert encoded.encoded_has_no_logical_nulls(wrapped)


@pytest.mark.parametrize(
    ("encoded_first", "empty"),
    [(False, False), (True, False), (False, True)],
)
def test_normalization_validates_all_encoded_columns_before_batch_fast_paths(
    encoded_first: bool,
    empty: bool,
) -> None:
    safe_encoded = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        pa.array([7], type=pa.int64()),
    )
    safe_plain = pa.array([1], type=pa.int64())
    sample_arrays = (safe_encoded, safe_plain) if encoded_first else (safe_plain, safe_encoded)
    compiled = compile_normalization_plan(
        _sample(sample_arrays, ("first", "second")),
        _parse_plan(batch_size=1),
    )
    root = pa.array(range(2_000_000), type=pa.int64())
    oversized_encoded = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        root.slice(0, 1),
    )
    runtime_plain = pa.array([], type=pa.int64()) if empty else safe_plain
    runtime_encoded = oversized_encoded.slice(0, 0) if empty else oversized_encoded
    runtime_arrays = (
        (runtime_encoded, runtime_plain) if encoded_first else (runtime_plain, runtime_encoded)
    )
    batch = pa.record_batch(list(runtime_arrays), schema=compiled.input_schema)

    with pytest.raises(ValueError, match="logical-view byte limit"):
        ArrowNormalizationOperation(compiled).normalize(batch)


@pytest.mark.parametrize(
    "operation",
    [
        "logical-view",
        "logical-validity",
        "logical-density",
        "ree-filter",
        "ree-direct-filter",
        "ree-mask",
        "ree-trim",
        "dictionary-filter",
        "dictionary-mask",
    ],
)
def test_encoded_buffer_parent_chain_is_bounded(operation: str) -> None:
    leaf = _deep_parent_buffer_array()
    run_values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        leaf,
    )
    dictionary_values = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        leaf,
    )
    keep = pa.array([True])
    mask = pa.array([False])
    calls: dict[str, Any] = {
        "logical-view": lambda: encoded.encoded_logical_view(run_values),
        "logical-validity": lambda: encoded.encoded_logical_validity(run_values),
        "logical-density": lambda: encoded.encoded_has_no_logical_nulls(run_values),
        "ree-filter": lambda: encoded.filter_encoded(run_values, keep),
        "ree-direct-filter": lambda: encoded.filter_run_end_encoded(run_values, keep),
        "ree-mask": lambda: encoded.mask_encoded(run_values, mask),
        "ree-trim": lambda: encoded.trim_run_end_encoded(run_values),
        "dictionary-filter": lambda: encoded.filter_encoded(dictionary_values, keep),
        "dictionary-mask": lambda: encoded.mask_encoded(dictionary_values, mask),
    }

    with pytest.raises(ValueError, match="logical-view structural limits"):
        calls[operation]()


@pytest.mark.parametrize(
    ("indices", "dictionary", "expected"),
    [
        ([0, 1], [None, 7], [7]),
        ([0, 0], [None], []),
        ([None, 0], [7], [7]),
    ],
)
def test_dictionary_logical_nulls_participate_in_all_null_row_filtering(
    indices: list[int | None],
    dictionary: list[int | None],
    expected: list[int],
) -> None:
    values = pa.DictionaryArray.from_arrays(
        pa.array(indices, type=pa.int8()),
        pa.array(dictionary, type=pa.int64()),
    )
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result.column(0).to_pylist() == expected
    assert result.schema.equals(batch.schema, check_metadata=True)


def test_unreferenced_null_dictionary_value_preserves_batch_identity() -> None:
    values = pa.DictionaryArray.from_arrays(
        pa.array([1, 1], type=pa.int8()),
        pa.array([None, 7], type=pa.int64()),
    )
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result is batch
    assert result.column(0).indices.buffers()[1].address == values.indices.buffers()[1].address


def test_nested_dictionary_logical_nulls_are_filtered_recursively() -> None:
    inner = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        pa.array([None, 7], type=pa.int64()),
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        inner,
    )
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result.column(0).to_pylist() == [7]
    assert result.column(0).type == values.type
    assert result.schema.equals(batch.schema, check_metadata=True)


def test_nested_dictionary_projection_decodes_only_referenced_ree_entries(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    dictionary = pa.RunEndEncodedArray.from_arrays(
        pa.array([500_000, 1_000_000], type=pa.int32()),
        pa.array([None, 7], type=pa.int64()),
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([999_999, None, 0, 999_999], type=pa.int32()),
        dictionary,
    )
    decoded_lengths: list[int] = []
    decoded_buffer_bytes: list[int] = []
    real_decode = encoded.pc.run_end_decode

    def counted_decode(array: pa.RunEndEncodedArray) -> pa.Array:
        decoded = real_decode(array)
        decoded_lengths.append(len(array))
        decoded_buffer_bytes.append(
            sum(buffer.size for buffer in decoded.buffers() if buffer is not None)
        )
        return decoded

    monkeypatch.setattr(encoded.pc, "run_end_decode", counted_decode)

    logical = encoded.encoded_logical_view(values)
    validity = encoded.encoded_logical_validity(values)

    assert logical.to_pylist() == [7, None, None, 7]
    assert validity.to_pylist() == [True, False, False, True]
    assert max(decoded_lengths, default=0) <= 2
    assert max(decoded_buffer_bytes, default=0) <= 64


def test_nested_dictionary_projection_takes_only_referenced_dictionary_entries(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    inner_indices = pa.concat_arrays(
        [
            pa.repeat(pa.scalar(0, type=pa.int8()), 500_000),
            pa.repeat(pa.scalar(1, type=pa.int8()), 500_000),
        ]
    )
    dictionary = pa.DictionaryArray.from_arrays(
        inner_indices,
        pa.array([None, 7], type=pa.int64()),
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([999_999, None, 0, 999_999], type=pa.int32()),
        dictionary,
    )
    taken_lengths: list[int] = []
    taken_buffer_bytes: list[int] = []
    real_take = encoded.pc.take

    def counted_take(array: pa.Array, indices: pa.Array) -> pa.Array:
        taken = real_take(array, indices)
        taken_lengths.append(len(taken))
        taken_buffer_bytes.append(
            sum(buffer.size for buffer in taken.buffers() if buffer is not None)
        )
        return taken

    monkeypatch.setattr(encoded.pc, "take", counted_take)

    logical = encoded.encoded_logical_view(values)
    validity = encoded.encoded_logical_validity(values)

    assert logical.to_pylist() == [7, None, None, 7]
    assert validity.to_pylist() == [True, False, False, True]
    assert max(taken_lengths, default=0) <= 4
    assert max(taken_buffer_bytes, default=0) <= 64


@pytest.mark.parametrize(
    ("indices", "expected", "unchanged"),
    [
        ([0, 1], [7], False),
        ([0, 0], [], False),
        ([1, 1], [7, 7], True),
        ([None, 1], [7], False),
    ],
)
def test_alternating_dictionary_ree_validity_is_recursive(
    indices: list[int | None],
    expected: list[int],
    unchanged: bool,
) -> None:
    leaf_dictionary = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        pa.array([None, 7], type=pa.int64()),
    )
    nested_runs = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        leaf_dictionary,
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array(indices, type=pa.int8()),
        nested_runs,
    )
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result.column(0).to_pylist() == expected
    assert result.column(0).type == values.type
    assert result.schema.equals(batch.schema, check_metadata=True)
    assert (result is batch) is unchanged


def test_nested_run_end_encoded_logical_nulls_preserve_recursive_schema() -> None:
    inner = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array([None, 7], type=pa.int64()),
    )
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        inner,
    )
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result.column(0).to_pylist() == [7]
    assert result.column(0).type == values.type
    assert result.schema.equals(batch.schema, check_metadata=True)


def test_run_end_encoded_dictionary_logical_nulls_keep_stable_schema() -> None:
    run_values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        pa.array([None, 7], type=pa.int64()),
    )
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 3], type=pa.int16()),
        run_values,
    )
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result.column(0).to_pylist() == [7, 7]
    assert result.schema.equals(batch.schema, check_metadata=True)
    assert result.column(0).type == values.type


@pytest.mark.parametrize("offset", [0, 499_999])
def test_sliced_ree_dictionary_filter_decodes_only_intersecting_runs(
    monkeypatch: pytest.MonkeyPatch,
    offset: int,
) -> None:
    run_values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        pa.array([None, 7], type=pa.int64()),
    )
    parent = pa.RunEndEncodedArray.from_arrays(
        pa.array([500_000, 1_000_000], type=pa.int32()),
        run_values,
    )
    values = parent.slice(offset, 1)
    batch = pa.record_batch([values], names=["value"])
    decoded_lengths: list[int] = []
    decoded_buffer_bytes: list[int] = []
    real_decode = arrow_pipeline.pc.run_end_decode

    def counted_decode(encoded: pa.RunEndEncodedArray) -> pa.Array:
        decoded = real_decode(encoded)
        decoded_lengths.append(len(encoded))
        decoded_buffer_bytes.append(
            sum(buffer.size for buffer in decoded.buffers() if buffer is not None)
        )
        return decoded

    monkeypatch.setattr(arrow_pipeline.pc, "run_end_decode", counted_decode)

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result.num_rows == 0
    assert decoded_lengths
    assert max(decoded_lengths) <= 1
    assert max(decoded_buffer_bytes) <= 64
    assert result.schema.equals(batch.schema, check_metadata=True)


def test_trailing_dense_ree_slice_preserves_batch_identity() -> None:
    run_values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        pa.array([None, 7], type=pa.int64()),
    )
    parent = pa.RunEndEncodedArray.from_arrays(
        pa.array([500_000, 1_000_000], type=pa.int32()),
        run_values,
    )
    values = parent.slice(999_999, 1)
    batch = pa.record_batch([values], names=["value"])

    result = arrow_pipeline._drop_all_null_rows(batch)

    assert result is batch


def test_passthrough_ree_regex_filter_preserves_encoded_schema() -> None:
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 3], type=pa.int16()),
        pa.array([1, 2], type=pa.int64()),
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_regex=r"^2$", batch_size=3),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert result.column(0).to_pylist() == [1]
    assert result.column(0).type == values.type
    assert result.schema.equals(compiled.schema, check_metadata=True)


def test_passthrough_ree_condition_uses_logical_leaf_and_encoded_filter() -> None:
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 3], type=pa.int16()),
        pa.array([1, 2], type=pa.int64()),
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(
            drop_conditions=[{"column": "value", "value": 2}],
            batch_size=3,
        ),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert compiled.drop_conditions[0].operands == (pa.scalar(2, type=pa.int64()),)
    assert result.column(0).to_pylist() == [1]
    assert result.column(0).type == values.type
    assert result.schema.equals(compiled.schema, check_metadata=True)


@pytest.mark.parametrize(
    "filter_config",
    [
        {"drop_regex": r"^999$"},
        {"drop_conditions": [{"column": "value", "value": 999}]},
    ],
)
def test_passthrough_ree_all_keep_and_empty_filters_preserve_identity(
    filter_config: dict[str, object],
) -> None:
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array([1, 2], type=pa.int64()),
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(**filter_config),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)
    empty_batch = pa.record_batch([values.slice(0, 0)], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)
    empty_result = ArrowNormalizationOperation(compiled).normalize(empty_batch)

    assert result is batch
    assert empty_result is empty_batch


@pytest.mark.parametrize(
    "filter_config",
    [
        {"drop_regex": r"^2$"},
        {"drop_conditions": [{"column": "value", "value": 2}]},
    ],
)
def test_passthrough_ree_all_drop_filters_preserve_empty_encoded_schema(
    filter_config: dict[str, object],
) -> None:
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([2], type=pa.int16()),
        pa.array([2], type=pa.int64()),
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(**filter_config),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.num_rows == 0
    assert result.column(0).type == values.type
    assert result.schema.equals(compiled.schema, check_metadata=True)


def test_passthrough_sliced_ree_filter_decodes_only_the_logical_slice(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    parent = pa.RunEndEncodedArray.from_arrays(
        pa.array([500_000, 1_000_000], type=pa.int32()),
        pa.array([1, 2], type=pa.int64()),
    )
    values = parent.slice(499_999, 2)
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_regex=r"^2$"),
    )
    decoded_lengths: list[int] = []
    decoded_buffer_bytes: list[int] = []
    real_decode = encoded.pc.run_end_decode

    def counted_decode(array: pa.RunEndEncodedArray) -> pa.Array:
        decoded = real_decode(array)
        decoded_lengths.append(len(array))
        decoded_buffer_bytes.append(
            sum(buffer.size for buffer in decoded.buffers() if buffer is not None)
        )
        return decoded

    monkeypatch.setattr(encoded.pc, "run_end_decode", counted_decode)

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == [1]
    assert result.column(0).type == values.type
    assert max(decoded_lengths, default=0) <= 2
    assert max(decoded_buffer_bytes, default=0) <= 64


@pytest.mark.parametrize("nesting", ["dictionary_ree", "ree_dictionary", "ree_ree"])
def test_passthrough_nested_encoded_conditions_preserve_schema(nesting: str) -> None:
    if nesting == "dictionary_ree":
        dictionary = pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2], type=pa.int16()),
            pa.array([1, 2], type=pa.int64()),
        )
        values: pa.Array = pa.DictionaryArray.from_arrays(
            pa.array([0, 1], type=pa.int8()),
            dictionary,
        )
    elif nesting == "ree_dictionary":
        run_values = pa.DictionaryArray.from_arrays(
            pa.array([0, 1], type=pa.int8()),
            pa.array([1, 2], type=pa.int64()),
        )
        values = pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2], type=pa.int16()),
            run_values,
        )
    else:
        inner = pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2], type=pa.int16()),
            pa.array([1, 2], type=pa.int64()),
        )
        values = pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2], type=pa.int16()),
            inner,
        )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": 2}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == [1]
    assert result.column(0).type == values.type
    assert result.schema.equals(compiled.schema, check_metadata=True)


def test_late_value_uses_global_input_offset_across_filtered_rows() -> None:
    sample = _sample(
        (pa.array(["1", "2"]), pa.array(["keep", "drop"])),
        ("amount", "note"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(drop_regex="drop", batch_size=3),
    )
    operation = ArrowNormalizationOperation(compiled)
    first = operation.normalize(
        pa.record_batch(
            [
                pa.array(["1", "2", "3"]),
                pa.array(["keep", "drop", "keep"]),
            ],
            schema=compiled.input_schema,
        )
    )

    with pytest.raises(StreamingTypeError) as captured:
        operation.normalize(
            pa.record_batch(
                [pa.array(["4", "late"]), pa.array(["keep", "keep"])],
                schema=compiled.input_schema,
            )
        )

    assert first.num_rows == 2
    assert captured.value.context["ordinal"] == 0
    assert captured.value.context["row_offset"] == 4
    assert captured.value.context["value_description"] == "str(length=4)"
    assert "late" not in str(captured.value)
    with pytest.raises(RuntimeError, match="terminal"):
        operation.normalize(
            pa.record_batch(
                [pa.array(["5"]), pa.array(["keep"])],
                schema=compiled.input_schema,
            )
        )


def test_input_schema_metadata_drift_is_rejected_before_normalization() -> None:
    sample = _sample((pa.array(["1"]),), ("amount",))
    compiled = compile_normalization_plan(sample, _parse_plan())
    drifted = pa.schema([pa.field("0", pa.string(), metadata={b"x": b"y"})])

    with pytest.raises(ValueError, match="input schema"):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch([pa.array(["1"])], schema=drifted)
        )


def test_operation_rejects_input_larger_than_compiled_batch_size() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1", "2"]),), ("amount",)),
        _parse_plan(batch_size=2),
    )

    with pytest.raises(ValueError, match="batch_size"):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch(
                [pa.array(["1", "2", "3"])],
                schema=compiled.input_schema,
            )
        )


def test_raw_numeric_dates_keep_legacy_epoch_even_for_1904_workbook() -> None:
    sample = _sample(
        (pa.array([1, 2], type=pa.int64()),),
        ("order_date",),
        date_system="1904",
    )
    compiled = compile_normalization_plan(sample, _parse_plan())

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([pa.array([1], type=pa.int64())], schema=compiled.input_schema)
    )

    assert compiled.columns[0].date_epoch.isoformat() == "1899-12-30"
    assert result.column(0)[0].as_py().date().isoformat() == "1899-12-31"


def test_already_decoded_datetime_values_are_not_re_epoched() -> None:
    values = pa.array(
        [datetime(2024, 1, 2, 3, 4, 5)],
        type=pa.timestamp("us"),
    )
    sample = _sample((values,), ("order_date",), date_system="1904")
    compiled = compile_normalization_plan(sample, _parse_plan())

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0)[0].as_py() == datetime(2024, 1, 2, 3, 4, 5)


def test_timezone_timestamp_schema_and_values_are_preserved_across_batches() -> None:
    timestamp_type = pa.timestamp("us", tz="Asia/Kuala_Lumpur")
    values = pa.array([datetime(2024, 1, 2, 3, 4, 5)], type=timestamp_type)
    compiled = compile_normalization_plan(
        _sample((values,), ("created_at",)),
        _parse_plan(batch_size=1),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.schema.types == [timestamp_type]
    assert result.column(0)[0].as_py() == values[0].as_py()


def test_all_duplicate_condition_masks_cells_but_keeps_rows() -> None:
    sample = _sample(
        (pa.array(["drop", "keep"]), pa.array(["keep", "drop"])),
        ("Status", "Status"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(
            drop_conditions=[{"column": "Status", "value": "drop"}],
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [pa.array(["drop", "keep"]), pa.array(["keep", "drop"])],
            schema=compiled.input_schema,
        )
    )

    assert result.to_pydict() == {"0": [None, "keep"], "1": ["keep", None]}


def test_duplicate_subset_condition_keeps_legacy_error() -> None:
    sample = _sample(
        (
            pa.array(["drop"]),
            pa.array(["keep"]),
            pa.array(["x"]),
        ),
        ("Status", "Status", "Other"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(
            drop_conditions=[{"column": "Status", "value": "drop"}],
        ),
    )

    with pytest.raises(ValueError, match="duplicate labels"):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch(list(sample.columns), schema=compiled.input_schema)
        )


def test_normalized_reader_validates_schema_transactionally_and_retains_schema() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    raw = _Reader(pa.schema([("0", pa.int64())]))

    with pytest.raises(ValueError, match="schema") as captured:
        NormalizedStreamingReader(raw, compiled)

    assert raw.close_calls == 1
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.CONFIGURATION


@pytest.mark.parametrize("step", ["empty_schema", "operation_schema"])
@pytest.mark.parametrize("failure_type", [OSError, MemoryError])
def test_normalized_reader_constructor_closes_every_post_ownership_failure(
    step: str,
    failure_type: type[BaseException],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    failure = failure_type(step)
    raw = _Reader(compiled.input_schema)

    def fail(*_args: object, **_kwargs: object) -> object:
        raise failure

    if step == "empty_schema":
        monkeypatch.setattr(arrow_pipeline.pa, "schema", fail)
    else:
        monkeypatch.setattr(ArrowNormalizationOperation, "schema", property(fail))

    with pytest.raises(failure_type) as captured:
        NormalizedStreamingReader(raw, compiled)

    assert captured.value is failure
    assert raw.close_calls == 1


@pytest.mark.parametrize(
    ("primary_type", "cleanup_type", "cleanup_wins"),
    [
        (OSError, RuntimeError, False),
        (OSError, MemoryError, True),
        (MemoryError, OSError, False),
    ],
)
def test_normalized_reader_constructor_cleanup_preserves_failure_precedence(
    primary_type: type[BaseException],
    cleanup_type: type[BaseException],
    cleanup_wins: bool,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    primary = primary_type("operation schema")
    cleanup = cleanup_type("cleanup")
    raw = _Reader(compiled.input_schema, close_error=cleanup)

    def fail_schema(_operation: object) -> object:
        raise primary

    monkeypatch.setattr(
        ArrowNormalizationOperation,
        "schema",
        property(fail_schema),
    )

    expected = cleanup if cleanup_wins else primary
    with pytest.raises(type(expected)) as captured:
        NormalizedStreamingReader(raw, compiled)

    assert captured.value is expected
    if cleanup_wins:
        assert cleanup.backend_context["operation_failure"] == {"type": primary_type.__name__}
    else:
        assert primary.backend_context["cleanup_failure"] == {"type": cleanup_type.__name__}
    assert raw.close_calls == 1


def test_normalized_reader_suppresses_empty_outputs_and_closes_at_sticky_eof() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1", "2"]),), ("amount",)),
        _parse_plan(),
    )
    empty = pa.record_batch([pa.array([None], type=pa.string())], schema=compiled.input_schema)
    kept = pa.record_batch([pa.array(["3"])], schema=compiled.input_schema)
    raw = _Reader(compiled.input_schema, [empty, kept, None])
    reader = NormalizedStreamingReader(raw, compiled)

    result = reader.read_next_batch()

    assert result is not None
    assert result.to_pydict() == {"0": [3]}
    assert reader.schema.equals(compiled.schema, check_metadata=True)
    assert reader.read_next_batch() is None
    assert raw.close_calls == 1
    assert reader.read_next_batch() is None
    reader.close()
    assert raw.close_calls == 1
    assert reader.schema.equals(compiled.schema, check_metadata=True)


def test_normalized_reader_failure_closes_failing_batch_and_becomes_sticky() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1", "2"]),), ("amount",)),
        _parse_plan(),
    )
    first = pa.record_batch([pa.array(["3"])], schema=compiled.input_schema)
    failing = pa.record_batch([pa.array(["bad"])], schema=compiled.input_schema)
    raw = _Reader(compiled.input_schema, [first, failing])
    reader = NormalizedStreamingReader(raw, compiled)

    assert reader.read_next_batch() is not None
    with pytest.raises(StreamingTypeError):
        reader.read_next_batch()

    assert raw.close_calls == 1
    assert reader.read_next_batch() is None


def test_normalized_reader_cleanup_process_failure_wins_over_semantic_error() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    failing = pa.record_batch([pa.array(["bad"])], schema=compiled.input_schema)
    cleanup = MemoryError("cleanup")
    raw = _Reader(compiled.input_schema, [failing], close_error=cleanup)
    reader = NormalizedStreamingReader(raw, compiled)

    with pytest.raises(MemoryError) as captured:
        reader.read_next_batch()

    assert captured.value is cleanup
    assert captured.value.backend_context["operation_failure"] == {"type": "StreamingTypeError"}


def test_ordinary_cleanup_failure_stays_attached_to_semantic_error() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    failing = pa.record_batch([pa.array(["bad"])], schema=compiled.input_schema)
    raw = _Reader(compiled.input_schema, [failing], close_error=OSError("cleanup"))
    reader = NormalizedStreamingReader(raw, compiled)

    with pytest.raises(StreamingTypeError) as captured:
        reader.read_next_batch()

    assert captured.value.backend_context["cleanup_failure"] == {"type": "OSError"}


def test_raw_compatibility_failure_retries_only_after_clean_internal_cleanup() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        output = pa.record_batch([pa.array([1.0])], schema=compiled.schema)
        return _Reader(compiled.schema, [output, None])

    primary = _Reader(
        compiled.input_schema,
        [CoordinateCompatibilityError("retry")],
    )
    result = list(
        FallbackCoordinator(lambda error: isinstance(error, CoordinateCompatibilityError)).batches(
            lambda: NormalizedStreamingReader(primary, compiled),
            fallback_factory,
        )
    )

    assert len(result) == 1
    assert fallback_calls == 1
    assert primary.close_calls == 1


def test_internal_cleanup_failure_blocks_otherwise_retryable_fallback() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    primary_error = CoordinateCompatibilityError("retry")
    primary = _Reader(
        compiled.input_schema,
        [primary_error],
        close_error=OSError("cleanup"),
    )
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(compiled.schema)

    with pytest.raises(CoordinateCompatibilityError) as captured:
        list(
            FallbackCoordinator(lambda _error: True).batches(
                lambda: NormalizedStreamingReader(primary, compiled),
                fallback_factory,
            )
        )

    assert captured.value is primary_error
    assert fallback_calls == 0
    assert _fallback_block_reason(primary_error) is _FallbackBlockReason.SOURCE_OWNERSHIP


def test_explicit_close_surfaces_cleanup_failure_once() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    cleanup = OSError("cleanup")
    raw = _Reader(compiled.input_schema, close_error=cleanup)
    reader = NormalizedStreamingReader(raw, compiled)

    with pytest.raises(OSError) as captured:
        reader.close()

    assert captured.value is cleanup
    assert raw.close_calls == 1
    reader.close()
    assert raw.close_calls == 1


def test_hinted_and_unhinted_typed_all_null_columns_fix_distinct_schemas() -> None:
    null_strings = pa.array([None, None], type=pa.string())
    sample = _sample(
        (null_strings, null_strings),
        ("hinted", "unhinted"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(type_hints={"hinted": "DATE"}),
    )

    assert compiled.schema.types == [pa.date32(), pa.null()]

    with pytest.raises(StreamingTypeError) as captured:
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch(
                [
                    pa.array([None], type=pa.string()),
                    pa.array(["late"], type=pa.string()),
                ],
                schema=compiled.input_schema,
            )
        )

    assert captured.value.context["ordinal"] == 1


def test_missing_marker_case_and_extended_policy_match_legacy_contract() -> None:
    values = pa.array(["NA", "na", "Null", "-", "keep"])
    rows = pa.array([1, 2, 3, 4, 5], type=pa.int64())
    default = compile_normalization_plan(
        _sample((values, rows), ("note", "row")),
        _parse_plan(batch_size=5),
    )
    extended = compile_normalization_plan(
        _sample((values, rows), ("note", "row")),
        _parse_plan(batch_size=5, use_extended_missing_list=True),
    )

    default_result = ArrowNormalizationOperation(default).normalize(
        pa.record_batch([values, rows], schema=default.input_schema)
    )
    extended_result = ArrowNormalizationOperation(extended).normalize(
        pa.record_batch([values, rows], schema=extended.input_schema)
    )

    assert default_result.column(0).to_pylist() == [None, "na", "Null", "-", "keep"]
    assert extended_result.column(0).to_pylist() == [None, "na", "Null", None, "keep"]


@pytest.mark.parametrize(
    ("values", "decimal", "thousands", "expected"),
    [
        (("$1,234.50", "(2.50)"), ".", ",", (1234.5, -2.5)),
        (("1.234,50", "2.345,75"), ",", ".", (1234.5, 2345.75)),
        (("1\xa0234,50", "2\xa0345,75"), ",", " ", (1234.5, 2345.75)),
    ],
)
def test_compiled_numeric_locale_currency_accounting_and_nbsp(
    values: tuple[str, ...],
    decimal: str,
    thousands: str,
    expected: tuple[float, ...],
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(values),), ("amount",)),
        _parse_plan(
            decimal_separator=decimal,
            thousands_separator=thousands,
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([pa.array(values)], schema=compiled.input_schema)
    )

    assert tuple(result.column(0).to_pylist()) == expected


def test_mixed_numeric_locale_is_fixed_once_and_applied_per_value() -> None:
    values = pa.array(["1,234.50", "1.234,50"])
    compiled = compile_normalization_plan(
        _sample((values,), ("amount",)),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.columns[0].numeric_mode == "mixed_locale"
    assert result.column(0).to_pylist() == [1234.5, 1234.5]


def test_mixed_numeric_locale_applies_same_per_value_rule_to_late_batches() -> None:
    sample_values = pa.array(["1,234.56", "2.345,67"])
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("amount",)),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [pa.array(["3.456,78", "4,567.89"])],
            schema=compiled.input_schema,
        )
    )

    assert compiled.columns[0].numeric_mode == "mixed_locale"
    assert result.column(0).to_pylist() == [3456.78, 4567.89]


def test_nan_is_missing_for_empty_rows_and_never_matches_regex() -> None:
    values = pa.array([float("nan"), 1.0], type=pa.float64())
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_regex="nan"),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == [1.0]


def test_nan_is_null_with_a_companion_and_drops_an_otherwise_empty_row() -> None:
    nan = pa.array([float("nan")], type=pa.float64())
    companion = pa.array([1], type=pa.int64())
    compiled = compile_normalization_plan(
        _sample((nan, companion), ("value", "row")),
        _parse_plan(),
    )

    with_companion = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([nan, companion], schema=compiled.input_schema)
    )

    assert with_companion.column(0).to_pylist() == [None]

    nan_only = compile_normalization_plan(
        _sample((nan,), ("value",)),
        _parse_plan(),
    )
    empty = ArrowNormalizationOperation(nan_only).normalize(
        pa.record_batch([nan], schema=nan_only.input_schema)
    )
    assert empty.num_rows == 0


def test_whitespace_only_is_missing_when_whitespace_normalization_is_disabled() -> None:
    blank = pa.array(["  \t\xa0 "])
    companion = pa.array([1], type=pa.int64())
    compiled = compile_normalization_plan(
        _sample((blank, companion), ("note", "row")),
        _parse_plan(normalize_whitespace=False),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([blank, companion], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == [None]


@pytest.mark.parametrize("late", ["1.5", "9223372036854775808"])
def test_late_fractional_or_overflow_value_rejects_fixed_integer_schema(late: str) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1", "2"]),), ("value",)),
        _parse_plan(),
    )

    with pytest.raises(StreamingTypeError) as captured:
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch([pa.array([late])], schema=compiled.input_schema)
        )

    assert captured.value.context["expected_type"] == "int64"


def test_identifier_rule_preserves_leading_zeroes() -> None:
    values = pa.array(["001", "002"])
    compiled = compile_normalization_plan(
        _sample((values,), ("customer_id",)),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.string()]
    assert result.column(0).to_pylist() == ["001", "002"]


def test_mixed_text_date_formats_are_compiled_and_late_invalid_is_strict() -> None:
    sample_values = pa.array(["2024-01-01", "02/01/2024"])
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("event_date",)),
        _parse_plan(),
    )
    operation = ArrowNormalizationOperation(compiled)

    valid = operation.normalize(
        pa.record_batch(
            [pa.array(["2024-03-01", "04/01/2024"])],
            schema=compiled.input_schema,
        )
    )
    with pytest.raises(StreamingTypeError):
        operation.normalize(
            pa.record_batch([pa.array(["not-a-date"])], schema=compiled.input_schema)
        )

    assert compiled.columns[0].date_format == "mixed"
    assert valid.num_rows == 2


def test_late_numeric_serial_outside_excel_range_is_strict() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array([1, 2], type=pa.int64()),), ("event_date",)),
        _parse_plan(),
    )

    with pytest.raises(StreamingTypeError):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch(
                [pa.array([60_001], type=pa.int64())],
                schema=compiled.input_schema,
            )
        )


def test_single_and_missing_conditions_resolve_against_final_names() -> None:
    sample = _sample(
        (pa.array(["drop", "keep"]), pa.array(["x", "y"])),
        ("Status Code", "Other"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(
            sanitize_column_names=True,
            column_renames={"status_code": "status"},
            drop_conditions=[
                {"column": "missing", "value": "ignored"},
                {"column": "status", "value": "drop"},
            ],
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(list(sample.columns), schema=compiled.input_schema)
    )

    assert compiled.final_display_names == ("status", "other")
    assert result.to_pydict() == {"0": ["keep"], "1": ["y"]}


def test_condition_operands_are_compiled_for_each_arrow_type() -> None:
    arrays = (
        pa.array([1.5, 2.5], type=pa.float64()),
        pa.array([date(2024, 1, 1), date(2024, 1, 2)], type=pa.date32()),
        pa.array(
            [datetime(2024, 1, 1, 12), datetime(2024, 1, 2, 12)],
            type=pa.timestamp("us"),
        ),
    )
    compiled = compile_normalization_plan(
        _sample(arrays, ("amount", "day", "moment")),
        _parse_plan(
            drop_conditions=[
                {"column": "amount", "value": 1.5},
                {"column": "day", "value": date(2024, 1, 2)},
                {"column": "moment", "value": datetime(2099, 1, 1)},
            ],
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(list(arrays), schema=compiled.input_schema)
    )

    assert result.num_rows == 0


@pytest.mark.parametrize("operand", [b"a", None, float("nan"), date(2024, 1, 1)])
def test_cross_type_condition_operand_is_a_harmless_non_match(operand: object) -> None:
    values = pa.array(["a", "b"])
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": operand}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == ["a", "b"]


@pytest.mark.parametrize(
    ("values", "operand", "expected"),
    [
        (pa.array([1, 2], type=pa.int64()), True, [2]),
        (pa.array([0, 1], type=pa.int64()), False, [1]),
        (pa.array([1.0, 2.0], type=pa.float64()), True, [2.0]),
        (pa.array([True, False]), 1, [False]),
        (pa.array([True, False]), 1.0, [False]),
        (pa.array([True, False]), 0, [True]),
    ],
)
def test_boolean_numeric_conditions_match_characterized_pandas_equality(
    values: pa.Array,
    operand: object,
    expected: list[object],
) -> None:
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": operand}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == expected


@pytest.mark.parametrize(
    ("values", "operand", "expected"),
    [
        (pa.array([1, 2], type=pa.int64()), Decimal("1"), [2]),
        (pa.array([1.5, 2.5], type=pa.float64()), Decimal("1.5"), [2.5]),
        (
            pa.array([Decimal("1.5"), Decimal("2.5")], type=pa.decimal128(3, 1)),
            1.5,
            [Decimal("2.5")],
        ),
        (pa.array([True, False]), Decimal("1"), [False]),
        (
            pa.array([Decimal("1.0"), Decimal("2.0")], type=pa.decimal128(3, 1)),
            True,
            [Decimal("2.0")],
        ),
    ],
)
def test_decimal_numeric_conditions_match_characterized_pandas_equality(
    values: pa.Array,
    operand: object,
    expected: list[object],
) -> None:
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": operand}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == expected


@pytest.mark.parametrize(
    ("values", "operand"),
    [
        (pa.array([1, 2], type=pa.int8()), Decimal("1.5")),
        (pa.array([1, 2], type=pa.int8()), Decimal("128")),
        (pa.array([0.1, 0.2], type=pa.float64()), Decimal("0.1")),
        (pa.array([1.0, 2.0], type=pa.float64()), Decimal("NaN")),
        (
            pa.array([Decimal("0.1"), Decimal("0.2")], type=pa.decimal128(3, 1)),
            0.1,
        ),
        (
            pa.array([Decimal("1.0"), Decimal("2.0")], type=pa.decimal128(3, 1)),
            float("inf"),
        ),
    ],
)
def test_unsafe_numeric_condition_conversions_are_harmless_non_matches(
    values: pa.Array,
    operand: object,
) -> None:
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": operand}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == values.to_pylist()
    assert compiled.drop_conditions[0].operands == (None,)


def test_decimal_condition_must_be_exact_in_the_arrow_float_width() -> None:
    values = pa.array([16_777_216.0, 2.0], type=pa.float32())
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": Decimal("16777217")}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == values.to_pylist()
    assert compiled.drop_conditions[0].operands == (None,)


def test_exact_decimal_condition_matches_float16_values() -> None:
    values = pa.array([1.5, 2.0], type=pa.float16())
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": Decimal("1.5")}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == [2.0]


@pytest.mark.parametrize(
    ("values", "operand"),
    [
        (pa.array([2_048.0, 2.0], type=pa.float16()), Decimal("2049")),
        (pa.array([0.0, 2.0], type=pa.float32()), Decimal("1e-50")),
        (pa.array([0.0, 2.0], type=pa.float32()), Decimal("1e100")),
    ],
)
def test_decimal_float_width_loss_is_a_harmless_non_match(
    values: pa.Array,
    operand: Decimal,
) -> None:
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_conditions=[{"column": "value", "value": operand}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == values.to_pylist()
    assert compiled.drop_conditions[0].operands == (None,)


def test_decimal_condition_preserves_duplicate_label_positional_masking() -> None:
    arrays = (
        pa.array([1, 2], type=pa.int64()),
        pa.array([1, 3], type=pa.int64()),
    )
    compiled = compile_normalization_plan(
        _sample(arrays, ("value", "value")),
        _parse_plan(drop_conditions=[{"column": "value", "value": Decimal("1")}]),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(list(arrays), schema=compiled.input_schema)
    )

    assert result.to_pydict() == {"0": [None, 2], "1": [None, 3]}


def test_passthrough_ree_duplicate_condition_masks_encoded_values() -> None:
    arrays = (
        pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2], type=pa.int16()),
            pa.array([1, 2], type=pa.int64()),
        ),
        pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2], type=pa.int16()),
            pa.array([1, 3], type=pa.int64()),
        ),
    )
    compiled = compile_normalization_plan(
        _sample(arrays, ("value", "value")),
        _parse_plan(
            drop_conditions=[{"column": "value", "value": 1}],
            batch_size=2,
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(list(arrays), schema=compiled.input_schema)
    )

    assert result.to_pydict() == {"0": [None, 2], "1": [None, 3]}
    assert result.schema.equals(compiled.schema, check_metadata=True)
    assert result.column(0).type == arrays[0].type
    assert result.column(1).type == arrays[1].type


def test_hostile_numeric_subclass_condition_is_never_executed_or_retained() -> None:
    callbacks: list[str] = []

    class HostileDecimal(Decimal):
        def __float__(self) -> float:
            callbacks.append("float")
            raise AssertionError("hostile numeric conversion")

        def __int__(self) -> int:
            callbacks.append("int")
            raise AssertionError("hostile numeric conversion")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("eq")
            raise AssertionError("hostile numeric comparison")

        def __hash__(self) -> int:
            callbacks.append("hash")
            raise AssertionError("hostile numeric hashing")

    operand = HostileDecimal("1")
    operand_ref = weakref.ref(operand)

    snapshot = normalization_plan._snapshot_condition_value(operand)

    assert isinstance(snapshot, tuple)
    assert snapshot[:1] == ("unsupported",)
    assert callbacks == []
    del operand
    gc.collect()
    assert operand_ref() is None


def test_duplicate_subset_condition_raises_after_prior_condition_drops_every_row() -> None:
    sample = _sample(
        (pa.array(["drop"]), pa.array(["drop"]), pa.array(["x"])),
        ("Status", "Status", "Other"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(
            drop_conditions=[
                {"column": "Other", "value": "x"},
                {"column": "Status", "value": "drop"},
            ],
        ),
    )

    with pytest.raises(ValueError, match="duplicate labels"):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch(list(sample.columns), schema=compiled.input_schema)
        )


def test_non_type_operation_failure_is_non_retryable_before_first_output() -> None:
    sample = _sample(
        (pa.array(["drop"]), pa.array(["keep"]), pa.array(["x"])),
        ("Status", "Status", "Other"),
    )
    compiled = compile_normalization_plan(
        sample,
        _parse_plan(drop_conditions=[{"column": "Status", "value": "drop"}]),
    )
    primary = _Reader(
        compiled.input_schema,
        [pa.record_batch(list(sample.columns), schema=compiled.input_schema)],
    )
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(compiled.schema)

    with pytest.raises(ValueError, match="duplicate labels") as captured:
        list(
            FallbackCoordinator(lambda _error: True).batches(
                lambda: NormalizedStreamingReader(primary, compiled),
                fallback_factory,
            )
        )

    assert fallback_calls == 0
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.CONFIGURATION


def test_reader_closes_owned_source_when_plan_validation_fails() -> None:
    raw = _Reader(pa.schema([]))

    with pytest.raises(TypeError, match="NormalizationPlan"):
        NormalizedStreamingReader(raw, object())  # type: ignore[arg-type]

    assert raw.close_calls == 1


def test_direct_operation_is_terminal_after_invalid_input() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]),), ("value",)),
        _parse_plan(),
    )
    operation = ArrowNormalizationOperation(compiled)

    with pytest.raises(TypeError, match="RecordBatch"):
        operation.normalize(object())  # type: ignore[arg-type]
    with pytest.raises(RuntimeError, match="terminal"):
        operation.normalize(pa.record_batch([pa.array(["x"])], schema=compiled.input_schema))


def test_eof_cleanup_compatibility_failure_cannot_trigger_fallback() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]),), ("value",)),
        _parse_plan(),
    )
    cleanup = CoordinateCompatibilityError("cleanup")
    primary = _Reader(compiled.input_schema, [None], close_error=cleanup)
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(compiled.schema)

    with pytest.raises(CoordinateCompatibilityError) as captured:
        list(
            FallbackCoordinator(lambda _error: True).batches(
                lambda: NormalizedStreamingReader(primary, compiled),
                fallback_factory,
            )
        )

    assert captured.value is cleanup
    assert fallback_calls == 0
    assert _fallback_block_reason(cleanup) is _FallbackBlockReason.SOURCE_OWNERSHIP


def test_early_close_is_idempotent_and_retains_schema() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]),), ("note",)),
        _parse_plan(),
    )
    raw = _Reader(
        compiled.input_schema,
        [pa.record_batch([pa.array(["x"])], schema=compiled.input_schema)],
    )
    reader = NormalizedStreamingReader(raw, compiled)

    reader.close()
    reader.close()

    assert reader.schema == compiled.schema
    assert reader.read_next_batch() is None
    assert raw.close_calls == 1


def test_context_body_error_wins_over_ordinary_cleanup_failure() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]),), ("note",)),
        _parse_plan(),
    )
    body_error = ValueError("body")
    cleanup = OSError("cleanup")
    reader = NormalizedStreamingReader(
        _Reader(compiled.input_schema, close_error=cleanup),
        compiled,
    )

    with pytest.raises(ValueError) as captured, reader:
        raise body_error

    assert captured.value is body_error
    assert captured.value.backend_context["cleanup_failure"] == {"type": "OSError"}


@pytest.mark.parametrize("cleanup", [OSError("close"), MemoryError("close")])
def test_hostile_close_descriptor_failure_surfaces_once(cleanup: BaseException) -> None:
    class DescriptorReader:
        schema = pa.schema([("0", pa.string())])

        def __init__(self) -> None:
            self.close_accesses = 0

        def read_next_batch(self) -> None:
            return None

        @property
        def close(self) -> object:
            self.close_accesses += 1
            raise cleanup

    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]),), ("note",)),
        _parse_plan(),
    )
    raw = DescriptorReader()
    reader = NormalizedStreamingReader(raw, compiled)

    with pytest.raises(type(cleanup)) as captured:
        reader.close()

    assert captured.value is cleanup
    reader.close()
    assert raw.close_accesses == 1


def test_arrow_fast_paths_avoid_scalar_normalization(monkeypatch: pytest.MonkeyPatch) -> None:
    text_values = pa.array(["  keep  ", "NA"])
    float_values = pa.array([float("nan"), 2.0], type=pa.float64())
    compiled = compile_normalization_plan(
        _sample((text_values, float_values), ("note", "value")),
        _parse_plan(),
    )

    def fail_scalar(*_args: object) -> object:
        raise AssertionError("Arrow-exact columns must not use the scalar fallback")

    monkeypatch.setattr(arrow_pipeline, "_normalize_scalar", fail_scalar)
    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [text_values, float_values],
            schema=compiled.input_schema,
        )
    )

    assert result.column(0).to_pylist() == ["keep", None]
    assert result.column(1).to_pylist() == [None, 2.0]


def test_string_regex_filter_uses_arrow_kernel(monkeypatch: pytest.MonkeyPatch) -> None:
    values = pa.array(["drop", "keep"])
    compiled = compile_normalization_plan(
        _sample((values,), ("note",)),
        _parse_plan(drop_regex="drop"),
    )

    def fail_scalar(_value: object) -> str:
        raise AssertionError("primitive regex filtering must be vectorized")

    monkeypatch.setattr(arrow_pipeline, "_regex_text", fail_scalar)
    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == ["keep"]


@pytest.mark.parametrize(
    "values",
    [
        pa.array([["drop"], ["keep"]]),
        pa.array(
            [{"field": "drop"}, {"field": "keep"}],
            type=pa.struct([("field", pa.string())]),
        ),
    ],
)
def test_nested_regex_columns_are_type_gated_before_scalar_materialization(
    values: pa.Array,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    compiled = compile_normalization_plan(
        _sample((values,), ("nested",)),
        _parse_plan(drop_regex="drop", batch_size=2),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    def fail_materialization(_value: object) -> str:
        raise AssertionError("nested regex values must not be materialized")

    monkeypatch.setattr(arrow_pipeline, "_regex_text", fail_materialization)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert result is batch


def test_mixed_date_fallback_calls_pandas_once_for_one_bounded_column(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sample_values = pa.array(["2024-01-01", "02/01/2024"])
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("event_date",)),
        _parse_plan(),
    )
    calls = 0
    original = arrow_pipeline.pd.to_datetime

    def counted(*args: object, **kwargs: object) -> object:
        nonlocal calls
        calls += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(arrow_pipeline.pd, "to_datetime", counted)
    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [pa.array(["2024-03-01", "04/01/2024"])],
            schema=compiled.input_schema,
        )
    )

    assert result.num_rows == 2
    assert calls == 1


@pytest.mark.parametrize(
    "values",
    [
        ("January 15, 2024", "February 16, 2024"),
        ("15 Jan 2024", "16 Feb 2024"),
        ("15/01/2024 10:30:00", "16/01/2024 11:45:00"),
    ],
)
def test_fixed_legacy_date_formats_use_arrow_without_pandas(
    values: tuple[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sample_values = pa.array(values)
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("event_date",)),
        _parse_plan(),
    )

    def fail_pandas(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("fixed date formats must use Arrow")

    monkeypatch.setattr(arrow_pipeline.pd, "to_datetime", fail_pandas)
    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([sample_values], schema=compiled.input_schema)
    )

    assert result.num_rows == 2


@pytest.mark.parametrize("late", ["inf", "-inf"])
def test_late_nonfinite_text_rejects_fixed_float_schema(late: str) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1.5", "2.5"]),), ("amount",)),
        _parse_plan(),
    )

    with pytest.raises(StreamingTypeError):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch([pa.array([late])], schema=compiled.input_schema)
        )


@pytest.mark.parametrize("late", [0, -1, 1.5, 60_001, float("inf")])
def test_late_numeric_serial_must_be_finite_integral_and_in_range(
    late: int | float,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array([1.0, 2.0], type=pa.float64()),), ("event_date",)),
        _parse_plan(),
    )
    late_batch = pa.array([late], type=pa.float64())

    with pytest.raises(StreamingTypeError):
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch([late_batch], schema=compiled.input_schema)
        )


@pytest.mark.parametrize(
    ("array", "pattern"),
    [
        (pa.array([b"drop", b"keep"], type=pa.binary()), "drop"),
        (
            pa.array(
                [Decimal("12.50"), Decimal("7.25")],
                type=pa.decimal128(8, 2),
            ),
            r"12\.50",
        ),
    ],
)
def test_regex_filters_binary_and_decimal_primitive_values(
    array: pa.Array,
    pattern: str,
) -> None:
    compiled = compile_normalization_plan(
        _sample((array,), ("value",)),
        _parse_plan(drop_regex=pattern),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([array], schema=compiled.input_schema)
    )

    assert result.num_rows == 1


@pytest.mark.parametrize("pattern", [r"(?<=pré)fix", r"\w+"])
def test_python_only_or_unicode_regex_uses_exact_fallback(pattern: str) -> None:
    values = pa.array(["préfix", "---"])
    compiled = compile_normalization_plan(
        _sample((values,), ("note",)),
        _parse_plan(drop_regex=pattern),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == ["---"]


def test_regex_dollar_anchor_matches_before_a_final_newline_like_python() -> None:
    values = pa.array(["x\n", "\nx", ""])
    keep = pa.array(["before-final-newline", "at-end", "empty"])
    compiled = compile_normalization_plan(
        _sample((values, keep), ("value", "keep")),
        _parse_plan(
            batch_size=3,
            drop_regex=r"x$",
            normalize_whitespace=False,
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values, keep], schema=compiled.input_schema)
    )

    assert result.column(1).to_pylist() == ["empty"]


def test_operation_does_not_retain_prior_batches_or_use_forbidden_materializers() -> None:
    source = inspect.getsource(arrow_pipeline)
    assert ".to_pylist(" not in source
    assert ".to_pydict(" not in source
    assert "pd.DataFrame(" not in source

    values = pa.array(["keep"])
    compiled = compile_normalization_plan(
        _sample((values,), ("note",)),
        _parse_plan(batch_size=1),
    )
    operation = ArrowNormalizationOperation(compiled)
    batch = pa.record_batch([values], schema=compiled.input_schema)
    batch_ref = weakref.ref(batch)

    operation.normalize(batch)
    del batch
    gc.collect()

    assert batch_ref() is None


@pytest.mark.parametrize(
    "values",
    [pa.array([1.5, 2.5], type=pa.float64()), pa.array(["alpha", "beta"])],
)
def test_clean_float_and_string_normalization_reuses_input_buffers(
    values: pa.Array,
) -> None:
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert result is batch
    assert tuple(
        None if buffer is None else buffer.address for buffer in result.column(0).buffers()
    ) == tuple(None if buffer is None else buffer.address for buffer in values.buffers())


def test_clean_string_preflight_skips_replacement_and_trim_candidates(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    values = pa.array(["alpha", "two words"])
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    def fail_candidate(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("clean strings must not materialize transform candidates")

    monkeypatch.setattr(arrow_pipeline.pc, "replace_substring_regex", fail_candidate)
    monkeypatch.setattr(arrow_pipeline.pc, "utf8_trim_whitespace", fail_candidate)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert result is batch


def test_missing_marker_array_is_cached_once_per_operation_and_released(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    left = pa.array(["alpha", "beta"])
    right = pa.array(["gamma", "delta"])
    compiled = compile_normalization_plan(
        _sample((left, right), ("left", "right")),
        _parse_plan(),
    )
    marker_values = tuple(sorted(compiled.columns[0].missing_values))
    real_array = arrow_pipeline.pa.array
    marker_constructions = 0
    marker_ref: weakref.ReferenceType[pa.Array] | None = None

    def counted_array(values: object, *args: object, **kwargs: object) -> pa.Array:
        nonlocal marker_constructions, marker_ref
        result = real_array(values, *args, **kwargs)
        if isinstance(values, (list, tuple)) and tuple(values) == marker_values:
            marker_constructions += 1
            if marker_ref is None:
                marker_ref = weakref.ref(result)
        return result

    monkeypatch.setattr(arrow_pipeline.pa, "array", counted_array)
    operation = ArrowNormalizationOperation(compiled)
    batch = pa.record_batch([left, right], schema=compiled.input_schema)

    assert operation.normalize(batch) is batch
    assert operation.normalize(batch) is batch
    assert marker_constructions == 1
    assert marker_ref is not None and marker_ref() is not None

    del operation
    gc.collect()
    assert marker_ref() is None


def test_compiled_plan_does_not_retain_hostile_label_or_condition_operand() -> None:
    class Hostile:
        def __str__(self) -> str:
            raise AssertionError("str must not run")

        def __repr__(self) -> str:
            raise AssertionError("repr must not run")

    label = Hostile()
    operand = Hostile()
    label_ref = weakref.ref(label)
    operand_ref = weakref.ref(operand)
    config = SheetConfig(
        auto_detect=False,
        sanitize_column_names=True,
        drop_conditions=[{"column": "missing", "value": operand}],
    )
    parse_plan = compile_parse_plan(
        config,
        structure=None,
        format_type="xlsx",
        output_mode=OutputMode.STREAMING,
        batch_size=1,
    )
    sample = _sample((pa.array(["x"]),), (label,))

    compiled = compile_normalization_plan(sample, parse_plan)

    del sample, parse_plan, config, label, operand
    gc.collect()
    assert label_ref() is None
    assert operand_ref() is None
    assert len(compiled.final_display_names) == 1
    assert type(compiled.final_display_names[0]) is str


@pytest.mark.parametrize("sanitize", [False, True])
def test_hostile_label_type_metadata_is_never_executed_or_retained(
    sanitize: bool,
) -> None:
    callbacks: list[str] = []

    class HostileMetadata:
        def __str__(self) -> str:
            callbacks.append("str")
            raise AssertionError("metadata str must not run")

        def __repr__(self) -> str:
            callbacks.append("repr")
            raise AssertionError("metadata repr must not run")

        def __format__(self, _spec: str) -> str:
            callbacks.append("format")
            raise AssertionError("metadata format must not run")

        def __hash__(self) -> int:
            callbacks.append("hash")
            raise AssertionError("metadata hash must not run")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("eq")
            raise AssertionError("metadata equality must not run")

    class HostileTextMetadata(str):
        def __str__(self) -> str:
            callbacks.append("str")
            raise AssertionError("metadata str must not run")

        def __repr__(self) -> str:
            callbacks.append("repr")
            raise AssertionError("metadata repr must not run")

        def __format__(self, _spec: str) -> str:
            callbacks.append("format")
            raise AssertionError("metadata format must not run")

        def __hash__(self) -> int:
            callbacks.append("hash")
            raise AssertionError("metadata hash must not run")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("eq")
            raise AssertionError("metadata equality must not run")

    class HostileLabel:
        pass

    module = HostileMetadata()
    qualname = HostileTextMetadata("HostileLabel")
    HostileLabel.__module__ = module  # type: ignore[assignment]
    HostileLabel.__qualname__ = qualname
    label = HostileLabel()
    label_ref = weakref.ref(label)
    module_ref = weakref.ref(module)
    sample = _sample((pa.array(["x"]),), (label,))

    first = compile_normalization_plan(
        sample,
        _parse_plan(sanitize_column_names=sanitize),
    )
    second = compile_normalization_plan(
        sample,
        _parse_plan(sanitize_column_names=sanitize),
    )

    assert first == second
    assert hash(first) == hash(second)
    assert callbacks == []
    HostileLabel.__module__ = "test_streaming_normalization"
    HostileLabel.__qualname__ = "HostileLabel"
    del sample, label, module, qualname
    gc.collect()
    assert label_ref() is None
    assert module_ref() is None


@pytest.mark.parametrize("sanitize", [False, True])
def test_hostile_metaclass_hash_is_not_used_for_label_classification(
    sanitize: bool,
) -> None:
    callbacks: list[str] = []

    class HostileMetaclass(type):
        def __hash__(cls) -> int:
            callbacks.append("metaclass-hash")
            return type.__hash__(cls)

    class HostileLabel(metaclass=HostileMetaclass):
        pass

    label = HostileLabel()
    label_ref = weakref.ref(label)
    sample = _sample((pa.array(["x"]),), (label,))

    compiled = compile_normalization_plan(
        sample,
        _parse_plan(sanitize_column_names=sanitize),
    )

    assert normalization_plan._safe_name_text(label) == ""
    assert arrow_pipeline._safe_display_label(label) == "non-string label"
    assert arrow_pipeline._safe_value_description(label) == "unsupported value"
    assert arrow_pipeline._regex_text(label) == ""
    assert callbacks == []
    assert hash(compiled)
    del sample, label
    gc.collect()
    assert label_ref() is None


@pytest.mark.parametrize(
    ("hostile_type", "expected"),
    [
        (type("HostileText", (str,), {}), "unsupported value"),
        (type("HostileBytes", (bytes,), {}), "unsupported value"),
    ],
)
def test_safe_value_description_accepts_only_exact_builtin_primitives(
    hostile_type: type[object],
    expected: str,
) -> None:
    callbacks: list[str] = []

    def hostile_len(_self: object) -> int:
        callbacks.append("len")
        raise RuntimeError("hostile length")

    hostile_type.__len__ = hostile_len  # type: ignore[attr-defined]
    raw_value = b"secret" if issubclass(hostile_type, bytes) else "secret"
    value = hostile_type(raw_value)
    value_ref = None if isinstance(value, bytes) else weakref.ref(value)

    assert arrow_pipeline._safe_value_description(value) == expected
    assert arrow_pipeline._safe_value_description("safe") == "str(length=4)"
    assert arrow_pipeline._safe_value_description(b"safe") == "bytes(length=4)"
    assert callbacks == []

    del value
    gc.collect()
    assert value_ref is None or value_ref() is None


def test_late_extension_scalar_error_description_is_sanitized_without_callbacks() -> None:
    callbacks: list[str] = []
    value_refs: list[weakref.ReferenceType[object]] = []

    class HostileText(str):
        def __len__(self) -> int:
            callbacks.append("len")
            raise RuntimeError("hostile length")

    class HostileScalar(pa.ExtensionScalar):
        def as_py(self, **_kwargs: object) -> object:
            value = HostileText(self.value.as_py())
            value_refs.append(weakref.ref(value))
            return value

    class HostileTextType(pa.ExtensionType):
        def __init__(self) -> None:
            super().__init__(pa.string(), "test.hostile-text-value")

        def __arrow_ext_serialize__(self) -> bytes:
            return b""

        @classmethod
        def __arrow_ext_deserialize__(
            cls,
            _storage_type: pa.DataType,
            _serialized: bytes,
        ) -> HostileTextType:
            return cls()

        def __arrow_ext_scalar_class__(self) -> type[pa.ExtensionScalar]:
            return HostileScalar

    extension_type = HostileTextType()
    sample_values = pa.ExtensionArray.from_storage(extension_type, pa.array(["1"]))
    late_values = pa.ExtensionArray.from_storage(extension_type, pa.array(["bad"]))
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("amount",)),
        _parse_plan(type_hints={"amount": "INTEGER"}),
    )
    operation = ArrowNormalizationOperation(compiled)

    with pytest.raises(StreamingTypeError) as captured:
        operation.normalize(
            pa.record_batch([late_values], schema=compiled.input_schema),
        )

    assert captured.value.context == {
        "ordinal": 0,
        "display_label": "str label(length=6)",
        "row_offset": 0,
        "value_description": "unsupported value",
        "expected_type": "int64",
    }
    assert callbacks == []

    del captured, operation
    gc.collect()
    assert value_refs
    assert all(value_ref() is None for value_ref in value_refs)


def test_hostile_metaclass_hash_is_not_used_for_condition_snapshotting() -> None:
    callbacks: list[str] = []

    class HostileMetaclass(type):
        def __hash__(cls) -> int:
            callbacks.append("metaclass-hash")
            return type.__hash__(cls)

    @dataclass
    class HostileOperand(metaclass=HostileMetaclass):
        value: int

    operand = HostileOperand(1)
    operand_ref = weakref.ref(operand)
    parse_plan = _parse_plan(
        drop_conditions=[{"column": "value", "value": operand}],
    )

    compiled = compile_normalization_plan(
        _sample((pa.array(["keep"]),), ("value",)),
        parse_plan,
    )

    assert compiled.drop_conditions[0].operands == (None,)
    assert callbacks == []
    del parse_plan, operand
    gc.collect()
    assert operand_ref() is None


def test_semantic_error_never_triggers_pre_output_fallback() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    failing = pa.record_batch([pa.array(["bad"])], schema=compiled.input_schema)
    primary = _Reader(compiled.input_schema, [failing])
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(compiled.schema)

    with pytest.raises(StreamingTypeError):
        list(
            FallbackCoordinator(lambda _error: True).batches(
                lambda: NormalizedStreamingReader(primary, compiled),
                fallback_factory,
            )
        )

    assert fallback_calls == 0


def test_process_operation_failure_remains_winner_over_process_cleanup() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    primary = KeyboardInterrupt()
    raw = _Reader(compiled.input_schema, [primary], close_error=MemoryError("cleanup"))
    reader = NormalizedStreamingReader(raw, compiled)

    with pytest.raises(KeyboardInterrupt) as captured:
        reader.read_next_batch()

    assert captured.value is primary


def test_eof_cleanup_failure_is_not_reported_as_clean_eof() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    cleanup = OSError("cleanup")
    reader = NormalizedStreamingReader(
        _Reader(compiled.input_schema, [None], close_error=cleanup),
        compiled,
    )

    with pytest.raises(OSError) as captured:
        reader.read_next_batch()

    assert captured.value is cleanup


@pytest.mark.parametrize(
    "cleanup",
    [OSError("close"), MemoryError("close"), SystemExit("close")],
)
def test_coordinator_early_close_exposes_nested_source_cleanup_failure(
    cleanup: BaseException,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    raw = _Reader(
        compiled.input_schema,
        [pa.record_batch([pa.array(["1"])], schema=compiled.input_schema)],
        close_error=cleanup,
    )
    stream = FallbackCoordinator(lambda _error: True).batches(
        lambda: NormalizedStreamingReader(raw, compiled),
        pytest.fail,
    )

    next(stream)
    with pytest.raises(type(cleanup)) as captured:
        stream.close()

    assert captured.value is cleanup
    assert raw.close_calls == 1


def test_source_generator_exit_still_wins_over_ordinary_cleanup_failure() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    primary = GeneratorExit()
    raw = _Reader(
        compiled.input_schema,
        [primary],
        close_error=OSError("close"),
    )
    reader = NormalizedStreamingReader(raw, compiled)

    with pytest.raises(GeneratorExit) as captured:
        reader.read_next_batch()

    assert captured.value is primary
    assert primary.backend_context["cleanup_failure"] == {"type": "OSError"}


def test_invalid_normalization_plan_closes_source_and_never_falls_back() -> None:
    raw = _Reader(pa.schema([]))
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(pa.schema([]))

    with pytest.raises(TypeError, match="NormalizationPlan") as captured:
        list(
            FallbackCoordinator(lambda _error: True).batches(
                lambda: NormalizedStreamingReader(raw, object()),  # type: ignore[arg-type]
                fallback_factory,
            )
        )

    assert raw.close_calls == 1
    assert fallback_calls == 0
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.CONFIGURATION


def test_reader_schema_compatibility_failure_can_fall_back_after_cleanup() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    schema_error = CoordinateCompatibilityError("unsupported schema")

    class SchemaFailureReader:
        close_calls = 0

        @property
        def schema(self) -> pa.Schema:
            raise schema_error

        def read_next_batch(self) -> pa.RecordBatch | None:
            raise AssertionError("schema failure must occur before reads")

        def close(self) -> None:
            self.close_calls += 1

    primary = SchemaFailureReader()
    fallback = _Reader(compiled.schema)

    assert (
        list(
            FallbackCoordinator(lambda error: error is schema_error).batches(
                lambda: NormalizedStreamingReader(primary, compiled),  # type: ignore[arg-type]
                lambda: fallback,
            )
        )
        == []
    )
    assert primary.close_calls == 1
    assert fallback.close_calls == 1
    assert _fallback_block_reason(schema_error) is None


def test_reader_schema_cleanup_failure_blocks_compatible_fallback() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    schema_error = CoordinateCompatibilityError("unsupported schema")
    cleanup_error = OSError("cleanup")

    class SchemaFailureReader:
        close_calls = 0

        @property
        def schema(self) -> pa.Schema:
            raise schema_error

        def read_next_batch(self) -> pa.RecordBatch | None:
            raise AssertionError("schema failure must occur before reads")

        def close(self) -> None:
            self.close_calls += 1
            raise cleanup_error

    primary = SchemaFailureReader()
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(compiled.schema)

    with pytest.raises(CoordinateCompatibilityError) as captured:
        list(
            FallbackCoordinator(lambda error: error is schema_error).batches(
                lambda: NormalizedStreamingReader(primary, compiled),  # type: ignore[arg-type]
                fallback_factory,
            )
        )

    assert captured.value is schema_error
    assert primary.close_calls == 1
    assert fallback_calls == 0
    assert _fallback_block_reason(schema_error) is _FallbackBlockReason.SOURCE_OWNERSHIP
    assert schema_error.backend_context["cleanup_failure"] == {"type": "OSError"}


@pytest.mark.parametrize("descriptor", [False, True])
def test_non_callable_owned_close_blocks_retryable_fallback(descriptor: bool) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]),), ("amount",)),
        _parse_plan(),
    )
    primary_error = CoordinateCompatibilityError("retry")

    class NoCloseReader:
        schema = compiled.input_schema
        close = None

        def read_next_batch(self) -> pa.RecordBatch | None:
            raise primary_error

    class DescriptorReader:
        schema = compiled.input_schema

        def read_next_batch(self) -> pa.RecordBatch | None:
            raise primary_error

        @property
        def close(self) -> object:
            return object()

    raw = DescriptorReader() if descriptor else NoCloseReader()
    fallback_calls = 0

    def fallback_factory() -> _Reader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _Reader(compiled.schema)

    with pytest.raises(CoordinateCompatibilityError) as captured:
        list(
            FallbackCoordinator(lambda _error: True).batches(
                lambda: NormalizedStreamingReader(raw, compiled),  # type: ignore[arg-type]
                fallback_factory,
            )
        )

    assert captured.value is primary_error
    assert fallback_calls == 0
    assert _fallback_block_reason(primary_error) is _FallbackBlockReason.SOURCE_OWNERSHIP
    assert primary_error.backend_context["cleanup_failure"] == {"type": "TypeError"}


def test_nan_display_labels_compile_to_equal_hash_stable_plans() -> None:
    first = compile_normalization_plan(
        _sample((pa.array(["x"]),), (float("nan"),)),
        _parse_plan(),
    )
    second = compile_normalization_plan(
        _sample((pa.array(["x"]),), (float("nan"),)),
        _parse_plan(),
    )

    assert first == second
    assert hash(first) == hash(second)
    assert isinstance(first.source_display_names[0], float)
    assert math.isnan(first.source_display_names[0])


@pytest.mark.parametrize(
    ("left", "right"),
    [(True, 1), (1, 1.0), (0, -0.0), (b"a", b"b")],
)
def test_typed_display_labels_remain_distinct_plan_identities(
    left: object,
    right: object,
) -> None:
    left_plan = compile_normalization_plan(
        _sample((pa.array(["x"]),), (left,)),
        _parse_plan(),
    )
    right_plan = compile_normalization_plan(
        _sample((pa.array(["x"]),), (right,)),
        _parse_plan(),
    )

    assert left_plan != right_plan
    assert hash(left_plan) != hash(right_plan)


def test_distinct_bytes_labels_keep_distinct_sanitized_names() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]), pa.array(["y"])), (b"a", b"b")),
        _parse_plan(sanitize_column_names=True),
    )

    assert compiled.source_display_names == (b"a", b"b")
    assert compiled.final_display_names == ("b_a", "b_b")


def test_distinct_tuple_labels_resolve_hints_renames_and_conditions_positionally() -> None:
    labels = (("left",), ("right",))
    arrays = (pa.array(["1", "2"]), pa.array(["keep", "drop"]))
    hinted = compile_normalization_plan(
        _sample(arrays, labels),
        _parse_plan(type_hints={("left",): "INTEGER"}),
    )
    renamed = compile_normalization_plan(
        _sample(arrays, labels),
        _parse_plan(column_renames={("right",): "renamed"}),
    )
    conditioned = compile_normalization_plan(
        _sample(arrays, labels),
        _parse_plan(drop_conditions=[{"column": ("right",), "value": "drop"}]),
    )

    result = ArrowNormalizationOperation(conditioned).normalize(
        pa.record_batch(list(arrays), schema=conditioned.input_schema)
    )

    assert hinted.schema.types == [pa.int64(), pa.string()]
    assert tuple(rule.explicit_hint for rule in hinted.columns) == ("INTEGER", None)
    assert renamed.final_display_names == (("left",), "renamed")
    assert result.to_pydict() == {"0": [1], "1": ["keep"]}


def test_tuple_label_snapshots_preserve_nested_values_and_sanitize_safely() -> None:
    nested = (
        "outer",
        (b"bytes", date(2024, 1, 2), datetime(2024, 1, 2, 3, 4), time(3, 4)),
    )
    raw = compile_normalization_plan(
        _sample((pa.array(["x"]),), (nested,)),
        _parse_plan(),
    )
    sanitized = compile_normalization_plan(
        _sample((pa.array(["x"]), pa.array(["y"])), (("left",), ("right",))),
        _parse_plan(sanitize_column_names=True),
    )

    assert raw.source_display_names == (nested,)
    assert raw.final_display_names == (nested,)
    assert sanitized.source_display_names == (("left",), ("right",))
    assert sanitized.final_display_names == ("unsafe_label", "unsafe_label_1")


@pytest.mark.parametrize(
    ("left", "right"),
    [(True, 1), (1, 1.0), (0.0, -0.0), (b"a", b"b")],
)
def test_tuple_label_members_keep_exact_type_tagged_identity(
    left: object,
    right: object,
) -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]), pa.array(["y"])), ((left,), (right,))),
        _parse_plan(),
    )

    assert compiled.source_label_tokens[0] != compiled.source_label_tokens[1]
    assert hash(compiled.source_label_tokens[0]) != hash(compiled.source_label_tokens[1])


def test_nested_tuple_nan_labels_have_deterministic_plan_identity() -> None:
    first = compile_normalization_plan(
        _sample((pa.array(["x"]),), (("value", (float("nan"),)),)),
        _parse_plan(),
    )
    second = compile_normalization_plan(
        _sample((pa.array(["x"]),), (("value", (float("nan"),)),)),
        _parse_plan(),
    )

    assert first == second
    assert hash(first) == hash(second)


def test_hostile_tuple_label_member_is_never_executed_or_retained() -> None:
    callbacks: list[str] = []

    class HostileMember:
        def __str__(self) -> str:
            callbacks.append("str")
            raise AssertionError("str must not run")

        def __repr__(self) -> str:
            callbacks.append("repr")
            raise AssertionError("repr must not run")

        def __format__(self, _spec: str) -> str:
            callbacks.append("format")
            raise AssertionError("format must not run")

        def __hash__(self) -> int:
            callbacks.append("hash")
            raise AssertionError("hash must not run")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("eq")
            raise AssertionError("equality must not run")

    member = HostileMember()
    member_ref = weakref.ref(member)
    label = ("safe", member)
    sample = _sample((pa.array(["x"]),), (label,))

    compiled = compile_normalization_plan(sample, _parse_plan())

    assert hash(compiled)
    assert callbacks == []
    del sample, label, member
    gc.collect()
    assert member_ref() is None


@pytest.mark.parametrize("resolution", ["hints", "renames", "conditions"])
def test_unsupported_tuple_member_collisions_are_rejected_before_resolution(
    resolution: str,
) -> None:
    class OpaqueMember:
        def __init__(self, side: str) -> None:
            self.side = side

        def __hash__(self) -> int:
            return hash(self.side)

        def __eq__(self, other: object) -> bool:
            return type(other) is OpaqueMember and self.side == other.side

    labels = ((OpaqueMember("right"),),)
    configuration: dict[str, object]
    if resolution == "hints":
        configuration = {"type_hints": {(OpaqueMember("left"),): "INTEGER"}}
    elif resolution == "renames":
        configuration = {"column_renames": {(OpaqueMember("left"),): "renamed"}}
    else:
        configuration = {"drop_conditions": [{"column": (OpaqueMember("left"),), "value": "drop"}]}
    plan = _parse_plan(**configuration)

    with pytest.raises(ValueError, match="unsupported tuple members"):
        compile_normalization_plan(
            _sample((pa.array(["drop"]),), labels),
            plan,
        )


def test_supported_duplicate_tuple_tokens_remain_resolvable() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1"]), pa.array(["2"])), (("same",), ("same",))),
        _parse_plan(type_hints={("same",): "INTEGER"}),
    )

    assert compiled.source_label_tokens[0] == compiled.source_label_tokens[1]
    assert tuple(rule.explicit_hint for rule in compiled.columns) == ("INTEGER", "INTEGER")


@pytest.mark.parametrize("resolution", ["hints", "renames", "conditions"])
def test_unrelated_safe_configuration_is_allowed_with_an_opaque_tuple_label(
    resolution: str,
) -> None:
    class OpaqueMember:
        pass

    opaque = OpaqueMember()
    opaque_ref = weakref.ref(opaque)
    arrays = (
        pa.array(["opaque", "opaque"]),
        pa.array(["1", "2"] if resolution == "hints" else ["keep", "drop"]),
    )
    configuration: dict[str, object]
    if resolution == "hints":
        configuration = {"type_hints": {"safe": "INTEGER"}}
    elif resolution == "renames":
        configuration = {"column_renames": {"safe": "renamed"}}
    else:
        configuration = {"drop_conditions": [{"column": "safe", "value": "drop"}]}
    sample = _sample(arrays, ((opaque,), "safe"))

    compiled = compile_normalization_plan(sample, _parse_plan(**configuration))

    if resolution == "hints":
        assert tuple(rule.explicit_hint for rule in compiled.columns) == (None, "INTEGER")
    elif resolution == "renames":
        assert compiled.final_display_names[1] == "renamed"
    else:
        result = ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch(list(arrays), schema=compiled.input_schema)
        )
        assert result.to_pydict() == {"0": ["opaque"], "1": ["keep"]}
    del sample, opaque
    gc.collect()
    assert opaque_ref() is None


def test_sanitized_opaque_tuple_name_can_be_renamed_and_conditioned() -> None:
    class OpaqueMember:
        pass

    opaque = OpaqueMember()
    arrays = (pa.array(["drop", "keep"]), pa.array(["x", "y"]))
    compiled = compile_normalization_plan(
        _sample(arrays, ((opaque,), "safe column")),
        _parse_plan(
            sanitize_column_names=True,
            column_renames={"unsafe_label": "opaque_safe"},
            drop_conditions=[{"column": "opaque_safe", "value": "drop"}],
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(list(arrays), schema=compiled.input_schema)
    )

    assert compiled.final_display_names == ("opaque_safe", "safe_column")
    assert result.to_pydict() == {"0": ["keep"], "1": ["y"]}


def test_ambiguous_hostile_tuple_members_are_never_executed_or_retained() -> None:
    callbacks: list[str] = []

    class HostileMember:
        def __str__(self) -> str:
            callbacks.append("str")
            raise AssertionError("str must not run")

        def __repr__(self) -> str:
            callbacks.append("repr")
            raise AssertionError("repr must not run")

        def __format__(self, _spec: str) -> str:
            callbacks.append("format")
            raise AssertionError("format must not run")

        def __hash__(self) -> int:
            callbacks.append("hash")
            raise AssertionError("hash must not run")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("eq")
            raise AssertionError("equality must not run")

    left = HostileMember()
    right = HostileMember()
    left_ref = weakref.ref(left)
    right_ref = weakref.ref(right)
    labels = ((left,), (right,))
    sample = _sample((pa.array(["left"]), pa.array(["right"])), labels)

    with pytest.raises(ValueError, match="unsupported tuple members"):
        compile_normalization_plan(sample, _parse_plan())

    assert callbacks == []
    del sample, labels, left, right
    gc.collect()
    assert left_ref() is None
    assert right_ref() is None


@pytest.mark.parametrize("unsafe_payload", ["offset", "name"])
def test_nested_unsafe_temporal_tuple_collisions_are_rejected_without_callbacks(
    unsafe_payload: str,
) -> None:
    callbacks: list[str] = []

    class HostileName(str):
        __slots__ = ("__weakref__",)

        def __hash__(self) -> int:
            callbacks.append("name-hash")
            raise AssertionError("timezone name must not be hashed")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("name-eq")
            raise AssertionError("timezone name must not be compared")

        def __str__(self) -> str:
            callbacks.append("name-str")
            raise AssertionError("timezone name must not be rendered")

        def __repr__(self) -> str:
            callbacks.append("name-repr")
            raise AssertionError("timezone name must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("name-format")
            raise AssertionError("timezone name must not be formatted")

    class HostileOffset(timedelta):
        __slots__ = ("__weakref__",)

        def __hash__(self) -> int:
            callbacks.append("offset-hash")
            raise AssertionError("timezone offset must not be hashed")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("offset-eq")
            raise AssertionError("timezone offset must not be compared")

        def __str__(self) -> str:
            callbacks.append("offset-str")
            raise AssertionError("timezone offset must not be rendered")

        def __repr__(self) -> str:
            callbacks.append("offset-repr")
            raise AssertionError("timezone offset must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("offset-format")
            raise AssertionError("timezone offset must not be formatted")

    def hostile_label(side: str) -> tuple[tuple[datetime], weakref.ReferenceType[object]]:
        if unsafe_payload == "offset":
            payload: object = HostileOffset(hours=1 if side == "left" else 2)
            exact_timezone = timezone(payload, "hostile")
        else:
            payload = HostileName(side)
            exact_timezone = timezone(timedelta(hours=3), payload)
        return (
            (datetime(2024, 1, 2, 3, 4, 5, tzinfo=exact_timezone),),
            weakref.ref(payload),
        )

    left, left_ref = hostile_label("left")
    right, right_ref = hostile_label("right")
    labels = (left, right)
    sample = _sample((pa.array(["left"]), pa.array(["right"])), labels)

    with pytest.raises(ValueError, match="opaque tuple"):
        compile_normalization_plan(sample, _parse_plan())

    assert callbacks == []
    del sample, labels, left, right
    gc.collect()
    assert left_ref() is None
    assert right_ref() is None


def test_safe_timezone_tuple_and_unrelated_configuration_remain_resolvable() -> None:
    temporal = datetime(
        2024,
        1,
        2,
        3,
        4,
        5,
        tzinfo=timezone(timedelta(hours=3), "safe"),
    )
    compiled = compile_normalization_plan(
        _sample((pa.array(["x"]), pa.array(["1"])), ((temporal,), "amount")),
        _parse_plan(
            type_hints={
                (temporal,): "VARCHAR",
                "amount": "INTEGER",
            }
        ),
    )

    assert tuple(rule.explicit_hint for rule in compiled.columns) == ("VARCHAR", "INTEGER")


@pytest.mark.parametrize("resolution", ["hints", "renames", "conditions"])
@pytest.mark.parametrize("unsafe_payload", ["offset", "name"])
def test_nested_unsafe_temporal_tuple_configuration_is_rejected_without_callbacks(
    resolution: str,
    unsafe_payload: str,
) -> None:
    callbacks: list[str] = []

    class HostileName(str):
        __slots__ = ("__weakref__",)

        def __hash__(self) -> int:
            callbacks.append("name-hash")
            raise AssertionError("timezone name must not be hashed")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("name-eq")
            raise AssertionError("timezone name must not be compared")

        def __str__(self) -> str:
            callbacks.append("name-str")
            raise AssertionError("timezone name must not be rendered")

        def __repr__(self) -> str:
            callbacks.append("name-repr")
            raise AssertionError("timezone name must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("name-format")
            raise AssertionError("timezone name must not be formatted")

    class HostileOffset(timedelta):
        __slots__ = ("__weakref__",)

        def __hash__(self) -> int:
            callbacks.append("offset-hash")
            raise AssertionError("timezone offset must not be hashed")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("offset-eq")
            raise AssertionError("timezone offset must not be compared")

        def __str__(self) -> str:
            callbacks.append("offset-str")
            raise AssertionError("timezone offset must not be rendered")

        def __repr__(self) -> str:
            callbacks.append("offset-repr")
            raise AssertionError("timezone offset must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("offset-format")
            raise AssertionError("timezone offset must not be formatted")

    payload: object
    if unsafe_payload == "offset":
        payload = HostileOffset(hours=3)
        exact_timezone = timezone(payload, "hostile")
    else:
        payload = HostileName("hostile")
        exact_timezone = timezone(timedelta(hours=3), payload)
    payload_ref = weakref.ref(payload)
    label = (datetime(2024, 1, 2, 3, 4, 5, tzinfo=exact_timezone),)
    configuration: dict[str, object]
    if resolution == "hints":
        configuration = {"type_hints": {label: "INTEGER"}}
    elif resolution == "renames":
        configuration = {"column_renames": {label: "renamed"}}
    else:
        configuration = {"drop_conditions": [{"column": label, "value": "drop"}]}
    callbacks.clear()

    with pytest.raises(TypeError, match="unsupported mutable configuration value: datetime"):
        _parse_plan(**configuration)

    assert callbacks == []
    del configuration, label, exact_timezone, payload
    gc.collect()
    assert payload_ref() is None


def test_tuple_label_structuralization_has_a_bounded_depth() -> None:
    label: object = "leaf"
    for _ in range(40):
        label = (label,)

    with pytest.raises(ValueError, match="tuple label exceeds structural limits"):
        compile_normalization_plan(
            _sample((pa.array(["x"]),), (label,)),
            _parse_plan(),
        )


def test_label_mapping_and_condition_resolution_are_linear_in_width(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    width = 80
    labels = tuple(f"column-{ordinal:03d}" for ordinal in range(width))
    renamed = tuple(f"renamed-{ordinal:03d}" for ordinal in range(width))
    token_calls = 0
    equality_calls = 0
    real_token = normalization_plan._display_label_token
    real_equality = normalization_plan._LabelToken.__eq__

    def counted_token(value: object, *args: object, **kwargs: object) -> object:
        nonlocal token_calls
        token_calls += 1
        return real_token(value, *args, **kwargs)  # type: ignore[arg-type]

    def counted_equality(left: object, right: object) -> object:
        nonlocal equality_calls
        equality_calls += 1
        return real_equality(left, right)

    monkeypatch.setattr(normalization_plan, "_display_label_token", counted_token)
    monkeypatch.setattr(normalization_plan._LabelToken, "__eq__", counted_equality)
    arrays = tuple(pa.array(["1"]) for _ in range(width))

    compiled = compile_normalization_plan(
        _sample(arrays, labels),
        _parse_plan(
            type_hints=dict.fromkeys(labels, "INTEGER"),
            column_renames=dict(zip(labels, renamed, strict=True)),
            drop_conditions=[{"column": label, "value": -1} for label in renamed],
        ),
    )

    assert len(compiled.columns) == width
    assert token_calls <= width * 6
    assert equality_calls <= width * 6


@pytest.mark.parametrize("sanitize", [False, True])
@pytest.mark.parametrize("temporal_kind", ["datetime", "time"])
def test_hostile_temporal_label_timezone_is_never_executed_or_retained(
    sanitize: bool,
    temporal_kind: str,
) -> None:
    calls = 0

    class HostileTimezone(tzinfo):
        def utcoffset(self, _value: datetime | None) -> timedelta:
            nonlocal calls
            calls += 1
            raise AssertionError("untrusted timezone must not execute")

        def dst(self, _value: datetime | None) -> timedelta:
            raise AssertionError("untrusted timezone must not execute")

        def tzname(self, _value: datetime | None) -> str:
            raise AssertionError("untrusted timezone must not execute")

    timezone = HostileTimezone()
    label = (
        datetime(2024, 1, 2, 3, 4, 5, tzinfo=timezone)
        if temporal_kind == "datetime"
        else time(3, 4, 5, tzinfo=timezone)
    )
    timezone_ref = weakref.ref(timezone)
    sample = _sample((pa.array(["x"]),), (label,))

    compiled = compile_normalization_plan(
        sample,
        _parse_plan(sanitize_column_names=sanitize),
    )

    del sample, label, timezone
    gc.collect()
    assert calls == 0
    assert timezone_ref() is None
    assert hash(compiled)


@pytest.mark.parametrize("sanitize", [False, True])
@pytest.mark.parametrize("temporal_kind", ["datetime", "time"])
@pytest.mark.parametrize("unsafe_payload", ["name", "offset"])
def test_exact_timezone_hostile_payload_is_never_executed_or_retained(
    sanitize: bool,
    temporal_kind: str,
    unsafe_payload: str,
) -> None:
    callbacks: list[str] = []

    class HostileName(str):
        __slots__ = ("__weakref__",)

        def __hash__(self) -> int:
            callbacks.append("name-hash")
            raise AssertionError("timezone name must not be hashed")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("name-eq")
            raise AssertionError("timezone name must not be compared")

        def __str__(self) -> str:
            callbacks.append("name-str")
            raise AssertionError("timezone name must not be rendered")

        def __repr__(self) -> str:
            callbacks.append("name-repr")
            raise AssertionError("timezone name must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("name-format")
            raise AssertionError("timezone name must not be formatted")

    class HostileOffset(timedelta):
        __slots__ = ("__weakref__",)

        def __hash__(self) -> int:
            callbacks.append("offset-hash")
            raise AssertionError("timezone offset must not be hashed")

        def __eq__(self, _other: object) -> bool:
            callbacks.append("offset-eq")
            raise AssertionError("timezone offset must not be compared")

        def __str__(self) -> str:
            callbacks.append("offset-str")
            raise AssertionError("timezone offset must not be rendered")

        def __repr__(self) -> str:
            callbacks.append("offset-repr")
            raise AssertionError("timezone offset must not be rendered")

        def __format__(self, _spec: str) -> str:
            callbacks.append("offset-format")
            raise AssertionError("timezone offset must not be formatted")

    def hostile_sample() -> tuple[NormalizationSample, weakref.ReferenceType[object]]:
        name = HostileName("hostile") if unsafe_payload == "name" else "safe"
        offset = HostileOffset(hours=3) if unsafe_payload == "offset" else timedelta(hours=3)
        payload_ref = weakref.ref(name if unsafe_payload == "name" else offset)
        exact_timezone = timezone(offset, name)
        label = (
            datetime(2024, 1, 2, 3, 4, 5, tzinfo=exact_timezone)
            if temporal_kind == "datetime"
            else time(3, 4, 5, tzinfo=exact_timezone)
        )
        return _sample((pa.array(["x"]),), (label,)), payload_ref

    first_sample, first_payload_ref = hostile_sample()
    second_sample, second_payload_ref = hostile_sample()

    first = compile_normalization_plan(
        first_sample,
        _parse_plan(sanitize_column_names=sanitize),
    )
    second = compile_normalization_plan(
        second_sample,
        _parse_plan(sanitize_column_names=sanitize),
    )

    assert callbacks == []
    assert type(first.source_display_names[0]) is str
    assert first.final_display_names == (
        ("unsafe_label" if sanitize else first.source_display_names[0]),
    )
    assert first == second
    assert hash(first) == hash(second)
    del first_sample, second_sample
    gc.collect()
    assert first_payload_ref() is None
    assert second_payload_ref() is None


def test_none_condition_label_is_ignored_while_nan_label_resolves_positionally() -> None:
    compiled = compile_normalization_plan(
        _sample(
            (pa.array(["drop", "keep"]), pa.array(["keep", "drop"])),
            (None, float("nan")),
        ),
        _parse_plan(
            drop_conditions=[
                {"column": None, "value": "drop"},
                {"column": float("nan"), "value": "drop"},
            ]
        ),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [pa.array(["drop", "keep"]), pa.array(["keep", "drop"])],
            schema=compiled.input_schema,
        )
    )

    assert result.to_pydict() == {"0": ["drop"], "1": ["keep"]}


@pytest.mark.parametrize(
    ("array", "pattern"),
    [
        (
            pa.array([datetime(2024, 1, 2, 3, 4, 5)], type=pa.timestamp("us")),
            r"05$",
        ),
        (pa.array([time(3, 4, 5)], type=pa.time64("us")), r"05$"),
        (pa.array([b"a'b"], type=pa.binary()), r'^b"a\'b"$'),
    ],
)
def test_regex_temporal_and_binary_text_matches_python_scalar_formatting(
    array: pa.Array,
    pattern: str,
) -> None:
    compiled = compile_normalization_plan(
        _sample((array,), ("value",)),
        _parse_plan(drop_regex=pattern),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([array], schema=compiled.input_schema)
    )

    assert result.num_rows == 0


def test_unicode_case_insensitive_regex_uses_python_semantics() -> None:
    values = pa.array(["İ", "keep"])
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(drop_regex=r"(?i)i"),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert result.column(0).to_pylist() == ["keep"]


def test_mixed_locale_late_comma_requires_exactly_two_decimal_digits() -> None:
    compiled = compile_normalization_plan(
        _sample((pa.array(["1,234.56", "2.345,67"]),), ("amount",)),
        _parse_plan(batch_size=1),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([pa.array(["3,4"])], schema=compiled.input_schema)
    )

    assert compiled.columns[0].numeric_mode == "mixed_locale"
    assert result.column(0).to_pylist() == [34.0]


def test_sample_budget_counts_unique_physical_backing_buffers() -> None:
    backing = pa.array([b"x" * (MAX_SAMPLE_BYTES + 1), b"ok"], type=pa.binary())
    tiny_logical_slice = backing.slice(1, 1)
    assert tiny_logical_slice.nbytes < 100

    with pytest.raises(ValueError, match="bytes"):
        _sample((tiny_logical_slice,), ("payload",))

    shared = pa.array([b"x" * (MAX_SAMPLE_BYTES // 2 + 1)], type=pa.binary())
    accepted = _sample((shared, shared), ("left", "right"))
    assert accepted.row_count == 1


def test_sample_budget_counts_nested_dictionary_backing_buffers() -> None:
    dictionary = pa.array([b"x" * (MAX_SAMPLE_BYTES + 1)], type=pa.binary())
    encoded = pa.DictionaryArray.from_arrays(pa.array([0], type=pa.int8()), dictionary)
    nested = pa.ListArray.from_arrays(pa.array([0, 1], type=pa.int32()), encoded)

    with pytest.raises(ValueError, match="bytes"):
        _sample((nested,), ("payload",))

    shared_dictionary = pa.array(["x" * (MAX_SAMPLE_BYTES // 2 + 1)])
    left = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        shared_dictionary,
    )
    right = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        shared_dictionary,
    )
    accepted = _sample((left, right), ("left", "right"))
    assert accepted.row_count == 1


def test_sample_buffer_parent_chain_is_bounded_with_configuration_error() -> None:
    values = _deep_parent_buffer_array()

    with pytest.raises(ValueError, match="physical-buffer structural limits"):
        _sample((values,), ("value",))


def test_sample_buffer_accounting_rejects_wide_type_before_eager_walk() -> None:
    wide = pa.StructArray.from_arrays(
        [pa.array([index], type=pa.int16()) for index in range(4_096)],
        names=[f"field_{index}" for index in range(4_096)],
    )

    with pytest.raises(ValueError, match="physical-buffer structural limits"):
        _sample((wide,), ("value",))


def test_sample_buffer_accounting_converts_deep_type_to_configuration_error() -> None:
    deep: pa.Array = pa.array([1], type=pa.int8())
    for _ in range(1_100):
        deep = pa.DictionaryArray.from_arrays(
            pa.array([0], type=pa.int8()),
            deep,
        )

    with pytest.raises(ValueError, match="physical-buffer structural limits"):
        _sample((deep,), ("value",))


def test_dictionary_encoded_strings_match_plain_normalization_across_late_batches() -> None:
    dictionary = pa.array(["  alpha  ", "NA", None])
    encoded = pa.DictionaryArray.from_arrays(
        pa.array([0, 1, 2, None], type=pa.int8()),
        dictionary,
    )
    plain = pa.array(encoded.to_pylist(), type=pa.string())
    encoded_plan = compile_normalization_plan(
        _sample((encoded,), ("value",)),
        _parse_plan(batch_size=4),
    )
    plain_plan = compile_normalization_plan(
        _sample((plain,), ("value",)),
        _parse_plan(batch_size=4),
    )
    encoded_operation = ArrowNormalizationOperation(encoded_plan)
    plain_operation = ArrowNormalizationOperation(plain_plan)

    encoded_result = encoded_operation.normalize(
        pa.record_batch([encoded], schema=encoded_plan.input_schema)
    )
    plain_result = plain_operation.normalize(
        pa.record_batch([plain], schema=plain_plan.input_schema)
    )
    late_dictionary = pa.array(["  beta  ", "NA", None])
    late_encoded = pa.DictionaryArray.from_arrays(
        pa.array([0, 1, 2, None], type=pa.int8()),
        late_dictionary,
    )
    late_plain = pa.array(late_encoded.to_pylist(), type=pa.string())
    encoded_late_result = encoded_operation.normalize(
        pa.record_batch([late_encoded], schema=encoded_plan.input_schema)
    )
    plain_late_result = plain_operation.normalize(
        pa.record_batch([late_plain], schema=plain_plan.input_schema)
    )

    assert encoded_plan.schema == plain_plan.schema == pa.schema([("0", pa.string())])
    assert encoded_result.to_pydict() == plain_result.to_pydict() == {"0": ["alpha"]}
    assert encoded_late_result.to_pydict() == plain_late_result.to_pydict() == {"0": ["beta"]}


def test_shared_string_dictionary_columns_normalize_independently() -> None:
    dictionary = pa.array(["  alpha  ", "NA"])
    left = pa.DictionaryArray.from_arrays(
        pa.array([0, 1, None], type=pa.int8()),
        dictionary,
    )
    right = pa.DictionaryArray.from_arrays(
        pa.array([0, 0, 0], type=pa.int8()),
        dictionary,
    )
    compiled = compile_normalization_plan(
        _sample((left, right), ("left", "right")),
        _parse_plan(batch_size=3),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([left, right], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.string(), pa.string()]
    assert result.to_pydict() == {
        "0": ["alpha", None, None],
        "1": ["alpha", "alpha", "alpha"],
    }


def test_run_end_encoded_strings_match_plain_normalization_across_late_batches() -> None:
    encoded = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2, 4], type=pa.int16()),
        pa.array(["  alpha  ", "NA", None]),
    )
    plain = pa.array(encoded.to_pylist(), type=pa.string())
    encoded_plan = compile_normalization_plan(
        _sample((encoded,), ("value",)),
        _parse_plan(batch_size=4),
    )
    plain_plan = compile_normalization_plan(
        _sample((plain,), ("value",)),
        _parse_plan(batch_size=4),
    )
    encoded_operation = ArrowNormalizationOperation(encoded_plan)
    plain_operation = ArrowNormalizationOperation(plain_plan)

    encoded_result = encoded_operation.normalize(
        pa.record_batch([encoded], schema=encoded_plan.input_schema)
    )
    plain_result = plain_operation.normalize(
        pa.record_batch([plain], schema=plain_plan.input_schema)
    )
    late_encoded = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2, 4], type=pa.int16()),
        pa.array(["  beta  ", "NA", None]),
    )
    late_plain = pa.array(late_encoded.to_pylist(), type=pa.string())
    encoded_late_result = encoded_operation.normalize(
        pa.record_batch([late_encoded], schema=encoded_plan.input_schema)
    )
    plain_late_result = plain_operation.normalize(
        pa.record_batch([late_plain], schema=plain_plan.input_schema)
    )

    assert encoded_plan.schema == plain_plan.schema == pa.schema([("0", pa.string())])
    assert encoded_result.to_pydict() == plain_result.to_pydict() == {"0": ["alpha"]}
    assert encoded_late_result.to_pydict() == plain_late_result.to_pydict() == {"0": ["beta"]}


def test_run_end_encoded_string_locale_inference_matches_plain_evidence() -> None:
    encoded = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array(["1,00", "2,00"]),
    )
    plain = pa.array(encoded.to_pylist(), type=pa.string())
    encoded_plan = compile_normalization_plan(
        _sample((encoded,), ("amount",)),
        _parse_plan(),
    )
    plain_plan = compile_normalization_plan(
        _sample((plain,), ("amount",)),
        _parse_plan(),
    )

    encoded_result = ArrowNormalizationOperation(encoded_plan).normalize(
        pa.record_batch([encoded], schema=encoded_plan.input_schema)
    )

    assert encoded_plan.schema == plain_plan.schema
    assert encoded_plan.columns[0].decimal_separator == ","
    assert encoded_result.column(0).to_pylist() == [1.0, 2.0]


def test_dictionary_of_run_end_encoded_strings_enforces_late_timezone() -> None:
    sample_dictionary = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array(["2024-01-01T00:00:00+08:00", "2024-01-02T00:00:00+08:00"]),
    )
    sample_values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        sample_dictionary,
    )
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("created_at",)),
        _parse_plan(batch_size=2),
    )
    late_dictionary = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        pa.array(["2024-01-03T00:00:00+09:00", "2024-01-04T00:00:00+08:00"]),
    )
    late_values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        late_dictionary,
    )

    with pytest.raises(StreamingTypeError) as captured:
        ArrowNormalizationOperation(compiled).normalize(
            pa.record_batch([late_values], schema=compiled.input_schema)
        )

    assert captured.value.context["ordinal"] == 0
    assert captured.value.context["row_offset"] == 0
    assert compiled.columns[0].timezone == "+08:00"


def test_alternating_encoded_strings_normalize_missing_whitespace_and_nulls() -> None:
    def encoded(dictionary_values: list[str | None]) -> pa.DictionaryArray:
        leaf = pa.DictionaryArray.from_arrays(
            pa.array([0, 1, 2], type=pa.int8()),
            pa.array(dictionary_values, type=pa.string()),
        )
        runs = pa.RunEndEncodedArray.from_arrays(
            pa.array([1, 2, 3], type=pa.int16()),
            leaf,
        )
        return pa.DictionaryArray.from_arrays(
            pa.array([0, 1, 2, None], type=pa.int8()),
            runs,
        )

    sample_values = encoded(["  alpha  ", "NA", None])
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("value",)),
        _parse_plan(batch_size=4),
    )
    late_values = encoded(["  beta  ", "NA", None])
    operation = ArrowNormalizationOperation(compiled)

    sample_result = operation.normalize(
        pa.record_batch([sample_values], schema=compiled.input_schema)
    )
    late_result = operation.normalize(pa.record_batch([late_values], schema=compiled.input_schema))

    assert compiled.schema == pa.schema([("0", pa.string())])
    assert sample_result.column(0).to_pylist() == ["alpha"]
    assert late_result.column(0).to_pylist() == ["beta"]


def test_alternating_encoded_string_locale_matches_plain_values() -> None:
    leaf = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        pa.array(["1,00", "2,00"]),
    )
    runs = pa.RunEndEncodedArray.from_arrays(
        pa.array([1, 2], type=pa.int16()),
        leaf,
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([0, 1], type=pa.int8()),
        runs,
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("amount",)),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.columns[0].decimal_separator == ","
    assert result.column(0).to_pylist() == [1.0, 2.0]


def test_normalize_false_preserves_alternating_encoded_string_identity() -> None:
    leaf = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        pa.array(["  unchanged  "]),
    )
    runs = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        leaf,
    )
    values = pa.DictionaryArray.from_arrays(
        pa.array([0], type=pa.int8()),
        runs,
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(normalize=False),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert result is batch


def test_normalize_false_preserves_run_end_encoded_string_identity() -> None:
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([2], type=pa.int16()),
        pa.array(["  unchanged  "]),
    )
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(normalize=False),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert compiled.schema == compiled.input_schema
    assert result is batch
    assert result.column(0).type == values.type


def test_sample_budget_counts_run_end_encoded_string_value_buffers() -> None:
    values = pa.RunEndEncodedArray.from_arrays(
        pa.array([1], type=pa.int16()),
        pa.array(["x" * (MAX_SAMPLE_BYTES + 1)]),
    )

    with pytest.raises(ValueError, match="bytes"):
        _sample((values,), ("value",))


@pytest.mark.parametrize(
    "filter_config",
    [
        {},
        {"drop_regex": r"^999$"},
        {"drop_conditions": [{"column": "value", "value": 999}]},
    ],
)
def test_all_keep_row_filters_preserve_original_batch_and_buffers(
    filter_config: dict[str, object],
) -> None:
    values = pa.array([1, 2], type=pa.int64())
    compiled = compile_normalization_plan(
        _sample((values,), ("value",)),
        _parse_plan(**filter_config),
    )
    batch = pa.record_batch([values], schema=compiled.input_schema)

    result = ArrowNormalizationOperation(compiled).normalize(batch)

    assert result is batch
    assert result.column(0).buffers()[1].address == values.buffers()[1].address


def test_uniform_timezone_text_hint_preserves_fixed_offset_schema_and_values() -> None:
    values = pa.array(["2024-01-01T12:00:00+08:00", "2024-01-02T13:30:00+08:00"])
    compiled = compile_normalization_plan(
        _sample((values,), ("created_at",)),
        _parse_plan(type_hints={"created_at": "TIMESTAMP"}),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.timestamp("ns", tz="+08:00")]
    assert [value.isoformat() for value in result.column(0).to_pylist()] == [
        "2024-01-01T12:00:00+08:00",
        "2024-01-02T13:30:00+08:00",
    ]


@pytest.mark.parametrize(
    "incompatible",
    [
        "2024-01-04T15:00:00",
        "2024-01-04T15:00:00+09:00",
    ],
)
def test_late_timestamp_text_must_match_compiled_fixed_timezone_policy(
    incompatible: str,
) -> None:
    sample_values = pa.array(["2024-01-01T12:00:00+08:00", "2024-01-02T13:30:00+08:00"])
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("created_at",)),
        _parse_plan(type_hints={"created_at": "TIMESTAMP"}),
    )
    compatible = pa.record_batch(
        [pa.array(["2024-01-03T14:00:00+08:00"])],
        schema=compiled.input_schema,
    )
    late = pa.record_batch([pa.array([incompatible])], schema=compiled.input_schema)
    raw = _Reader(compiled.input_schema, [compatible, late])
    reader = NormalizedStreamingReader(raw, compiled)

    assert reader.read_next_batch() is not None
    with pytest.raises(StreamingTypeError) as captured:
        reader.read_next_batch()

    assert captured.value.context["row_offset"] == 1
    assert captured.value.context["expected_type"] == "timestamp[ns, tz=+08:00]"
    assert raw.close_calls == 1


def test_late_timestamp_text_with_the_compiled_fixed_offset_remains_compatible() -> None:
    sample_values = pa.array(["2024-01-01T12:00:00+08:00", "2024-01-02T13:30:00+08:00"])
    compiled = compile_normalization_plan(
        _sample((sample_values,), ("created_at",)),
        _parse_plan(type_hints={"created_at": "TIMESTAMP"}),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch(
            [pa.array(["2024-01-03T14:00:00+08:00"])],
            schema=compiled.input_schema,
        )
    )

    assert result.column(0)[0].as_py().isoformat() == "2024-01-03T14:00:00+08:00"


def test_mixed_timezone_text_hint_is_rejected_during_compilation() -> None:
    values = pa.array(["2024-01-01T12:00:00+08:00", "2024-01-02T13:30:00+09:00"])

    with pytest.raises(ValueError, match="timezone"):
        compile_normalization_plan(
            _sample((values,), ("created_at",)),
            _parse_plan(type_hints={"created_at": "TIMESTAMP"}),
        )


def test_timestamp_hint_preserves_compatible_observed_timezone_and_unit() -> None:
    timestamp_type = pa.timestamp("us", tz="Asia/Kuala_Lumpur")
    values = pa.array([datetime(2024, 1, 1, 12)], type=timestamp_type)
    compiled = compile_normalization_plan(
        _sample((values,), ("created_at",)),
        _parse_plan(type_hints={"created_at": "TIMESTAMP"}),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [timestamp_type]
    assert result.column(0)[0].as_py() == values[0].as_py()


def test_inferred_uniform_timezone_text_preserves_fixed_offset() -> None:
    values = pa.array(["2024-01-01T12:00:00+08:00", "2024-01-02T13:30:00+08:00"])
    compiled = compile_normalization_plan(
        _sample((values,), ("created_at",)),
        _parse_plan(),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.timestamp("us", tz="+08:00")]
    assert result.column(0)[0].as_py().isoformat() == "2024-01-01T12:00:00+08:00"


def test_known_invalid_date_sample_evidence_does_not_compile_strict_date() -> None:
    values = pa.array([*(f"2024-01-{day:02d}" for day in range(1, 10)), "not-a-date"])
    compiled = compile_normalization_plan(
        _sample((values,), ("event_date",)),
        _parse_plan(batch_size=10),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.string()]
    assert result.column(0)[-1].as_py() == "not-a-date"


@pytest.mark.parametrize(
    "values",
    [
        ["1", "2", "NA"],
        ["2024-01-01", "NA"],
    ],
)
def test_missing_markers_are_applied_after_numeric_and_date_inference(
    values: list[str],
) -> None:
    evidence = pa.array(values)
    compiled = compile_normalization_plan(
        _sample((evidence,), ("value",)),
        _parse_plan(batch_size=len(values)),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([evidence], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.string()]
    assert result.column(0).to_pylist() == values[:-1]


def test_known_invalid_numeric_date_sample_stays_numeric() -> None:
    values = pa.array([*range(1, 10), 60_001], type=pa.int64())
    compiled = compile_normalization_plan(
        _sample((values,), ("event_date",)),
        _parse_plan(batch_size=10),
    )

    result = ArrowNormalizationOperation(compiled).normalize(
        pa.record_batch([values], schema=compiled.input_schema)
    )

    assert compiled.schema.types == [pa.int64()]
    assert result.column(0)[-1].as_py() == 60_001
