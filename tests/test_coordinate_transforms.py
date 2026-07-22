"""Coordinate-aware Arrow transform contracts."""

from __future__ import annotations

import gc
import inspect
import json
import weakref
from pathlib import Path
from typing import Any

import pyarrow as pa
import pytest

import messy_xlsx.parsing.coordinates as coordinates_module
from messy_xlsx import MessyWorkbook, SheetConfig
from messy_xlsx.enums import MergeStrategy
from messy_xlsx.ooxml.models import (
    Interval,
    IntervalIndex,
    MergeRange,
    SheetManifest,
)
from messy_xlsx.parsing.contracts import OutputMode
from messy_xlsx.parsing.coordinates import (
    ColumnIdentity,
    CoordinateBatch,
    CoordinateCompatibilityError,
    CoordinateOperation,
    CoordinateTransform,
)
from messy_xlsx.parsing.parse_plan import ParsePlan, compile_parse_plan


def _plan(**overrides: Any) -> ParsePlan:
    values: dict[str, Any] = {
        "auto_detect": False,
        "include_hidden": True,
        "merge_strategy": MergeStrategy.SKIP,
        "evaluate_formulas": True,
        "header_rows": 0,
    }
    values.update(overrides)
    return compile_parse_plan(
        SheetConfig(**values),
        structure=None,
        format_type="xlsx",
        output_mode=OutputMode.MATERIALIZED,
        batch_size=None,
    )


def _batch(
    row_numbers: tuple[int, ...] = (1, 2),
    *,
    identities: tuple[ColumnIdentity, ...] = (),
) -> CoordinateBatch:
    return CoordinateBatch(
        batch=pa.record_batch(
            [list(row_numbers), [number * 10 for number in row_numbers]],
            names=["0", "1"],
        ),
        row_numbers=pa.array(row_numbers, type=pa.int64()),
        column_numbers=(1, 2),
        column_identities=identities,
    )


def _transform() -> CoordinateTransform:
    return CoordinateTransform(
        hidden_rows=IntervalIndex(()),
        hidden_columns=IntervalIndex(()),
        merged_ranges=(),
    )


def _grid_batch(
    rows: tuple[int, ...] = (1, 2, 3),
    columns: tuple[int, ...] = (1, 2, 3),
) -> CoordinateBatch:
    return CoordinateBatch(
        batch=pa.record_batch(
            [[row * 10 + column for row in rows] for column in columns],
            names=[str(index) for index in range(len(columns))],
        ),
        row_numbers=pa.array(rows, type=pa.int64()),
        column_numbers=columns,
    )


def _run(
    transform: CoordinateTransform,
    plan: ParsePlan,
    *batches: CoordinateBatch,
) -> tuple[CoordinateBatch, ...]:
    operation = transform.open(plan)
    emitted = tuple(output for batch in batches for output in operation.push(batch))
    return (*emitted, *operation.finish())


def _result_contract(
    result: tuple[CoordinateBatch, ...],
) -> tuple[tuple[int, ...], tuple[int, ...], tuple[tuple[object, ...], ...]]:
    if not result:
        return (), (), ()
    rows = tuple(coordinate.as_py() for batch in result for coordinate in batch.row_numbers)
    columns = result[0].column_numbers
    values = tuple(
        tuple(batch.batch.column(column)[row].as_py() for column in range(batch.batch.num_columns))
        for batch in result
        for row in range(batch.batch.num_rows)
    )
    return rows, columns, values


def test_coordinate_batch_slices_rows_and_preserves_original_columns() -> None:
    identities = (ColumnIdentity(0, "left"), ColumnIdentity(1, "right"))

    result = _batch(identities=identities).slice_rows(1, 1)

    assert result.batch.to_pydict() == {"0": [2], "1": [20]}
    assert result.row_numbers.to_pylist() == [2]
    assert result.column_numbers == (1, 2)
    assert result.column_identities == identities


def test_coordinate_batch_rejects_row_length_mismatch() -> None:
    with pytest.raises(
        ValueError,
        match=r"^row coordinate count does not match batch rows$",
    ):
        CoordinateBatch(
            batch=pa.record_batch([[1, 2]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )


def test_coordinate_batch_rejects_column_length_mismatch() -> None:
    with pytest.raises(
        ValueError,
        match=r"^column coordinate count does not match batch columns$",
    ):
        CoordinateBatch(
            batch=pa.record_batch([[1]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1, 2),
        )


def test_coordinate_operation_rejects_out_of_order_batches() -> None:
    operation = _transform().open(_plan())
    operation.push(_batch((3, 4)))

    with pytest.raises(ValueError, match=r"^coordinate batches are out of order$"):
        operation.push(_batch((2,)))


def test_coordinate_operation_rejects_null_row_coordinates() -> None:
    operation = _transform().open(_plan())
    batch = CoordinateBatch(
        batch=pa.record_batch([[1, 2]], names=["0"]),
        row_numbers=pa.array([1, None], type=pa.int64()),
        column_numbers=(1,),
    )

    with pytest.raises(ValueError, match=r"^row coordinates cannot contain nulls$"):
        operation.push(batch)


@pytest.mark.parametrize("row_number", [0, -1])
def test_coordinate_operation_rejects_nonpositive_row_coordinates(
    row_number: int,
) -> None:
    operation = _transform().open(_plan())

    with pytest.raises(ValueError, match=r"^row coordinates must be positive$"):
        operation.push(_batch((row_number,)))


def test_coordinate_operation_does_not_materialize_coordinate_sidecar() -> None:
    source = inspect.getsource(CoordinateOperation.push)

    assert "to_pylist" not in source


def test_coordinate_operation_finish_is_idempotent() -> None:
    operation = _transform().open(_plan())

    assert operation.finish() == ()
    assert operation.finish() == ()


def test_coordinate_operation_rejects_push_after_finish() -> None:
    operation = _transform().open(_plan())
    operation.finish()

    with pytest.raises(RuntimeError, match=r"^coordinate operation is finished$"):
        operation.push(_batch())


RANGE_AND_HIDDEN_CASES = (
    (None, True, (1, 3), (1, 3)),
    ("A1:C3", True, (1, 2, 3), (1, 2, 3)),
    ("B2:C3", True, (2, 3), (2, 3)),
    ("$B$2:$C$3", True, (2, 3), (2, 3)),
    ("A1:C3", False, (1, 2, 3), (1, 2, 3)),
)


@pytest.mark.parametrize(
    ("cell_range", "ignore_hidden", "expected_rows", "expected_columns"),
    RANGE_AND_HIDDEN_CASES,
)
def test_range_and_hidden_precedence(
    cell_range: str | None,
    ignore_hidden: bool,
    expected_rows: tuple[int, ...],
    expected_columns: tuple[int, ...],
) -> None:
    from messy_xlsx.ooxml.models import Interval

    transform = CoordinateTransform(
        hidden_rows=IntervalIndex((Interval(2, 2),)),
        hidden_columns=IntervalIndex((Interval(2, 2),)),
        merged_ranges=(),
    )

    result = _run(
        transform,
        _plan(cell_range=cell_range, include_hidden=not ignore_hidden),
        _grid_batch(),
    )

    rows, columns, values = _result_contract(result)
    assert rows == expected_rows
    assert columns == expected_columns
    assert values == tuple(
        tuple(row * 10 + column for column in expected_columns) for row in expected_rows
    )


def test_hidden_row_filtering_does_not_scan_the_batch_per_interval(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    intervals = tuple(Interval(row, row) for row in range(1, 4_001, 2))
    transform = CoordinateTransform(
        hidden_rows=IntervalIndex(intervals),
        hidden_columns=IntervalIndex(()),
        merged_ranges=(),
    )
    calls = 0
    original = coordinates_module.pc.greater_equal

    def recording_greater_equal(*args: object, **kwargs: object) -> object:
        nonlocal calls
        calls += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(coordinates_module.pc, "greater_equal", recording_greater_equal)
    result = _run(
        transform,
        _plan(include_hidden=False),
        _grid_batch(rows=tuple(range(1, 4_001)), columns=(1,)),
    )

    assert calls < 10
    assert _result_contract(result)[0] == tuple(range(2, 4_001, 2))


def test_hidden_row_filtering_builds_only_each_batch_overlap_window(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    intervals = tuple(Interval(row, row) for row in range(10_001, 12_001, 2))
    transform = CoordinateTransform(
        hidden_rows=IntervalIndex(intervals),
        hidden_columns=IntervalIndex(()),
        merged_ranges=(),
    )
    requested_counts: list[int] = []
    original = coordinates_module.np.fromiter

    def recording_fromiter(*args: object, **kwargs: object) -> object:
        requested_counts.append(int(kwargs["count"]))
        return original(*args, **kwargs)

    monkeypatch.setattr(coordinates_module.np, "fromiter", recording_fromiter)
    operation = transform.open(_plan(include_hidden=False))
    first = operation.push(_grid_batch(rows=tuple(range(1, 101)), columns=(1,)))
    second = operation.push(_grid_batch(rows=tuple(range(101, 201)), columns=(1,)))

    assert requested_counts == []
    assert _result_contract((*first, *second))[0] == tuple(range(1, 201))


def test_single_cell_range_preserves_original_coordinate() -> None:
    result = _run(_transform(), _plan(cell_range="B2:B2"), _grid_batch())

    assert _result_contract(result) == ((2,), (2,), ((22,),))


def test_bare_single_cell_range_preserves_legacy_invalid_range_behavior() -> None:
    with pytest.raises(ValueError, match=r"^B2 is not a valid coordinate or range$"):
        _transform().open(_plan(cell_range="B2"))


@pytest.mark.parametrize(
    ("cell_range", "message"),
    [
        ("A0:B1", "cell range coordinates must be one-based"),
        ("B2:A1", "cell range coordinates must be ordered"),
        ("XFE1:XFE2", "cell range exceeds Excel worksheet bounds"),
        ("A1048577:A1048577", "cell range exceeds Excel worksheet bounds"),
    ],
)
def test_range_projection_rejects_invalid_or_unbounded_coordinates(
    cell_range: str,
    message: str,
) -> None:
    with pytest.raises(ValueError, match=f"^{message}$"):
        _transform().open(_plan(cell_range=cell_range))


def test_leading_blank_rows_remain_addressable_by_range() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([None, None, "value"], type=pa.string())],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _run(_transform(), _plan(cell_range="A1:A3"), raw)

    assert _result_contract(result) == (
        (1, 2, 3),
        (1,),
        ((None,), (None,), ("value",)),
    )


def test_empty_raw_input_does_not_invent_a_row() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch([], names=[]),
        row_numbers=pa.array([], type=pa.int64()),
        column_numbers=(),
    )

    result = _run(_transform(), _plan(), raw)

    assert _result_contract(result) == ((), (), ())


def test_range_beyond_observed_dimensions_is_null_padded() -> None:
    result = _run(
        _transform(),
        _plan(cell_range="A1:C3"),
        _grid_batch(rows=(1, 2), columns=(1, 2)),
    )

    assert _result_contract(result) == (
        (1, 2, 3),
        (1, 2, 3),
        ((11, 12, None), (21, 22, None), (None, None, None)),
    )


def test_range_selection_preserves_schema_until_global_materialized_inference() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([1.0, None], type=pa.float64())],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _run(_transform(), _plan(cell_range="A2:A2"), raw)

    assert result[0].batch.column(0).type == pa.float64()
    assert _result_contract(result)[2] == ((None,),)


def test_range_batches_keep_one_physical_schema_before_global_inference() -> None:
    operation = _transform().open(_plan(cell_range="A1:A2"))
    first = operation.push(
        CoordinateBatch(
            batch=pa.record_batch(
                [pa.array([None], type=pa.float64())],
                names=["0"],
            ),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )
    )
    second = operation.push(
        CoordinateBatch(
            batch=pa.record_batch([[2.0]], names=["0"]),
            row_numbers=pa.array([2], type=pa.int64()),
            column_numbers=(1,),
        )
    )

    result = (*first, *second, *operation.finish())
    assert {batch.batch.schema for batch in result} == {pa.schema([pa.field("0", pa.float64())])}
    assert _result_contract(result)[2] == ((None,), (2.0,))


def test_range_operation_never_emits_incompatible_schema_for_late_column() -> None:
    operation = _transform().open(_plan(cell_range="A1:B2", header_rows=0))
    first = operation.push(
        CoordinateBatch(
            batch=pa.record_batch([[11]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )
    )
    try:
        second = operation.push(
            CoordinateBatch(
                batch=pa.record_batch([[21], [22]], names=["0", "1"]),
                row_numbers=pa.array([2], type=pa.int64()),
                column_numbers=(1, 2),
            )
        )
    except CoordinateCompatibilityError as error:
        assert str(error) == "coordinate input schema changed across batches"
    else:
        table = pa.Table.from_batches(
            [batch.batch for batch in (*first, *second, *operation.finish())]
        )
        assert table.to_pydict() == {"0": [11, 21], "1": [None, 22]}


def test_coordinate_operation_rejects_field_type_drift_before_output() -> None:
    operation = _transform().open(_plan(header_rows=0))
    operation.push(
        CoordinateBatch(
            batch=pa.record_batch([[11]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )
    )

    with pytest.raises(
        CoordinateCompatibilityError,
        match=r"^coordinate input schema changed across batches$",
    ):
        operation.push(
            CoordinateBatch(
                batch=pa.record_batch([["text"]], names=["0"]),
                row_numbers=pa.array([2], type=pa.int64()),
                column_numbers=(1,),
            )
        )


def test_malformed_range_is_rejected_when_operation_opens() -> None:
    with pytest.raises(
        ValueError,
        match=r"^invalid-range is not a valid coordinate or range$",
    ):
        _transform().open(_plan(cell_range="invalid-range"))


def test_malformed_range_preserves_frozen_public_error_contract() -> None:
    from tests.compatibility._contract import exception_contract

    root = Path(__file__).resolve().parents[1]
    expected = json.loads(
        (root / "tests/compatibility/golden/v010-errors.json").read_text(encoding="utf-8")
    )["invalid_range"]
    with MessyWorkbook(root / "tests/samples/accounts_receivable.xlsx") as workbook:
        actual = exception_contract(
            lambda: workbook.to_dataframe(
                config=SheetConfig(auto_detect=False, cell_range="invalid-range")
            )
        )

    assert actual == expected


@pytest.mark.parametrize("strategy", list(MergeStrategy))
@pytest.mark.parametrize(
    "merged_range",
    [
        MergeRange(1, 1, 1, 2),
        MergeRange(1, 1, 2, 1),
        MergeRange(1, 1, 2, 2),
    ],
    ids=["horizontal", "vertical", "two-dimensional"],
)
def test_merge_strategy_preserves_legacy_values(
    strategy: MergeStrategy,
    merged_range: MergeRange,
) -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [["a1", "a2"], ["b1", "b2"]],
            names=["0", "1"],
        ),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1, 2),
    )
    transform = CoordinateTransform(IntervalIndex(()), IntervalIndex(()), (merged_range,))

    result = _run(transform, _plan(merge_strategy=strategy), raw)

    original = (("a1", "b1"), ("a2", "b2"))
    if strategy is MergeStrategy.SKIP:
        expected = original
    elif strategy is MergeStrategy.FILL:
        expected = tuple(
            tuple(
                "a1"
                if merged_range.min_row <= row <= merged_range.max_row
                and merged_range.min_col <= column <= merged_range.max_col
                else original[row - 1][column - 1]
                for column in (1, 2)
            )
            for row in (1, 2)
        )
    else:
        expected = tuple(
            tuple(
                None
                if merged_range.min_row <= row <= merged_range.max_row
                and merged_range.min_col <= column <= merged_range.max_col
                and (row, column) != (merged_range.min_row, merged_range.min_col)
                else original[row - 1][column - 1]
                for column in (1, 2)
            )
            for row in (1, 2)
        )
    assert _result_contract(result)[2] == expected


def test_fill_carries_anchor_across_batches() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 2, 1),),
    )
    operation = transform.open(_plan(merge_strategy=MergeStrategy.FILL))

    first = operation.push(_grid_batch(rows=(1,), columns=(1,)))
    second = (*operation.push(_grid_batch(rows=(2,), columns=(1,))), *operation.finish())

    assert _result_contract(first)[2] == ((11,),)
    assert _result_contract(second)[2] == ((11,),)

    independent = transform.open(_plan(merge_strategy=MergeStrategy.FILL))
    independent_first = independent.push(
        CoordinateBatch(
            batch=pa.record_batch([[99]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )
    )
    independent_second = independent.push(_grid_batch(rows=(2,), columns=(1,)))
    assert _result_contract((*independent_first, *independent_second))[2] == (
        (99,),
        (99,),
    )


def test_fill_keeps_promoted_null_follower_schema_after_merge_ends() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    operation = transform.open(_plan(merge_strategy=MergeStrategy.FILL))
    raw_schema = pa.schema([pa.field("0", pa.int64()), pa.field("1", pa.null())])

    first = operation.push(
        CoordinateBatch(
            batch=pa.RecordBatch.from_arrays(
                [pa.array([11], type=pa.int64()), pa.nulls(1)],
                schema=raw_schema,
            ),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1, 2),
        )
    )
    second = operation.push(
        CoordinateBatch(
            batch=pa.RecordBatch.from_arrays(
                [pa.array([21], type=pa.int64()), pa.nulls(1)],
                schema=raw_schema,
            ),
            row_numbers=pa.array([2], type=pa.int64()),
            column_numbers=(1, 2),
        )
    )

    table = pa.Table.from_batches([batch.batch for batch in (*first, *second, *operation.finish())])
    assert table.to_pydict() == {"0": [11, 21], "1": [11, None]}


def test_fill_keeps_projected_synthetic_follower_schema_after_merge_ends() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    operation = transform.open(_plan(cell_range="B1:B2", merge_strategy=MergeStrategy.FILL))
    raw_schema = pa.schema([pa.field("0", pa.int64())])

    emitted = tuple(
        output
        for row, value in ((1, 11), (2, 21))
        for output in operation.push(
            CoordinateBatch(
                batch=pa.RecordBatch.from_arrays(
                    [pa.array([value], type=pa.int64())],
                    schema=raw_schema,
                ),
                row_numbers=pa.array([row], type=pa.int64()),
                column_numbers=(1,),
            )
        )
    )

    table = pa.Table.from_batches([batch.batch for batch in (*emitted, *operation.finish())])
    assert table.to_pydict() == {"0": [11, None]}


def test_fill_rejects_conflicting_null_follower_promotions_before_output() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (
            MergeRange(1, 1, 1, 3),
            MergeRange(2, 2, 2, 3),
        ),
    )
    operation = transform.open(_plan(merge_strategy=MergeStrategy.FILL))
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [
                pa.array([11, 21], type=pa.int64()),
                pa.array([1.5, 2.5], type=pa.float64()),
                pa.nulls(2),
            ],
            names=["0", "1", "2"],
        ),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1, 2, 3),
    )

    with pytest.raises(
        CoordinateCompatibilityError,
        match=r"^merged-cell fill cannot establish one Arrow schema$",
    ):
        operation.push(raw)


def test_projected_out_anchor_fills_requested_follower() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch([["anchor"], [None]], names=["0", "1"]),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1, 2),
    )

    result = _run(
        transform,
        _plan(cell_range="B1:B1", merge_strategy=MergeStrategy.FILL),
        raw,
    )

    assert _result_contract(result) == ((1,), (2,), (("anchor",),))


def test_hidden_anchor_fills_visible_follower_before_filtering() -> None:
    transform = CoordinateTransform(
        IntervalIndex((Interval(1, 1),)),
        IntervalIndex(()),
        (MergeRange(1, 1, 2, 1),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch([["anchor", None]], names=["0"]),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _run(
        transform,
        _plan(include_hidden=False, merge_strategy=MergeStrategy.FILL),
        raw,
    )

    assert _result_contract(result) == ((2,), (1,), (("anchor",),))


def test_auxiliary_anchor_row_and_column_are_removed() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 2, 2),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [["anchor", None, "a3"], [None, None, "b3"], ["c1", "c2", "c3"]],
            names=["0", "1", "2"],
        ),
        row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        column_numbers=(1, 2, 3),
    )

    result = _run(
        transform,
        _plan(cell_range="B2:C3", merge_strategy=MergeStrategy.FILL),
        raw,
    )

    assert _result_contract(result) == (
        (2, 3),
        (2, 3),
        (("anchor", "c2"), ("b3", "c3")),
    )


def test_completed_merge_anchor_state_is_released() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 2, 1),),
    )
    operation = transform.open(_plan(merge_strategy=MergeStrategy.FILL))

    operation.push(_grid_batch(rows=(1,), columns=(1,)))
    assert operation._active_anchors
    operation.push(_grid_batch(rows=(2,), columns=(1,)))

    assert not operation._active_anchors


def test_mixed_type_merge_fill_raises_exact_compatibility_signal() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [
                pa.array(["anchor"], type=pa.string()),
                pa.array([None], type=pa.int64()),
            ],
            names=["0", "1"],
        ),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1, 2),
    )

    operation = transform.open(_plan(merge_strategy=MergeStrategy.FILL))
    with pytest.raises(CoordinateCompatibilityError):
        operation.push(raw)


def test_typed_null_anchor_does_not_preemptively_reject_non_null_follower() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [
                pa.array([None], type=pa.string()),
                pa.array([7], type=pa.int64()),
            ],
            names=["0", "1"],
        ),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1, 2),
    )

    result = _run(transform, _plan(merge_strategy=MergeStrategy.FILL), raw)

    assert _result_contract(result)[2] == ((None, None),)
    assert result[0].batch.schema.types == [pa.string(), pa.int64()]


@pytest.mark.parametrize("anchor", [1.0, 1.5])
def test_float_merge_anchor_never_coerces_to_integer_follower(
    anchor: float,
) -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [
                pa.array([anchor], type=pa.float64()),
                pa.array([None], type=pa.int64()),
            ],
            names=["0", "1"],
        ),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1, 2),
    )

    with pytest.raises(CoordinateCompatibilityError):
        transform.open(_plan(merge_strategy=MergeStrategy.FILL)).push(raw)


def test_large_integer_merge_anchor_does_not_lose_precision_in_float_column() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 1, 2),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [
                pa.array([2**53 + 1], type=pa.int64()),
                pa.array([None], type=pa.float64()),
            ],
            names=["0", "1"],
        ),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1, 2),
    )

    with pytest.raises(CoordinateCompatibilityError):
        transform.open(_plan(merge_strategy=MergeStrategy.FILL)).push(raw)


def test_merge_fill_applies_to_trailing_synthetic_range_coordinates() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 3, 3),),
    )
    raw = CoordinateBatch(
        batch=pa.record_batch([["anchor"]], names=["0"]),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _run(
        transform,
        _plan(cell_range="A1:C3", merge_strategy=MergeStrategy.FILL),
        raw,
    )

    assert _result_contract(result)[2] == (("anchor",) * 3,) * 3


def test_merge_fill_applies_to_synthetic_row_between_batches() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 3, 1),),
    )
    operation = transform.open(_plan(cell_range="A1:A3", merge_strategy=MergeStrategy.FILL))
    first = operation.push(
        CoordinateBatch(
            batch=pa.record_batch([["anchor"]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )
    )
    second = operation.push(
        CoordinateBatch(
            batch=pa.record_batch(
                [pa.array([None], type=pa.string())],
                names=["0"],
            ),
            row_numbers=pa.array([3], type=pa.int64()),
            column_numbers=(1,),
        )
    )

    result = (*first, *second, *operation.finish())
    assert _result_contract(result)[2] == (("anchor",),) * 3


def test_finish_releases_anchor_state_after_terminal_synthetic_processing() -> None:
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        (MergeRange(1, 1, 5, 1),),
    )
    operation = transform.open(_plan(cell_range="A1:A2", merge_strategy=MergeStrategy.FILL))
    operation.push(
        CoordinateBatch(
            batch=pa.record_batch([["anchor"]], names=["0"]),
            row_numbers=pa.array([1], type=pa.int64()),
            column_numbers=(1,),
        )
    )

    operation.finish()

    assert not operation._active_anchors
    assert operation.finish() == ()


def _mixed_string_int_array(values: tuple[str | int, ...]) -> pa.UnionArray:
    strings: list[str] = []
    integers: list[int] = []
    type_codes: list[int] = []
    offsets: list[int] = []
    for value in values:
        if isinstance(value, str):
            type_codes.append(0)
            offsets.append(len(strings))
            strings.append(value)
        else:
            type_codes.append(1)
            offsets.append(len(integers))
            integers.append(value)
    return pa.UnionArray.from_dense(
        pa.array(type_codes, type=pa.int8()),
        pa.array(offsets, type=pa.int32()),
        [pa.array(strings), pa.array(integers, type=pa.int64())],
        field_names=["string", "integer"],
    )


def _row_batch(rows: tuple[tuple[object, ...], ...]) -> CoordinateBatch:
    columns = tuple(zip(*rows, strict=True)) if rows else ()
    arrays = [
        _mixed_string_int_array(column)  # type: ignore[arg-type]
        if {type(value) for value in column} == {str, int}
        else pa.array(column)
        for column in columns
    ]
    return CoordinateBatch(
        batch=pa.record_batch(
            arrays,
            names=[str(index) for index in range(len(columns))],
        ),
        row_numbers=pa.array(range(1, len(rows) + 1), type=pa.int64()),
        column_numbers=tuple(range(1, len(columns) + 1)),
    )


@pytest.mark.parametrize(
    (
        "cell_range",
        "skip_rows",
        "header_rows",
        "skip_footer",
        "expected_labels",
        "expected_rows",
    ),
    [
        (None, 1, 1, 1, ("h1", "h2"), ((1, 2),)),
        (
            "A1:B4",
            1,
            1,
            1,
            ("skip-me", "skip-me-too"),
            (("h1", "h2"), (1, 2)),
        ),
        (
            None,
            0,
            0,
            1,
            ("col_0", "col_1"),
            (("skip-me", "skip-me-too"), ("h1", "h2"), (1, 2)),
        ),
        (
            None,
            0,
            2,
            0,
            ("skip-me__h1", "skip-me-too__h2"),
            ((1, 2), ("footer", "footer")),
        ),
    ],
)
def test_row_framing_precedence(
    cell_range: str | None,
    skip_rows: int,
    header_rows: int,
    skip_footer: int,
    expected_labels: tuple[str, ...],
    expected_rows: tuple[tuple[object, ...], ...],
) -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [
                _mixed_string_int_array(("skip-me", "h1", 1, "footer")),
                _mixed_string_int_array(("skip-me-too", "h2", 2, "footer")),
            ],
            names=["0", "1"],
        ),
        row_numbers=pa.array([1, 2, 3, 4], type=pa.int64()),
        column_numbers=(1, 2),
    )

    result = _run(
        _transform(),
        _plan(
            cell_range=cell_range,
            skip_rows=skip_rows,
            header_rows=header_rows,
            skip_footer=skip_footer,
        ),
        raw,
    )

    assert result[0].column_identities == tuple(
        ColumnIdentity(index, label) for index, label in enumerate(expected_labels)
    )
    assert _result_contract(result)[2] == expected_rows


def test_range_operation_ignores_compiled_skip_rows() -> None:
    result = _run(
        _transform(),
        _plan(cell_range="A1:A2", skip_rows=1, header_rows=0),
        CoordinateBatch(
            batch=pa.record_batch([["first", "second"]], names=["0"]),
            row_numbers=pa.array([1, 2], type=pa.int64()),
            column_numbers=(1,),
        ),
    )

    assert _result_contract(result)[2] == (("first",), ("second",))


def test_coordinate_operation_does_not_retain_manifest_on_plan() -> None:
    manifest = SheetManifest(
        name="Data",
        target="xl/worksheets/sheet1.xml",
        declared_dimension=(1, 1, 1, 1),
        observed_max_row=1,
        observed_max_col=1,
        hidden_rows=IntervalIndex(()),
        hidden_columns=IntervalIndex(()),
        merged_ranges=(),
        has_formulas=False,
        formula_samples=(),
    )
    manifest_reference = weakref.ref(manifest)
    transform = CoordinateTransform.from_manifest(manifest)
    plan = _plan()
    operation = transform.open(plan)

    del manifest
    gc.collect()

    assert manifest_reference() is None
    assert operation._plan is plan
    assert not hasattr(plan, "manifest")


def test_insufficient_multi_row_header_retains_rows_with_generic_labels() -> None:
    result = _run(
        _transform(),
        _plan(header_rows=2),
        _row_batch((("only", "row"),)),
    )

    assert result[0].column_identities == (
        ColumnIdentity(0, "col_0"),
        ColumnIdentity(1, "col_1"),
    )
    assert _result_contract(result)[2] == (("only", "row"),)


def test_footer_rows_split_across_batches_are_removed_terminally() -> None:
    operation = _transform().open(_plan(header_rows=0, skip_footer=2))

    first = operation.push(_row_batch((("data",), ("footer-1",))))
    second_raw = CoordinateBatch(
        batch=pa.record_batch([["footer-2"]], names=["0"]),
        row_numbers=pa.array([3], type=pa.int64()),
        column_numbers=(1,),
    )
    second = (*operation.push(second_raw), *operation.finish())

    assert _result_contract((*first, *second))[2] == (("data",),)


def test_duplicate_header_labels_retain_positional_identities() -> None:
    result = _run(
        _transform(),
        _plan(header_rows=1),
        _row_batch((("duplicate", "duplicate"), (1, 2))),
    )

    assert result[0].column_identities == (
        ColumnIdentity(0, "duplicate"),
        ColumnIdentity(1, "duplicate"),
    )
    assert _result_contract(result)[2] == ((1, 2),)


def test_null_header_uses_generic_positional_label() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([None, "value"], type=pa.string())],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _run(_transform(), _plan(header_rows=1), raw)

    assert result[0].column_identities == (ColumnIdentity(0, "col_0"),)


def test_numeric_header_preserves_arrow_scalar_string_form() -> None:
    result = _run(
        _transform(),
        _plan(header_rows=1),
        _row_batch(((1,), (2,), (3,))),
    )

    assert result[0].column_identities == (ColumnIdentity(0, "1"),)
    assert result[0].batch.column(0).type == pa.int64()
    assert _result_contract(result)[2] == ((2,), (3,))


def test_numeric_header_reflects_null_upcast_before_header_extraction() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([1.0, None, 3.0], type=pa.float64())],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _run(_transform(), _plan(header_rows=1), raw)

    assert result[0].column_identities == (ColumnIdentity(0, "1.0"),)
    assert result[0].batch.column(0).type == pa.float64()
    assert _result_contract(result)[2] == ((None,), (3.0,))


def test_whitespace_in_multi_row_headers_is_stripped_before_joining() -> None:
    result = _run(
        _transform(),
        _plan(header_rows=2),
        _row_batch(((" left ", "  "), (" right ", "nan"), (1, 2))),
    )

    assert result[0].column_identities == (
        ColumnIdentity(0, "left__right"),
        ColumnIdentity(1, "col_1"),
    )


def test_header_only_sheet_emits_schema_without_inventing_data_row() -> None:
    result = _run(
        _transform(),
        _plan(header_rows=1),
        _row_batch((("header",),)),
    )

    assert len(result) == 1
    assert result[0].batch.num_rows == 0
    assert result[0].column_identities == (ColumnIdentity(0, "header"),)


def test_footer_removal_can_make_sheet_too_short_for_header_count() -> None:
    result = _run(
        _transform(),
        _plan(header_rows=2, skip_footer=1),
        _row_batch((("retained",), ("footer",))),
    )

    assert result[0].column_identities == (ColumnIdentity(0, "col_0"),)
    assert _result_contract(result)[2] == (("retained",),)


def test_empty_template_does_not_retain_first_batch_buffers() -> None:
    values = ["x" * 1_024 for _ in range(2_000)]
    raw = CoordinateBatch(
        batch=pa.record_batch([values], names=["0"]),
        row_numbers=pa.array(range(1, 2_001), type=pa.int64()),
        column_numbers=(1,),
    )
    operation = _transform().open(_plan(header_rows=0))

    operation.push(raw)

    assert operation._template is not None
    assert operation._template.batch.get_total_buffer_size() < 1_024


def test_footer_holdback_detaches_small_tail_from_large_batch_buffers() -> None:
    values = ["x" * 1_024 for _ in range(2_000)]
    raw = CoordinateBatch(
        batch=pa.record_batch([values], names=["0"]),
        row_numbers=pa.array(range(1, 2_001), type=pa.int64()),
        column_numbers=(1,),
    )
    operation = _transform().open(_plan(header_rows=0, skip_footer=1))

    operation.push(raw)

    retained = sum(batch.batch.get_total_buffer_size() for batch in operation._buffer)
    assert retained < 4_096


def test_materialized_integral_numeric_inference_precedes_header_framing() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch([[1.0, 2.0, 3.0]], names=["0"]),
        row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _transform().open(_plan(header_rows=1))._materialize_complete(raw)

    assert result[0].column_identities == (ColumnIdentity(0, "1"),)
    assert result[0].batch.column(0).type == pa.int64()
    assert _result_contract(result)[2] == ((2,), (3,))


@pytest.mark.parametrize("arrow_type", [pa.float64(), pa.int64()])
def test_materialized_nullable_numeric_inference_precedes_header_framing(
    arrow_type: pa.DataType,
) -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([1, None, 3], type=arrow_type)],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _transform().open(_plan(header_rows=1))._materialize_complete(raw)

    assert result[0].column_identities == (ColumnIdentity(0, "1.0"),)
    assert result[0].batch.column(0).type == pa.float64()


def test_materialized_fractional_footer_forces_float_before_footer_removal() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch([[1.0, 2.0, 2.5]], names=["0"]),
        row_numbers=pa.array([1, 2, 3], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _transform().open(_plan(header_rows=1, skip_footer=1))._materialize_complete(raw)

    assert result[0].column_identities == (ColumnIdentity(0, "1.0"),)
    assert result[0].batch.column(0).type == pa.float64()
    assert _result_contract(result)[2] == ((2.0,),)


def test_materialized_unsafe_integral_double_raises_compatibility_signal() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch([[float(2**53)]], names=["0"]),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1,),
    )

    with pytest.raises(CoordinateCompatibilityError):
        _transform().open(_plan())._materialize_complete(raw)


def test_materialized_nullable_large_integer_requires_legacy_fallback() -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([2**53 + 1, None], type=pa.int64())],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1,),
    )

    with pytest.raises(CoordinateCompatibilityError):
        _transform().open(_plan())._materialize_complete(raw)


def test_nullable_integer_exactness_is_checked_before_arrow_cast(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([2**53 + 1, None], type=pa.int64())],
            names=["0"],
        ),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1,),
    )

    def permissive_cast(
        array: pa.Array,
        target: pa.DataType,
        *,
        safe: bool,
    ) -> pa.Array:
        assert safe is True
        return pa.array(
            [None if value is None else float(value) for value in array.to_pylist()],
            type=target,
        )

    monkeypatch.setattr(coordinates_module.pc, "cast", permissive_cast)

    with pytest.raises(CoordinateCompatibilityError):
        _transform().open(_plan())._materialize_complete(raw)


def test_range_projection_indexes_original_arrays_without_filter_copy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def forbidden_filter(*_args: object, **_kwargs: object) -> None:
        pytest.fail("range projection must not filter before taking rows")

    monkeypatch.setattr(coordinates_module.pc, "filter", forbidden_filter)

    result = _run(
        _transform(),
        _plan(cell_range="A2:B3"),
        _grid_batch(rows=(1, 2, 3), columns=(1, 2)),
    )

    assert _result_contract(result) == (
        (2, 3),
        (1, 2),
        ((21, 22), (31, 32)),
    )


def test_many_disjoint_merges_rebuild_each_column_at_most_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    merged_ranges = tuple(MergeRange(row, 1, row + 1, 1) for row in range(1, 200, 2))
    transform = CoordinateTransform(
        IntervalIndex(()),
        IntervalIndex(()),
        merged_ranges,
    )
    raw = CoordinateBatch(
        batch=pa.record_batch([list(range(1, 201))], names=["0"]),
        row_numbers=pa.array(range(1, 201), type=pa.int64()),
        column_numbers=(1,),
    )
    calls = 0
    original_if_else = coordinates_module.pc.if_else

    def counted_if_else(*args: object, **kwargs: object) -> pa.Array:
        nonlocal calls
        calls += 1
        return original_if_else(*args, **kwargs)

    monkeypatch.setattr(coordinates_module.pc, "if_else", counted_if_else)

    result = _run(
        transform,
        _plan(merge_strategy=MergeStrategy.FILL),
        raw,
    )

    assert calls <= 1
    assert _result_contract(result)[2][:4] == ((1,), (1,), (3,), (3,))


def test_materialized_timestamp_normalizes_to_microseconds_before_framing() -> None:
    from datetime import datetime

    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([datetime(2024, 1, 1)], type=pa.timestamp("ms"))],
            names=["0"],
        ),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1,),
    )

    result = _transform().open(_plan())._materialize_complete(raw)

    assert result[0].batch.column(0).type == pa.timestamp("us")


@pytest.mark.parametrize("epoch", ["1899-12-31T12:00:00", "1904-01-01T12:00:00"])
def test_materialized_time_only_epoch_requires_legacy_fallback(epoch: str) -> None:
    from datetime import datetime

    raw = CoordinateBatch(
        batch=pa.record_batch(
            [pa.array([datetime.fromisoformat(epoch)], type=pa.timestamp("ms"))],
            names=["0"],
        ),
        row_numbers=pa.array([1], type=pa.int64()),
        column_numbers=(1,),
    )

    with pytest.raises(CoordinateCompatibilityError):
        _transform().open(_plan())._materialize_complete(raw)


def test_materialized_mode_rejects_operation_after_streaming_push() -> None:
    operation = _transform().open(_plan())
    operation.push(_grid_batch(rows=(1,), columns=(1,)))

    with pytest.raises(
        RuntimeError,
        match=r"^materialized coordinate operation must be fresh$",
    ):
        operation._materialize_complete(_grid_batch(rows=(2,), columns=(1,)))
