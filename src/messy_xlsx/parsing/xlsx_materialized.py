"""Whole-sheet fastexcel materialization into one Arrow table."""

from __future__ import annotations

from itertools import pairwise

import pyarrow as pa

from messy_xlsx.parsing.coordinates import (
    CoordinateBatch,
    CoordinateCompatibilityError,
    CoordinateTransform,
)
from messy_xlsx.parsing.fastexcel_session import FastexcelSession
from messy_xlsx.parsing.parse_plan import ParsePlan


def _coerce_materialized_table(materialized: object) -> pa.Table:
    """Normalize fastexcel's eager and wrapper return shapes to one table."""
    if isinstance(materialized, pa.Table):
        return materialized
    if isinstance(materialized, pa.RecordBatch):
        return pa.Table.from_batches([materialized])

    to_arrow = getattr(materialized, "to_arrow", None)
    if not callable(to_arrow):
        raise TypeError("fastexcel materialization must produce a pyarrow Table or RecordBatch")
    converted = to_arrow()
    if isinstance(converted, pa.Table):
        return converted
    if isinstance(converted, pa.RecordBatch):
        return pa.Table.from_batches([converted])
    raise TypeError("fastexcel to_arrow() must produce a pyarrow Table or RecordBatch")


class FastexcelMaterializedReader:
    """A thin, non-owning operation reader with a bound parse plan."""

    def __init__(
        self,
        session: FastexcelSession,
        sheet: str,
        plan: ParsePlan,
        transform: CoordinateTransform | None = None,
    ) -> None:
        self._session = session
        self._sheet = sheet
        self._plan = plan
        self._transform = transform

    def read_table(self) -> pa.Table:
        """Materialize the raw worksheet once without applying transforms."""
        if self._transform is None:
            materialized = self._session.materialize(self._sheet, skip_rows=0)
            return _coerce_materialized_table(materialized)

        materialized = self._session.materialize(
            self._sheet,
            skip_rows=0,
            dtype_coercion="strict",
            eager=False,
        )
        column_numbers = _selected_column_numbers(materialized)
        table = _coerce_materialized_table(materialized)
        operation = self._transform.open(self._plan)
        projection = operation._projection
        raw = _coordinate_envelopes(
            table,
            column_numbers,
            self._transform,
            self._plan,
            range_max_row=projection.max_row if projection is not None else None,
        )
        output = operation._materialize_complete(raw)
        if not output:
            return pa.table({})
        identities = output[0].column_identities
        if any(batch.column_identities != identities for batch in output):
            raise CoordinateCompatibilityError(
                "coordinate output identities changed during materialization"
            )
        result = pa.Table.from_batches([batch.batch for batch in output])
        return result.rename_columns([str(identity.display_name) for identity in identities])


def _selected_column_numbers(materialized: object) -> tuple[int, ...]:
    selected = getattr(materialized, "selected_columns", None)
    if selected is None:
        raise CoordinateCompatibilityError("fastexcel selected-column origins are unavailable")
    try:
        absolute_indices = tuple(column.absolute_index for column in selected)
    except (AttributeError, TypeError) as error:
        raise CoordinateCompatibilityError(
            "fastexcel selected-column origins are invalid"
        ) from error
    if any(type(index) is not int or not 0 <= index < 16_384 for index in absolute_indices):
        raise CoordinateCompatibilityError("fastexcel selected-column origins are invalid")
    numbers = tuple(index + 1 for index in absolute_indices)
    if any(left >= right for left, right in pairwise(numbers)):
        raise CoordinateCompatibilityError("fastexcel selected-column origins are invalid")
    return numbers


def _coordinate_envelopes(
    table: pa.Table,
    column_numbers: tuple[int, ...],
    transform: CoordinateTransform,
    plan: ParsePlan,
    *,
    range_max_row: int | None = None,
) -> tuple[CoordinateBatch, ...]:
    if len(column_numbers) != table.num_columns:
        raise CoordinateCompatibilityError(
            "fastexcel selected-column origins do not match Arrow columns"
        )
    if plan.cell_range:
        envelope_columns = column_numbers
    else:
        max_column = max(
            transform._observed_max_col,
            column_numbers[-1] if column_numbers else 0,
        )
        envelope_columns = tuple(range(1, max_column + 1))
    row_count = (
        min(table.num_rows, range_max_row)
        if range_max_row is not None
        else max(table.num_rows, transform._observed_max_row)
    )
    if not envelope_columns and row_count:
        raise CoordinateCompatibilityError(
            "Arrow cannot represent worksheet rows with zero visible columns"
        )
    type_by_column = {
        number: table.schema.field(index).type for index, number in enumerate(column_numbers)
    }
    batches: list[CoordinateBatch] = []
    row_offset = 0
    for raw_batch in table.to_batches():
        if row_offset >= row_count:
            break
        remaining = row_count - row_offset
        if raw_batch.num_rows > remaining:
            raw_batch = raw_batch.slice(0, remaining)
        arrays_by_column = {
            number: raw_batch.column(index) for index, number in enumerate(column_numbers)
        }
        batches.append(
            _envelope_batch(
                arrays_by_column,
                envelope_columns,
                row_offset + 1,
                raw_batch.num_rows,
            )
        )
        row_offset += raw_batch.num_rows
    if row_offset < row_count:
        batches.append(
            _envelope_batch(
                {
                    number: pa.nulls(row_count - row_offset, type=column_type)
                    for number, column_type in type_by_column.items()
                },
                envelope_columns,
                row_offset + 1,
                row_count - row_offset,
            )
        )
    if not batches:
        batches.append(
            _envelope_batch(
                {
                    number: pa.array([], type=column_type)
                    for number, column_type in type_by_column.items()
                },
                envelope_columns,
                1,
                0,
            )
        )
    return tuple(batches)


def _envelope_batch(
    arrays_by_column: dict[int, pa.Array],
    envelope_columns: tuple[int, ...],
    start_row: int,
    row_count: int,
) -> CoordinateBatch:
    arrays = [arrays_by_column.get(number, pa.nulls(row_count)) for number in envelope_columns]
    return CoordinateBatch(
        batch=pa.record_batch(
            arrays,
            names=[str(index) for index in range(len(arrays))],
        ),
        row_numbers=pa.array(
            range(start_row, start_row + row_count),
            type=pa.int64(),
        ),
        column_numbers=envelope_columns,
    )
