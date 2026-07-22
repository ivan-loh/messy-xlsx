"""Whole-sheet fastexcel materialization into one Arrow table."""

from __future__ import annotations

import pyarrow as pa

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
    ) -> None:
        self._session = session
        self._sheet = sheet
        self._plan = plan

    def read_table(self) -> pa.Table:
        """Materialize the raw worksheet once without applying transforms."""
        materialized = self._session.materialize(self._sheet, skip_rows=0)
        return _coerce_materialized_table(materialized)
