"""The single Arrow-to-pandas bridge for legacy materialized APIs."""

from __future__ import annotations

import pandas as pd
import pyarrow as pa

from messy_xlsx.parsing.parse_plan import ParsePlan


class LegacyDataFrameAdapter:
    """Convert one materialized Arrow table exactly once for legacy framing."""

    def to_dataframe(
        self,
        table: pa.Table,
        plan: ParsePlan | None,
    ) -> pd.DataFrame:
        """Return the raw DataFrame; existing owners retain all transforms."""
        del plan
        return table.to_pandas()
