"""Type coercion for BigQuery/Arrow compatibility."""

# ============================================================================
# Imports
# ============================================================================

import numpy as np
import pandas as pd


# ============================================================================
# Core
# ============================================================================

class TypeCoercionNormalizer:
    """
    Ensure columns have consistent types for BigQuery/Arrow compatibility.

    Mixed-type object columns (containing both strings and numbers) are
    converted to strings to prevent Arrow conversion failures.
    """

    def normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Coerce mixed-type columns to consistent types.

        - Object columns with mixed types → all strings
        - Preserves None/NaN as null values
        """
        df = df.copy()

        for col in df.columns:
            if df[col].dtype == object:
                df[col] = self._coerce_object_column(df[col])

        return df

    def _coerce_object_column(self, series: pd.Series) -> pd.Series:
        """
        Coerce object column to consistent type.

        If column has mixed types (e.g., strings and ints), convert all to string.
        """
        non_null = series.dropna()

        if len(non_null) == 0:
            return series

        # Check what types are present (excluding NoneType)
        types = set()
        for val in non_null:
            if val is not None:
                # Group numeric types together
                if isinstance(val, (int, np.integer)):
                    types.add("int")
                elif isinstance(val, (float, np.floating)):
                    if not np.isnan(val):
                        types.add("float")
                elif isinstance(val, str):
                    types.add("str")
                else:
                    types.add(type(val).__name__)

        # If only one type (or empty), no coercion needed
        if len(types) <= 1:
            return series

        # Mixed types detected - convert all to string
        # This ensures BigQuery/Arrow compatibility
        def to_string(val):
            if val is None:
                return None
            if isinstance(val, float) and np.isnan(val):
                return None
            return str(val)

        return series.apply(to_string)
