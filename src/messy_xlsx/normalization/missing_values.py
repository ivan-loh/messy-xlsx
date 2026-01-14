"""Missing value handling for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import numpy as np
import pandas as pd


# ============================================================================
# Config
# ============================================================================

# Conservative list - clearly "missing" indicators
DEFAULT_MISSING_VALUES = [
    "NA",
    "N/A",
    "n/a",
    "#N/A",
    "null",
    "NULL",
    "None",
    "NONE",
    "nan",
    "NaN",
    "NAN",
    "<NA>",
    "#NA",
    "missing",
    "MISSING",
    "nil",
    "NIL",
]

# Aggressive list - includes ambiguous values (opt-in)
EXTENDED_MISSING_VALUES = [
    "-",
    "--",
    "---",
    ".",
    "..",
    "...",
    "?",
    "??",
    "???",
]


# ============================================================================
# Core
# ============================================================================

class MissingValueHandler:
    """
    Standardize missing value representations.

    By default, uses type-preserving replacement:
    - String columns: missing values -> None (not np.nan)
    - Numeric columns: missing values -> np.nan

    This prevents mixed-type columns that break downstream tools
    like PyArrow, BigQuery, and Parquet.
    """

    def __init__(
        self,
        extra_values: list[str] | None = None,
        empty_string_as_na: bool = True,
        use_extended_list: bool = False,
        preserve_types: bool = True,
    ):
        """
        Initialize handler.

        Args:
            extra_values: Additional strings to treat as missing
            empty_string_as_na: Treat empty strings as missing
            use_extended_list: Include ambiguous values like "-", ".", "?"
            preserve_types: Use type-appropriate null values (recommended)
        """
        self.missing_values = DEFAULT_MISSING_VALUES.copy()

        if use_extended_list:
            self.missing_values.extend(EXTENDED_MISSING_VALUES)

        if extra_values:
            self.missing_values.extend(extra_values)

        self.empty_string_as_na = empty_string_as_na
        self.preserve_types = preserve_types

    def normalize(
        self,
        df: pd.DataFrame,
        drop_empty_rows: bool = True,
        drop_empty_cols: bool = True,
    ) -> pd.DataFrame:
        """
        Handle missing values in DataFrame.

        Args:
            df: Input DataFrame
            drop_empty_rows: Remove rows that are entirely empty
            drop_empty_cols: Remove columns that are entirely empty

        Returns:
            DataFrame with standardized missing values
        """
        df = df.copy()

        if self.preserve_types:
            df = self._type_aware_replace(df)
        else:
            # Legacy behavior - replace all with np.nan
            df = df.replace(self.missing_values, np.nan)
            if self.empty_string_as_na:
                df = df.replace("", np.nan)
                df = df.replace(r"^\s*$", np.nan, regex=True)

        if drop_empty_rows:
            df = df.dropna(how="all")

        if drop_empty_cols:
            df = df.dropna(axis=1, how="all")

        return df.reset_index(drop=True)

    def _type_aware_replace(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Replace missing values using type-appropriate null values.

        - String/object columns: use None (preserves object dtype)
        - Numeric columns: use np.nan (standard for floats)
        - Other types: use pd.NA (pandas nullable)
        """
        # Pre-compute missing values set for O(1) lookups
        missing_set = set(self.missing_values)

        for col in df.columns:
            col_dtype = df[col].dtype

            # Determine appropriate null value
            if col_dtype == "object":
                null_value = None
            elif np.issubdtype(col_dtype, np.floating):
                null_value = np.nan
            elif np.issubdtype(col_dtype, np.integer):
                null_value = np.nan
            else:
                null_value = pd.NA

            # Vectorized replacement using isin() - single pass instead of N passes
            mask = df[col].isin(missing_set)
            if mask.any():
                df.loc[mask, col] = null_value

            # Handle empty strings
            if self.empty_string_as_na and col_dtype == "object":
                # Empty string and whitespace-only in single mask operation
                str_series = df[col].astype(str)
                empty_mask = (df[col] == "") | str_series.str.fullmatch(r"\s*", na=False)
                if empty_mask.any():
                    df.loc[empty_mask, col] = null_value

        return df
