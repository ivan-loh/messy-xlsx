"""Missing value handling for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import numpy as np
import pandas as pd


# ============================================================================
# Config
# ============================================================================

DEFAULT_MISSING_VALUES = [
    "NA",
    "N/A",
    "n/a",
    "#N/A",
    "null",
    "NULL",
    "None",
    "NONE",
    "-",
    "--",
    "---",
    ".",
    "..",
    "...",
    "?",
    "??",
    "???",
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


# ============================================================================
# Core
# ============================================================================

class MissingValueHandler:
    """Standardize missing value representations."""

    def __init__(
        self,
        extra_values: list[str] | None = None,
        empty_string_as_na: bool = True,
    ):
        """Initialize handler."""
        self.missing_values      = DEFAULT_MISSING_VALUES.copy()
        if extra_values:
            self.missing_values.extend(extra_values)
        self.empty_string_as_na = empty_string_as_na

    def normalize(
        self,
        df: pd.DataFrame,
        drop_empty_rows: bool = True,
        drop_empty_cols: bool = True,
    ) -> pd.DataFrame:
        """Handle missing values in DataFrame."""
        df = df.copy()

        df = df.replace(self.missing_values, np.nan).infer_objects(copy=False)

        if self.empty_string_as_na:
            df = df.replace("", np.nan).infer_objects(copy=False)
            df = df.replace(r"^\s*$", np.nan, regex=True).infer_objects(copy=False)

        if drop_empty_rows:
            df = df.dropna(how="all")

        if drop_empty_cols:
            df = df.dropna(axis=1, how="all")

        return df.reset_index(drop=True)
