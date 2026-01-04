"""Date normalization for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import re

import numpy as np
import pandas as pd


# ============================================================================
# Config
# ============================================================================

EXCEL_EPOCH     = "1899-12-30"
EXCEL_DATE_MIN  = 1
EXCEL_DATE_MAX  = 60000

# Patterns that indicate a column likely contains dates
DATE_COLUMN_PATTERNS = [
    r"(?i)date",
    r"(?i)time",
    r"(?i)timestamp",
    r"(?i)created",
    r"(?i)updated",
    r"(?i)modified",
    r"(?i)born",
    r"(?i)expired?",
    r"(?i)due",
    r"(?i)start",
    r"(?i)end",
    r"(?i)period",
    r"(?i)day",
    r"(?i)month",
    r"(?i)year",
]

# Patterns that indicate a column should NOT be converted to dates
NON_DATE_COLUMN_PATTERNS = [
    r"(?i)count",
    r"(?i)total",
    r"(?i)sum",
    r"(?i)qty",
    r"(?i)quantity",
    r"(?i)amount",
    r"(?i)number",
    r"(?i)num",
    r"(?i)id$",
    r"(?i)_id$",
    r"(?i)unique",
    r"(?i)transactions?",
    r"(?i)customers?",
    r"(?i)users?",
    r"(?i)orders?",
    r"(?i)items?",
    r"(?i)units?",
    r"(?i)price",
    r"(?i)cost",
    r"(?i)revenue",
    r"(?i)sales",
    r"(?i)score",
    r"(?i)rating",
    r"(?i)rank",
    r"(?i)index",
    r"(?i)age",
    r"(?i)size",
    r"(?i)length",
    r"(?i)width",
    r"(?i)height",
    r"(?i)weight",
    r"(?i)percent",
    r"(?i)rate",
    r"(?i)ratio",
]


# ============================================================================
# Core
# ============================================================================

class DateNormalizer:
    """Normalize dates with multiple format support."""

    def normalize(
        self,
        df: pd.DataFrame,
        semantic_hints: dict[str, str] | None = None,
    ) -> pd.DataFrame:
        """Normalize dates in DataFrame."""
        df             = df.copy()
        semantic_hints = semantic_hints or {}

        for col in df.columns:
            # Skip if semantic hint says not a date
            if col in semantic_hints:
                hint = semantic_hints[col].upper()
                if any(t in hint for t in ["DECIMAL", "NUMERIC", "INTEGER", "FLOAT", "VARCHAR", "TEXT"]):
                    continue
                # Explicitly marked as timestamp - always convert
                if "TIMESTAMP" in hint or "DATE" in hint:
                    if self._is_numeric_date_candidate(df[col]):
                        df[col] = self._convert_excel_dates(df[col])
                    elif self._looks_like_text_dates(df[col]):
                        df[col] = self._convert_text_dates(df[col])
                    continue

            # For numeric columns, only convert if column name suggests it's a date
            if self._is_numeric_date_candidate(df[col]):
                if self._column_name_suggests_date(str(col)):
                    df[col] = self._convert_excel_dates(df[col])
            # For text columns, be more permissive
            elif self._looks_like_text_dates(df[col]):
                df[col] = self._convert_text_dates(df[col])

        return df

    def _column_name_suggests_date(self, col_name: str) -> bool:
        """Check if column name suggests it contains dates."""
        # First check if name suggests NON-date (more specific patterns)
        for pattern in NON_DATE_COLUMN_PATTERNS:
            if re.search(pattern, col_name):
                return False

        # Then check if name suggests date
        for pattern in DATE_COLUMN_PATTERNS:
            if re.search(pattern, col_name):
                return True

        return False

    def _is_numeric_date_candidate(self, series: pd.Series) -> bool:
        """Check if column could be Excel serial dates (numeric check only)."""
        if not pd.api.types.is_numeric_dtype(series):
            return False

        sample = series.dropna()
        if len(sample) == 0:
            return False

        in_range   = (sample >= EXCEL_DATE_MIN) & (sample <= EXCEL_DATE_MAX)
        is_integer = (sample % 1 == 0)

        return (in_range & is_integer).mean() > 0.8

    def _looks_like_text_dates(self, series: pd.Series) -> bool:
        """Check if column contains text dates."""
        if series.dtype != object:
            return False

        sample = series.dropna().head(50).astype(str)
        if len(sample) == 0:
            return False

        try:
            # Use format="mixed" to avoid UserWarning about format inference
            parsed = pd.to_datetime(sample, errors="coerce", format="mixed")
            return parsed.notna().sum() > len(sample) * 0.5
        except Exception:
            return False

    def _convert_excel_dates(self, series: pd.Series) -> pd.Series:
        """Convert Excel serial dates to datetime."""
        try:
            return pd.to_datetime(
                series,
                unit   = "D",
                origin = EXCEL_EPOCH,
                errors = "coerce",
            )
        except Exception:
            return series

    def _convert_text_dates(self, series: pd.Series) -> pd.Series:
        """Convert text dates to datetime."""
        try:
            # Use format="mixed" instead of deprecated infer_datetime_format=True
            return pd.to_datetime(series, errors="coerce", format="mixed")
        except Exception:
            return series
