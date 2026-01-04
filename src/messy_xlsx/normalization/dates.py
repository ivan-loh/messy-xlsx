"""Date normalization for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import numpy as np
import pandas as pd


# ============================================================================
# Config
# ============================================================================

EXCEL_EPOCH     = "1899-12-30"
EXCEL_DATE_MIN  = 1
EXCEL_DATE_MAX  = 60000


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
            if col in semantic_hints:
                hint = semantic_hints[col].upper()
                if any(t in hint for t in ["DECIMAL", "NUMERIC", "INTEGER", "FLOAT", "VARCHAR", "TEXT"]):
                    continue

            if self._looks_like_excel_dates(df[col]):
                df[col] = self._convert_excel_dates(df[col])
            elif self._looks_like_text_dates(df[col]):
                df[col] = self._convert_text_dates(df[col])

        return df

    def _looks_like_excel_dates(self, series: pd.Series) -> bool:
        """Check if column contains Excel serial dates."""
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
            parsed = pd.to_datetime(sample, errors="coerce")
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
            return pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
        except Exception:
            return series
