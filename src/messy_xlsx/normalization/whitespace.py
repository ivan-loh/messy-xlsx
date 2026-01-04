"""Whitespace normalization for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import pandas as pd


# ============================================================================
# Core
# ============================================================================

class WhitespaceNormalizer:
    """Clean whitespace issues in text data."""

    def normalize(
        self,
        df: pd.DataFrame,
        preserve_linebreaks: bool = False,
    ) -> pd.DataFrame:
        """Normalize whitespace in all string columns."""
        df = df.copy()

        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = self._normalize_column(df[col], preserve_linebreaks)

        return df

    def _normalize_column(
        self,
        series: pd.Series,
        preserve_linebreaks: bool,
    ) -> pd.Series:
        """Normalize whitespace in a single column."""
        result = series.copy()

        mask = result.notna() & (result.apply(type) == str)

        if not mask.any():
            return result

        text = result[mask].astype(str)

        text = text.str.strip()

        text = text.str.replace("\xa0", " ", regex=False)
        text = text.str.replace("\u00a0", " ", regex=False)

        if not preserve_linebreaks:
            text = text.str.replace(r"[\r\n]+", " ", regex=True)

        text = text.str.replace(r"\s+", " ", regex=True)

        text = text.str.strip()

        result[mask] = text

        return result
