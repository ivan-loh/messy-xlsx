"""Whitespace normalization for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import re

import pandas as pd

# ============================================================================
# Compiled patterns (avoid recompiling on each call)
# ============================================================================

# Matches horizontal whitespace (space, tab, NBSP) - preserves newlines
_HORIZONTAL_WS_PATTERN = re.compile(r"[ \t\xa0]+")
# Matches all whitespace including NBSP and newlines
_ALL_WS_PATTERN = re.compile(r"[\s\xa0]+")


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

        non_null_mask = result.notna()
        if not non_null_mask.any():
            return result

        # Fast path: sample up to 100 non-null values to check if all are strings.
        # Object columns are usually all-string; avoid per-element isinstance scan.
        non_null_vals = result[non_null_mask]
        sample = non_null_vals.iloc[:100]
        all_strings = all(isinstance(v, str) for v in sample)

        if all_strings:
            mask = non_null_mask
        else:
            is_str = result.apply(lambda x: isinstance(x, str))
            mask = non_null_mask & is_str.astype(bool)

        if not mask.any():
            return result

        text = result[mask].astype(str)

        # Single-pass replacement using pre-compiled pattern
        # Replaces NBSP and collapses whitespace sequences to single space
        if preserve_linebreaks:
            # Only collapse horizontal whitespace (space, tab, NBSP), preserve newlines
            text = text.str.replace(_HORIZONTAL_WS_PATTERN, " ", regex=True)
        else:
            # Collapse all whitespace including newlines and NBSP
            text = text.str.replace(_ALL_WS_PATTERN, " ", regex=True)

        text = text.str.strip()

        result[mask] = text

        return result
