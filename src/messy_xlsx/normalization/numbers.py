"""Locale-aware number normalization."""

# ============================================================================
# Imports
# ============================================================================

import re

import numpy as np
import pandas as pd

# ============================================================================
# Config - Compiled patterns at module level for performance
# ============================================================================

CURRENCY_SYMBOLS = ["$", "€", "£", "¥", "₹", "CHF", "kr", "zł"]

# Pre-compile patterns
ACCOUNTING_PATTERN = re.compile(r"^\s*\(([^)]+)\)\s*$")
NUMBER_PATTERN = re.compile(r"^[+-]?[\d,.\s]+$|^\([0-9,.\s]+\)$|^[$€£¥₹][0-9,.\s]+$")
COMMA_DECIMAL_PATTERN = re.compile(r"\d,\d{2}$")
DOT_DECIMAL_PATTERN = re.compile(r"\d\.\d{2}$")
DOT_THOUSANDS_PATTERN = re.compile(r"\d\.\d{3}")
COMMA_THOUSANDS_PATTERN = re.compile(r"\d,\d{3}")
NUMERIC_CHARS_PATTERN = re.compile(r"[\d.,\s]+")

# Pre-build currency removal pattern
_currency_pattern = re.compile("|".join(re.escape(s) for s in CURRENCY_SYMBOLS))


# ============================================================================
# Core
# ============================================================================


class NumberNormalizer:
    """Normalize numbers with locale-aware parsing."""

    def __init__(
        self,
        decimal_separator: str | None = None,
        thousands_separator: str | None = None,
    ):
        """Initialize normalizer."""
        self.decimal_separator = decimal_separator
        self.thousands_separator = thousands_separator

    def normalize(
        self,
        df: pd.DataFrame,
        semantic_hints: dict[str, str] | None = None,
    ) -> pd.DataFrame:
        """Normalize numbers in DataFrame."""
        df = df.copy()
        semantic_hints = semantic_hints or {}

        if self.decimal_separator is None:
            self.decimal_separator, self.thousands_separator = self._detect_locale(df)

        for idx in self._string_positions(df):
            col = df.columns[idx]
            series = df.iloc[:, idx]
            if col in semantic_hints:
                hint = semantic_hints[col].upper()
                if any(t in hint for t in ["VARCHAR", "TEXT", "STRING", "CHAR"]):
                    continue

            if self._looks_like_numbers(series):
                df.isetitem(idx, self._normalize_column(series))

        return df

    def _detect_locale(self, df: pd.DataFrame) -> tuple[str, str]:
        """Detect number locale from DataFrame."""
        samples = []

        for idx in self._string_positions(df):
            sample = df.iloc[:, idx].dropna().head(50).astype(str)
            for val in sample:
                if NUMERIC_CHARS_PATTERN.match(val):
                    samples.append(val)

        if not samples:
            return ".", ","

        comma_decimal = sum(1 for s in samples if COMMA_DECIMAL_PATTERN.search(s))
        dot_decimal = sum(1 for s in samples if DOT_DECIMAL_PATTERN.search(s))
        dot_thousands = sum(1 for s in samples if DOT_THOUSANDS_PATTERN.search(s))
        comma_thousands = sum(1 for s in samples if COMMA_THOUSANDS_PATTERN.search(s))

        # Strong evidence: both decimal and thousands patterns agree
        if comma_decimal > dot_decimal and dot_thousands >= comma_thousands:
            return ",", "."
        if dot_decimal > comma_decimal and comma_thousands >= dot_thousands:
            return ".", ","

        # Weaker evidence: only decimal pattern, no contradicting thousands pattern
        if comma_decimal > dot_decimal and comma_thousands == 0:
            return ",", "."

        return ".", ","

    def _looks_like_numbers(self, series: pd.Series) -> bool:
        """Check if column looks numeric."""
        sample = series.dropna().head(50).astype(str)

        if len(sample) == 0:
            return False

        matches = sum(1 for val in sample if NUMBER_PATTERN.match(val.strip()))
        return matches > len(sample) * 0.5

    def _normalize_column(self, series: pd.Series) -> pd.Series:
        """
        Normalize numbers in a column using vectorized operations.

        Converts in single pass - if any value fails, returns original series.
        """
        # Preserve values that are already numeric (int/float) in object dtype columns.
        # Stringifying them and re-parsing through locale rules corrupts values
        # (e.g., 850000.0 → "850000.0" → strip "." as thousands sep → 8500000).
        if series.dtype == object:
            already_numeric = series.apply(
                lambda x: isinstance(x, (int, float)) and not isinstance(x, bool)
            )
            if already_numeric.any():
                str_mask = ~already_numeric & series.notna()
                if str_mask.any():
                    str_part = self._normalize_str_part(series[str_mask])
                    # If string values couldn't be converted to numeric,
                    # the column is truly mixed — return original unchanged
                    if not pd.api.types.is_numeric_dtype(str_part):
                        return series
                    result = pd.to_numeric(series.where(already_numeric), errors="coerce")
                    result = result.copy()
                    result[str_mask] = str_part
                    return result
                # All non-null values are already numeric
                return pd.to_numeric(series, errors="coerce")

        # Work with string representation
        return self._normalize_str_part(series)

    def _normalize_str_part(self, series: pd.Series) -> pd.Series:
        """Normalize string values, detecting per-column mixed locale."""
        str_vals = series.dropna().astype(str)
        has_comma_dec = str_vals.str.contains(COMMA_DECIMAL_PATTERN, regex=True).any()
        has_dot_dec = str_vals.str.contains(DOT_DECIMAL_PATTERN, regex=True).any()

        if has_comma_dec and has_dot_dec:
            # Mixed separators in the same column — normalize per-value
            return self._normalize_mixed_locale_series(series)

        return self._normalize_string_series(series)

    def _normalize_mixed_locale_series(self, series: pd.Series) -> pd.Series:
        """Normalize values when the same column has mixed decimal separators.

        For each value, determine the decimal separator by looking at
        which separator appears last and is followed by exactly 1-2 digits.
        """

        def _per_value(val: object) -> object:
            if pd.isna(val):
                return np.nan
            s = str(val).strip()
            s = _currency_pattern.sub("", s).strip()
            # Accounting format
            m = ACCOUNTING_PATTERN.match(s)
            if m:
                s = "-" + m.group(1).strip()
            s = s.replace(" ", "")

            has_comma_2 = bool(COMMA_DECIMAL_PATTERN.search(s))
            has_dot_2 = bool(DOT_DECIMAL_PATTERN.search(s))

            if has_comma_2 and has_dot_2:
                # Both present — last separator is the decimal
                if s.rfind(",") > s.rfind("."):
                    s = s.replace(".", "").replace(",", ".")
                else:
                    s = s.replace(",", "")
            elif has_comma_2:
                s = s.replace(".", "").replace(",", ".")
            elif has_dot_2:
                s = s.replace(",", "")
            else:
                s = s.replace(",", "")

            if not s or s in ("-", "+"):
                return np.nan
            try:
                return float(s)
            except ValueError:
                return np.nan

        result = series.apply(_per_value)
        result = pd.to_numeric(result, errors="coerce")

        original_nulls = series.isna()
        new_nulls = result.isna() & ~original_nulls
        if new_nulls.any():
            return series
        return result

    def _normalize_string_series(self, series: pd.Series) -> pd.Series:
        """Normalize a series of string values to numeric."""
        str_series = series.astype(str)

        # Vectorized: remove currency symbols
        str_series = str_series.str.replace(_currency_pattern, "", regex=True)
        str_series = str_series.str.strip()

        # Handle accounting format (xxx) -> -xxx
        is_accounting = str_series.str.match(r"^\s*\([^)]+\)\s*$", na=False)
        if is_accounting.any():
            # Extract content from parentheses and add negative sign
            str_series = str_series.where(
                ~is_accounting, "-" + str_series.str.replace(r"[()]", "", regex=True).str.strip()
            )

        # Vectorized: remove thousands separator
        if self.thousands_separator:
            str_series = str_series.str.replace(self.thousands_separator, "", regex=False)

        # Vectorized: convert decimal separator to dot
        if self.decimal_separator and self.decimal_separator != ".":
            str_series = str_series.str.replace(self.decimal_separator, ".", regex=False)

        # Vectorized: remove spaces
        str_series = str_series.str.replace(" ", "", regex=False)

        # Handle empty strings and original NaN
        str_series = str_series.replace("", np.nan)
        str_series = str_series.replace("nan", np.nan)

        # Try to convert to numeric - if fails, return original
        result = pd.to_numeric(str_series, errors="coerce")

        # Check if conversion created new NaNs (excluding original NaNs)
        original_nulls = series.isna()
        new_nulls = result.isna() & ~original_nulls

        # If we have values that couldn't convert (and weren't originally null),
        # return the original series to avoid mixed types
        if new_nulls.any():
            return series

        return result

    def _string_positions(self, df: pd.DataFrame) -> list[int]:
        """Return positional indices for string-like columns."""
        return [
            idx
            for idx, dtype in enumerate(df.dtypes)
            if pd.api.types.is_object_dtype(dtype) or isinstance(dtype, pd.StringDtype)
        ]
