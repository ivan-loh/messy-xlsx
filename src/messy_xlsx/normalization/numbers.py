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

        for col in df.select_dtypes(include=["object"]).columns:
            if col in semantic_hints:
                hint = semantic_hints[col].upper()
                if any(t in hint for t in ["VARCHAR", "TEXT", "STRING", "CHAR"]):
                    continue

            if self._looks_like_numbers(df[col]):
                df[col] = self._normalize_column(df[col])

        return df

    def _detect_locale(self, df: pd.DataFrame) -> tuple[str, str]:
        """Detect number locale from DataFrame."""
        samples = []

        for col in df.select_dtypes(include=["object"]).columns:
            sample = df[col].dropna().head(50).astype(str)
            for val in sample:
                if NUMERIC_CHARS_PATTERN.match(val):
                    samples.append(val)

        if not samples:
            return ".", ","

        comma_decimal = sum(1 for s in samples if COMMA_DECIMAL_PATTERN.search(s))
        dot_decimal = sum(1 for s in samples if DOT_DECIMAL_PATTERN.search(s))
        dot_thousands = sum(1 for s in samples if DOT_THOUSANDS_PATTERN.search(s))
        comma_thousands = sum(1 for s in samples if COMMA_THOUSANDS_PATTERN.search(s))

        if comma_decimal > dot_decimal and dot_thousands > comma_thousands:
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
        # Work with string representation
        str_series = series.astype(str)

        # Vectorized: remove currency symbols
        str_series = str_series.str.replace(_currency_pattern, "", regex=True)
        str_series = str_series.str.strip()

        # Handle accounting format (xxx) -> -xxx
        is_accounting = str_series.str.match(r"^\s*\([^)]+\)\s*$", na=False)
        if is_accounting.any():
            # Extract content from parentheses and add negative sign
            str_series = str_series.where(
                ~is_accounting,
                "-" + str_series.str.replace(r"[()]", "", regex=True).str.strip()
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
