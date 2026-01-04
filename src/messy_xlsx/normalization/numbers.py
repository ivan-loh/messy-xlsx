"""Locale-aware number normalization."""

# ============================================================================
# Imports
# ============================================================================

import re

import numpy as np
import pandas as pd


# ============================================================================
# Config
# ============================================================================

CURRENCY_SYMBOLS = ["$", "€", "£", "¥", "₹", "CHF", "kr", "zł"]

ACCOUNTING_PATTERN = re.compile(r"^\s*\(([^)]+)\)\s*$")

NUMBER_PATTERN = re.compile(r"^[+-]?[\d,.\s]+$|^\([0-9,.\s]+\)$|^[$€£¥₹][0-9,.\s]+$")


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
        self.decimal_separator   = decimal_separator
        self.thousands_separator = thousands_separator

    def normalize(
        self,
        df: pd.DataFrame,
        semantic_hints: dict[str, str] | None = None,
    ) -> pd.DataFrame:
        """Normalize numbers in DataFrame."""
        df             = df.copy()
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
                if re.match(r"[\d.,\s]+", val):
                    samples.append(val)

        if not samples:
            return ".", ","

        comma_decimal   = sum(1 for s in samples if re.search(r"\d,\d{2}$", s))
        dot_decimal     = sum(1 for s in samples if re.search(r"\d\.\d{2}$", s))
        dot_thousands   = sum(1 for s in samples if re.search(r"\d\.\d{3}", s))
        comma_thousands = sum(1 for s in samples if re.search(r"\d,\d{3}", s))

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
        """Normalize numbers in a column."""
        result = series.copy()

        def normalize_value(val):
            if pd.isna(val):
                return np.nan

            val_str = str(val).strip()

            if not val_str:
                return np.nan

            for symbol in CURRENCY_SYMBOLS:
                val_str = val_str.replace(symbol, "")

            val_str = val_str.strip()

            match = ACCOUNTING_PATTERN.match(val_str)
            if match:
                val_str = "-" + match.group(1)

            if self.thousands_separator:
                val_str = val_str.replace(self.thousands_separator, "")

            if self.decimal_separator and self.decimal_separator != ".":
                val_str = val_str.replace(self.decimal_separator, ".")

            val_str = val_str.replace(" ", "")

            try:
                return float(val_str)
            except ValueError:
                return val

        return result.apply(normalize_value)
