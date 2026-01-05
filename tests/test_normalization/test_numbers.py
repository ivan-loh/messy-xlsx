"""Unit tests for NumberNormalizer."""

import pandas as pd
import pytest

from messy_xlsx.normalization import NumberNormalizer


class TestNumberNormalizer:
    """Test number normalization functionality."""

    def test_normalize_us_format(self):
        """Test normalizing US number format."""
        normalizer = NumberNormalizer(decimal_separator=".", thousands_separator=",")

        df = pd.DataFrame({
            "amount": ["1,234.56", "2,345.67", "3,456.78"]
        })

        result = normalizer.normalize(df)

        assert pd.api.types.is_numeric_dtype(result["amount"])
        assert result["amount"].iloc[0] == pytest.approx(1234.56)

    def test_normalize_european_format(self):
        """Test normalizing European number format."""
        normalizer = NumberNormalizer(decimal_separator=",", thousands_separator=".")

        df = pd.DataFrame({
            "amount": ["1.234,56", "2.345,67", "3.456,78"]
        })

        result = normalizer.normalize(df)

        assert pd.api.types.is_numeric_dtype(result["amount"])
        assert result["amount"].iloc[0] == pytest.approx(1234.56)

    def test_remove_currency_symbols(self):
        """Test removing currency symbols."""
        normalizer = NumberNormalizer(decimal_separator=".", thousands_separator=",")

        df = pd.DataFrame({
            "amount": ["$1,234.56", "€2,345.67", "£3,456.78"]
        })

        result = normalizer.normalize(df)

        assert pd.api.types.is_numeric_dtype(result["amount"])
        assert result["amount"].iloc[0] == pytest.approx(1234.56)

    def test_accounting_format(self):
        """Test handling accounting format (negative in parentheses)."""
        normalizer = NumberNormalizer(decimal_separator=".", thousands_separator=",")

        df = pd.DataFrame({
            "amount": ["1,234.56", "(234.56)", "567.89"]
        })

        result = normalizer.normalize(df)

        assert result["amount"].iloc[1] == pytest.approx(-234.56)

    def test_auto_locale_detection(self):
        """Test automatic locale detection."""
        normalizer = NumberNormalizer()  # Auto-detect when None

        df = pd.DataFrame({
            "us_format": ["1,234.56", "2,345.67"],
            "eu_format": ["1.234,56", "2.345,67"]
        })

        result = normalizer.normalize(df)

        # Should detect and normalize correctly
        assert pd.api.types.is_numeric_dtype(result["us_format"])

    def test_preserve_integers(self):
        """Test that integer values are preserved."""
        normalizer = NumberNormalizer(decimal_separator=".", thousands_separator=",")

        df = pd.DataFrame({
            "count": [100, 200, 300]
        })

        result = normalizer.normalize(df)

        assert pd.api.types.is_numeric_dtype(result["count"])
        assert result["count"].iloc[0] == 100
