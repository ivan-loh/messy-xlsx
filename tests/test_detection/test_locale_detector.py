"""Unit tests for LocaleDetector."""

import pandas as pd
import pytest

from messy_xlsx.detection import LocaleDetector


class TestLocaleDetector:
    """Test locale detection functionality."""

    def test_detect_us_locale(self):
        """Test detecting US number format (1,234.56)."""
        detector = LocaleDetector()

        df = pd.DataFrame({
            "amount": ["1,234.56", "2,345.67", "3,456.78"]
        })

        locale_info = detector.detect(df)

        assert locale_info["decimal_separator"] == "."
        assert locale_info["thousands_separator"] == ","

    def test_detect_european_locale(self):
        """Test detecting European number format (1.234,56)."""
        detector = LocaleDetector()

        df = pd.DataFrame({
            "amount": ["1.234,56", "2.345,67", "3.456,78"]
        })

        locale_info = detector.detect(df)

        assert locale_info["decimal_separator"] == ","
        assert locale_info["thousands_separator"] == "."

    def test_mixed_formats(self):
        """Test handling mixed number formats."""
        detector = LocaleDetector()

        df = pd.DataFrame({
            "amount": ["1,234.56", "1.234,56", "1000"]
        })

        locale_info = detector.detect(df)

        # Should return most common format
        assert locale_info["decimal_separator"] in [".", ","]

    def test_no_formatted_numbers(self):
        """Test handling data with no formatted numbers."""
        detector = LocaleDetector()

        df = pd.DataFrame({
            "amount": ["1000", "2000", "3000"]
        })

        locale_info = detector.detect(df)

        # Should have default values
        assert "decimal_separator" in locale_info
        assert "thousands_separator" in locale_info

    def test_empty_dataframe(self):
        """Test handling empty DataFrame."""
        detector = LocaleDetector()

        df = pd.DataFrame()

        locale_info = detector.detect(df)

        # Should return default locale
        assert locale_info is not None
