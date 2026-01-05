"""Unit tests for LocaleDetector."""

from messy_xlsx.detection import LocaleDetector


class TestLocaleDetector:
    """Test locale detection functionality."""

    def test_detect_us_locale(self):
        """Test detecting US number format (1,234.56)."""
        detector = LocaleDetector()

        samples = ["1,234.56", "2,345.67", "3,456.78"]

        locale_info = detector.detect_from_samples(samples)

        assert locale_info.decimal_separator == "."
        assert locale_info.thousands_separator == ","

    def test_detect_european_locale(self):
        """Test detecting European number format (1.234,56)."""
        detector = LocaleDetector()

        samples = ["1.234,56", "2.345,67", "3.456,78"]

        locale_info = detector.detect_from_samples(samples)

        assert locale_info.decimal_separator == ","
        assert locale_info.thousands_separator == "."

    def test_mixed_formats(self):
        """Test handling mixed number formats."""
        detector = LocaleDetector()

        samples = ["1,234.56", "1.234,56", "1000"]

        locale_info = detector.detect_from_samples(samples)

        # Should return most common format
        assert locale_info.decimal_separator in [".", ","]

    def test_no_formatted_numbers(self):
        """Test handling data with no formatted numbers."""
        detector = LocaleDetector()

        samples = ["1000", "2000", "3000"]

        locale_info = detector.detect_from_samples(samples)

        # Should have default values
        assert locale_info.decimal_separator is not None
        assert locale_info.thousands_separator is not None

    def test_empty_dataframe(self):
        """Test handling empty DataFrame."""
        detector = LocaleDetector()

        samples = []

        locale_info = detector.detect_from_samples(samples)

        # Should return default locale
        assert locale_info is not None
