"""Unit tests for HandlerRegistry."""

import pytest
from messy_xlsx.parsing import HandlerRegistry, XLSXHandler, CSVHandler


class TestHandlerRegistry:
    """Test handler registry functionality."""

    def test_get_xlsx_handler(self):
        """Test getting XLSX handler."""
        registry = HandlerRegistry()

        handler = registry.get_handler("xlsx")

        assert handler is not None
        assert isinstance(handler, XLSXHandler)

    def test_get_csv_handler(self):
        """Test getting CSV handler."""
        registry = HandlerRegistry()

        handler = registry.get_handler("csv")

        assert handler is not None
        assert isinstance(handler, CSVHandler)

    def test_get_xlsm_handler(self):
        """Test getting XLSM handler (same as XLSX)."""
        registry = HandlerRegistry()

        handler = registry.get_handler("xlsm")

        assert handler is not None
        assert isinstance(handler, XLSXHandler)

    def test_unsupported_format(self):
        """Test handling unsupported format."""
        registry = HandlerRegistry()

        with pytest.raises(Exception):
            registry.get_handler("unsupported")

    def test_parse_with_fallback(self, sample_xlsx):
        """Test parsing with fallback chain."""
        registry = HandlerRegistry()

        df = registry.parse_with_fallback(
            file_path=sample_xlsx,
            sheet="Data",
            config=None,
            format_hint="xlsx"
        )

        assert df is not None
