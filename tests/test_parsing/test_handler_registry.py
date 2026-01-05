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

        handler = registry.get_handler("unsupported")
        assert handler is None

    def test_parse_with_fallback(self, sample_xlsx):
        """Test parsing with fallback chain."""
        from messy_xlsx.parsing import ParseOptions
        registry = HandlerRegistry()

        # parse() method already implements fallback logic
        df = registry.parse(
            file_source=sample_xlsx,
            sheet="Data",
            options=ParseOptions(),
            format_type="xlsx"
        )

        assert df is not None
