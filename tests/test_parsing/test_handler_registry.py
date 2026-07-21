"""Unit tests for HandlerRegistry."""

import pandas as pd
import pytest

from messy_xlsx.exceptions import FormatError
from messy_xlsx.parsing import CSVHandler, HandlerRegistry, XLSXHandler
from messy_xlsx.parsing.base_handler import FormatHandler, ParseOptions


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
            file_source=sample_xlsx, sheet="Data", options=ParseOptions(), format_type="xlsx"
        )

        assert df is not None

    def test_fallback_only_uses_handlers_for_detected_format(self, sample_xlsx):
        """An incompatible handler must never be used as a recovery path."""

        class FailingXlsxHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                raise ValueError("primary failed")

            def get_sheet_names(self, file_source):
                return ["Data"]

            def validate(self, file_source):
                return True, None

        class IncompatibleHandler(FormatHandler):
            called = False

            def can_handle(self, format_type):
                return format_type == "unrelated"

            def parse(self, file_source, sheet, options):
                self.called = True
                return pd.DataFrame({"wrong": [1]})

            def get_sheet_names(self, file_source):
                return ["Wrong"]

            def validate(self, file_source):
                return True, None

        incompatible = IncompatibleHandler()
        registry = HandlerRegistry(handlers=[FailingXlsxHandler(), incompatible])

        with pytest.raises(FormatError, match="All handlers failed"):
            registry.parse(sample_xlsx, "Data", ParseOptions(), format_type="xlsx")

        assert incompatible.called is False

    def test_fallback_preserves_compatible_custom_handlers(self, sample_xlsx):
        """A second handler for the same format remains a valid fallback."""

        class FailingXlsxHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                raise ValueError("primary failed")

            def get_sheet_names(self, file_source):
                return ["Data"]

            def validate(self, file_source):
                return True, None

        class BackupXlsxHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                return pd.DataFrame({"recovered": [True]})

            def get_sheet_names(self, file_source):
                return ["Data"]

            def validate(self, file_source):
                return True, None

        registry = HandlerRegistry(handlers=[FailingXlsxHandler(), BackupXlsxHandler()])

        result = registry.parse(sample_xlsx, "Data", ParseOptions(), format_type="xlsx")

        assert result.to_dict(orient="list") == {"recovered": [True]}
