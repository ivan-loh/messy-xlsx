"""Unit tests for HandlerRegistry."""

import pandas as pd
import pytest

from messy_xlsx._fallback_signals import (
    _FallbackBlockReason,
    _mark_fallback_blocked,
)
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

    def test_structured_parse_failure_blocks_legacy_registry_fallback(
        self,
        sample_xlsx,
    ):
        primary_error = _mark_fallback_blocked(
            ValueError("paramètre non valide"),
            _FallbackBlockReason.CONFIGURATION,
        )
        fallback_calls = 0

        class PrimaryHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                raise primary_error

            def get_sheet_names(self, file_source):
                raise primary_error

            def validate(self, file_source):
                return True, None

        class FallbackHandler(PrimaryHandler):
            def parse(self, file_source, sheet, options):
                nonlocal fallback_calls
                fallback_calls += 1
                return pd.DataFrame({"wrong": [True]})

        registry = HandlerRegistry(handlers=[PrimaryHandler(), FallbackHandler()])

        with pytest.raises(ValueError) as captured:
            registry.parse(sample_xlsx, format_type="xlsx")

        assert captured.value is primary_error
        assert fallback_calls == 0

    def test_structured_sheet_name_failure_blocks_legacy_registry_fallback(
        self,
        sample_xlsx,
    ):
        primary_error = _mark_fallback_blocked(
            RuntimeError("source occupée"),
            _FallbackBlockReason.SOURCE_OWNERSHIP,
        )
        fallback_calls = 0

        class PrimaryHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                return pd.DataFrame()

            def get_sheet_names(self, file_source):
                raise primary_error

            def validate(self, file_source):
                return True, None

        class FallbackHandler(PrimaryHandler):
            def get_sheet_names(self, file_source):
                nonlocal fallback_calls
                fallback_calls += 1
                return ["Wrong"]

        registry = HandlerRegistry(handlers=[PrimaryHandler(), FallbackHandler()])

        with pytest.raises(RuntimeError) as captured:
            registry.get_sheet_names(sample_xlsx, format_type="xlsx")

        assert captured.value is primary_error
        assert fallback_calls == 0

    @pytest.mark.parametrize(
        "primary_error",
        [
            ExceptionGroup(
                "outer",
                [ExceptionGroup("inner", [MemoryError("capacity")])],
            ),
            ExceptionGroup(
                "outer",
                [
                    ExceptionGroup(
                        "inner",
                        [
                            _mark_fallback_blocked(
                                ValueError("paramètre non valide"),
                                _FallbackBlockReason.CONFIGURATION,
                            )
                        ],
                    )
                ],
            ),
        ],
    )
    def test_nested_unsafe_group_blocks_legacy_registry_retry(
        self,
        sample_xlsx,
        primary_error: ExceptionGroup,
    ):
        fallback_calls = 0

        class PrimaryHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                raise primary_error

            def get_sheet_names(self, file_source):
                return ["Data"]

            def validate(self, file_source):
                return True, None

        class FallbackHandler(PrimaryHandler):
            def parse(self, file_source, sheet, options):
                nonlocal fallback_calls
                fallback_calls += 1
                return pd.DataFrame({"wrong": [True]})

        registry = HandlerRegistry(handlers=[PrimaryHandler(), FallbackHandler()])

        with pytest.raises(ExceptionGroup) as captured:
            registry.parse(sample_xlsx, format_type="xlsx")

        assert captured.value is primary_error
        assert fallback_calls == 0

    @pytest.mark.parametrize(
        "primary_error",
        [
            ExceptionGroup(
                "outer",
                [ExceptionGroup("inner", [MemoryError("capacity")])],
            ),
            ExceptionGroup(
                "outer",
                [
                    ExceptionGroup(
                        "inner",
                        [
                            _mark_fallback_blocked(
                                RuntimeError("source occupée"),
                                _FallbackBlockReason.SOURCE_OWNERSHIP,
                            )
                        ],
                    )
                ],
            ),
        ],
    )
    def test_nested_unsafe_group_blocks_legacy_sheet_name_retry(
        self,
        sample_xlsx,
        primary_error: ExceptionGroup,
    ):
        fallback_calls = 0

        class PrimaryHandler(FormatHandler):
            def can_handle(self, format_type):
                return format_type == "xlsx"

            def parse(self, file_source, sheet, options):
                return pd.DataFrame()

            def get_sheet_names(self, file_source):
                raise primary_error

            def validate(self, file_source):
                return True, None

        class FallbackHandler(PrimaryHandler):
            def get_sheet_names(self, file_source):
                nonlocal fallback_calls
                fallback_calls += 1
                return ["Wrong"]

        registry = HandlerRegistry(handlers=[PrimaryHandler(), FallbackHandler()])

        with pytest.raises(ExceptionGroup) as captured:
            registry.get_sheet_names(sample_xlsx, format_type="xlsx")

        assert captured.value is primary_error
        assert fallback_calls == 0
