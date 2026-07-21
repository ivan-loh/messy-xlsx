"""Handler registry for routing files to appropriate format handlers."""

# ============================================================================
# Imports
# ============================================================================

from collections.abc import Iterator
from contextlib import contextmanager

import pandas as pd

from messy_xlsx._source import SourceHandle, SourceInput
from messy_xlsx.detection.format_detector import FormatDetector
from messy_xlsx.exceptions import FormatError
from messy_xlsx.models import FormatInfo
from messy_xlsx.parsing.base_handler import FormatHandler, ParseOptions
from messy_xlsx.parsing.csv_handler import CSVHandler
from messy_xlsx.parsing.xls_handler import XLSHandler
from messy_xlsx.parsing.xlsx_handler import XLSXHandler

# ============================================================================
# Core
# ============================================================================


class HandlerRegistry:
    """Registry of format handlers with automatic detection and fallback."""

    _accepts_source_handle = True

    def __init__(
        self,
        handlers: list[FormatHandler] | None = None,
        detector: FormatDetector | None = None,
    ) -> None:
        """Initialize a registry, optionally with caller-supplied components."""
        self.handlers = (
            list(handlers)
            if handlers is not None
            else [
                XLSXHandler(),
                XLSHandler(),
                CSVHandler(),
            ]
        )
        self.detector = detector or FormatDetector()

    def register_handler(self, handler: FormatHandler, priority: int = -1) -> None:
        """Register a custom handler."""
        if priority < 0:
            self.handlers.append(handler)
        else:
            self.handlers.insert(priority, handler)

    def get_handler(self, format_type: str) -> FormatHandler | None:
        """Get handler for a specific format type."""
        for handler in self.handlers:
            if handler.can_handle(format_type):
                return handler
        return None

    def detect_format(
        self,
        file_source: SourceInput | SourceHandle,
        filename: str | None = None,
    ) -> FormatInfo:
        """Detect file format."""
        with self._source_handle(file_source, filename) as source:
            return self._detect(source, filename)

    def parse(
        self,
        file_source: SourceInput | SourceHandle,
        sheet: str | None = None,
        options: ParseOptions | None = None,
        format_type: str | None = None,
    ) -> pd.DataFrame:
        """Parse file with automatic format detection and fallback."""
        with self._source_handle(file_source) as source:
            file_desc = source.description
            options = options or ParseOptions()

            if format_type is None:
                format_info = self._detect(source)
                format_type = format_info.format_type

            handler = self.get_handler(format_type)

            if handler is None:
                raise FormatError(
                    f"No handler available for format: {format_type}",
                    file_path=file_desc,
                    detected_format=format_type,
                )

            try:
                return self._parse_with(handler, source, sheet, options)
            except (PermissionError, FileNotFoundError, MemoryError):
                raise
            except Exception:
                pass

            attempted_handlers = [handler.__class__.__name__]
            for fallback_handler in self.handlers:
                if fallback_handler == handler or not fallback_handler.can_handle(format_type):
                    continue

                attempted_handlers.append(fallback_handler.__class__.__name__)
                try:
                    return self._parse_with(fallback_handler, source, sheet, options)
                except (PermissionError, FileNotFoundError, MemoryError):
                    raise
                except Exception:
                    continue

            name = source.path.name if source.path is not None else file_desc
            raise FormatError(
                f"All handlers failed for {name}",
                file_path=file_desc,
                detected_format=format_type,
                attempted_formats=attempted_handlers,
            )

    def get_sheet_names(
        self,
        file_source: SourceInput | SourceHandle,
        format_type: str | None = None,
    ) -> list[str]:
        """Get sheet names from file."""
        with self._source_handle(file_source) as source:
            if format_type is None:
                format_info = self._detect(source)
                format_type = format_info.format_type

            handler = self.get_handler(format_type)

            if handler is None:
                return ["Sheet1"]

            try:
                return self._sheet_names_with(handler, source)
            except (PermissionError, FileNotFoundError, MemoryError):
                raise
            except Exception:
                pass

            for fallback_handler in self.handlers:
                if fallback_handler == handler or not fallback_handler.can_handle(format_type):
                    continue

                try:
                    return self._sheet_names_with(fallback_handler, source)
                except (PermissionError, FileNotFoundError, MemoryError):
                    raise
                except Exception:
                    continue

            return ["Sheet1"]

    def validate(
        self,
        file_source: SourceInput | SourceHandle,
        format_type: str | None = None,
    ) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        with self._source_handle(file_source) as source:
            if format_type is None:
                try:
                    format_info = self._detect(source)
                    format_type = format_info.format_type
                except FormatError as e:
                    return False, str(e)

            if format_type == "unknown":
                return False, "Unknown file format"

            handler = self.get_handler(format_type)

            if handler is None:
                return False, f"No handler for format: {format_type}"

            return self._validate_with(handler, source)

    @contextmanager
    def _source_handle(
        self,
        source: SourceInput | SourceHandle,
        filename: str | None = None,
    ) -> Iterator[SourceHandle]:
        """Keep caller-owned handles alive and close local snapshots."""
        created = not isinstance(source, SourceHandle)
        handle = SourceHandle.coerce(source, filename=filename)
        try:
            yield handle
        finally:
            if created:
                handle.close()

    def _detect(self, source: SourceHandle, filename: str | None = None) -> FormatInfo:
        accepts_handle = bool(type(self.detector).__dict__.get("_accepts_source_handle", False))
        if accepts_handle:
            if filename is None:
                return self.detector.detect(source)
            return self.detector.detect(source, filename=filename)
        with source.open_legacy() as legacy_source:
            if filename is None:
                return self.detector.detect(legacy_source)
            return self.detector.detect(legacy_source, filename=filename)

    @staticmethod
    def _handler_accepts_source_handle(handler: FormatHandler) -> bool:
        return bool(type(handler).__dict__.get("_accepts_source_handle", False))

    def _parse_with(
        self,
        handler: FormatHandler,
        source: SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        if self._handler_accepts_source_handle(handler):
            return handler.parse(source, sheet, options)  # type: ignore[arg-type]
        with source.open_legacy() as legacy_source:
            return handler.parse(legacy_source, sheet, options)

    def _sheet_names_with(
        self,
        handler: FormatHandler,
        source: SourceHandle,
    ) -> list[str]:
        if self._handler_accepts_source_handle(handler):
            return handler.get_sheet_names(source)  # type: ignore[arg-type]
        with source.open_legacy() as legacy_source:
            return handler.get_sheet_names(legacy_source)

    def _validate_with(
        self,
        handler: FormatHandler,
        source: SourceHandle,
    ) -> tuple[bool, str | None]:
        if self._handler_accepts_source_handle(handler):
            return handler.validate(source)  # type: ignore[arg-type]
        with source.open_legacy() as legacy_source:
            return handler.validate(legacy_source)


# ============================================================================
# Module Entrypoint
# ============================================================================

_registry = HandlerRegistry()


def get_registry() -> HandlerRegistry:
    """Get the global handler registry."""
    return _registry
