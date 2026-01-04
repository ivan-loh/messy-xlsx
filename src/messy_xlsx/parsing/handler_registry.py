"""Handler registry for routing files to appropriate format handlers."""

# ============================================================================
# Imports
# ============================================================================

from pathlib import Path
from typing import BinaryIO

import pandas as pd

from messy_xlsx.detection.format_detector import FormatDetector
from messy_xlsx.exceptions import FormatError
from messy_xlsx.models import FormatInfo
from messy_xlsx.parsing.base_handler import FileSource, FormatHandler, ParseOptions
from messy_xlsx.parsing.csv_handler import CSVHandler
from messy_xlsx.parsing.xls_handler import XLSHandler
from messy_xlsx.parsing.xlsx_handler import XLSXHandler


# ============================================================================
# Core
# ============================================================================

class HandlerRegistry:
    """Registry of format handlers with automatic detection and fallback."""

    def __init__(self):
        """Initialize registry with default handlers."""
        self.handlers: list[FormatHandler] = [
            XLSXHandler(),
            XLSHandler(),
            CSVHandler(),
        ]
        self.detector = FormatDetector()

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

    def detect_format(self, file_source: FileSource, filename: str | None = None) -> FormatInfo:
        """Detect file format."""
        return self.detector.detect(file_source, filename=filename)

    def parse(
        self,
        file_source: FileSource,
        sheet: str | None = None,
        options: ParseOptions | None = None,
        format_type: str | None = None,
    ) -> pd.DataFrame:
        """Parse file with automatic format detection and fallback."""
        is_fileobj = hasattr(file_source, "read")
        file_desc = "<stream>" if is_fileobj else str(file_source)
        options = options or ParseOptions()

        if format_type is None:
            format_info = self.detector.detect(file_source)
            format_type = format_info.format_type

        handler = self.get_handler(format_type)

        if handler is None:
            raise FormatError(
                f"No handler available for format: {format_type}",
                file_path        = file_desc,
                detected_format  = format_type,
            )

        errors = []
        try:
            return handler.parse(file_source, sheet, options)
        except (PermissionError, FileNotFoundError, MemoryError):
            raise
        except Exception as e:
            errors.append(f"{handler.__class__.__name__}: {e}")

        for fallback_handler in self.handlers:
            if fallback_handler == handler:
                continue

            try:
                df = fallback_handler.parse(file_source, sheet, options)
                return df
            except (PermissionError, FileNotFoundError, MemoryError):
                raise
            except Exception as e:
                errors.append(f"{fallback_handler.__class__.__name__}: {e}")
                continue

        name = file_desc if is_fileobj else Path(file_source).name
        raise FormatError(
            f"All handlers failed for {name}",
            file_path          = file_desc,
            detected_format    = format_type,
            attempted_formats  = [h.__class__.__name__ for h in self.handlers],
        )

    def get_sheet_names(
        self,
        file_source: FileSource,
        format_type: str | None = None,
    ) -> list[str]:
        """Get sheet names from file."""
        is_fileobj = hasattr(file_source, "read")

        if format_type is None:
            format_info = self.detector.detect(file_source)
            format_type = format_info.format_type

        handler = self.get_handler(format_type)

        if handler is None:
            return ["Sheet1"]

        errors = []

        try:
            return handler.get_sheet_names(file_source)
        except (PermissionError, FileNotFoundError, MemoryError):
            raise
        except Exception as e:
            errors.append(f"{handler.__class__.__name__}: {e}")

        for fallback_handler in self.handlers:
            if fallback_handler == handler:
                continue

            try:
                return fallback_handler.get_sheet_names(file_source)
            except (PermissionError, FileNotFoundError, MemoryError):
                raise
            except Exception:
                continue

        return ["Sheet1"]

    def validate(
        self,
        file_source: FileSource,
        format_type: str | None = None,
    ) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        if format_type is None:
            try:
                format_info = self.detector.detect(file_source)
                format_type = format_info.format_type
            except FormatError as e:
                return False, str(e)

        if format_type == "unknown":
            return False, "Unknown file format"

        handler = self.get_handler(format_type)

        if handler is None:
            return False, f"No handler for format: {format_type}"

        return handler.validate(file_source)


# ============================================================================
# Module Entrypoint
# ============================================================================

_registry = HandlerRegistry()


def get_registry() -> HandlerRegistry:
    """Get the global handler registry."""
    return _registry
