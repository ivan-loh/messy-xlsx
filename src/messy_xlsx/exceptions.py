"""Custom exception hierarchy for messy-xlsx."""

# ============================================================================
# Imports
# ============================================================================

from typing import Any


# ============================================================================
# Base Exception
# ============================================================================

class MessyXlsxError(Exception):
    """Base exception for all messy-xlsx errors."""

    def __init__(self, message: str, context: dict[str, Any] | None = None):
        self.message = message
        self.context = context or {}
        super().__init__(message)

    def __str__(self) -> str:
        if self.context:
            context_str = ", ".join(f"{k}={v!r}" for k, v in self.context.items())
            return f"{self.message} ({context_str})"
        return self.message

    def to_dict(self) -> dict[str, Any]:
        """Convert exception to dictionary for serialization."""
        return {
            "error": self.__class__.__name__,
            "message": self.message,
            "context": self.context,
        }


# ============================================================================
# File-Related Exceptions
# ============================================================================

class FileError(MessyXlsxError):
    """Raised for file I/O issues."""

    def __init__(
        self,
        message: str,
        file_path: str | None = None,
        operation: str | None = None,
        **kwargs: Any,
    ):
        context = {"file_path": file_path, "operation": operation, **kwargs}
        context = {k: v for k, v in context.items() if v is not None}
        super().__init__(message, context)


class FormatError(MessyXlsxError):
    """Raised when file format cannot be determined or is unsupported."""

    def __init__(
        self,
        message: str,
        file_path: str | None = None,
        detected_format: str | None = None,
        attempted_formats: list[str] | None = None,
        **kwargs: Any,
    ):
        context = {
            "file_path": file_path,
            "detected_format": detected_format,
            "attempted_formats": attempted_formats,
            **kwargs,
        }
        context = {k: v for k, v in context.items() if v is not None}
        super().__init__(message, context)


class StructureError(MessyXlsxError):
    """Raised when structure detection fails."""

    def __init__(
        self,
        message: str,
        sheet: str | None = None,
        detection_phase: str | None = None,
        **kwargs: Any,
    ):
        context = {"sheet": sheet, "detection_phase": detection_phase, **kwargs}
        context = {k: v for k, v in context.items() if v is not None}
        super().__init__(message, context)


# ============================================================================
# Data Processing Exceptions
# ============================================================================

class NormalizationError(MessyXlsxError):
    """Raised when data normalization fails."""

    def __init__(
        self,
        message: str,
        column: str | None = None,
        row: int | None = None,
        value: Any = None,
        expected_type: str | None = None,
        **kwargs: Any,
    ):
        context = {
            "column": column,
            "row": row,
            "value": repr(value) if value is not None else None,
            "expected_type": expected_type,
            **kwargs,
        }
        context = {k: v for k, v in context.items() if v is not None}
        super().__init__(message, context)


# ============================================================================
# Formula-Related Exceptions
# ============================================================================

class FormulaError(MessyXlsxError):
    """Base exception for formula evaluation failures."""

    def __init__(
        self,
        message: str,
        cell_ref: str | None = None,
        formula: str | None = None,
        **kwargs: Any,
    ):
        context = {"cell_ref": cell_ref, "formula": formula, **kwargs}
        context = {k: v for k, v in context.items() if v is not None}
        super().__init__(message, context)


class CircularReferenceError(FormulaError):
    """Raised when a circular reference is detected during formula evaluation."""

    def __init__(
        self,
        message: str,
        cycle: list[str] | None = None,
        **kwargs: Any,
    ):
        self.cycle = cycle or []
        super().__init__(message, cycle=cycle, **kwargs)

    def __str__(self) -> str:
        if self.cycle:
            cycle_str = " -> ".join(self.cycle)
            return f"{self.message}: {cycle_str}"
        return self.message


class UnsupportedFunctionError(FormulaError):
    """Raised when a formula uses an unsupported Excel function."""

    def __init__(
        self,
        function_name: str,
        cell_ref: str | None = None,
        **kwargs: Any,
    ):
        self.function_name = function_name
        message = f"Unsupported function: {function_name}"
        super().__init__(message, cell_ref=cell_ref, function_name=function_name, **kwargs)
