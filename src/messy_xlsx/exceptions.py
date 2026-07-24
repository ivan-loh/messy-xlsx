"""Custom exception hierarchy for messy-xlsx."""

# ============================================================================
# Imports
# ============================================================================

import re
from typing import Any

from messy_xlsx._fallback_signals import (
    _FallbackBlockReason,
    _mark_fallback_blocked,
)

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


class StreamingTypeError(NormalizationError):
    """A streamed value cannot fit the schema fixed before iteration."""

    _MAX_CONTEXT_TEXT = 160
    _MESSAGE = "streamed value is incompatible with the fixed schema"
    _STRUCTURAL_TEXT = re.compile(
        r"^(?:str|bytes)(?: label)?\(length=\d+\)$|^(?:(?:int|float|bool) label|int|float|bool|date|datetime|time|timedelta|unsupported value|non-string label)$"
    )

    def __init__(
        self,
        message: str,
        *,
        ordinal: int,
        display_label: str,
        row_offset: int,
        value_description: str,
        expected_type: str,
    ) -> None:
        if type(ordinal) is not int or ordinal < 0:
            raise TypeError("ordinal must be a non-negative exact int")
        if type(row_offset) is not int or row_offset < 0:
            raise TypeError("row_offset must be a non-negative exact int")
        context_values = {
            "display_label": display_label,
            "value_description": value_description,
            "expected_type": expected_type,
        }
        if any(type(value) is not str for value in context_values.values()):
            raise TypeError("streaming error text context must use exact strings")
        del message
        safe_label = self._safe_label(display_label)
        safe_value = self._safe_value_description(value_description)
        safe_expected = self._safe_expected_type(expected_type)
        super().__init__(
            self._MESSAGE,
            value=None,
            expected_type=safe_expected,
            ordinal=ordinal,
            display_label=safe_label,
            row_offset=row_offset,
            value_description=safe_value,
        )
        _mark_fallback_blocked(self, _FallbackBlockReason.CONFIGURATION)

    @classmethod
    def _safe_label(cls, value: str) -> str:
        bounded = value[: cls._MAX_CONTEXT_TEXT]
        if cls._STRUCTURAL_TEXT.fullmatch(bounded):
            return bounded
        return f"str label(length={len(value)})"

    @classmethod
    def _safe_value_description(cls, value: str) -> str:
        bounded = value[: cls._MAX_CONTEXT_TEXT]
        return bounded if cls._STRUCTURAL_TEXT.fullmatch(bounded) else "unsupported value"

    @classmethod
    def _safe_expected_type(cls, value: str) -> str:
        bounded = value[: cls._MAX_CONTEXT_TEXT]
        return bounded if re.fullmatch(r"[A-Za-z0-9_+\-.,: =\[\]()]+", bounded) else "arrow type"


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
