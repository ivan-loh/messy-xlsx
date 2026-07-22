"""Internal capability and output-mode backend routing."""

from __future__ import annotations

from dataclasses import dataclass

from messy_xlsx.enums import FormatType
from messy_xlsx.parsing.contracts import BackendKind, OutputMode, ReaderDecision

_OOXML_FORMATS = frozenset({"xlsx", "xlsm", "xltx", "xltm"})
_TEXT_FORMATS = frozenset({"csv", "tsv", "txt"})


@dataclass(frozen=True, slots=True)
class WorkbookContext:
    """Routing inputs compiled without opening a parser backend."""

    format_type: FormatType | str
    output_mode: OutputMode
    evaluate_formulas: bool
    has_custom_registry: bool


class BackendRouter:
    """Select an internal backend from capabilities and requested output."""

    def select(self, context: WorkbookContext) -> ReaderDecision:
        """Return the backend decision without performing parser I/O."""
        if context.has_custom_registry:
            return ReaderDecision(
                BackendKind.CUSTOM_DATAFRAME,
                "caller extension compatibility",
            )

        format_type = str(context.format_type).lower().removeprefix(".")
        output_mode = OutputMode(context.output_mode)

        if format_type in _TEXT_FORMATS:
            return ReaderDecision(BackendKind.CSV_STREAMING, "text chunk reader")
        if format_type == "xls":
            return ReaderDecision(BackendKind.XLS_STREAMING, "optional xlrd row reader")
        if format_type not in _OOXML_FORMATS:
            raise ValueError(f"Unsupported format for backend routing: {format_type}")
        if output_mode is OutputMode.STREAMING:
            return ReaderDecision(
                BackendKind.OPENPYXL_STREAMING,
                "fastexcel has no batch iterator",
            )
        if context.evaluate_formulas:
            return ReaderDecision(
                BackendKind.FASTEXCEL,
                "fast cached-value materialization",
            )
        return ReaderDecision(
            BackendKind.OPENPYXL_COMPAT,
            "formula expressions required",
        )
