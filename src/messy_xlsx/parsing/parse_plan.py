"""Pure compilation of sheet configuration into an immutable parse plan.

This module deliberately contains no source, workbook, worksheet, or DataFrame
access.  It resolves configuration and optional structure evidence into stable
projections consumed by the existing parsing and normalization layers.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Final

from messy_xlsx.enums import (
    FormatType,
    HeaderDetectionMode,
    HeaderFallback,
    MergeStrategy,
)
from messy_xlsx.exceptions import StructureError
from messy_xlsx.models import SheetConfig, StructureInfo
from messy_xlsx.parsing.base_handler import ParseOptions

_STRUCTURE_FORMATS: Final = frozenset({"xlsx", "xlsm", "xltx", "xltm"})
_TEXT_FORMATS: Final = frozenset({"csv", "tsv", "txt"})

_COMMA_DECIMAL_DOT_THOUSANDS: Final = frozenset(
    {
        "de",
        "nl",
        "it",
        "es",
        "pt",
        "el",
        "tr",
        "id",
    }
)
_COMMA_DECIMAL_SPACE_THOUSANDS: Final = frozenset(
    {
        "fr",
        "sv",
        "nb",
        "nn",
        "fi",
        "pl",
        "cs",
        "sk",
        "hu",
        "ro",
        "bg",
        "hr",
        "sl",
        "sr",
        "da",
        "ru",
        "uk",
    }
)


@dataclass(frozen=True, slots=True)
class ParsePlan:
    """Frozen internal decisions for one sheet parse.

    Configuration containers are projected to tuples. Arbitrary condition
    payload objects retain their identity so comparison behavior is unchanged.
    """

    # Projection for format handlers.
    skip_rows: int
    header_rows: int
    skip_footer: int
    merge_strategy: MergeStrategy | str
    ignore_hidden: bool
    cell_range: str | None
    data_only: bool
    auto_detect_header: bool

    # Projection for normalization.
    normalize: bool
    decimal_separator: str | None
    thousands_separator: str | None
    use_extended_missing_list: bool
    preserve_types: bool
    type_hints: tuple[tuple[str, str], ...]
    skip_normalization_steps: tuple[str, ...]

    # Projection for post-processing.
    sanitize_column_names: bool
    column_renames: tuple[tuple[str, str], ...]
    drop_regex: str | None
    drop_conditions: tuple[tuple[Any, Any], ...]

    def to_parse_options(self) -> ParseOptions:
        """Return a fresh mutable projection for the existing handler API."""
        return ParseOptions(
            skip_rows=self.skip_rows,
            header_rows=self.header_rows,
            skip_footer=self.skip_footer,
            merge_strategy=self.merge_strategy,
            ignore_hidden=self.ignore_hidden,
            cell_range=self.cell_range,
            data_only=self.data_only,
            auto_detect_header=self.auto_detect_header,
        )


def requires_structure_analysis(
    config: SheetConfig,
    format_type: FormatType | str,
) -> bool:
    """Return whether current behavior requires OOXML structure evidence."""
    return config.auto_detect and format_type in _STRUCTURE_FORMATS


def compile_parse_plan(
    config: SheetConfig,
    structure: StructureInfo | None,
    format_type: FormatType | str,
) -> ParsePlan:
    """Compile configuration and optional structure evidence without mutation or I/O."""
    use_structure = requires_structure_analysis(config, format_type)
    if use_structure and structure is None:
        raise ValueError("StructureInfo is required for auto-detected OOXML parsing")

    active_structure = structure if use_structure else None

    skip_rows = config.skip_rows
    header_rows = config.header_rows
    skip_footer = config.skip_footer
    merge_strategy = config.merge_strategy
    ignore_hidden = not config.include_hidden
    locale = config.locale

    if active_structure is not None:
        skip_rows, header_rows = _resolve_header_rows(config, active_structure)
        skip_footer = _resolve_skip_footer(config, active_structure)
        locale = config.locale or active_structure.detected_locale

        if not active_structure.merged_ranges:
            merge_strategy = "skip"
        if not active_structure.hidden_rows and not active_structure.hidden_columns:
            ignore_hidden = False

    decimal_separator, thousands_separator = _resolve_separators(config, locale)

    skipped_steps: list[str] = []
    if not config.normalize_whitespace:
        skipped_steps.append("whitespace")
    if not config.normalize_numbers:
        skipped_steps.append("numbers")
    if not config.normalize_dates:
        skipped_steps.append("dates")
    if not config.ensure_type_consistency:
        skipped_steps.append("type_coercion")

    return ParsePlan(
        skip_rows=skip_rows,
        header_rows=header_rows,
        skip_footer=skip_footer,
        merge_strategy=merge_strategy,
        ignore_hidden=ignore_hidden,
        cell_range=config.cell_range,
        data_only=config.evaluate_formulas,
        auto_detect_header=config.auto_detect and format_type in _TEXT_FORMATS,
        normalize=config.normalize,
        decimal_separator=decimal_separator,
        thousands_separator=thousands_separator,
        use_extended_missing_list=config.use_extended_missing_list,
        preserve_types=config.preserve_types,
        type_hints=tuple(config.type_hints.items()),
        skip_normalization_steps=tuple(skipped_steps),
        sanitize_column_names=config.sanitize_column_names,
        column_renames=tuple(config.column_renames.items()),
        drop_regex=config.drop_regex,
        drop_conditions=tuple(
            (
                condition.get("column"),
                condition.get("value"),
            )
            for condition in config.drop_conditions
        ),
    )


def _resolve_header_rows(
    config: SheetConfig,
    structure: StructureInfo,
) -> tuple[int, int]:
    """Resolve row offsets while preserving existing header-mode semantics."""
    skip_rows = config.skip_rows
    header_rows = config.header_rows
    header_is_usable = (
        structure.header_row is not None
        and structure.header_confidence >= config.header_confidence_threshold
    )

    if config.header_detection_mode == HeaderDetectionMode.AUTO:
        if header_is_usable:
            skip_rows, header_rows = _detected_header_rows(config, structure)
        elif config.header_fallback == HeaderFallback.FIRST_ROW:
            skip_rows = 0
            header_rows = 1
        elif config.header_fallback == HeaderFallback.NONE:
            header_rows = 0
        elif config.header_fallback == HeaderFallback.ERROR:
            raise StructureError(
                f"No header detected with sufficient confidence "
                f"(found: {structure.header_confidence:.2f}, "
                f"required: {config.header_confidence_threshold:.2f})"
            )
    elif (
        config.header_detection_mode == HeaderDetectionMode.SMART
        and config.skip_rows == 0
        and header_is_usable
    ):
        skip_rows, header_rows = _detected_header_rows(config, structure)

    return skip_rows, header_rows


def _detected_header_rows(
    config: SheetConfig,
    structure: StructureInfo,
) -> tuple[int, int]:
    """Convert a detected one-based header into handler row offsets."""
    if structure.header_row is None:
        raise ValueError("A detected header row is required")

    skip_rows = max(0, structure.header_row - 1)
    if structure.hidden_rows and not config.include_hidden:
        hidden_before_header = sum(row < structure.header_row for row in structure.hidden_rows)
        skip_rows = max(0, skip_rows - hidden_before_header)

    return skip_rows, structure.header_rows_count


def _resolve_skip_footer(config: SheetConfig, structure: StructureInfo) -> int:
    """Resolve footer trimming with caller and detected-table precedence."""
    skip_footer = structure.suggested_skip_footer
    if config.skip_footer > 0:
        return config.skip_footer

    # A caller-selected range already constrains the parse. Sheet-global
    # detected footer evidence cannot be mapped safely into that local range.
    if config.cell_range:
        return 0

    if structure.num_tables > 1 and structure.table_ranges:
        first_table_end = structure.table_ranges[0]["end_row"]
        rows_after_first_table = max(0, structure.data_end_row - first_table_end)

        if structure.hidden_rows and not config.include_hidden:
            hidden_after = sum(
                first_table_end < row <= structure.data_end_row for row in structure.hidden_rows
            )
            rows_after_first_table = max(0, rows_after_first_table - hidden_after)

        if rows_after_first_table > 0:
            skip_footer = max(skip_footer, rows_after_first_table)

    return skip_footer


def _resolve_separators(
    config: SheetConfig,
    locale: str | None,
) -> tuple[str | None, str | None]:
    """Resolve explicit or locale-derived number separators."""
    decimal_separator = config.decimal_separator
    thousands_separator = config.thousands_separator
    if decimal_separator is None and thousands_separator is None and locale and locale != "auto":
        return _locale_to_separators(locale)
    return decimal_separator, thousands_separator


def _locale_to_separators(locale: str) -> tuple[str, str]:
    """Convert an existing locale convention to number separators."""
    lang = locale.split("_")[0].lower()
    country = locale.split("_")[1].upper() if "_" in locale else ""

    if lang == "de" and country == "CH":
        return ".", "'"
    if lang in _COMMA_DECIMAL_DOT_THOUSANDS:
        return ",", "."
    if lang in _COMMA_DECIMAL_SPACE_THOUSANDS:
        return ",", " "
    return ".", ","
