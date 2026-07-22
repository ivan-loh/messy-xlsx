"""Immutable models for bounded OOXML metadata."""

from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass, field


@dataclass(frozen=True)
class OoxmlLimits:
    """Resource ceilings applied before and during OOXML metadata parsing."""

    max_members: int = 10_000
    max_total_uncompressed: int = 2 * 1024**3
    max_xml_uncompressed: int = 512 * 1024**2
    suspicious_ratio_size: int = 64 * 1024**2
    max_compression_ratio: float = 1_000.0
    max_formula_samples: int = 256
    max_xml_depth: int = 256
    max_element_attributes: int = 256
    max_element_text: int = 16 * 1024 * 1024


DEFAULT_LIMITS = OoxmlLimits()


@dataclass(frozen=True, order=True)
class Interval:
    """Inclusive, one-based coordinate interval."""

    start: int
    end: int

    def __post_init__(self) -> None:
        if self.start < 1 or self.end < self.start:
            raise ValueError("invalid one-based interval")


@dataclass(frozen=True)
class IntervalIndex:
    """Normalized compact intervals with logarithmic coordinate lookup."""

    intervals: tuple[Interval, ...]
    starts: tuple[int, ...] = field(init=False)

    def __post_init__(self) -> None:
        normalized: list[Interval] = []
        for interval in sorted(self.intervals):
            if normalized and interval.start <= normalized[-1].end + 1:
                previous = normalized[-1]
                normalized[-1] = Interval(previous.start, max(previous.end, interval.end))
            else:
                normalized.append(interval)
        compact = tuple(normalized)
        object.__setattr__(self, "intervals", compact)
        object.__setattr__(self, "starts", tuple(interval.start for interval in compact))

    def contains(self, value: int) -> bool:
        """Return whether a one-based coordinate is covered."""
        position = bisect_right(self.starts, value) - 1
        return position >= 0 and value <= self.intervals[position].end


@dataclass(frozen=True)
class MergeRange:
    """Inclusive one-based rectangular merged-cell range."""

    min_row: int
    min_col: int
    max_row: int
    max_col: int

    def __post_init__(self) -> None:
        if (
            self.min_row < 1
            or self.min_col < 1
            or self.max_row < self.min_row
            or self.max_col < self.min_col
        ):
            raise ValueError("invalid one-based merge range")


@dataclass(frozen=True)
class CellEvidence:
    """Value-free OOXML provenance for one structurally sampled cell."""

    row: int
    column: int
    data_type: str
    has_value: bool
    has_formula: bool
    number_format: str = "General"

    def __post_init__(self) -> None:
        if self.row < 1 or self.column < 1:
            raise ValueError("cell evidence requires positive one-based coordinates")


@dataclass(frozen=True)
class SheetManifest:
    """Metadata-only index from one worksheet XML pass."""

    name: str
    target: str
    declared_dimension: tuple[int, int, int, int] | None
    observed_max_row: int
    observed_max_col: int
    hidden_rows: IntervalIndex
    hidden_columns: IntervalIndex
    merged_ranges: tuple[MergeRange, ...]
    has_formulas: bool
    formula_samples: tuple[str, ...]
    number_format_codes: tuple[str, ...] = ()
    observed_min_col: int = 0
    semantic_data_region: tuple[int, int, int, int] = (1, 1, 1, 1)
    semantic_nonempty_rows: IntervalIndex = field(default_factory=lambda: IntervalIndex(()))
    cell_evidence: tuple[CellEvidence, ...] = ()
    legacy_has_formulas: bool = False


@dataclass(frozen=True)
class SheetDescriptor:
    """Ordered workbook-level metadata for one worksheet."""

    name: str
    relationship_id: str
    target: str
    state: str


@dataclass(frozen=True)
class StyleManifest:
    """Number-format evidence needed by later structural planning."""

    custom_number_formats: tuple[tuple[int, str], ...]
    date_style_ids: tuple[int, ...]
    number_format_codes: tuple[str, ...] = ()


@dataclass(frozen=True)
class WorkbookManifest:
    """Metadata-only description of one OOXML workbook package."""

    workbook_type: str
    date_system: str
    sheets: tuple[SheetDescriptor, ...]
    has_shared_strings: bool
    shared_strings_uncompressed_size: int
    styles: StyleManifest
    external_relationships: tuple[str, ...] = field(default_factory=tuple)
