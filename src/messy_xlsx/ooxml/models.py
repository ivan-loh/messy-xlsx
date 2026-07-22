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


CELL_KIND_MASK = 0x07
CELL_KIND_TEXT = 0x01
CELL_KIND_NUMBER = 0x02
CELL_KIND_BOOLEAN = 0x03
CELL_KIND_DATE = 0x04
CELL_KIND_OTHER = 0x05
CELL_HAS_VALUE = 0x08
CELL_HAS_FORMULA = 0x10
CELL_EUROPEAN_FORMAT = 0x20


@dataclass(frozen=True)
class RowBitSet:
    """Compact fixed-domain truth index for one-based worksheet rows."""

    bits: bytes = b""

    def contains(self, value: int) -> bool:
        """Return whether *value* has its bit set."""
        if value < 1:
            return False
        bit = value - 1
        byte = bit >> 3
        return byte < len(self.bits) and bool(self.bits[byte] & (1 << (bit & 7)))


@dataclass(frozen=True)
class CellEvidenceIndex:
    """Packed provenance for only the coordinates used by legacy scoring."""

    start_row: int = 1
    start_col: int = 1
    end_col: int = 0
    header_row_count: int = 0
    header_codes: bytes = b""
    locale_row_offset: int = 0
    locale_row_count: int = 0
    locale_column_count: int = 0
    locale_codes: bytes = b""

    def __len__(self) -> int:
        """Return the number of retained packed coordinate slots."""
        return len(self.header_codes) + len(self.locale_codes)

    def code(self, row: int, column: int) -> int:
        """Return packed provenance for one coordinate, or zero when unneeded."""
        row_offset = row - self.start_row
        width = self.end_col - self.start_col + 1
        if 0 <= row_offset < self.header_row_count and self.start_col <= column <= self.end_col:
            offset = row_offset * width + column - self.start_col
            return self.header_codes[offset]
        locale_row = row_offset - self.locale_row_offset
        if (
            0 <= locale_row < self.locale_row_count
            and self.start_col <= column < self.start_col + self.locale_column_count
        ):
            offset = locale_row * self.locale_column_count + column - self.start_col
            return self.locale_codes[offset]
        return 0


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
    observed_min_col: int = 0
    semantic_data_region: tuple[int, int, int, int] = (1, 1, 1, 1)
    semantic_nonempty_rows: RowBitSet = field(default_factory=RowBitSet)
    cell_evidence: CellEvidenceIndex = field(default_factory=CellEvidenceIndex)
    sparse_filled_counts: bytes = b""
    locale_has_european_format: bool = False
    legacy_has_formulas: bool = False

    def sparse_filled_count(self, column: int) -> int:
        """Return the compact first-1,000-row filled count for *column*."""
        offset = (column - 1) * 2
        if column < 1 or offset + 2 > len(self.sparse_filled_counts):
            return 0
        return int.from_bytes(self.sparse_filled_counts[offset : offset + 2], "little")


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
