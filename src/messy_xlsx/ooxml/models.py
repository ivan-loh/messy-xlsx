"""Immutable models for bounded OOXML metadata."""

from __future__ import annotations

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
