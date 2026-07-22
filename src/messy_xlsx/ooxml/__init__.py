"""Bounded, metadata-only OOXML workbook inspection."""

from messy_xlsx.ooxml.manifest import ManifestReader, build_manifest
from messy_xlsx.ooxml.models import (
    DEFAULT_LIMITS,
    Interval,
    IntervalIndex,
    MergeRange,
    OoxmlLimits,
    SheetDescriptor,
    SheetManifest,
    StyleManifest,
    WorkbookManifest,
)

__all__ = [
    "DEFAULT_LIMITS",
    "Interval",
    "IntervalIndex",
    "ManifestReader",
    "MergeRange",
    "OoxmlLimits",
    "SheetDescriptor",
    "SheetManifest",
    "StyleManifest",
    "WorkbookManifest",
    "build_manifest",
]
