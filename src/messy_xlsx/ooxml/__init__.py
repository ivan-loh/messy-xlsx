"""Bounded, metadata-only OOXML workbook inspection."""

from messy_xlsx.ooxml.manifest import build_manifest
from messy_xlsx.ooxml.models import (
    DEFAULT_LIMITS,
    OoxmlLimits,
    SheetDescriptor,
    StyleManifest,
    WorkbookManifest,
)

__all__ = [
    "DEFAULT_LIMITS",
    "OoxmlLimits",
    "SheetDescriptor",
    "StyleManifest",
    "WorkbookManifest",
    "build_manifest",
]
