"""Enumeration types for messy-xlsx configuration."""

# ============================================================================
# Imports
# ============================================================================

from enum import StrEnum

# ============================================================================
# Data Type Enums
# ============================================================================


class DataType(StrEnum):
    """Cell data types."""

    EMPTY = "empty"
    TEXT = "text"
    NUMBER = "number"
    BOOLEAN = "boolean"
    DATE = "date"
    ERROR = "error"
    FORMULA = "formula"


class FormatType(StrEnum):
    """File format types."""

    XLSX = "xlsx"
    XLSM = "xlsm"
    XLSB = "xlsb"
    XLS = "xls"
    CSV = "csv"
    TSV = "tsv"
    TXT = "txt"
    UNKNOWN = "unknown"


# ============================================================================
# Configuration Enums
# ============================================================================


class MergeStrategy(StrEnum):
    """Strategies for handling merged cells."""

    FILL = "fill"
    SKIP = "skip"
    FIRST_ONLY = "first_only"


class HeaderDetectionMode(StrEnum):
    """Modes for header row detection."""

    SMART = "smart"
    AUTO = "auto"
    MANUAL = "manual"


class HeaderFallback(StrEnum):
    """Fallback strategies when header detection fails."""

    FIRST_ROW = "first_row"
    NONE = "none"
    ERROR = "error"
