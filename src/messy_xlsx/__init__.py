"""messy-xlsx: A Python library for parsing messy Excel files."""

# ============================================================================
# Imports
# ============================================================================

from messy_xlsx.exceptions import (
    CircularReferenceError,
    FileError,
    FormulaError,
    FormatError,
    MessyXlsxError,
    NormalizationError,
    StructureError,
    UnsupportedFunctionError,
)
from messy_xlsx.formulas.config import (
    CircularRefStrategy,
    FormulaConfig,
    FormulaEvaluationMode,
)
from messy_xlsx.models import (
    CellValue,
    FormatInfo,
    SheetConfig,
    StructureInfo,
)
from messy_xlsx.workbook import MessyWorkbook


# ============================================================================
# Package Metadata
# ============================================================================

__version__ = "0.1.0"

__all__ = [
    "MessyWorkbook",
    "CellValue",
    "FormatInfo",
    "SheetConfig",
    "StructureInfo",
    "MessyXlsxError",
    "FileError",
    "FormatError",
    "StructureError",
    "NormalizationError",
    "FormulaError",
    "CircularReferenceError",
    "UnsupportedFunctionError",
    "FormulaConfig",
    "FormulaEvaluationMode",
    "CircularRefStrategy",
]


# ============================================================================
# Convenience Functions
# ============================================================================

def read_excel(file_path: str, sheet: str | None = None, **config_kwargs):
    """Quick function to read an Excel file to a pandas DataFrame."""
    config = SheetConfig(**config_kwargs) if config_kwargs else None
    wb = MessyWorkbook(file_path, sheet_config=config)
    return wb.to_dataframe(sheet=sheet)


def read_excel_tables(file_path: str, sheet: str | None = None) -> list:
    """Read all detected tables from a sheet."""
    wb = MessyWorkbook(file_path)
    sheet_obj = wb.get_sheet(sheet or wb.sheet_names[0])
    return [table.to_dataframe() for table in sheet_obj.tables]


def analyze_structure(file_path: str, sheet: str | None = None) -> StructureInfo:
    """Analyze and return structure info without loading data."""
    wb = MessyWorkbook(file_path)
    return wb.get_structure(sheet or wb.sheet_names[0])
