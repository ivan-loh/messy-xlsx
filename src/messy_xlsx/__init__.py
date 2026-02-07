"""messy-xlsx: A Python library for parsing messy Excel files."""

# ============================================================================
# Imports
# ============================================================================

from typing import Any

import pandas as pd

from messy_xlsx.exceptions import (
    CircularReferenceError,
    FileError,
    FormatError,
    FormulaError,
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
    SheetError,
    StructureInfo,
)
from messy_xlsx.multi_sheet import (
    MultiSheetOptions,
    MultiSheetParser,
    SheetInfo,
    analyze_excel,
    read_all_sheets,
)
from messy_xlsx.utils import sanitize_column_name
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
    "SheetError",
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
    # Multi-sheet parsing
    "MultiSheetParser",
    "MultiSheetOptions",
    "SheetInfo",
    "read_all_sheets",
    "analyze_excel",
    # Utilities
    "sanitize_column_name",
]


# ============================================================================
# Convenience Functions
# ============================================================================

def read_excel(
    file_path: str, sheet: str | None = None, **config_kwargs: Any
) -> pd.DataFrame:
    """Quick function to read an Excel file to a pandas DataFrame."""
    config = SheetConfig(**config_kwargs) if config_kwargs else None
    wb = MessyWorkbook(file_path, sheet_config=config)
    return wb.to_dataframe(sheet=sheet)


def read_excel_tables(
    file_path: str, sheet: str | None = None
) -> list[pd.DataFrame]:
    """Read all detected tables from a sheet."""
    wb = MessyWorkbook(file_path)
    sheet_obj = wb.get_sheet(sheet or wb.sheet_names[0])
    return [table.to_dataframe() for table in sheet_obj.tables]


def analyze_structure(file_path: str, sheet: str | None = None) -> StructureInfo:
    """Analyze and return structure info without loading data."""
    wb = MessyWorkbook(file_path)
    return wb.get_structure(sheet or wb.sheet_names[0])
