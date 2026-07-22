"""messy-xlsx: A Python library for parsing messy Excel files."""

# ============================================================================
# Imports
# ============================================================================

import warnings as _warnings
from typing import Any

import pandas as pd

from messy_xlsx.enums import (
    DataType,
    FormatType,
    HeaderDetectionMode,
    HeaderFallback,
    MergeStrategy,
)
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
from messy_xlsx.sheet import MessyTable as _MessyTable
from messy_xlsx.utils import sanitize_column_name
from messy_xlsx.warnings import LegacyAPIWarning
from messy_xlsx.warnings import warn_legacy as _warn_legacy
from messy_xlsx.workbook import MessyWorkbook

_MESSY_TABLE_TO_DATAFRAME = _MessyTable.to_dataframe
_MESSY_WORKBOOK_TO_DATAFRAME = MessyWorkbook.to_dataframe

# ============================================================================
# Package Metadata
# ============================================================================

__version__ = "0.10.0"

__all__ = [
    "CellValue",
    "CircularRefStrategy",
    "CircularReferenceError",
    "DataType",
    "FileError",
    "FormatError",
    "FormatInfo",
    "FormatType",
    "FormulaConfig",
    "FormulaError",
    "FormulaEvaluationMode",
    "HeaderDetectionMode",
    "HeaderFallback",
    "LegacyAPIWarning",
    "MergeStrategy",
    "MessyWorkbook",
    "MessyXlsxError",
    "MultiSheetOptions",
    "MultiSheetParser",
    "NormalizationError",
    "SheetConfig",
    "SheetError",
    "SheetInfo",
    "StructureError",
    "StructureInfo",
    "UnsupportedFunctionError",
    "analyze_excel",
    "analyze_structure",
    "read_all_sheets",
    "read_excel",
    "read_excel_tables",
    "sanitize_column_name",
]


# ============================================================================
# Convenience Functions
# ============================================================================


def _workbook_to_dataframe_compat(
    workbook: MessyWorkbook,
    sheet: str | None,
) -> pd.DataFrame:
    if getattr(type(workbook), "to_dataframe", None) is _MESSY_WORKBOOK_TO_DATAFRAME:
        return workbook._to_dataframe_compat(sheet=sheet)
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore", LegacyAPIWarning)
        return workbook.to_dataframe(sheet=sheet)


def _table_to_dataframe_compat(table: _MessyTable) -> pd.DataFrame:
    if getattr(type(table), "to_dataframe", None) is _MESSY_TABLE_TO_DATAFRAME:
        return table._to_dataframe_compat()
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore", LegacyAPIWarning)
        return table.to_dataframe()


def read_excel(file_path: str, sheet: str | None = None, **config_kwargs: Any) -> pd.DataFrame:
    """Quick function to read an Excel file to a pandas DataFrame."""
    _warn_legacy("read_excel")
    config = SheetConfig(**config_kwargs) if config_kwargs else None
    with MessyWorkbook(file_path, sheet_config=config) as wb:
        return _workbook_to_dataframe_compat(wb, sheet)


def read_excel_tables(file_path: str, sheet: str | None = None) -> list[pd.DataFrame]:
    """Read all detected tables from a sheet."""
    _warn_legacy("read_excel_tables")
    with MessyWorkbook(file_path) as wb:
        sheet_obj = wb.get_sheet(sheet or wb.sheet_names[0])
        return [_table_to_dataframe_compat(table) for table in sheet_obj.tables]


def analyze_structure(file_path: str, sheet: str | None = None) -> StructureInfo:
    """Analyze and return structure info without loading data."""
    with MessyWorkbook(file_path) as wb:
        return wb.get_structure(sheet or wb.sheet_names[0])
