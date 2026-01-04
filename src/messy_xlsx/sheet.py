"""MessySheet class for accessing individual sheets."""

# ============================================================================
# Imports
# ============================================================================

from typing import TYPE_CHECKING, Any, Iterator

import pandas as pd

from messy_xlsx.models import CellValue, SheetConfig, StructureInfo, TableInfo
from messy_xlsx.utils import column_index_to_letter

if TYPE_CHECKING:
    from messy_xlsx.workbook import MessyWorkbook


# ============================================================================
# Helper Classes
# ============================================================================

class MessyTable:
    """Represents a detected table within a sheet."""

    def __init__(
        self,
        sheet: "MessySheet",
        table_info: TableInfo | dict,
    ):
        """Initialize table."""
        self._sheet = sheet

        if isinstance(table_info, dict):
            self._info = TableInfo(**table_info)
        else:
            self._info = table_info

    @property
    def start_row(self) -> int:
        """First row of table (1-indexed)."""
        return self._info.start_row

    @property
    def end_row(self) -> int:
        """Last row of table (1-indexed)."""
        return self._info.end_row

    @property
    def row_count(self) -> int:
        """Number of data rows."""
        return self._info.row_count

    @property
    def column_count(self) -> int:
        """Number of columns."""
        return self._info.column_count

    def to_dataframe(self, config: SheetConfig | None = None) -> pd.DataFrame:
        """Convert table to DataFrame."""
        range_str = self._info.to_range_string()

        config            = config or SheetConfig()
        config.cell_range = range_str

        return self._sheet._workbook._parse_sheet(self._sheet.name, config)


# ============================================================================
# Core
# ============================================================================

class MessySheet:
    """Wrapper for a single sheet with structure detection."""

    def __init__(
        self,
        workbook: "MessyWorkbook",
        name: str,
    ):
        """Initialize sheet."""
        self._workbook   = workbook
        self._name       = name
        self._structure: StructureInfo | None = None

    @property
    def name(self) -> str:
        """Sheet name."""
        return self._name

    @property
    def structure(self) -> StructureInfo:
        """Get detected structure for this sheet."""
        if self._structure is None:
            self._structure = self._workbook._analyze_structure(self._name)
        return self._structure

    @property
    def tables(self) -> list[MessyTable]:
        """Get detected tables in this sheet."""
        structure = self.structure
        return [
            MessyTable(self, info)
            for info in structure.table_ranges
        ]

    @property
    def has_multiple_tables(self) -> bool:
        """Check if sheet contains multiple tables."""
        return self.structure.num_tables > 1

    def to_dataframe(self, config: SheetConfig | None = None) -> pd.DataFrame:
        """Convert sheet to DataFrame."""
        return self._workbook._parse_sheet(self._name, config)

    def get_cell(self, row: int, col: int) -> CellValue:
        """Get a single cell value."""
        return self._workbook.get_cell(self._name, row, col)

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
    ) -> Iterator[list[CellValue]]:
        """Iterate over rows of cells."""
        structure = self.structure
        max_row   = max_row or structure.data_end_row
        max_col   = max_col or structure.data_end_col

        for row in range(min_row, max_row + 1):
            row_cells = []
            for col in range(min_col, max_col + 1):
                cell = self.get_cell(row, col)
                row_cells.append(cell)
            yield row_cells

    def __getitem__(self, ref: str) -> CellValue | list[list[CellValue]]:
        """Access cells by reference."""
        if ":" in ref:
            from messy_xlsx.utils import parse_range

            _, start_row, start_col, end_row, end_col = parse_range(ref)
            result = []
            for row in range(start_row, end_row + 1):
                row_cells = []
                for col in range(start_col, end_col + 1):
                    row_cells.append(self.get_cell(row, col))
                result.append(row_cells)
            return result
        else:
            from messy_xlsx.utils import cell_ref_to_coords

            _, row, col = cell_ref_to_coords(ref)
            return self.get_cell(row, col)

    def __repr__(self) -> str:
        return f"MessySheet({self._name!r})"
