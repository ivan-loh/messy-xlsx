"""Multi-sheet Excel parser with auto-detection and cleaning."""

# ============================================================================
# Imports
# ============================================================================

import warnings as _warnings
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from messy_xlsx.detection.header_detector import detect_header_row
from messy_xlsx.models import SheetConfig
from messy_xlsx.warnings import LegacyAPIWarning, warn_legacy

# ============================================================================
# Config
# ============================================================================

# Patterns that indicate a pivot table (summary, not raw data)
PIVOT_INDICATORS = [
    "Row Labels",
    "Column Labels",
    "Grand Total",
    "Count of",
    "Sum of",
    "Average of",
]


# ============================================================================
# Models
# ============================================================================


@dataclass
class SheetInfo:
    """Information about a sheet's structure."""

    name: str
    row_count: int
    col_count: int
    header_row: int  # 0-indexed row where headers are
    is_empty: bool = False
    is_pivot: bool = False
    skip_reason: str | None = None

    @property
    def column_count(self) -> int:
        """Number of columns (descriptive alias for ``col_count``)."""
        return self.col_count


@dataclass
class MultiSheetOptions:
    """Options for multi-sheet parsing."""

    # Which sheets to parse (None = auto-detect data sheets)
    sheets: list[str] | None = None

    # Skip sheets that look like pivot tables
    skip_pivots: bool = True

    # Skip empty sheets
    skip_empty: bool = True

    # Minimum rows to consider a sheet as "data"
    min_rows: int = 2

    # Minimum columns to consider a sheet as "data"
    min_cols: int = 2

    # Maximum rows to scan for header detection
    header_scan_rows: int = 10

    # Clean column names for BigQuery/database compatibility
    clean_column_names: bool = True

    # Ensure type consistency (no mixed str/float columns)
    ensure_type_consistency: bool = True

    # Missing value handling options
    use_extended_missing_list: bool = False
    preserve_types: bool = True

    # Number normalization
    normalize_numbers: bool = True
    decimal_separator: str | None = None
    thousands_separator: str | None = None

    # Custom sheet filter function
    sheet_filter: Callable[[SheetInfo], bool] | None = None


# ============================================================================
# Core
# ============================================================================


class MultiSheetParser:
    """
    Parse multiple sheets from an Excel file with auto-detection.

    Handles:
    - Automatic header row detection (skips metadata rows)
    - Empty sheet filtering
    - Pivot table detection
    - Type consistency (no mixed str/float columns)
    - Column name cleaning for BigQuery compatibility
    """

    def __init__(self, file_path: str | Path, options: MultiSheetOptions | None = None):
        """
        Initialize parser.

        Args:
            file_path: Path to Excel file
            options: Parsing options
        """
        self.file_path = Path(file_path)
        self.options = options or MultiSheetOptions()

        # Multi-sheet parsing is intentionally limited to Excel workbooks.
        ext = self.file_path.suffix.lower()
        if ext not in (".xlsx", ".xlsm", ".xls"):
            raise ValueError(f"Unsupported file type: {ext}")

    def analyze_sheets(self) -> list[SheetInfo]:
        """
        Analyze all sheets and return their structure info.

        Returns:
            List of SheetInfo objects describing each sheet
        """
        from messy_xlsx.workbook import MessyWorkbook

        with MessyWorkbook(self.file_path) as workbook:
            return [self._analyze_sheet(workbook, name) for name in workbook.sheet_names]

    def parse_all(self) -> dict[str, pd.DataFrame]:
        """
        Parse all data sheets and return cleaned DataFrames.

        Returns:
            Dict mapping sheet name to cleaned DataFrame
        """
        warn_legacy("MultiSheetParser.parse_all")
        return self._parse_all_compat()

    def _parse_all_compat(self) -> dict[str, pd.DataFrame]:
        """Parse all data sheets without emitting a nested legacy warning."""
        from messy_xlsx.workbook import MessyWorkbook

        results = {}
        with MessyWorkbook(self.file_path) as workbook:
            sheet_infos = [self._analyze_sheet(workbook, name) for name in workbook.sheet_names]

            for info in sheet_infos:
                if info.skip_reason:
                    continue
                if self.options.sheet_filter and not self.options.sheet_filter(info):
                    continue
                if self.options.sheets and info.name not in self.options.sheets:
                    continue

                df = self._parse_sheet(workbook, info)
                if not df.empty:
                    results[info.name] = df

        return results

    def parse_sheet(self, sheet_name: str) -> pd.DataFrame:
        """
        Parse a specific sheet.

        Args:
            sheet_name: Name of sheet to parse

        Returns:
            Cleaned DataFrame
        """
        warn_legacy("MultiSheetParser.parse_sheet")
        return self._parse_sheet_compat(sheet_name)

    def _parse_sheet_compat(self, sheet_name: str) -> pd.DataFrame:
        """Parse one sheet without emitting a nested legacy warning."""
        from messy_xlsx.workbook import MessyWorkbook

        with MessyWorkbook(self.file_path) as workbook:
            info = self._analyze_sheet(workbook, sheet_name)
            return self._parse_sheet(workbook, info)

    def _analyze_sheet(self, workbook: Any, sheet_name: str) -> SheetInfo:
        """Analyze a single sheet's structure."""
        raw_config = SheetConfig(
            auto_detect=False,
            header_rows=0,
            include_hidden=True,
            merge_strategy="skip",
            normalize=False,
            sanitize_column_names=False,
        )

        try:
            df = workbook._to_dataframe_compat(sheet_name, config=raw_config)
        except Exception as e:
            return SheetInfo(
                name=sheet_name,
                row_count=0,
                col_count=0,
                header_row=0,
                is_empty=True,
                skip_reason=f"Parse error: {e}",
            )

        # Check if empty
        if df.empty or len(df) == 0 or len(df.columns) == 0:
            return SheetInfo(
                name=sheet_name,
                row_count=0,
                col_count=0,
                header_row=0,
                is_empty=True,
                skip_reason="Empty sheet" if self.options.skip_empty else None,
            )

        # Check minimum size
        if len(df) < self.options.min_rows or len(df.columns) < self.options.min_cols:
            return SheetInfo(
                name=sheet_name,
                row_count=len(df),
                col_count=len(df.columns),
                header_row=0,
                is_empty=True,
                skip_reason="Too small" if self.options.skip_empty else None,
            )

        # Check for pivot table
        is_pivot = self._looks_like_pivot(df)
        if is_pivot and self.options.skip_pivots:
            return SheetInfo(
                name=sheet_name,
                row_count=len(df),
                col_count=len(df.columns),
                header_row=0,
                is_pivot=True,
                skip_reason="Pivot table",
            )

        # Reuse the core analyzer so every public parsing path agrees about
        # header placement.
        if workbook.format_type == "xls":
            header_row = detect_header_row(df, self.options.header_scan_rows)
        else:
            structure = workbook.get_structure(sheet_name)
            header_row = max(0, (structure.header_row or 1) - 1)
        if header_row >= self.options.header_scan_rows:
            header_row = 0

        return SheetInfo(
            name=sheet_name,
            row_count=len(df),
            col_count=len(df.columns),
            header_row=header_row,
            is_pivot=is_pivot,
        )

    def _looks_like_pivot(self, df: pd.DataFrame) -> bool:
        """Check if sheet looks like a pivot table."""
        # Check first few rows for pivot indicators
        sample = df.head(20)

        for col in sample.columns:
            for val in sample[col].dropna().astype(str):
                for indicator in PIVOT_INDICATORS:
                    if indicator.lower() in val.lower():
                        return True

        return False

    def _parse_sheet(self, workbook: Any, info: SheetInfo) -> pd.DataFrame:
        """Parse and clean a single sheet."""
        config = SheetConfig(
            auto_detect=False,
            skip_rows=info.header_row,
            header_rows=1,
            include_hidden=True,
            sanitize_column_names=self.options.clean_column_names,
            normalize_numbers=self.options.normalize_numbers,
            use_extended_missing_list=self.options.use_extended_missing_list,
            preserve_types=self.options.preserve_types,
            ensure_type_consistency=self.options.ensure_type_consistency,
            decimal_separator=self.options.decimal_separator,
            thousands_separator=self.options.thousands_separator,
        )
        return workbook._to_dataframe_compat(info.name, config=config)


# ============================================================================
# Convenience Functions
# ============================================================================

_MULTI_SHEET_PARSER_PARSE_ALL = MultiSheetParser.parse_all


def _parse_all_compat(parser: MultiSheetParser) -> dict[str, pd.DataFrame]:
    parse_all = parser.parse_all
    if (
        getattr(parse_all, "__func__", None) is _MULTI_SHEET_PARSER_PARSE_ALL
        and getattr(parse_all, "__self__", None) is parser
    ):
        return parser._parse_all_compat()
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore", LegacyAPIWarning)
        return parse_all()


def read_all_sheets(
    file_path: str | Path,
    **options_kwargs: Any,
) -> dict[str, pd.DataFrame]:
    """
    Read all data sheets from an Excel file.

    Automatically:
    - Skips empty sheets and pivot tables
    - Detects header rows (skips metadata)
    - Cleans column names for BigQuery/databases
    - Ensures type consistency per column

    Args:
        file_path: Path to Excel file
        **options_kwargs: Options passed to MultiSheetOptions

    Returns:
        Dict mapping sheet name to cleaned DataFrame

    Example:
        >>> sheets = read_all_sheets("messy_file.xlsx")
        >>> for name, df in sheets.items():
        ...     print(f"{name}: {len(df)} rows")
    """
    warn_legacy("read_all_sheets")
    options = MultiSheetOptions(**options_kwargs)
    parser = MultiSheetParser(file_path, options)
    return _parse_all_compat(parser)


def analyze_excel(file_path: str | Path) -> list[SheetInfo]:
    """
    Analyze an Excel file's structure without loading data.

    Returns:
        List of SheetInfo describing each sheet

    Example:
        >>> infos = analyze_excel("messy_file.xlsx")
        >>> for info in infos:
        ...     print(f"{info.name}: {info.row_count} rows, header at row {info.header_row}")
    """
    parser = MultiSheetParser(file_path)
    return parser.analyze_sheets()
