"""Multi-sheet Excel parser with auto-detection and cleaning."""

# ============================================================================
# Imports
# ============================================================================

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import pandas as pd

from messy_xlsx.normalization import MissingValueHandler, NumberNormalizer
from messy_xlsx.parsing import XLSXHandler, XLSHandler, ParseOptions
from messy_xlsx.utils import sanitize_column_name


# ============================================================================
# Config
# ============================================================================

# Patterns that indicate a metadata row (not actual data)
METADATA_PATTERNS = [
    re.compile(r"printed\s*(date|on|by)", re.I),
    re.compile(r"page\s*[-:]?\s*\d+\s*(of|/)\s*\d+", re.I),
    re.compile(r"generated\s*(on|by|at)", re.I),
    re.compile(r"report\s*(date|name|title)", re.I),
    re.compile(r"^(as\s+of|date|run\s+date)\s*:", re.I),
]

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

        # Select handler based on extension
        ext = self.file_path.suffix.lower()
        if ext in (".xlsx", ".xlsm"):
            self._handler = XLSXHandler()
        elif ext == ".xls":
            self._handler = XLSHandler()
        else:
            raise ValueError(f"Unsupported file type: {ext}")

    def analyze_sheets(self) -> list[SheetInfo]:
        """
        Analyze all sheets and return their structure info.

        Returns:
            List of SheetInfo objects describing each sheet
        """
        sheet_names = self._handler.get_sheet_names(self.file_path)
        results = []

        for name in sheet_names:
            info = self._analyze_sheet(name)
            results.append(info)

        return results

    def parse_all(self) -> dict[str, pd.DataFrame]:
        """
        Parse all data sheets and return cleaned DataFrames.

        Returns:
            Dict mapping sheet name to cleaned DataFrame
        """
        sheet_infos = self.analyze_sheets()
        results = {}

        for info in sheet_infos:
            # Skip based on analysis
            if info.skip_reason:
                continue

            # Apply custom filter if provided
            if self.options.sheet_filter and not self.options.sheet_filter(info):
                continue

            # Only parse sheets in explicit list if provided
            if self.options.sheets and info.name not in self.options.sheets:
                continue

            df = self._parse_sheet(info)
            if df is not None and not df.empty:
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
        info = self._analyze_sheet(sheet_name)
        return self._parse_sheet(info)

    def _analyze_sheet(self, sheet_name: str) -> SheetInfo:
        """Analyze a single sheet's structure."""
        # Read first few rows to analyze
        parse_options = ParseOptions(
            header_rows=0,  # Read raw data
            skip_rows=0,
        )

        try:
            df = self._handler.parse(self.file_path, sheet_name, parse_options)
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

        # Detect header row
        header_row = self._detect_header_row(df)

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

    def _detect_header_row(self, df: pd.DataFrame) -> int:
        """
        Detect which row contains the actual headers.

        Skips metadata rows like "Printed Date: ...", "Page 1 of 5", etc.

        Returns:
            0-indexed row number where headers are
        """
        max_scan = min(self.options.header_scan_rows, len(df))

        for row_idx in range(max_scan):
            row_values = df.iloc[row_idx].dropna().astype(str).tolist()

            # Skip completely empty rows
            if not row_values:
                continue

            # Check if this row looks like metadata
            is_metadata = False
            for val in row_values:
                for pattern in METADATA_PATTERNS:
                    if pattern.search(val):
                        is_metadata = True
                        break
                if is_metadata:
                    break

            if is_metadata:
                continue

            # Check if this row looks like headers
            # Headers typically have multiple non-empty string values
            non_empty_count = len(row_values)
            if non_empty_count >= 2:
                # Check if values look like headers (short strings, no dates)
                looks_like_header = True
                for val in row_values:
                    # If value is a datetime string, probably not headers
                    if re.match(r"^\d{4}-\d{2}-\d{2}", val):
                        looks_like_header = False
                        break
                    # If value is a pure number, probably not headers
                    if re.match(r"^[\d,.]+$", val) and len(val) > 0:
                        looks_like_header = False
                        break

                if looks_like_header:
                    return row_idx

        # Default to first row
        return 0

    def _parse_sheet(self, info: SheetInfo) -> pd.DataFrame:
        """Parse and clean a single sheet."""
        parse_options = ParseOptions(
            skip_rows=info.header_row,
            header_rows=1,
        )

        df = self._handler.parse(self.file_path, info.name, parse_options)

        if df.empty:
            return df

        # Clean column names
        if self.options.clean_column_names:
            df = self._clean_column_names(df)

        # Ensure type consistency
        if self.options.ensure_type_consistency:
            df = self._ensure_type_consistency(df)

        # Handle missing values
        handler = MissingValueHandler(
            use_extended_list=self.options.use_extended_missing_list,
            preserve_types=self.options.preserve_types,
        )
        df = handler.normalize(df, drop_empty_rows=True, drop_empty_cols=True)

        # Normalize numbers
        if self.options.normalize_numbers:
            normalizer = NumberNormalizer(
                decimal_separator=self.options.decimal_separator,
                thousands_separator=self.options.thousands_separator,
            )
            df = normalizer.normalize(df)

        return df

    def _clean_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean column names for database compatibility."""
        new_columns = []
        seen = {}

        for col in df.columns:
            clean = sanitize_column_name(col)

            # Handle duplicates
            if clean in seen:
                seen[clean] += 1
                clean = f"{clean}_{seen[clean]}"
            else:
                seen[clean] = 0

            new_columns.append(clean)

        df.columns = new_columns
        return df

    def _ensure_type_consistency(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Ensure each column has consistent types.

        If a column has mixed types (e.g., strings and numbers),
        convert all values to strings to prevent PyArrow errors.
        """
        for col in df.columns:
            if df[col].dtype == object:
                # Check for mixed numeric/string types
                types = set()
                for val in df[col].dropna():
                    if isinstance(val, (int, float)):
                        types.add("numeric")
                    elif isinstance(val, str):
                        types.add("string")
                    else:
                        types.add(type(val).__name__)

                # If mixed, convert all to string
                if len(types) > 1:
                    df[col] = df[col].apply(
                        lambda x: str(x) if pd.notna(x) and not isinstance(x, str) else x
                    )

        return df


# ============================================================================
# Convenience Functions
# ============================================================================

def read_all_sheets(
    file_path: str | Path,
    **options_kwargs,
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
    options = MultiSheetOptions(**options_kwargs)
    parser = MultiSheetParser(file_path, options)
    return parser.parse_all()


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
