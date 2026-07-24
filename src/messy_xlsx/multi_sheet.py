"""Multi-sheet Excel parser with auto-detection and cleaning."""

# ============================================================================
# Imports
# ============================================================================

import warnings as _warnings
from collections.abc import Callable
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast

import pandas as pd

from messy_xlsx.detection.header_detector import detect_header_row
from messy_xlsx.detection.structure_sampler import SampleWindow
from messy_xlsx.models import SheetConfig, SheetInfo
from messy_xlsx.parsing.parse_plan import ParsePlan, compile_parse_plan
from messy_xlsx.parsing.sheet_planner import (
    PlannedSheet,
    PlannedSheetState,
    PlanningFailureStage,
    SheetPlanner,
)
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
            if workbook._uses_builtin_ooxml_planner():
                planned = self._plan_shared_ooxml(
                    workbook,
                    compile_outputs=False,
                    select_all=True,
                )
                return [item.info for item in planned]
            if workbook._uses_builtin_xls_planner():
                planned = self._plan_shared_xls(
                    workbook,
                    compile_outputs=False,
                    select_all=True,
                )
                return [item.info for item in planned]
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
            if workbook._uses_builtin_ooxml_planner() or workbook._uses_builtin_xls_planner():
                lease = workbook._materialized_operation()
                with lease:
                    try:
                        lease._body_started()
                        planned = (
                            self._plan_shared_ooxml(
                                workbook,
                                compile_outputs=True,
                                select_all=False,
                            )
                            if workbook._uses_builtin_ooxml_planner()
                            else self._plan_shared_xls(
                                workbook,
                                compile_outputs=True,
                                select_all=False,
                            )
                        )
                        for item in planned:
                            if item.state is PlannedSheetState.SKIPPED:
                                continue
                            if item.state is PlannedSheetState.ERROR:
                                assert item.error is not None
                                if item.failure_stage is PlanningFailureStage.ANALYSIS:
                                    continue
                                raise item.error
                            assert item.parse_plan is not None
                            frame = workbook._materialize_compiled_plan(
                                item.name,
                                item.parse_plan,
                            )
                            if not frame.empty:
                                results[item.name] = frame
                        return results
                    finally:
                        lease._body_complete()

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

    def _plan_shared_ooxml(
        self,
        workbook: Any,
        *,
        compile_outputs: bool,
        select_all: bool,
    ) -> tuple[PlannedSheet, ...]:
        """Plan built-in OOXML sheets with bounded evidence and no raw parse."""
        structures: dict[str, Any] = {}

        def analyze(name: str) -> SheetInfo:
            return self._analyze_sheet_bounded(workbook, name, structures)

        def compile_selected(name: str, info: SheetInfo) -> ParsePlan:
            return compile_parse_plan(
                self._config_for_info(info),
                structures.get(name),
                workbook.format_type,
            )

        planner = SheetPlanner(
            analyze,
            compile_selected,
            should_propagate=lambda error: workbook._should_propagate_sheet_error(error),
            analysis_failure_info=lambda name, error: SheetInfo(
                name=name,
                row_count=0,
                col_count=0,
                header_row=0,
                is_empty=True,
                skip_reason=f"Parse error: {error}",
            ),
        )
        return planner.plan(
            workbook.sheet_names,
            options=None if select_all else self.options,
            select_all=select_all,
            compile_outputs=compile_outputs,
        )

    def _plan_shared_xls(
        self,
        workbook: Any,
        *,
        compile_outputs: bool,
        select_all: bool,
    ) -> tuple[PlannedSheet, ...]:
        """Plan untouched built-in XLS sheets from one bounded xlrd inspection."""
        structures: dict[str, Any] = {}
        with (
            workbook._source_handle.open_backend() as backend_source,
            closing(pd.ExcelFile(backend_source, engine="xlrd")) as excel,
        ):

            def analyze(name: str) -> SheetInfo:
                return self._analyze_sheet_bounded_xls(excel, name)

            def compile_selected(name: str, info: SheetInfo) -> ParsePlan:
                return compile_parse_plan(
                    self._config_for_info(info),
                    structures.get(name),
                    workbook.format_type,
                )

            planner = SheetPlanner(
                analyze,
                compile_selected,
                should_propagate=lambda error: workbook._should_propagate_sheet_error(error),
                analysis_failure_info=lambda name, error: SheetInfo(
                    name=name,
                    row_count=0,
                    col_count=0,
                    header_row=0,
                    is_empty=True,
                    skip_reason=f"Parse error: {error}",
                ),
            )
            return planner.plan(
                workbook.sheet_names,
                options=None if select_all else self.options,
                select_all=select_all,
                compile_outputs=compile_outputs,
            )

    def _analyze_sheet_bounded(
        self,
        workbook: Any,
        sheet_name: str,
        structures: dict[str, Any],
    ) -> SheetInfo:
        """Analyze built-in OOXML using manifest dimensions and bounded values."""
        manifest = workbook._get_sheet_manifest(sheet_name)
        start_row, end_row, start_col, end_col = manifest.semantic_data_region
        has_values = manifest.semantic_nonempty_rows.contains(start_row)
        row_count = end_row - start_row + 1 if has_values else 0
        col_count = end_col - start_col + 1 if has_values else 0

        if row_count == 0 or col_count == 0:
            return SheetInfo(
                name=sheet_name,
                row_count=0,
                col_count=0,
                header_row=0,
                is_empty=True,
                skip_reason="Empty sheet" if self.options.skip_empty else None,
            )

        if row_count < self.options.min_rows or col_count < self.options.min_cols:
            return SheetInfo(
                name=sheet_name,
                row_count=row_count,
                col_count=col_count,
                header_row=0,
                is_empty=True,
                skip_reason="Too small" if self.options.skip_empty else None,
            )

        pivot_sample = (
            workbook._get_fastexcel_session()
            .sample_windows(
                sheet_name,
                windows=(SampleWindow(start_row, min(20, row_count)),),
                min_column=start_col,
                max_column=end_col,
            )
            .values
        )
        workbook.parse_metrics.sample_reads += 1
        is_pivot = self._looks_like_pivot(pivot_sample)
        if is_pivot and self.options.skip_pivots:
            return SheetInfo(
                name=sheet_name,
                row_count=row_count,
                col_count=col_count,
                header_row=0,
                is_pivot=True,
                skip_reason="Pivot table",
            )

        structure = workbook._analyze_stream_structure(sheet_name, SheetConfig())
        structures[sheet_name] = structure
        header_row = max(0, (structure.header_row or 1) - 1)
        if header_row >= self.options.header_scan_rows:
            header_row = 0
        return SheetInfo(
            name=sheet_name,
            row_count=row_count,
            col_count=col_count,
            header_row=header_row,
            is_pivot=is_pivot,
        )

    def _analyze_sheet_bounded_xls(
        self,
        excel: pd.ExcelFile,
        sheet_name: str,
    ) -> SheetInfo:
        """Analyze one XLS sheet from metadata plus a bounded value preview."""
        worksheet = excel.book.sheet_by_name(sheet_name)
        row_count = int(worksheet.nrows)
        col_count = int(worksheet.ncols)
        if row_count == 0 or col_count == 0:
            return SheetInfo(
                name=sheet_name,
                row_count=0,
                col_count=0,
                header_row=0,
                is_empty=True,
                skip_reason="Empty sheet" if self.options.skip_empty else None,
            )
        if row_count < self.options.min_rows or col_count < self.options.min_cols:
            return SheetInfo(
                name=sheet_name,
                row_count=row_count,
                col_count=col_count,
                header_row=0,
                is_empty=True,
                skip_reason="Too small" if self.options.skip_empty else None,
            )

        preview_rows = max(20, self.options.header_scan_rows)
        preview = excel.parse(
            sheet_name=sheet_name,
            header=None,
            nrows=preview_rows,
            na_values=[],
        )
        is_pivot = self._looks_like_pivot(preview)
        if is_pivot and self.options.skip_pivots:
            return SheetInfo(
                name=sheet_name,
                row_count=row_count,
                col_count=col_count,
                header_row=0,
                is_pivot=True,
                skip_reason="Pivot table",
            )
        header_row = detect_header_row(preview, self.options.header_scan_rows)
        if header_row >= self.options.header_scan_rows:
            header_row = 0
        return SheetInfo(
            name=sheet_name,
            row_count=row_count,
            col_count=col_count,
            header_row=header_row,
            is_pivot=is_pivot,
        )

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
            if workbook._uses_builtin_ooxml_planner() and sheet_name in workbook.sheet_names:
                lease = workbook._materialized_operation()
                with lease:
                    try:
                        lease._body_started()
                        structures: dict[str, Any] = {}
                        info = self._analyze_sheet_bounded(
                            workbook,
                            sheet_name,
                            structures,
                        )
                        plan = compile_parse_plan(
                            self._config_for_info(info),
                            structures.get(sheet_name),
                            workbook.format_type,
                        )
                        return workbook._materialize_compiled_plan(sheet_name, plan)
                    finally:
                        lease._body_complete()
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
        return workbook._to_dataframe_compat(
            info.name,
            config=self._config_for_info(info),
        )

    def _config_for_info(self, info: SheetInfo) -> SheetConfig:
        """Compile the legacy final-parse configuration for one analyzed sheet."""
        return SheetConfig(
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
        return cast(dict[str, pd.DataFrame], parser._parse_all_compat())
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
