"""XLSX/XLSM file handler using fastexcel with openpyxl fallback."""

# ============================================================================
# Imports
# ============================================================================

import logging
from contextlib import ExitStack, closing
from typing import Any

import fastexcel
import numpy as np
import openpyxl
import pandas as pd

from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FileError, FormatError
from messy_xlsx.parsing.base_handler import (
    FileSource,
    FormatHandler,
    ParseOptions,
)

logger = logging.getLogger(__name__)

# ============================================================================
# Constants
# ============================================================================

EXCEL_ERRORS = {
    "#DIV/0!",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#REF!",
    "#VALUE!",
    "#GETTING_DATA",
}


# ============================================================================
# Core
# ============================================================================


class XLSXHandler(FormatHandler):
    """Handler for XLSX and XLSM files."""

    _accepts_source_handle = True

    def can_handle(self, format_type: str) -> bool:
        return format_type in ("xlsx", "xlsm")

    def parse(
        self,
        file_source: FileSource | SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Parse XLSX/XLSM file to DataFrame."""
        source = SourceHandle.coerce(file_source)
        try:
            needs_openpyxl = (
                options.merge_strategy != "skip"
                or options.ignore_hidden
                or options.cell_range
                or not options.data_only
            )

            if needs_openpyxl:
                return self._parse_openpyxl(source, sheet, options)
            return self._parse_fastexcel(source, sheet, options)
        finally:
            if source is not file_source:
                source.close()

    # -------------------------------------------------------------------------
    # Fastexcel (fast path)
    # -------------------------------------------------------------------------

    def _parse_fastexcel(
        self,
        source: SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Fast path using fastexcel."""
        file_desc = source.description

        try:
            content = source.path if source.path is not None else source.read_bytes()
            excel_file = fastexcel.read_excel(content)
        except PermissionError as e:
            raise FileError(
                f"Permission denied: {file_desc}", file_path=file_desc, operation="open"
            ) from e
        except Exception as e:
            raise FormatError(f"Cannot open Excel file: {e}", file_path=file_desc) from e

        try:
            sheet_idx = self._resolve_sheet_index(excel_file, sheet, file_desc)

            df = excel_file.load_sheet(
                sheet_idx,
                skip_rows=options.skip_rows,
                header_row=None,
            ).to_pandas()

            if df.empty:
                return pd.DataFrame()

            df = self._apply_options(df, options)
            return self._clean_excel_data(df, options)

        except FormatError:
            raise
        except Exception as e:
            raise FormatError(f"Error reading Excel file: {e}", file_path=file_desc) from e

    def _resolve_sheet_index(self, excel_file: Any, sheet: str | None, file_desc: str) -> int:
        if not sheet:
            return 0
        if sheet not in excel_file.sheet_names:
            raise FormatError(
                f"Sheet '{sheet}' not found", file_path=file_desc, detected_format="xlsx"
            )
        return int(excel_file.sheet_names.index(sheet))

    # -------------------------------------------------------------------------
    # Openpyxl (fallback for advanced features)
    # -------------------------------------------------------------------------

    def _parse_openpyxl(
        self,
        source: SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Fallback path using openpyxl for advanced features."""
        read_only = options.merge_strategy == "skip" and not options.ignore_hidden
        file_desc = source.description

        with ExitStack() as stack:
            try:
                backend_source = stack.enter_context(source.open_backend())
                wb = openpyxl.load_workbook(
                    backend_source,
                    read_only=read_only,
                    data_only=options.data_only,
                    keep_links=False,
                )
            except PermissionError as e:
                raise FileError(
                    f"Permission denied: {file_desc}", file_path=file_desc, operation="open"
                ) from e
            except Exception as e:
                raise FormatError(f"Cannot open Excel file: {e}", file_path=file_desc) from e

            # Close the workbook before the borrowed source view is restored or
            # released. This matters for read-only openpyxl archives.
            stack.callback(wb.close)
            ws = self._resolve_worksheet(wb, sheet, file_desc)

            if options.merge_strategy != "skip" and not read_only:
                self._handle_merged_cells(ws, options.merge_strategy)

            data = self._read_worksheet(ws, options)
            if not data:
                return pd.DataFrame()

            df = pd.DataFrame(data)

            if options.skip_rows > 0 and not options.cell_range:
                df = df.iloc[options.skip_rows :]

            df = self._apply_options(df, options)
            return self._clean_excel_data(df, options)

    def _resolve_worksheet(self, wb: Any, sheet: str | None, file_desc: str) -> Any:
        if not sheet:
            return wb.active
        if sheet not in wb.sheetnames:
            raise FormatError(
                f"Sheet '{sheet}' not found", file_path=file_desc, detected_format="xlsx"
            )
        return wb[sheet]

    def _read_worksheet(self, ws: Any, options: ParseOptions) -> list[list[Any]]:
        """Read worksheet data using openpyxl."""
        if options.cell_range:
            return self._read_cell_range(ws, options.cell_range)
        return self._read_full_sheet(ws, options.ignore_hidden)

    def _read_cell_range(self, ws: Any, cell_range: str) -> list[list[Any]]:
        try:
            rows_iter = ws[cell_range]
            if not isinstance(rows_iter[0], tuple):
                rows_iter = [rows_iter]

            data = []
            for row in rows_iter:
                if not isinstance(row, tuple):
                    row = [row]
                data.append([cell.value for cell in row])
            return data

        except Exception as e:
            raise FormatError(f"Invalid cell range: {cell_range}", detected_format="xlsx") from e

    def _read_full_sheet(self, ws: Any, ignore_hidden: bool) -> list[list[Any]]:
        hidden_rows: set[int] = set()
        hidden_cols: set[str] = set()

        if ignore_hidden:
            hidden_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
            hidden_cols = {c for c, d in ws.column_dimensions.items() if d.hidden}

        data = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=False), start=1):
            if row_idx in hidden_rows:
                continue

            if hidden_cols:
                from openpyxl.utils import get_column_letter

                row_values = [
                    cell.value for cell in row if get_column_letter(cell.column) not in hidden_cols
                ]
            else:
                row_values = [cell.value for cell in row]
            data.append(row_values)

        return data

    def _handle_merged_cells(self, ws: Any, strategy: str) -> None:
        """Handle merged cells according to strategy."""
        for merged_range in list(ws.merged_cells.ranges):
            top_left = ws.cell(merged_range.min_row, merged_range.min_col).value
            ws.unmerge_cells(str(merged_range))

            if strategy == "fill":
                for r in range(merged_range.min_row, merged_range.max_row + 1):
                    for c in range(merged_range.min_col, merged_range.max_col + 1):
                        ws.cell(r, c).value = top_left

            elif strategy == "first_only":
                for r in range(merged_range.min_row, merged_range.max_row + 1):
                    for c in range(merged_range.min_col, merged_range.max_col + 1):
                        if r != merged_range.min_row or c != merged_range.min_col:
                            ws.cell(r, c).value = None

    # -------------------------------------------------------------------------
    # Common
    # -------------------------------------------------------------------------

    def _apply_options(self, df: pd.DataFrame, options: ParseOptions) -> pd.DataFrame:
        """Apply skip_footer and header options."""
        if options.skip_footer > 0:
            df = df.iloc[: -options.skip_footer]

        if options.header_rows > 0 and len(df) >= options.header_rows:
            df, columns = self._generate_column_names(df, options.header_rows)
            df.columns = columns
            df = df.reset_index(drop=True)
        else:
            df.columns = [f"col_{i}" for i in range(len(df.columns))]

        return df

    def _clean_excel_data(self, df: pd.DataFrame, options: ParseOptions) -> pd.DataFrame:
        """Replace Excel errors and custom NA values with NaN."""
        na_values = EXCEL_ERRORS | set(options.na_values or [])

        for idx in range(len(df.columns)):
            series = df.iloc[:, idx]
            if series.dtype == object:
                df.iloc[:, idx] = series.map(lambda x: np.nan if x in na_values else x)

        return df

    # -------------------------------------------------------------------------
    # Public API
    # -------------------------------------------------------------------------

    def get_sheet_names(self, file_source: FileSource | SourceHandle) -> list[str]:
        """Get list of sheet names."""
        source = SourceHandle.coerce(file_source)
        try:
            file_desc = source.description
            try:
                content = source.path if source.path is not None else source.read_bytes()
                excel_file = fastexcel.read_excel(content)
                return list(excel_file.sheet_names)
            except (OSError, ValueError, RuntimeError) as e:
                logger.debug("fastexcel sheet name read failed, falling back to openpyxl: %s", e)

            # Keep the openpyxl workbook inside the borrowed source lifetime.
            try:
                with (
                    source.open_backend() as backend_source,
                    closing(openpyxl.load_workbook(backend_source, read_only=True)) as wb,
                ):
                    return list(wb.sheetnames)
            except PermissionError as e:
                raise FileError(
                    f"Permission denied: {file_desc}",
                    file_path=file_desc,
                    operation="get_sheets",
                ) from e
            except Exception as e:
                raise FormatError(f"Cannot read sheet names: {e}", file_path=file_desc) from e
        finally:
            if source is not file_source:
                source.close()

    def validate(self, file_source: FileSource | SourceHandle) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        source = SourceHandle.coerce(file_source)
        try:
            content = source.path if source.path is not None else source.read_bytes()
            fastexcel.read_excel(content)
            return True, None
        except Exception as e:
            return False, str(e)
        finally:
            if source is not file_source:
                source.close()
