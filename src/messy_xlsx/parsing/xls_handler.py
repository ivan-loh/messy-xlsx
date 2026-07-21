"""XLS file handler for legacy Excel formats."""

# ============================================================================
# Imports
# ============================================================================

from contextlib import closing

import pandas as pd

from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FileError, FormatError
from messy_xlsx.parsing.base_handler import (
    FileSource,
    FormatHandler,
    ParseOptions,
)

# ============================================================================
# Core
# ============================================================================


class XLSHandler(FormatHandler):
    """Handler for legacy XLS files (Excel 97-2003)."""

    _accepts_source_handle = True

    def can_handle(self, format_type: str) -> bool:
        """Check if this handler can process the format."""
        return format_type == "xls"

    def parse(
        self,
        file_source: FileSource | SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Parse XLS file to DataFrame."""
        source = SourceHandle.coerce(file_source)
        try:
            return self._parse_source(source, sheet, options)
        finally:
            if source is not file_source:
                source.close()

    def _parse_source(
        self,
        source: SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Parse one replayable source view, borrowing afresh for retries."""
        file_desc = source.description
        header = 0 if options.header_rows == 1 else None

        try:
            with source.open_backend() as backend_source:
                df = pd.read_excel(
                    backend_source,
                    sheet_name=sheet if sheet else 0,
                    skiprows=options.skip_rows if options.header_rows <= 1 else 0,
                    skipfooter=options.skip_footer,
                    na_values=options.na_values,
                    header=header,
                    engine="xlrd",
                )
        except ImportError:
            try:
                with source.open_backend() as backend_source:
                    df = pd.read_excel(
                        backend_source,
                        sheet_name=sheet if sheet else 0,
                        skiprows=options.skip_rows if options.header_rows <= 1 else 0,
                        skipfooter=options.skip_footer,
                        na_values=options.na_values,
                        header=header,
                    )
            except Exception as e:
                raise FormatError(
                    f"Cannot parse XLS file (xlrd may be required): {e}",
                    file_path=file_desc,
                    detected_format="xls",
                ) from e
        except PermissionError as e:
            raise FileError(
                f"Permission denied: {file_desc}",
                file_path=file_desc,
                operation="open",
            ) from e
        except Exception as e:
            raise FormatError(
                f"Cannot parse XLS file: {e}",
                file_path=file_desc,
                detected_format="xls",
            ) from e

        if options.header_rows > 1:
            if options.skip_rows > 0:
                df = df.iloc[options.skip_rows :]

            df, columns = self._generate_column_names(df, options.header_rows)
            df.columns = columns
            df = df.reset_index(drop=True)
        elif options.header_rows == 0:
            df.columns = [f"col_{i}" for i in range(len(df.columns))]

        if options.max_rows is not None:
            df = df.iloc[: options.max_rows]

        return df

    def get_sheet_names(self, file_source: FileSource | SourceHandle) -> list[str]:
        """Get list of sheet names."""
        source = SourceHandle.coerce(file_source)
        try:
            file_desc = source.description
            try:
                with (
                    source.open_backend() as backend_source,
                    closing(pd.ExcelFile(backend_source, engine="xlrd")) as xl_file,
                ):
                    return list(xl_file.sheet_names)
            except ImportError:
                try:
                    with (
                        source.open_backend() as backend_source,
                        closing(pd.ExcelFile(backend_source)) as xl_file,
                    ):
                        return list(xl_file.sheet_names)
                except Exception:
                    return ["Sheet1"]
            except PermissionError as e:
                raise FileError(
                    f"Permission denied: {file_desc}",
                    file_path=file_desc,
                    operation="get_sheets",
                ) from e
            except Exception:
                return ["Sheet1"]
        finally:
            if source is not file_source:
                source.close()

    def validate(self, file_source: FileSource | SourceHandle) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        source = SourceHandle.coerce(file_source)
        try:
            try:
                with (
                    source.open_backend() as backend_source,
                    closing(pd.ExcelFile(backend_source, engine="xlrd")),
                ):
                    pass
                return True, None
            except ImportError:
                try:
                    with (
                        source.open_backend() as backend_source,
                        closing(pd.ExcelFile(backend_source)),
                    ):
                        pass
                    return True, None
                except Exception as e:
                    return False, str(e)
            except Exception as e:
                return False, str(e)
        finally:
            if source is not file_source:
                source.close()
