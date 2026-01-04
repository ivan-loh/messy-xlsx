"""XLS file handler for legacy Excel formats."""

# ============================================================================
# Imports
# ============================================================================

from pathlib import Path

import pandas as pd

from messy_xlsx.exceptions import FileError, FormatError
from messy_xlsx.parsing.base_handler import FormatHandler, ParseOptions


# ============================================================================
# Core
# ============================================================================

class XLSHandler(FormatHandler):
    """Handler for legacy XLS files (Excel 97-2003)."""

    def can_handle(self, format_type: str) -> bool:
        """Check if this handler can process the format."""
        return format_type == "xls"

    def parse(
        self,
        file_path: Path,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Parse XLS file to DataFrame."""
        header = 0 if options.header_rows == 1 else None

        try:
            df = pd.read_excel(
                file_path,
                sheet_name   = sheet if sheet else 0,
                skiprows     = options.skip_rows if options.header_rows <= 1 else 0,
                skipfooter   = options.skip_footer,
                na_values    = options.na_values,
                header       = header,
                engine       = "xlrd",
            )
        except ImportError:
            try:
                df = pd.read_excel(
                    file_path,
                    sheet_name  = sheet if sheet else 0,
                    skiprows    = options.skip_rows if options.header_rows <= 1 else 0,
                    skipfooter  = options.skip_footer,
                    na_values   = options.na_values,
                    header      = header,
                )
            except Exception as e:
                raise FormatError(
                    f"Cannot parse XLS file (xlrd may be required): {e}",
                    file_path        = str(file_path),
                    detected_format  = "xls",
                ) from e
        except PermissionError as e:
            raise FileError(
                f"Permission denied: {file_path}",
                file_path  = str(file_path),
                operation  = "open",
            ) from e
        except Exception as e:
            raise FormatError(
                f"Cannot parse XLS file: {e}",
                file_path        = str(file_path),
                detected_format  = "xls",
            ) from e

        if options.header_rows > 1:
            if options.skip_rows > 0:
                df = df.iloc[options.skip_rows:]

            df, columns = self._generate_column_names(df, options.header_rows)
            df.columns  = columns
            df          = df.reset_index(drop=True)
        elif options.header_rows == 0:
            df.columns = [f"col_{i}" for i in range(len(df.columns))]

        if options.max_rows is not None:
            df = df.iloc[:options.max_rows]

        return df

    def get_sheet_names(self, file_path: Path) -> list[str]:
        """Get list of sheet names."""
        try:
            xl_file = pd.ExcelFile(file_path, engine="xlrd")
            sheets  = xl_file.sheet_names
            xl_file.close()
            return sheets
        except ImportError:
            try:
                xl_file = pd.ExcelFile(file_path)
                sheets  = xl_file.sheet_names
                xl_file.close()
                return sheets
            except Exception:
                return ["Sheet1"]
        except PermissionError as e:
            raise FileError(
                f"Permission denied: {file_path}",
                file_path  = str(file_path),
                operation  = "get_sheets",
            ) from e
        except Exception as e:
            return ["Sheet1"]

    def validate(self, file_path: Path) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        try:
            xl_file = pd.ExcelFile(file_path, engine="xlrd")
            xl_file.close()
            return True, None
        except ImportError:
            try:
                xl_file = pd.ExcelFile(file_path)
                xl_file.close()
                return True, None
            except Exception as e:
                return False, str(e)
        except Exception as e:
            return False, str(e)
