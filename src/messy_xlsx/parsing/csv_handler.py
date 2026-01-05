"""CSV/TSV file handler with intelligent dialect detection."""

# ============================================================================
# Imports
# ============================================================================

import csv
import io
from pathlib import Path
from typing import BinaryIO

import pandas as pd

from messy_xlsx.exceptions import FileError, FormatError
from messy_xlsx.parsing.base_handler import FileSource, FormatHandler, ParseOptions


# ============================================================================
# Config
# ============================================================================

DEFAULT_NA_VALUES = ["", "NA", "N/A", "n/a", "null", "NULL", "None", "#N/A"]

ENCODING_FALLBACKS = ["latin-1", "windows-1252", "iso-8859-1"]


# ============================================================================
# Core
# ============================================================================

class CSVHandler(FormatHandler):
    """Handler for CSV and TSV files."""

    def can_handle(self, format_type: str) -> bool:
        """Check if this handler can process the format."""
        return format_type in ("csv", "tsv", "txt")

    def parse(
        self,
        file_source: FileSource,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Parse CSV/TSV file to DataFrame."""
        is_fileobj = hasattr(file_source, "read")
        file_desc = "<stream>" if is_fileobj else str(file_source)

        if is_fileobj:
            # For file-like objects, read content and detect from bytes
            if hasattr(file_source, "seek"):
                file_source.seek(0)
            raw_data = file_source.read()
            if hasattr(file_source, "seek"):
                file_source.seek(0)

            encoding = self._detect_encoding_from_bytes(raw_data, options.encoding)
            delimiter = self._detect_delimiter_from_bytes(raw_data, encoding)

            # Create a new StringIO for pandas
            text_data = raw_data.decode(encoding, errors="ignore")
            source_for_pandas = io.StringIO(text_data)
        else:
            encoding = self._detect_encoding(file_source, options.encoding)
            delimiter = self._detect_delimiter(file_source, encoding)
            source_for_pandas = file_source

        na_values = options.na_values or DEFAULT_NA_VALUES

        header = 0 if options.header_rows > 0 else None

        engine = "python" if options.skip_footer > 0 else "c"

        try:
            df = pd.read_csv(
                source_for_pandas,
                encoding      = None if is_fileobj else encoding,  # Already decoded for StringIO
                delimiter     = delimiter,
                skiprows      = options.skip_rows if options.header_rows <= 1 else 0,
                skipfooter    = options.skip_footer,
                nrows         = options.max_rows,
                na_values     = na_values,
                header        = header,
                engine        = engine,
                on_bad_lines  = "warn",  # Handle malformed rows gracefully
            )
        except UnicodeDecodeError:
            if is_fileobj:
                raise FormatError(
                    f"Cannot decode CSV data",
                    file_path        = file_desc,
                    detected_format  = "csv",
                )
            df = self._read_with_encoding_fallback(
                file_source,
                delimiter,
                options,
                na_values,
            )
        except Exception as e:
            raise FormatError(
                f"Cannot parse CSV file: {e}",
                file_path        = file_desc,
                detected_format  = "csv",
            ) from e

        if options.header_rows > 1:
            if options.skip_rows > 0:
                df = df.iloc[options.skip_rows:]

            df, columns = self._generate_column_names(df, options.header_rows)
            df.columns  = columns
            df          = df.reset_index(drop=True)
        elif options.header_rows == 0:
            df.columns = [f"col_{i}" for i in range(len(df.columns))]

        return df

    def _detect_encoding(self, file_path: Path, default: str) -> str:
        """Detect file encoding."""
        try:
            with open(file_path, "rb") as f:
                raw_data = f.read(10000)
        except Exception:
            return default

        return self._detect_encoding_from_bytes(raw_data, default)

    def _detect_encoding_from_bytes(self, raw_data: bytes, default: str) -> str:
        """Detect encoding from raw bytes."""
        if not raw_data:
            return default

        if raw_data.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        if raw_data.startswith(b"\xff\xfe"):
            return "utf-16-le"
        if raw_data.startswith(b"\xfe\xff"):
            return "utf-16-be"

        try:
            raw_data[:10000].decode("utf-8")
            return "utf-8"
        except UnicodeDecodeError:
            pass

        return "latin-1"

    def _detect_delimiter(self, file_path: Path, encoding: str) -> str:
        """Detect CSV delimiter."""
        try:
            with open(file_path, "r", encoding=encoding, errors="ignore") as f:
                sample = f.read(8192)
        except Exception:
            return ","

        return self._detect_delimiter_from_text(sample)

    def _detect_delimiter_from_bytes(self, raw_data: bytes, encoding: str) -> str:
        """Detect delimiter from raw bytes."""
        try:
            sample = raw_data[:8192].decode(encoding, errors="ignore")
        except Exception:
            return ","

        return self._detect_delimiter_from_text(sample)

    def _detect_delimiter_from_text(self, sample: str) -> str:
        """Detect delimiter from text sample."""
        try:
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample)
            return dialect.delimiter
        except csv.Error:
            pass

        lines = sample.split("\n")[:10]
        lines = [line for line in lines if line.strip()]

        if not lines:
            return ","

        delimiters     = [",", "\t", ";", "|"]
        best_delimiter = ","
        best_score     = 0.0

        for delim in delimiters:
            counts = [line.count(delim) for line in lines]

            if not counts or counts[0] == 0:
                continue

            avg_count = sum(counts) / len(counts)
            if len(counts) > 1:
                variance = sum((c - avg_count) ** 2 for c in counts) / len(counts)
            else:
                variance = 0

            score = avg_count / (variance + 1)

            if score > best_score:
                best_score     = score
                best_delimiter = delim

        return best_delimiter

    def _read_with_encoding_fallback(
        self,
        file_path: Path,
        delimiter: str,
        options: ParseOptions,
        na_values: list[str],
    ) -> pd.DataFrame:
        """Try reading with fallback encodings."""
        header = 0 if options.header_rows > 0 else None
        engine = "python" if options.skip_footer > 0 else "c"
        errors = []

        for encoding in ENCODING_FALLBACKS:
            try:
                df = pd.read_csv(
                    file_path,
                    encoding   = encoding,
                    delimiter  = delimiter,
                    skiprows   = options.skip_rows if options.header_rows <= 1 else 0,
                    skipfooter = options.skip_footer,
                    nrows      = options.max_rows,
                    na_values  = na_values,
                    header     = header,
                    engine     = engine,
                )
                return df
            except UnicodeDecodeError as e:
                errors.append(f"{encoding}: {e}")
                continue
            except Exception as e:
                errors.append(f"{encoding}: {e}")
                continue

        raise FormatError(
            f"Cannot read CSV with any encoding",
            file_path          = str(file_path),
            detected_format    = "csv",
            attempted_formats  = [f"csv[{enc}]" for enc in ENCODING_FALLBACKS],
        )

    def get_sheet_names(self, file_source: FileSource) -> list[str]:
        """Get sheet names (always returns single element for CSV)."""
        return ["Sheet1"]

    def validate(self, file_source: FileSource) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        is_fileobj = hasattr(file_source, "read")

        if is_fileobj:
            try:
                if hasattr(file_source, "seek"):
                    file_source.seek(0)
                data = file_source.read(1024)
                if hasattr(file_source, "seek"):
                    file_source.seek(0)
                return True, None
            except Exception as e:
                return False, str(e)
        else:
            try:
                encoding = self._detect_encoding(file_source, "utf-8")
                with open(file_source, "r", encoding=encoding, errors="ignore") as f:
                    f.read(1024)
                return True, None
            except Exception as e:
                return False, str(e)
