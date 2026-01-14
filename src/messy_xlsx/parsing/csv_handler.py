"""CSV/TSV file handler with intelligent dialect detection."""

# ============================================================================
# Imports
# ============================================================================

import csv
import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

import numpy as np
import pandas as pd

from messy_xlsx.exceptions import FileError, FormatError
from messy_xlsx.parsing.base_handler import FileSource, FormatHandler, ParseOptions


# ============================================================================
# Config
# ============================================================================

DEFAULT_NA_VALUES = ["", "NA", "N/A", "n/a", "null", "NULL", "None", "#N/A"]

ENCODING_FALLBACKS = ["latin-1", "windows-1252", "iso-8859-1"]

METADATA_PATTERNS = [
    r"(?i)printed\s*(date)?",
    r"(?i)report\s*(date|name|:)",
    r"(?i)generated\s*(on|at|:)?",
    r"(?i)exported",
    r"(?i)^date\s*:",
    r"(?i)^\w+\s*:\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",
]


# ============================================================================
# Header Detection
# ============================================================================

@dataclass
class _RowProfile:
    """Profile of a row's characteristics for metadata detection."""

    empty_ratio: float
    numeric_ratio: float
    avg_cell_length: float
    has_metadata_pattern: bool
    cell_count: int
    is_empty: bool


class MetadataRowDetector:
    """
    Detector for metadata/header rows in CSV files.

    Uses statistical comparison to identify rows that don't match
    the structure of typical data rows (e.g., report headers, dates).
    """

    def _profile_row(self, row: pd.Series) -> _RowProfile:
        """Extract profile for a single row."""
        cells = [v for v in row if pd.notna(v)]
        non_empty = len([c for c in cells if str(c).strip()])
        total = len(row)

        if non_empty == 0:
            return _RowProfile(1.0, 0, 0, False, 0, True)

        numeric = sum(1 for c in cells if self._is_numeric(c))
        lengths = [len(str(c)) for c in cells if str(c).strip()]
        avg_len = float(np.mean(lengths)) if lengths else 0.0

        row_text = " ".join(str(v) for v in row if pd.notna(v))
        has_pattern = any(re.search(p, row_text) for p in METADATA_PATTERNS)

        return _RowProfile(
            empty_ratio=1 - (non_empty / total) if total > 0 else 1,
            numeric_ratio=numeric / non_empty,
            avg_cell_length=avg_len,
            has_metadata_pattern=has_pattern,
            cell_count=non_empty,
            is_empty=False,
        )

    def _is_numeric(self, val: object) -> bool:
        """Check if value is numeric."""
        try:
            float(str(val).replace(",", "").replace("%", ""))
            return True
        except (ValueError, TypeError):
            return False

    def _score_as_metadata(
        self, profile: _RowProfile, consensus: _RowProfile
    ) -> float:
        """Score how likely a row is metadata (0-1, higher = more likely)."""
        if profile.is_empty:
            return 1.0

        score = 0.0

        if profile.empty_ratio > consensus.empty_ratio + 0.3:
            score += 0.35

        if consensus.cell_count > 0:
            if profile.cell_count / consensus.cell_count < 0.4:
                score += 0.35

        if profile.has_metadata_pattern:
            score += 0.4

        return min(score, 1.0)

    def detect_skip_rows(
        self,
        file_path: Path,
        encoding: str,
        delimiter: str,
        max_check: int = 15,
    ) -> int:
        """
        Detect how many metadata rows to skip at the start of a CSV.

        Args:
            file_path: Path to the CSV file
            encoding: File encoding
            delimiter: CSV delimiter
            max_check: Maximum rows to analyze

        Returns:
            Number of rows to skip (0 if no metadata detected)
        """
        try:
            df = pd.read_csv(
                file_path,
                header=None,
                nrows=max_check,
                encoding=encoding,
                delimiter=delimiter,
                on_bad_lines="warn",
                skip_blank_lines=False,
            )
        except Exception:
            return 0

        if len(df) < 3:
            return 0

        profiles = [self._profile_row(df.iloc[i]) for i in range(len(df))]

        data_profiles = [p for p in profiles[2:10] if not p.is_empty]
        if not data_profiles:
            return 0

        consensus = _RowProfile(
            empty_ratio=float(np.mean([p.empty_ratio for p in data_profiles])),
            numeric_ratio=float(np.mean([p.numeric_ratio for p in data_profiles])),
            avg_cell_length=float(np.mean([p.avg_cell_length for p in data_profiles])),
            has_metadata_pattern=False,
            cell_count=int(np.mean([p.cell_count for p in data_profiles])),
            is_empty=False,
        )

        skip_rows = 0
        for i in range(min(6, len(profiles))):
            score = self._score_as_metadata(profiles[i], consensus)
            if score >= 0.35:
                skip_rows = i + 1
            else:
                break

        return skip_rows

    def detect_skip_rows_from_text(
        self,
        text_data: str,
        delimiter: str,
        max_check: int = 15,
    ) -> int:
        """Detect skip rows from text content (for file-like objects)."""
        try:
            df = pd.read_csv(
                io.StringIO(text_data),
                header=None,
                nrows=max_check,
                delimiter=delimiter,
                on_bad_lines="warn",
                skip_blank_lines=False,
            )
        except Exception:
            return 0

        if len(df) < 3:
            return 0

        profiles = [self._profile_row(df.iloc[i]) for i in range(len(df))]

        data_profiles = [p for p in profiles[2:10] if not p.is_empty]
        if not data_profiles:
            return 0

        consensus = _RowProfile(
            empty_ratio=float(np.mean([p.empty_ratio for p in data_profiles])),
            numeric_ratio=float(np.mean([p.numeric_ratio for p in data_profiles])),
            avg_cell_length=float(np.mean([p.avg_cell_length for p in data_profiles])),
            has_metadata_pattern=False,
            cell_count=int(np.mean([p.cell_count for p in data_profiles])),
            is_empty=False,
        )

        skip_rows = 0
        for i in range(min(6, len(profiles))):
            score = self._score_as_metadata(profiles[i], consensus)
            if score >= 0.35:
                skip_rows = i + 1
            else:
                break

        return skip_rows


# ============================================================================
# Core
# ============================================================================

class CSVHandler(FormatHandler):
    """Handler for CSV and TSV files."""

    def can_handle(self, format_type: str) -> bool:
        """Check if this handler can process the format."""
        return format_type in ("csv", "tsv", "txt")

    def __init__(self) -> None:
        """Initialize handler with metadata detector."""
        self._detector = MetadataRowDetector()

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

        # Auto-detect metadata rows to skip
        skip_rows = options.skip_rows
        if options.auto_detect_header and skip_rows == 0:
            if is_fileobj:
                skip_rows = self._detector.detect_skip_rows_from_text(text_data, delimiter)
                # Reset StringIO after detection
                source_for_pandas = io.StringIO(text_data)
            else:
                skip_rows = self._detector.detect_skip_rows(
                    file_source, encoding, delimiter
                )

        na_values = options.na_values or DEFAULT_NA_VALUES

        header = 0 if options.header_rows > 0 else None

        engine = "python" if options.skip_footer > 0 else "c"

        try:
            df = pd.read_csv(
                source_for_pandas,
                encoding      = None if is_fileobj else encoding,  # Already decoded for StringIO
                delimiter     = delimiter,
                skiprows      = skip_rows if options.header_rows <= 1 else 0,
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
