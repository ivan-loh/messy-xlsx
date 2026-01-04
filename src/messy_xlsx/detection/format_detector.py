"""File format detection using binary signatures and content analysis."""

# ============================================================================
# Imports
# ============================================================================

import io
import zipfile
from pathlib import Path
from typing import BinaryIO

from messy_xlsx.exceptions import FormatError
from messy_xlsx.models import FormatInfo


# ============================================================================
# Configuration
# ============================================================================

SIGNATURES = {
    b"PK\x03\x04": "zip_based",
    b"PK\x05\x06": "zip_based",
    b"PK\x07\x08": "zip_based",
    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1": "ole2",
    b"\x09\x08\x10\x00\x00\x06\x05\x00": "xls_biff8",
    b"\x09\x08\x08\x00\x00\x06\x05\x00": "xls_biff8",
}

HEADER_SIZE = 8192


# ============================================================================
# Format Detector
# ============================================================================

class FormatDetector:
    """Detect file format using binary signatures and content analysis."""

    def detect(self, file_or_path: Path | str | BinaryIO, filename: str | None = None) -> FormatInfo:
        """Detect file format from path or file-like object.

        Args:
            file_or_path: Path to file or file-like object (BytesIO, etc.)
            filename: Optional filename hint for extension-based detection when using file-like objects
        """
        # Handle file-like objects
        if hasattr(file_or_path, "read"):
            return self._detect_from_fileobj(file_or_path, filename)

        # Handle path
        file_path = Path(file_or_path)

        if not file_path.exists():
            raise FormatError(f"File not found: {file_path}", file_path=str(file_path))

        try:
            with open(file_path, "rb") as f:
                header = f.read(HEADER_SIZE)
        except PermissionError as e:
            raise FormatError(f"Permission denied: {file_path}", file_path=str(file_path)) from e
        except OSError as e:
            raise FormatError(f"Cannot read file: {file_path}", file_path=str(file_path)) from e

        if not header:
            raise FormatError(f"File is empty: {file_path}", file_path=str(file_path))

        for signature, format_family in SIGNATURES.items():
            if header.startswith(signature):
                return self._analyze_format_family(file_path, format_family, header)

        if self._is_text_based(header):
            return self._analyze_text_format(file_path, header)

        return self._detect_from_extension(file_path)

    def _detect_from_fileobj(self, fileobj: BinaryIO, filename: str | None = None) -> FormatInfo:
        """Detect format from file-like object."""
        # Save position
        start_pos = fileobj.tell() if hasattr(fileobj, "tell") else 0

        try:
            header = fileobj.read(HEADER_SIZE)
        except Exception as e:
            raise FormatError(f"Cannot read from file object: {e}", file_path="<stream>") from e
        finally:
            # Reset position
            if hasattr(fileobj, "seek"):
                fileobj.seek(start_pos)

        if not header:
            raise FormatError("File object is empty", file_path="<stream>")

        for signature, format_family in SIGNATURES.items():
            if header.startswith(signature):
                return self._analyze_format_family_from_bytes(header, format_family, fileobj)

        if self._is_text_based(header):
            return self._analyze_text_format_from_bytes(header, filename)

        # Fall back to extension if filename provided
        if filename:
            return self._detect_from_extension(Path(filename))

        return FormatInfo(format_type="unknown", confidence=0.0)

    def _analyze_format_family_from_bytes(
        self, header: bytes, format_family: str, fileobj: BinaryIO
    ) -> FormatInfo:
        """Analyze format family from file-like object."""
        if format_family == "zip_based":
            return self._analyze_zip_format_from_fileobj(fileobj)
        elif format_family == "ole2":
            return FormatInfo(format_type="xls", confidence=0.95, version="OLE2 Compound Document")
        elif format_family.startswith("xls_biff"):
            return FormatInfo(format_type="xls", confidence=0.95, version=format_family.upper())
        else:
            return FormatInfo(format_type="unknown", confidence=0.0)

    def _analyze_zip_format_from_fileobj(self, fileobj: BinaryIO) -> FormatInfo:
        """Analyze ZIP-based format from file-like object."""
        start_pos = fileobj.tell() if hasattr(fileobj, "tell") else 0

        try:
            # Reset to beginning for zipfile
            if hasattr(fileobj, "seek"):
                fileobj.seek(0)

            with zipfile.ZipFile(fileobj, "r") as zf:
                filelist = set(zf.namelist())

                if "xl/workbook.xml" in filelist:
                    has_macros = "xl/vbaProject.bin" in filelist
                    is_encrypted = "EncryptionInfo" in filelist

                    return FormatInfo(
                        format_type="xlsm" if has_macros else "xlsx",
                        confidence=0.95,
                        version="Office Open XML",
                        has_macros=has_macros,
                        is_encrypted=is_encrypted,
                        is_compressed=True,
                    )

                if "xl/workbook.bin" in filelist:
                    has_macros = "xl/vbaProject.bin" in filelist

                    return FormatInfo(
                        format_type="xlsb",
                        confidence=0.95,
                        version="Excel Binary",
                        has_macros=has_macros,
                        is_compressed=True,
                    )

                return FormatInfo(
                    format_type="unknown",
                    confidence=0.3,
                    version="ZIP archive (not Excel)",
                    is_compressed=True,
                )

        except zipfile.BadZipFile:
            return FormatInfo(format_type="unknown", confidence=0.0)
        finally:
            if hasattr(fileobj, "seek"):
                fileobj.seek(start_pos)

    def _analyze_text_format_from_bytes(self, header: bytes, filename: str | None) -> FormatInfo:
        """Analyze text-based format from bytes."""
        try:
            text_sample = header.decode("utf-8", errors="ignore")
        except Exception:
            text_sample = header.decode("latin-1", errors="ignore")

        lines = [line for line in text_sample.split("\n")[:10] if line.strip()]

        if len(lines) < 2:
            return FormatInfo(format_type="csv", confidence=0.5, encoding="utf-8")

        delimiter, confidence = self._detect_delimiter(lines)
        format_type = "tsv" if delimiter == "\t" else "csv"
        encoding = self._detect_encoding(header)

        return FormatInfo(format_type=format_type, confidence=confidence, encoding=encoding)

    def _analyze_format_family(self, file_path: Path, format_family: str, header: bytes) -> FormatInfo:
        """Analyze format within a detected family."""
        if format_family == "zip_based":
            return self._analyze_zip_format(file_path)
        elif format_family == "ole2":
            return FormatInfo(format_type="xls", confidence=0.95, version="OLE2 Compound Document")
        elif format_family.startswith("xls_biff"):
            return FormatInfo(format_type="xls", confidence=0.95, version=format_family.upper())
        else:
            return FormatInfo(format_type="unknown", confidence=0.0)

    def _analyze_zip_format(self, file_path: Path) -> FormatInfo:
        """Analyze ZIP-based Office Open XML formats."""
        try:
            with zipfile.ZipFile(file_path, "r") as zf:
                filelist = set(zf.namelist())

                if "xl/workbook.xml" in filelist:
                    has_macros = "xl/vbaProject.bin" in filelist
                    is_encrypted = "EncryptionInfo" in filelist

                    return FormatInfo(
                        format_type="xlsm" if has_macros else "xlsx",
                        confidence=0.95,
                        version="Office Open XML",
                        has_macros=has_macros,
                        is_encrypted=is_encrypted,
                        is_compressed=True,
                    )

                if "xl/workbook.bin" in filelist:
                    has_macros = "xl/vbaProject.bin" in filelist

                    return FormatInfo(
                        format_type="xlsb",
                        confidence=0.95,
                        version="Excel Binary",
                        has_macros=has_macros,
                        is_compressed=True,
                    )

                return FormatInfo(
                    format_type="unknown",
                    confidence=0.3,
                    version="ZIP archive (not Excel)",
                    is_compressed=True,
                )

        except zipfile.BadZipFile:
            return self._detect_from_extension(file_path)

    def _is_text_based(self, header: bytes) -> bool:
        """Check if file appears to be text-based."""
        text_chars = bytes(range(32, 127)) + b"\n\r\t"
        sample = header[:1000]

        if not sample:
            return False

        text_ratio = sum(1 for byte in sample if byte in text_chars) / len(sample)
        return text_ratio > 0.8

    def _analyze_text_format(self, file_path: Path, header: bytes) -> FormatInfo:
        """Analyze text-based formats (CSV, TSV)."""
        try:
            text_sample = header.decode("utf-8", errors="ignore")
        except Exception:
            text_sample = header.decode("latin-1", errors="ignore")

        lines = [line for line in text_sample.split("\n")[:10] if line.strip()]

        if len(lines) < 2:
            return FormatInfo(format_type="csv", confidence=0.5, encoding="utf-8")

        delimiter, confidence = self._detect_delimiter(lines)
        format_type = "tsv" if delimiter == "\t" else "csv"
        encoding = self._detect_encoding(header)

        return FormatInfo(format_type=format_type, confidence=confidence, encoding=encoding)

    def _detect_delimiter(self, lines: list[str]) -> tuple[str, float]:
        """Detect CSV delimiter from sample lines."""
        delimiters = [",", "\t", ";", "|"]
        best_delimiter = ","
        best_score = 0.0

        for delim in delimiters:
            counts = [line.count(delim) for line in lines if line]

            if not counts or counts[0] == 0:
                continue

            avg_count = sum(counts) / len(counts)
            if len(counts) > 1:
                variance = sum((c - avg_count) ** 2 for c in counts) / len(counts)
            else:
                variance = 0

            score = avg_count / (variance + 1)

            if score > best_score:
                best_score = score
                best_delimiter = delim

        confidence = min(0.9, 0.5 + best_score / 20)

        return best_delimiter, confidence

    def _detect_encoding(self, header: bytes) -> str:
        """Detect text encoding from header bytes."""
        if header.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        if header.startswith(b"\xff\xfe"):
            return "utf-16-le"
        if header.startswith(b"\xfe\xff"):
            return "utf-16-be"

        try:
            header[:1000].decode("utf-8")
            return "utf-8"
        except UnicodeDecodeError:
            pass

        return "latin-1"

    def _detect_from_extension(self, file_path: Path) -> FormatInfo:
        """Fall back to extension-based detection."""
        ext = file_path.suffix.lower()

        extension_map = {
            ".xlsx": ("xlsx", "Office Open XML"),
            ".xlsm": ("xlsm", "Office Open XML with Macros"),
            ".xlsb": ("xlsb", "Excel Binary"),
            ".xls": ("xls", "Legacy Excel"),
            ".csv": ("csv", "Comma-Separated Values"),
            ".tsv": ("tsv", "Tab-Separated Values"),
            ".txt": ("csv", "Text file (assumed CSV)"),
        }

        if ext in extension_map:
            format_type, version = extension_map[ext]
            return FormatInfo(format_type=format_type, confidence=0.5, version=version)

        return FormatInfo(format_type="unknown", confidence=0.0)

    def validate(self, file_path: Path | str) -> tuple[bool, str | None]:
        """Validate that a file can be parsed."""
        try:
            info = self.detect(file_path)

            if info.format_type == "unknown":
                return False, "Unknown file format"

            if info.is_encrypted:
                return False, "File is encrypted"

            return True, None

        except FormatError as e:
            return False, str(e)
        except Exception as e:
            return False, f"Validation error: {e}"
