"""File format detection using binary signatures and content analysis."""

# ============================================================================
# Imports
# ============================================================================

import zipfile
from pathlib import Path

from messy_xlsx._source import SourceHandle, SourceInput
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

    _accepts_source_handle = True

    def detect(
        self,
        file_or_path: SourceInput | SourceHandle,
        filename: str | None = None,
    ) -> FormatInfo:
        """Detect file format from path or file-like object.

        Args:
            file_or_path: Path to file or file-like object (BytesIO, etc.)
            filename: Optional filename hint for extension-based detection when using file-like objects
        """
        owns_handle = not isinstance(file_or_path, SourceHandle)
        try:
            handle = SourceHandle.coerce(file_or_path, filename=filename)
        except Exception as e:
            raise FormatError(
                f"Cannot read from file object: {e}",
                file_path=filename or "<stream>",
            ) from e

        try:
            return self._detect_from_handle(handle)
        finally:
            if owns_handle:
                handle.close()

    def _detect_from_handle(self, handle: SourceHandle) -> FormatInfo:
        """Detect a format through a repeatable source handle."""
        file_path = handle.path
        if file_path is not None and not file_path.exists():
            raise FormatError(f"File not found: {file_path}", file_path=str(file_path))

        header = self._read_header(handle)
        if not header:
            if file_path is not None:
                raise FormatError(f"File is empty: {file_path}", file_path=str(file_path))
            raise FormatError("File object is empty", file_path=handle.description)

        for signature, format_family in SIGNATURES.items():
            if header.startswith(signature):
                return self._analyze_format_family(handle, format_family)

        if self._is_text_based(header):
            return self._analyze_text_format_from_bytes(header, handle.filename_hint)

        extension_source = file_path
        if extension_source is None and handle.filename_hint:
            extension_source = Path(handle.filename_hint)
        if extension_source is not None:
            return self._detect_from_extension(extension_source)

        return FormatInfo(format_type="unknown", confidence=0.0)

    def _read_header(self, handle: SourceHandle) -> bytes:
        """Read a bounded header while preserving a caller stream's position."""
        try:
            with handle.open_binary() as stream:
                header = stream.read(HEADER_SIZE)
            if isinstance(header, bytes):
                return header
            if isinstance(header, (bytearray, memoryview)):
                return bytes(header)
            raise TypeError(
                "Binary source read() must return bytes, bytearray, or memoryview; "
                f"got {type(header).__name__}"
            )
        except PermissionError as e:
            if handle.path is not None:
                raise FormatError(
                    f"Permission denied: {handle.path}",
                    file_path=str(handle.path),
                ) from e
            raise FormatError(
                f"Cannot read from file object: {e}",
                file_path=handle.description,
            ) from e
        except OSError as e:
            if handle.path is not None:
                raise FormatError(
                    f"Cannot read file: {handle.path}",
                    file_path=str(handle.path),
                ) from e
            raise FormatError(
                f"Cannot read from file object: {e}",
                file_path=handle.description,
            ) from e
        except Exception as e:
            raise FormatError(
                f"Cannot read from file object: {e}",
                file_path=handle.description,
            ) from e

    def _analyze_format_family(
        self,
        handle: SourceHandle,
        format_family: str,
    ) -> FormatInfo:
        """Analyze a signature family through the source handle."""
        if format_family == "zip_based":
            return self._analyze_zip_format(handle)
        if format_family == "ole2":
            return FormatInfo(format_type="xls", confidence=0.95, version="OLE2 Compound Document")
        if format_family.startswith("xls_biff"):
            return FormatInfo(format_type="xls", confidence=0.95, version=format_family.upper())
        return FormatInfo(format_type="unknown", confidence=0.0)

    def _analyze_zip_format(self, handle: SourceHandle) -> FormatInfo:
        """Analyze a ZIP-based source without sharing its cursor."""
        try:
            with handle.open_binary() as stream, zipfile.ZipFile(stream, "r") as archive:
                filelist = set(archive.namelist())

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
            if handle.path is not None:
                return self._detect_from_extension(handle.path)
            return FormatInfo(format_type="unknown", confidence=0.0)

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

    def _is_text_based(self, header: bytes) -> bool:
        """Check if file appears to be text-based."""
        text_chars = bytes(range(32, 127)) + b"\n\r\t"
        sample = header[:1000]

        if not sample:
            return False

        text_ratio = sum(1 for byte in sample if byte in text_chars) / len(sample)
        return text_ratio > 0.8

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

    def validate(
        self,
        file_path: SourceInput | SourceHandle,
    ) -> tuple[bool, str | None]:
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
