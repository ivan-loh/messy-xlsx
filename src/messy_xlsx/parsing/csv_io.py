"""Neutral bounded binary primitives shared by CSV parser frontends."""

from __future__ import annotations

import codecs
import csv
import io
import warnings
from collections import deque
from typing import Any, Literal

import pandas as pd

from messy_xlsx.exceptions import FormatError
from messy_xlsx.normalization.plan import MAX_SAMPLE_BYTES

_READ_CHUNK_BYTES = 64 * 1024


def _coerce_binary_read(value: object) -> bytes:
    if isinstance(value, bytes):
        return value
    if isinstance(value, bytearray):
        return bytes(value)
    if isinstance(value, memoryview):
        return value.tobytes()
    raise TypeError(
        "Binary source read() must return bytes, bytearray, or memoryview; "
        f"got {type(value).__name__}"
    )


def _encoding_layout(encoding: str) -> tuple[int, Literal["little", "big"]]:
    normalized = codecs.lookup(encoding).name.replace("_", "-")
    if normalized in {"utf-16", "utf-16-le"}:
        return 2, "little"
    if normalized == "utf-16-be":
        return 2, "big"
    if normalized in {"utf-32", "utf-32-le"}:
        return 4, "little"
    if normalized == "utf-32-be":
        return 4, "big"
    return 1, "little"


class LogicalRecordByteScanner:
    """Reject oversized records/windows without retaining their byte payload."""

    def __init__(
        self,
        encoding: str,
        delimiter: str,
        description: str,
        *,
        enforce_total: bool,
        ignored_prefix_records: int = 0,
    ) -> None:
        normalized = codecs.lookup(encoding).name.replace("_", "-")
        self._unit_width, self._byteorder = _encoding_layout(encoding)
        self._uses_utf8_sig = normalized == "utf-8-sig"
        self._utf8_bom_pending = self._uses_utf8_sig
        self._bom_buffer = b""
        self._delimiter = ord(delimiter[0]) if delimiter else ord(",")
        self._description = description
        self._enforce_total = enforce_total
        self._ignored_records = ignored_prefix_records
        self._carry = b""
        self._record_bytes = 0
        self._window_bytes = 0
        self._in_quotes = False
        self._quote_pending = False
        self._at_field_start = True
        self._at_start = True
        self._pending_cr = False

    def reset(self) -> None:
        self._carry = b""
        self._record_bytes = 0
        self._window_bytes = 0
        self._in_quotes = False
        self._quote_pending = False
        self._at_field_start = True
        self._at_start = True
        self._pending_cr = False
        self._bom_buffer = b""
        self._utf8_bom_pending = self._uses_utf8_sig

    def feed(self, data: bytes) -> None:  # noqa: C901
        if not data:
            return
        if self._utf8_bom_pending:
            combined_bom = self._bom_buffer + data
            if len(combined_bom) < len(codecs.BOM_UTF8) and codecs.BOM_UTF8.startswith(
                combined_bom
            ):
                self._bom_buffer = combined_bom
                return
            self._utf8_bom_pending = False
            self._bom_buffer = b""
            if combined_bom.startswith(codecs.BOM_UTF8):
                self._record_bytes += len(codecs.BOM_UTF8)
                if self._ignored_records == 0:
                    self._window_bytes += len(codecs.BOM_UTF8)
                self._at_start = False
                data = combined_bom[len(codecs.BOM_UTF8) :]
            else:
                data = combined_bom
        if self._unit_width == 1:
            for codepoint in data:
                self._accept(codepoint, 1)
        else:
            combined = self._carry + data
            complete = len(combined) - (len(combined) % self._unit_width)
            for offset in range(0, complete, self._unit_width):
                codepoint = int.from_bytes(
                    combined[offset : offset + self._unit_width],
                    self._byteorder,
                )
                self._accept(codepoint, self._unit_width)
            self._carry = combined[complete:]
            if self._record_bytes + len(self._carry) > MAX_SAMPLE_BYTES:
                self._raise_oversized("logical CSV record")
        if self._enforce_total and self._window_bytes > MAX_SAMPLE_BYTES:
            self._raise_oversized("sample window")

    def _accept(self, codepoint: int, byte_count: int) -> None:  # noqa: C901
        if self._pending_cr and codepoint != ord("\n"):
            self._finish_record()
        self._record_bytes += byte_count
        if self._ignored_records == 0:
            self._window_bytes += byte_count
        if self._record_bytes > MAX_SAMPLE_BYTES:
            self._raise_oversized("logical CSV record")
        if self._at_start and codepoint == 0xFEFF:
            self._at_start = False
            return
        self._at_start = False
        if self._in_quotes:
            if self._quote_pending:
                if codepoint == ord('"'):
                    self._quote_pending = False
                    return
                self._in_quotes = False
                self._quote_pending = False
            elif codepoint == ord('"'):
                self._quote_pending = True
                return
            else:
                return
        if self._pending_cr and codepoint == ord("\n"):
            self._finish_record()
        elif codepoint == ord("\r"):
            self._pending_cr = True
        elif codepoint == ord("\n"):
            self._finish_record()
        elif self._at_field_start and codepoint == ord('"'):
            self._in_quotes = True
            self._at_field_start = False
        elif codepoint == self._delimiter:
            self._at_field_start = True
        else:
            self._at_field_start = False

    def _finish_record(self) -> None:
        if self._ignored_records:
            self._ignored_records -= 1
        self._record_bytes = 0
        self._at_field_start = True
        self._pending_cr = False

    def _raise_oversized(self, subject: str) -> None:
        raise FormatError(
            f"Cannot parse CSV file: {subject} exceeds the {MAX_SAMPLE_BYTES}-byte sample budget",
            file_path=self._description,
            detected_format="csv",
        )


class LogicalRecordBudgetReader(io.RawIOBase):
    """Non-owning binary proxy enforcing record and optional window budgets."""

    def __init__(
        self,
        stream: Any,
        encoding: str,
        delimiter: str,
        description: str,
        *,
        enforce_total: bool = True,
        ignored_prefix_records: int = 0,
    ) -> None:
        super().__init__()
        self._stream = stream
        self._scanner = LogicalRecordByteScanner(
            encoding,
            delimiter,
            description,
            enforce_total=enforce_total,
            ignored_prefix_records=ignored_prefix_records,
        )

    def readable(self) -> bool:
        return True

    def seekable(self) -> bool:
        seekable = getattr(self._stream, "seekable", None)
        return bool(seekable()) if callable(seekable) else False

    def read(self, size: int = -1) -> bytes:
        bounded = _READ_CHUNK_BYTES if size is None or size < 0 else size
        raw = _coerce_binary_read(self._stream.read(bounded))
        self._scanner.feed(raw)
        return raw

    def read1(self, size: int = -1) -> bytes:
        bounded = _READ_CHUNK_BYTES if size is None or size < 0 else size
        read1 = getattr(self._stream, "read1", None)
        data = read1(bounded) if callable(read1) else self._stream.read(bounded)
        raw = _coerce_binary_read(data)
        self._scanner.feed(raw)
        return raw

    def readinto(self, buffer: Any) -> int:
        readinto = getattr(self._stream, "readinto", None)
        if callable(readinto):
            count = readinto(buffer)
            if count is None:
                return 0
            self._scanner.feed(bytes(memoryview(buffer)[:count]))
            return int(count)
        raw = self.read(len(buffer))
        memoryview(buffer)[: len(raw)] = raw
        return len(raw)

    def readline(self, size: int | None = -1) -> bytes:
        bounded = _READ_CHUNK_BYTES if size is None or size < 0 else size
        raw = _coerce_binary_read(self._stream.readline(bounded))
        self._scanner.feed(raw)
        return raw

    def tell(self) -> int:
        return int(self._stream.tell())

    def seek(self, offset: int, whence: int = 0) -> int:
        position = int(self._stream.seek(offset, whence))
        self._scanner.reset()
        return position

    def fileno(self) -> int:
        return int(self._stream.fileno())


class LogicalRecordFramer:
    """Split binary CSV input into quote-aware logical records."""

    def __init__(
        self,
        encoding: str,
        delimiter: str,
        description: str,
        *,
        max_record_bytes: int | None = None,
    ) -> None:
        normalized = codecs.lookup(encoding).name.replace("_", "-")
        self._unit_width, self._byteorder = _encoding_layout(encoding)
        self._utf8_bom_pending = normalized == "utf-8-sig"
        self._bom_buffer = b""
        self._delimiter = ord(delimiter[0]) if delimiter else ord(",")
        self._description = description
        self._max_record_bytes = max_record_bytes
        self._carry = b""
        self._record = bytearray()
        self._in_quotes = False
        self._quote_pending = False
        self._at_field_start = True
        self._pending_cr = False
        self._at_start = True

    @property
    def at_record_boundary(self) -> bool:
        """Return whether another reader can safely resume raw record splitting."""
        return not (
            self._carry
            or self._record
            or self._in_quotes
            or self._quote_pending
            or self._pending_cr
        )

    def feed(self, data: bytes) -> list[bytes]:
        completed: list[bytes] = []
        if self._utf8_bom_pending:
            combined_bom = self._bom_buffer + data
            if len(combined_bom) < len(codecs.BOM_UTF8) and codecs.BOM_UTF8.startswith(
                combined_bom
            ):
                self._bom_buffer = combined_bom
                return completed
            self._utf8_bom_pending = False
            self._bom_buffer = b""
            if combined_bom.startswith(codecs.BOM_UTF8):
                self._record.extend(codecs.BOM_UTF8)
                self._check_record_size()
                self._at_start = False
                data = combined_bom[len(codecs.BOM_UTF8) :]
            else:
                data = combined_bom
        combined = self._carry + data
        complete = len(combined) - (len(combined) % self._unit_width)
        if self._unit_width == 1:
            self._feed_single_byte(combined[:complete], completed)
            self._carry = combined[complete:]
            self._check_record_size(len(self._carry))
            return completed
        for offset in range(0, complete, self._unit_width):
            unit = combined[offset : offset + self._unit_width]
            codepoint = int.from_bytes(unit, self._byteorder)
            self._accept(unit, codepoint, completed)
        self._carry = combined[complete:]
        self._check_record_size(len(self._carry))
        return completed

    def _feed_single_byte(self, data: bytes, completed: list[bytes]) -> None:
        """Use C-level LF splitting outside quoted or CR-delimited regions."""
        offset = 0
        while offset < len(data):
            if self._in_quotes or self._quote_pending or self._pending_cr:
                self._accept(data[offset : offset + 1], data[offset], completed)
                offset += 1
                continue
            quote_at = data.find(b'"', offset)
            cr_at = data.find(b"\r", offset)
            boundaries = [position for position in (quote_at, cr_at) if position >= 0]
            boundary = min(boundaries) if boundaries else len(data)
            if boundary > offset:
                self._accept_unquoted_lf_bytes(data[offset:boundary], completed)
                offset = boundary
                continue
            self._accept(data[offset : offset + 1], data[offset], completed)
            offset += 1

    def _accept_unquoted_lf_bytes(
        self,
        data: bytes,
        completed: list[bytes],
    ) -> None:
        """Accept a quote-free, CR-free byte run without a Python byte loop."""
        if not data:
            return
        self._at_start = False
        parts = data.split(b"\n")
        for part in parts[:-1]:
            self._record.extend(part)
            self._record.extend(b"\n")
            self._check_record_size()
            self._finish_record(completed)
        tail = parts[-1]
        if tail:
            self._record.extend(tail)
            self._check_record_size()
            self._at_field_start = tail[-1] == self._delimiter

    def finish(self) -> list[bytes]:
        completed: list[bytes] = []
        if self._carry:
            self._record.extend(self._carry)
            self._carry = b""
            self._check_record_size()
        if self._record:
            completed.append(bytes(self._record))
            self._record.clear()
        self._pending_cr = False
        return completed

    def _accept(  # noqa: C901
        self,
        unit: bytes,
        codepoint: int,
        completed: list[bytes],
    ) -> None:
        if self._pending_cr:
            if codepoint == ord("\n"):
                self._record.extend(unit)
                self._check_record_size()
                self._finish_record(completed)
                return
            self._finish_record(completed)
        self._record.extend(unit)
        self._check_record_size()
        if self._at_start and codepoint == 0xFEFF:
            self._at_start = False
            return
        self._at_start = False
        if self._in_quotes:
            if self._quote_pending:
                if codepoint == ord('"'):
                    self._quote_pending = False
                    return
                self._in_quotes = False
                self._quote_pending = False
            elif codepoint == ord('"'):
                self._quote_pending = True
                return
            else:
                return
        if codepoint == ord("\r"):
            self._pending_cr = True
        elif codepoint == ord("\n"):
            self._finish_record(completed)
        elif self._at_field_start and codepoint == ord('"'):
            self._in_quotes = True
            self._at_field_start = False
        elif codepoint == self._delimiter:
            self._at_field_start = True
        else:
            self._at_field_start = False

    def _finish_record(self, completed: list[bytes]) -> None:
        completed.append(bytes(self._record))
        self._record.clear()
        self._in_quotes = False
        self._quote_pending = False
        self._at_field_start = True
        self._pending_cr = False

    def _check_record_size(self, extra: int = 0) -> None:
        if (
            self._max_record_bytes is not None
            and len(self._record) + extra > self._max_record_bytes
        ):
            raise FormatError(
                "Cannot parse CSV file: logical CSV record exceeds "
                f"the {self._max_record_bytes}-byte sample budget",
                file_path=self._description,
                detected_format="csv",
            )


class _BufferedTransformReader(io.RawIOBase):
    """Shared bounded binary read surface for record transforms."""

    def __init__(self) -> None:
        super().__init__()
        self._output = bytearray()
        self._eof = False

    def readable(self) -> bool:
        return True

    def read(self, size: int = -1) -> bytes:
        bounded = _READ_CHUNK_BYTES if size is None or size < 0 else size
        self._fill(bounded)
        count = min(bounded, len(self._output))
        result = bytes(self._output[:count])
        del self._output[:count]
        return result

    def read1(self, size: int = -1) -> bytes:
        return self.read(size)

    def readinto(self, buffer: Any) -> int:
        raw = self.read(len(buffer))
        memoryview(buffer)[: len(raw)] = raw
        return len(raw)

    def readline(self, size: int | None = -1) -> bytes:
        bounded = _READ_CHUNK_BYTES if size is None or size < 0 else size
        self._fill(1)
        while len(self._output) < bounded and b"\n" not in self._output and not self._eof:
            self._fill(len(self._output) + 1)
        newline = self._output.find(b"\n")
        count = min(bounded, newline + 1) if newline >= 0 else min(bounded, len(self._output))
        result = bytes(self._output[:count])
        del self._output[:count]
        return result

    def _fill(self, minimum: int) -> None:
        raise NotImplementedError


class RecordLimitingReader(_BufferedTransformReader):
    """Expose at most a fixed count of complete logical records."""

    def __init__(
        self,
        stream: Any,
        encoding: str,
        delimiter: str,
        description: str,
        max_records: int,
        *,
        owns_stream: bool = False,
    ) -> None:
        super().__init__()
        self._stream = stream
        self._owns_stream = owns_stream
        self._description = description
        self._framer = LogicalRecordFramer(
            encoding,
            delimiter,
            description,
            max_record_bytes=MAX_SAMPLE_BYTES,
        )
        self._remaining = max_records

    def close(self) -> None:
        if self.closed and not self._owns_stream:
            return
        if self._owns_stream:
            self._stream.close()
            self._owns_stream = False
        super().close()

    def _fill(self, minimum: int) -> None:
        while len(self._output) < minimum and not self._eof:
            raw = _coerce_binary_read(self._stream.read(_READ_CHUNK_BYTES))
            records = self._framer.feed(raw) if raw else self._framer.finish()
            for record in records:
                if self._remaining <= 0:
                    self._eof = True
                    break
                self._remaining -= 1
                self._output.extend(record)
            if self._remaining <= 0 or not raw:
                self._eof = True


class MalformedRecordFilteringReader(_BufferedTransformReader):
    """Remove excess-field records before pandas chunk boundaries can affect them."""

    def __init__(
        self,
        stream: Any,
        encoding: str,
        encoding_errors: str,
        delimiter: str,
        description: str,
        expected_fields: int,
        protected_prefix_records: int,
        *,
        owns_stream: bool,
    ) -> None:
        super().__init__()
        self._stream = stream
        self._owns_stream = owns_stream
        self._encoding = encoding
        self._encoding_errors = encoding_errors
        self._delimiter = delimiter
        self._framer = LogicalRecordFramer(encoding, delimiter, description)
        unit_width, _byteorder = _encoding_layout(encoding)
        encoded_delimiter = delimiter.encode(encoding)
        if encoded_delimiter.startswith(codecs.BOM_UTF8):
            encoded_delimiter = encoded_delimiter[len(codecs.BOM_UTF8) :]
        self._raw_delimiter = (
            encoded_delimiter if unit_width == 1 and len(encoded_delimiter) == 1 else None
        )
        self._raw_blank = unit_width == 1
        self._raw_fast_active = self._raw_delimiter is not None
        self._raw_fast_pending = bytearray()
        self._expected_fields = expected_fields
        self._accepted_data_width = expected_fields
        self._protected = protected_prefix_records
        self._seen_data_record = False

    def close(self) -> None:
        if self.closed and not self._owns_stream:
            return
        if self._owns_stream:
            self._stream.close()
            self._owns_stream = False
        super().close()

    def _fill(self, minimum: int) -> None:
        while len(self._output) < minimum and not self._eof:
            raw = _coerce_binary_read(self._stream.read(max(_READ_CHUNK_BYTES, minimum)))
            if raw and self._raw_fast_active and b'"' not in raw and b"\r" not in raw:
                self._feed_raw_unquoted_lf(raw)
                continue
            if raw and self._raw_fast_active:
                raw = bytes(self._raw_fast_pending) + raw
                self._raw_fast_pending.clear()
                self._raw_fast_active = False
            records = self._framer.feed(raw) if raw else self._framer.finish()
            self._accept_framed_records(records)
            if raw and self._framer.at_record_boundary:
                self._raw_fast_active = self._raw_delimiter is not None
            if raw:
                continue
            if self._raw_fast_active and self._raw_fast_pending:
                record = bytes(self._raw_fast_pending)
                self._raw_fast_pending.clear()
                if self._accept_record(record):
                    self._output.extend(record)
            self._eof = True

    def _feed_raw_unquoted_lf(self, raw: bytes) -> None:
        """Classify complete LF records with C-level split/count/join operations."""
        if self._raw_fast_pending:
            data = bytes(self._raw_fast_pending) + raw
            self._raw_fast_pending.clear()
        else:
            data = raw
        last_newline = data.rfind(b"\n")
        if last_newline < 0:
            self._raw_fast_pending.extend(data)
            return
        complete_blob = data[: last_newline + 1]
        self._raw_fast_pending.extend(data[last_newline + 1 :])
        complete = complete_blob[:-1].split(b"\n")
        accepted: list[bytes] = []
        delimiter = self._raw_delimiter
        assert delimiter is not None
        index = 0
        while index < len(complete) and self._protected:
            accepted.append(complete[index])
            self._protected -= 1
            index += 1
        while index < len(complete) and not self._seen_data_record:
            record = complete[index]
            accepted.append(record)
            index += 1
            if not record.strip(b" \t\r\n"):
                continue
            self._seen_data_record = True
            self._accepted_data_width = max(
                self._expected_fields,
                record.count(delimiter) + 1,
            )
        accepted_width = self._accepted_data_width
        for record in complete[index:]:
            if record.count(delimiter) + 1 > accepted_width:
                warnings.warn(
                    "Skipping malformed CSV record with excess fields",
                    pd.errors.ParserWarning,
                    stacklevel=4,
                )
                continue
            accepted.append(record)
        if not accepted:
            return
        self._output.extend(b"\n".join(accepted))
        self._output.extend(b"\n")

    def _accept_framed_records(self, records: list[bytes]) -> None:
        for record in records:
            if self._accept_record(record):
                self._output.extend(record)

    def _accept_record(self, record: bytes) -> bool:
        if self._protected:
            self._protected -= 1
            return True
        field_count = self._field_count(record)
        if not self._seen_data_record:
            if self._is_blank(record):
                return True
            self._seen_data_record = True
            self._accepted_data_width = max(
                self._expected_fields,
                field_count,
            )
            return True
        if field_count > self._accepted_data_width:
            warnings.warn(
                "Skipping malformed CSV record with excess fields",
                pd.errors.ParserWarning,
                stacklevel=4,
            )
            return False
        return True

    def _field_count(self, record: bytes) -> int:
        if self._raw_delimiter is not None and b'"' not in record:
            return record.count(self._raw_delimiter) + 1
        text = record.decode(self._encoding, errors=self._encoding_errors)
        try:
            return len(next(csv.reader([text], delimiter=self._delimiter)))
        except (csv.Error, StopIteration):
            return self._expected_fields

    def _is_blank(self, record: bytes) -> bool:
        if self._raw_blank:
            return not record.strip(b" \t\r\n")
        text = record.decode(self._encoding, errors=self._encoding_errors)
        return not text.strip(" \t\r\n")


class FooterTrimmingReader(_BufferedTransformReader):
    """Lazily withhold the final physical logical records from pandas."""

    def __init__(
        self,
        stream: Any,
        encoding: str,
        delimiter: str,
        description: str,
        skip_footer: int,
        protected_prefix_records: int,
        *,
        owns_stream: bool = False,
        max_record_bytes: int | None = None,
        max_pending_bytes: int | None = None,
    ) -> None:
        super().__init__()
        self._stream = stream
        self._owns_stream = owns_stream
        self._description = description
        self._framer = LogicalRecordFramer(
            encoding,
            delimiter,
            description,
            max_record_bytes=max_record_bytes,
        )
        self._skip_footer = skip_footer
        self._protected = protected_prefix_records
        self._pending: deque[bytes] = deque()
        self._pending_bytes = 0
        self._max_pending_bytes = max_pending_bytes

    @property
    def pending_record_count(self) -> int:
        return len(self._pending)

    def close(self) -> None:
        if self.closed and not self._owns_stream:
            return
        if self._owns_stream:
            self._stream.close()
            self._owns_stream = False
        super().close()

    def _fill(self, minimum: int) -> None:
        while len(self._output) < minimum and not self._eof:
            raw = _coerce_binary_read(self._stream.read(_READ_CHUNK_BYTES))
            records = self._framer.feed(raw) if raw else self._framer.finish()
            for record in records:
                if self._protected:
                    self._protected -= 1
                    self._output.extend(record)
                    continue
                self._pending.append(record)
                self._pending_bytes += len(record)
                if len(self._pending) > self._skip_footer:
                    released = self._pending.popleft()
                    self._pending_bytes -= len(released)
                    self._output.extend(released)
                if (
                    self._max_pending_bytes is not None
                    and self._pending_bytes > self._max_pending_bytes
                ):
                    raise FormatError(
                        "Cannot parse CSV file: footer lookahead exceeds "
                        f"the {self._max_pending_bytes}-byte sample budget",
                        file_path=self._description,
                        detected_format="csv",
                    )
            if not raw:
                self._pending.clear()
                self._pending_bytes = 0
                self._eof = True
