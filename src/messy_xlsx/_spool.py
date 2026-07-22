"""Bounded-memory replay storage for caller-owned binary streams."""

from __future__ import annotations

import io
import os
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import BinaryIO, cast

DEFAULT_MEMORY_LIMIT = 8 * 1024 * 1024
COPY_CHUNK_SIZE = 1024 * 1024


class _SpoolStorageError(OSError):
    """Identify temporary-storage failures at the source boundary."""


def _coerce_bytes(value: object) -> bytes:
    """Normalize a binary read result and reject text/non-blocking streams."""
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


def _storage_error(error: OSError) -> _SpoolStorageError:
    return _SpoolStorageError(str(error))


def _is_seekable(stream: BinaryIO) -> bool:
    try:
        seekable = getattr(stream, "seekable", None)
        if callable(seekable) and not seekable():
            return False
        position = stream.tell()
        stream.seek(position)
    except (AttributeError, OSError, ValueError):
        return False
    return True


def _capture_position(stream: BinaryIO) -> tuple[bool, int]:
    seekable = _is_seekable(stream)
    if seekable:
        return True, stream.tell()
    try:
        position = stream.tell()
    except (AttributeError, OSError, ValueError):
        position = 0
    if position != 0:
        raise ValueError("A non-seekable source must be positioned at byte 0")
    return False, 0


def _cleanup_pending_spill(
    descriptor: int | None,
    opened: BinaryIO | None,
    raw_path: str,
    error: BaseException,
) -> None:
    if opened is not None:
        try:
            opened.close()
        except BaseException as cleanup_error:
            error.add_note(f"temporary file close also failed: {cleanup_error!r}")
    elif descriptor is not None:
        try:
            os.close(descriptor)
        except BaseException as cleanup_error:
            error.add_note(f"temporary descriptor close also failed: {cleanup_error!r}")
    try:
        os.unlink(raw_path)
    except BaseException as cleanup_error:
        error.add_note(f"temporary file removal also failed: {cleanup_error!r}")


def _create_spill() -> tuple[Path, BinaryIO]:
    try:
        raw_descriptor, raw_path = tempfile.mkstemp(
            prefix="messy-xlsx-",
            suffix=".spool",
        )
    except OSError as error:
        raise _storage_error(error) from error

    descriptor_owner: int | None = raw_descriptor
    opened: BinaryIO | None = None
    try:
        os.chmod(raw_path, 0o600)
        opened = cast(BinaryIO, os.fdopen(raw_descriptor, "wb"))
        descriptor_owner = None
        return Path(raw_path), opened
    except BaseException as error:
        if isinstance(error, OSError):
            storage_error = _storage_error(error)
            _cleanup_pending_spill(descriptor_owner, opened, raw_path, storage_error)
            raise storage_error from error
        _cleanup_pending_spill(descriptor_owner, opened, raw_path, error)
        raise


def _write_spill(stream: BinaryIO, content: bytes | bytearray) -> None:
    try:
        stream.write(content)
    except OSError as error:
        raise _storage_error(error) from error


def _close_spill(stream: BinaryIO) -> None:
    try:
        stream.close()
    except OSError as error:
        raise _storage_error(error) from error


def _cleanup_spill(
    opened: BinaryIO | None,
    path: Path | None,
    error: BaseException,
) -> None:
    if opened is not None:
        try:
            opened.close()
        except BaseException as cleanup_error:
            error.add_note(f"temporary file close also failed: {cleanup_error!r}")
    if path is not None:
        try:
            path.unlink(missing_ok=True)
        except BaseException as cleanup_error:
            error.add_note(f"temporary file removal also failed: {cleanup_error!r}")


def _restore_position(
    stream: BinaryIO,
    entry: int,
    primary_error: BaseException | None,
    completed: ReplaySpool | None,
) -> None:
    try:
        stream.seek(entry)
    except BaseException as restore_error:
        if primary_error is not None:
            primary_error.add_note(f"cursor restoration also failed: {restore_error!r}")
            return
        if completed is not None:
            try:
                completed.close()
            except BaseException as cleanup_error:
                restore_error.add_note(f"temporary spool cleanup also failed: {cleanup_error!r}")
        raise


class ReplaySpool:
    """Replay a complete source from bounded memory or a private temp path."""

    def __init__(self, memory: bytes | None, path: Path | None) -> None:
        self._memory = memory
        self._path = path
        self._closed = False

    @classmethod
    def from_stream(
        cls,
        stream: BinaryIO,
        memory_limit: int = DEFAULT_MEMORY_LIMIT,
    ) -> ReplaySpool:
        """Copy *stream* once, restoring seekable caller cursor state."""
        seekable, entry = _capture_position(stream)
        buffer = bytearray()
        path: Path | None = None
        opened: BinaryIO | None = None
        completed: ReplaySpool | None = None
        primary_error: BaseException | None = None
        try:
            if seekable:
                stream.seek(0)
            while True:
                chunk = _coerce_bytes(stream.read(COPY_CHUNK_SIZE))
                if not chunk:
                    break
                if path is None and len(buffer) + len(chunk) <= memory_limit:
                    buffer.extend(chunk)
                    continue
                if path is None:
                    path, opened = _create_spill()
                    _write_spill(opened, buffer)
                    buffer.clear()
                assert opened is not None
                _write_spill(opened, chunk)
            if opened is not None:
                _close_spill(opened)
                opened = None
            completed = cls(bytes(buffer) if path is None else None, path)
        except BaseException as error:
            primary_error = error
            _cleanup_spill(opened, path, error)
            raise
        finally:
            if seekable:
                _restore_position(stream, entry, primary_error, completed)

        assert completed is not None
        return completed

    @contextmanager
    def open_binary(self) -> Iterator[BinaryIO]:
        """Yield a fresh seekable binary replay."""
        self._ensure_open()
        if self._path is None:
            with io.BytesIO(self._memory or b"") as stream:
                yield stream
            return
        with self._path.open("rb") as stream:
            yield stream

    @contextmanager
    def open_path_or_bytes(self) -> Iterator[Path | bytes]:
        """Yield the bounded-memory bytes or closed spill path."""
        self._ensure_open()
        yield self._path if self._path is not None else (self._memory or b"")

    def close(self) -> None:
        """Release memory and remove any spill path; repeated calls are safe."""
        if self._closed:
            return
        if self._path is not None:
            self._path.unlink(missing_ok=True)
            self._path = None
        self._memory = None
        self._closed = True

    def _ensure_open(self) -> None:
        if self._closed:
            raise ValueError("ReplaySpool is closed")
