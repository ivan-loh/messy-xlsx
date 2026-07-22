"""Bounded-memory replay storage for caller-owned binary streams."""

from __future__ import annotations

import io
import os
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from types import TracebackType
from typing import BinaryIO, cast

from messy_xlsx._fallback_signals import (
    _contains_process_failure,
    _exception_traceback,
    _FallbackBlockReason,
    _mark_fallback_blocked,
    _raise_with_traceback,
    _safe_add_note,
    _type_name,
)

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
    error = TypeError(
        "Binary source read() must return bytes, bytearray, or memoryview; "
        f"got {type(value).__name__}"
    )
    raise _mark_fallback_blocked(
        error,
        _FallbackBlockReason.SOURCE_OWNERSHIP,
    )


def _storage_error(error: OSError) -> _SpoolStorageError:
    return _SpoolStorageError(str(error))


def _cleanup_takes_precedence(error: BaseException) -> bool:
    """Return whether source teardown must replace an operation failure."""
    return _contains_process_failure(error)


_CleanupWinner = tuple[BaseException, TracebackType | None]


def _record_cleanup_failure(
    primary_error: BaseException,
    cleanup_error: BaseException,
    label: str,
    winner: _CleanupWinner | None,
) -> _CleanupWinner | None:
    """Record one teardown failure and retain the first process-level winner."""
    if winner is not None:
        _safe_add_note(winner[0], f"{label}: {_type_name(cleanup_error)}")
        return winner
    if _contains_process_failure(cleanup_error):
        _mark_fallback_blocked(
            cleanup_error,
            _FallbackBlockReason.SOURCE_OWNERSHIP,
        )
        _safe_add_note(
            cleanup_error,
            f"source operation also failed: {_type_name(primary_error)}",
        )
        return cleanup_error, _exception_traceback(cleanup_error)
    _mark_fallback_blocked(
        primary_error,
        _FallbackBlockReason.SOURCE_OWNERSHIP,
    )
    _safe_add_note(primary_error, f"{label}: {_type_name(cleanup_error)}")
    return None


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
        error = ValueError("A non-seekable source must be positioned at byte 0")
        raise _mark_fallback_blocked(
            error,
            _FallbackBlockReason.SOURCE_OWNERSHIP,
        )
    return False, 0


def _cleanup_pending_spill(
    descriptor: int | None,
    opened: BinaryIO | None,
    raw_path: str,
    error: BaseException,
) -> None:
    winner: _CleanupWinner | None = None
    if opened is not None:
        try:
            opened.close()
        except BaseException as cleanup_error:
            winner = _record_cleanup_failure(
                error,
                cleanup_error,
                "temporary file close also failed",
                winner,
            )
    elif descriptor is not None:
        try:
            os.close(descriptor)
        except BaseException as cleanup_error:
            winner = _record_cleanup_failure(
                error,
                cleanup_error,
                "temporary descriptor close also failed",
                winner,
            )
    try:
        os.unlink(raw_path)
    except BaseException as cleanup_error:
        winner = _record_cleanup_failure(
            error,
            cleanup_error,
            "temporary file removal also failed",
            winner,
        )
    if winner is not None:
        _raise_with_traceback(*winner)


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
    winner: _CleanupWinner | None = None
    if opened is not None:
        try:
            opened.close()
        except BaseException as cleanup_error:
            winner = _record_cleanup_failure(
                error,
                cleanup_error,
                "temporary file close also failed",
                winner,
            )
    if path is not None:
        try:
            path.unlink(missing_ok=True)
        except BaseException as cleanup_error:
            winner = _record_cleanup_failure(
                error,
                cleanup_error,
                "temporary file removal also failed",
                winner,
            )
    if winner is not None:
        _raise_with_traceback(*winner)


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
            if _cleanup_takes_precedence(restore_error):
                _mark_fallback_blocked(
                    restore_error,
                    _FallbackBlockReason.SOURCE_OWNERSHIP,
                )
                raise
            _mark_fallback_blocked(
                primary_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            _safe_add_note(
                primary_error,
                f"cursor restoration also failed: {_type_name(restore_error)}",
            )
            return
        _mark_fallback_blocked(
            restore_error,
            _FallbackBlockReason.SOURCE_OWNERSHIP,
        )
        if completed is not None:
            try:
                completed.close()
            except BaseException as cleanup_error:
                if _cleanup_takes_precedence(cleanup_error):
                    _mark_fallback_blocked(
                        cleanup_error,
                        _FallbackBlockReason.SOURCE_OWNERSHIP,
                    )
                    _safe_add_note(
                        cleanup_error,
                        f"cursor restoration also failed: {_type_name(restore_error)}",
                    )
                    raise
                _safe_add_note(
                    restore_error,
                    f"temporary spool cleanup also failed: {_type_name(cleanup_error)}",
                )
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
            error = ValueError("ReplaySpool is closed")
            raise _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
