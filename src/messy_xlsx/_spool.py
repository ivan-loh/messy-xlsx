"""Bounded-memory replay storage for caller-owned binary streams."""

from __future__ import annotations

import atexit
import io
import os
import tempfile
import weakref
from collections import deque
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from pathlib import Path
from types import TracebackType
from typing import BinaryIO, cast

from messy_xlsx._fallback_signals import (
    _attach_cleanup_failure,
    _attach_operation_failure,
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
_MAX_ORPHAN_STORAGE_OWNERS = 64


class _ReplayStorageOwner:
    """Independent cleanup state that never retains its public spool."""

    __slots__ = ("closed", "memory", "path")

    def __init__(self, memory: bytes | None, path: Path | None) -> None:
        self.memory = memory
        self.path = path
        self.closed = False

    def close(self) -> None:
        if self.closed:
            return
        path = self.path
        if path is not None:
            path.unlink(missing_ok=True)
        self.path = None
        self.memory = None
        self.closed = True


_ORPHAN_STORAGE_OWNERS: deque[_ReplayStorageOwner] = deque()


def _queue_orphan_storage(owner: _ReplayStorageOwner) -> None:
    """Retain a failed finalizer for one bounded retry on a later safe point."""
    if owner.closed or any(candidate is owner for candidate in _ORPHAN_STORAGE_OWNERS):
        return
    if len(_ORPHAN_STORAGE_OWNERS) >= _MAX_ORPHAN_STORAGE_OWNERS:
        _drain_orphan_storages()
    if len(_ORPHAN_STORAGE_OWNERS) < _MAX_ORPHAN_STORAGE_OWNERS:
        _ORPHAN_STORAGE_OWNERS.append(owner)


def _drain_orphan_storages() -> None:
    """Attempt every currently queued owner once; never retry in a loop."""
    pending = tuple(_ORPHAN_STORAGE_OWNERS)
    _ORPHAN_STORAGE_OWNERS.clear()
    for owner in pending:
        try:
            owner.close()
        except BaseException:
            if len(_ORPHAN_STORAGE_OWNERS) < _MAX_ORPHAN_STORAGE_OWNERS:
                _ORPHAN_STORAGE_OWNERS.append(owner)


def _finalize_replay_storage(owner: _ReplayStorageOwner) -> None:
    try:
        owner.close()
    except BaseException:
        _queue_orphan_storage(owner)


atexit.register(_drain_orphan_storages)


def _adopter_owns(
    adopter: Callable[[ReplaySpool], None] | None,
    spool: ReplaySpool,
) -> bool:
    """Confirm an internal adoption without consulting the spool itself."""
    owns = getattr(adopter, "owns", None)
    return bool(owns(spool)) if callable(owns) else False


def _register_adopter_cursor_restore(
    adopter: Callable[[ReplaySpool], None] | None,
    stream: BinaryIO,
    entry: int,
) -> bool:
    """Transfer a cursor obligation to an adopter before moving the stream."""
    register = getattr(adopter, "register_cursor_restore", None)
    if not callable(register):
        return False
    register(stream, entry)
    return True


def _confirm_adopter_cursor_restore(
    adopter: Callable[[ReplaySpool], None] | None,
    stream: BinaryIO,
    entry: int,
) -> None:
    """Tell an adopter that the transferred cursor obligation completed."""
    confirm = getattr(adopter, "confirm_cursor_restore", None)
    if callable(confirm):
        confirm(stream, entry)


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


def _cleanup_takes_precedence(
    error: BaseException,
    primary_error: BaseException,
) -> bool:
    """Return whether source teardown must replace an operation failure."""
    return bool(_contains_process_failure(error, exclude=primary_error))


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
    if _contains_process_failure(cleanup_error, exclude=primary_error):
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
    descriptor_owner: int | None = None
    raw_path: str | None = None
    opened: BinaryIO | None = None
    try:
        descriptor_owner, raw_path = tempfile.mkstemp(
            prefix="messy-xlsx-",
            suffix=".spool",
        )
        os.chmod(raw_path, 0o600)
        opened = cast(BinaryIO, os.fdopen(descriptor_owner, "wb"))
        descriptor_owner = None
        return Path(raw_path), opened
    except BaseException as error:
        if raw_path is None:
            if isinstance(error, OSError):
                raise _storage_error(error) from error
            raise
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
) -> bool:
    try:
        stream.seek(entry)
    except BaseException as restore_error:
        if primary_error is not None:
            if _cleanup_takes_precedence(restore_error, primary_error):
                _mark_fallback_blocked(
                    restore_error,
                    _FallbackBlockReason.SOURCE_OWNERSHIP,
                )
                _attach_operation_failure(restore_error, primary_error)
                _safe_add_note(
                    restore_error,
                    f"source operation also failed: {_type_name(primary_error)}",
                )
                raise
            _mark_fallback_blocked(
                primary_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            _attach_cleanup_failure(primary_error, restore_error)
            _safe_add_note(
                primary_error,
                f"cursor restoration also failed: {_type_name(restore_error)}",
            )
            return False
        _mark_fallback_blocked(
            restore_error,
            _FallbackBlockReason.SOURCE_OWNERSHIP,
        )
        if completed is not None:
            try:
                completed.close()
            except BaseException as cleanup_error:
                if _cleanup_takes_precedence(cleanup_error, restore_error):
                    _mark_fallback_blocked(
                        cleanup_error,
                        _FallbackBlockReason.SOURCE_OWNERSHIP,
                    )
                    _attach_operation_failure(cleanup_error, restore_error)
                    _safe_add_note(
                        cleanup_error,
                        f"cursor restoration also failed: {_type_name(restore_error)}",
                    )
                    raise
                _attach_cleanup_failure(restore_error, cleanup_error)
                _safe_add_note(
                    restore_error,
                    f"temporary spool cleanup also failed: {_type_name(cleanup_error)}",
                )
        raise
    return True


class ReplaySpool:
    """Replay a complete source from bounded memory or a private temp path."""

    def __init__(self, memory: bytes | None, path: Path | None) -> None:
        _drain_orphan_storages()
        self._storage = _ReplayStorageOwner(memory, path)
        self._storage_finalizer = weakref.finalize(
            self,
            _finalize_replay_storage,
            self._storage,
        )

    @property
    def _memory(self) -> bytes | None:
        return self._storage.memory

    @property
    def _path(self) -> Path | None:
        return self._storage.path

    @property
    def _closed(self) -> bool:
        return self._storage.closed

    @classmethod
    def from_stream(
        cls,
        stream: BinaryIO,
        memory_limit: int = DEFAULT_MEMORY_LIMIT,
        *,
        _adopter: Callable[[ReplaySpool], None] | None = None,
    ) -> ReplaySpool:
        """Copy *stream* once, restoring seekable caller cursor state."""
        try:
            seekable, entry = _capture_position(stream)
        except BaseException as error:
            _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            raise
        completed: ReplaySpool | None = None
        adopted = False
        restore_registered = False
        primary_error: BaseException | None = None
        try:
            if seekable:
                restore_registered = _register_adopter_cursor_restore(
                    _adopter,
                    stream,
                    entry,
                )
                stream.seek(0)
            completed = cls._copy_and_adopt(stream, memory_limit, _adopter)
            adopted = _adopter_owns(_adopter, completed)
        except BaseException as error:
            primary_error = error
            _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            if completed is not None:
                adopted = _adopter_owns(_adopter, completed)
                if not adopted:
                    _cleanup_spill(None, completed._path, error)
            raise
        finally:
            if seekable:
                restored = _restore_position(
                    stream,
                    entry,
                    primary_error,
                    None if adopted else completed,
                )
                if restored and restore_registered:
                    _confirm_adopter_cursor_restore(_adopter, stream, entry)

        assert completed is not None
        return completed

    @classmethod
    def _copy_and_adopt(
        cls,
        stream: BinaryIO,
        memory_limit: int,
        adopter: Callable[[ReplaySpool], None] | None,
    ) -> ReplaySpool:
        """Build a complete spool and transfer it before crossing a return gap."""
        buffer = bytearray()
        path: Path | None = None
        opened: BinaryIO | None = None
        completed: ReplaySpool | None = None
        try:
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
            if adopter is not None:
                adopter(completed)
                if not _adopter_owns(adopter, completed):
                    raise RuntimeError("replay spool adopter did not confirm ownership")
            return completed
        except BaseException as error:
            if completed is None or not _adopter_owns(adopter, completed):
                _cleanup_spill(opened, path, error)
            raise

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
        _drain_orphan_storages()
        self._storage.close()
        self._storage_finalizer.detach()

    def _ensure_open(self) -> None:
        if self._closed:
            error = ValueError("ReplaySpool is closed")
            raise _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
