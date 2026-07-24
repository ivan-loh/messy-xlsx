"""Private source and ownership abstraction.

``SourceHandle`` gives the parsing pipeline one repeatable view of a path or a
caller-owned binary stream. It never retains an unbounded complete byte cache;
replay-required sources use a bounded-memory, spillable spool.
"""

from __future__ import annotations

import atexit
import io
import os
import weakref
from collections import deque
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from pathlib import Path
from types import TracebackType
from typing import BinaryIO, Literal, TypeAlias, cast

from messy_xlsx._fallback_signals import (
    _attach_cleanup_failure,
    _attach_operation_failure,
    _contains_process_failure,
    _exception_traceback,
    _FallbackBlockReason,
    _mark_fallback_blocked,
    _safe_add_note,
    _type_name,
)
from messy_xlsx._spool import ReplaySpool, _SpoolStorageError
from messy_xlsx.exceptions import FileError

SourceInput: TypeAlias = str | Path | BinaryIO
SourceIdentity: TypeAlias = tuple[Literal["path", "stream"], str | int]
BackendSource: TypeAlias = Path | BinaryIO
_MAX_ORPHAN_SOURCE_STATES = 64


class _SourceCleanupState:
    """Finalizable cursor/spool obligations independent of ``SourceHandle``."""

    __slots__ = ("caller_stream", "pending_restore_position", "spool")

    def __init__(self) -> None:
        self.caller_stream: BinaryIO | None = None
        self.pending_restore_position: int | None = None
        self.spool: ReplaySpool | None = None

    def close_once(self) -> None:
        failures: list[BaseException] = []
        if self.pending_restore_position is not None and self.caller_stream is not None:
            try:
                self.caller_stream.seek(self.pending_restore_position)
            except BaseException as error:
                failures.append(error)
            else:
                self.pending_restore_position = None
        spool = self.spool
        if spool is not None:
            try:
                spool.close()
            except BaseException as error:
                failures.append(error)
            else:
                self.spool = None
        if failures:
            process_failure = next(
                (error for error in failures if _contains_process_failure(error)),
                None,
            )
            raise process_failure if process_failure is not None else failures[0]

    @property
    def pending(self) -> bool:
        return self.pending_restore_position is not None or self.spool is not None


_ORPHAN_SOURCE_STATES: deque[_SourceCleanupState] = deque()


def _queue_orphan_source_state(state: _SourceCleanupState) -> None:
    if not state.pending or any(candidate is state for candidate in _ORPHAN_SOURCE_STATES):
        return
    if len(_ORPHAN_SOURCE_STATES) >= _MAX_ORPHAN_SOURCE_STATES:
        _drain_orphan_source_states()
    if len(_ORPHAN_SOURCE_STATES) < _MAX_ORPHAN_SOURCE_STATES:
        _ORPHAN_SOURCE_STATES.append(state)


def _drain_orphan_source_states() -> None:
    pending = tuple(_ORPHAN_SOURCE_STATES)
    _ORPHAN_SOURCE_STATES.clear()
    for state in pending:
        try:
            state.close_once()
        except BaseException:
            if len(_ORPHAN_SOURCE_STATES) < _MAX_ORPHAN_SOURCE_STATES:
                _ORPHAN_SOURCE_STATES.append(state)


def _finalize_source_state(state: _SourceCleanupState) -> None:
    try:
        state.close_once()
    except BaseException:
        _queue_orphan_source_state(state)


atexit.register(_drain_orphan_source_states)


class _SourceSpoolAdopter:
    """Expose both transfer and hook-free ownership confirmation."""

    __slots__ = ("_handle",)

    def __init__(self, handle: SourceHandle) -> None:
        self._handle = handle

    def __call__(self, spool: ReplaySpool) -> None:
        self._handle._adopt_spool(spool)

    def owns(self, spool: ReplaySpool) -> bool:
        return self._handle._spool is spool

    def register_cursor_restore(self, stream: BinaryIO, position: int) -> None:
        self._handle._register_pending_cursor_restoration(stream, position)

    def confirm_cursor_restore(self, stream: BinaryIO, position: int) -> None:
        self._handle._confirm_pending_cursor_restoration(stream, position)


def _run_source_cleanups(
    cleanups: list[tuple[str, Callable[[], object]]],
    *,
    primary_error: BaseException,
    primary_traceback: TracebackType | None,
) -> None:
    """Load the shared lifecycle helper lazily to avoid a parsing import cycle."""
    from messy_xlsx.parsing.streams import _run_cleanups

    _run_cleanups(
        cleanups,
        primary_error=primary_error,
        primary_traceback=primary_traceback,
    )


def _is_readable(value: object) -> bool:
    """Return whether *value* exposes a callable binary-style ``read`` method."""
    try:
        return callable(getattr(value, "read", None))
    except Exception:
        return False


def _stream_name(stream: BinaryIO) -> str | None:
    """Return a useful filename carried by a named stream, if it has one."""
    try:
        name = getattr(stream, "name", None)
    except Exception:
        return None

    if isinstance(name, str):
        return name or None
    if isinstance(name, os.PathLike):
        try:
            path_name = os.fspath(name)
        except Exception:
            return None
        return path_name if isinstance(path_name, str) and path_name else None
    return None


def describe_source(source: SourceInput, filename: str | None = None) -> str:
    """Describe a raw source even when constructing its handle fails."""
    if _is_readable(source):
        return filename or _stream_name(cast(BinaryIO, source)) or "<stream>"
    try:
        return str(Path(cast(str | Path, source)))
    except Exception:
        return "<source>"


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


class SourceHandle:
    """Repeatable, caller-ownership-preserving access to one input source.

    Paths are kept as paths for backends that can open them directly. Seekable
    caller streams are borrowed at byte zero and restored to their entry
    position after every borrow. Non-seekable streams are consumed exactly once
    into an owned replay spool and are never closed by the handle.
    """

    _cleanup_state: _SourceCleanupState
    _cleanup_finalizer: weakref.finalize

    def __new__(cls, *_args: object, **_kwargs: object) -> SourceHandle:
        handle = super().__new__(cls)
        state = _SourceCleanupState()
        handle._cleanup_state = state
        handle._cleanup_finalizer = weakref.finalize(
            handle,
            _finalize_source_state,
            state,
        )
        _drain_orphan_source_states()
        return handle

    def __init__(self, source: SourceInput, filename: str | None = None) -> None:
        self._initialize(source, filename)
        try:
            self.start()
        except BaseException as error:
            _run_source_cleanups(
                [
                    ("source construction rollback", self._close_spool),
                    ("source construction rollback retry", self._close_spool),
                ],
                primary_error=error,
                primary_traceback=_exception_traceback(error),
            )
            raise

    @classmethod
    def prepare(cls, source: SourceInput, filename: str | None = None) -> SourceHandle:
        """Create an inert handle whose eager acquisition has not started."""
        handle = cls.__new__(cls)
        handle._initialize(source, filename)
        return handle

    def _initialize(self, source: SourceInput, filename: str | None) -> None:
        self._path: Path | None = None
        self._stream: BinaryIO | None = None
        self._spool: ReplaySpool | None = None
        self._backend_requires_copy: bool | None = None
        self._active_borrow = False
        self._pending_restore_position: int | None = None
        self._closed = False
        self._started = False

        if _is_readable(source):
            stream = cast(BinaryIO, source)
            self._stream = stream
            self._cleanup_state.caller_stream = stream
            self._original: BackendSource = stream
            self._filename_hint = filename or _stream_name(stream)
            self._stream_is_seekable = self._probe_seekability(stream)
            self._identity: SourceIdentity = ("stream", id(stream))
        else:
            path = Path(cast(str | Path, source))
            self._path = path
            self._original = path
            self._filename_hint = filename or None
            self._stream_is_seekable = False
            try:
                resolved = path.resolve(strict=False)
            except (OSError, RuntimeError):
                resolved = path.absolute()
            self._identity = ("path", str(resolved))

    @property
    def _spool(self) -> ReplaySpool | None:
        return self._cleanup_state.spool

    @_spool.setter
    def _spool(self, value: ReplaySpool | None) -> None:
        self._cleanup_state.spool = value

    @property
    def _pending_restore_position(self) -> int | None:
        return self._cleanup_state.pending_restore_position

    @_pending_restore_position.setter
    def _pending_restore_position(self, value: int | None) -> None:
        self._cleanup_state.pending_restore_position = value

    def start(self) -> None:
        """Perform eager replay acquisition after an owner records this handle."""
        self._ensure_open()
        if self._started:
            return
        try:
            if self._stream is not None and not self._stream_is_seekable:
                self._ensure_spool()
            self._started = True
        except BaseException as error:
            _run_source_cleanups(
                [("source start rollback", self._close_spool)],
                primary_error=error,
                primary_traceback=_exception_traceback(error),
            )
            raise

    @classmethod
    def coerce(
        cls,
        source: SourceInput | SourceHandle,
        filename: str | None = None,
    ) -> SourceHandle:
        """Return *source* unchanged when it is already a handle, else wrap it.

        A handle's filename metadata is fixed at its first public boundary. A
        later adapter therefore cannot accidentally replace an explicit hint.
        """
        if isinstance(source, cls):
            return source
        return cls(cast(SourceInput, source), filename=filename)

    @staticmethod
    def _probe_seekability(stream: BinaryIO) -> bool:
        """Probe seek/tell behavior without trusting attribute presence alone."""
        try:
            seekable = getattr(stream, "seekable", None)
        except Exception:
            return False

        if callable(seekable):
            try:
                if not seekable():
                    return False
            except Exception:
                return False

        try:
            tell = stream.tell
            seek = stream.seek
        except Exception:
            return False
        if not callable(tell) or not callable(seek):
            return False

        try:
            position = tell()
            seek(position)
        except Exception:
            return False
        return True

    @property
    def original(self) -> BackendSource:
        """Original caller stream, or the normalized ``Path`` for path input."""
        return self._original

    @property
    def path(self) -> Path | None:
        """Filesystem path when the source was path-based, otherwise ``None``."""
        return self._path

    @property
    def filename_hint(self) -> str | None:
        """Explicit filename, or a named stream's filename when available."""
        return self._filename_hint

    @property
    def description(self) -> str:
        """Stable human-readable source description for diagnostics."""
        if self._path is not None:
            return str(self._path)
        return self._filename_hint or "<stream>"

    @property
    def identity(self) -> SourceIdentity:
        """Stable, hashable identity for this path or original stream object."""
        return self._identity

    @property
    def is_path(self) -> bool:
        """Whether this source originated from a filesystem path."""
        return self._path is not None

    @property
    def is_stream(self) -> bool:
        """Whether this source originated from a caller-owned stream."""
        return self._stream is not None

    @property
    def is_seekable(self) -> bool:
        """Whether the original source can be replayed without a spool."""
        return self._path is not None or self._stream_is_seekable

    @property
    def was_snapshotted(self) -> bool:
        """Whether a non-seekable original required a one-time replay copy."""
        return not self._stream_is_seekable and self._spool is not None

    @property
    def owns_stream(self) -> bool:
        """Whether the handle owns the original stream (always false)."""
        return False

    @property
    def caller_owned(self) -> bool:
        """Whether the original source is a stream owned by the caller."""
        return self._stream is not None

    @property
    def closed(self) -> bool:
        """Whether this handle has released its internal replay spool."""
        return self._closed

    def rewind(self) -> None:
        """Move a seekable caller stream to byte zero.

        Adapters should normally prefer :meth:`open_binary`, which also restores
        the entry position. Paths and snapshotted streams have no shared cursor,
        so rewinding them is a no-op.
        """
        self._ensure_open()
        if self._path is not None or not self._stream_is_seekable:
            return
        stream = self._require_stream()
        stream.seek(0)

    def read_bytes(self) -> bytes:
        """Return complete bytes without retaining an unbounded handle cache."""
        self._ensure_open()
        with self.open_binary() as stream:
            return _coerce_bytes(stream.read())

    def detached_binary(self) -> BinaryIO:
        """Return an owned seekable copy for a backend that outlives a borrow.

        The caller must close the returned stream. Ordinary synchronous
        adapters should prefer :meth:`open_binary` to avoid this extra view.
        """
        return io.BytesIO(self.read_bytes())

    @contextmanager
    def open_binary(self) -> Iterator[BinaryIO]:
        """Borrow a complete seekable binary view and clean it up safely.

        The context owns path-opened files and replay views. It never closes a
        caller stream, and it restores a seekable caller stream even when the
        consumer raises.
        """
        with self._borrow(), self._open_binary_unchecked() as stream:
            yield stream

    @contextmanager
    def open_path_or_bytes(self) -> Iterator[Path | bytes]:
        """Yield a path or bounded-memory bytes for path/bytes-only backends."""
        with self._borrow():
            if self._path is not None:
                yield self._path
                return
            spool = self._ensure_spool()
            with spool.open_path_or_bytes() as source:
                yield source

    @contextmanager
    def open_legacy(self) -> Iterator[BackendSource]:
        """Yield the ordinary path/stream type expected by legacy extensions."""
        with self._borrow():
            if self._path is not None:
                yield self._path
                return
            with self._open_binary_unchecked() as stream:
                yield stream

    @contextmanager
    def open_backend(self) -> Iterator[BackendSource]:
        """Yield a path or a byte-normalized stream suitable for parsers.

        A few valid binary stream implementations return ``bytearray`` or
        ``memoryview`` from ``read()``. The source boundary accepts those
        values, but third-party parsers generally require actual ``bytes``.
        Detect that limitation once and provide an owned normalized view only
        for those adapters.
        """
        with self._borrow():
            if self._path is not None:
                yield self._path
                return

            if not self._stream_is_seekable:
                with self._ensure_spool().open_binary() as stream:
                    yield stream
                return

            if self._backend_requires_copy is None:
                with self._open_binary_unchecked() as stream:
                    probe = stream.read(1)
                _coerce_bytes(probe)
                self._backend_requires_copy = not isinstance(probe, bytes)

            if self._backend_requires_copy:
                with self._ensure_spool().open_binary() as normalized_stream:
                    yield normalized_stream
                return

            with self._open_binary_unchecked() as stream:
                yield stream

    def close(self) -> None:
        """Release handle-owned replay storage; never close caller streams."""
        if self._closed:
            return
        cleanups: list[tuple[str, Callable[[], object]]] = []
        if self._pending_restore_position is not None:
            cleanups.append(("caller cursor restoration", self._retry_pending_cursor_restoration))
        if self._spool is not None:
            cleanups.append(("temporary spool cleanup", self._close_spool))
        if cleanups:
            from messy_xlsx.parsing.streams import _run_cleanups

            _run_cleanups(cleanups)
        if self._pending_restore_position is not None or self._spool is not None:
            return
        self._closed = True
        self._backend_requires_copy = None
        self._cleanup_finalizer.detach()

    def _ensure_spool(self) -> ReplaySpool:
        if self._spool is not None:
            return self._spool
        try:
            adopter = _SourceSpoolAdopter(self)
            spool = ReplaySpool.from_stream(
                self._require_stream(),
                _adopter=adopter,
            )
            self._adopt_spool(spool)
            return spool
        except _SpoolStorageError as error:
            source_error = FileError(
                f"Cannot spool source: {error}",
                file_path=self.description,
                operation="spool",
            )
            raise _mark_fallback_blocked(
                source_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            ) from error
        except BaseException as error:
            owned_spool = self._spool
            if owned_spool is None:
                raise
            _run_source_cleanups(
                [("temporary spool cleanup", self._close_spool)],
                primary_error=error,
                primary_traceback=_exception_traceback(error),
            )
            raise

    def _adopt_spool(self, spool: ReplaySpool) -> None:
        """Record a completed spool before its constructor restores or returns."""
        current = self._spool
        if current is not None and current is not spool:
            raise RuntimeError("SourceHandle already owns a different replay spool")
        self._spool = spool

    def _close_spool(self) -> None:
        spool = self._spool
        if spool is None:
            return
        try:
            spool.close()
        except BaseException:
            raise
        self._spool = None

    @contextmanager
    def _borrow(self) -> Iterator[None]:
        self._ensure_open()
        self._retry_pending_cursor_restoration()
        if self._active_borrow:
            error = RuntimeError("SourceHandle already has an active borrow")
            raise _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
        try:
            try:
                self._active_borrow = True
                yield
            finally:
                self._active_borrow = False
        finally:
            self._active_borrow = False

    @contextmanager
    def _open_binary_unchecked(self) -> Iterator[BinaryIO]:
        if self._path is not None:
            with self._path.open("rb") as opened_stream:
                yield opened_stream
            return
        if not self._stream_is_seekable:
            with self._ensure_spool().open_binary() as replay:
                yield replay
            return

        caller_stream = self._require_stream()
        try:
            entry_position = caller_stream.tell()
        except BaseException as source_error:
            _mark_fallback_blocked(
                source_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            raise
        consumer_error: BaseException | None = None
        try:
            try:
                caller_stream.seek(0)
            except BaseException as source_error:
                _mark_fallback_blocked(
                    source_error,
                    _FallbackBlockReason.SOURCE_OWNERSHIP,
                )
                raise
            yield caller_stream
        except BaseException as error:
            consumer_error = error
            raise
        finally:
            self._restore_caller_cursor(
                caller_stream,
                entry_position,
                consumer_error,
            )

    def _restore_caller_cursor(
        self,
        caller_stream: BinaryIO,
        entry_position: int,
        consumer_error: BaseException | None,
    ) -> None:
        self._pending_restore_position = entry_position
        try:
            caller_stream.seek(entry_position)
        except BaseException as restore_error:
            if consumer_error is None:
                _mark_fallback_blocked(
                    restore_error,
                    _FallbackBlockReason.SOURCE_OWNERSHIP,
                )
                raise
            if _cleanup_takes_precedence(restore_error, consumer_error):
                _mark_fallback_blocked(
                    restore_error,
                    _FallbackBlockReason.SOURCE_OWNERSHIP,
                )
                _attach_operation_failure(restore_error, consumer_error)
                _safe_add_note(
                    restore_error,
                    f"source operation also failed: {_type_name(consumer_error)}",
                )
                raise
            _mark_fallback_blocked(
                consumer_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            _attach_cleanup_failure(consumer_error, restore_error)
            _safe_add_note(
                consumer_error,
                f"cursor restoration also failed: {_type_name(restore_error)}",
            )
        else:
            self._pending_restore_position = None

    def _register_pending_cursor_restoration(
        self,
        caller_stream: BinaryIO,
        position: int,
    ) -> None:
        """Own an exact caller cursor restoration before the stream is moved."""
        if caller_stream is not self._require_stream():
            raise RuntimeError("Cursor restoration belongs to a different source")
        pending = self._pending_restore_position
        if pending is not None and pending != position:
            raise RuntimeError("SourceHandle already owns a different cursor restoration")
        self._pending_restore_position = position

    def _confirm_pending_cursor_restoration(
        self,
        caller_stream: BinaryIO,
        position: int,
    ) -> None:
        """Clear one exact cursor obligation only after restoration succeeds."""
        if caller_stream is not self._require_stream():
            raise RuntimeError("Cursor restoration belongs to a different source")
        if self._pending_restore_position == position:
            self._pending_restore_position = None

    def _retry_pending_cursor_restoration(self) -> None:
        """Retry one interrupted caller cursor restoration."""
        position = self._pending_restore_position
        if position is None:
            return
        stream = self._require_stream()
        try:
            stream.seek(position)
        except BaseException as restore_error:
            _mark_fallback_blocked(
                restore_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            raise
        self._pending_restore_position = None

    def _ensure_open(self) -> None:
        if self._closed:
            error = ValueError("SourceHandle is closed")
            raise _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )

    def _require_stream(self) -> BinaryIO:
        if self._stream is None:
            error = ValueError("SourceHandle does not contain a stream")
            raise _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
        return self._stream

    def __enter__(self) -> SourceHandle:
        self._ensure_open()
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        del exc_type, traceback
        try:
            self.close()
        except BaseException as cleanup_error:
            _mark_fallback_blocked(
                cleanup_error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            _attach_cleanup_failure(cleanup_error, cleanup_error)
            if not isinstance(exc_value, BaseException):
                _safe_add_note(
                    cleanup_error,
                    f"source cleanup failed: {_type_name(cleanup_error)}",
                )
                raise
            if _contains_process_failure(cleanup_error, exclude=exc_value):
                _attach_operation_failure(cleanup_error, exc_value)
                _safe_add_note(
                    cleanup_error,
                    f"source operation also failed: {_type_name(exc_value)}",
                )
                raise
            _mark_fallback_blocked(
                exc_value,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
            _attach_cleanup_failure(exc_value, cleanup_error)
            _safe_add_note(
                exc_value,
                f"source cleanup also failed: {_type_name(cleanup_error)}",
            )


def _cleanup_takes_precedence(
    error: BaseException,
    primary_error: BaseException,
) -> bool:
    """Return whether source teardown must replace an operation failure."""
    return bool(_contains_process_failure(error, exclude=primary_error))
