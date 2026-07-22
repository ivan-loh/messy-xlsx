"""Private source and ownership abstraction.

``SourceHandle`` gives the parsing pipeline one repeatable view of a path or a
caller-owned binary stream. It never retains an unbounded complete byte cache;
replay-required sources use a bounded-memory, spillable spool.
"""

from __future__ import annotations

import io
import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import BinaryIO, Literal, TypeAlias, cast

from messy_xlsx._fallback_signals import (
    _attach_cleanup_failure,
    _attach_operation_failure,
    _contains_process_failure,
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

    def __init__(self, source: SourceInput, filename: str | None = None) -> None:
        self._path: Path | None = None
        self._stream: BinaryIO | None = None
        self._spool: ReplaySpool | None = None
        self._backend_requires_copy: bool | None = None
        self._active_borrow = False
        self._closed = False

        if _is_readable(source):
            stream = cast(BinaryIO, source)
            self._stream = stream
            self._original: BackendSource = stream
            self._filename_hint = filename or _stream_name(stream)
            self._stream_is_seekable = self._probe_seekability(stream)
            self._identity: SourceIdentity = ("stream", id(stream))

            if not self._stream_is_seekable:
                self._ensure_spool()
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
        if self._spool is not None:
            self._spool.close()
            self._spool = None
        self._closed = True
        self._backend_requires_copy = None

    def _ensure_spool(self) -> ReplaySpool:
        if self._spool is not None:
            return self._spool
        try:
            spool = ReplaySpool.from_stream(self._require_stream())
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
        self._spool = spool
        return spool

    @contextmanager
    def _borrow(self) -> Iterator[None]:
        self._ensure_open()
        if self._active_borrow:
            error = RuntimeError("SourceHandle already has an active borrow")
            raise _mark_fallback_blocked(
                error,
                _FallbackBlockReason.SOURCE_OWNERSHIP,
            )
        self._active_borrow = True
        try:
            yield
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
        entry_position = caller_stream.tell()
        consumer_error: BaseException | None = None
        try:
            caller_stream.seek(0)
            yield caller_stream
        except BaseException as error:
            consumer_error = error
            raise
        finally:
            try:
                caller_stream.seek(entry_position)
            except BaseException as restore_error:
                if consumer_error is None or _cleanup_takes_precedence(restore_error):
                    _mark_fallback_blocked(
                        restore_error,
                        _FallbackBlockReason.SOURCE_OWNERSHIP,
                    )
                    raise
                _mark_fallback_blocked(
                    consumer_error,
                    _FallbackBlockReason.SOURCE_OWNERSHIP,
                )

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
            if _contains_process_failure(cleanup_error):
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


def _cleanup_takes_precedence(error: BaseException) -> bool:
    """Return whether source teardown must replace an operation failure."""
    return _contains_process_failure(error)
