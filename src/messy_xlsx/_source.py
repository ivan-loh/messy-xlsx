"""Private source and ownership abstraction.

``SourceHandle`` gives the parsing pipeline one repeatable view of a path or a
caller-owned binary stream.  It deliberately does not cache parser backends or
workbook objects; its only persistent payload cache is an immutable byte view
requested by an adapter, plus the mandatory snapshot for a read-once stream.
"""

from __future__ import annotations

import io
import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import BinaryIO, Literal, TypeAlias, cast

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
    raise TypeError(
        "Binary source read() must return bytes, bytearray, or memoryview; "
        f"got {type(value).__name__}"
    )


class SourceHandle:
    """Repeatable, caller-ownership-preserving access to one input source.

    Paths are kept as paths for backends that can open them directly. Seekable
    caller streams are borrowed at byte zero and restored to their entry
    position after every borrow. Non-seekable streams are consumed exactly once
    into an internal immutable snapshot and are never closed by the handle.
    """

    def __init__(self, source: SourceInput, filename: str | None = None) -> None:
        self._path: Path | None = None
        self._stream: BinaryIO | None = None
        self._snapshot: bytes | None = None
        self._byte_cache: bytes | None = None
        self._has_byte_cache = False
        self._backend_requires_copy: bool | None = None
        self._closed = False

        if _is_readable(source):
            stream = cast(BinaryIO, source)
            self._stream = stream
            self._original: BackendSource = stream
            self._filename_hint = filename or _stream_name(stream)
            self._stream_is_seekable = self._probe_seekability(stream)
            self._identity: SourceIdentity = ("stream", id(stream))

            if not self._stream_is_seekable:
                self._snapshot_nonseekable()
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

    def _snapshot_nonseekable(self) -> None:
        """Consume a read-once stream into the handle's canonical byte view."""
        stream = self._require_stream()

        # If a read-once stream can report that a prefix is already gone, do not
        # silently interpret the remaining suffix as a complete workbook.
        try:
            tell = getattr(stream, "tell", None)
        except Exception:
            tell = None
        if callable(tell):
            try:
                position = tell()
            except Exception:
                position = None
            if position not in (None, 0):
                raise ValueError(
                    "A non-seekable binary source must be positioned at byte 0 when it is supplied"
                )

        content = _coerce_bytes(stream.read())
        self._snapshot = content
        self._byte_cache = content
        self._has_byte_cache = True

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
        """Whether the original source can be replayed without a snapshot."""
        return self._path is not None or self._stream_is_seekable

    @property
    def was_snapshotted(self) -> bool:
        """Whether a non-seekable original required a one-time snapshot."""
        return self._snapshot is not None

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
        """Whether this handle has released its internal snapshot/cache."""
        return self._closed

    def rewind(self) -> None:
        """Move a seekable caller stream to byte zero.

        Adapters should normally prefer :meth:`open_binary`, which also restores
        the entry position. Paths and snapshotted streams have no shared cursor,
        so rewinding them is a no-op.
        """
        self._ensure_open()
        if self._path is not None or self._snapshot is not None:
            return
        stream = self._require_stream()
        stream.seek(0)

    def read_bytes(self) -> bytes:
        """Return one memoized immutable byte view of the complete source."""
        self._ensure_open()
        if self._has_byte_cache:
            assert self._byte_cache is not None
            return self._byte_cache

        with self.open_binary() as stream:
            content = _coerce_bytes(stream.read())

        self._byte_cache = content
        self._has_byte_cache = True
        return content

    def detached_binary(self) -> BinaryIO:
        """Return an owned seekable copy for a backend that outlives a borrow.

        The caller must close the returned stream. Ordinary synchronous
        adapters should prefer :meth:`open_binary` to avoid this extra view.
        """
        return io.BytesIO(self.read_bytes())

    @contextmanager
    def open_binary(self) -> Iterator[BinaryIO]:
        """Borrow a complete seekable binary view and clean it up safely.

        The context owns path-opened files and snapshot views. It never closes a
        caller stream, and it restores a seekable caller stream even when the
        consumer raises.
        """
        self._ensure_open()

        if self._path is not None:
            with self._path.open("rb") as opened_stream:
                yield opened_stream
            return

        if self._snapshot is not None:
            with io.BytesIO(self._snapshot) as snapshot_stream:
                yield snapshot_stream
            return

        caller_stream = self._require_stream()
        entry_position = caller_stream.tell()
        consumer_failed = False
        try:
            caller_stream.seek(0)
            yield caller_stream
        except BaseException:
            consumer_failed = True
            raise
        finally:
            try:
                caller_stream.seek(entry_position)
            except Exception:
                # Keep the consumer's exception as the primary failure. A
                # restoration error after successful consumption is actionable.
                if not consumer_failed:
                    raise

    @contextmanager
    def open_legacy(self) -> Iterator[BackendSource]:
        """Yield the ordinary path/stream type expected by legacy extensions."""
        self._ensure_open()
        if self._path is not None:
            yield self._path
            return
        with self.open_binary() as stream:
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
        self._ensure_open()
        if self._path is not None:
            yield self._path
            return

        if self._snapshot is not None:
            with self.open_binary() as stream:
                yield stream
            return

        if self._backend_requires_copy is None:
            with self.open_binary() as stream:
                probe = stream.read(1)
                stream.seek(0)
            _coerce_bytes(probe)
            self._backend_requires_copy = not isinstance(probe, bytes)

        if self._backend_requires_copy:
            with io.BytesIO(self.read_bytes()) as normalized_stream:
                yield normalized_stream
            return

        with self.open_binary() as stream:
            yield stream

    def close(self) -> None:
        """Release only handle-owned memory; never close the caller's stream."""
        if self._closed:
            return
        self._closed = True
        self._snapshot = None
        self._byte_cache = None
        self._has_byte_cache = False
        self._backend_requires_copy = None

    def _ensure_open(self) -> None:
        if self._closed:
            raise ValueError("SourceHandle is closed")

    def _require_stream(self) -> BinaryIO:
        if self._stream is None:
            raise ValueError("SourceHandle does not contain a stream")
        return self._stream

    def __enter__(self) -> SourceHandle:
        self._ensure_open()
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        self.close()
