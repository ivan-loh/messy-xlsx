from __future__ import annotations

import io
import os
from pathlib import Path
from typing import BinaryIO, NoReturn

import pytest

import messy_xlsx._spool as spool_module
from messy_xlsx._fallback_signals import (
    _fallback_block_reason,
    _FallbackBlockReason,
)
from messy_xlsx._spool import ReplaySpool


class _HostilePrimaryError(RuntimeError):
    def __getattribute__(self, name: str) -> object:
        if name == "add_note":
            raise AssertionError("cleanup diagnostics must bypass add_note lookup")
        return BaseException.__getattribute__(self, name)

    def add_note(self, note: str) -> None:
        del note
        raise AssertionError("cleanup diagnostics must bypass add_note overrides")


class _HostileCleanupError(RuntimeError):
    def __repr__(self) -> str:
        raise AssertionError("cleanup diagnostics must tolerate hostile repr")


class _FatalSpoolCleanup(BaseException):
    pass


def _descriptor_is_closed(descriptor: int) -> bool:
    try:
        os.fstat(descriptor)
    except OSError:
        return True
    return False


def _remove_test_resources(descriptor: int, path: Path) -> None:
    try:
        os.close(descriptor)
    except OSError:
        pass
    path.unlink(missing_ok=True)


def _track_temp_creation(
    monkeypatch: pytest.MonkeyPatch,
) -> list[tuple[int, Path]]:
    created: list[tuple[int, Path]] = []
    original_mkstemp = spool_module.tempfile.mkstemp

    def tracking_mkstemp(*args: object, **kwargs: object) -> tuple[int, str]:
        descriptor, raw_path = original_mkstemp(*args, **kwargs)
        created.append((descriptor, Path(raw_path)))
        return descriptor, raw_path

    monkeypatch.setattr(spool_module.tempfile, "mkstemp", tracking_mkstemp)
    return created


def test_small_spool_stays_in_memory_and_restores_cursor() -> None:
    source = io.BytesIO(b"abcdef")
    source.seek(3)

    spool = ReplaySpool.from_stream(source, memory_limit=16)

    assert source.tell() == 3
    with spool.open_path_or_bytes() as backend:
        assert backend == b"abcdef"
    spool.close()


def test_large_spool_uses_private_path_and_deletes_it() -> None:
    spool = ReplaySpool.from_stream(io.BytesIO(b"x" * 32), memory_limit=8)

    with spool.open_path_or_bytes() as backend:
        assert isinstance(backend, Path)
        path = backend
        assert path.exists()
        assert path.stat().st_mode & 0o777 == 0o600

    spool.close()
    assert not path.exists()


def test_open_binary_returns_fresh_complete_replays() -> None:
    spool = ReplaySpool.from_stream(io.BytesIO(b"complete"), memory_limit=4)

    with spool.open_binary() as first:
        assert first.read() == b"complete"
    with spool.open_binary() as second:
        assert second.read() == b"complete"

    spool.close()


def test_close_is_idempotent() -> None:
    spool = ReplaySpool.from_stream(io.BytesIO(b"data"), memory_limit=8)

    spool.close()
    spool.close()

    with (
        pytest.raises(ValueError, match="ReplaySpool is closed"),
        spool.open_binary(),
    ):
        pass


def test_spool_restores_cursor_when_read_fails() -> None:
    class Broken(io.BytesIO):
        def read(self, size: int = -1) -> bytes:
            raise OSError("injected read failure")

    source = Broken(b"data")
    source.seek(2)

    with pytest.raises(OSError, match="injected read failure"):
        ReplaySpool.from_stream(source, memory_limit=8)

    assert source.tell() == 2


@pytest.mark.parametrize(
    "source_error",
    [
        OSError("initial rewind failed"),
        ExceptionGroup("outer", [OSError("nested initial rewind failed")]),
    ],
)
def test_initial_source_seek_failure_is_marked_without_replacement(
    source_error: BaseException,
) -> None:
    class InitialRewindFailure(io.BytesIO):
        def __init__(self) -> None:
            super().__init__(b"source")
            self.seek_calls = 0

        def seek(self, position: int, whence: int = 0) -> int:
            self.seek_calls += 1
            if self.seek_calls == 2 and position == 0 and whence == 0:
                raise source_error
            return super().seek(position, whence)

    source = InitialRewindFailure()
    io.BytesIO.seek(source, 3)

    with pytest.raises(type(source_error)) as captured:
        ReplaySpool.from_stream(source)

    assert captured.value is source_error
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP


def test_spill_file_is_removed_when_a_later_source_read_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created_paths: list[Path] = []
    original_mkstemp = spool_module.tempfile.mkstemp

    def tracking_mkstemp(*args: object, **kwargs: object) -> tuple[int, str]:
        descriptor, raw_path = original_mkstemp(*args, **kwargs)
        created_paths.append(Path(raw_path))
        return descriptor, raw_path

    class BrokenAfterSpill(io.BytesIO):
        def __init__(self) -> None:
            super().__init__(b"x" * 32)
            self._read_calls = 0

        def read(self, size: int = -1) -> bytes:
            self._read_calls += 1
            if self._read_calls == 2:
                raise OSError("source failed after spill")
            return super().read(size)

    monkeypatch.setattr(spool_module.tempfile, "mkstemp", tracking_mkstemp)
    source = BrokenAfterSpill()
    source.seek(3)

    with pytest.raises(OSError, match="source failed after spill"):
        ReplaySpool.from_stream(source, memory_limit=8)

    assert source.tell() == 3
    assert len(created_paths) == 1
    assert not created_paths[0].exists()


def test_spill_file_is_removed_when_a_write_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created_paths: list[Path] = []
    original_mkstemp = spool_module.tempfile.mkstemp
    original_fdopen = os.fdopen

    def tracking_mkstemp(*args: object, **kwargs: object) -> tuple[int, str]:
        descriptor, raw_path = original_mkstemp(*args, **kwargs)
        created_paths.append(Path(raw_path))
        return descriptor, raw_path

    class BrokenWriter:
        def __init__(self, wrapped: BinaryIO) -> None:
            self._wrapped = wrapped

        def write(self, _content: bytes | bytearray) -> int:
            raise OSError("injected capacity failure")

        def close(self) -> None:
            self._wrapped.close()

    def broken_fdopen(descriptor: int, mode: str) -> BrokenWriter:
        return BrokenWriter(original_fdopen(descriptor, mode))

    monkeypatch.setattr(spool_module.tempfile, "mkstemp", tracking_mkstemp)
    monkeypatch.setattr(spool_module.os, "fdopen", broken_fdopen)
    source = io.BytesIO(b"x" * 32)
    source.seek(4)

    with pytest.raises(OSError, match="injected capacity failure"):
        ReplaySpool.from_stream(source, memory_limit=8)

    assert source.tell() == 4
    assert len(created_paths) == 1
    assert not created_paths[0].exists()


@pytest.mark.parametrize("operation", ["chmod", "fdopen"])
def test_spill_resources_are_removed_when_setup_raises_non_oserror(
    monkeypatch: pytest.MonkeyPatch,
    operation: str,
) -> None:
    created = _track_temp_creation(monkeypatch)
    primary_error = RuntimeError(f"injected {operation} failure")

    def fail_operation(*_args: object) -> NoReturn:
        raise primary_error

    monkeypatch.setattr(spool_module.os, operation, fail_operation)
    source = io.BytesIO(b"x" * 32)
    source.seek(5)

    try:
        with pytest.raises(RuntimeError, match=f"injected {operation} failure") as captured:
            ReplaySpool.from_stream(source, memory_limit=8)

        assert captured.value is primary_error
        assert source.tell() == 5
        assert len(created) == 1
        descriptor, path = created[0]
        assert _descriptor_is_closed(descriptor)
        assert not path.exists()
    finally:
        if created:
            _remove_test_resources(*created[0])


def test_non_oserror_remains_primary_when_spill_cleanup_also_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created = _track_temp_creation(monkeypatch)
    original_close = os.close
    original_unlink = os.unlink
    primary_error = RuntimeError("injected fdopen failure")

    def fail_fdopen(_descriptor: int, _mode: str) -> BinaryIO:
        raise primary_error

    def fail_close(_descriptor: int) -> None:
        raise RuntimeError("injected descriptor cleanup failure")

    def fail_unlink(_path: str) -> None:
        raise RuntimeError("injected path cleanup failure")

    monkeypatch.setattr(spool_module.os, "fdopen", fail_fdopen)
    monkeypatch.setattr(spool_module.os, "close", fail_close)
    monkeypatch.setattr(spool_module.os, "unlink", fail_unlink)

    try:
        with pytest.raises(RuntimeError, match="injected fdopen failure") as captured:
            ReplaySpool.from_stream(io.BytesIO(b"x" * 32), memory_limit=8)

        assert captured.value is primary_error
        assert captured.value.__notes__ == [
            "temporary descriptor close also failed: RuntimeError",
            "temporary file removal also failed: RuntimeError",
        ]
        assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP
    finally:
        if created:
            descriptor, path = created[0]
            try:
                original_close(descriptor)
            except OSError:
                pass
            try:
                original_unlink(path)
            except FileNotFoundError:
                pass


def test_hostile_spill_cleanup_diagnostics_never_mask_the_primary(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created = _track_temp_creation(monkeypatch)
    original_close = os.close
    original_unlink = os.unlink
    primary_error = _HostilePrimaryError("setup failed")

    def fail_fdopen(_descriptor: int, _mode: str) -> BinaryIO:
        raise primary_error

    def fail_close(_descriptor: int) -> None:
        raise _HostileCleanupError("cleanup failed")

    monkeypatch.setattr(spool_module.os, "fdopen", fail_fdopen)
    monkeypatch.setattr(spool_module.os, "close", fail_close)

    try:
        with pytest.raises(RuntimeError) as captured:
            ReplaySpool.from_stream(io.BytesIO(b"x" * 32), memory_limit=8)

        assert captured.value is primary_error
        notes = BaseException.__getattribute__(primary_error, "__notes__")
        assert notes == ["temporary descriptor close also failed: _HostileCleanupError"]
        assert _fallback_block_reason(primary_error) is _FallbackBlockReason.SOURCE_OWNERSHIP
    finally:
        if created:
            descriptor, path = created[0]
            try:
                original_close(descriptor)
            except OSError:
                pass
            try:
                original_unlink(path)
            except FileNotFoundError:
                pass


@pytest.mark.parametrize("operation", ["close", "unlink"])
@pytest.mark.parametrize(
    "cleanup_error",
    [
        MemoryError("capacity"),
        KeyboardInterrupt(),
        SystemExit(2),
        _FatalSpoolCleanup(),
        ExceptionGroup("outer", [MemoryError("nested capacity")]),
        BaseExceptionGroup("outer", [KeyboardInterrupt()]),
    ],
)
def test_pending_spill_process_cleanup_wins_over_fdopen_failure(
    monkeypatch: pytest.MonkeyPatch,
    operation: str,
    cleanup_error: BaseException,
) -> None:
    created = _track_temp_creation(monkeypatch)
    original_close = os.close
    original_unlink = os.unlink
    primary_error = RuntimeError("fdopen failed")

    def fail_fdopen(_descriptor: int, _mode: str) -> BinaryIO:
        raise primary_error

    def fail_cleanup(*_args: object) -> None:
        raise cleanup_error

    monkeypatch.setattr(spool_module.os, "fdopen", fail_fdopen)
    monkeypatch.setattr(spool_module.os, operation, fail_cleanup)

    try:
        with pytest.raises(type(cleanup_error)) as captured:
            ReplaySpool.from_stream(io.BytesIO(b"x" * 32), memory_limit=8)

        assert captured.value is cleanup_error
    finally:
        if created:
            descriptor, path = created[0]
            try:
                original_close(descriptor)
            except OSError:
                pass
            try:
                original_unlink(path)
            except FileNotFoundError:
                pass


@pytest.mark.parametrize(
    "cleanup_error",
    [
        MemoryError("capacity"),
        KeyboardInterrupt(),
        SystemExit(2),
        _FatalSpoolCleanup(),
        ExceptionGroup("outer", [MemoryError("nested capacity")]),
        BaseExceptionGroup("outer", [KeyboardInterrupt()]),
    ],
)
def test_active_spill_process_cleanup_wins_over_primary_failure(
    cleanup_error: BaseException,
) -> None:
    primary_error = RuntimeError("read failed")

    class FailingClose:
        def close(self) -> None:
            raise cleanup_error

    with pytest.raises(type(cleanup_error)) as captured:
        spool_module._cleanup_spill(  # type: ignore[arg-type]
            FailingClose(),
            None,
            primary_error,
        )

    assert captured.value is cleanup_error


@pytest.mark.parametrize(
    "cleanup_error",
    [
        MemoryError("capacity"),
        KeyboardInterrupt(),
        SystemExit(2),
        _FatalSpoolCleanup(),
        ExceptionGroup("outer", [MemoryError("nested capacity")]),
        BaseExceptionGroup("outer", [KeyboardInterrupt()]),
    ],
)
def test_completed_spool_process_cleanup_wins_over_restoration_failure(
    monkeypatch: pytest.MonkeyPatch,
    cleanup_error: BaseException,
) -> None:
    class RestoreFailureAfterRead(io.BytesIO):
        def __init__(self) -> None:
            super().__init__(b"source")
            self.read_started = False

        def read(self, size: int = -1) -> bytes:
            self.read_started = True
            return super().read(size)

        def seek(self, position: int, whence: int = 0) -> int:
            if self.read_started and position == 3 and whence == 0:
                raise OSError("restore failed")
            return super().seek(position, whence)

    source = RestoreFailureAfterRead()
    source.seek(3)

    def fail_close(_spool: ReplaySpool) -> None:
        raise cleanup_error

    monkeypatch.setattr(ReplaySpool, "close", fail_close)

    with pytest.raises(type(cleanup_error)) as captured:
        ReplaySpool.from_stream(source)

    assert captured.value is cleanup_error
