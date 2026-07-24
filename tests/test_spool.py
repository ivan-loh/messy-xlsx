from __future__ import annotations

import gc
import io
import os
import sys
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


@pytest.mark.parametrize(
    ("restore_error", "cleanup_wins"),
    [
        (OSError("restore failed"), False),
        (ExceptionGroup("outer", [OSError("nested restore failed")]), False),
        (MemoryError("capacity"), True),
        (ExceptionGroup("outer", [MemoryError("nested capacity")]), True),
        (BaseExceptionGroup("outer", [KeyboardInterrupt()]), True),
    ],
)
def test_spool_restoration_diagnostics_preserve_the_exact_precedence_winner(
    restore_error: BaseException,
    cleanup_wins: bool,
) -> None:
    primary_error = OSError("source read failed")

    class ReadAndRestoreFailure(io.BytesIO):
        def __init__(self) -> None:
            super().__init__(b"source")
            self.read_failed = False

        def read(self, _size: int = -1) -> bytes:
            self.read_failed = True
            raise primary_error

        def seek(self, position: int, whence: int = 0) -> int:
            if self.read_failed and position == 2 and whence == 0:
                raise restore_error
            return super().seek(position, whence)

    source = ReadAndRestoreFailure()
    io.BytesIO.seek(source, 2)
    expected = restore_error if cleanup_wins else primary_error

    with pytest.raises(type(expected)) as captured:
        ReplaySpool.from_stream(source)

    assert captured.value is expected
    traceback = BaseException.__getattribute__(captured.value, "__traceback__")
    frame_names: list[str] = []
    while traceback is not None:
        frame_names.append(traceback.tb_frame.f_code.co_name)
        traceback = traceback.tb_next
    assert ("seek" if cleanup_wins else "read") in frame_names
    if cleanup_wins:
        assert captured.value.backend_context == {  # type: ignore[attr-defined]
            "operation_failure": {"type": "OSError"}
        }
        assert captured.value.__notes__ == ["source operation also failed: OSError"]
    else:
        assert captured.value.backend_context == {  # type: ignore[attr-defined]
            "cleanup_failure": {"type": type(restore_error).__name__}
        }
        assert captured.value.__notes__ == [
            f"cursor restoration also failed: {type(restore_error).__name__}"
        ]
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


@pytest.mark.parametrize(
    "interruption",
    [
        OSError("spill ownership transition interrupted"),
        MemoryError("spill ownership transition interrupted"),
    ],
    ids=["ordinary", "process"],
)
def test_mkstemp_result_is_owned_before_the_next_python_line(
    monkeypatch: pytest.MonkeyPatch,
    interruption: BaseException,
) -> None:
    created = _track_temp_creation(monkeypatch)
    real_close = spool_module.os.close
    real_unlink = spool_module.os.unlink
    close_calls = 0
    unlink_calls = 0
    interrupted = False
    target_code = spool_module._create_spill.__code__

    def track_close(descriptor: int) -> None:
        nonlocal close_calls
        if created and descriptor == created[0][0]:
            close_calls += 1
        real_close(descriptor)

    def track_unlink(path: str | bytes | os.PathLike[str] | os.PathLike[bytes]) -> None:
        nonlocal unlink_calls
        if created and Path(path) == created[0][1]:
            unlink_calls += 1
        real_unlink(path)

    def interrupt_after_mkstemp_return(
        frame: object,
        event: str,
        _arg: object,
    ) -> object:
        nonlocal interrupted
        frame_locals = getattr(frame, "f_locals", {})
        created_descriptor = created[0][0] if created else None
        if (
            getattr(frame, "f_code", None) is target_code
            and event == "line"
            and created_descriptor is not None
            and created_descriptor
            in {
                frame_locals.get("raw_descriptor"),
                frame_locals.get("descriptor_owner"),
            }
            and not interrupted
        ):
            interrupted = True
            sys.settrace(None)
            frame.f_trace = None  # type: ignore[attr-defined]
            raise type(interruption)(*interruption.args)
        return interrupt_after_mkstemp_return

    monkeypatch.setattr(spool_module.os, "close", track_close)
    monkeypatch.setattr(spool_module.os, "unlink", track_unlink)
    sys.settrace(interrupt_after_mkstemp_return)
    try:
        with pytest.raises(type(interruption), match="ownership transition interrupted"):
            ReplaySpool.from_stream(io.BytesIO(b"x" * 32), memory_limit=1)
    finally:
        sys.settrace(None)

    try:
        assert interrupted
        assert len(created) == 1
        descriptor, path = created[0]
        assert _descriptor_is_closed(descriptor)
        assert not path.exists()
        assert close_calls == 1
        assert unlink_calls == 1
    finally:
        if created:
            descriptor, path = created[0]
            try:
                real_close(descriptor)
            except OSError:
                pass
            try:
                real_unlink(path)
            except FileNotFoundError:
                pass


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


def test_replay_spool_return_gap_finalizer_removes_unadopted_spill(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created = _track_temp_creation(monkeypatch)
    target_code = ReplaySpool.from_stream.__func__.__code__
    interrupted = False

    def interrupt_public_return(frame: object, event: str, _arg: object) -> object:
        nonlocal interrupted
        if getattr(frame, "f_code", None) is target_code and event == "return" and not interrupted:
            interrupted = True
            sys.settrace(None)
            frame.f_trace = None  # type: ignore[attr-defined]
            raise MemoryError("replay spool return interrupted")
        return interrupt_public_return

    sys.settrace(interrupt_public_return)
    try:
        with pytest.raises(MemoryError, match="replay spool return interrupted"):
            ReplaySpool.from_stream(io.BytesIO(b"x" * 64), memory_limit=1)
    finally:
        sys.settrace(None)
    gc.collect()

    assert interrupted
    assert len(created) == 1
    assert not created[0][1].exists()


def test_replay_spool_rejects_adopter_that_does_not_confirm_ownership(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created = _track_temp_creation(monkeypatch)

    class NoopAdopter:
        def __call__(self, _spool: ReplaySpool) -> None:
            return

        def owns(self, _spool: ReplaySpool) -> bool:
            return False

    with pytest.raises(RuntimeError, match="adopt"):
        ReplaySpool.from_stream(
            io.BytesIO(b"x" * 64),
            memory_limit=1,
            _adopter=NoopAdopter(),
        )

    assert len(created) == 1
    assert not created[0][1].exists()
