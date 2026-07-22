"""Seekable source I/O failures must block every backend retry path."""

from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
import pytest

from messy_xlsx._fallback_signals import (
    _fallback_block_reason,
    _FallbackBlockReason,
)
from messy_xlsx._source import SourceHandle
from messy_xlsx.parsing.base_handler import FormatHandler
from messy_xlsx.parsing.fallback import FallbackCoordinator
from messy_xlsx.parsing.handler_registry import HandlerRegistry


class _InitialSeekFailure(BytesIO):
    def __init__(self, content: bytes, error: OSError) -> None:
        super().__init__(content)
        self.error = error
        self.fail_initial_seek = False

    def seek(self, position: int, whence: int = 0) -> int:
        if self.fail_initial_seek and position == 0 and whence == 0:
            raise self.error
        return super().seek(position, whence)


class _EntryTellFailure(BytesIO):
    def __init__(self, content: bytes, error: OSError) -> None:
        super().__init__(content)
        self.error = error
        self.fail_entry_tell = False

    def tell(self) -> int:
        if self.fail_entry_tell:
            raise self.error
        return super().tell()


class _SourceReader:
    def __init__(self, handle: SourceHandle) -> None:
        self.handle = handle

    def read_table(self) -> bytes:
        return self.handle.read_bytes()

    def close(self) -> None:
        return None


class _SuccessfulReader:
    def read_table(self) -> bytes:
        return b"fallback"

    def close(self) -> None:
        return None


def _frame_names(error: BaseException) -> list[str]:
    traceback = BaseException.__getattribute__(error, "__traceback__")
    names: list[str] = []
    while traceback is not None:
        names.append(traceback.tb_frame.f_code.co_name)
        traceback = traceback.tb_next
    return names


def _failing_seekable_source() -> tuple[_InitialSeekFailure, SourceHandle, OSError]:
    error = OSError("initial seek failed")
    source = _InitialSeekFailure(b"complete source", error)
    source.seek(5)
    handle = SourceHandle(source)
    source.fail_initial_seek = True
    return source, handle, error


def test_seekable_initial_seek_failure_blocks_coordinator_fallback() -> None:
    source, handle, source_error = _failing_seekable_source()
    fallback_calls = 0
    classifier_calls = 0

    def classifier(_error: Exception) -> bool:
        nonlocal classifier_calls
        classifier_calls += 1
        return True

    def fallback_factory() -> _SuccessfulReader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _SuccessfulReader()

    with pytest.raises(OSError) as captured:
        FallbackCoordinator(classifier).materialize(
            lambda: _SourceReader(handle),
            fallback_factory,
        )

    assert captured.value is source_error
    assert classifier_calls == 0
    assert fallback_calls == 0
    assert source.tell() == 5
    assert "_open_binary_unchecked" in _frame_names(captured.value)
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP


def test_seekable_entry_tell_failure_blocks_coordinator_fallback() -> None:
    source_error = OSError("entry tell failed")
    source = _EntryTellFailure(b"complete source", source_error)
    source.seek(5)
    handle = SourceHandle(source)
    source.fail_entry_tell = True
    fallback_calls = 0
    classifier_calls = 0

    def classifier(_error: Exception) -> bool:
        nonlocal classifier_calls
        classifier_calls += 1
        return True

    def fallback_factory() -> _SuccessfulReader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _SuccessfulReader()

    with pytest.raises(OSError) as captured:
        FallbackCoordinator(classifier).materialize(
            lambda: _SourceReader(handle),
            fallback_factory,
        )

    source.fail_entry_tell = False
    assert captured.value is source_error
    assert classifier_calls == 0
    assert fallback_calls == 0
    assert source.tell() == 5
    assert "_open_binary_unchecked" in _frame_names(captured.value)
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP


def test_seekable_initial_seek_failure_blocks_legacy_handler_retry() -> None:
    source, handle, source_error = _failing_seekable_source()
    fallback_calls = 0

    class PrimaryHandler(FormatHandler):
        _accepts_source_handle = True

        def can_handle(self, format_type: str) -> bool:
            return format_type == "xlsx"

        def parse(self, file_source: Any, sheet: Any, options: Any) -> pd.DataFrame:
            file_source.read_bytes()
            raise AssertionError("source failure must propagate")

        def get_sheet_names(self, file_source: Any) -> list[str]:
            return ["Data"]

        def validate(self, file_source: Any) -> tuple[bool, None]:
            return True, None

    class FallbackHandler(PrimaryHandler):
        _accepts_source_handle = True

        def parse(self, file_source: Any, sheet: Any, options: Any) -> pd.DataFrame:
            nonlocal fallback_calls
            fallback_calls += 1
            return pd.DataFrame({"wrong": [True]})

    registry = HandlerRegistry(handlers=[PrimaryHandler(), FallbackHandler()])

    with pytest.raises(OSError) as captured:
        registry.parse(handle, format_type="xlsx")

    assert captured.value is source_error
    assert fallback_calls == 0
    assert source.tell() == 5
    assert "_open_binary_unchecked" in _frame_names(captured.value)
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP
