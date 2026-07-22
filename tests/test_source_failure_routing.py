"""Source-ownership failures must never retry a partially consumed stream."""

from __future__ import annotations

import pytest

from messy_xlsx._fallback_signals import (
    _fallback_block_reason,
    _FallbackBlockReason,
)
from messy_xlsx._source import SourceHandle
from messy_xlsx.parsing.fallback import FallbackCoordinator


class _HostileMarkerStorageError(OSError):
    @property
    def __dict__(self) -> dict[str, object]:
        raise AssertionError("marker storage must not depend on __dict__")

    @property
    def _messy_xlsx_fallback_block_reason(self) -> object:
        raise AssertionError("private marker descriptor blocks direct reads")

    @_messy_xlsx_fallback_block_reason.setter
    def _messy_xlsx_fallback_block_reason(self, _value: object) -> None:
        raise AssertionError("private marker descriptor blocks direct writes")

    def __getattribute__(self, name: str) -> object:
        if name == "caller_controlled_probe":
            raise AssertionError("virtual exception getter must not be invoked")
        return super().__getattribute__(name)

    def __setattr__(self, _name: str, _value: object) -> None:
        raise AssertionError("virtual exception setter must not be invoked")


class _SuccessfulReader:
    def read_table(self) -> bytes:
        return b"fallback"

    def close(self) -> None:
        return None


@pytest.mark.parametrize(
    "source_error",
    [
        OSError("source read failed"),
        ExceptionGroup("outer", [OSError("nested source read failed")]),
    ],
)
def test_partial_source_read_failure_is_marked_and_never_retried(
    source_error: BaseException,
) -> None:
    class PartialFailureSource:
        def __init__(self) -> None:
            self.read_calls = 0

        def seekable(self) -> bool:
            return False

        def tell(self) -> int:
            return 0

        def read(self, _size: int = -1) -> bytes:
            self.read_calls += 1
            if self.read_calls == 1:
                return b"consumed prefix"
            raise source_error

    source = PartialFailureSource()
    fallback_calls = 0

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    with pytest.raises(type(source_error)) as captured:
        FallbackCoordinator(lambda _error: True).materialize(
            lambda: SourceHandle(source),  # type: ignore[arg-type]
            fallback_factory,
        )

    assert captured.value is source_error
    assert source.read_calls == 2
    assert fallback_calls == 0
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP


def test_hostile_exception_marker_storage_fails_closed_after_partial_read() -> None:
    source_error = _HostileMarkerStorageError("source read failed")

    class PartialFailureSource:
        def __init__(self) -> None:
            self.read_calls = 0

        def seekable(self) -> bool:
            return False

        def tell(self) -> int:
            return 0

        def read(self, _size: int = -1) -> bytes:
            self.read_calls += 1
            if self.read_calls == 1:
                return b"consumed prefix"
            raise source_error

    source = PartialFailureSource()
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

    with pytest.raises(_HostileMarkerStorageError) as captured:
        FallbackCoordinator(classifier).materialize(
            lambda: SourceHandle(source),  # type: ignore[arg-type]
            fallback_factory,
        )

    assert captured.value is source_error
    assert source.read_calls == 2
    assert classifier_calls == 0
    assert fallback_calls == 0
    assert _fallback_block_reason(captured.value) is _FallbackBlockReason.SOURCE_OWNERSHIP
