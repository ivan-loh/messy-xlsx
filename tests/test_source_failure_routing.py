"""Source-ownership failures must never retry a partially consumed stream."""

from __future__ import annotations

import pytest

from messy_xlsx._fallback_signals import (
    _fallback_block_reason,
    _FallbackBlockReason,
)
from messy_xlsx._source import SourceHandle
from messy_xlsx.parsing.fallback import FallbackCoordinator


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
