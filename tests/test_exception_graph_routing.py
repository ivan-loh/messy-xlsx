"""Fallback exclusions must traverse the complete bounded exception graph."""

from __future__ import annotations

from typing import Any

import pytest

from messy_xlsx._fallback_signals import (
    _FallbackBlockReason,
    _mark_fallback_blocked,
)
from messy_xlsx.parsing.fallback import FallbackCoordinator


class _FailingReader:
    def __init__(self, error: BaseException, *, suppress: bool = False) -> None:
        self.error = error
        self.suppress = suppress

    def __enter__(self) -> _FailingReader:
        return self

    def __exit__(self, *_args: object) -> bool:
        return self.suppress

    def read_table(self) -> Any:
        raise self.error


class _SuccessfulReader:
    def read_table(self) -> str:
        return "fallback"

    def close(self) -> None:
        return None


def _raised_from(
    outer: Exception,
    inner: BaseException,
    *,
    suppress_context: bool = False,
) -> Exception:
    try:
        raise inner
    except BaseException:
        try:
            if suppress_context:
                raise outer from None
            raise outer from inner
        except Exception as captured:
            return captured


def _raised_with_context(outer: Exception, inner: Exception) -> Exception:
    try:
        raise inner
    except Exception:
        try:
            raise outer
        except Exception as captured:
            return captured


def _assert_no_fallback(error: BaseException, *, suppress: bool = False) -> None:
    fallback_calls = 0

    def fallback_factory() -> _SuccessfulReader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _SuccessfulReader()

    with pytest.raises(type(error)) as captured:
        FallbackCoordinator(lambda _error: True).materialize(
            lambda: _FailingReader(error, suppress=suppress),
            fallback_factory,
        )

    assert captured.value is error
    assert fallback_calls == 0


@pytest.mark.parametrize(
    "inner",
    [
        _mark_fallback_blocked(
            OSError("partially consumed source"),
            _FallbackBlockReason.SOURCE_OWNERSHIP,
        ),
        PermissionError("denied"),
        FileNotFoundError("missing"),
    ],
    ids=["consumed-source", "permission", "missing-file"],
)
def test_explicit_causes_block_fallback_and_preserve_outer_identity(
    inner: Exception,
) -> None:
    outer = _raised_from(RuntimeError("wrapped"), inner)

    _assert_no_fallback(outer)


def test_unsuppressed_hard_failure_context_blocks_fallback() -> None:
    outer = _raised_with_context(RuntimeError("wrapped"), PermissionError("denied"))

    _assert_no_fallback(outer)


def test_suppressed_hard_failure_context_is_not_traversed() -> None:
    outer = _raised_from(
        RuntimeError("wrapped"),
        PermissionError("denied"),
        suppress_context=True,
    )
    fallback_calls = 0

    def fallback_factory() -> _SuccessfulReader:
        nonlocal fallback_calls
        fallback_calls += 1
        return _SuccessfulReader()

    result = FallbackCoordinator(lambda _error: True).materialize(
        lambda: _FailingReader(outer),
        fallback_factory,
    )

    assert result == "fallback"
    assert fallback_calls == 1


@pytest.mark.parametrize("link", ["cause", "context"])
def test_wrapped_process_failure_cannot_be_suppressed_or_retried(link: str) -> None:
    if link == "cause":
        outer = _raised_from(RuntimeError("wrapped"), MemoryError("capacity"))
    else:
        outer = _raised_with_context(
            RuntimeError("wrapped"),
            MemoryError("capacity"),  # type: ignore[arg-type]
        )

    _assert_no_fallback(outer, suppress=True)


def test_exception_cause_cycles_terminate_and_remain_retryable() -> None:
    first = RuntimeError("first")
    second = RuntimeError("second")
    first.__cause__ = second
    second.__cause__ = first

    result = FallbackCoordinator(lambda _error: True).materialize(
        lambda: _FailingReader(first),
        _SuccessfulReader,
    )

    assert result == "fallback"


def test_oversized_exception_cause_graph_fails_closed() -> None:
    outer = RuntimeError("node 0")
    cursor = outer
    for index in range(1, 10_002):
        nested = RuntimeError(f"node {index}")
        cursor.__cause__ = nested
        cursor = nested

    _assert_no_fallback(outer)
