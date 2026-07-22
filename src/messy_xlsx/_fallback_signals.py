"""Private structured reasons that make a backend retry unsafe."""

from __future__ import annotations

from enum import Enum, auto
from typing import TypeVar

_FALLBACK_BLOCK_REASON_KEY = "_messy_xlsx_fallback_block_reason"
_MAX_EXCEPTION_TREE_NODES = 10_000
_ErrorT = TypeVar("_ErrorT", bound=BaseException)


class _FallbackBlockReason(Enum):
    """Internal reason why the same operation must not use another backend."""

    CONFIGURATION = auto()
    SOURCE_OWNERSHIP = auto()


def _mark_fallback_blocked(
    error: _ErrorT,
    reason: _FallbackBlockReason,
) -> _ErrorT:
    """Attach a private typed signal while preserving the exact exception."""
    if not isinstance(reason, _FallbackBlockReason):
        return error
    try:
        state = BaseException.__getattribute__(error, "__dict__")
        if type(state) is dict:
            dict.__setitem__(state, _FALLBACK_BLOCK_REASON_KEY, reason)
    except BaseException:
        pass
    return error


def _fallback_block_reason(
    error: BaseException,
) -> _FallbackBlockReason | None:
    """Read the first valid signal from a bounded nested exception tree."""
    stack = [error]
    seen: set[int] = set()
    while stack:
        candidate = stack.pop()
        identity = id(candidate)
        if identity in seen:
            continue
        if len(seen) >= _MAX_EXCEPTION_TREE_NODES:
            return _FallbackBlockReason.SOURCE_OWNERSHIP
        seen.add(identity)
        reason = _direct_fallback_block_reason(candidate)
        if reason is not None:
            return reason
        nested = _nested_exceptions(candidate)
        if nested is None:
            if isinstance(candidate, BaseExceptionGroup):
                return _FallbackBlockReason.SOURCE_OWNERSHIP
            continue
        stack.extend(reversed(nested))
    return None


def _direct_fallback_block_reason(
    error: BaseException,
) -> _FallbackBlockReason | None:
    """Read one exception's private marker without virtual hooks."""
    try:
        state = BaseException.__getattribute__(error, "__dict__")
        if type(state) is not dict:
            return None
        reason = dict.get(state, _FALLBACK_BLOCK_REASON_KEY)
    except BaseException:
        return None
    return reason if isinstance(reason, _FallbackBlockReason) else None


def _is_fallback_blocked(error: BaseException) -> bool:
    """Return whether any nested failure carries an unsafe retry marker."""
    return _fallback_block_reason(error) is not None


def _contains_process_failure(error: BaseException) -> bool:
    """Return whether a bounded nested tree contains a process-level failure."""
    stack = [error]
    seen: set[int] = set()
    while stack:
        candidate = stack.pop()
        identity = id(candidate)
        if identity in seen:
            continue
        if len(seen) >= _MAX_EXCEPTION_TREE_NODES:
            return True
        seen.add(identity)
        if isinstance(candidate, MemoryError) or not isinstance(candidate, Exception):
            return True
        nested = _nested_exceptions(candidate)
        if nested is None:
            if isinstance(candidate, BaseExceptionGroup):
                return True
            continue
        stack.extend(reversed(nested))
    return False


def _blocks_backend_retry(error: BaseException) -> bool:
    """Return whether process semantics or an internal marker forbid retry."""
    return _contains_process_failure(error) or _is_fallback_blocked(error)


def _nested_exceptions(
    error: BaseException,
) -> tuple[BaseException, ...] | None:
    if not isinstance(error, BaseExceptionGroup):
        return ()
    try:
        nested = BaseException.__getattribute__(error, "exceptions")
    except BaseException:
        return None
    if type(nested) is not tuple or len(nested) > _MAX_EXCEPTION_TREE_NODES:
        return None
    return nested
