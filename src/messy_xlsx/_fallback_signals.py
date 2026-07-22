"""Private retry signals and non-masking exception diagnostics."""

from __future__ import annotations

from enum import Enum, auto
from types import TracebackType
from typing import NoReturn, TypeVar

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
    """Find any bounded-tree leaf whose semantics make a retry unsafe."""
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
        if (
            isinstance(candidate, (PermissionError, FileNotFoundError, MemoryError))
            or not isinstance(candidate, Exception)
            or _direct_fallback_block_reason(candidate) is not None
        ):
            return True
        nested = _nested_exceptions(candidate)
        if nested is None:
            if isinstance(candidate, BaseExceptionGroup):
                return True
            continue
        stack.extend(reversed(nested))
    return False


def _failure_summary(error: BaseException) -> dict[str, str]:
    """Return useful context without source paths, values, or error messages."""
    return {"type": _type_name(error)}


def _attach_cleanup_failure(
    primary_error: BaseException,
    cleanup_error: BaseException,
) -> None:
    _merge_backend_context(
        primary_error,
        cleanup_failure=_failure_summary(cleanup_error),
    )


def _attach_operation_failure(
    cleanup_error: BaseException,
    operation_error: BaseException,
) -> None:
    _merge_backend_context(
        cleanup_error,
        operation_failure=_failure_summary(operation_error),
    )


def _merge_backend_context(
    error: BaseException,
    **updates: dict[str, str],
) -> None:
    """Best-effort merge of sanitized summaries without virtual hooks."""
    try:
        state = BaseException.__getattribute__(error, "__dict__")
        if type(state) is not dict:
            return
        context: dict[str, dict[str, str]] = {}
        existing = dict.get(state, "backend_context")
        if type(existing) is dict:
            for name in (
                "primary_failure",
                "fallback_failure",
                "operation_failure",
                "cleanup_failure",
                "classifier_failure",
            ):
                summary = dict.get(existing, name)
                if (
                    type(summary) is dict
                    and set(summary) == {"type"}
                    and isinstance(dict.__getitem__(summary, "type"), str)
                ):
                    context[name] = {"type": dict.__getitem__(summary, "type")}
        context.update({name: dict(summary) for name, summary in updates.items()})
        dict.__setitem__(state, "backend_context", context)
    except BaseException:
        return


def _type_name(value: object) -> str:
    """Read a concrete type name without invoking metaclass overrides."""
    try:
        name = type.__getattribute__(type(value), "__name__")
    except BaseException:
        return "<unknown>"
    return name if isinstance(name, str) else "<unknown>"


def _exception_traceback(error: BaseException) -> TracebackType | None:
    """Read traceback state without invoking exception attribute overrides."""
    try:
        traceback = BaseException.__getattribute__(error, "__traceback__")
    except BaseException:
        return None
    return traceback if isinstance(traceback, TracebackType) else None


def _safe_add_note(error: BaseException, note: str) -> None:
    """Attach sanitized diagnostics without allowing annotation failures to mask."""
    try:
        BaseException.add_note(error, note)
    except BaseException:
        return


def _raise_with_traceback(
    error: BaseException,
    traceback: TracebackType | None,
) -> NoReturn:
    """Raise the exact exception with captured traceback state when possible."""
    try:
        error = BaseException.with_traceback(error, traceback)
    except BaseException:
        pass
    raise error


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
