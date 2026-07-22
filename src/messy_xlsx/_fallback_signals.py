"""Private structured reasons that make a backend retry unsafe."""

from __future__ import annotations

from enum import Enum, auto
from typing import TypeVar

_FALLBACK_BLOCK_REASON_KEY = "_messy_xlsx_fallback_block_reason"
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
    """Read a valid private signal without invoking exception overrides."""
    try:
        state = BaseException.__getattribute__(error, "__dict__")
        if type(state) is not dict:
            return None
        reason = dict.get(state, _FALLBACK_BLOCK_REASON_KEY)
    except BaseException:
        return None
    return reason if isinstance(reason, _FallbackBlockReason) else None


def _is_fallback_blocked(error: BaseException) -> bool:
    """Return whether an internal producer marked this failure as unsafe."""
    return _fallback_block_reason(error) is not None
