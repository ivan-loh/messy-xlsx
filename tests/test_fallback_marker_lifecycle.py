"""Fallback markers must remain local to the exact exception they annotate."""

from __future__ import annotations

import gc
import weakref

import pytest

from messy_xlsx._fallback_signals import (
    _fallback_block_reason,
    _FallbackBlockReason,
    _mark_fallback_blocked,
)


class _HostileMarkerStorageError(RuntimeError):
    @property
    def __dict__(self) -> dict[str, object]:
        raise AssertionError("virtual exception dictionary must not be read")

    @property
    def _messy_xlsx_fallback_block_reason(self) -> object:
        raise AssertionError("hostile marker descriptor must not be read")

    @_messy_xlsx_fallback_block_reason.setter
    def _messy_xlsx_fallback_block_reason(self, _value: object) -> None:
        raise AssertionError("hostile marker descriptor must not be written")

    def __setattr__(self, _name: str, _value: object) -> None:
        raise AssertionError("virtual exception setter must not be called")


class _WeakPayload:
    pass


@pytest.mark.parametrize("reason", list(_FallbackBlockReason))
def test_hostile_exception_marker_preserves_identity_and_typed_reason(
    reason: _FallbackBlockReason,
) -> None:
    error = _HostileMarkerStorageError("failure")

    marked = _mark_fallback_blocked(error, reason)

    assert marked is error
    assert _fallback_block_reason(error) is reason


def test_marker_retrieval_ignores_untyped_marker_spoof() -> None:
    error = RuntimeError("unmarked")
    error.__dict__["_messy_xlsx_fallback_block_reason"] = "SOURCE_OWNERSHIP"

    assert _fallback_block_reason(error) is None


def _mark_raised_error_and_release_graph() -> weakref.ReferenceType[_WeakPayload]:
    payload = _WeakPayload()
    payload_reference = weakref.ref(payload)
    error = _HostileMarkerStorageError(payload)

    try:
        raise error
    except _HostileMarkerStorageError as caught:
        traceback = BaseException.__getattribute__(caught, "__traceback__")
        assert traceback is not None
        assert _mark_fallback_blocked(caught, _FallbackBlockReason.SOURCE_OWNERSHIP) is caught

    del traceback
    del error
    del payload
    return payload_reference


def test_marker_storage_does_not_retain_exception_traceback_graph() -> None:
    payload_reference = _mark_raised_error_and_release_graph()

    gc.collect()

    assert payload_reference() is None


def test_many_hostile_markers_do_not_mark_an_unrelated_error() -> None:
    for index in range(10_001):
        _mark_fallback_blocked(
            _HostileMarkerStorageError(index),
            _FallbackBlockReason.CONFIGURATION,
        )

    assert _fallback_block_reason(RuntimeError("unrelated")) is None
