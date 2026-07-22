"""Bounded behavior fingerprints for project types referenced by the registry."""

from __future__ import annotations

import messy_xlsx._source as source_module
from messy_xlsx._spool import ReplaySpool
from messy_xlsx.exceptions import MessyXlsxError
from messy_xlsx.parsing.handler_registry import HandlerRegistry


def test_referenced_project_type_global_rebind_and_restore_is_detected() -> None:
    registry = HandlerRegistry()
    original = source_module._is_readable

    def replacement(_value: object) -> bool:
        return False

    try:
        source_module._is_readable = replacement
        assert registry._uses_builtin_components() is False
    finally:
        source_module._is_readable = original

    assert registry._uses_builtin_components() is True


def test_transitively_referenced_project_type_behavior_is_detected() -> None:
    registry = HandlerRegistry()
    descriptor = vars(ReplaySpool)["from_stream"]

    try:
        ReplaySpool.from_stream = classmethod(  # type: ignore[method-assign]
            lambda cls, stream, memory_limit=0: cls(None, None)
        )
        assert registry._uses_builtin_components() is False
    finally:
        ReplaySpool.from_stream = descriptor  # type: ignore[method-assign]

    assert registry._uses_builtin_components() is True


def test_inherited_project_base_behavior_is_detected_and_restored() -> None:
    registry = HandlerRegistry()
    original = MessyXlsxError.__init__

    try:
        MessyXlsxError.__init__ = lambda self, message, context=None: Exception.__init__(  # type: ignore[method-assign]
            self,
            message,
        )
        assert registry._uses_builtin_components() is False
    finally:
        MessyXlsxError.__init__ = original  # type: ignore[method-assign]

    assert registry._uses_builtin_components() is True
