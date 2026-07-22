"""Bounded behavior fingerprints for project types referenced by the registry."""

from __future__ import annotations

import messy_xlsx._source as source_module
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
