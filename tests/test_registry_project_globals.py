"""Bounded behavior fingerprints for project types referenced by the registry."""

from __future__ import annotations

import pandas as pd

import messy_xlsx._source as source_module
import messy_xlsx._spool as spool_module
from messy_xlsx._spool import ReplaySpool
from messy_xlsx.exceptions import FormatError, MessyXlsxError
from messy_xlsx.parsing.base_handler import FormatHandler, ParseOptions
from messy_xlsx.parsing.csv_handler import CSVHandler
from messy_xlsx.parsing.handler_registry import HandlerRegistry


class _BehaviorChangingInt(int):
    def __add__(self, _other: object) -> int:
        return -1


class _InheritedBehaviorOverride(FormatHandler):
    def _apply_row_limits(self, df: pd.DataFrame, options: ParseOptions) -> pd.DataFrame:
        return df


class _ExternalFormatBehavior:
    def to_dict(self) -> dict[str, bool]:
        return {"external": True}


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


def test_scalar_subclass_in_project_function_defaults_forces_compatibility() -> None:
    registry = HandlerRegistry()
    function = vars(ReplaySpool)["from_stream"].__func__
    original_defaults = function.__defaults__

    try:
        function.__defaults__ = (_BehaviorChangingInt(spool_module.DEFAULT_MEMORY_LIMIT),)
        assert registry._uses_builtin_components() is False
    finally:
        function.__defaults__ = original_defaults

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


def test_transitive_project_type_external_mro_change_is_detected_and_restored() -> None:
    registry = HandlerRegistry()
    original_bases = FormatError.__bases__

    try:
        FormatError.__bases__ = (_ExternalFormatBehavior, *original_bases)
        assert FormatError("changed").to_dict() == {"external": True}
        assert registry._uses_builtin_components() is False
    finally:
        FormatError.__bases__ = original_bases

    assert registry._uses_builtin_components() is True


def test_exact_handler_mro_mutation_is_detected_and_restored() -> None:
    registry = HandlerRegistry()
    original_bases = CSVHandler.__bases__

    try:
        CSVHandler.__bases__ = (_InheritedBehaviorOverride,)
        assert registry._uses_builtin_components() is False
    finally:
        CSVHandler.__bases__ = original_bases

    assert registry._uses_builtin_components() is True
