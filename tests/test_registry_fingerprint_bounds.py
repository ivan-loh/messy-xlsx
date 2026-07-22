"""Cost bounds must be checked before large registry graph allocations."""

from __future__ import annotations

import builtins
from collections.abc import Callable
from types import SimpleNamespace
from typing import Any

import pytest

import messy_xlsx.parsing.handler_registry as registry_module


class _BehaviorChangingInt(int):
    def __add__(self, _other: object) -> int:
        return -1


class _HostileScalarInt(int):
    hash_calls = 0

    def __hash__(self) -> int:
        type(self).hash_calls += 1
        raise AssertionError("fingerprint must not invoke scalar subclass hash")

    def __eq__(self, _other: object) -> bool:
        raise AssertionError("fingerprint must not invoke scalar subclass equality")


def test_oversized_instance_namespace_fails_before_copy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    component = SimpleNamespace()
    vars(component).update({f"oversized_{index}": index for index in range(10_001)})

    def guarded_dict(*args: object, **kwargs: object) -> dict[object, object]:
        if args and len(args[0]) > 10_000:  # type: ignore[arg-type]
            raise AssertionError("oversized namespace was copied before budget rejection")
        return builtins.dict(*args, **kwargs)  # type: ignore[call-overload]

    monkeypatch.setattr(registry_module, "dict", guarded_dict, raising=False)

    with pytest.raises(registry_module._FingerprintError, match="budget exceeded"):
        registry_module._CompositionFingerprinter(include_identity=True).token(component)


def test_small_instance_namespace_still_uses_normal_copy_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    component = SimpleNamespace(value=1)
    copy_calls = 0

    def tracking_dict(*args: object, **kwargs: object) -> dict[object, object]:
        nonlocal copy_calls
        copy_calls += 1
        return builtins.dict(*args, **kwargs)  # type: ignore[call-overload]

    monkeypatch.setattr(registry_module, "dict", tracking_dict, raising=False)

    registry_module._CompositionFingerprinter(include_identity=True).token(component)

    assert copy_calls == 1


def test_component_scalar_subclass_mutation_and_restore_changes_token() -> None:
    component = SimpleNamespace(limit=1)
    expected = registry_module._CompositionFingerprinter(include_identity=False).token(component)

    component.limit = _BehaviorChangingInt(1)
    changed = registry_module._CompositionFingerprinter(include_identity=False).token(component)
    component.limit = 1
    restored = registry_module._CompositionFingerprinter(include_identity=False).token(component)

    assert changed != expected
    assert restored == expected


def test_scalar_subclass_fingerprint_never_invokes_virtual_hash_or_equality() -> None:
    component = SimpleNamespace(limit=_HostileScalarInt(1))
    _HostileScalarInt.hash_calls = 0

    token = registry_module._CompositionFingerprinter(include_identity=False).token(component)

    hash(token)
    assert _HostileScalarInt.hash_calls == 0


@pytest.mark.parametrize(
    "fingerprint",
    [
        lambda value_type: registry_module._BehaviorFingerprinter()._project_type_token(value_type),
        lambda value_type: registry_module._class_behavior_token((value_type,)),
    ],
    ids=["referenced-project-type", "canonical-class-behavior"],
)
def test_oversized_class_namespace_fails_before_sort(
    monkeypatch: pytest.MonkeyPatch,
    fingerprint: Callable[[type[object]], object],
) -> None:
    value_type = type(
        "OversizedBehavior",
        (),
        {f"behavior_{index}": index for index in range(10_001)},
    )

    def guarded_sorted(iterable: Any, *args: object, **kwargs: object) -> list[Any]:
        items = list(iterable)
        if len(items) > 10_000:
            raise AssertionError("oversized namespace was sorted before budget rejection")
        return builtins.sorted(items, *args, **kwargs)

    monkeypatch.setattr(registry_module, "sorted", guarded_sorted, raising=False)

    with pytest.raises(registry_module._FingerprintError, match="budget exceeded"):
        fingerprint(value_type)


def test_small_class_namespace_still_uses_normal_sort_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class SmallBehavior:
        def method(self) -> None:
            return None

    sort_calls = 0

    def tracking_sorted(iterable: Any, *args: object, **kwargs: object) -> list[Any]:
        nonlocal sort_calls
        sort_calls += 1
        return builtins.sorted(iterable, *args, **kwargs)

    monkeypatch.setattr(registry_module, "sorted", tracking_sorted, raising=False)

    registry_module._BehaviorFingerprinter()._project_type_token(SmallBehavior)

    assert sort_calls >= 1


def test_oversized_transitive_project_type_fails_closed_before_sort(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    value_type = type(
        "OversizedTransitiveBehavior",
        (),
        {
            "__module__": "messy_xlsx.oversized_test",
            **{f"behavior_{index}": index for index in range(10_001)},
        },
    )

    def guarded_sorted(iterable: Any, *args: object, **kwargs: object) -> list[Any]:
        items = list(iterable)
        if len(items) > 10_000:
            raise AssertionError("oversized transitive namespace was sorted")
        return builtins.sorted(items, *args, **kwargs)

    monkeypatch.setattr(registry_module, "sorted", guarded_sorted, raising=False)

    with pytest.raises(registry_module._FingerprintError, match="budget exceeded"):
        registry_module._BehaviorFingerprinter()._project_global_reference_token(value_type)


def test_canonical_class_fingerprint_bypasses_hostile_metaclass_hooks() -> None:
    virtual_hook_calls: list[str] = []

    class HostileMeta(type):
        def __getattribute__(cls, name: str) -> object:
            if name in {"__bases__", "__dict__", "__mro__"}:
                virtual_hook_calls.append(name)
                raise AssertionError("fingerprint invoked a virtual metaclass hook")
            return type.__getattribute__(cls, name)

        @property
        def __dict__(cls) -> object:
            virtual_hook_calls.append("__dict__ property")
            raise AssertionError("fingerprint invoked a metaclass property")

        @property
        def __bases__(cls) -> object:
            virtual_hook_calls.append("__bases__ property")
            raise AssertionError("fingerprint invoked a metaclass property")

        @property
        def __mro__(cls) -> object:
            virtual_hook_calls.append("__mro__ property")
            raise AssertionError("fingerprint invoked a metaclass property")

    cycle: list[object] = []
    cycle.append(cycle)

    class HostileBehavior(metaclass=HostileMeta):
        behavior = cycle

    virtual_hook_calls.clear()
    registry_module._class_behavior_token((HostileBehavior,))

    assert virtual_hook_calls == []


def test_non_exact_handler_list_fails_closed_without_iteration() -> None:
    class TrackingHandlerList(list[object]):
        def __init__(self, values: list[object]) -> None:
            super().__init__(values)
            self.iter_calls = 0

        def __iter__(self):  # type: ignore[no-untyped-def]
            self.iter_calls += 1
            return super().__iter__()

    registry = registry_module.HandlerRegistry()
    handlers = TrackingHandlerList(registry.handlers)
    registry.handlers = handlers  # type: ignore[assignment]

    assert registry._uses_builtin_components() is False
    assert handlers.iter_calls == 0


def test_oversized_handler_list_fails_closed_before_scanning(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    registry = registry_module.HandlerRegistry()
    handler = registry.handlers[0]
    registry.handlers = [handler] * (registry_module._MAX_FINGERPRINT_NODES + 1)
    handler_type_calls = 0

    def tracking_type(value: object) -> type[object]:
        nonlocal handler_type_calls
        if value is handler:
            handler_type_calls += 1
        return builtins.type(value)

    monkeypatch.setattr(registry_module, "type", tracking_type, raising=False)

    assert registry._uses_builtin_components() is False
    assert handler_type_calls == 0


def test_oversized_registry_state_fails_closed_before_key_materialization(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    registry = registry_module.HandlerRegistry()
    state = vars(registry)
    state.update(
        {
            f"unexpected_state_{index}": None
            for index in range(registry_module._MAX_FINGERPRINT_NODES + 1)
        }
    )
    state_set_calls = 0

    def tracking_set(*args: object) -> set[object]:
        nonlocal state_set_calls
        if args and args[0] is state:
            state_set_calls += 1
        return builtins.set(*args)

    monkeypatch.setattr(registry_module, "set", tracking_set, raising=False)

    assert registry._uses_builtin_components() is False
    assert state_set_calls == 0
