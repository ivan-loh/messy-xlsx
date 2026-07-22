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
