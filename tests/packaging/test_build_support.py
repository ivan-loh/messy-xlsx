from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BUILD_SUPPORT_PATH = PROJECT_ROOT / "build_support.py"


def _load_build_support() -> ModuleType:
    assert BUILD_SUPPORT_PATH.is_file(), "Task 2 must provide build_support.py"
    spec = importlib.util.spec_from_file_location("messy_xlsx_build_support", BUILD_SUPPORT_PATH)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.mark.parametrize("mode", ["native", "fallback"])
def test_explicit_build_modes_are_authoritative(mode: str) -> None:
    build_support = _load_build_support()

    assert (
        build_support.resolve_build_mode(
            {"MESSY_XLSX_BUILD_MODE": mode},
            implementation="pypy",
            version_info=(3, 20),
            system="Plan9",
            machine="mips64",
            free_threaded=True,
        )
        == mode
    )


@pytest.mark.parametrize(
    "mode",
    ["", "auto", "NATIVE", " native", "fallback ", "0"],
)
def test_invalid_explicit_build_modes_fail_closed(mode: str) -> None:
    build_support = _load_build_support()

    with pytest.raises(ValueError, match="MESSY_XLSX_BUILD_MODE"):
        build_support.resolve_build_mode({"MESSY_XLSX_BUILD_MODE": mode})


@pytest.mark.parametrize(
    ("system", "machine"),
    [
        ("Linux", "x86_64"),
        ("Linux", "aarch64"),
        ("Darwin", "x86_64"),
        ("Darwin", "arm64"),
        ("Windows", "AMD64"),
    ],
)
@pytest.mark.parametrize("minor", [11, 12, 13, 14])
def test_supported_cpython_defaults_to_native(
    system: str,
    machine: str,
    minor: int,
) -> None:
    build_support = _load_build_support()

    assert (
        build_support.resolve_build_mode(
            {},
            implementation="cpython",
            version_info=(3, minor),
            system=system,
            machine=machine,
            free_threaded=False,
        )
        == "native"
    )


@pytest.mark.parametrize(
    (
        "implementation",
        "version_info",
        "system",
        "machine",
        "free_threaded",
    ),
    [
        ("pypy", (3, 11), "Linux", "x86_64", False),
        ("cpython", (3, 10), "Linux", "x86_64", False),
        ("cpython", (3, 15), "Linux", "x86_64", False),
        ("cpython", (3, 14), "Linux", "x86_64", True),
        ("cpython", (3, 14), "Linux", "ppc64le", False),
        ("cpython", (3, 14), "Darwin", "aarch64", False),
        ("cpython", (3, 14), "Windows", "ARM64", False),
        ("cpython", (3, 14), "FreeBSD", "x86_64", False),
    ],
)
def test_unsupported_runtime_defaults_to_fallback(
    implementation: str,
    version_info: tuple[int, int],
    system: str,
    machine: str,
    free_threaded: bool,
) -> None:
    build_support = _load_build_support()

    assert (
        build_support.resolve_build_mode(
            {},
            implementation=implementation,
            version_info=version_info,
            system=system,
            machine=machine,
            free_threaded=free_threaded,
        )
        == "fallback"
    )
