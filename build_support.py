"""Deterministic native/fallback build-mode selection."""

from __future__ import annotations

import os
import platform
import sys
import sysconfig
from collections.abc import Mapping
from typing import Literal

BuildMode = Literal["native", "fallback"]
_BUILD_MODE_ENV = "MESSY_XLSX_BUILD_MODE"
_SUPPORTED_PLATFORMS = frozenset(
    {
        ("linux", "x86_64"),
        ("linux", "aarch64"),
        ("darwin", "x86_64"),
        ("darwin", "arm64"),
        ("windows", "amd64"),
    }
)


def _runtime_is_free_threaded() -> bool:
    return bool(sysconfig.get_config_var("Py_GIL_DISABLED"))


def resolve_build_mode(
    environ: Mapping[str, str] | None = None,
    *,
    implementation: str | None = None,
    version_info: tuple[int, int] | None = None,
    system: str | None = None,
    machine: str | None = None,
    free_threaded: bool | None = None,
) -> BuildMode:
    """Resolve the only two supported build modes, failing closed on bad input."""

    selected_environ = os.environ if environ is None else environ
    if _BUILD_MODE_ENV in selected_environ:
        explicit_mode = selected_environ[_BUILD_MODE_ENV]
        if explicit_mode not in {"native", "fallback"}:
            raise ValueError(
                f"{_BUILD_MODE_ENV} must be exactly 'native' or 'fallback'; got {explicit_mode!r}"
            )
        return explicit_mode

    runtime_implementation = sys.implementation.name if implementation is None else implementation
    runtime_version = (
        (sys.version_info.major, sys.version_info.minor) if version_info is None else version_info
    )
    runtime_system = platform.system() if system is None else system
    runtime_machine = platform.machine() if machine is None else machine
    runtime_free_threaded = _runtime_is_free_threaded() if free_threaded is None else free_threaded

    supported = (
        runtime_implementation.casefold() == "cpython"
        and runtime_version[0] == 3
        and 11 <= runtime_version[1] <= 14
        and not runtime_free_threaded
        and (runtime_system.casefold(), runtime_machine.casefold()) in _SUPPORTED_PLATFORMS
    )
    return "native" if supported else "fallback"
