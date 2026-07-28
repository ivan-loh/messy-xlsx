from __future__ import annotations

from hashlib import sha256
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

from setuptools import Extension, setup

PROJECT_ROOT = Path(__file__).resolve().parent
NATIVE_SOURCE = PROJECT_ROOT / "src" / "messy_xlsx" / "_csv_tokenizer.pyx"
BUILD_SUPPORT = spec_from_file_location(
    "messy_xlsx_build_support",
    PROJECT_ROOT / "build_support.py",
)
if BUILD_SUPPORT is None or BUILD_SUPPORT.loader is None:
    raise RuntimeError("cannot load build_support.py")
BUILD_SUPPORT_MODULE = module_from_spec(BUILD_SUPPORT)
BUILD_SUPPORT.loader.exec_module(BUILD_SUPPORT_MODULE)
resolve_build_mode = BUILD_SUPPORT_MODULE.resolve_build_mode


def _native_extensions() -> list[Extension]:
    from Cython.Build import cythonize

    source_hash = sha256(NATIVE_SOURCE.read_bytes()).hexdigest()
    extension = Extension(
        "messy_xlsx._csv_tokenizer",
        [str(NATIVE_SOURCE.relative_to(PROJECT_ROOT))],
        define_macros=[("Py_LIMITED_API", "0x030B0000")],
        py_limited_api=True,
    )
    return cythonize(
        [extension],
        build_dir=str(PROJECT_ROOT / "build" / "cython"),
        compile_time_env={"NATIVE_SOURCE_SHA256_VALUE": source_hash},
        compiler_directives={"language_level": 3},
        force=True,
    )


setup(
    ext_modules=_native_extensions() if resolve_build_mode() == "native" else [],
)
