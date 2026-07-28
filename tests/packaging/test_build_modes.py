from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tarfile
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]


@dataclass(frozen=True)
class BuildResult:
    process: subprocess.CompletedProcess[str]
    distributions: tuple[Path, ...]


def _copy_source(destination: Path) -> Path:
    return Path(
        shutil.copytree(
            PROJECT_ROOT,
            destination,
            ignore=shutil.ignore_patterns(
                ".git",
                ".venv",
                ".worktrees",
                ".superpowers",
                ".pytest_cache",
                ".ruff_cache",
                ".mypy_cache",
                "__pycache__",
                "build",
                "dist",
                "*.egg-info",
                "*.pyc",
                "*.pyd",
                "*.so",
                "CONTINUE.md",
                "uv.lock",
            ),
        )
    )


def _run_build(source: Path, output: Path, mode: str, *kinds: str) -> BuildResult:
    env = os.environ.copy()
    env["MESSY_XLSX_BUILD_MODE"] = mode
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    process = subprocess.run(
        [
            sys.executable,
            "-m",
            "build",
            *kinds,
            "--outdir",
            str(output),
        ],
        cwd=source,
        env=env,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=300,
        check=False,
    )
    return BuildResult(process=process, distributions=tuple(sorted(output.glob("*"))))


@pytest.fixture(scope="module")
def mode_builds(tmp_path_factory: pytest.TempPathFactory) -> dict[str, BuildResult]:
    build_root = tmp_path_factory.mktemp("mode-builds")
    native_source = _copy_source(build_root / "native-source")
    fallback_source = _copy_source(build_root / "fallback-source")
    return {
        "native": _run_build(
            native_source,
            build_root / "native-dist",
            "native",
            "--wheel",
        ),
        "fallback": _run_build(
            fallback_source,
            build_root / "fallback-dist",
            "fallback",
            "--sdist",
            "--wheel",
        ),
    }


def _only_wheel(result: BuildResult) -> Path:
    assert result.process.returncode == 0, result.process.stdout
    wheels = [path for path in result.distributions if path.suffix == ".whl"]
    assert len(wheels) == 1
    return wheels[0]


def _wheel_metadata(wheel: Path, filename: str) -> str:
    with ZipFile(wheel) as archive:
        matches = [name for name in archive.namelist() if name.endswith(f".dist-info/{filename}")]
        assert len(matches) == 1
        return archive.read(matches[0]).decode("utf-8")


def test_native_mode_builds_cp311_abi3_extension(
    mode_builds: dict[str, BuildResult],
) -> None:
    wheel = _only_wheel(mode_builds["native"])

    assert "-cp311-abi3-" in wheel.name
    with ZipFile(wheel) as archive:
        names = archive.namelist()
    assert any(
        name.startswith("messy_xlsx/_csv_tokenizer") and name.endswith((".abi3.so", ".pyd"))
        for name in names
    )
    wheel_metadata = _wheel_metadata(wheel, "WHEEL")
    assert "Root-Is-Purelib: false" in wheel_metadata
    assert "Tag: cp311-abi3-" in wheel_metadata


def test_fallback_mode_builds_universal_wheel_without_extension(
    mode_builds: dict[str, BuildResult],
) -> None:
    wheel = _only_wheel(mode_builds["fallback"])

    assert wheel.name.endswith("-py3-none-any.whl")
    with ZipFile(wheel) as archive:
        names = archive.namelist()
    assert not any("_csv_tokenizer" in name or name.endswith((".so", ".pyd")) for name in names)
    wheel_metadata = _wheel_metadata(wheel, "WHEEL")
    assert "Root-Is-Purelib: true" in wheel_metadata
    assert "Tag: py3-none-any" in wheel_metadata


def test_native_and_fallback_metadata_pin_pandas_3_0_5(
    mode_builds: dict[str, BuildResult],
) -> None:
    native_metadata = _wheel_metadata(_only_wheel(mode_builds["native"]), "METADATA")
    fallback_metadata = _wheel_metadata(_only_wheel(mode_builds["fallback"]), "METADATA")

    assert "Requires-Dist: pandas==3.0.5" in native_metadata
    assert native_metadata == fallback_metadata


def test_source_distribution_contains_native_source(
    mode_builds: dict[str, BuildResult],
) -> None:
    result = mode_builds["fallback"]
    assert result.process.returncode == 0, result.process.stdout
    sdists = [path for path in result.distributions if path.name.endswith(".tar.gz")]
    assert len(sdists) == 1

    with tarfile.open(sdists[0], "r:gz") as archive:
        names = archive.getnames()
    assert any(name.endswith("/src/messy_xlsx/_csv_tokenizer.pyx") for name in names)
    assert any(name.endswith("/requirements/native-release.txt") for name in names)


def test_invalid_build_mode_stops_the_build(
    tmp_path: Path,
) -> None:
    source = _copy_source(tmp_path / "invalid-source")
    result = _run_build(source, tmp_path / "invalid-dist", "auto", "--wheel")

    assert result.process.returncode != 0, result.process.stdout
    assert "MESSY_XLSX_BUILD_MODE" in result.process.stdout
