"""Tests for package metadata and models."""

from pathlib import Path


def test_version_matches_pyproject():
    """Ensure __version__ in __init__.py matches pyproject.toml."""
    import tomllib

    import messy_xlsx

    pyproject_path = Path(__file__).parent.parent / "pyproject.toml"
    with open(pyproject_path, "rb") as f:
        pyproject = tomllib.load(f)

    assert messy_xlsx.__version__ == pyproject["project"]["version"], (
        f"Version mismatch: __init__.py has {messy_xlsx.__version__!r}, "
        f"pyproject.toml has {pyproject['project']['version']!r}"
    )


def test_py_typed_exists():
    """Ensure py.typed marker exists for PEP 561 compliance."""
    import messy_xlsx

    package_dir = Path(messy_xlsx.__file__).parent
    py_typed = package_dir / "py.typed"
    assert py_typed.exists(), f"py.typed marker not found at {py_typed}"
