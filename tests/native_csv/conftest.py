from __future__ import annotations

import importlib
from hashlib import sha256
from pathlib import Path

import pytest

PYX_PATH = Path(__file__).resolve().parents[2] / "src" / "messy_xlsx" / "_csv_tokenizer.pyx"


@pytest.fixture(autouse=True)
def require_fresh_native_extension() -> None:
    try:
        native = importlib.import_module("messy_xlsx._csv_tokenizer")
    except ModuleNotFoundError:
        # The test body owns the ordinary missing-extension assertion. This
        # fixture's specific job is to reject an importable but stale binary.
        return
    except (ImportError, OSError) as exc:
        pytest.fail(f"native CSV extension is not importable: {exc}")

    expected_hash = sha256(PYX_PATH.read_bytes()).hexdigest()
    assert expected_hash == native.NATIVE_SOURCE_SHA256, (
        "stale native CSV extension loaded: "
        f"expected source hash {expected_hash}, got {native.NATIVE_SOURCE_SHA256}"
    )
