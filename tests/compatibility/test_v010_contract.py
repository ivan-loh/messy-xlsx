from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "compatibility" / "golden"
FRAMES = json.loads((GOLDEN / "v010-frames.json").read_text(encoding="utf-8"))

pytestmark = pytest.mark.compatibility


def _json(name: str) -> object:
    return json.loads((GOLDEN / name).read_text(encoding="utf-8"))


def test_frame_contract_serializes_deterministically() -> None:
    contract_module = ROOT / "tests" / "compatibility" / "_contract.py"
    assert contract_module.is_file(), "the v0.10.0 contract serializer is not implemented"

    from ._contract import frame_contract

    frame = pd.DataFrame(
        [[1, None], [2, "é"]],
        columns=["amount", 7],
        index=["row", 3],
    )

    assert frame_contract(frame) == {
        "shape": [2, 2],
        "columns": [
            {"type": "str", "repr": "'amount'"},
            {"type": "int", "repr": "7"},
        ],
        "index": [
            {"type": "str", "repr": "'row'"},
            {"type": "int", "repr": "3"},
        ],
        "dtypes": ["int64", "str"],
        "value_sha256": "be5abcb94b67033ffd4687600819e8b901cae8f1b151bd0ce20521db007c6f0a",
    }


def test_v010_contract_assets_are_versioned() -> None:
    expected = {
        "v010-frames.json",
        "v010-structures.json",
        "v010-cells.json",
        "v010-errors.json",
        "v010-signatures.json",
    }
    assert {path.name for path in GOLDEN.glob("v010-*.json")} == expected
    assert FRAMES["version"] == "0.10.0"


@pytest.mark.parametrize("sample_name", sorted(FRAMES["samples"]))
def test_default_frames_match_v010_contract(sample_name: str) -> None:
    from messy_xlsx import MessyWorkbook, read_excel

    from ._contract import frame_contract

    path = ROOT / "tests" / "samples" / sample_name
    assert frame_contract(read_excel(str(path))) == FRAMES["samples"][sample_name]["default"]
    with MessyWorkbook(path) as workbook:
        actual = {
            name: frame_contract(workbook.to_dataframe(name)) for name in workbook.sheet_names
        }
    assert actual == FRAMES["samples"][sample_name]["workbook_sheets"]


def test_read_all_sheets_matches_v010_contract() -> None:
    from messy_xlsx import read_all_sheets

    from ._contract import frame_contract

    path = ROOT / "tests" / "samples" / "financial_statements.xlsx"
    actual = {name: frame_contract(frame) for name, frame in read_all_sheets(path).items()}
    assert actual == FRAMES["read_all_sheets"]


def test_missing_sheet_exception_matches_v010(sample_xlsx: Path) -> None:
    from messy_xlsx import MessyWorkbook

    from ._contract import exception_contract

    errors = _json("v010-errors.json")
    with MessyWorkbook(sample_xlsx) as workbook:
        actual = exception_contract(lambda: workbook.to_dataframe("missing-sheet"))
    assert actual == errors["missing_sheet"]


def test_error_contracts_match_v010() -> None:
    from scripts.capture_v010_contract import _error_contracts

    errors = _json("v010-errors.json")
    sample = ROOT / "tests" / "samples" / "accounts_receivable.xlsx"
    assert _error_contracts(sample) == errors


def test_unsupported_binary_payload_raises_format_error(tmp_path: Path) -> None:
    from messy_xlsx import MessyWorkbook

    from ._contract import exception_contract

    unsupported = tmp_path / "unsupported.bin"
    unsupported.write_bytes(b"\x00" * 32 + b"not a supported spreadsheet")

    actual = exception_contract(lambda: MessyWorkbook(unsupported))

    assert actual["type"] == "messy_xlsx.exceptions.FormatError"
    assert "Unknown file format" in actual["message"]
    assert actual["context"]["file_path"] == "unsupported.bin"


def test_existing_public_signatures_match_v010() -> None:
    from scripts.capture_v010_contract import _public_signatures

    expected = _json("v010-signatures.json")
    current = _public_signatures()
    assert {name: current[name] for name in expected} == expected


def test_structure_and_multi_sheet_analysis_match_v010() -> None:
    from scripts.capture_v010_contract import _structure_contracts

    expected = _json("v010-structures.json")
    assert _structure_contracts() == expected


def test_cell_rows_and_tables_match_v010() -> None:
    from scripts.capture_v010_contract import _cell_contracts

    expected = _json("v010-cells.json")
    assert _cell_contracts() == expected


def test_reference_benchmark_metadata_matches_measured_v010() -> None:
    assert json.loads(
        (ROOT / "benchmarks" / "v010-reference.json").read_text(encoding="utf-8")
    ) == {
        "version": "0.10.0",
        "xlsx_100k": {"elapsed_seconds": 9.99, "peak_rss_mb": 627},
        "csv_300k_path": {"normalized_seconds": 1.58, "peak_rss_mb": 267},
        "csv_300k_seekable": {"normalized_seconds": 1.68, "peak_rss_mb": 352},
        "multi_sheet": {
            "to_dataframes_openpyxl_loads": 6,
            "read_all_sheets_openpyxl_loads": 9,
        },
    }
