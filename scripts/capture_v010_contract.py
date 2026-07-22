from __future__ import annotations

import inspect
import json
import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from messy_xlsx import (  # noqa: E402
    MessyWorkbook,
    SheetConfig,
    analyze_excel,
    analyze_structure,
    read_all_sheets,
    read_excel,
)
from messy_xlsx.models import CellValue  # noqa: E402
from messy_xlsx.utils import coords_to_cell_ref  # noqa: E402
from tests.compatibility._contract import exception_contract, frame_contract  # noqa: E402

SAMPLES = sorted((ROOT / "tests" / "samples").glob("*.xlsx"))
GOLDEN = ROOT / "tests" / "compatibility" / "golden"
OUTPUT = GOLDEN / "v010-frames.json"

FEATURE_SAMPLES = {
    "merged": ROOT
    / "tests/generated_messy/parseable"
    / (
        "messy__preset_financial_statements_summary_sheet__seed_1004__"
        "metadata_preamble_merged_headers_irrelevant_summary_sheet_hidden_rows.xlsx"
    ),
    "hidden": ROOT
    / "tests/generated_messy/parseable"
    / (
        "messy__preset_expense_reports_date_unicode_hidden__seed_1006__"
        "date_noise_unicode_whitespace_noise_hidden_cols_blank_row_noise.xlsx"
    ),
    "formula": ROOT
    / "tests/generated_messy/parseable"
    / (
        "messy__preset_cash_flow_offset_formula_multitable__seed_1008__"
        "offset_table_formula_noise_multi_table_sheet_footer_noise.xlsx"
    ),
    "multi_table": ROOT
    / "tests/generated_messy/parseable"
    / (
        "messy__preset_inventory_multitable_ragged_hidden__seed_1010__"
        "multi_table_sheet_ragged_rows_hidden_cols_blank_row_noise.xlsx"
    ),
    "multi_sheet": ROOT / "tests/samples/financial_statements.xlsx",
}


def _relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def _json_native(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, dict):
        return {str(key): _json_native(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_native(item) for item in value]
    return repr(value)


def _cell_value(cell: CellValue) -> dict[str, Any]:
    return _json_native(asdict(cell))


def _public_signatures() -> dict[str, str]:
    import messy_xlsx
    from messy_xlsx.parsing import FormatHandler, HandlerRegistry
    from messy_xlsx.sheet import MessySheet, MessyTable

    targets = {
        f"messy_xlsx.{name}": getattr(messy_xlsx, name)
        for name in messy_xlsx.__all__
        if callable(getattr(messy_xlsx, name))
    }
    for owner in (
        messy_xlsx.MessyWorkbook,
        messy_xlsx.MultiSheetParser,
        MessySheet,
        MessyTable,
        HandlerRegistry,
        FormatHandler,
    ):
        for name, member in inspect.getmembers(owner, inspect.isfunction):
            if not name.startswith("_"):
                targets[f"{owner.__module__}.{owner.__qualname__}.{name}"] = member
    return {name: str(inspect.signature(value)) for name, value in sorted(targets.items())}


def _error_contracts(sample: Path) -> dict[str, object]:
    missing = ROOT / "tests" / "samples" / "missing.xlsx"
    malformed = (
        ROOT
        / "tests"
        / "generated_messy"
        / "malformed"
        / ("messy__preset_malformed_missing_workbook_xml__seed_1020__missing_workbook_xml.xlsx")
    )
    unsupported = GOLDEN / "unsupported.bin"
    with MessyWorkbook(sample) as workbook:
        missing_sheet = exception_contract(lambda: workbook.to_dataframe("missing-sheet"))
        invalid_range = exception_contract(
            lambda: workbook.to_dataframe(
                config=SheetConfig(auto_detect=False, cell_range="invalid-range")
            )
        )
    unsupported.write_bytes(b"\x00" * 32 + b"not a supported spreadsheet")
    try:
        unsupported_contract = exception_contract(lambda: MessyWorkbook(unsupported))
    finally:
        unsupported.unlink(missing_ok=True)
    return {
        "missing_file": exception_contract(lambda: MessyWorkbook(missing)),
        "missing_sheet": missing_sheet,
        "invalid_range": invalid_range,
        "malformed_ooxml": exception_contract(lambda: MessyWorkbook(malformed)),
        "unsupported_format": unsupported_contract,
    }


def _structure_contracts() -> dict[str, object]:
    contracts: dict[str, object] = {}
    for fixture_name, path in FEATURE_SAMPLES.items():
        contracts[fixture_name] = {
            "path": _relative(path),
            "analyze_structure": _json_native(asdict(analyze_structure(str(path)))),
        }
    multi_sheet = FEATURE_SAMPLES["multi_sheet"]
    multi_contract = contracts["multi_sheet"]
    assert isinstance(multi_contract, dict)
    multi_contract["analyze_excel"] = [
        _json_native(asdict(info)) for info in analyze_excel(multi_sheet)
    ]
    return contracts


def _target_cell(fixture_name: str, workbook: MessyWorkbook, sheet_name: str) -> tuple[int, int]:
    structure = workbook.get_structure(sheet_name)
    if fixture_name == "merged" and structure.merged_ranges:
        start_row, start_col, _, _ = structure.merged_ranges[0]
        return start_row, start_col
    if fixture_name == "hidden":
        if structure.hidden_rows:
            return structure.hidden_rows[0], structure.data_start_col
        if structure.hidden_columns:
            return structure.data_start_row, structure.hidden_columns[0]
    if fixture_name == "formula":
        for row in range(structure.data_start_row, structure.data_end_row + 1):
            for col in range(structure.data_start_col, structure.data_end_col + 1):
                if workbook.get_cell(sheet_name, row, col).formula is not None:
                    return row, col
    if fixture_name == "multi_table" and structure.table_ranges:
        first = structure.table_ranges[0]
        return int(first["start_row"]), int(first["start_col"])
    return structure.data_start_row, structure.data_start_col


def _cell_contracts() -> dict[str, object]:
    contracts: dict[str, object] = {}
    for fixture_name, path in FEATURE_SAMPLES.items():
        with MessyWorkbook(path) as workbook:
            sheet_name = workbook.sheet_names[0]
            sheet = workbook.get_sheet(sheet_name)
            structure = sheet.structure
            row, col = _target_cell(fixture_name, workbook, sheet_name)
            ref = coords_to_cell_ref(row, col, sheet_name)
            max_row = min(structure.data_end_row, structure.data_start_row + 2)
            max_col = min(structure.data_end_col, structure.data_start_col + 3)
            rows = [
                [_cell_value(cell) for cell in cells]
                for cells in sheet.iter_rows(
                    min_row=structure.data_start_row,
                    max_row=max_row,
                    min_col=structure.data_start_col,
                    max_col=max_col,
                )
            ]
            tables = [
                {
                    "start_row": table.start_row,
                    "end_row": table.end_row,
                    "row_count": table.row_count,
                    "column_count": table.column_count,
                    "frame": frame_contract(table.to_dataframe()),
                }
                for table in sheet.tables
            ]
            contracts[fixture_name] = {
                "path": _relative(path),
                "sheet": sheet_name,
                "target_ref": ref,
                "get_cell": _cell_value(workbook.get_cell(sheet_name, row, col)),
                "get_cell_by_ref": _cell_value(workbook.get_cell_by_ref(ref)),
                "iter_rows": rows,
                "table_ranges": _json_native(asdict(structure)["table_ranges"]),
                "tables": tables,
            }
    return contracts


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True), encoding="utf-8")


def _write_auxiliary_contracts() -> None:
    _write_json(GOLDEN / "v010-structures.json", _structure_contracts())
    _write_json(GOLDEN / "v010-cells.json", _cell_contracts())
    _write_json(GOLDEN / "v010-errors.json", _error_contracts(SAMPLES[0]))
    _write_json(GOLDEN / "v010-signatures.json", _public_signatures())


def main() -> None:
    contract: dict[str, object] = {"version": "0.10.0", "samples": {}}
    samples = contract["samples"]
    assert isinstance(samples, dict)
    for path in SAMPLES:
        with MessyWorkbook(path) as workbook:
            sheets = {
                name: frame_contract(workbook.to_dataframe(name)) for name in workbook.sheet_names
            }
        samples[path.name] = {
            "default": frame_contract(read_excel(str(path))),
            "workbook_sheets": sheets,
        }
    multi = ROOT / "tests" / "samples" / "financial_statements.xlsx"
    contract["read_all_sheets"] = {
        name: frame_contract(frame) for name, frame in read_all_sheets(multi).items()
    }
    _write_json(OUTPUT, contract)
    _write_auxiliary_contracts()


if __name__ == "__main__":
    main()
