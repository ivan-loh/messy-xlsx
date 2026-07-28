# Parser Performance v1.0.0 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship messy-xlsx v1.0.0 with materially faster materialized parsing, bounded-row streaming APIs, lower stream memory, and complete v0.10.0 public compatibility.

**Architecture:** Ordinary materialized OOXML parsing uses fastexcel and converges on Arrow before the legacy pandas adapter. Bounded-row OOXML iteration uses openpyxl read-only, while CSV and XLS use native chunk/row readers; every streaming path yields coordinate-aware Arrow `RecordBatch` objects through one closable lifecycle. A lazy, hardened OOXML manifest supplies structural metadata once per workbook, and a spillable source boundary prevents unbounded stream copies.

**Tech Stack:** Python 3.11-3.14, pandas 3.x, pyarrow 23+, fastexcel 0.19+ with 0.20.2 as the design baseline, openpyxl 3.1.5+, defusedxml 0.7.1+, optional xlrd 2.0.2+, pytest 9+, Hypothesis, Ruff, mypy, MkDocs, GitHub Actions.

## Global Constraints

- Target version is exactly `1.0.0`; the release tag is exactly `v1.0.0`.
- Every v0.10.0 public API remains callable throughout v1.x with the same signature, return shape, values, columns, dtypes, indexes, parsing defaults, and public exception contracts.
- Materialized DataFrame entry points emit exactly one caller-facing `LegacyAPIWarning`, derived from `DeprecationWarning`; extension SPI calls do not leak warnings.
- `pyarrow.RecordBatch` is the canonical streaming and transform representation; `pyarrow.Table` is the materialized Arrow container.
- Fastexcel 0.20.2 is materialized-only. It must never be labeled bounded streaming, and repeated `skip_rows` windows are not an acceptable streaming implementation.
- OOXML streaming has a bounded row working set but may retain openpyxl shared-string/style tables; this format-level overhead must be capped, measured, and reported separately.
- A source replay spool keeps at most 8 MiB in memory and spills larger sources to a private mode-`0600` temporary path.
- Caller-owned streams are never closed and are restored to their entry cursor on success, failure, and early close.
- One `MessyWorkbook` supports one active parse or stream; concurrent and re-entrant use raises `RuntimeError`.
- No public backend selector is added in v1.0.0.
- Any custom registry, registry subclass, detector override, or custom handler uses the compatibility SPI path.
- Legacy materialized normalization remains exact. Streaming inference is bounded and raises `StreamingTypeError` on a late incompatible value.
- OOXML parsing performs no network access, follows no external relationship, and enforces the security budgets in the design spec.
- All behavior-changing tasks use red-green-refactor, focused review, and a dedicated commit.
- Design authority: `docs/superpowers/specs/2026-07-22-parser-performance-design.md`.

---

## File and Responsibility Map

### New runtime modules

- `src/messy_xlsx/warnings.py` — public legacy warning category and single-boundary warning helper.
- `src/messy_xlsx/_spool.py` — spillable replay source and secure temporary-path lifecycle.
- `src/messy_xlsx/ooxml/models.py` — immutable manifest, sheet metadata, interval, and merge models.
- `src/messy_xlsx/ooxml/security.py` — ZIP/XML budgets and relationship-target validation.
- `src/messy_xlsx/ooxml/manifest.py` — eager workbook manifest and lazy per-sheet metadata parsing.
- `src/messy_xlsx/detection/structure_sampler.py` — bounded, reusable structure analysis without discarded DataFrames.
- `src/messy_xlsx/parsing/contracts.py` — output modes, reader decisions, metrics, and reader protocols.
- `src/messy_xlsx/parsing/fastexcel_session.py` — one workbook-scoped fastexcel reader shared by sampling and materialization.
- `src/messy_xlsx/parsing/router.py` — capability/output-mode routing and custom-registry escape hatch.
- `src/messy_xlsx/parsing/fallback.py` — transactional classified fallback before observable output.
- `src/messy_xlsx/parsing/coordinates.py` — ordinal column identity and coordinate-aware batch transforms.
- `src/messy_xlsx/parsing/streams.py` — shared closable iterator lifecycle plus public stream classes.
- `src/messy_xlsx/parsing/xlsx_materialized.py` — fastexcel materialized Arrow reader.
- `src/messy_xlsx/parsing/legacy_adapter.py` — Arrow-to-pandas compatibility authority for existing APIs.
- `src/messy_xlsx/parsing/xlsx_streaming.py` — openpyxl read-only Arrow batch reader.
- `src/messy_xlsx/parsing/csv_streaming.py` — bounded CSV/TSV/TXT inspection
  and Arrow adapter shell; the full pass is the internal native reader.
- `src/messy_xlsx/parsing/xls_streaming.py` — optional xlrd row-window Arrow reader.
- `src/messy_xlsx/parsing/sheet_planner.py` — one sheet-selection and per-sheet configuration authority.
- `src/messy_xlsx/normalization/plan.py` — immutable normalization schema and late-value policy.
- `src/messy_xlsx/normalization/arrow_pipeline.py` — stream-safe Arrow normalization.
- `src/messy_xlsx/cell_index.py` — compact merge and hidden-coordinate lookup for cell APIs.

### Existing runtime modules to modify

- `src/messy_xlsx/__init__.py` — public exports and top-level Arrow/batch conveniences.
- `src/messy_xlsx/_source.py` — spool integration and borrow/re-entrancy rules.
- `src/messy_xlsx/cache.py` — robust path identity and pre/post-stat validation.
- `src/messy_xlsx/exceptions.py` — public `StreamingTypeError`.
- `src/messy_xlsx/models.py` — frozen `SheetResult`.
- `src/messy_xlsx/workbook.py` — orchestration, lifecycle, public APIs, and workbook-local caches.
- `src/messy_xlsx/sheet.py` — legacy warnings and indexed cell compatibility.
- `src/messy_xlsx/multi_sheet.py` — shared planning path and warning-safe legacy adapters.
- `src/messy_xlsx/parsing/parse_plan.py` — immutable configuration snapshot and output-mode fields.
- `src/messy_xlsx/parsing/handler_registry.py` — explicit compatibility routing and transactional fallback.
- `src/messy_xlsx/parsing/xlsx_handler.py` — legacy adapter over materialized reader.
- `src/messy_xlsx/parsing/csv_handler.py` — direct binary-stream parsing and legacy adapter.
- `src/messy_xlsx/parsing/xls_handler.py` — legacy adapter over optional XLS reader.
- `src/messy_xlsx/normalization/pipeline.py` and normalizers — consolidate legacy copies without changing output.

### Test, benchmark, documentation, and release files

- `tests/compatibility/` — v0.10.0 golden contracts and public API inventory.
- `tests/test_spool.py`, `tests/test_ooxml_manifest.py`, `tests/test_ooxml_security.py`.
- `tests/test_reader_routing.py`, `tests/test_coordinate_transforms.py`, `tests/test_stream_lifecycle.py`.
- `tests/test_xlsx_streaming.py`, `tests/test_streaming_normalization.py`, `tests/test_arrow_api.py`.
- `tests/test_multi_sheet_streaming.py`, `tests/test_csv_streaming.py`, `tests/test_xls_streaming.py`.
- `tests/test_cell_indexes.py`, existing regression/property/resource suites.
- `scripts/capture_v010_contract.py`, `scripts/benchmark_worker.py`, `scripts/compare_benchmarks.py`.
- `benchmarks/v010-reference.json`, `.github/workflows/performance.yml`, `.github/workflows/test.yml`.
- `README.md`, `docs/index.md`, `docs/getting-started.md`, `docs/configuration.md`, `docs/api.md`, `CHANGELOG.md`, `pyproject.toml`.

## Progress Tracker

| Slice | Deliverable | Depends on | Status |
|---:|---|---|:---:|
| 1 | v0.10.0 compatibility and performance contract | — | [x] |
| 2 | Legacy warning and API classification | 1 | [x] |
| 3 | Spillable source lifecycle | 1 | [x] |
| 4 | OOXML archive security and eager manifest | 3 | [x] |
| 5 | Lazy sheet metadata, interval indexes, and bounded structure sampling | 4 | [x] |
| 6 | Immutable plans, reader contracts, and router | 2, 5 | [x] |
| 7 | Fastexcel materialized Arrow reader | 6 | [x] |
| 8 | Coordinate-aware Arrow transforms | 5, 7 | [x] |
| 9 | Closable stream lifecycle | 3, 6 | [x] |
| 10 | Openpyxl bounded-row OOXML reader | 8, 9 | [x] |
| 11 | Streaming normalization and schema enforcement | 9, 10 | [x] |
| 12 | Public Arrow, batch, and pandas-chunk APIs | 7, 9, 11 | [x] |
| 13 | Unified multi-sheet planning and `SheetStream` | 5, 7, 12 | [x] |
| 14 | CSV/TXT direct-stream and batch optimization | 3, 9, 11 | [ ] |
| 15 | XLS streaming and custom registry compatibility | 6, 9, 12 | [ ] |
| 16 | Legacy normalization copy reduction | 1, 7 | [ ] |
| 17 | Cell-access indexes and formula boundary | 5 | [ ] |
| 18 | Security, failure, property, and corpus integration | 10-17 | [ ] |
| 19 | Reproducible performance CI and acceptance gates | 18 | [ ] |
| 20 | Documentation, versioning, packaging, and release gate | 19 | [ ] |

---

### Task 1: Freeze the v0.10.0 Compatibility and Performance Contract

**Files:**
- Create: `tests/compatibility/__init__.py`
- Create: `tests/compatibility/_contract.py`
- Create: `tests/compatibility/test_v010_contract.py`
- Create: `tests/compatibility/golden/v010-frames.json`
- Create: `tests/compatibility/golden/v010-structures.json`
- Create: `tests/compatibility/golden/v010-cells.json`
- Create: `tests/compatibility/golden/v010-errors.json`
- Create: `tests/compatibility/golden/v010-signatures.json`
- Create: `scripts/capture_v010_contract.py`
- Create: `benchmarks/v010-reference.json`
- Modify: `pyproject.toml`

**Interfaces:**
- Produces: `frame_contract(df: pd.DataFrame) -> dict[str, object]` and immutable v0.10.0 golden records used by every replacement task.
- Produces: pytest marker `compatibility` and reference performance metadata for the maintained 100,000-row XLSX sample and 300,000-row generated CSV.

- [x] **Step 1: Add deterministic DataFrame contract serialization**

```python
# tests/compatibility/_contract.py
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pandas as pd


def _label(value: object) -> dict[str, str]:
    return {"type": type(value).__qualname__, "repr": repr(value)}


def frame_contract(frame: pd.DataFrame) -> dict[str, Any]:
    normalized = frame.astype(object).where(frame.notna(), None)
    records = normalized.to_dict(orient="split")
    payload = json.dumps(records, default=str, ensure_ascii=False, sort_keys=True)
    return {
        "shape": list(frame.shape),
        "columns": [_label(value) for value in frame.columns],
        "index": [_label(value) for value in frame.index],
        "dtypes": [str(dtype) for dtype in frame.dtypes],
        "value_sha256": hashlib.sha256(payload.encode("utf-8")).hexdigest(),
    }


def exception_contract(callable_object) -> dict[str, Any]:
    try:
        callable_object()
    except Exception as error:
        context = dict(getattr(error, "context", {}))
        message = str(error)
        if "file_path" in context:
            original = str(context["file_path"])
            normalized = Path(original).name
            context["file_path"] = normalized
            message = message.replace(original, normalized)
        return {
            "type": f"{type(error).__module__}.{type(error).__qualname__}",
            "message": message,
            "context": context,
        }
    raise AssertionError("expected callable to raise")
```

- [x] **Step 2: Add and run the baseline capture script before runtime changes**

```python
# scripts/capture_v010_contract.py
from __future__ import annotations

import json
from pathlib import Path

from messy_xlsx import MessyWorkbook, read_all_sheets, read_excel
from tests.compatibility._contract import exception_contract, frame_contract


ROOT = Path(__file__).resolve().parents[1]
SAMPLES = sorted((ROOT / "tests" / "samples").glob("*.xlsx"))
OUTPUT = ROOT / "tests" / "compatibility" / "golden" / "v010-frames.json"


def main() -> None:
    contract: dict[str, object] = {"version": "0.10.0", "samples": {}}
    samples = contract["samples"]
    assert isinstance(samples, dict)
    for path in SAMPLES:
        with MessyWorkbook(path) as workbook:
            sheets = {
                name: frame_contract(workbook.to_dataframe(name))
                for name in workbook.sheet_names
            }
        samples[path.name] = {
            "default": frame_contract(read_excel(str(path))),
            "workbook_sheets": sheets,
        }
    multi = ROOT / "tests" / "samples" / "financial_statements.xlsx"
    contract["read_all_sheets"] = {
        name: frame_contract(frame) for name, frame in read_all_sheets(multi).items()
    }
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(contract, indent=2, sort_keys=True), encoding="utf-8")


if __name__ == "__main__":
    main()
```

Run: `.venv/bin/python scripts/capture_v010_contract.py`

Expected: `tests/compatibility/golden/v010-frames.json` records version `0.10.0` and every maintained XLSX sample.

Extend the same script to write `v010-structures.json` for `analyze_structure()`
and `analyze_excel()`, `v010-cells.json` for `get_cell()`, `get_cell_by_ref()`,
`MessySheet.iter_rows()`, and table ranges, and `v010-errors.json` for missing
files, missing sheets, invalid cell ranges, malformed OOXML, and unsupported formats. Write
`v010-signatures.json` using `str(inspect.signature(...))` for every existing
public function, class constructor, workbook method, sheet method, table method,
`MultiSheetParser` method, `HandlerRegistry` method, and `FormatHandler` method.

```python
def _public_signatures() -> dict[str, str]:
    import inspect
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
    from messy_xlsx import MessyWorkbook, SheetConfig

    missing = ROOT / "tests" / "samples" / "missing.xlsx"
    malformed = ROOT / "tests" / "generated_messy" / "malformed" / (
        "messy__preset_malformed_missing_workbook_xml__seed_1020__missing_workbook_xml.xlsx"
    )
    unsupported = ROOT / "tests" / "compatibility" / "golden" / "unsupported.bin"
    with MessyWorkbook(sample) as workbook:
        missing_sheet = exception_contract(lambda: workbook.to_dataframe("missing-sheet"))
        invalid_range = exception_contract(
            lambda: workbook.to_dataframe(
                config=SheetConfig(auto_detect=False, cell_range="invalid-range")
            )
        )
    unsupported.write_bytes(b"not a supported spreadsheet")
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


def _write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True), encoding="utf-8")


def _write_auxiliary_contracts() -> None:
    golden = ROOT / "tests" / "compatibility" / "golden"
    _write_json(golden / "v010-structures.json", _structure_contracts())
    _write_json(golden / "v010-cells.json", _cell_contracts())
    _write_json(golden / "v010-errors.json", _error_contracts(SAMPLES[0]))
    _write_json(golden / "v010-signatures.json", _public_signatures())
```

Add `_structure_contracts()` and `_cell_contracts()` to the same script. Serialize
dataclasses with `dataclasses.asdict()`, serialize every table DataFrame with
`frame_contract()`, and serialize `CellValue` instances as their `asdict()` value
after replacing any non-JSON-native value with `repr()`. Include the maintained
merged, hidden, formula, multi-table, and multi-sheet fixtures. Write their
results to `v010-structures.json` and `v010-cells.json`, then call all four
writers at the end of `main()`. The unsupported-format contract uses a byte
payload whose temporary filename is normalized by `exception_contract()`.

- [x] **Step 3: Add the characterization test**

```python
# tests/compatibility/test_v010_contract.py
import json
from pathlib import Path

import pytest

from messy_xlsx import MessyWorkbook, read_excel
from scripts.capture_v010_contract import (
    _cell_contracts,
    _public_signatures,
    _structure_contracts,
)
from tests.compatibility._contract import exception_contract, frame_contract


ROOT = Path(__file__).resolve().parents[2]
GOLDEN = json.loads(
    (ROOT / "tests/compatibility/golden/v010-frames.json").read_text(encoding="utf-8")
)


@pytest.mark.compatibility
@pytest.mark.parametrize("sample_name", sorted(GOLDEN["samples"]))
def test_default_frames_match_v010_contract(sample_name: str) -> None:
    path = ROOT / "tests" / "samples" / sample_name
    assert frame_contract(read_excel(str(path))) == GOLDEN["samples"][sample_name]["default"]
    with MessyWorkbook(path) as workbook:
        actual = {name: frame_contract(workbook.to_dataframe(name)) for name in workbook.sheet_names}
    assert actual == GOLDEN["samples"][sample_name]["workbook_sheets"]


def test_missing_sheet_exception_matches_v010(sample_xlsx) -> None:
    errors = json.loads(
        (ROOT / "tests/compatibility/golden/v010-errors.json").read_text(encoding="utf-8")
    )
    with MessyWorkbook(sample_xlsx) as workbook:
        actual = exception_contract(lambda: workbook.to_dataframe("missing-sheet"))
    assert actual == errors["missing_sheet"]


def test_existing_public_signatures_match_v010() -> None:
    expected = json.loads(
        (ROOT / "tests/compatibility/golden/v010-signatures.json").read_text(encoding="utf-8")
    )
    current = _public_signatures()
    assert {name: current[name] for name in expected} == expected


def test_structure_and_multi_sheet_analysis_match_v010() -> None:
    expected = json.loads(
        (ROOT / "tests/compatibility/golden/v010-structures.json").read_text(
            encoding="utf-8"
        )
    )
    assert _structure_contracts() == expected


def test_cell_rows_and_tables_match_v010() -> None:
    expected = json.loads(
        (ROOT / "tests/compatibility/golden/v010-cells.json").read_text(
            encoding="utf-8"
        )
    )
    assert _cell_contracts() == expected
```

- [x] **Step 4: Register the compatibility marker and run the frozen contract**

Add `"compatibility: v0.10.0 public behavior contract"` to `tool.pytest.ini_options.markers` in `pyproject.toml`.

Run: `.venv/bin/pytest tests/compatibility/test_v010_contract.py -q`

Expected: all maintained sample, exception, and public-signature cases pass against the freshly captured v0.10.0 contract.

- [x] **Step 5: Record the reference benchmark metadata**

Write `benchmarks/v010-reference.json` from the already measured baseline:

```json
{
  "version": "0.10.0",
  "xlsx_100k": {"elapsed_seconds": 9.99, "peak_rss_mb": 627},
  "csv_300k_path": {"normalized_seconds": 1.58, "peak_rss_mb": 267},
  "csv_300k_seekable": {"normalized_seconds": 1.68, "peak_rss_mb": 352},
  "multi_sheet": {"to_dataframes_openpyxl_loads": 6, "read_all_sheets_openpyxl_loads": 9}
}
```

- [x] **Step 6: Commit the compatibility checkpoint**

```bash
git add pyproject.toml benchmarks/v010-reference.json scripts/capture_v010_contract.py tests/compatibility
git commit -m "test: freeze v0.10 parser compatibility"
```

---

### Task 2: Classify Legacy APIs and Emit One Warning

**Files:**
- Create: `src/messy_xlsx/warnings.py`
- Create: `tests/compatibility/test_legacy_warnings.py`
- Modify: `src/messy_xlsx/__init__.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `src/messy_xlsx/sheet.py`
- Modify: `src/messy_xlsx/multi_sheet.py`

**Interfaces:**
- Produces: `LegacyAPIWarning(DeprecationWarning)` and `warn_legacy(api_name: str) -> None`.
- Preserves: extension SPI calls remain warning-free; direct listed legacy entry points emit exactly one warning with the user's filename.

- [x] **Step 1: Write failing warning-boundary tests**

```python
# tests/compatibility/test_legacy_warnings.py
import warnings

from messy_xlsx import LegacyAPIWarning, MultiSheetParser, read_all_sheets, read_excel


def _legacy_records(callable_object):
    with warnings.catch_warnings(record=True) as records:
        warnings.simplefilter("always", LegacyAPIWarning)
        callable_object()
    return [record for record in records if record.category is LegacyAPIWarning]


def test_read_excel_emits_one_caller_facing_warning(sample_xlsx) -> None:
    records = _legacy_records(lambda: read_excel(str(sample_xlsx)))
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_all_sheets_suppresses_nested_parse_all_warning(sample_xlsx) -> None:
    records = _legacy_records(lambda: read_all_sheets(sample_xlsx))
    assert len(records) == 1


def test_direct_multi_sheet_methods_warn_once(sample_xlsx) -> None:
    parser = MultiSheetParser(sample_xlsx)
    assert len(_legacy_records(parser.parse_all)) == 1
    assert len(_legacy_records(lambda: parser.parse_sheet("Data"))) == 1
```

Run: `.venv/bin/pytest tests/compatibility/test_legacy_warnings.py -q`

Expected: collection fails because `LegacyAPIWarning` is not defined.

- [x] **Step 2: Add the public warning type and helper**

```python
# src/messy_xlsx/warnings.py
from __future__ import annotations

import warnings


class LegacyAPIWarning(DeprecationWarning):
    """A materialized compatibility API has a preferred bounded alternative."""


def warn_legacy(api_name: str) -> None:
    warnings.warn(
        f"{api_name} is a legacy materialized API retained through messy-xlsx v1.x",
        LegacyAPIWarning,
        stacklevel=3,
    )
```

Export `LegacyAPIWarning` from `messy_xlsx.__all__`.

- [x] **Step 3: Put warnings only at public boundaries**

Refactor each legacy method into a warning boundary plus a private implementation. Use this exact pattern:

```python
def to_dataframe(self, sheet=None, config=None):
    warn_legacy("MessyWorkbook.to_dataframe")
    return self._to_dataframe_compat(sheet, config)


def _to_dataframe_compat(self, sheet=None, config=None):
    sheet_name = sheet or self._sheet_names[0]
    return self._parse_sheet(sheet_name, config)
```

Make `read_excel()` call `_to_dataframe_compat()`, `read_all_sheets()` call `MultiSheetParser._parse_all_compat()`, and `read_excel_tables()` call a private table conversion method so nested adapters cannot emit duplicates. Apply the same boundary to every API listed in the design spec.

- [x] **Step 4: Run focused warnings and compatibility tests**

Run: `.venv/bin/pytest tests/compatibility/test_legacy_warnings.py tests/compatibility/test_v010_contract.py -q`

Expected: warning tests and all golden contracts pass.

- [x] **Step 5: Commit the warning contract**

```bash
git add src/messy_xlsx/warnings.py src/messy_xlsx/__init__.py src/messy_xlsx/workbook.py src/messy_xlsx/sheet.py src/messy_xlsx/multi_sheet.py tests/compatibility/test_legacy_warnings.py
git commit -m "feat: mark materialized APIs as legacy"
```

---

### Task 3: Replace Whole-Source Caches with a Spillable Replay Spool

**Files:**
- Create: `src/messy_xlsx/_spool.py`
- Create: `tests/test_spool.py`
- Modify: `src/messy_xlsx/_source.py`
- Modify: `tests/test_source_handle.py`
- Modify: `tests/test_resource_lifecycle.py`

**Interfaces:**
- Produces: `ReplaySpool.from_stream(stream, memory_limit=8_388_608) -> ReplaySpool`.
- Produces: `ReplaySpool.open_binary()`, `ReplaySpool.open_path_or_bytes()`, `ReplaySpool.close()`.
- `SourceHandle.open_path_or_bytes()` yields `Path | bytes` for path/bytes-only backends without an unbounded byte cache.

- [x] **Step 1: Write failing spill, cursor, and cleanup tests**

```python
# tests/test_spool.py
import io
from pathlib import Path

import pytest

from messy_xlsx._spool import ReplaySpool


def test_small_spool_stays_in_memory() -> None:
    source = io.BytesIO(b"abcdef")
    source.seek(3)
    spool = ReplaySpool.from_stream(source, memory_limit=16)
    assert source.tell() == 3
    with spool.open_path_or_bytes() as backend:
        assert backend == b"abcdef"
    spool.close()


def test_large_spool_uses_private_path_and_deletes_it() -> None:
    spool = ReplaySpool.from_stream(io.BytesIO(b"x" * 32), memory_limit=8)
    with spool.open_path_or_bytes() as backend:
        assert isinstance(backend, Path)
        path = backend
        assert path.exists()
        assert path.stat().st_mode & 0o777 == 0o600
    spool.close()
    assert not path.exists()


def test_close_is_idempotent() -> None:
    spool = ReplaySpool.from_stream(io.BytesIO(b"data"), memory_limit=8)
    spool.close()
    spool.close()


def test_spool_restores_cursor_when_read_fails() -> None:
    class Broken(io.BytesIO):
        def read(self, size: int = -1) -> bytes:
            raise OSError("injected read failure")

    source = Broken(b"data")
    source.seek(2)
    with pytest.raises(OSError, match="injected read failure"):
        ReplaySpool.from_stream(source, memory_limit=8)
    assert source.tell() == 2
```

Run: `.venv/bin/pytest tests/test_spool.py -q`

Expected: collection fails because `messy_xlsx._spool` does not exist.

- [x] **Step 2: Implement the spill state machine**

```python
# src/messy_xlsx/_spool.py
from __future__ import annotations

import io
import os
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import BinaryIO, Iterator


DEFAULT_MEMORY_LIMIT = 8 * 1024 * 1024
COPY_CHUNK_SIZE = 1024 * 1024


class ReplaySpool:
    def __init__(self, memory: bytes | None, path: Path | None) -> None:
        self._memory = memory
        self._path = path
        self._closed = False

    @classmethod
    def from_stream(
        cls,
        stream: BinaryIO,
        memory_limit: int = DEFAULT_MEMORY_LIMIT,
    ) -> "ReplaySpool":
        seekable = False
        entry = 0
        try:
            seekable = bool(stream.seekable())
            if seekable:
                entry = stream.tell()
        except (AttributeError, OSError, ValueError):
            seekable = False
        if not seekable:
            try:
                position = stream.tell()
            except (AttributeError, OSError, ValueError):
                position = 0
            if position != 0:
                raise ValueError("A non-seekable source must be positioned at byte 0")
        buffer = bytearray()
        path: Path | None = None
        opened = None
        primary_error: BaseException | None = None
        try:
            if seekable:
                stream.seek(0)
            while True:
                chunk = stream.read(COPY_CHUNK_SIZE)
                if not chunk:
                    break
                if path is None and len(buffer) + len(chunk) <= memory_limit:
                    buffer.extend(chunk)
                    continue
                if path is None:
                    descriptor, raw_path = tempfile.mkstemp(prefix="messy-xlsx-", suffix=".spool")
                    path = Path(raw_path)
                    try:
                        os.chmod(path, 0o600)
                        opened = os.fdopen(descriptor, "wb")
                    except BaseException:
                        try:
                            os.close(descriptor)
                        except OSError:
                            pass
                        path.unlink(missing_ok=True)
                        path = None
                        raise
                    opened.write(buffer)
                    buffer.clear()
                opened.write(chunk)
            if opened is not None:
                opened.close()
                opened = None
            return cls(bytes(buffer) if path is None else None, path)
        except BaseException as error:
            primary_error = error
            if opened is not None:
                try:
                    opened.close()
                except BaseException as cleanup_error:
                    error.add_note(f"temporary file close also failed: {cleanup_error!r}")
            if path is not None:
                try:
                    path.unlink(missing_ok=True)
                except BaseException as cleanup_error:
                    error.add_note(f"temporary file removal also failed: {cleanup_error!r}")
            raise
        finally:
            if seekable:
                try:
                    stream.seek(entry)
                except BaseException as restore_error:
                    if primary_error is None:
                        raise
                    primary_error.add_note(f"cursor restoration also failed: {restore_error!r}")

    @contextmanager
    def open_binary(self) -> Iterator[BinaryIO]:
        self._ensure_open()
        if self._path is None:
            with io.BytesIO(self._memory or b"") as stream:
                yield stream
            return
        with self._path.open("rb") as stream:
            yield stream

    @contextmanager
    def open_path_or_bytes(self) -> Iterator[Path | bytes]:
        self._ensure_open()
        yield self._path if self._path is not None else (self._memory or b"")

    def close(self) -> None:
        if self._closed:
            return
        self._memory = None
        if self._path is not None:
            self._path.unlink(missing_ok=True)
            self._path = None
        self._closed = True

    def _ensure_open(self) -> None:
        if self._closed:
            raise ValueError("ReplaySpool is closed")
```

- [x] **Step 3: Integrate the spool into `SourceHandle`**

Replace `_snapshot`, `_byte_cache`, and `_has_byte_cache` with one optional `ReplaySpool`. Keep `open_binary()` borrowing seekable caller streams directly. Add:

```python
@contextmanager
def open_path_or_bytes(self) -> Iterator[Path | bytes]:
    self._ensure_open()
    if self._path is not None:
        yield self._path
        return
    spool = self._ensure_spool()
    with spool.open_path_or_bytes() as source:
        yield source
```

Track an active-borrow flag around caller stream contexts and raise `RuntimeError("SourceHandle already has an active borrow")` on nesting. Translate temporary-file creation/write/capacity `OSError` into `FileError(operation="spool")` at the public source boundary.

- [x] **Step 4: Run source and lifecycle regression tests**

Run: `.venv/bin/pytest tests/test_spool.py tests/test_source_handle.py tests/test_resource_lifecycle.py -q`

Expected: all tests pass; seekable streams are restored and large replay sources leave no temporary files.

- [x] **Step 5: Commit the source lifecycle**

```bash
git add src/messy_xlsx/_spool.py src/messy_xlsx/_source.py tests/test_spool.py tests/test_source_handle.py tests/test_resource_lifecycle.py
git commit -m "perf: add spillable source replay"
```

---

### Task 4: Harden OOXML Archives and Build the Eager Workbook Manifest

**Files:**
- Create: `src/messy_xlsx/ooxml/__init__.py`
- Create: `src/messy_xlsx/ooxml/models.py`
- Create: `src/messy_xlsx/ooxml/security.py`
- Create: `src/messy_xlsx/ooxml/manifest.py`
- Create: `tests/test_ooxml_security.py`
- Create: `tests/test_ooxml_manifest.py`
- Modify: `pyproject.toml`

**Interfaces:**
- Produces: frozen `OoxmlLimits`, `SheetDescriptor`, and `WorkbookManifest`.
- Produces: `build_manifest(source: SourceHandle, limits: OoxmlLimits = DEFAULT_LIMITS) -> WorkbookManifest`.
- Raises: existing `FormatError` for unsafe or malformed packages without reading worksheet values.

- [x] **Step 1: Write failing security and workbook-order tests**

```python
# tests/test_ooxml_security.py
import io
import zipfile

import pytest

from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import OoxmlLimits
from messy_xlsx.ooxml.security import validate_archive


def _archive(entries: list[tuple[str, bytes]]) -> zipfile.ZipFile:
    raw = io.BytesIO()
    with zipfile.ZipFile(raw, "w", zipfile.ZIP_DEFLATED) as package:
        for name, value in entries:
            package.writestr(name, value)
    raw.seek(0)
    return zipfile.ZipFile(raw)


def test_duplicate_members_are_rejected() -> None:
    package = _archive([("xl/workbook.xml", b"a"), ("xl/workbook.xml", b"b")])
    with pytest.raises(FormatError, match="duplicate"):
        validate_archive(package, OoxmlLimits())


@pytest.mark.parametrize("name", ["../escape.xml", "/absolute.xml", "xl/../../escape.xml"])
def test_unsafe_member_paths_are_rejected(name: str) -> None:
    package = _archive([(name, b"content")])
    with pytest.raises(FormatError, match="unsafe archive path"):
        validate_archive(package, OoxmlLimits())
```

```python
# tests/test_ooxml_manifest.py
from messy_xlsx._source import SourceHandle
from messy_xlsx.ooxml.manifest import build_manifest


def test_manifest_preserves_sheet_order_without_cell_values(sample_xlsx) -> None:
    with SourceHandle(sample_xlsx) as source:
        manifest = build_manifest(source)
    assert [sheet.name for sheet in manifest.sheets] == ["Data"]
    assert not hasattr(manifest, "dataframe")
    assert not hasattr(manifest, "shared_strings")
```

Run: `.venv/bin/pytest tests/test_ooxml_security.py tests/test_ooxml_manifest.py -q`

Expected: collection fails because the `messy_xlsx.ooxml` package does not exist.

- [x] **Step 2: Define immutable limits and manifest models**

Add `"defusedxml>=0.7.1"` to the required project dependencies before importing
the hardened parser.

```python
# src/messy_xlsx/ooxml/models.py
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class OoxmlLimits:
    max_members: int = 10_000
    max_total_uncompressed: int = 2 * 1024**3
    max_xml_uncompressed: int = 512 * 1024**2
    suspicious_ratio_size: int = 64 * 1024**2
    max_compression_ratio: float = 1_000.0
    max_formula_samples: int = 256
    max_xml_depth: int = 256
    max_element_attributes: int = 256
    max_element_text: int = 16 * 1024 * 1024


@dataclass(frozen=True)
class SheetDescriptor:
    name: str
    relationship_id: str
    target: str
    state: str


@dataclass(frozen=True)
class StyleManifest:
    custom_number_formats: tuple[tuple[int, str], ...]
    date_style_ids: tuple[int, ...]


@dataclass(frozen=True)
class WorkbookManifest:
    workbook_type: str
    date_system: str
    sheets: tuple[SheetDescriptor, ...]
    has_shared_strings: bool
    shared_strings_uncompressed_size: int
    styles: StyleManifest
    external_relationships: tuple[str, ...] = field(default_factory=tuple)
```

- [x] **Step 3: Implement archive validation before XML parsing**

```python
# src/messy_xlsx/ooxml/security.py
from pathlib import PurePosixPath
from zipfile import ZipFile

from defusedxml import ElementTree as SafeElementTree

from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import OoxmlLimits


def _safe_member(name: str) -> bool:
    path = PurePosixPath(name)
    drive_like = bool(path.parts and ":" in path.parts[0])
    return bool(
        name
        and "\x00" not in name
        and "\\" not in name
        and not drive_like
        and not path.is_absolute()
        and ".." not in path.parts
    )


def validate_archive(package: ZipFile, limits: OoxmlLimits) -> None:
    members = package.infolist()
    if len(members) > limits.max_members:
        raise FormatError("OOXML archive exceeds member limit", member_count=len(members))
    names = [member.filename for member in members]
    if len(names) != len(set(names)):
        raise FormatError("OOXML archive contains duplicate member names")
    if any(not _safe_member(name) for name in names):
        raise FormatError("OOXML archive contains unsafe archive path")
    total = sum(member.file_size for member in members)
    if total > limits.max_total_uncompressed:
        raise FormatError("OOXML archive exceeds total uncompressed limit", uncompressed=total)
    for member in members:
        if member.filename.endswith(".xml") and member.file_size > limits.max_xml_uncompressed:
            raise FormatError("OOXML XML member exceeds size limit", member=member.filename)
        ratio = member.file_size / max(member.compress_size, 1)
        if member.file_size > limits.suspicious_ratio_size and ratio > limits.max_compression_ratio:
            raise FormatError("OOXML member has suspicious compression ratio", member=member.filename)


def reject_unsafe_xml_prefix(prefix: bytes, member: str) -> None:
    lowered = prefix.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise FormatError("OOXML XML declarations are not allowed", member=member)


def safe_iterparse(source, member: str, limits: OoxmlLimits):
    depth = 0
    for event, element in SafeElementTree.iterparse(
        source,
        events=("start", "end"),
        forbid_dtd=True,
        forbid_entities=True,
        forbid_external=True,
    ):
        if event == "start":
            depth += 1
            if depth > limits.max_xml_depth:
                raise FormatError("OOXML XML exceeds depth limit", member=member)
            if len(element.attrib) > limits.max_element_attributes:
                raise FormatError("OOXML element exceeds attribute limit", member=member)
        else:
            if len(element.text or "") > limits.max_element_text:
                raise FormatError("OOXML element exceeds text limit", member=member)
            depth -= 1
        yield event, element
```

- [x] **Step 4: Parse workbook relationships without following external targets**

Implement `build_manifest()` with `zipfile.ZipFile`, `validate_archive()`, and `safe_iterparse()`. Resolve only normalized internal relationship targets below the archive root. Record `TargetMode="External"` relationships as evidence and never open them. Open the archive through `SourceHandle.open_binary()` and close it before returning the frozen manifest.

The parser must return only `SheetDescriptor` metadata, date system, shared-string presence/declared size, and external relationship names. It must not read `sharedStrings.xml` values or worksheet cell data.

```python
def build_manifest(source, limits=OoxmlLimits()):
    with source.open_binary() as binary, ZipFile(binary) as package:
        validate_archive(package, limits)
        relationships, external_names = _read_relationships(
            package, "xl/_rels/workbook.xml.rels"
        )
        workbook = _read_workbook_xml(package, "xl/workbook.xml")
        sheets = tuple(
            SheetDescriptor(
                name=sheet.name,
                relationship_id=sheet.relationship_id,
                target=_internal_target(relationships[sheet.relationship_id]),
                state=sheet.state,
            )
            for sheet in workbook.sheets
        )
        shared = package.getinfo("xl/sharedStrings.xml") if "xl/sharedStrings.xml" in package.namelist() else None
        styles = _read_styles(package, "xl/styles.xml")
        return WorkbookManifest(
            workbook_type=workbook.workbook_type,
            date_system=workbook.date_system,
            sheets=sheets,
            has_shared_strings=shared is not None,
            shared_strings_uncompressed_size=0 if shared is None else shared.file_size,
            styles=styles,
            external_relationships=tuple(external_names),
        )
```

Implement `_read_styles()` as a bounded `iterparse()` over `xl/styles.xml`:

```python
def _read_styles(package, member):
    from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format

    if member not in package.namelist():
        return StyleManifest((), ())
    custom = {}
    date_styles = []
    in_cell_xfs = False
    style_index = 0
    with package.open(member) as source:
        for event, element in safe_iterparse(source, member, OoxmlLimits()):
            local = element.tag.rsplit("}", 1)[-1]
            if event == "start" and local == "cellXfs":
                in_cell_xfs = True
            elif event == "end" and local == "numFmt":
                custom[int(element.attrib["numFmtId"])] = element.attrib["formatCode"]
            elif event == "end" and local == "xf" and in_cell_xfs:
                number_format_id = int(element.attrib.get("numFmtId", "0"))
                code = custom.get(number_format_id, BUILTIN_FORMATS.get(number_format_id, ""))
                if code and is_date_format(code):
                    date_styles.append(style_index)
                style_index += 1
            elif event == "end" and local == "cellXfs":
                in_cell_xfs = False
            if event == "end":
                element.clear()
    return StyleManifest(tuple(sorted(custom.items())), tuple(date_styles))
```

- [x] **Step 5: Run security, manifest, malformed-file, and format tests**

Run: `.venv/bin/pytest tests/test_ooxml_security.py tests/test_ooxml_manifest.py tests/test_edge_cases/test_malformed_files.py tests/test_detection/test_format_detector.py -q`

Expected: all tests pass with external relationships never opened.

- [x] **Step 6: Commit the hardened eager manifest**

```bash
git add pyproject.toml src/messy_xlsx/ooxml tests/test_ooxml_security.py tests/test_ooxml_manifest.py
git commit -m "feat: add hardened OOXML manifest"
```

---

### Task 5: Add Lazy Sheet Metadata, Bounded Structure Sampling, and Robust Cache Identity

**Files:**
- Modify: `src/messy_xlsx/ooxml/models.py`
- Modify: `src/messy_xlsx/ooxml/manifest.py`
- Create: `src/messy_xlsx/detection/structure_sampler.py`
- Create: `src/messy_xlsx/parsing/fastexcel_session.py`
- Modify: `src/messy_xlsx/detection/structure_analyzer.py`
- Modify: `src/messy_xlsx/detection/locale_detector.py`
- Create: `tests/test_ooxml_sheet_metadata.py`
- Create: `tests/test_detection/test_structure_sampler.py`
- Modify: `src/messy_xlsx/cache.py`
- Modify: `tests/test_detection/test_cache.py`

**Interfaces:**
- Produces: `Interval`, `IntervalIndex`, `MergeRange`, `SheetManifest`, and `ManifestReader.sheet(name) -> SheetManifest`.
- Produces: `StructureSampler.analyze(sheet, header_patterns=None) -> StructureInfo`, cached by sheet and pattern tuple, with no complete DataFrame.
- Produces: closable `FastexcelSession` that opens one workbook and serves bounded samples plus one materialization per selected sheet.
- Produces: `PathIdentity.before(path)`, `PathIdentity.unchanged(path)`, and cache insertion only after a stable pre/post stat comparison.

- [x] **Step 1: Write failing interval, lazy-load, and cache-identity tests**

```python
# tests/test_ooxml_sheet_metadata.py
from messy_xlsx._source import SourceHandle
from messy_xlsx.ooxml.manifest import ManifestReader
from messy_xlsx.ooxml.models import Interval, IntervalIndex


def test_interval_index_uses_ranges_without_expanding_cells() -> None:
    index = IntervalIndex((Interval(2, 4), Interval(10, 12)))
    assert index.contains(3)
    assert not index.contains(7)
    assert len(index.intervals) == 2


def test_sheet_xml_is_loaded_lazily(sample_xlsx, monkeypatch) -> None:
    opened: list[str] = []
    with SourceHandle(sample_xlsx) as source:
        reader = ManifestReader(source, on_member_open=opened.append)
        assert not any(name.startswith("xl/worksheets/") for name in opened)
        first = reader.sheet(reader.workbook.sheets[0].name)
        assert first.name == reader.workbook.sheets[0].name
        assert sum(name.startswith("xl/worksheets/") for name in opened) == 1
```

Add a cache test that rewrites a path to equal size and restored `mtime_ns`, then asserts a changed `ctime_ns` or inode prevents a stale hit.

Run: `.venv/bin/pytest tests/test_ooxml_sheet_metadata.py tests/test_detection/test_cache.py -q`

Expected: failures because interval and lazy reader types do not exist.

- [x] **Step 2: Add immutable interval and sheet models**

```python
@dataclass(frozen=True, order=True)
class Interval:
    start: int
    end: int

    def __post_init__(self) -> None:
        if self.start < 1 or self.end < self.start:
            raise ValueError("invalid one-based interval")


@dataclass(frozen=True)
class IntervalIndex:
    intervals: tuple[Interval, ...]
    starts: tuple[int, ...] = field(init=False)

    def __post_init__(self) -> None:
        object.__setattr__(self, "starts", tuple(interval.start for interval in self.intervals))

    def contains(self, value: int) -> bool:
        position = bisect_right(self.starts, value) - 1
        return position >= 0 and value <= self.intervals[position].end


@dataclass(frozen=True)
class MergeRange:
    min_row: int
    min_col: int
    max_row: int
    max_col: int


@dataclass(frozen=True)
class SheetManifest:
    name: str
    target: str
    declared_dimension: tuple[int, int, int, int] | None
    observed_max_row: int
    observed_max_col: int
    hidden_rows: IntervalIndex
    hidden_columns: IntervalIndex
    merged_ranges: tuple[MergeRange, ...]
    has_formulas: bool
    formula_samples: tuple[str, ...]
```

Normalize and merge overlapping adjacent intervals before constructing `IntervalIndex` so lookup remains logarithmic and memory scales with ranges, not covered cells.

- [x] **Step 3: Implement one lazy SAX-style worksheet metadata pass**

Use `ElementTree.iterparse()` over the selected worksheet member. On `end` events, record row/column hidden intervals, merge references, formula presence and at most 256 formula coordinates, declared dimension, and maximum observed cell coordinate. Clear each element immediately after processing it.

```python
class ManifestReader:
    def __init__(self, source, limits=OoxmlLimits(), on_member_open=None):
        self._source = source
        self._limits = limits
        self._on_member_open = on_member_open or (lambda _name: None)
        self.workbook = build_manifest(source, limits)
        self._sheets: dict[str, SheetManifest] = {}

    def sheet(self, name: str) -> SheetManifest:
        if name not in self._sheets:
            descriptor = next(sheet for sheet in self.workbook.sheets if sheet.name == name)
            self._on_member_open(descriptor.target)
            self._sheets[name] = self._parse_sheet(descriptor)
        return self._sheets[name]
```

- [x] **Step 4: Strengthen path cache identity**

```python
@dataclass(frozen=True)
class PathIdentity:
    resolved: str
    device: int
    inode: int
    size: int
    mtime_ns: int
    ctime_ns: int

    @classmethod
    def before(cls, path: Path) -> "PathIdentity":
        stat = path.stat()
        return cls(
            str(path.resolve()), stat.st_dev, stat.st_ino, stat.st_size,
            stat.st_mtime_ns, stat.st_ctime_ns,
        )

    def unchanged(self, path: Path) -> bool:
        return self == type(self).before(path)
```

Make `StructureCache.put()` accept the identity captured before analysis and skip insertion when `identity.unchanged(path)` is false.

- [x] **Step 5: Characterize and replace full-frame structure analysis**

```python
# tests/test_detection/test_structure_sampler.py
import pandas as pd
import pytest

from messy_xlsx._source import SourceHandle
from messy_xlsx.detection.structure_sampler import (
    SampleWindow,
    StructureEvidence,
    StructureSampler,
)
from messy_xlsx.ooxml.manifest import ManifestReader


class RecordingExcelReader:
    def __init__(self, manifest) -> None:
        self.manifest = manifest
        self.open_count = 1
        self.requests = []
        self.complete_dataframe_count = 0

    def sample_windows(self, sheet, windows, max_column):
        self.requests.append(
            {"sheet": sheet, "windows": windows, "max_column": max_column}
        )
        row_numbers = tuple(
            row
            for window in windows
            for row in range(window.start_row, window.start_row + window.n_rows)
        )
        return StructureEvidence(
            row_numbers=row_numbers,
            values=pd.DataFrame([["name", "value"], ["a", 1]], index=row_numbers[:2]),
        )


@pytest.fixture
def recording_excel_reader(sample_xlsx):
    source = SourceHandle(sample_xlsx)
    manifest = ManifestReader(source)
    reader = RecordingExcelReader(manifest)
    try:
        yield reader
    finally:
        source.close()


def test_sampler_uses_bounded_rows_and_reuses_one_reader(sample_xlsx, recording_excel_reader) -> None:
    sampler = StructureSampler(recording_excel_reader, manifest_reader=recording_excel_reader.manifest)
    first = sampler.analyze("Data")
    second = sampler.analyze("Data")
    assert first == second
    assert recording_excel_reader.open_count == 1
    request = recording_excel_reader.requests[0]
    assert request["sheet"] == "Data"
    assert request["windows"][0].start_row == 1
    assert request["windows"][0].n_rows <= 10_000
    assert sum(window.n_rows for window in request["windows"]) <= 10_500
    assert recording_excel_reader.complete_dataframe_count == 0
```

Before replacing the analyzer, compare `StructureInfo` from the current implementation and the sampler across maintained and generated workbooks. The sampler reads at most 10,000 data-region rows, scores the first 16 header candidates, takes exact dimensions from manifest/fastexcel metadata, and caches by `(sheet_name, tuple(header_patterns or ()))`.

Extract the current analyzer's scoring bodies into internal pure helpers instead
of rewriting the heuristics. `StructureAnalyzer` calls those helpers through its
existing worksheet adapter; `StructureSampler` calls the same helpers with
`StructureEvidence`. Build the evidence row set as the union of rows `1..10_000`,
the current blank-row sampling positions from `_detect_blank_rows_sampled()`,
and the final ten worksheet rows used by `_suggest_skip_footer()`. The union is
sorted, deduplicated, and coalesced into contiguous `SampleWindow` values, and
remains at most 10,500 retained rows. Each window uses fastexcel's integer
`skip_rows` plus `n_rows`; do not use list/callable `skip_rows`, because a Task 5
characterization test must preserve the observed 0.20.2 behavior that those
forms do not filter reliably. Repeated windows are allowed only for bounded
structure evidence, never for a public streaming iterator.

```python
@dataclass(frozen=True)
class StructureEvidence:
    row_numbers: tuple[int, ...]
    values: pd.DataFrame

    def row(self, row_number: int) -> tuple[object, ...]:
        if row_number not in self.values.index:
            return ()
        return tuple(self.values.loc[row_number].tolist())


@dataclass(frozen=True)
class SampleWindow:
    start_row: int
    n_rows: int


def structure_sample_windows(max_row: int) -> tuple[SampleWindow, ...]:
    head = range(1, min(max_row, 10_000) + 1)
    tail = range(max(1, max_row - 9), max_row + 1)
    sampled_blanks = blank_row_sample_positions(1, max_row)
    return coalesce_rows(tuple(sorted(set(head) | set(tail) | set(sampled_blanks))))
```

The shared pure composition function must preserve this exact source mapping:

- data bounds: existing `_detect_data_region` logic over head evidence, with
  `SheetManifest.observed_max_row/observed_max_col` as the current max hints;
- merge, hidden, and formula fields: exact `SheetManifest` values;
- header and metadata: existing first-16-row scoring and pre-header rules;
- blank rows and table splits: existing sampled positions and grouping rules;
- locale: `LocaleDetector.detect_from_evidence(text_values, format_codes)` using
  the current text scoring and manifest number-format codes;
- sparse columns: the current first-1,000-row threshold;
- footer: the current patterns over the final ten evidence rows;
- suggestions and merge flags: the current helper rules unchanged.

```python
def analyze_structure_evidence(
    evidence: StructureEvidence,
    manifest: SheetManifest,
    header_patterns: tuple[str, ...],
) -> StructureInfo:
    data_region = detect_data_region(evidence, manifest)
    merged = [
        (item.min_row, item.min_col, item.max_row, item.max_col)
        for item in manifest.merged_ranges
    ]
    header = detect_headers(evidence, data_region, merged, header_patterns)
    blank_rows = detect_blank_rows(evidence, data_region)
    tables = detect_multiple_tables(data_region, header, blank_rows)
    locale = detect_locale_evidence(evidence, manifest)
    return compose_structure_info(
        evidence=evidence,
        manifest=manifest,
        data_region=data_region,
        merged=merged,
        header=header,
        blank_rows=blank_rows,
        tables=tables,
        locale=locale,
    )
```

```python
class StructureSampler:
    def __init__(self, excel_reader, manifest_reader, metrics=None) -> None:
        self._excel_reader = excel_reader
        self._manifest_reader = manifest_reader
        self._metrics = metrics
        self._cache = {}

    def analyze(self, sheet, header_patterns=None):
        key = (sheet, tuple(header_patterns or ()))
        if key not in self._cache:
            manifest = self._manifest_reader.sheet(sheet)
            windows = structure_sample_windows(manifest.observed_max_row)
            sample = self._excel_reader.sample_windows(
                sheet, windows=windows, max_column=manifest.observed_max_col
            )
            self._cache[key] = analyze_structure_evidence(sample, manifest, key[1])
            if self._metrics is not None:
                self._metrics.sample_reads += 1
        return self._cache[key]
```

```python
# src/messy_xlsx/parsing/fastexcel_session.py
class FastexcelSession:
    def __init__(self, source) -> None:
        self._context = source.open_path_or_bytes()
        backend = self._context.__enter__()
        try:
            self._reader = fastexcel.read_excel(backend)
        except BaseException:
            self._context.__exit__(None, None, None)
            raise
        self._closed = False

    @property
    def sheet_names(self):
        return tuple(self._reader.sheet_names)

    def sample_windows(self, sheet, windows, max_column):
        frames = []
        rows = []
        for window in windows:
            batch = self._reader.load_sheet(
                sheet,
                header_row=None,
                skip_rows=window.start_row - 1,
                n_rows=window.n_rows,
                schema_sample_rows=min(1_000, window.n_rows),
                use_columns=list(range(max_column)),
                eager=True,
            )
            frame = batch.to_pandas()
            frame.index = pd.RangeIndex(
                window.start_row,
                window.start_row + len(frame),
                name="worksheet_row",
            )
            rows.extend(frame.index.tolist())
            frames.append(frame)
        values = pd.concat(frames, axis=0) if frames else pd.DataFrame()
        return StructureEvidence(tuple(rows), values)

    def materialize(self, sheet):
        return self._reader.load_sheet(
            sheet,
            header_row=None,
            schema_sample_rows=1_000,
            dtype_coercion="coerce",
            eager=True,
        )

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._reader = None
        self._context.__exit__(None, None, None)
```

`MessyWorkbook` owns at most one lazy session and closes it before its source
handle. The bounded sample is allowed to become a bounded pandas frame for the
existing detector; only complete discarded DataFrames are forbidden.

- [x] **Step 6: Run focused metadata, sampler, and cache tests**

Run: `.venv/bin/pytest tests/test_ooxml_sheet_metadata.py tests/test_detection/test_structure_sampler.py tests/test_detection/test_cache.py tests/test_detection/test_structure_analyzer.py -q`

Expected: all tests pass and worksheet metadata is parsed only on first request per sheet.

- [x] **Step 7: Commit lazy metadata, sampler, and cache identity**

```bash
git add src/messy_xlsx/ooxml src/messy_xlsx/detection/structure_analyzer.py src/messy_xlsx/detection/structure_sampler.py src/messy_xlsx/detection/locale_detector.py src/messy_xlsx/parsing/fastexcel_session.py src/messy_xlsx/cache.py tests/test_ooxml_sheet_metadata.py tests/test_detection/test_structure_sampler.py tests/test_detection/test_cache.py
git commit -m "perf: index OOXML sheet metadata lazily"
```

---

### Task 6: Compile Immutable Plans and Route by Capability and Output Mode

**Files:**
- Create: `src/messy_xlsx/parsing/contracts.py`
- Create: `src/messy_xlsx/parsing/router.py`
- Create: `src/messy_xlsx/parsing/fallback.py`
- Create: `tests/test_reader_routing.py`
- Modify: `src/messy_xlsx/parsing/parse_plan.py`
- Modify: `src/messy_xlsx/parsing/handler_registry.py`
- Modify: `tests/test_parse_plan.py`

**Interfaces:**
- Produces: `OutputMode`, `BackendKind`, `ReaderDecision`, `ParseMetrics`, `MaterializedArrowReader`, and `StreamingBatchReader`; materialized reader factories bind the immutable `ParsePlan`, so the operation protocol is `read_table() -> pa.Table`.
- Produces: `compile_parse_plan(config, structure, format_type, output_mode: OutputMode, batch_size: int | None) -> ParsePlan` with a deep immutable configuration snapshot.
- Produces: `BackendRouter.select(workbook_context) -> ReaderDecision`.
- Produces: `FallbackCoordinator.materialize()` and `FallbackCoordinator.batches()` with cleanup-before-fallback and no retry after observable output.

- [x] **Step 1: Write the failing routing decision matrix**

```python
# tests/test_reader_routing.py
import pytest

from messy_xlsx.parsing.contracts import BackendKind, OutputMode
from messy_xlsx.parsing.router import BackendRouter, WorkbookContext


@pytest.mark.parametrize(
    ("mode", "evaluate_formulas", "custom_registry", "expected"),
    [
        (OutputMode.MATERIALIZED, True, False, BackendKind.FASTEXCEL),
        (OutputMode.MATERIALIZED, False, False, BackendKind.OPENPYXL_COMPAT),
        (OutputMode.STREAMING, True, False, BackendKind.OPENPYXL_STREAMING),
        (OutputMode.STREAMING, False, False, BackendKind.OPENPYXL_STREAMING),
        (OutputMode.MATERIALIZED, True, True, BackendKind.CUSTOM_DATAFRAME),
        (OutputMode.STREAMING, True, True, BackendKind.CUSTOM_DATAFRAME),
    ],
)
def test_ooxml_routing_matrix(mode, evaluate_formulas, custom_registry, expected) -> None:
    context = WorkbookContext(
        format_type="xlsx",
        output_mode=mode,
        evaluate_formulas=evaluate_formulas,
        has_custom_registry=custom_registry,
    )
    assert BackendRouter().select(context).backend is expected
```

Add tests asserting `batch_size=0` fails before backend initialization and mutating nested `type_hints` or `drop_conditions` after stream creation cannot change the plan.

Run: `.venv/bin/pytest tests/test_reader_routing.py tests/test_parse_plan.py -q`

Expected: collection fails because reader contract modules do not exist.

- [x] **Step 2: Define exact internal reader contracts**

```python
# src/messy_xlsx/parsing/contracts.py
from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from typing import Protocol

import pyarrow as pa


class OutputMode(StrEnum):
    MATERIALIZED = "materialized"
    STREAMING = "streaming"


class BackendKind(StrEnum):
    FASTEXCEL = "fastexcel"
    OPENPYXL_COMPAT = "openpyxl_compat"
    OPENPYXL_STREAMING = "openpyxl_streaming"
    CSV_STREAMING = "csv_streaming"
    XLS_STREAMING = "xls_streaming"
    CUSTOM_DATAFRAME = "custom_dataframe"


@dataclass(frozen=True)
class ReaderDecision:
    backend: BackendKind
    reason: str


@dataclass
class ParseMetrics:
    manifest_builds: int = 0
    sample_reads: int = 0
    full_materializations: int = 0
    streaming_passes: int = 0
    failed_attempts: int = 0


class MaterializedArrowReader(Protocol):
    def read_table(self) -> pa.Table:
        raise NotImplementedError


class StreamingBatchReader(Protocol):
    @property
    def schema(self) -> pa.Schema:
        raise NotImplementedError

    def read_next_batch(self) -> pa.RecordBatch | None:
        raise NotImplementedError

    def close(self) -> None:
        raise NotImplementedError
```

- [x] **Step 3: Add output mode, batch size, and deep snapshots to `ParsePlan`**

Create a frozen `ConfigSnapshot` whose dict/list fields become sorted tuples and whose values are recursively copied. Validate `batch_size` when compiling, before any I/O:

```python
def _freeze(value):
    if isinstance(value, dict):
        return tuple((key, _freeze(item)) for key, item in sorted(value.items()))
    if isinstance(value, list):
        return tuple(_freeze(item) for item in value)
    if isinstance(value, tuple):
        return tuple(_freeze(item) for item in value)
    return value


def _validated_batch_size(mode: OutputMode, batch_size: int | None) -> int | None:
    if mode is OutputMode.STREAMING and (batch_size is None or batch_size < 1):
        raise ValueError("batch_size must be >= 1 for streaming output")
    return batch_size
```

- [x] **Step 4: Implement the output-mode router and custom-registry escape hatch**

```python
# src/messy_xlsx/parsing/router.py
@dataclass(frozen=True)
class WorkbookContext:
    format_type: str
    output_mode: OutputMode
    evaluate_formulas: bool
    has_custom_registry: bool


class BackendRouter:
    def select(self, context: WorkbookContext) -> ReaderDecision:
        if context.has_custom_registry:
            return ReaderDecision(BackendKind.CUSTOM_DATAFRAME, "caller extension compatibility")
        if context.format_type in {"csv", "tsv", "txt"}:
            return ReaderDecision(BackendKind.CSV_STREAMING, "text chunk reader")
        if context.format_type == "xls":
            return ReaderDecision(BackendKind.XLS_STREAMING, "optional xlrd row reader")
        if context.output_mode is OutputMode.STREAMING:
            return ReaderDecision(BackendKind.OPENPYXL_STREAMING, "fastexcel has no batch iterator")
        if context.evaluate_formulas:
            return ReaderDecision(BackendKind.FASTEXCEL, "fast cached-value materialization")
        return ReaderDecision(BackendKind.OPENPYXL_COMPAT, "formula expressions required")
```

Treat a registry as custom when it is not the exact built-in instance/type composition created by `MessyWorkbook`; do not infer compatibility merely from the file format.

- [x] **Step 5: Implement transactional classified fallback**

```python
# src/messy_xlsx/parsing/fallback.py
class FallbackCoordinator:
    def __init__(self, is_compatibility_error) -> None:
        self._is_compatibility_error = is_compatibility_error

    def materialize(self, primary_factory, fallback_factory):
        try:
            with primary_factory() as primary:
                return primary.read_table()
        except Exception as error:
            if not self._is_compatibility_error(error):
                raise
            primary_summary = f"{type(error).__name__}: {error}"
        try:
            with fallback_factory() as fallback:
                return fallback.read_table()
        except Exception as fallback_error:
            setattr(
                fallback_error,
                "backend_context",
                {"primary_failure": primary_summary, "fallback": type(fallback_error).__name__},
            )
            fallback_error.add_note(f"primary backend failed: {primary_summary}")
            raise

    def batches(self, primary_factory, fallback_factory):
        reader = None
        yielded = False
        using_fallback = False
        primary_summary = None
        try:
            while True:
                if reader is None:
                    try:
                        factory = fallback_factory if using_fallback else primary_factory
                        reader = factory()
                    except Exception as error:
                        if not using_fallback and not yielded and self._is_compatibility_error(error):
                            primary_summary = f"{type(error).__name__}: {error}"
                            using_fallback = True
                            continue
                        if primary_summary is not None:
                            error.add_note(f"primary backend failed: {primary_summary}")
                        raise
                try:
                    batch = reader.read_next_batch()
                except Exception as error:
                    try:
                        reader.close()
                    except BaseException as cleanup_error:
                        error.add_note(f"reader cleanup also failed: {cleanup_error!r}")
                    reader = None
                    if yielded or using_fallback or not self._is_compatibility_error(error):
                        raise
                    primary_summary = f"{type(error).__name__}: {error}"
                    using_fallback = True
                    continue
                if batch is None:
                    return
                yielded = True
                yield batch
        finally:
            if reader is not None:
                reader.close()
```

The factories own source restoration and spool cleanup. The primary reader is closed before `fallback_factory()` executes. `PermissionError`, `FileNotFoundError`, `MemoryError`, invalid configuration, and source-ownership failures must fail the classifier.

- [x] **Step 6: Run plan, routing, and fallback tests**

Run: `.venv/bin/pytest tests/test_reader_routing.py tests/test_parse_plan.py tests/test_architecture_contracts.py -q`

Expected: all tests pass; custom registry injections continue to drive parsing.

- [x] **Step 7: Commit the reader contracts and router**

```bash
git add src/messy_xlsx/parsing/contracts.py src/messy_xlsx/parsing/router.py src/messy_xlsx/parsing/fallback.py src/messy_xlsx/parsing/parse_plan.py src/messy_xlsx/parsing/handler_registry.py tests/test_reader_routing.py tests/test_parse_plan.py
git commit -m "refactor: route parsers by output capability"
```

---

### Task 7: Implement the Fastexcel Materialized Arrow Reader

**Files:**
- Create: `src/messy_xlsx/parsing/xlsx_materialized.py`
- Create: `src/messy_xlsx/parsing/legacy_adapter.py`
- Create: `tests/test_xlsx_materialized.py`
- Modify: `src/messy_xlsx/parsing/fastexcel_session.py`
- Modify: `src/messy_xlsx/parsing/xlsx_handler.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `tests/test_architecture_contracts.py`
- Modify: `tests/test_resource_lifecycle.py`

**Interfaces:**
- Produces: `FastexcelMaterializedReader(session, sheet, plan).read_table() -> pa.Table` using the workbook-scoped `FastexcelSession`; the factory binds the immutable `ParsePlan` when it constructs the operation reader.
- Produces: `LegacyDataFrameAdapter.to_dataframe(table, plan) -> pd.DataFrame`, the only built-in bridge for existing materialized APIs.
- Preserves: the legacy `XLSXHandler.parse() -> pd.DataFrame` SPI through an adapter.

- [x] **Step 1: Write a failing zero-openpyxl materialization test**

```python
# tests/test_xlsx_materialized.py
import openpyxl
import pytest

from messy_xlsx import SheetConfig
from messy_xlsx._source import SourceHandle
from messy_xlsx.parsing.contracts import OutputMode
from messy_xlsx.parsing.parse_plan import compile_parse_plan
from messy_xlsx.parsing.fastexcel_session import FastexcelSession
from messy_xlsx.parsing.xlsx_materialized import FastexcelMaterializedReader


@pytest.fixture
def basic_parse_plan():
    return compile_parse_plan(
        SheetConfig(auto_detect=False, normalize=False, sanitize_column_names=False),
        structure=None,
        format_type="xlsx",
        output_mode=OutputMode.MATERIALIZED,
        batch_size=None,
    )


def test_ordinary_materialization_never_loads_openpyxl(sample_xlsx, basic_parse_plan, monkeypatch) -> None:
    def forbidden(*args, **kwargs):
        raise AssertionError("openpyxl must not load on the ordinary materialized path")

    monkeypatch.setattr(openpyxl, "load_workbook", forbidden)
    with SourceHandle(sample_xlsx) as source:
        session = FastexcelSession(source)
        try:
            table = FastexcelMaterializedReader(
                session, "Data", basic_parse_plan
            ).read_table()
        finally:
            session.close()
    assert table.num_rows > 0
```

Add tests for seekable and non-seekable buffers larger than 8 MiB. Fastexcel
receives the same private spill `Path` during initialization and materialization;
reader/session closure does not delete the SourceHandle-owned path, and workbook
or locally owned handler closure deletes it on success and failure.

Run: `.venv/bin/pytest tests/test_xlsx_materialized.py -q`

Expected: collection fails because `xlsx_materialized` does not exist.

- [x] **Step 2: Implement one whole-sheet fastexcel read**

```python
# src/messy_xlsx/parsing/xlsx_materialized.py
from __future__ import annotations

import pyarrow as pa


class FastexcelMaterializedReader:
    def __init__(self, session, sheet: str, plan) -> None:
        self._session = session
        self._sheet = sheet
        self._plan = plan

    def read_table(self) -> pa.Table:
        value = self._session.materialize(self._sheet, skip_rows=0)
        return _coerce_materialized_table(value)
```

The reader accepts a `pa.Table`, wraps a `pa.RecordBatch`, or calls a backend
wrapper's `to_arrow()` exactly once. It performs no slicing, pandas conversion,
metric update, or source cleanup. `FallbackCoordinator` owns metrics. A
workbook-scoped `FastexcelSession` is reused across eligible sheet reads and is
closed before the workbook's SourceHandle.

For a no-range plan with nonzero `skip_rows`, retain backend-pushed fastexcel
skipping through the legacy handler path. Whole-sheet coercion can turn numeric
cells below text metadata into strings before pandas slicing can recover their
types; the mixed `("metadata",), (1,), (2,)` regression test freezes this gate.

- [x] **Step 3: Adapt the legacy XLSX handler without changing output**

```python
class LegacyDataFrameAdapter:
    def to_dataframe(self, table, plan):
        del plan
        return table.to_pandas()
```

Keep handler framing/cleaning in `XLSXHandler` and keep normalization,
positional sanitization, thawed renames, `normalize=False`, regex filtering, and
sequential condition filtering in `MessyWorkbook` in that exact order. The
private bound-plan seam is available only to the exact untouched built-in
registry. Custom registries, subclasses, handler mutations, formula-expression,
merge, hidden, range, and streaming cases retain compatibility routing.

- [x] **Step 4: Run XLSX, architecture, and golden contracts**

Run: `.venv/bin/pytest tests/test_xlsx_materialized.py tests/test_parsing/test_xlsx_handler.py tests/test_architecture_contracts.py tests/compatibility/test_v010_contract.py -q`

Expected: all tests pass and ordinary materialization records one fastexcel full materialization and zero openpyxl loads.

- [x] **Step 5: Commit the materialized reader**

```bash
git add src/messy_xlsx/parsing/fastexcel_session.py src/messy_xlsx/parsing/xlsx_materialized.py src/messy_xlsx/parsing/legacy_adapter.py src/messy_xlsx/parsing/xlsx_handler.py src/messy_xlsx/workbook.py tests/test_xlsx_materialized.py tests/test_architecture_contracts.py tests/test_resource_lifecycle.py
git commit -m "perf: materialize OOXML through fastexcel Arrow"
```

---

### Task 8: Implement Coordinate-Aware Arrow Transforms

**Files:**
- Create: `src/messy_xlsx/parsing/coordinates.py`
- Create: `tests/test_coordinate_transforms.py`
- Modify: `src/messy_xlsx/parsing/xlsx_materialized.py`
- Modify as required: `src/messy_xlsx/parsing/legacy_adapter.py`
- Modify as required: `src/messy_xlsx/parsing/router.py`
- Modify as required: `src/messy_xlsx/parsing/xlsx_handler.py`
- Modify as required: `src/messy_xlsx/workbook.py`
- Modify: `tests/test_xlsx_materialized.py`
- Modify as required: existing configuration, architecture, and compatibility tests

**Interfaces:**
- Produces: `ColumnIdentity`, `CoordinateBatch`, immutable
  `CoordinateTransform`, and a fresh stateful operation with `push()` and
  idempotent `finish()` for merge-anchor and footer/header carry-over.
- Consumes: `SheetManifest`, raw Arrow batches, and immutable `ParsePlan`.

**Mandatory corrections:**

- Do not modify `ParsePlan` or add `projection`, `legacy_mode`, or
  data-derived `display_names`; its existing resolved fields are sufficient.
- `CoordinateBatch` carries original one-based row/column coordinates plus a
  positional `column_identities` sidecar. Arrow physical field names remain
  unique ordinal strings while transforms run.
- Range, merge, hidden, auxiliary-anchor removal, skip/header/footer, and final
  identity attachment execute in that order. A range bypasses hidden filtering
  and `skip_rows`; explicit `skip_footer` still applies.
- The materialized reader still returns `pa.Table`. It opens one transform
  operation, pushes raw coordinate batches, calls `finish()`, and assembles the
  final table before the single pandas conversion.
- A no-range plan with nonzero effective `skip_rows` remains on Task 7's
  backend-pushed legacy fastexcel path. The coordinate path must not broaden
  this gate until the mixed metadata/numeric dtype contract can be matched
  exactly.
- A mixed-type merge fill that Arrow cannot represent exactly raises only a
  private `CoordinateCompatibilityError` before output. Add only that exact
  signal to transactional fallback; never classify generic Arrow, type, range,
  configuration, source, permission, memory, or process failures.

- [x] **Step 1: Write failing coordinate-precedence tests**

```python
# tests/test_coordinate_transforms.py
import pyarrow as pa

from messy_xlsx.ooxml.models import Interval, IntervalIndex, MergeRange
from messy_xlsx.parsing.coordinates import CoordinateBatch, CoordinateTransform


def _batch() -> CoordinateBatch:
    return CoordinateBatch(
        batch=pa.record_batch([["anchor", None], [1, 2]], names=["A", "B"]),
        row_numbers=pa.array([1, 2], type=pa.int64()),
        column_numbers=(1, 2),
    )


def test_range_does_not_filter_hidden_coordinates() -> None:
    transform = CoordinateTransform(
        hidden_rows=IntervalIndex((Interval(2, 2),)),
        hidden_columns=IntervalIndex(()),
        merged_ranges=(),
    )
    operation = transform.open(range_plan("A1:B2"))
    result = (*operation.push(_batch()), *operation.finish())
    assert sum(batch.batch.num_rows for batch in result) == 2


def test_merge_anchor_is_carried_across_batches() -> None:
    transform = CoordinateTransform(
        hidden_rows=IntervalIndex(()),
        hidden_columns=IntervalIndex(()),
        merged_ranges=(MergeRange(1, 1, 2, 1),),
    )
    operation = transform.open(merge_plan("fill"))
    first = operation.push(_batch().slice_rows(0, 1))
    second = (*operation.push(_batch().slice_rows(1, 1)), *operation.finish())
    assert first[0].batch.column(0)[0].as_py() == "anchor"
    assert second[0].batch.column(0)[0].as_py() == "anchor"
```

Add parameterized fixtures for range × hidden rows/columns, explicit/detected skips, every merge strategy, hidden anchors, projected-out anchors, and merges crossing boundaries.
Also retain the Task 7 mixed `("metadata",), (1,), (2,)` regression and
forbid the coordinate reader when effective no-range `skip_rows != 0`.

Run: `.venv/bin/pytest tests/test_coordinate_transforms.py -q`

Expected: collection fails because coordinate types do not exist.

- [x] **Step 2: Define positional batch identity**

```python
@dataclass(frozen=True)
class ColumnIdentity:
    ordinal: int
    display_name: object


@dataclass(frozen=True)
class CoordinateBatch:
    batch: pa.RecordBatch
    row_numbers: pa.Int64Array
    column_numbers: tuple[int, ...]
    column_identities: tuple[ColumnIdentity, ...] = ()

    def slice_rows(self, offset: int, length: int) -> "CoordinateBatch":
        return CoordinateBatch(
            self.batch.slice(offset, length),
            self.row_numbers.slice(offset, length),
            self.column_numbers,
            self.column_identities,
        )
```

Every transform addresses columns by ordinal. Display names remain payload until final output, preserving duplicate and non-string labels.

- [x] **Step 3: Implement transform ordering and anchor projection expansion**

Implement the design matrix in this exact order:

1. Determine requested original-coordinate projection.
2. Expand the read projection for any intersecting merge anchor.
3. Apply merge `fill`, `skip`, or `first_only` using persistent anchor state.
4. If no legacy `cell_range` is active, remove hidden coordinates when configured.
5. Remove auxiliary anchor coordinates added in step 2.
6. Apply explicit/detected row header/footer rules according to `ParsePlan`.
7. Attach final `ColumnIdentity` ordinals and display names.

Use Arrow arrays and boolean masks; never build a complete row list.

```python
class CoordinateTransform:
    def open(self, plan: ParsePlan) -> CoordinateOperation:
        return CoordinateOperation(self, plan)


class CoordinateOperation:
    def push(self, batch: CoordinateBatch) -> tuple[CoordinateBatch, ...]: ...

    def finish(self) -> tuple[CoordinateBatch, ...]: ...
```

The operation interprets `plan.cell_range` in original coordinates, retains
any intersecting merge anchor, applies merge semantics, filters hidden
coordinates only without a range, removes all auxiliary rows/columns/cells,
and finalizes footer/header state before attaching positional identities. The
compiler has already resolved explicit versus detected values; do not reread
`SheetConfig`. Tests enumerate the complete precedence matrix first.

- [x] **Step 4: Run transform, merge, header, and range regression tests**

Run: `.venv/bin/pytest tests/test_coordinate_transforms.py tests/test_configurations/test_merge_strategies.py tests/test_configurations/test_header_modes.py tests/test_parse_plan.py -q`

Expected: all precedence cases and existing configuration tests pass.

- [x] **Step 5: Commit coordinate transforms**

```bash
git add src/messy_xlsx/parsing/coordinates.py src/messy_xlsx/parsing/xlsx_materialized.py src/messy_xlsx/parsing/legacy_adapter.py src/messy_xlsx/parsing/router.py src/messy_xlsx/parsing/xlsx_handler.py src/messy_xlsx/workbook.py tests/test_coordinate_transforms.py tests/test_xlsx_materialized.py
git commit -m "feat: transform Arrow batches by worksheet coordinates"
```

---

### Task 9: Add the Shared Closable Stream Lifecycle

**Files:**
- Create: `src/messy_xlsx/parsing/streams.py`
- Create: `tests/test_stream_lifecycle.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `tests/test_resource_lifecycle.py`

**Interfaces:**
- Produces: public `BatchStream`, `DataFrameChunkStream`, and `SheetStream`.
- Each stream is one-shot, context-managed, idempotently closable, owner-aware, and registered as the workbook's sole active operation.

- [x] **Step 1: Write failing lifecycle contract tests**

```python
# tests/test_stream_lifecycle.py
import pyarrow as pa
import pytest

from messy_xlsx.parsing.streams import BatchStream


def test_batch_stream_is_one_shot_and_close_is_idempotent() -> None:
    closed: list[bool] = []
    batch = pa.record_batch([[1, 2]], names=["value"])
    stream = BatchStream(iter([batch]), batch.schema, lambda: closed.append(True))
    assert iter(stream) is stream
    assert next(stream).equals(batch)
    with pytest.raises(StopIteration):
        next(stream)
    stream.close()
    stream.close()
    assert closed == [True]


def test_empty_stream_exposes_schema() -> None:
    schema = pa.schema([("value", pa.int64())])
    stream = BatchStream(iter(()), schema, lambda: None)
    assert stream.schema == schema
    with pytest.raises(StopIteration):
        next(stream)


def test_workbook_rejects_a_second_active_operation(sample_xlsx) -> None:
    from messy_xlsx import MessyWorkbook

    with MessyWorkbook(sample_xlsx) as workbook:
        token = workbook._begin_operation()
        try:
            with pytest.raises(RuntimeError, match="active parse or stream"):
                workbook._begin_operation()
        finally:
            workbook._end_operation(token)
```

Run: `.venv/bin/pytest tests/test_stream_lifecycle.py -q`

Expected: collection fails because public stream types do not exist.

- [x] **Step 2: Implement one generic lifecycle core and typed wrappers**

```python
# src/messy_xlsx/parsing/streams.py
from collections.abc import Callable, Iterator
from typing import Generic, TypeVar

import pandas as pd
import pyarrow as pa


T = TypeVar("T")


class _ResultStream(Generic[T], Iterator[T]):
    def __init__(self, source: Iterator[T], close_callback: Callable[[], None]) -> None:
        self._source = source
        self._close_callback = close_callback
        self._closed = False
        self._owner_invalidated = False

    def __iter__(self):
        return self

    def __next__(self) -> T:
        if self._owner_invalidated:
            raise RuntimeError("MessyWorkbook is closed")
        if self._closed:
            raise StopIteration
        try:
            return next(self._source)
        except StopIteration:
            self.close()
            raise
        except BaseException as error:
            try:
                self.close()
            except BaseException as cleanup_error:
                error.add_note(f"stream cleanup also failed: {cleanup_error!r}")
            raise

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        primary = None
        close = getattr(self._source, "close", None)
        try:
            if close is not None:
                close()
        except BaseException as error:
            primary = error
        try:
            self._close_callback()
        except BaseException:
            if primary is None:
                raise
        if primary is not None:
            raise primary

    def invalidate_from_owner(self) -> None:
        if self._closed:
            self._owner_invalidated = True
            return
        try:
            self.close()
        finally:
            self._owner_invalidated = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()


class BatchStream(_ResultStream[pa.RecordBatch]):
    def __init__(self, source, schema: pa.Schema, close_callback) -> None:
        super().__init__(source, close_callback)
        self.schema = schema


class DataFrameChunkStream(_ResultStream[pd.DataFrame]):
    """One-shot stream of pandas chunks with deterministic cleanup."""


class SheetStream(_ResultStream["SheetResult"]):
    """One-shot stream of ordered sheet results with deterministic cleanup."""
```

- [x] **Step 3: Add workbook active-operation ownership**

Add `_active_stream`, `_closed`, `_begin_operation()`, and `_end_operation()` to `MessyWorkbook`. Validate configuration and reserve the operation before returning a stream. `MessyWorkbook.close()` calls `invalidate_from_owner()` on the child stream first. A child manually closed returns the reservation exactly once. A stream invalidated by parent closure raises `RuntimeError("MessyWorkbook is closed")` on further reads; an explicitly closed stream raises `StopIteration`.

- [x] **Step 4: Run lifecycle and failure injection tests**

Run: `.venv/bin/pytest tests/test_stream_lifecycle.py tests/test_resource_lifecycle.py tests/test_source_handle.py -q`

Expected: all tests pass, including early close and primary-exception preservation.

- [x] **Step 5: Commit the stream lifecycle**

```bash
git add src/messy_xlsx/parsing/streams.py src/messy_xlsx/workbook.py tests/test_stream_lifecycle.py tests/test_resource_lifecycle.py
git commit -m "feat: add deterministic parser streams"
```

---

### Task 10: Implement the Openpyxl Bounded-Row OOXML Reader

Completed in `e3a9e30..e78c975`; full suite: 1,740 passed. Two independent
Critical/Important rereviews are clean. The implementation follows the
hardened Task 10 brief: fallback owns streaming metrics, exact transformed
schemas are precompiled before parser I/O, and format-overhead telemetry is
deferred until the acceptance/performance slices add dedicated evidence.

**Files:**
- Create: `src/messy_xlsx/parsing/xlsx_streaming.py`
- Create: `tests/test_xlsx_streaming.py`
- Modify: `src/messy_xlsx/parsing/router.py`
- Modify: `src/messy_xlsx/parsing/coordinates.py`
- Modify: `tests/test_resource_lifecycle.py`

**Interfaces:**
- Produces: `OpenpyxlStreamingReader`, implementing `StreamingBatchReader`.
- Consumes: `SourceHandle`, `SheetManifest`, `ParsePlan`, `CoordinateTransform`, and a precompiled raw Arrow schema.
- Guarantees: one read-only worksheet pass, column-oriented buffers only, at most `batch_size` output rows per yielded batch.

- [x] **Step 1: Write failing batch-size, one-pass, formula, and cleanup tests**

```python
# tests/test_xlsx_streaming.py
from contextlib import contextmanager

import openpyxl
import pyarrow as pa
import pytest

from messy_xlsx import SheetConfig
from messy_xlsx._source import SourceHandle
from messy_xlsx.ooxml.manifest import ManifestReader
from messy_xlsx.parsing.contracts import OutputMode
from messy_xlsx.parsing.coordinates import CoordinateTransform
from messy_xlsx.parsing.parse_plan import compile_parse_plan
from messy_xlsx.parsing.xlsx_streaming import OpenpyxlStreamingReader, reader_batches


@pytest.fixture
def streaming_xlsx(tmp_path):
    path = tmp_path / "streaming.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["name", "amount"])
    for index in range(7):
        sheet.append([f"item-{index}", str(index)])
    workbook.save(path)
    workbook.close()
    return path


@contextmanager
def _reader(path, *, batch_size=2, evaluate_formulas=True):
    source = SourceHandle(path)
    manifest_reader = ManifestReader(source)
    sheet_manifest = manifest_reader.sheet("Data")
    plan = compile_parse_plan(
        SheetConfig(
            auto_detect=False,
            header_rows=0,
            normalize=False,
            sanitize_column_names=False,
            evaluate_formulas=evaluate_formulas,
        ),
        structure=None,
        format_type="xlsx",
        output_mode=OutputMode.STREAMING,
        batch_size=batch_size,
    )
    schema = pa.schema([("0", pa.string()), ("1", pa.string())])
    transform = CoordinateTransform.from_manifest(sheet_manifest)
    reader = OpenpyxlStreamingReader(source, "Data", plan, schema, transform)
    try:
        yield reader
    finally:
        reader.close()
        source.close()


def test_xlsx_batches_never_exceed_requested_output_size(streaming_xlsx) -> None:
    with _reader(streaming_xlsx, batch_size=2) as reader:
        batches = list(reader_batches(reader))
    assert batches
    assert all(0 < batch.num_rows <= 2 for batch in batches)


def test_streaming_uses_one_read_only_openpyxl_load(streaming_xlsx, monkeypatch) -> None:
    original = openpyxl.load_workbook
    calls: list[dict[str, object]] = []

    def recording(*args, **kwargs):
        calls.append(dict(kwargs))
        return original(*args, **kwargs)

    monkeypatch.setattr(openpyxl, "load_workbook", recording)
    with _reader(streaming_xlsx, batch_size=2) as reader:
        list(reader_batches(reader))
    assert len(calls) == 1
    assert calls[0]["read_only"] is True
    assert calls[0]["keep_links"] is False


def test_formula_expression_stream_uses_data_only_false(streaming_xlsx) -> None:
    workbook = openpyxl.load_workbook(streaming_xlsx)
    workbook["Data"]["A2"] = "=1+1"
    workbook.save(streaming_xlsx)
    workbook.close()
    with _reader(streaming_xlsx, evaluate_formulas=False) as reader:
        values = [value.as_py() for batch in reader_batches(reader) for value in batch.column(0)]
    assert any(isinstance(value, str) and value.startswith("=") for value in values)
```

Add failure-injection tests for first, middle, and final row batches, verifying workbook/source closure and caller cursor restoration.

Run: `.venv/bin/pytest tests/test_xlsx_streaming.py -q`

Expected: collection fails because the streaming reader does not exist.

- [x] **Step 2: Implement the read-only reader with column buffers**

```python
# src/messy_xlsx/parsing/xlsx_streaming.py
from __future__ import annotations

from contextlib import ExitStack

import openpyxl
import pyarrow as pa


class OpenpyxlStreamingReader:
    def __init__(self, source, sheet, plan, raw_schema, transform, metrics=None) -> None:
        self._stack = ExitStack()
        try:
            backend = self._stack.enter_context(source.open_backend())
            self._workbook = openpyxl.load_workbook(
                backend,
                read_only=True,
                data_only=plan.data_only,
                keep_links=False,
            )
            self._stack.callback(self._workbook.close)
            self._rows = iter(self._workbook[sheet].iter_rows(values_only=True))
        except BaseException:
            self._stack.close()
            raise
        self._plan = plan
        self._schema = raw_schema
        self._transform = transform
        self._metrics = metrics
        self._next_row = 1
        self._closed = False
        if metrics is not None:
            metrics.streaming_passes += 1

    @property
    def schema(self) -> pa.Schema:
        return self._schema

    def read_next_batch(self) -> pa.RecordBatch | None:
        columns: list[list[object]] = [[] for _field in self._schema]
        row_numbers: list[int] = []
        while len(row_numbers) < self._plan.batch_size:
            try:
                row = next(self._rows)
            except StopIteration:
                break
            row_number = self._next_row
            self._next_row += 1
            for index in range(len(columns)):
                columns[index].append(row[index] if index < len(row) else None)
            row_numbers.append(row_number)
        if not row_numbers:
            self.close()
            return None
        raw = pa.RecordBatch.from_arrays(
            [pa.array(values, type=field.type) for values, field in zip(columns, self._schema)],
            schema=self._schema,
        )
        return self._transform.apply_raw(raw, row_numbers, self._plan)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._stack.close()
```

The production implementation must accumulate transformed rows until it has at most `batch_size` output rows, because hidden/filtered input rows may shrink a raw window. It may retain openpyxl's shared-string/style tables; dedicated format-overhead metrics are deferred until the metrics model and manifest expose the required immutable evidence.

- [x] **Step 3: Add an iterator adapter that closes before propagation**

```python
def reader_batches(reader: StreamingBatchReader):
    try:
        while True:
            batch = reader.read_next_batch()
            if batch is None:
                return
            yield batch
    finally:
        reader.close()
```

Do not transparently retry after the first batch is yielded. Any later backend error closes and propagates.

- [x] **Step 4: Run streaming, coordinate, and lifecycle suites**

Run: `.venv/bin/pytest tests/test_xlsx_streaming.py tests/test_coordinate_transforms.py tests/test_stream_lifecycle.py tests/test_resource_lifecycle.py -q`

Expected: all tests pass with one read-only workbook load and deterministic cleanup.

- [x] **Step 5: Commit the OOXML streaming reader**

```bash
git add src/messy_xlsx/parsing/xlsx_streaming.py src/messy_xlsx/parsing/router.py src/messy_xlsx/parsing/coordinates.py tests/test_xlsx_streaming.py tests/test_resource_lifecycle.py
git commit -m "feat: stream OOXML rows into Arrow batches"
```

---

### Task 11: Compile Streaming Normalization and Enforce Stable Schemas

**Files:**
- Create: `src/messy_xlsx/normalization/plan.py`
- Create: `src/messy_xlsx/normalization/arrow_pipeline.py`
- Create: `tests/test_streaming_normalization.py`
- Modify: `src/messy_xlsx/exceptions.py`
- Modify: `src/messy_xlsx/normalization/__init__.py`
- Modify: `src/messy_xlsx/parsing/xlsx_streaming.py`

**Interfaces:**
- Produces: public `StreamingTypeError(NormalizationError)`.
- Produces: frozen `NormalizationPlan` and `compile_normalization_plan(sample, config) -> NormalizationPlan`.
- Produces: `ArrowNormalizationPipeline.normalize(batch, plan) -> pa.RecordBatch` with stable schema and ordinal fields.

- [x] **Step 1: Write failing schema and late-value tests**

```python
# tests/test_streaming_normalization.py
import pyarrow as pa
import pytest

from messy_xlsx import StreamingTypeError
from messy_xlsx.normalization.arrow_pipeline import ArrowNormalizationPipeline
from messy_xlsx.normalization.plan import NormalizationPlan


def test_late_incompatible_value_closes_with_typed_error() -> None:
    plan = NormalizationPlan(
        schema=pa.schema([("amount", pa.float64())]),
        display_names=("amount",),
        normalize=True,
        bypass_row_filters=False,
    )
    batch = pa.record_batch([pa.array(["1.00", "late-nonnumeric"])], names=["amount"])
    with pytest.raises(StreamingTypeError) as error:
        ArrowNormalizationPipeline().normalize(batch, plan)
    assert error.value.context["column"] == "amount"
    assert error.value.context["row_offset"] == 1


def test_all_null_columns_remain_in_stream_schema() -> None:
    plan = NormalizationPlan(
        schema=pa.schema([("empty", pa.null())]),
        display_names=("empty",),
        normalize=True,
        bypass_row_filters=False,
    )
    result = ArrowNormalizationPipeline().normalize(
        pa.record_batch([pa.nulls(3)], names=["empty"]), plan
    )
    assert result.num_columns == 1
    assert result.num_rows == 3
```

Add tests for explicit `type_hints`, duplicate labels by ordinal, locale numbers, dates, missing values, and `normalize=False` bypassing row filters.

Run: `.venv/bin/pytest tests/test_streaming_normalization.py -q`

Expected: collection fails because normalization-plan modules and `StreamingTypeError` do not exist.

- [x] **Step 2: Add the public typed error**

```python
class StreamingTypeError(NormalizationError):
    """A streamed value cannot fit the schema fixed before iteration."""

    def __init__(self, message, column=None, row_offset=None, value=None, expected_type=None):
        super().__init__(
            message,
            column=column,
            value=value,
            expected_type=expected_type,
            row_offset=row_offset,
        )
```

Export it from `messy_xlsx.__all__` in Task 12.

- [x] **Step 3: Define and compile the immutable normalization plan**

```python
# src/messy_xlsx/normalization/plan.py
from dataclasses import dataclass

import pyarrow as pa


@dataclass(frozen=True)
class NormalizationPlan:
    schema: pa.Schema
    display_names: tuple[object, ...]
    normalize: bool
    bypass_row_filters: bool


def compile_normalization_plan(sample, config) -> NormalizationPlan:
    fields = []
    for ordinal, field in enumerate(sample.schema):
        display_name = sample.display_names[ordinal]
        hinted = config.type_hints.get(display_name)
        fields.append(pa.field(str(ordinal), _arrow_type(field.type, hinted)))
    return NormalizationPlan(
        schema=pa.schema(fields),
        display_names=tuple(sample.display_names),
        normalize=config.normalize,
        bypass_row_filters=not config.normalize,
    )


def _arrow_type(observed: pa.DataType, hinted: str | None) -> pa.DataType:
    if hinted is None:
        return observed
    normalized = hinted.upper()
    if any(name in normalized for name in ("VARCHAR", "TEXT", "STRING", "CHAR")):
        return pa.string()
    if any(name in normalized for name in ("INTEGER", "INT64", "INT32")):
        return pa.int64()
    if any(name in normalized for name in ("DECIMAL", "NUMERIC", "FLOAT", "DOUBLE")):
        return pa.float64()
    if "TIMESTAMP" in normalized:
        return pa.timestamp("ns")
    if normalized == "DATE":
        return pa.date32()
    raise ValueError(f"Unsupported type hint: {hinted}")
```

Without a hint, refine the observed type with the bounded sample plus current name-based inference rules before calling `_arrow_type()`.

- [x] **Step 4: Normalize each ordinal and report the first incompatible value**

Use `pyarrow.compute` for whitespace, missing markers, numeric casts, temporal casts, masks, and row filters. If an exact Arrow operation is unavailable, convert only that column to pandas and rebuild that one Arrow array. Catch cast failures, locate the first invalid offset, and raise `StreamingTypeError` with display label, offset, safe `repr(value)`, and expected Arrow type. Never silently substitute null for a non-null invalid input.

```python
class ArrowNormalizationPipeline:
    def normalize(self, batch: pa.RecordBatch, plan: NormalizationPlan) -> pa.RecordBatch:
        if not plan.normalize:
            return batch
        arrays = []
        for ordinal, field in enumerate(plan.schema):
            arrays.append(self._normalize_column(batch.column(ordinal), field, plan, ordinal))
        return pa.RecordBatch.from_arrays(arrays, schema=plan.schema)

    def _cast_strict(self, array, field, display_name, row_base):
        try:
            return pc.cast(array, target_type=field.type, safe=True)
        except (pa.ArrowInvalid, pa.ArrowNotImplementedError) as original:
            for offset, scalar in enumerate(array):
                if not scalar.is_valid:
                    continue
                try:
                    pc.cast(pa.array([scalar.as_py()]), target_type=field.type, safe=True)
                except (pa.ArrowInvalid, pa.ArrowNotImplementedError):
                    raise StreamingTypeError(
                        "streamed value is incompatible with the fixed schema",
                        column=display_name,
                        row_offset=row_base + offset,
                        value=repr(scalar.as_py()),
                        expected_type=str(field.type),
                    ) from original
            raise
```

- [x] **Step 5: Run streaming and legacy normalization tests**

Run: `.venv/bin/pytest tests/test_streaming_normalization.py tests/test_normalization tests/test_bigquery_compatibility.py -q`

Expected: streaming rules pass and all legacy pandas normalization behavior remains unchanged.

- [x] **Step 6: Commit stable streaming normalization**

```bash
git add src/messy_xlsx/exceptions.py src/messy_xlsx/normalization src/messy_xlsx/parsing/xlsx_streaming.py tests/test_streaming_normalization.py
git commit -m "feat: normalize streaming Arrow schemas"
```

Completed through `b3c4278`. Final acceptance evidence: 2,076 tests passed,
Ruff and formatting checks passed, scoped mypy passed for all eight changed
source files, the cumulative diff and worktree were clean, and independent
architecture, lifecycle, and performance reviews returned CLEAN.

---

### Task 12: Expose Arrow, Batch, and Pandas-Chunk APIs

**Files:**
- Create: `tests/test_arrow_api.py`
- Modify: `src/messy_xlsx/__init__.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `src/messy_xlsx/parsing/streams.py`
- Modify: `src/messy_xlsx/models.py`
- Modify: `tests/test_integration.py`

**Interfaces:**
- Produces: `MessyWorkbook.to_arrow() -> pa.Table`, `iter_batches() -> BatchStream`, and `iter_dataframe_chunks() -> DataFrameChunkStream` with the exact signatures in the design spec.
- Produces: top-level `read_excel_arrow()` and `read_excel_batches()`.
- Produces: frozen public `SheetResult`, ready for the multi-sheet stream in Task 13.
- Exports: every new public type/function through `messy_xlsx.__all__`.

- [x] **Step 1: Write failing public API and ownership tests**

```python
# tests/test_arrow_api.py
import io

import pandas as pd
import pyarrow as pa

import messy_xlsx
from messy_xlsx import MessyWorkbook, read_excel_arrow, read_excel_batches


def test_new_public_exports_are_complete() -> None:
    expected = {
        "BatchStream", "DataFrameChunkStream", "SheetStream", "SheetResult",
        "LegacyAPIWarning", "StreamingTypeError", "read_excel_arrow", "read_excel_batches",
    }
    assert expected <= set(messy_xlsx.__all__)


def test_to_arrow_is_materialized_and_drops_global_null_columns(sample_xlsx) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        table = workbook.to_arrow("Data")
    assert isinstance(table, pa.Table)
    assert all(table.column(name).null_count < table.num_rows for name in table.column_names)


def test_dataframe_chunks_have_monotonic_global_range_index(sample_xlsx) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        with workbook.iter_dataframe_chunks("Data", batch_size=2) as chunks:
            frames = list(chunks)
    combined = pd.concat(frames)
    assert combined.index.tolist() == list(range(len(combined)))


def test_top_level_batch_stream_owns_workbook_until_close(sample_xlsx) -> None:
    raw = io.BytesIO(sample_xlsx.read_bytes())
    raw.seek(7)
    with read_excel_batches(raw, filename="sample.xlsx", batch_size=2) as batches:
        next(batches)
        assert raw.tell() != 7
    assert raw.tell() == 7
    assert not raw.closed
```

Run: `.venv/bin/pytest tests/test_arrow_api.py -q`

Expected: failures because public methods and conveniences are absent.

- [x] **Step 2: Add the exact workbook method signatures**

```python
@dataclass(frozen=True)
class SheetResult:
    name: str
    dataframe: pd.DataFrame | None = None
    error: SheetError | None = None

    def __post_init__(self) -> None:
        if (self.dataframe is None) == (self.error is None):
            raise ValueError("exactly one of dataframe and error must be set")


def to_arrow(
    self,
    sheet: str | None = None,
    config: SheetConfig | None = None,
) -> pa.Table:
    request = self._validate_operation_request(
        sheet, config, OutputMode.MATERIALIZED, None
    )
    with self._operation():
        plan, context = self._compile_operation(request)
        return self._materialize_arrow(plan, context)


def iter_batches(
    self,
    sheet: str | None = None,
    batch_size: int = 65_536,
    config: SheetConfig | None = None,
) -> BatchStream:
    request = self._validate_operation_request(
        sheet, config, OutputMode.STREAMING, batch_size
    )
    token = self._begin_operation()
    try:
        plan, context = self._compile_operation(request)
        reader = self._open_streaming_reader(plan, context)
        return BatchStream(reader_batches(reader), reader.schema, lambda: self._end_operation(token))
    except BaseException:
        self._end_operation(token)
        raise


def iter_dataframe_chunks(
    self,
    sheet: str | None = None,
    batch_size: int = 65_536,
    config: SheetConfig | None = None,
) -> DataFrameChunkStream:
    batches = self.iter_batches(sheet, batch_size, config)
    offset = 0

    def frames():
        nonlocal offset
        for batch in batches:
            frame = batch.to_pandas(types_mapper=pd.ArrowDtype)
            frame.index = pd.RangeIndex(offset, offset + len(frame))
            offset += len(frame)
            yield frame

    return DataFrameChunkStream(frames(), batches.close)
```

`to_arrow()` applies materialized global normalization and all-null-column removal. It does not emit `LegacyAPIWarning`.

- [x] **Step 3: Add top-level ownership-safe conveniences**

```python
def read_excel_arrow(
    file_path_or_buffer: str | Path | BinaryIO,
    sheet: str | None = None,
    config: SheetConfig | None = None,
    filename: str | None = None,
) -> pa.Table:
    with MessyWorkbook(file_path_or_buffer, filename=filename) as workbook:
        return workbook.to_arrow(sheet, config)


def read_excel_batches(
    file_path_or_buffer: str | Path | BinaryIO,
    sheet: str | None = None,
    batch_size: int = 65_536,
    config: SheetConfig | None = None,
    filename: str | None = None,
) -> BatchStream:
    workbook = MessyWorkbook(file_path_or_buffer, filename=filename)
    try:
        child = workbook.iter_batches(sheet, batch_size, config)
    except BaseException:
        workbook.close()
        raise

    def close_owned() -> None:
        primary = None
        try:
            child.close()
        except BaseException as error:
            primary = error
        try:
            workbook.close()
        except BaseException:
            if primary is None:
                raise
        if primary is not None:
            raise primary

    return BatchStream(child, child.schema, close_owned)
```

- [x] **Step 4: Export all public types and verify typing**

Update `messy_xlsx.__all__`, keep `py.typed` in the wheel, and add explicit return annotations for every new public symbol.

Run: `.venv/bin/pytest tests/test_arrow_api.py tests/test_integration.py -q`

Run: `.venv/bin/mypy src/messy_xlsx --ignore-missing-imports`

Expected: public API tests pass and mypy reports no errors.

- [x] **Step 5: Commit the additive public APIs**

```bash
git add src/messy_xlsx/__init__.py src/messy_xlsx/workbook.py src/messy_xlsx/parsing/streams.py src/messy_xlsx/models.py tests/test_arrow_api.py tests/test_integration.py
git commit -m "feat: expose Arrow and batch parsing APIs"
```

---

### Task 13: Unify Multi-Sheet Planning and Add `SheetStream`

**Files:**
- Create: `src/messy_xlsx/parsing/sheet_planner.py`
- Create: `tests/test_multi_sheet_streaming.py`
- Modify: `src/messy_xlsx/models.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `src/messy_xlsx/multi_sheet.py`
- Modify: `tests/test_parsing/test_multi_sheet.py`
- Modify: `tests/test_parsing/test_multi_sheet_robustness.py`

**Interfaces:**
- Consumes: the frozen public `SheetResult` from Task 12, with exactly one of `dataframe` and `error` non-null.
- Produces: `SheetPlanner` shared by `to_dataframes()`, `iter_sheets()`, `MultiSheetParser`, `read_all_sheets()`, and `analyze_excel()`.
- Produces: `MessyWorkbook.iter_sheets(config=None) -> SheetStream` in workbook order.

- [x] **Step 1: Write failing one-manifest, one-materialization, and error tests**

```python
# tests/test_multi_sheet_streaming.py
import openpyxl
import pytest

from messy_xlsx import MessyWorkbook, SheetResult


@pytest.fixture
def streaming_multi_sheet_xlsx(tmp_path):
    path = tmp_path / "multi.xlsx"
    source = openpyxl.Workbook()
    source.active.title = "First"
    source["First"].append(["name", "value"])
    source["First"].append(["a", 1])
    second = source.create_sheet("Second")
    second.append(["name", "value"])
    second.append(["b", 2])
    source.save(path)
    source.close()
    return path


def test_iter_sheets_preserves_order_and_result_invariant(streaming_multi_sheet_xlsx) -> None:
    with MessyWorkbook(streaming_multi_sheet_xlsx) as workbook:
        expected = workbook.sheet_names
        with workbook.iter_sheets() as stream:
            results = list(stream)
    assert [result.name for result in results] == expected
    assert all((result.dataframe is None) != (result.error is None) for result in results)


def test_iter_sheets_does_not_swallow_memory_error(streaming_multi_sheet_xlsx, monkeypatch) -> None:
    with MessyWorkbook(streaming_multi_sheet_xlsx) as workbook:
        monkeypatch.setattr(workbook, "_parse_sheet", lambda *args, **kwargs: (_ for _ in ()).throw(MemoryError()))
        with workbook.iter_sheets() as stream:
            with pytest.raises(MemoryError):
                list(stream)


def test_multi_sheet_metrics_count_one_manifest_and_one_materialization_per_sheet(
    streaming_multi_sheet_xlsx,
) -> None:
    with MessyWorkbook(streaming_multi_sheet_xlsx) as workbook:
        workbook.to_dataframes()
        assert workbook.parse_metrics.manifest_builds == 1
        assert workbook.parse_metrics.full_materializations == len(workbook.sheet_names)
```

Run: `.venv/bin/pytest tests/test_multi_sheet_streaming.py -q`

Expected: failures because `iter_sheets()` and multi-sheet metrics are incomplete.

- [x] **Step 2: Verify the frozen result invariant at the stream boundary**

```python
def _success(name: str, dataframe: pd.DataFrame) -> SheetResult:
    return SheetResult(name=name, dataframe=dataframe)


def _failure(name: str, error: BaseException) -> SheetResult:
    return SheetResult(
        name=name,
        error=SheetError(
            sheet_name=name,
            error_type=type(error).__name__,
            message=str(error),
            context=getattr(error, "context", {}),
        ),
    )
```

- [x] **Step 3: Implement one shared sheet-planning pass**

`SheetPlanner` consumes one `ManifestReader`, bounded structure samples, `SheetConfig`, and optional `MultiSheetOptions`. It returns immutable per-sheet plans in workbook order. Apply explicit sheet lists and `sheet_filter` before full value reads. Preserve legacy pivot/empty/minimum-size decisions in `MultiSheetParser` characterization tests.

Move the existing `SheetInfo` dataclass to `models.py` and re-export it from
`multi_sheet.py` and the package root. This lets `sheet_planner.py` consume the
model without importing `multi_sheet.py` back into the parser layer.

```python
@dataclass(frozen=True)
class PlannedSheet:
    name: str
    info: SheetInfo
    parse_plan: ParsePlan | None


class SheetPlanner:
    def plan(self, names, config, options=None) -> tuple[PlannedSheet, ...]:
        planned = []
        for name in names:
            info = self._analyze_bounded(name, options)
            selected = self._selected(info, options)
            parse_plan = self._compile(name, config) if selected else None
            planned.append(PlannedSheet(name, info, parse_plan))
        return tuple(planned)
```

- [x] **Step 4: Implement `SheetStream` and warning-safe legacy adapters**

Close each sheet-local parser before yielding its materialized DataFrame. Convert ordinary per-sheet `Exception` failures into `SheetError`; close and propagate `MemoryError`, `KeyboardInterrupt`, and `SystemExit`. Make `read_all_sheets()` call the private shared plan and materialization methods so it emits one warning and performs no raw-then-final duplicate parse.

- [x] **Step 5: Run all multi-sheet and compatibility tests**

Run: `.venv/bin/pytest tests/test_multi_sheet_streaming.py tests/test_parsing/test_multi_sheet.py tests/test_parsing/test_multi_sheet_robustness.py tests/compatibility/test_legacy_warnings.py tests/compatibility/test_v010_contract.py -q`

Expected: all tests pass; one manifest and one successful full materialization are recorded per selected sheet.

- [x] **Step 6: Commit unified multi-sheet planning**

```bash
git add src/messy_xlsx/models.py src/messy_xlsx/workbook.py src/messy_xlsx/multi_sheet.py src/messy_xlsx/parsing/sheet_planner.py tests/test_multi_sheet_streaming.py tests/test_parsing/test_multi_sheet.py tests/test_parsing/test_multi_sheet_robustness.py
git commit -m "refactor: unify multi-sheet planning"
```

---

### Task 14: Optimize CSV, TSV, and TXT Paths and Add Streaming Batches

> **Superseded and expanded on 2026-07-27.** The pandas-chunk full-pass
> implementation below was invalidated by adversarial review: pandas changes
> malformed-row behavior at chunk boundaries and reads far beyond a public
> batch. The normative design is now
> `docs/superpowers/specs/2026-07-26-native-csv-tokenizer-design.md`.
> The approved executable plan is
> `docs/superpowers/plans/2026-07-28-native-csv-tokenizer.md`; it is the only
> Task 14 implementation authority. Task 14 cannot be marked complete until
> that plan's exact-SHA acceptance task passes.

**Files:**
- Create: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `src/messy_xlsx/parsing/csv_native.py`
- Create: `src/messy_xlsx/parsing/csv_value_adapter.py`
- Create: `src/messy_xlsx/parsing/csv_streaming.py`
- Create: native oracle, ABI, safety, wheel, and benchmark tests defined by the
  native-tokenizer implementation plan
- Create: `tests/test_csv_streaming.py`
- Modify: `src/messy_xlsx/parsing/csv_handler.py`
- Delete: `src/messy_xlsx/parsing/csv_io.py` after moving its retained probe
  helpers into `csv_probe.py` and its reference implementation under tests
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `pyproject.toml`
- Modify: `.github/workflows/test.yml`
- Modify: `.github/workflows/publish.yml`
- Modify: `tests/test_parsing/test_csv_handler.py`
- Modify: `tests/test_edge_cases/test_csv_variations.py`
- Modify: `tests/test_source_handle.py`

**Interfaces:**
- Produces: internal `NativeCSVReader`, implementing `StreamingBatchReader`
  through the native tokenizer only after all routing gates pass.
- Produces: `NativeCSVTokenizer`, an internal Cython component with an exact
  API/version handshake and bounded evidence/full-pass contracts.
- Produces: typed per-operation CSV execution decisions and reason counters for
  native, built-in materialized fallback, and custom SPI execution.
- Preserves: schema-compatible `CSVHandler` behavior under `pandas==3.0.5`,
  using C-engine semantics without a footer and Python-engine semantics with a
  footer, subject to the approved bounded stable-schema and late path-decode
  streaming exceptions.
- Preserves: materialized custom-registry authority and caller-stream
  ownership.

- [ ] **Stage 0: Prove the native ABI and safety shell**

Build and execute the exact Cython extension type, typed memoryview,
Python-source read, callback, allocation, reentrancy, and cleanup shell across
all claimed platforms and CPython 3.11–3.14. Pass `abi3audit --strict` before
tokenizer implementation. Run a shell-level sanitizer/debug smoke, but defer
the complete semantic ASan/UBSan, debug-allocator, and allocation-failure
matrix until the tokenizer and its test-owned allocation manifest exist. If
ABI3 fails, use reviewed per-minor native wheels.

- [ ] **Stage 1: Freeze engine-specific materialized behavior**

Add deterministic fixtures and the fixed-seed fuzz generator/worker contract
for both materialized branches before semantic implementation. Freeze headers,
multi-headers, skip rows, physical pandas scalar types, malformed rows,
NUL/quote behavior, footer ordering, encodings, warnings, errors, and
lifecycle. Activate and run each engine's generated native differential subset
as that semantic slice lands; run the complete 5,000-case-per-engine gate
after both engines are implemented.

- [ ] **Stage 2: Implement bounded evidence and native routing shell**

Implement evidence statuses, header plans, stable schema compilation,
materialized fallback, exact built-in registry eligibility, backend metrics,
the exact production evidence budgets, the environment kill switch, and the
native API/semantic handshake. The source-controlled production gate remains
false.

- [ ] **Stage 3: Implement C semantic mode**

Implement bounded decoding, framing, quoting, NUL/quote-junk behavior, implicit
indexes, malformed rows, physical scalar conversion, and fixed-buffer
lifecycle for `skip_footer == 0`.

- [ ] **Stage 4: Implement Python/footer semantic mode**

Implement physical-line skipping, Python quote-error behavior, footer
retention, row limits, and deterministic `batch_size + skip_footer` counters.
Generated multi-header processing selects materialized fallback in v1.0.0 to
preserve its whole-column pandas type-rendering contract.

- [ ] **Stage 5: Integrate with production routing disabled**

Connect native physical values to the existing Arrow and normalization
pipeline. Remove the pandas-chunk reader and installed Python
framing/filter/footer full pass, retaining only test references and inspection
probes. Pass semantic, lifecycle, memory, safety, and performance gates while
default production routing remains disabled.

- [ ] **Stage 6: Build and verify disabled candidate artifacts**

Switch to the approved setuptools dual build modes. Build all native and
universal wheels from one sdist, execute the complete platform/runtime matrix,
audit artifacts, and verify the internal native adapter, ABI3 runtimes, and
resolver/runtime selection through the private candidate smoke seam while
public production routing remains disabled.

- [ ] **Stage 7: Enable routing and rebuild final artifacts**

After the candidate matrix passes, make the production-gate constant the only
functional source change that enables exact built-in native routing. Build a
new final sdist and all final wheels from that exact revision and rerun the
complete artifact, public-routing, kill-switch, resolver, and runtime matrix.
This second set proves the implementation revision, but later
README/changelog/release-metadata changes invalidate its hashes. Only parent
Task 20's newly rebuilt and reverified final set from the exact
post-documentation release SHA is releasable.

- [ ] **Stage 8: Commit Task 14**

Close only after independent compatibility, native-safety, packaging,
performance, and whole-repository reviews are clean. Record the checkpoint
outside Git so the accepted final SHA remains unchanged; parent Task 20 folds
the SDD ledger and Task 14 tracker update into its release-documentation commit
and rebuilds the final artifact matrix on that exact new SHA.

---

### Task 15: Add XLS Row Batches and Preserve Custom Registry Overrides

**Files:**
- Create: `src/messy_xlsx/parsing/xls_streaming.py`
- Create: `tests/test_xls_streaming.py`
- Modify: `src/messy_xlsx/parsing/xls_handler.py`
- Modify: `src/messy_xlsx/parsing/handler_registry.py`
- Modify: `tests/test_parsing/test_xls_handler.py`
- Modify: `tests/test_architecture_contracts.py`
- Modify: `tests/test_source_handle.py`

**Interfaces:**
- Produces: optional `XLSStreamingReader` with bounded row/column buffers over xlrd's workbook model.
- Produces: `CustomDataFrameReader` that preserves every existing registry/detector/handler override and explicitly reports materialized compatibility mode.

- [ ] **Step 1: Write failing XLS batch and extension-routing tests**

```python
# tests/test_xls_streaming.py
import pytest

from messy_xlsx import MessyWorkbook


@pytest.fixture
def streaming_xls(tmp_path):
    xlwt = pytest.importorskip("xlwt")
    pytest.importorskip("xlrd")
    path = tmp_path / "streaming.xls"
    book = xlwt.Workbook()
    sheet = book.add_sheet("Data")
    for row, values in enumerate((("name", "value"), ("a", 1), ("b", 2), ("c", 3))):
        for column, value in enumerate(values):
            sheet.write(row, column, value)
    book.save(str(path))
    return path


def test_xls_batches_respect_batch_size(streaming_xls) -> None:
    with MessyWorkbook(streaming_xls) as workbook:
        with workbook.iter_batches(batch_size=3) as stream:
            assert all(batch.num_rows <= 3 for batch in stream)
```

Add architecture tests for a registry subclass overriding each of `parse`, `detect_format`, `validate`, and `get_sheet_names`. Assert every override is called and no built-in reader bypasses it. Assert a custom DataFrame batch stream is labeled `CUSTOM_DATAFRAME` and documented as materialized.

Run: `.venv/bin/pytest tests/test_xls_streaming.py tests/test_architecture_contracts.py -q`

Expected: XLS streaming is absent and at least one custom-registry route is bypassed.

- [ ] **Step 2: Implement optional xlrd row-window batches**

```python
from contextlib import ExitStack
from pathlib import Path

import pyarrow as pa
import xlrd


class XLSStreamingReader:
    def __init__(self, source, sheet_name, plan, schema, transform, normalizer) -> None:
        self._stack = ExitStack()
        try:
            backend = self._stack.enter_context(source.open_path_or_bytes())
            kwargs = {"on_demand": True}
            if isinstance(backend, Path):
                kwargs["filename"] = str(backend)
            else:
                kwargs["file_contents"] = backend
            self._book = xlrd.open_workbook(**kwargs)
            self._stack.callback(self._book.release_resources)
            self._sheet = self._book.sheet_by_name(sheet_name)
        except BaseException:
            self._stack.close()
            raise
        self._plan = plan
        self._schema = schema
        self._transform = transform
        self._normalizer = normalizer
        self._row = 0
        self._closed = False

    @property
    def schema(self) -> pa.Schema:
        return self._schema

    def read_next_batch(self):
        if self._row >= self._sheet.nrows:
            return None
        end = min(self._row + self._plan.batch_size, self._sheet.nrows)
        columns = [
            [self._sheet.cell_value(row, column) for row in range(self._row, end)]
            for column in range(self._sheet.ncols)
        ]
        start = self._row
        self._row = end
        raw = pa.record_batch(columns, names=[str(index) for index in range(len(columns))])
        transformed = self._transform.apply_raw(raw, range(start + 1, end + 1), self._plan)
        return self._normalizer.normalize(transformed.batch, self._plan.normalization)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._stack.close()
```

- [ ] **Step 3: Implement the custom DataFrame compatibility reader**

```python
class CustomDataFrameReader:
    def __init__(self, registry, source, sheet, options, format_type) -> None:
        self._registry = registry
        self._source = source
        self._sheet = sheet
        self._options = options
        self._format_type = format_type

    def read_table(self) -> pa.Table:
        frame = self._registry.parse(
            self._source,
            sheet=self._sheet,
            options=self._options,
            format_type=self._format_type,
        )
        return pa.Table.from_pandas(frame, preserve_index=False)
```

Batch APIs over this adapter may slice the materialized table for output convenience, but metrics and docs must state that input memory is not bounded.

- [ ] **Step 4: Run XLS, extension, and lifecycle tests**

Run: `.venv/bin/pytest tests/test_xls_streaming.py tests/test_parsing/test_xls_handler.py tests/test_architecture_contracts.py tests/test_source_handle.py -q`

Expected: all tests pass with and without the optional XLS dependency; custom overrides remain authoritative.

- [ ] **Step 5: Commit XLS and extension compatibility**

```bash
git add src/messy_xlsx/parsing/xls_streaming.py src/messy_xlsx/parsing/xls_handler.py src/messy_xlsx/parsing/handler_registry.py tests/test_xls_streaming.py tests/test_parsing/test_xls_handler.py tests/test_architecture_contracts.py tests/test_source_handle.py
git commit -m "feat: stream XLS rows and preserve custom handlers"
```

---

### Task 16: Reduce Legacy Normalization Copies Without Changing Results

**Files:**
- Create: `tests/test_normalization/test_copy_budget.py`
- Modify: `src/messy_xlsx/normalization/pipeline.py`
- Modify: `src/messy_xlsx/normalization/whitespace.py`
- Modify: `src/messy_xlsx/normalization/numbers.py`
- Modify: `src/messy_xlsx/normalization/dates.py`
- Modify: `src/messy_xlsx/normalization/missing_values.py`
- Modify: `src/messy_xlsx/normalization/type_coercion.py`
- Modify: `src/messy_xlsx/workbook.py`

**Interfaces:**
- Produces: `NormalizationPipeline.normalize_owned(frame, semantic_hints=None, skip_steps=None, drop_regex=None, drop_conditions=()) -> pd.DataFrame`.
- Guarantees: one explicit full-frame ownership copy, one combined regex/condition row mask, and exact v0.10.0 golden output.

- [ ] **Step 1: Write failing copy-budget and combined-filter tests**

```python
# tests/test_normalization/test_copy_budget.py
import pandas as pd

from messy_xlsx.normalization.pipeline import NormalizationPipeline


def test_pipeline_takes_one_explicit_ownership_copy(monkeypatch) -> None:
    pipeline = NormalizationPipeline()
    calls = 0
    original = pipeline._take_ownership

    def recording(frame):
        nonlocal calls
        calls += 1
        return original(frame)

    monkeypatch.setattr(pipeline, "_take_ownership", recording)
    frame = pd.DataFrame({"amount": [" 1.00 ", "2.00"], "note": ["ok", "drop"]})
    pipeline.normalize_owned(frame)
    assert calls == 1


def test_combined_row_filters_apply_one_final_index_reset() -> None:
    frame = pd.DataFrame({"kind": ["keep", "footer", "drop"], "value": [1, 2, 3]})
    result = NormalizationPipeline().normalize_owned(
        frame,
        drop_regex="footer",
        drop_conditions=(("kind", "drop"),),
    )
    assert result.to_dict(orient="list") == {"kind": ["keep"], "value": [1]}
    assert result.index.tolist() == [0]
```

Run: `.venv/bin/pytest tests/test_normalization/test_copy_budget.py -q`

Expected: failure because each normalizer currently copies and `normalize_owned()` is absent.

- [ ] **Step 2: Give the pipeline sole ownership of the working frame**

```python
def normalize_owned(
    self,
    frame,
    semantic_hints=None,
    skip_steps=None,
    drop_regex=None,
    drop_conditions=(),
):
    owned = self._take_ownership(frame)
    hints = semantic_hints or self.type_inference.infer_types(owned)
    for step_name, normalizer in self._ordered_steps():
        if step_name not in (skip_steps or ()):
            normalizer.normalize_owned(owned, hints)
    mask = self._combined_drop_mask(owned, drop_regex, drop_conditions)
    if mask is not None:
        owned = owned.loc[~mask]
    return owned.reset_index(drop=True)

def _take_ownership(self, frame):
    return frame.copy(deep=True)
```

Change each stage to `normalize_owned()` and positional `DataFrame.isetitem()` mutation without calling `DataFrame.copy()`. Keep existing `normalize()` methods as compatibility wrappers that copy once and delegate.

- [ ] **Step 3: Build regex and condition masks column by column**

Initialize one boolean Series aligned to the frame index. OR regex matches one column at a time; OR all valid condition matches; apply one final filter and reset. Preserve the current rule that row filters are bypassed when normalization is disabled.

- [ ] **Step 4: Run normalization, BigQuery, and golden contracts**

Run: `.venv/bin/pytest tests/test_normalization tests/test_bigquery_compatibility.py tests/compatibility/test_v010_contract.py -q`

Expected: all values/dtypes remain identical and the explicit copy budget passes.

- [ ] **Step 5: Commit normalization consolidation**

```bash
git add src/messy_xlsx/normalization src/messy_xlsx/workbook.py tests/test_normalization/test_copy_budget.py
git commit -m "perf: consolidate legacy normalization copies"
```

---

### Task 17: Index Cell Metadata and Preserve the Formula Boundary

**Files:**
- Create: `src/messy_xlsx/cell_index.py`
- Create: `tests/test_cell_indexes.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `src/messy_xlsx/sheet.py`
- Modify: `tests/test_configurations/test_formula_modes.py`
- Modify: `tests/test_formulas/test_engine.py`

**Interfaces:**
- Produces: `CellMetadataIndex` for logarithmic merge lookup and constant-time hidden membership.
- Preserves: `get_cell()` and `iter_rows()` full `CellValue` metadata; `FormulaConfig` remains cell-only and `SheetConfig.evaluate_formulas` remains table-only.

- [ ] **Step 1: Write failing index reuse and formula-boundary tests**

```python
# tests/test_cell_indexes.py
import openpyxl

from messy_xlsx import MessyWorkbook, SheetConfig


def test_cell_metadata_index_is_built_once(sample_xlsx, monkeypatch) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        calls = 0
        original = workbook._build_cell_index

        def recording(sheet):
            nonlocal calls
            calls += 1
            return original(sheet)

        monkeypatch.setattr(workbook, "_build_cell_index", recording)
        workbook.get_cell("Data", 1, 1)
        workbook.get_cell("Data", 2, 1)
    assert calls == 1


def test_iter_rows_never_calls_batch_reader(sample_xlsx, monkeypatch) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        monkeypatch.setattr(
            workbook,
            "iter_batches",
            lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("batch path forbidden")),
        )
        rows = list(workbook.get_sheet("Data").iter_rows(max_row=2, max_col=2))
    assert len(rows) == 2


def test_formula_config_does_not_trigger_table_calculation(tmp_path, monkeypatch) -> None:
    path = tmp_path / "formula.xlsx"
    source = openpyxl.Workbook()
    source.active.title = "Data"
    source["Data"].append(["value"])
    source["Data"].append(["=1+1"])
    source.save(path)
    source.close()
    with MessyWorkbook(path) as workbook:
        monkeypatch.setattr(
            workbook._formula_engine,
            "evaluate",
            lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("cell engine called")),
        )
        workbook.to_dataframe(config=SheetConfig(evaluate_formulas=True))
```

Run: `.venv/bin/pytest tests/test_cell_indexes.py -q`

Expected: failures because `_build_cell_index` and the cached index do not exist.

- [ ] **Step 2: Implement compact cell metadata indexes**

```python
# src/messy_xlsx/cell_index.py
from bisect import bisect_right
from dataclasses import dataclass

from messy_xlsx.ooxml.models import IntervalIndex, MergeRange


@dataclass(frozen=True)
class CellMetadataIndex:
    merge_starts: tuple[int, ...]
    merge_prefix_max_rows: tuple[int, ...]
    merge_ranges: tuple[MergeRange, ...]
    hidden_rows: IntervalIndex
    hidden_columns: IntervalIndex

    def merge_at(self, row: int, column: int) -> MergeRange | None:
        position = bisect_right(self.merge_starts, row) - 1
        while position >= 0 and self.merge_prefix_max_rows[position] >= row:
            candidate = self.merge_ranges[position]
            if candidate.min_row <= row <= candidate.max_row and candidate.min_col <= column <= candidate.max_col:
                return candidate
            position -= 1
        return None

    def is_hidden(self, row: int, column: int) -> bool:
        return self.hidden_rows.contains(row) or self.hidden_columns.contains(column)
```

Sort merge ranges by `min_row` and build `merge_prefix_max_rows` as the running maximum of `max_row`. Cache one index per sheet after the lazy compatibility workbook is opened. Do not expand merged rectangles or hidden intervals into individual cells.

- [ ] **Step 3: Route cell methods through the index and keep `iter_rows()` unchanged**

Replace linear `_is_cell_merged()` and repeated dimension-set construction with index queries. Keep `MessySheet.iter_rows()` calling `get_cell()` so formulas, formats, merges, and hidden flags remain present.

- [ ] **Step 4: Run cell, formula, lifecycle, and compatibility tests**

Run: `.venv/bin/pytest tests/test_cell_indexes.py tests/test_configurations/test_formula_modes.py tests/test_formulas tests/test_resource_lifecycle.py tests/compatibility/test_v010_contract.py -q`

Expected: all cell metadata and formula contracts pass with one index construction per sheet.

- [ ] **Step 5: Commit cell metadata indexes**

```bash
git add src/messy_xlsx/cell_index.py src/messy_xlsx/workbook.py src/messy_xlsx/sheet.py tests/test_cell_indexes.py tests/test_configurations/test_formula_modes.py tests/test_formulas/test_engine.py
git commit -m "perf: index merged and hidden cell metadata"
```

---

### Task 18: Complete Security, Failure, Property, and Corpus Integration

**Files:**
- Create: `tests/test_transactional_fallback.py`
- Create: `tests/test_supported_formats.py`
- Modify: `tests/test_property_based/test_hypothesis.py`
- Modify: `tests/test_resource_lifecycle.py`
- Modify: `tests/test_edge_cases/test_malformed_files.py`
- Modify: `tests/test_architecture_contracts.py`
- Modify: `tests/test_samples.py`
- Modify: `tests/compatibility/test_v010_contract.py`

**Interfaces:**
- Verifies: transactional fallback, hostile OOXML limits, duplicate headers, late schemas, batch-boundary transforms, all supported extensions, full maintained/generated corpus, and resource cleanup.
- Produces no new public runtime interface; runtime fixes remain scoped to the failing component.

- [ ] **Step 1: Write transactional fallback teardown tests**

```python
# tests/test_transactional_fallback.py
import pyarrow as pa
import pytest

from messy_xlsx.parsing.fallback import FallbackCoordinator


class MaterializedFake:
    def __init__(self, events, name, error=None) -> None:
        self.events = events
        self.name = name
        self.error = error

    def __enter__(self):
        self.events.append(f"{self.name}-open")
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.events.append(f"{self.name}-close")

    def read_table(self):
        if self.error is not None:
            raise self.error
        return pa.table({"value": [1]})


class StreamingFake:
    def __init__(self, batches, error, opens) -> None:
        self.batches = iter(batches)
        self.error = error
        self.opens = opens
        self.closed = False

    def read_next_batch(self):
        try:
            return next(self.batches)
        except StopIteration:
            if self.error is not None:
                error, self.error = self.error, None
                raise error
            return None

    def close(self) -> None:
        self.closed = True


def test_failed_reader_is_closed_before_fallback_opens() -> None:
    events: list[str] = []
    coordinator = FallbackCoordinator(lambda error: isinstance(error, ValueError))
    result = coordinator.materialize(
        lambda: MaterializedFake(events, "primary", ValueError("classified compatibility")),
        lambda: MaterializedFake(events, "fallback"),
    )
    assert result.num_rows == 1
    assert events.index("primary-close") < events.index("fallback-open")


def test_stream_never_retries_after_first_yield() -> None:
    opens = {"fallback": 0}
    first = pa.record_batch([[1]], names=["value"])
    primary = StreamingFake([first], ValueError("injected batch failure"), opens)

    def fallback_factory():
        opens["fallback"] += 1
        return StreamingFake([], None, opens)

    stream = FallbackCoordinator(lambda error: isinstance(error, ValueError)).batches(
        lambda: primary,
        fallback_factory,
    )
    assert next(stream).equals(first)
    with pytest.raises(ValueError, match="injected batch failure"):
        next(stream)
    assert primary.closed
    assert opens["fallback"] == 0
```

- [ ] **Step 2: Add hostile archive and relationship properties**

Extend Hypothesis strategies to generate member paths, duplicate names, compressed/uncompressed sizes, XML prefixes, merge intervals, hidden intervals, and formula samples. Assert every rejected archive raises `FormatError`, never performs network access, and leaves no spool path.

- [ ] **Step 3: Add schema and coordinate boundary properties**

Generate batch sizes from 1 to 128, merges spanning arbitrary boundaries, hidden rows around headers/footers, duplicate display names, and late incompatible values. Assert concatenated successful streams preserve row order, never exceed batch size, maintain one schema, and either exhaust deterministically or raise `StreamingTypeError` at the first incompatible batch.

- [ ] **Step 4: Cover every accepted extension and source shape**

```python
# tests/test_supported_formats.py
import io

import pytest

from messy_xlsx import MessyWorkbook


@pytest.mark.parametrize("suffix", [".xlsx", ".xlsm", ".xltx", ".xltm"])
def test_ooxml_extensions_use_the_same_content_contract(sample_xlsx, tmp_path, suffix) -> None:
    path = tmp_path / f"sample{suffix}"
    path.write_bytes(sample_xlsx.read_bytes())
    with MessyWorkbook(path) as workbook:
        assert workbook.to_dataframe().shape[0] > 0


@pytest.mark.parametrize("filename", ["data.csv", "data.tsv", "data.txt"])
def test_text_extensions_accept_seekable_buffers(filename) -> None:
    delimiter = "\t" if filename.endswith(".tsv") else ","
    source = io.BytesIO(f"a{delimiter}b\n1{delimiter}2\n".encode())
    with MessyWorkbook(source, filename=filename) as workbook:
        assert workbook.to_dataframe().shape == (1, 2)
```

- [ ] **Step 5: Run the complete compatibility and generated corpus**

Run: `.venv/bin/pytest tests/compatibility tests/test_property_based tests/test_resource_lifecycle.py tests/test_edge_cases tests/test_architecture_contracts.py tests/test_supported_formats.py tests/test_samples.py -q`

Expected: every test passes; no unresolved compatibility, security, or cleanup failure remains.

- [ ] **Step 6: Commit integration hardening**

```bash
git add tests/test_transactional_fallback.py tests/test_supported_formats.py tests/test_property_based tests/test_resource_lifecycle.py tests/test_edge_cases tests/test_architecture_contracts.py tests/test_samples.py tests/compatibility
git commit -m "test: harden parser integration contracts"
```

---

### Task 19: Add Reproducible Performance Comparison and CI Gates

**Files:**
- Create: `scripts/benchmark_worker.py`
- Create: `scripts/compare_benchmarks.py`
- Create: `benchmarks/requirements.txt`
- Create: `tests/test_performance/test_benchmark_contract.py`
- Create: `.github/workflows/performance.yml`
- Modify: `.github/workflows/test.yml`
- Modify: `pyproject.toml`

**Interfaces:**
- Produces: JSON benchmark records with wall time, Linux `ru_maxrss`, shape, schema, stable output hash, parser counters, Python/dependency versions, and CPU model.
- Produces: paired v0.10.0/candidate median gates on the same runner with one warm-up and five measured processes.

- [ ] **Step 1: Write failing benchmark schema and comparison tests**

```python
# tests/test_performance/test_benchmark_contract.py
from scripts.compare_benchmarks import compare


def test_xlsx_relative_gate_requires_four_times_speed_and_half_rss() -> None:
    baseline = {"median_seconds": 10.0, "median_rss_mb": 600.0, "output_hash": "same"}
    candidate = {"median_seconds": 2.0, "median_rss_mb": 290.0, "output_hash": "same"}
    assert compare("xlsx_100k", baseline, candidate) == []


def test_output_mismatch_always_fails() -> None:
    baseline = {"median_seconds": 10.0, "median_rss_mb": 600.0, "output_hash": "old"}
    candidate = {"median_seconds": 1.0, "median_rss_mb": 100.0, "output_hash": "new"}
    assert "output hash differs" in compare("xlsx_100k", baseline, candidate)


def test_csv_seekable_rss_must_be_within_twenty_percent_of_path() -> None:
    record = {"csv_path_rss_mb": 200.0, "csv_seekable_rss_mb": 245.0}
    assert "CSV seekable RSS exceeds 120% of path RSS" in compare("csv_300k", record, record)
```

Run: `.venv/bin/pytest tests/test_performance/test_benchmark_contract.py -q`

Expected: collection fails because benchmark scripts do not exist.

- [ ] **Step 2: Implement the isolated benchmark worker**

```python
# scripts/benchmark_worker.py
from __future__ import annotations

import argparse
import hashlib
import json
import resource
import time

import pandas as pd

from messy_xlsx import read_excel


def run_xlsx(path: str) -> dict[str, object]:
    started = time.perf_counter()
    frame = read_excel(path)
    elapsed = time.perf_counter() - started
    payload = frame.to_json(orient="split", date_format="iso", default_handler=str)
    return {
        "seconds": elapsed,
        "rss_mb": resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024,
        "shape": list(frame.shape),
        "dtypes": [str(dtype) for dtype in frame.dtypes],
        "output_hash": hashlib.sha256(payload.encode()).hexdigest(),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", required=True)
    args = parser.parse_args()
    print(json.dumps(run_xlsx(args.xlsx), sort_keys=True))


if __name__ == "__main__":
    main()
```

Add worker modes for raw/normalized CSV path and seekable buffer, multi-sheet counters, Arrow batches, and shared-string-heavy OOXML. Include dependency versions and `/proc/cpuinfo` model where available.

- [ ] **Step 3: Implement paired orchestration and median comparison**

`compare_benchmarks.py` creates temporary baseline and candidate virtual environments, installs the `v0.10.0` tag and current checkout with the same pinned benchmark dependency set, performs one unrecorded warm-up, runs five fresh worker subprocesses per workload, and writes `baseline.json`, `candidate.json`, and `comparison.json`.

```text
# benchmarks/requirements.txt
defusedxml==0.7.1
fastexcel==0.20.2
numpy==2.5.1
openpyxl==3.1.5
pandas==3.0.3
pyarrow==25.0.0
xlrd==2.0.2
```

Install this file first in both environments, then install each checkout with
`--no-deps` so code revision is the only deliberate variable.

```python
def compare(workload, baseline, candidate):
    failures = []
    if baseline.get("output_hash") != candidate.get("output_hash"):
        failures.append("output hash differs")
    if workload == "xlsx_100k":
        if candidate["median_seconds"] > baseline["median_seconds"] / 4:
            failures.append("XLSX runtime is not at least 4x faster")
        if candidate["median_rss_mb"] > baseline["median_rss_mb"] * 0.5:
            failures.append("XLSX RSS is not at least 50% lower")
    if workload == "csv_300k":
        if candidate["csv_seekable_rss_mb"] > candidate["csv_path_rss_mb"] * 1.2:
            failures.append("CSV seekable RSS exceeds 120% of path RSS")
    return failures
```

- [ ] **Step 4: Add the dedicated Linux performance workflow**

```yaml
# .github/workflows/performance.yml
name: Performance

on:
  workflow_dispatch:
  pull_request:
    branches: [main]
    paths:
      - "src/messy_xlsx/**"
      - "scripts/benchmark_*.py"
      - "scripts/compare_benchmarks.py"
      - "tests/samples/**"

permissions:
  contents: read

jobs:
  compare:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@d23441a48e516b6c34aea4fa41551a30e30af803
        with:
          fetch-depth: 0
      - uses: actions/setup-python@ece7cb06caefa5fff74198d8649806c4678c61a1
        with:
          python-version: "3.12"
      - run: python scripts/compare_benchmarks.py --baseline v0.10.0 --candidate . --output benchmark-results
      - uses: actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a
        with:
          name: parser-performance
          path: benchmark-results/
          if-no-files-found: error
```

- [ ] **Step 5: Run benchmark contract tests and a local smoke workload**

Run: `.venv/bin/pytest tests/test_performance/test_benchmark_contract.py -q`

Expected: comparison contract tests pass.

Run: `.venv/bin/python scripts/benchmark_worker.py --xlsx tests/samples/sales_transactions.xlsx`

Expected: one valid JSON record with matching shape/hash metadata and non-zero timing/RSS.

- [ ] **Step 6: Commit performance automation**

```bash
git add scripts/benchmark_worker.py scripts/compare_benchmarks.py benchmarks/requirements.txt tests/test_performance/test_benchmark_contract.py .github/workflows/performance.yml .github/workflows/test.yml pyproject.toml
git commit -m "ci: gate parser performance against v0.10.0"
```

---

### Task 20: Finish Documentation, Versioning, Packaging, and the v1.0.0 Gate

**Files:**
- Modify: `README.md`
- Modify: `docs/index.md`
- Modify: `docs/getting-started.md`
- Modify: `docs/configuration.md`
- Modify: `docs/api.md`
- Modify: `docs/changelog.md`
- Modify: `mkdocs.yml`
- Modify: `CHANGELOG.md`
- Modify: `pyproject.toml`
- Modify: `src/messy_xlsx/__init__.py`
- Modify: `.github/workflows/test.yml`
- Modify: `.github/workflows/publish.yml`
- Modify: `tests/test_integration.py`
- Create: `scripts/smoke_release_install.py`
- Create: `tests/packaging/test_release_smoke_cli.py`

**Interfaces:**
- Publishes: documented legacy migration table, Arrow/stream examples, source ownership, formula boundary, streaming normalization boundary, custom-handler limits, security limits, and performance guidance.
- Verifies: package and module versions already equal `1.0.0`; release workflow
  continues to publish only the `v1.0.0` tag at the tip of `main` after all
  gates pass.

- [ ] **Step 1: Write failing version/export/documentation smoke tests**

```python
def test_v100_version_and_exports() -> None:
    import messy_xlsx

    assert messy_xlsx.__version__ == "1.0.0"
    assert {
        "LegacyAPIWarning", "StreamingTypeError", "BatchStream", "DataFrameChunkStream",
        "SheetStream", "SheetResult", "read_excel_arrow", "read_excel_batches",
    } <= set(messy_xlsx.__all__)
```

Add wheel/sdist smoke code that imports every new public name, reads an Arrow table, exhausts one batch stream, verifies a legacy warning under `warnings.simplefilter("error", LegacyAPIWarning)`, and checks the caller stream remains open/restored. Add documentation/changelog assertions for every item in Step 2 and for the `v1.0.0` release section.
Implement those install checks in `scripts/smoke_release_install.py
--artifact PATH`; its subprocess tests in
`tests/packaging/test_release_smoke_cli.py` prove that it creates an isolated
environment outside the repository, installs exactly the supplied wheel or
sdist with `[all]`, runs `pip check`, and exercises the same behavior for both
artifact kinds.

Run: `.venv/bin/pytest tests/test_integration.py::test_v100_version_and_exports -q`

Expected: the version/export portion passes; the new documentation/changelog
smoke assertions fail until Steps 2–3 update those files.

- [ ] **Step 2: Update all public documentation**

Document:

- legacy APIs and one-warning behavior;
- `to_arrow`, `iter_batches`, `iter_dataframe_chunks`, and `iter_sheets` examples using context managers;
- exact stream lifecycle and one-active-operation rules;
- materialized versus bounded-row backend behavior;
- the 8 MiB spool threshold and temporary-file failure/cleanup semantics;
- openpyxl shared-string/style overhead;
- `StreamingTypeError` and explicit type-hint guidance;
- formula cached values versus expressions and cell-only `FormulaConfig`;
- custom registry materialization limits;
- supported XLSX/XLSM/XLTX/XLTM/XLS/CSV/TSV/TXT formats;
- v0.10.0-to-v1.0.0 migration mappings.

Add this to `mkdocs.yml` so internal workflow files are not published:

```yaml
exclude_docs: |
  superpowers/
```

- [ ] **Step 3: Verify version and set changelog metadata**

Assert both `pyproject.toml` and `src/messy_xlsx/__init__.py` already equal
`1.0.0`; do not create a second version transition. Add
`## [1.0.0] - 2026-07-22` to `CHANGELOG.md` with compatibility, performance,
API, security, and migration notes. Keep `docs/changelog.md` including the
canonical root changelog.

- [ ] **Step 4: Run the full local release gate**

Verify `build>=1.3`, `twine>=6.2`, and the pinned native-release tools are
already present in the development configuration, then synchronize:

```bash
uv pip install --python .venv/bin/python -e ".[dev,docs,all]" -r requirements/native-release.txt
.venv/bin/python -c "import build, twine; assert tuple(map(int, build.__version__.split('.')[:2])) >= (1, 3); assert tuple(map(int, twine.__version__.split('.')[:2])) >= (6, 2)"
```

Run: `.venv/bin/ruff check src/messy_xlsx tests scripts`

Expected: no lint errors.

Run: `.venv/bin/ruff format --check src/messy_xlsx tests scripts`

Expected: no formatting differences.

Run: `.venv/bin/mypy src/messy_xlsx --ignore-missing-imports`

Expected: no typing errors.

Run: `.venv/bin/pytest tests -q --cov=messy_xlsx --cov-report=term-missing --cov-fail-under=75`

Expected: all tests pass and coverage is at least 75%.

Run: `.venv/bin/mkdocs build --strict --site-dir /tmp/messy-xlsx-v100-site`

Expected: documentation builds with no warnings or errors.

Build local fallback and native artifacts from separate clean extractions of
one new source archive:

```bash
mx_release_local_root="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-release-local-XXXXXX")"
mkdir -p "$mx_release_local_root/sdist" "$mx_release_local_root/fallback-source" "$mx_release_local_root/native-source" "$mx_release_local_root/fallback" "$mx_release_local_root/native"
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_release_local_root/sdist"
tar -xzf "$mx_release_local_root"/sdist/*.tar.gz -C "$mx_release_local_root/fallback-source"
tar -xzf "$mx_release_local_root"/sdist/*.tar.gz -C "$mx_release_local_root/native-source"
mx_release_fallback_tree="$(find "$mx_release_local_root/fallback-source" -mindepth 1 -maxdepth 1 -type d)"
mx_release_native_tree="$(find "$mx_release_local_root/native-source" -mindepth 1 -maxdepth 1 -type d)"
(
  cd "$mx_release_fallback_tree"
  MESSY_XLSX_BUILD_MODE=fallback /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_release_local_root/fallback"
)
(
  cd "$mx_release_native_tree"
  MESSY_XLSX_BUILD_MODE=native /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_release_local_root/native"
)
.venv/bin/twine check "$mx_release_local_root"/sdist/* "$mx_release_local_root"/fallback/* "$mx_release_local_root"/native/*
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_release_local_root"/native/*abi3*.whl
```

These are local diagnostics; the committed release SHA owns the authoritative
complete matrix in Step 7.

- [ ] **Step 5: Smoke-test wheel and source distribution in clean environments**

Create fresh local artifacts and smoke them through the tested CLI:

```bash
mx_release_smoke_root="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-release-smoke-XXXXXX")"
mkdir -p "$mx_release_smoke_root/dist"
MESSY_XLSX_BUILD_MODE=native .venv/bin/python -m build --sdist --wheel --outdir "$mx_release_smoke_root/dist"
mx_release_smoke_wheel="$(find "$mx_release_smoke_root/dist" -maxdepth 1 -type f -name '*.whl' -print -quit)"
mx_release_smoke_sdist="$(find "$mx_release_smoke_root/dist" -maxdepth 1 -type f -name '*.tar.gz' -print -quit)"
test -n "$mx_release_smoke_wheel"
test -n "$mx_release_smoke_sdist"
.venv/bin/pytest tests/packaging/test_release_smoke_cli.py -q
.venv/bin/python scripts/smoke_release_install.py --artifact "$mx_release_smoke_wheel"
.venv/bin/python scripts/smoke_release_install.py --artifact "$mx_release_smoke_sdist"
```

Expected: both isolated installs pass `pip check`, import every public name,
parse `tests/samples/budget_vs_actuals.xlsx`, preserve the caller stream, emit
the expected legacy warning, and pass the context-managed batch API.

- [ ] **Step 6: Commit the v1.0.0 release candidate**

```bash
git add README.md docs mkdocs.yml CHANGELOG.md pyproject.toml src/messy_xlsx/__init__.py .github/workflows tests/test_integration.py scripts/smoke_release_install.py tests/packaging/test_release_smoke_cli.py
git commit -m "release: prepare messy-xlsx v1.0.0"
```

- [ ] **Step 7: Verify the publish preconditions without tagging**

Run: `git status --short --branch`

Expected: clean branch with the release commit at `HEAD`.

Run:

```bash
gh workflow view test.yml
gh workflow view publish.yml
gh workflow view native-artifacts.yml
```

After an explicitly authorized branch push, dispatch and accept the single
complete native artifact orchestrator on this exact release SHA:

```bash
mx_release_sha="$(git rev-parse HEAD)"
mx_release_branch="$(git branch --show-current)"
git fetch origin "$mx_release_branch"
test "$(git rev-parse "origin/$mx_release_branch")" = "$mx_release_sha"
gh workflow run native-artifacts.yml --ref "$mx_release_branch"
mx_release_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-release-$mx_release_sha"
mkdir -p "$mx_release_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_release_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_release_review_dir/final-run-ledger.json"
mx_release_run="$(
  .venv/bin/python scripts/verify_native_ci.py print-run-id \
    --ledger "$mx_release_review_dir/final-run-ledger.json" \
    --workflow native-artifacts.yml
)"
mx_release_download="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-native-release-set-$mx_release_sha-XXXXXX")"
gh run download "$mx_release_run" \
  --name "final-$mx_release_sha-release-set" \
  --dir "$mx_release_download"
.venv/bin/python scripts/release_artifacts.py verify \
  --phase final \
  --revision "$mx_release_sha" \
  --dist "$mx_release_download/release-set" \
  --manifest "$mx_release_download/final-manifest.json"
.venv/bin/python scripts/check_wheel_resolution.py \
  --wheelhouse "$mx_release_download/release-set" \
  --manifest "$mx_release_download/final-manifest.json"
mx_release_performance_report="$(
  find "$mx_release_download" -type f -name 'native-csv-performance.json' -print -quit
)"
test -n "$mx_release_performance_report"
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --phase final \
  --validate-report "$mx_release_performance_report"
.venv/bin/twine check "$mx_release_download"/release-set/*
test "$(git rev-parse HEAD)" = "$mx_release_sha"
```

Expected: test/publish workflows are valid, and one source archive, seven
native ABI3 wheels, and one universal fallback wheel build from that source
archive and pass manifest, resolver, ABI, smoke, safety, performance, and
provenance verification. No earlier candidate/final manifest may be reused
after documentation/changelog changes. The publish workflow requires a `v*`
tag whose version matches package, module, changelog, and `main` tip.

Only after the branch is pushed, every required GitHub check passes, the paired benchmark artifact is accepted, and the user explicitly authorizes publication should execution create and push tag `v1.0.0`.

---

## Final Cross-Slice Verification

- [ ] Every progress row is checked and every task commit exists in order.
- [ ] `git diff v0.10.0..HEAD --check` reports no whitespace errors.
- [ ] The full compatibility corpus matches `tests/compatibility/golden/v010-frames.json`.
- [ ] Ordinary materialized XLSX parsing records zero complete openpyxl loads.
- [ ] Every selected multi-sheet parse records one manifest and one successful full materialization.
- [ ] Batch schemas remain stable and every yielded batch respects `batch_size`.
- [ ] Seekable caller streams remain open and return to their entry cursor.
- [ ] Temporary spool files are absent after success, failure, and early close.
- [ ] External OOXML relationships are never followed and hostile archives fail within limits.
- [ ] Candidate XLSX median runtime is at least 4x faster and median RSS at least 50% lower than v0.10.0 on the paired runner.
- [ ] Seekable CSV median RSS is no more than 120% of path-backed CSV RSS.
- [ ] Python 3.11, 3.12, 3.13, and 3.14 pass on Linux, macOS, and Windows.
- [ ] Wheel and source distribution smoke tests pass with optional formula and XLS dependencies.
- [ ] The package remains untagged until explicit publication authorization after CI is green.
