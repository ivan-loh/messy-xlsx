# Native CSV Tokenizer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the invalid pandas-chunk CSV/TSV/TXT full pass with a bounded native tokenizer while preserving the v0.10.0 materialized API and the approved v1.0.0 streaming contracts.

**Architecture:** A Cython `cp311-abi3` extension owns byte reads, decoding, framing, C/Python-engine structural semantics, footer retention, and deterministic native counters. Python owns source borrowing, bounded evidence replay, pandas 3.0.5 physical-value sampling, stable-schema conversion, Arrow normalization, typed backend decisions, and materialized fallback. Public native routing stays behind a source-controlled gate until semantic, safety, memory, performance, and candidate-wheel gates pass.

**Tech Stack:** CPython 3.11–3.14, Cython 3.2.9, CPython Limited API `0x030B0000`, setuptools 83.0.0, pandas 3.0.5, PyArrow, pytest 9, Hypothesis, cibuildwheel 4.1.1, abi3audit 0.0.26, ASan, and UBSan.

**Approved design:** `docs/superpowers/specs/2026-07-26-native-csv-tokenizer-design.md`

## Global Constraints

- Target release is exactly messy-xlsx `1.0.0`; the release tag remains `v1.0.0`.
- Materialized `CSVHandler` remains the pandas 3.0.5 oracle: C engine when `skip_footer == 0`, Python engine when `skip_footer > 0`.
- The only native-streaming compatibility exceptions are bounded stable-schema `StreamingTypeError` and late path-decoding failure.
- No public tokenizer or backend-selector API is added.
- Any registry subclass, custom detector, handler replacement/subclass, class mutation, or parse override remains authoritative through `CUSTOM_DATAFRAME`.
- Caller-owned streams are never closed and exact entry cursors are restored on success, failure, and early close.
- The full-pass read request is always `1..65_536` bytes; an over-return or unsupported memoryview is rejected.
- Full-pass completed rows are bounded by `batch_size + skip_footer`, plus one current record and the fixed read buffer.
- Evidence limits are exactly 1,000 requested rows, 1,000,000 examined records, 256 MiB examined payload, 16,000,000 examined cells, 8 MiB retained replay, and 1,000,000 retained cells.
- Generated multi-row headers always use materialized fallback in v1.0.0.
- Native footer execution requires evidence to reach physical EOF inside every hard budget.
- Pandas is pinned exactly to `pandas==3.0.5`.
- Build pins are exactly `Cython==3.2.9`, `setuptools==83.0.0`, `cibuildwheel==4.1.1`, and `abi3audit==0.0.26`.
- Stable ABI floor is CPython 3.11; free-threaded runtimes are excluded.
- The native release matrix contains seven native wheels, one universal fallback wheel, and one source archive.
- `_NATIVE_CSV_PRODUCTION_READY` remains `False` through candidate artifacts; its flip to `True` is the only functional source change in the enablement commit.
- `CONTINUE.md` remains untracked and is never staged.
- Do not add `uv.lock`.
- Every behavior task follows red-green-refactor, receives a fresh spec and quality review, and ends in a focused commit.

## Existing Work That Must Be Preserved

The starting worktree intentionally contains:

```text
 M src/messy_xlsx/parsing/csv_handler.py
 M src/messy_xlsx/workbook.py
 M tests/test_resource_lifecycle.py
?? CONTINUE.md
?? src/messy_xlsx/parsing/csv_io.py
?? src/messy_xlsx/parsing/csv_streaming.py
?? tests/test_csv_streaming.py
```

The existing 91 focused tests pass. Keep bounded inspection, source borrowing,
schema compilation, lifecycle fixes, registry eligibility, and compatible
tests. Replace the pandas chunk full pass. Preserve `CONTINUE.md` outside every
commit.

## File and Responsibility Map

### Runtime

- `src/messy_xlsx/_csv_tokenizer.pyx` — native evidence/full-pass state machines, source protocol, framing, engine semantics, footer, counters, and test fault seams.
- `src/messy_xlsx/parsing/csv_contracts.py` — immutable enums/configuration/evidence/read/debug/converter/execution-decision types and exact budgets.
- `src/messy_xlsx/parsing/csv_native.py` — runtime eligibility, extension handshake, production gate, evidence orchestration, warning/error translation, materialized selection, candidate smoke seam, and native reader lifecycle.
- `src/messy_xlsx/parsing/csv_value_adapter.py` — pandas evidence classification, NA/scalar conversion, and late type-error context.
- `src/messy_xlsx/parsing/csv_streaming.py` — bounded inspection plus the physical Arrow reader shell; the existing pandas full-pass reader is removed.
- `src/messy_xlsx/parsing/csv_probe.py` — the small Python logical-record helpers still required by legacy metadata inspection.
- `src/messy_xlsx/parsing/contracts.py` — workbook metrics extended with typed per-operation CSV execution observations.
- `src/messy_xlsx/parsing/csv_handler.py` — unchanged materialized oracle except for importing retained probe helpers.
- `src/messy_xlsx/workbook.py` — exact built-in/custom selection and reader construction.

### Build and release

- `setup.py`, `setup.cfg`, `MANIFEST.in`, `pyproject.toml` — conditional native/fallback setuptools builds and pinned metadata.
- `build_support.py`, `requirements/native-release.txt` — deterministic build-mode selection and exact native-release tool pins.
- `scripts/release_artifacts.py` — exact-sdist artifact recording, assembly, provenance, and verification.
- `scripts/check_wheel_resolution.py` — isolated native/fallback resolver checks for supported, unsupported, future, and free-threaded tags.
- `scripts/smoke_csv_artifact.py` — candidate private-adapter, final public-route, and fallback-wheel smoke.
- `.github/workflows/native-wheels.yml` — reusable exact-sdist seven-wheel build and ABI3 runtime-smoke matrix.
- `.github/workflows/native-safety.yml` — sanitizer, debug-allocator, allocation-fault, and lifecycle gates.
- `.github/workflows/native-artifacts.yml` — immutable candidate/final assembly and verification.
- `.github/workflows/test.yml`, `.github/workflows/publish.yml` — call the reusable matrix and publish only verified final artifacts.

### Tests and benchmarks

- `tests/native_csv/__init__.py`, `tests/native_csv/oracle.py` — exact pandas oracle helpers and warning/error capture.
- `tests/native_csv/fixtures/` — checked-in minimized byte fixtures.
- `tests/native_csv/test_abi_shell.py`
- `tests/native_csv/test_execution_decisions.py`
- `tests/native_csv/test_evidence.py`
- `tests/native_csv/test_value_adapter.py`
- `tests/native_csv/test_tokenizer_c.py`
- `tests/native_csv/test_tokenizer_python.py`
- `tests/native_csv/test_lifecycle.py`
- `tests/native_csv/test_bounds.py`
- `tests/native_csv/test_differential.py`
- `tests/native_csv/test_fuzz.py`
- `tests/packaging/` — build-mode, artifact, resolver, smoke, and publish-contract tests.
- `benchmarks/native_csv.py`, `tests/test_performance/test_native_csv_contract.py` — deterministic 300,000-row generator and comparison contract.

## Progress Tracker

- [ ] Task 1 — Quarantine and commit the existing Task 14 baseline
- [ ] Task 2 — Prove the Stable-ABI build and native safety shell
- [ ] Task 3 — Freeze the engine-specific pandas oracle
- [ ] Task 4 — Add native contracts, evidence budgets, and deterministic test doubles
- [ ] Task 5 — Implement the native source protocol and lifecycle state machine
- [ ] Task 6 — Implement valid C-mode decoding and framing
- [ ] Task 7 — Complete C-mode pandas structural parity
- [ ] Task 8 — Implement hard-budget evidence replay and routing
- [ ] Task 9 — Implement pandas evidence and physical value conversion
- [ ] Task 10 — Implement Python-engine parsing order
- [ ] Task 11 — Implement footer retention and no-lookahead bounds
- [ ] Task 12 — Integrate the native reader and existing normalization pipeline
- [ ] Task 13 — Preserve warning, error, encoding, and cleanup semantics
- [ ] Task 14 — Close deterministic bounds, fault injection, and fuzz gates
- [ ] Task 15 — Pass the authoritative performance gate
- [ ] Task 16 — Build and verify disabled candidate artifacts
- [ ] Task 17 — Enable native routing and rebuild final artifacts
- [ ] Task 18 — Remove superseded runtime code and run whole-repository acceptance

---

### Task 1: Quarantine and Commit the Existing Task 14 Baseline

**Files:**
- Create: `src/messy_xlsx/parsing/csv_contracts.py`
- Create: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/contracts.py`
- Modify: `src/messy_xlsx/parsing/csv_streaming.py`
- Modify: `src/messy_xlsx/parsing/csv_handler.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `tests/test_csv_streaming.py`
- Modify: `tests/test_resource_lifecycle.py`
- Create: `tests/native_csv/__init__.py`
- Create: `tests/native_csv/test_execution_decisions.py`
- Preserve untracked: `CONTINUE.md`

**Interfaces:**
- Produces: `CSVExecutionKind`, `CSVExecutionReason`, and `CSVExecutionDecision`.
- Produces: `ParseMetrics.record_csv_execution(kind, reason) -> CSVExecutionDecision`.
- Produces: `_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False`.
- Preserves: the current pandas-chunk implementation only as a source-checkout test scaffold; public candidate routing materializes.

- [ ] **Step 1: Add failing execution-decision and gate tests**

```python
def test_candidate_public_csv_route_is_materialized(tmp_path: Path) -> None:
    source = tmp_path / "rows.csv"
    source.write_text("a,b\n1,2\n", encoding="utf-8")
    with MessyWorkbook(source) as workbook:
        with workbook.iter_batches() as stream:
            assert pa.Table.from_batches(list(stream)).to_pydict() == {"a": [1], "b": [2]}
        assert workbook.parse_metrics.last_csv_execution == CSVExecutionDecision(
            operation_id=1,
            kind=CSVExecutionKind.MATERIALIZED_FALLBACK,
            reason=CSVExecutionReason.PRODUCTION_GATE_DISABLED,
        )


def test_custom_csv_keeps_custom_backend_and_csv_decision(custom_csv_registry) -> None:
    with MessyWorkbook(io.BytesIO(b"a\n1\n"), filename="x.csv", registry=custom_csv_registry) as wb:
        with wb.iter_batches() as stream:
            list(stream)
        assert wb.parse_metrics.last_csv_execution.kind is CSVExecutionKind.CUSTOM_SPI
        assert wb.parse_metrics.last_csv_execution.reason is CSVExecutionReason.CUSTOM_SPI
```

Add an autouse fixture in `tests/test_csv_streaming.py` that temporarily sets
`csv_native._NATIVE_CSV_PRODUCTION_READY = True` so the existing 91 prototype
tests remain characterization coverage. The candidate-public test explicitly
sets it back to `False`.

- [ ] **Step 2: Run the tests and confirm the red state**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_execution_decisions.py tests/test_csv_streaming.py -q
```

Expected: collection fails because the execution types, metric field, and gate
selection do not exist.

- [ ] **Step 3: Add the typed decision contract and false production gate**

```python
class CSVExecutionKind(StrEnum):
    NATIVE = "csv_native"
    MATERIALIZED_FALLBACK = "csv_materialized_fallback"
    CUSTOM_SPI = "custom_dataframe"


class CSVExecutionReason(StrEnum):
    NATIVE_SELECTED = "native_selected"
    CUSTOM_SPI = "custom_spi"
    PRODUCTION_GATE_DISABLED = "production_gate_disabled"
    KILL_SWITCH = "kill_switch"
    UNSUPPORTED_RUNTIME = "unsupported_runtime"
    IMPORT_OR_LOAD_FAILURE = "import_or_load_failure"
    HANDSHAKE_MISMATCH = "handshake_mismatch"
    EVIDENCE_BUDGET = "evidence_budget"
    MULTI_HEADER_EXACTNESS = "multi_header_exactness"
    UNSUPPORTED_EVIDENCE_TYPE = "unsupported_evidence_type"


@dataclass(frozen=True, slots=True)
class CSVExecutionDecision:
    operation_id: int
    kind: CSVExecutionKind
    reason: CSVExecutionReason
```

Extend `ParseMetrics` with a sequence, last decision, and typed count map. The
recording method increments exactly once:

```python
def record_csv_execution(
    self,
    kind: CSVExecutionKind,
    reason: CSVExecutionReason,
) -> CSVExecutionDecision:
    self.csv_operation_sequence += 1
    decision = CSVExecutionDecision(self.csv_operation_sequence, kind, reason)
    self.last_csv_execution = decision
    key = (kind, reason)
    self.csv_execution_counts[key] = self.csv_execution_counts.get(key, 0) + 1
    return decision
```

In `csv_native.py`, define the source-controlled constant and a pure selection
function. Only exact `"1"` disables native:

```python
_NATIVE_CSV_PRODUCTION_READY: Final = False


def capability_reason() -> CSVExecutionReason | None:
    if not _NATIVE_CSV_PRODUCTION_READY:
        return CSVExecutionReason.PRODUCTION_GATE_DISABLED
    if os.environ.get("MESSY_XLSX_DISABLE_NATIVE") == "1":
        return CSVExecutionReason.KILL_SWITCH
    return None
```

Route the candidate public path to the existing materialized streaming adapter
before the pandas-chunk reader starts. Record `CUSTOM_SPI` separately without
changing `BackendKind.CUSTOM_DATAFRAME`.

- [ ] **Step 4: Run the focused baseline and lifecycle gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_execution_decisions.py tests/test_csv_streaming.py tests/test_resource_lifecycle.py -q
.venv/bin/ruff check src/messy_xlsx/parsing/csv_contracts.py src/messy_xlsx/parsing/csv_native.py src/messy_xlsx/parsing/contracts.py src/messy_xlsx/parsing/csv_handler.py src/messy_xlsx/parsing/csv_io.py src/messy_xlsx/parsing/csv_streaming.py src/messy_xlsx/workbook.py tests/native_csv tests/test_csv_streaming.py tests/test_resource_lifecycle.py
git diff --check
```

Expected: all focused tests pass, the default public route is materialized,
prototype tests run only under the source-checkout fixture, and lint/diff
checks are clean.

- [ ] **Step 5: Commit the quarantined baseline**

```bash
git add src/messy_xlsx/parsing/csv_contracts.py src/messy_xlsx/parsing/csv_native.py src/messy_xlsx/parsing/contracts.py src/messy_xlsx/parsing/csv_handler.py src/messy_xlsx/parsing/csv_io.py src/messy_xlsx/parsing/csv_streaming.py src/messy_xlsx/workbook.py tests/native_csv tests/test_csv_streaming.py tests/test_resource_lifecycle.py
git diff --cached --name-only
git commit -m "test: quarantine CSV streaming prototype"
```

Verify `CONTINUE.md` is absent from `git diff --cached --name-only` before
committing.

---

### Task 2: Prove the Stable-ABI Build and Native Safety Shell

**Files:**
- Create: `build_support.py`
- Create: `setup.py`
- Create: `setup.cfg`
- Create: `MANIFEST.in`
- Create: `requirements/native-release.txt`
- Create: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/packaging/__init__.py`
- Create: `tests/packaging/test_build_support.py`
- Create: `tests/packaging/test_build_modes.py`
- Create: `tests/native_csv/test_abi_shell.py`
- Create: `.github/workflows/native-abi.yml`
- Modify: `pyproject.toml`
- Modify: `.gitignore`

**Interfaces:**
- Produces: native module constants `NATIVE_API_VERSION == 1` and `PANDAS_SEMANTIC_VERSION == "3.0.5"`.
- Produces: `NativeCSVTokenizer` with inert constructor, `bind`, `debug_state`, `read_batch`, and idempotent `close` surface.
- Proves: `cp311-abi3` import on CPython 3.11–3.14 before parser implementation.

- [ ] **Step 1: Write failing build-mode and ABI-shell tests**

```python
def test_native_module_handshake_and_initial_state() -> None:
    import messy_xlsx._csv_tokenizer as native

    assert native.NATIVE_API_VERSION == 1
    assert native.PANDAS_SEMANTIC_VERSION == "3.0.5"
    tokenizer = native.NativeCSVTokenizer(object())
    assert tokenizer.debug_state.state == "new"
    tokenizer.close()
    tokenizer.close()
    assert tokenizer.debug_state.state == "closed"
```

Add subprocess tests in `tests/packaging/` that build with
`MESSY_XLSX_BUILD_MODE=native` and `fallback`, assert invalid modes fail,
assert native wheel tag `cp311-abi3-*`, and assert fallback tag
`py3-none-any`.

- [ ] **Step 2: Run the red build tests**

Run:

```bash
.venv/bin/pytest tests/packaging/test_build_support.py tests/packaging/test_build_modes.py tests/native_csv/test_abi_shell.py -q
```

Expected: import/build assertions fail because no extension or setuptools build
configuration exists.

- [ ] **Step 3: Migrate the build backend and add the minimal Cython shell**

Set:

```toml
[build-system]
requires = ["setuptools==83.0.0", "wheel", "Cython==3.2.9"]
build-backend = "setuptools.build_meta"

[tool.setuptools]
package-dir = {"" = "src"}

[tool.setuptools.packages.find]
where = ["src"]
```

Pin the runtime dependency to `pandas==3.0.5`. Put all mode/platform decisions
in `build_support.resolve_build_mode()`: accept only explicit `native` or
`fallback`; otherwise default native only on supported non-free-threaded
CPython 3.11–3.14 platform/architectures and fallback everywhere else.
`setup.py` consumes that result and declares:

```python
Extension(
    "messy_xlsx._csv_tokenizer",
    ["src/messy_xlsx/_csv_tokenizer.pyx"],
    define_macros=[("Py_LIMITED_API", "0x030B0000")],
    py_limited_api=True,
)
```

Set `[bdist_wheel] py_limited_api = cp311` in `setup.cfg`. Include the `.pyx`
in `MANIFEST.in`; generate C only below `build/cython/`; do not track generated
`.c`, `.so`, or `.pyd` files.

`requirements/native-release.txt` pins:

```text
abi3audit==0.0.26
cibuildwheel==4.1.1
Cython==3.2.9
setuptools==83.0.0
```

The Cython shell must exercise an extension type, typed byte memoryview,
Python `source.read`, callback invocation, `PyMem_Malloc/Realloc/Free`,
overflow checks, terminal state, and no-throw `__dealloc__`, without parsing
CSV yet.

- [ ] **Step 4: Build, audit, and test both modes locally**

Run:

```bash
mx_abi_native="$(mktemp -d)"
mx_abi_fallback="$(mktemp -d)"
mx_abi_sdist="$(mktemp -d)"
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/packaging/test_build_support.py tests/packaging/test_build_modes.py tests/native_csv/test_abi_shell.py -q
MESSY_XLSX_BUILD_MODE=native .venv/bin/python -m build --wheel --outdir "$mx_abi_native"
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_abi_native"/*abi3*.whl
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --wheel --outdir "$mx_abi_fallback"
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_abi_sdist"
.venv/bin/python -m zipfile -l "$mx_abi_fallback"/*.whl
```

Expected: native build/test passes and audits as `cp311-abi3`; fallback wheel
contains no extension and is `py3-none-any`. Native/fallback editable modes
are also built from separate clean source extractions.

`native-abi.yml` is both dispatchable and reusable. Before Task 3 it must build
the ABI shell on the complete claimed platform matrix: manylinux and
musllinux x86-64/aarch64, macOS x86-64/arm64, and Windows AMD64. Each platform
compiles only `cp311-abi3`; clean jobs then install the exact already-built
artifact by path outside the repository on CPython 3.11, 3.12, 3.13, and 3.14.
It runs `abi3audit==0.0.26 --strict` on all seven wheels and fails on every
empty, skipped, or unsupported matrix leg.

- [ ] **Step 5: Commit the ABI proof**

```bash
git add build_support.py setup.py setup.cfg MANIFEST.in requirements/native-release.txt pyproject.toml .gitignore src/messy_xlsx/_csv_tokenizer.pyx tests/packaging tests/native_csv/test_abi_shell.py .github/workflows/native-abi.yml
git commit -m "build: prove native CSV stable ABI"
```

If `abi3audit --strict` or any runtime import fails, stop tokenizer work. Amend
the approved design and artifact matrix to per-minor wheels before continuing.
The complete remote ABI-shell matrix for this commit must succeed before Task 3
begins; a local x86-64 wheel alone is not sufficient feasibility evidence.

---

### Task 3: Freeze the Engine-Specific Pandas Oracle

**Files:**
- Create: `tests/native_csv/oracle.py`
- Create: `tests/native_csv/fixtures/*.csv.bin`
- Create: `tests/native_csv/test_oracle.py`
- Create: `tests/native_csv/test_differential.py`
- Modify: `tests/test_parsing/test_csv_handler.py`
- Modify: `tests/test_edge_cases/test_csv_variations.py`

**Interfaces:**
- Produces: `OracleResult(columns, rows, scalar_types, warnings, error)` from exact materialized `CSVHandler`.
- Produces: checked-in engine-labeled byte cases shared by all later native tasks.
- Consumes: pandas exactly `3.0.5`.

- [ ] **Step 1: Write the oracle harness and semantic fixtures**

```python
@dataclass(frozen=True, slots=True)
class OracleResult:
    columns: tuple[object, ...] | None
    rows: tuple[tuple[object, ...], ...] | None
    scalar_types: tuple[tuple[str, ...], ...] | None
    warnings: tuple[tuple[type[Warning], str], ...]
    error: tuple[type[BaseException], str, tuple[tuple[str, object], ...]] | None


def materialized_oracle(data: bytes, options: ParseOptions) -> OracleResult:
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            frame = CSVHandler().parse(io.BytesIO(data), None, options)
        except BaseException as error:
            return OracleResult(None, None, None, captured(caught), error_contract(error))
    return OracleResult(
        tuple(frame.columns),
        tuple(map(tuple, frame.itertuples(index=False, name=None))),
        tuple(tuple(type(value).__qualname__ for value in frame.iloc[:, i]) for i in range(frame.shape[1])),
        captured(caught),
        None,
    )
```

Add explicit C/Python variants for LF/CRLF/CR-only, NUL, quote junk,
unterminated quote, embedded newline, short/wide rows, implicit index, blank
rows, malformed footer rows, quote errors before footer removal, skiprows that
bisect multiline records, empty/header-only/all-footer input, fallback encoding,
warning-as-error, and `max_rows + skip_footer`.

- [ ] **Step 2: Run oracle tests against pandas 3.0.5**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_oracle.py tests/test_parsing/test_csv_handler.py tests/test_edge_cases/test_csv_variations.py -q
```

Expected: all oracle characterizations pass and record engine-specific
differences; no native comparison is enabled yet.

- [ ] **Step 3: Add generated-header and physical-scalar proofs**

Assert that:

```python
assert generated_name(b"orig\n001\n002\n003\n004\n") == "1__2"
assert generated_name(b"orig\n001\n002\n003\n004\nlate\n") == "001__002"
```

Add object arbitrary-precision integers, boolean-with-missing, string-with-NaN,
missing-promoted integer, unsigned integer, overflow, and late mixed text
cases. The generated-header cases must state `MATERIALIZED_FALLBACK`, not a
native expected result.

- [ ] **Step 4: Run and commit the frozen oracle**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_oracle.py tests/native_csv/test_differential.py tests/test_parsing/test_csv_handler.py tests/test_edge_cases/test_csv_variations.py -q
```

Then:

```bash
git add tests/native_csv tests/test_parsing/test_csv_handler.py tests/test_edge_cases/test_csv_variations.py
git commit -m "test: freeze pandas CSV engine semantics"
```

---

### Task 4: Add Native Contracts, Evidence Budgets, and Test Doubles

**Files:**
- Modify: `src/messy_xlsx/parsing/csv_contracts.py`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Create: `tests/native_csv/test_contracts.py`
- Create: `tests/native_csv/test_evidence.py`

**Interfaces:**
- Produces: every immutable type in the approved design.
- Produces: `NativeModule(Protocol)` and a deterministic Python fake used before the real tokenizer is complete.
- Produces: `native_evidence_limits(operation_max_rows) -> NativeEvidenceLimits`.

- [ ] **Step 1: Write failing configuration and budget tests**

```python
@pytest.mark.parametrize(
    ("maximum", "expected"),
    [(None, 1_000), (0, 0), (3, 3), (10_000, 1_000)],
)
def test_production_evidence_limits_are_exact(maximum, expected) -> None:
    limits = native_evidence_limits(maximum)
    assert limits.requested_data_rows == expected
    assert limits.max_records_examined == 1_000_000
    assert limits.max_payload_bytes_examined == 256 * 1024**2
    assert limits.max_cells_examined == 16_000_000
    assert limits.max_replay_bytes == 8 * 1024**2
    assert limits.max_retained_cells == 1_000_000
```

Also test frozen dataclasses, negative rejection, status/EOF consistency,
column-width consistency, `done/source_eof` consistency, and the complete
reason enum.

- [ ] **Step 2: Run the red contract tests**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_contracts.py tests/native_csv/test_evidence.py -q
```

Expected: failures identify every missing type and constructor invariant.

- [ ] **Step 3: Implement the immutable Python contracts**

Define:

```python
class NativeSemanticEngine(StrEnum): C = "c"; PYTHON = "python"
class NativeEvidenceStatus(StrEnum): COMPLETE = "complete"; SAMPLE_FULL = "sample_full"; BUDGET_EXHAUSTED = "budget_exhausted"
class PandasValueKind(StrEnum): INT64 = "int64"; UINT64 = "uint64"; FLOAT64 = "float64"; BOOL = "bool"; TEXT = "text"; OBJECT_INTEGER = "object_integer"; OBJECT_BOOLEAN = "object_boolean"; OBJECT_TEXT = "object_text"
class PandasMissingKind(StrEnum): FLOAT_NAN = "float_nan"; PANDAS_NA = "pandas_na"; NONE = "none"
```

Add `NativeEvidenceLimits`, `CSVHeaderPlan`, `NativeCSVFramingConfig`,
`NativeEvidenceReplay`, `NativeEvidence`, `PandasValueConverter`,
`ResolvedNativeCSVConfig`, `NativeCSVWarning`, `NativeDebugState`,
`NativeCSVRead`, and `NativeModule`.

The fake module must simulate `COMPLETE`, `SAMPLE_FULL`, budget exhaustion,
warnings, terminal failures, `done` without source EOF, and immutable debug
snapshots. It is test-only and never selected in production.

- [ ] **Step 4: Run focused tests and commit**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_contracts.py tests/native_csv/test_evidence.py -q
.venv/bin/ruff check src/messy_xlsx/parsing/csv_contracts.py src/messy_xlsx/parsing/csv_native.py tests/native_csv
```

Then:

```bash
git add src/messy_xlsx/parsing/csv_contracts.py src/messy_xlsx/parsing/csv_native.py tests/native_csv/test_contracts.py tests/native_csv/test_evidence.py
git commit -m "feat: define native CSV contracts"
```

---

### Task 5: Implement the Native Source Protocol and Lifecycle State Machine

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/native_csv/test_native_api.py`
- Create: `tests/native_csv/test_lifecycle.py`

**Interfaces:**
- Consumes: `ResolvedNativeCSVConfig`, `NativeCSVRead`, and `NativeDebugState`.
- Produces: one-shot `bind(source)`, non-reentrant `read_batch`, safe
  `debug_state`, test observer, and terminal/idempotent `close`.
- Preserves: GIL and ownership rules before semantic parsing exists.

- [ ] **Step 1: Write failing state/source protocol tests**

```python
@pytest.mark.parametrize("value", [b"x", bytearray(b"x"), memoryview(b"x")])
def test_source_protocol_copies_supported_binary_results(value) -> None:
    source = ScriptedSource([value, b""])
    tokenizer = tokenizer_for_no_header()
    tokenizer.bind(source)
    read = tokenizer.read_batch(1, lambda warning: None)
    assert all(1 <= size <= 65_536 for size in source.requested_sizes)
    assert read.done


def test_over_return_is_terminal() -> None:
    tokenizer = tokenizer_for_no_header()
    tokenizer.bind(OverReturningSource())
    with pytest.raises(TypeError, match="returned more than requested"):
        tokenizer.read_batch(1, lambda warning: None)
    assert tokenizer.debug_state.state == "terminal"
```

Cover non-byte, multidimensional/noncontiguous/non-byte-format memoryviews,
partial and zero-length reads, one-shot bind, recursive calls from `read`,
recursive calls from warning/test observers, `requested_rows <= 0`,
`operation_max_rows == 0` without bind/read, callback failure, close during
each state, partial construction, and no-throw finalization.

- [ ] **Step 2: Rebuild and confirm red**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
```

Expected: failures show that the ABI shell does not yet implement the exact
state/source protocol.

- [ ] **Step 3: Implement the state machine and owned input buffer**

Use explicit native states:

```text
NEW -> BOUND -> READING -> BOUND
                    \----> TERMINAL -> CLOSED
NEW/BOUND --------------------------> CLOSED
```

The native read helper must:

```cython
cdef object result = self._source.read(requested)
if not isinstance(result, (bytes, bytearray, memoryview)):
    raise TypeError(...)
if len(result) > requested:
    raise TypeError("Binary source read() returned more than requested")
```

Normalize only one-dimensional C-contiguous byte memoryviews, copy into owned
fixed storage with the GIL held, and release every allocation on all exception
paths. The observer receives immutable snapshots before/after source/warning
callbacks and before return; it never observes itself.

- [ ] **Step 4: Run lifecycle, allocator, and debug-heap tests**

Run:

```bash
PYTHONMALLOC=debug .venv/bin/pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
.venv/bin/python -X dev -m pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
```

Expected: all state/source/lifecycle tests pass without debug allocator output.

- [ ] **Step 5: Commit the lifecycle shell**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py
git commit -m "feat: add native CSV lifecycle shell"
```

---

### Task 6: Implement Valid C-Mode Decoding and Framing

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/native_csv/test_tokenizer_c.py`
- Add: `tests/native_csv/fixtures/c-valid-*.csv.bin`

**Interfaces:**
- Consumes: `NativeSemanticEngine.C` configuration.
- Produces: valid decoded `list[str | None]` fields with physical/logical line
  metadata and fixed-buffer streaming.

- [ ] **Step 1: Add failing valid-record tests**

Parameterize every fixture across source chunk sizes `1, 2, 7, 65_535,
65_536`, requested rows `1, 2, 3, 127`, and LF/CRLF/CR-only:

```python
@pytest.mark.parametrize("batch_size", [1, 2, 3, 127])
def test_valid_c_records_match_oracle(valid_c_case, batch_size) -> None:
    expected = materialized_oracle(valid_c_case.data, valid_c_case.options)
    actual = native_rows(valid_c_case.data, valid_c_case.options, batch_size)
    assert actual.rows == expected.rows
    assert actual.columns == expected.columns
```

Include quoted delimiters, doubled quotes, embedded CR/LF, missing final
terminator, arbitrary one-character delimiter, skip-initial-space, UTF-8 BOM,
UTF-16 LE/BE BOM/code-unit splits, Latin-1, decoder ignore/strict, and short-row
right padding.

- [ ] **Step 2: Rebuild and confirm the semantic red**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py -q -k valid
```

Expected: valid semantic cases fail while Task 5 lifecycle tests stay green.

- [ ] **Step 3: Implement incremental decoding and valid record states**

Implement decoder states for field-start, unquoted, quoted, post-quote, CR
pending, and EOF. Never decode unread prefetched bytes. Grow only the current
record with checked `Py_ssize_t` arithmetic. On record completion, split fields
without a second complete-record copy where possible, then transfer owned
Python strings into the returned read.

At the moment the requested accepted row is releasable, return without another
read, framing step, field split, or callback.

- [ ] **Step 4: Run valid C and lifecycle gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py -q -k valid
.venv/bin/pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
```

Expected: all valid-record/source split combinations pass.

- [ ] **Step 5: Commit valid C-mode framing**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv/test_tokenizer_c.py tests/native_csv/fixtures
git commit -m "feat: tokenize valid pandas C-mode CSV"
```

---

### Task 7: Complete C-Mode Pandas Structural Parity

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Modify: `tests/native_csv/test_tokenizer_c.py`
- Add: `tests/native_csv/fixtures/c-edge-*.csv.bin`
- Modify: `tests/test_csv_streaming.py`

**Interfaces:**
- Produces: exact C-mode NUL, quote-junk, width, blank, implicit-index, warning,
  and accepted-row-limit behavior.
- Preserves: behavior independent of public batch boundaries.

- [ ] **Step 1: Add failing C-edge oracle comparisons**

```python
@pytest.mark.parametrize("batch_size", [1, 2, 3, 127])
def test_later_wide_row_is_batch_boundary_independent(batch_size) -> None:
    data = b"a,b\n1,2\n3,4,5\n6,7\n"
    assert native_outcome(data, batch_size) == materialized_oracle(data, ParseOptions())
```

Add first-row implicit indexes with one and multiple leading fields, blanks
before width inference, NUL truncation/delimiter classification, quote junk,
quotes in unquoted fields, unterminated quoted EOF, later excess fields,
consecutive/all bad rows, exact one-based warning lines/messages, and
`max_rows` counting only accepted rows.

- [ ] **Step 2: Run the red C-edge matrix**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py tests/test_csv_streaming.py -q -k "c_engine or implicit or nul or malformed or max_rows"
```

Expected: at least the known pandas-chunk boundary case and implicit-index
cases fail.

- [ ] **Step 3: Implement the C compatibility transitions**

Establish physical width/index fields from the first pandas-eligible row.
Right-pad short rows. Emit synchronous `NativeCSVWarning` and discard later
wide rows. Implement pandas C NUL and post-quote transitions from the frozen
fixture corpus rather than Python `csv` behavior.

Keep `operation_max_rows` distinct from `source_eof`; reaching the limit marks
the final nonempty read `done=True` and performs no lookahead.

- [ ] **Step 4: Run complete C differential tests**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py tests/native_csv/test_differential.py tests/test_csv_streaming.py -q -k "not python_mode and not footer"
```

Expected: schema-compatible C cases match values, columns, warnings, errors,
and physical line context at all batch sizes.

- [ ] **Step 5: Commit C structural parity**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv tests/test_csv_streaming.py
git commit -m "feat: match pandas C CSV edge semantics"
```

---

### Task 8: Implement Hard-Budget Evidence Replay and Routing

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/csv_contracts.py`
- Modify: `tests/native_csv/test_evidence.py`
- Create: `tests/native_csv/test_selection.py`

**Interfaces:**
- Produces: native structural `_scan_evidence(source, framing, limits)`.
- Produces: Python `scan_evidence(source, framing, limits) -> NativeEvidence`.
- Produces: exact pre-full-pass fallback decisions.

- [ ] **Step 1: Add failing evidence budget/replay tests**

```python
def test_payload_budget_stops_inside_unterminated_record() -> None:
    limits = replace(native_evidence_limits(None), max_payload_bytes_examined=16)
    evidence = scan_structural(b'a\n"' + b"x" * 1_000_000, limits)
    assert evidence.status is NativeEvidenceStatus.BUDGET_EXHAUSTED
    assert evidence.payload_bytes_examined == 16
    assert evidence.replay_bytes_retained <= limits.max_replay_bytes
    assert not evidence.eof
```

Cover all five budgets independently; fixed-buffer prefetch excluded from
examined bytes; original quoting/terminators preserved in replay; no
reserialization; `COMPLETE`/`SAMPLE_FULL`/`BUDGET_EXHAUSTED`; zero target;
width too large for one row; empty/header-only/all-bad input; fresh borrow
after evidence; multi-header immediate fallback; unsupported physical evidence
fallback; and no warning emission during evidence.

- [ ] **Step 2: Run the red evidence suite**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_evidence.py tests/native_csv/test_selection.py -q
```

Expected: missing structural evidence and routing behavior fail.

- [ ] **Step 3: Implement incremental budget checks and exact replay**

Before examining each next source byte/cell/record or retaining each replay
byte/cell, check its independent counter. Budget exhaustion may discard an
incomplete current record immediately and never searches for its terminator.

For C/no-footer mode, return `SAMPLE_FULL` when target rows are fixed. For
Python/footer mode, do not return `SAMPLE_FULL`; continue to EOF or a budget.
Evidence cleanup restores the borrow before returning a materialized decision.
No native-to-materialized transition is permitted after full-pass binding.

- [ ] **Step 4: Run evidence, cursor, and source tests**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_evidence.py tests/native_csv/test_selection.py tests/test_source_handle.py tests/test_csv_streaming.py -q -k "sample or evidence or cursor or replay or multi_header"
```

Expected: all evidence states, decisions, and cursor restores are exact.

- [ ] **Step 5: Commit bounded evidence**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx src/messy_xlsx/parsing/csv_contracts.py src/messy_xlsx/parsing/csv_native.py tests/native_csv tests/test_source_handle.py tests/test_csv_streaming.py
git commit -m "feat: add bounded native CSV evidence"
```

---

### Task 9: Implement Pandas Evidence and Physical Value Conversion

**Files:**
- Create: `src/messy_xlsx/parsing/csv_value_adapter.py`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Create: `tests/native_csv/test_pandas_evidence.py`
- Create: `tests/native_csv/test_value_adapter.py`

**Interfaces:**
- Produces: `compile_pandas_evidence(replay, framing, options) -> PandasEvidence`.
- Produces: `PandasCSVValueAdapter(converters, columns).convert(read, row_offset)`.
- Selects: `UNSUPPORTED_EVIDENCE_TYPE` before a public reader for unsupported
  extension or heterogeneous object evidence.

- [ ] **Step 1: Add failing converter-classification tests**

```python
@pytest.mark.parametrize(
    ("csv", "kind", "missing"),
    [
        (b"v\n1\n2\n", PandasValueKind.INT64, PandasMissingKind.NONE),
        (b"v,x\n1,1\n,2\n", PandasValueKind.FLOAT64, PandasMissingKind.FLOAT_NAN),
        (b"v,x\nTrue,1\n,2\nFalse,3\n", PandasValueKind.OBJECT_BOOLEAN, PandasMissingKind.FLOAT_NAN),
        (b"v\n18446744073709551616\n", PandasValueKind.OBJECT_INTEGER, PandasMissingKind.NONE),
    ],
)
def test_pandas_evidence_classifies_exact_scalar_family(csv, kind, missing) -> None:
    evidence = compile_evidence(csv)
    assert evidence.converters[0] == PandasValueConverter(kind, missing)
```

Add string dtype with float NaN, uint64, object text, default/configured NA
markers, whitespace, quoted/unquoted empties, structural missing fields,
footer exclusion, heterogeneous-object fallback, unsupported extension dtype,
and original replay quoting.

- [ ] **Step 2: Run converter tests and confirm red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_pandas_evidence.py tests/native_csv/test_value_adapter.py -q
```

Expected: failures show the converter authority and conversion code are absent.

- [ ] **Step 3: Compile exact converters from pandas-typed replay**

Parse the bounded original-byte replay with pandas 3.0.5 and exact handler
kwargs while capturing rather than emitting warnings. Pair pandas typed values
with native raw lexemes and classify nonmissing Python scalar families.

Convert later lexemes only according to the compiled descriptor. On the first
incompatible non-null lexeme raise:

```python
raise StreamingTypeError(
    "streamed value is incompatible with the fixed schema",
    ordinal=ordinal,
    display_label=physical_label_description(column),
    row_offset=row_offset + relative,
    value_description=physical_value_description(raw),
    expected_type=converter.kind.value,
)
```

Do not reproduce pandas whole-file widening or `DtypeWarning`.

- [ ] **Step 4: Run raw-value and normalization compatibility gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_pandas_evidence.py tests/native_csv/test_value_adapter.py tests/test_streaming_normalization.py tests/test_arrow_api.py -q -k "csv or physical or normalize_false or late"
```

Expected: exact supported physical scalars and stable-schema failures pass.

- [ ] **Step 5: Commit the value adapter**

```bash
git add src/messy_xlsx/parsing/csv_value_adapter.py src/messy_xlsx/parsing/csv_native.py tests/native_csv/test_pandas_evidence.py tests/native_csv/test_value_adapter.py
git commit -m "feat: convert native CSV physical values"
```

---

### Task 10: Implement Python-Engine Parsing Order

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/native_csv/test_tokenizer_python.py`
- Add: `tests/native_csv/fixtures/python-edge-*.csv.bin`

**Interfaces:**
- Consumes: `NativeSemanticEngine.PYTHON`.
- Produces: Python-engine record parsing, physical-line skiprows, parser-error
  diagnostics, width/index establishment, and blank classification before
  footer retention.

- [ ] **Step 1: Add failing Python-engine tests without footer assertions**

```python
def test_python_skiprows_can_bisect_multiline_record() -> None:
    data = b'a,b\n"first\nsecond",x\n3,4\n'
    options = ParseOptions(skip_rows=2, skip_footer=1)
    assert native_structural_outcome(data, options) == materialized_oracle(data, options)
```

Cover NUL, quote-junk, quotes in unquoted fields, unterminated quote, CR-only,
blank rows, `csv.Error` recovery, physical-line skiprows, implicit-index width
fixed before later post-processing, and malformed warning metadata. Use a
nonzero footer to select Python semantics but isolate pre-footer transitions.

- [ ] **Step 2: Rebuild and confirm the Python semantic red**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_tokenizer_python.py -q -k "not footer"
```

Expected: failures identify C/Python transition differences and physical-line
skip behavior.

- [ ] **Step 3: Add the Python semantic state branch**

Implement Python-specific quote errors and line accounting from the oracle
fixtures. A `csv.Error` record emits one synchronous parser diagnostic and is
discarded before footer processing. Establish physical width and any implicit
index on the first pandas-eligible data row before post-parse skip/header
stages.

- [ ] **Step 4: Run Python pre-footer and C regression gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_tokenizer_python.py -q -k "not footer"
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py -q
```

Expected: Python pre-footer and all existing C cases pass.

- [ ] **Step 5: Commit Python parsing order**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv/test_tokenizer_python.py tests/native_csv/fixtures
git commit -m "feat: match pandas Python CSV parsing order"
```

---

### Task 11: Implement Footer Retention and No-Lookahead Bounds

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Modify: `tests/native_csv/test_tokenizer_python.py`
- Create: `tests/native_csv/test_footer.py`
- Modify: `tests/test_csv_streaming.py`

**Interfaces:**
- Produces: successful parsed-row footer deque bounded by `skip_footer`.
- Produces: accepted output bounded by `requested_rows`, with no work after the
  requested row becomes releasable.

- [ ] **Step 1: Add failing footer-order and bound tests**

```python
def test_quote_error_is_discarded_before_footer_removal() -> None:
    data = b"a,b\n1,2\n3,4\n\"bad\n5,6\n"
    options = ParseOptions(skip_footer=1)
    assert native_outcome(data, options, batch_size=1) == materialized_oracle(data, options)


def test_wide_trailing_row_can_disappear_as_footer_without_warning() -> None:
    data = b"a,b\n1,2\n3,4,5\n"
    outcome = native_outcome(data, ParseOptions(skip_footer=1), batch_size=1)
    assert outcome.rows == ((1, 2),)
    assert outcome.warnings == ()
```

Cover blank/wide rows occupying footer slots, consecutive parser errors,
footer zero/one/all/greater-than-rows, multiline records, header-none
all-footer, batch sizes `1,2,3,127`, and exact successor-row counters.

- [ ] **Step 2: Run and confirm the footer red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_footer.py tests/native_csv/test_tokenizer_python.py -q -k footer
```

Expected: footer ordering/bounds fail.

- [ ] **Step 3: Implement parsed-success footer ownership**

Retain at most `skip_footer` successfully parsed rows. Parser-error records
never occupy a slot; blank or over-wide successfully parsed records do until
their later classification stage. Release an output row only after the
required successor count is known. Stop all source reads, framing, field
tokenization, and callbacks immediately after the requested output becomes
releasable.

- [ ] **Step 4: Run Python/footer differential and architecture tests**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_footer.py tests/native_csv/test_tokenizer_python.py tests/native_csv/test_differential.py tests/test_csv_streaming.py -q -k "footer or python_mode"
```

Expected: Python/footer cases match the materialized oracle and every retained
row counter stays within `batch_size + skip_footer`.

- [ ] **Step 5: Commit footer semantics**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv/test_footer.py tests/native_csv/test_tokenizer_python.py tests/test_csv_streaming.py
git commit -m "feat: add bounded Python-mode CSV footers"
```

---

### Task 12: Integrate the Native Reader and Existing Normalization Pipeline

**Files:**
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/csv_streaming.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `src/messy_xlsx/parsing/contracts.py`
- Create: `tests/native_csv/test_integration.py`
- Modify: `tests/test_csv_streaming.py`
- Modify: `tests/test_arrow_api.py`

**Interfaces:**
- Preserves: `prepare_csv_streaming_reader(source, plan, metrics, *, construction_owner=None) -> PreparedStreamingReader`.
- Produces: internal `NativeCSVReader(StreamingBatchReader)`.
- Removes: `_PreownedPandasReader` and pandas `chunksize` from the full pass.
- Retains: current bounded inspection, `PreparedStreamingReader`,
  `_CloseOnceReader`, normalization compilation, physical encoding, and public
  display-name handling.

- [ ] **Step 1: Add failing end-to-end integration tests**

```python
@pytest.mark.parametrize("batch_size", [1, 2, 3, 127])
def test_private_candidate_native_stream_matches_oracle(csv_case, batch_size) -> None:
    with candidate_native_stream(csv_case.data, csv_case.options, batch_size) as stream:
        table = pa.Table.from_batches(list(stream))
    assert table_rows(table) == materialized_oracle(csv_case.data, csv_case.options).rows
```

Cover `normalize=True/False`, all-null columns, duplicate/non-string labels,
dataframe chunks and global `RangeIndex`, path/seekable/nonseekable/one-byte
sources, final nonempty `done`, `max_rows == 0`, early close, stable schema
before return, and records larger than 8 MiB after the evidence sample.

- [ ] **Step 2: Run the integration red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_integration.py tests/test_csv_streaming.py tests/test_arrow_api.py -q -k "csv"
```

Expected: private native integration fails because the full pass still uses
pandas chunks.

- [ ] **Step 3: Replace only the full-pass reader**

`prepare_csv_streaming_reader` performs inspection, native evidence, pandas
converter compilation, normalization-plan compilation, and routing while no
full-pass borrow is open. It returns an inert `NativeCSVReader`; first
`read_next_batch()` opens `SourceHandle.open_binary()`, binds the tokenizer,
and converts `NativeCSVRead` columns through `PandasCSVValueAdapter` and the
existing normalization wrapper.

If capability/evidence selection chooses materialized fallback, restore the
evidence borrow first and construct the existing materialized streaming
adapter. Never catch a native execution failure as fallback.

- [ ] **Step 4: Run public, Arrow, and custom-registry gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_integration.py tests/test_csv_streaming.py tests/test_arrow_api.py tests/test_reader_routing.py tests/test_architecture_contracts.py -q
```

Expected: private native integration passes; public default remains
`PRODUCTION_GATE_DISABLED`; custom registry behavior remains materialized and
authoritative.

- [ ] **Step 5: Commit native integration**

```bash
git add src/messy_xlsx/parsing/csv_native.py src/messy_xlsx/parsing/csv_streaming.py src/messy_xlsx/parsing/contracts.py src/messy_xlsx/workbook.py tests/native_csv/test_integration.py tests/test_csv_streaming.py tests/test_arrow_api.py
git commit -m "feat: integrate native CSV streaming"
```

---

### Task 13: Preserve Warning, Error, Encoding, and Cleanup Semantics

**Files:**
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/csv_value_adapter.py`
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/native_csv/test_failures.py`
- Modify: `tests/native_csv/test_lifecycle.py`
- Modify: `tests/test_source_handle.py`
- Modify: `tests/test_resource_lifecycle.py`
- Modify: `tests/test_stream_lifecycle.py`

**Interfaces:**
- Produces: exact eager evidence and lazy full-pass `FormatError` boundaries.
- Preserves: process-failure and Task 13 cleanup precedence.
- Produces: warning emission once in full pass, never from evidence.

- [ ] **Step 1: Add failing warning/error/lifecycle tests**

```python
def test_warning_promoted_to_error_uses_materialized_format_boundary() -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("error", pd.errors.ParserWarning)
        with pytest.raises(FormatError) as caught:
            exhaust_native(b"a,b\n1,2,3\n")
    assert isinstance(caught.value.__cause__, pd.errors.ParserWarning)


def test_fallback_encoding_footer_evidence_uses_legacy_terminal_context(tmp_path: Path) -> None:
    source = write_late_invalid_utf8_with_python_parser_error(tmp_path)
    with pytest.raises(FormatError, match="Cannot read CSV with any encoding") as caught:
        open_native(source, ParseOptions(skip_footer=1))
    assert caught.value.context["attempted_formats"] == [
        "csv[latin-1]", "csv[windows-1252]", "csv[iso-8859-1]"
    ]
```

Cover evidence warnings suppressed/full warnings once, warning callback
failure, eager/lazy decode/parser/I/O errors, late strict path decode,
fallback-evidence replay failure, ordinary and process failures during read and
cleanup, cursor restoration failure preventing final batch return, return-gap
interruption, active-operation release, finalizer cleanup, and caller stream
never closed.

- [ ] **Step 2: Run the failure suite and confirm red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_failures.py tests/native_csv/test_lifecycle.py tests/test_source_handle.py tests/test_resource_lifecycle.py tests/test_stream_lifecycle.py -q
```

Expected: missing native error mapping/cleanup cases fail.

- [ ] **Step 3: Implement exact adapter translation and cleanup order**

Ordinary native/parser/decoder/source failures become contextual `FormatError`.
After fallback encoding is selected, evidence, pandas replay, and full-pass
parser failures use the legacy `"Cannot read CSV with any encoding"` message
and attempted-format context. `MemoryError`, `KeyboardInterrupt`,
`SystemExit`, or chained process failures propagate unchanged.

Close tokenizer allocations before restoring the source borrow. A process-level
cleanup failure replaces an ordinary primary; ordinary cleanup never replaces
an active primary; an active process failure remains authoritative.

- [ ] **Step 4: Run lifecycle and compatibility gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_failures.py tests/native_csv/test_lifecycle.py tests/test_source_handle.py tests/test_resource_lifecycle.py tests/test_stream_lifecycle.py tests/test_task13_stream_lifecycle.py tests/compatibility -q
```

Expected: all failure chains, lifecycle obligations, and compatibility
fixtures pass.

- [ ] **Step 5: Commit error/lifecycle semantics**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx src/messy_xlsx/parsing/csv_native.py src/messy_xlsx/parsing/csv_value_adapter.py tests/native_csv/test_failures.py tests/native_csv/test_lifecycle.py tests/test_source_handle.py tests/test_resource_lifecycle.py tests/test_stream_lifecycle.py
git commit -m "fix: preserve native CSV failure semantics"
```

---

### Task 14: Close Deterministic Bounds, Fault Injection, and Fuzz Gates

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/native_csv/test_bounds.py`
- Create: `tests/native_csv/test_failure_injection.py`
- Create: `tests/native_csv/test_reentrancy.py`
- Create: `tests/native_csv/test_fuzz.py`
- Create: `tests/native_csv/fuzz_worker.py`
- Add: `tests/native_csv/regressions/*.bin`
- Create: `scripts/run_native_csv_sanitizers.sh`
- Create: `scripts/run_native_csv_fuzz.py`
- Modify: `.github/workflows/native-abi.yml`

**Interfaces:**
- Produces: observer events `before_source_read`, `after_source_read`,
  `before_warning`, `after_warning`, and `before_return`.
- Produces: `_allocation_sites_for_tests()` and
  `_set_allocation_failure_for_tests(site)`.
- Proves: native safety and literal memory/row bounds before performance work.

- [ ] **Step 1: Add failing deterministic counter and allocation tests**

```python
def assert_snapshot(snapshot: NativeDebugState, batch_size: int, footer: int) -> None:
    assert snapshot.output_rows_retained <= batch_size
    assert snapshot.post_output_rows_retained <= footer
    assert snapshot.field_tokenized_successor_rows <= footer
    assert snapshot.undecoded_buffer_bytes <= 65_536
    assert snapshot.current_record_payload_bytes >= 0
    assert snapshot.footer_payload_bytes >= 0
    assert snapshot.output_payload_bytes >= 0


@pytest.mark.parametrize("site", native_allocation_sites())
def test_every_native_allocation_failure_is_safe(site) -> None:
    set_allocation_failure(site)
    tokenizer = tokenizer_for_fixture()
    with pytest.raises(MemoryError):
        tokenizer.read_batch(1, lambda warning: None)
    tokenizer.close()
```

Assert no source read/tokenization/callback after a requested batch becomes
releasable, oversized full-pass records remain valid, overflow/realloc paths
are checked, recursive mutating calls fail, `debug_state` is allowed
reentrantly, and observer failure is terminal.

- [ ] **Step 2: Run the red safety suites**

Run:

```bash
PYTHONMALLOC=debug .venv/bin/pytest tests/native_csv/test_bounds.py tests/native_csv/test_failure_injection.py tests/native_csv/test_reentrancy.py -q
```

Expected: missing observers/fault sites or counter violations fail.

- [ ] **Step 3: Implement fault seams and close all deterministic bounds**

Give each `PyMem_*` allocation site a stable test name. Reallocation uses a
temporary pointer; checked arithmetic precedes every allocation; each owner has
one unwind path. Allocation hooks are internal and default disabled.
`native_allocation_bytes` is cross-checked with hooks but remains report-only.

- [ ] **Step 4: Add and run seeded differential fuzzing**

Generate valid/malformed bytes and randomized source chunk splits. Run 5,000
fixed-seed examples for each engine:

```bash
.venv/bin/python scripts/run_native_csv_fuzz.py --c-seed 0x0C5A14 --python-seed 0xBADC5EED --examples 5000 --timeout 300
```

Schema-compatible cases must match pandas values/scalars/columns/warnings/error
classes. Late incompatible and late decode cases assert their separate
streaming exceptions. Minimize every mismatch into
`tests/native_csv/regressions/` before fixing it.

- [ ] **Step 5: Run sanitizers and commit safety gates**

Run:

```bash
bash scripts/run_native_csv_sanitizers.sh
PYTHONMALLOC=debug .venv/bin/pytest tests/native_csv -q
```

The sanitizer script builds with:

```text
-O1 -g -fno-omit-frame-pointer -fsanitize=address,undefined -Werror
ASAN_OPTIONS=detect_leaks=0:abort_on_error=1
UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1
```

Then:

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv scripts/run_native_csv_sanitizers.sh scripts/run_native_csv_fuzz.py .github/workflows/native-abi.yml
git commit -m "test: fuzz and sanitize native CSV tokenizer"
```

---

### Task 15: Pass the Authoritative Performance Gate

**Files:**
- Create: `benchmarks/native_csv.py`
- Create: `scripts/run_native_csv_benchmarks.py`
- Create: `tests/native_csv/reference_streaming.py`
- Create: `tests/test_performance/test_native_csv_contract.py`
- Create: `.github/workflows/native-performance.yml`
- Modify: `.github/workflows/test.yml`

**Interfaces:**
- Produces: deterministic benchmark corpora from seed `0x0C5A14`.
- Produces: machine-readable per-run and aggregate benchmark reports.
- Compares: native no-footer routing with direct pandas C-engine parsing.
- Compares: native footer routing with both the retained Python streaming
  reference and materialized `CSVHandler`.
- Gates: the production constant remains `False` if any timing or deterministic
  ownership threshold fails.

- [ ] **Step 1: Write failing benchmark-contract tests**

```python
def test_native_csv_corpus_contract() -> None:
    corpora = build_native_csv_corpora(rows=300_000, seed=0x0C5A14)
    assert {
        "clean_unquoted_lf",
        "quoted_lf",
        "crlf",
        "multiline_quoted",
        "sparse_malformed",
        "batch_size_one",
        "large_logical_record",
        "footer_ten",
    } == set(corpora)
    for corpus in corpora.values():
        assert corpus.logical_records < 310_000
        assert corpus.examined_payload_bytes < 48 * 1024**2
        assert corpus.examined_cells < 4_800_000
        assert corpus.replay_payload_bytes < 2 * 1024**2
        assert corpus.replay_cells < 100_000
```

Add tests that reject a report unless it contains exactly three warmups,
seven alternating measured runs per contender, individual medians, geometric
means, execution decisions, Python/pandas/native/compiler/platform/CPU/build
metadata, throughput, process and Python peak memory, deterministic native
counters, source position at each batch, and a stable output hash.

The large logical record must appear after the first 1,000 accepted data rows.
The footer corpus must use `skip_footer=10`, reach physical EOF within every
evidence limit, and record `CSVExecutionKind.NATIVE`.

- [ ] **Step 2: Run the contract tests and confirm the red state**

Run:

```bash
.venv/bin/pytest tests/test_performance/test_native_csv_contract.py -q
```

Expected: collection or contract assertions fail because the deterministic
harness, reference implementation, and report validator do not exist.

- [ ] **Step 3: Implement deterministic generation, alternating runs, and validation**

`benchmarks/native_csv.py` generates each corpus without timing file creation.
`scripts/run_native_csv_benchmarks.py` installs or accepts the exact native
wheel path, validates native routing before starting a timer, performs three
warmups, then alternates contenders for seven measured runs. It reports
medians and computes geometric means from per-corpus ratios.

Copy only the superseded Python framing/filter/footer implementation needed by
the footer baseline into `tests/native_csv/reference_streaming.py`. It is a
benchmark/differential reference, never imported by installed runtime code.

Reject a performance report unless all of these are true:

```text
each no-footer native/direct-pandas-C median ratio <= 3.0
geometric mean of no-footer native/direct-pandas-C ratios <= 2.0
geometric mean of footer reference/native median ratios >= 4.0
every public case records CSVExecutionKind.NATIVE
every deterministic row, fixed-buffer, and logical-payload bound passes
```

Direct pandas C is not an equivalent footer baseline. Footer reports include
the retained Python reference and end-to-end materialized `CSVHandler`.

- [ ] **Step 4: Add the authoritative and corroborating CI jobs**

`.github/workflows/native-performance.yml` builds or downloads the exact
manylinux native artifact and runs the gate on the dedicated Ubuntu 24.04
x86-64 benchmark runner, its pinned image/CPU identity, and CPython 3.12.
Other supported native platforms produce corroborating reports without
relaxing or replacing the authoritative thresholds. Store the JSON report,
corpora manifest, wheel hash, and runner identity as SHA-scoped artifacts.

There is no automatic waiver. A threshold miss leaves
`_NATIVE_CSV_PRODUCTION_READY = False` until the design and this plan receive
explicit user-approved amendments.

- [ ] **Step 5: Run the local contract and representative benchmark**

Run:

```bash
.venv/bin/pytest tests/test_performance/test_native_csv_contract.py -q
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --rows 300000 \
  --seed 0x0C5A14 \
  --warmups 3 \
  --runs 7 \
  --output /tmp/messy-xlsx-native-csv-performance.json
```

Expected: the report validator passes. Local timing is diagnostic; only the
dedicated CI runner decides the timing gate.

- [ ] **Step 6: Commit the performance gate**

```bash
git add benchmarks/native_csv.py scripts/run_native_csv_benchmarks.py tests/native_csv/reference_streaming.py tests/test_performance/test_native_csv_contract.py .github/workflows/native-performance.yml .github/workflows/test.yml
git commit -m "perf: gate native CSV production routing"
```

---

### Task 16: Build and Verify Disabled Candidate Artifacts

**Files:**
- Modify: `build_support.py`
- Modify: `requirements/native-release.txt`
- Create: `scripts/release_artifacts.py`
- Create: `scripts/check_wheel_resolution.py`
- Create: `scripts/smoke_csv_artifact.py`
- Modify: `tests/packaging/test_build_support.py`
- Modify: `tests/packaging/test_build_modes.py`
- Create: `tests/packaging/test_artifact_smoke_cli.py`
- Create: `tests/packaging/test_release_artifacts.py`
- Create: `tests/packaging/test_wheel_resolution.py`
- Create: `tests/packaging/test_publish_contract.py`
- Create: `.github/workflows/native-wheels.yml`
- Create: `.github/workflows/native-safety.yml`
- Create: `.github/workflows/native-artifacts.yml`
- Modify: `setup.py`
- Modify: `pyproject.toml`
- Modify: `.github/workflows/native-abi.yml`
- Modify: `.github/workflows/test.yml`
- Modify: `.github/workflows/publish.yml`
- Modify: `.gitignore`
- Modify: `Makefile`

**Interfaces:**
- Produces: exact native/fallback build modes from one source archive.
- Produces: `release_artifacts.py record|assemble|verify`.
- Produces: a private module-token-protected
  `csv_native._run_candidate_artifact_smoke(...)` that bypasses only the false
  production constant.
- Proves: the public candidate route remains materialized with reason
  `PRODUCTION_GATE_DISABLED`.
- Produces: exactly seven `cp311-abi3` native wheels, one `py3-none-any`
  fallback wheel, one source archive, and a SHA-256 manifest outside the
  publish directory.

- [ ] **Step 1: Write failing build, candidate-seam, artifact, and resolver tests**

Test `build_support.resolve_build_mode()` with explicit `native`, explicit
`fallback`, invalid values, supported CPython, unsupported/free-threaded
runtimes, and supported/unsupported architectures. Test clean native and
fallback editable installations from separate source extractions so a stale
extension cannot satisfy fallback assertions.

The candidate-seam tests must prove:

```python
assert csv_native._NATIVE_CSV_PRODUCTION_READY is False
assert public_csv_decision().reason is CSVExecutionReason.PRODUCTION_GATE_DISABLED
with pytest.raises(PermissionError):
    csv_native._run_candidate_artifact_smoke(object(), source, plan)
```

Only the module-owned token may use the seam. It bypasses only the production
constant: supported-runtime, kill-switch, import, handshake, evidence, and
semantic guards still execute. No environment variable or public API exposes
the seam.

Create corrupt synthetic artifacts and manifests that must be rejected for
missing/extra/duplicate files, wrong tags or purity, extension absence or
presence in the wrong variant, cross-wheel `METADATA` drift, missing `.pyx` or
build files in the source archive, forbidden generated native files,
`CONTINUE.md`, `.superpowers`, or `uv.lock`, wrong source-archive lineage,
phase/namespace mismatch, or altered SHA-256 content.

- [ ] **Step 2: Run the tests and confirm the red state**

Run:

```bash
.venv/bin/pytest \
  tests/packaging/test_build_support.py \
  tests/packaging/test_build_modes.py \
  tests/packaging/test_artifact_smoke_cli.py \
  tests/packaging/test_release_artifacts.py \
  tests/packaging/test_wheel_resolution.py \
  tests/packaging/test_publish_contract.py -q
```

Expected: missing helper, CLI, gate, artifact, and publish contracts fail.

- [ ] **Step 3: Finish deterministic dual-mode build support**

`requirements/native-release.txt` contains exactly:

```text
abi3audit==0.0.26
cibuildwheel==4.1.1
Cython==3.2.9
setuptools==83.0.0
```

`setup.py` delegates all platform/default logic to `build_support.py`,
generates C only below `build/cython/`, and fails closed in explicit native
mode. Official workflows always set `MESSY_XLSX_BUILD_MODE`. Native builds use
`Py_LIMITED_API=0x030B0000`; fallback builds contain no extension or generated
native artifact and report `Root-Is-Purelib: true`.

Add `build>=1.3`, `twine>=6.2`, `bandit>=1.9`, and the pinned native-release
tools needed by the documented local gate to the development configuration
without weakening the build-system pins.

- [ ] **Step 4: Implement artifact/provenance and isolated resolver tools**

`scripts/release_artifacts.py` records source and wheel hashes, assembles an
allowlisted release set, and verifies phase, revision, nine-file count,
filenames, wheel tags, purity, extension inventory, cross-wheel `METADATA`,
source inventory, exact source-archive lineage, SHA-256 content,
`abi3audit --strict`, `twine check`, and clean-environment `pip check`.

`scripts/check_wheel_resolution.py` uses temporary environments outside the
repository with `--no-index --find-links` and `--no-deps`. It proves native
preference on supported tags, universal fallback on unsupported and
free-threaded tags, and native ABI3 preference on future supported CPython
tags while the runtime guard still controls execution.

- [ ] **Step 5: Build the exact seven-wheel and ABI-smoke matrix**

`.github/workflows/native-wheels.yml` first creates one fallback-mode source
archive. Every wheel is built from a clean extraction of that exact archive.
Linux passes `MESSY_XLSX_BUILD_MODE=native` through `CIBW_ENVIRONMENT`; host
environment inheritance is not assumed. `CIBW_BUILD=cp311-*` compiles once,
and separate jobs install each exact wheel path outside the repository on
CPython 3.12, 3.13, and 3.14 in addition to cibuildwheel's 3.11 test.

The required native matrix is:

| Family | Runner and architecture | Required platform tag |
|---|---|---|
| manylinux x86-64 | `ubuntu-24.04`, `x86_64`, `manylinux2014` pinned by cibuildwheel 4.1.1 | `manylinux_2_17_x86_64.manylinux2014_x86_64` |
| musllinux x86-64 | `ubuntu-24.04`, `x86_64`, `musllinux_1_2` pinned by cibuildwheel 4.1.1 | `musllinux_1_2_x86_64` |
| manylinux aarch64 | `ubuntu-24.04-arm`, `aarch64`, `manylinux2014` pinned by cibuildwheel 4.1.1 | `manylinux_2_17_aarch64.manylinux2014_aarch64` |
| musllinux aarch64 | `ubuntu-24.04-arm`, `aarch64`, `musllinux_1_2` pinned by cibuildwheel 4.1.1 | `musllinux_1_2_aarch64` |
| macOS x86-64 | `macos-15-intel`, deployment target 10.13 | `macosx_10_13_x86_64` |
| macOS arm64 | `macos-15`, deployment target 11.0 | `macosx_11_0_arm64` |
| Windows x86-64 | `windows-2025`, `CIBW_ARCHS_WINDOWS=AMD64` | `win_amd64` |

All seven filenames contain `cp311-abi3`. Do not use `allow-empty`, floating
Linux images, `macos-latest`, Windows `auto`, or an unreviewed `test-skip`.
Run `abi3audit==0.0.26 --strict` on every native wheel.

- [ ] **Step 6: Wire immutable candidate assembly, safety, and publishing**

`.github/workflows/native-safety.yml` runs GCC ASan/UBSan with `-Werror`,
`PYTHONMALLOC=debug`, allocation failure at every named `PyMem_*` site,
callback/source reentrancy, repeated partial construction, lifecycle stress,
and no-throw deallocation.

`.github/workflows/native-artifacts.yml` derives `candidate` versus `final`
from the source-controlled constant; workflow inputs cannot request `final`.
Candidate artifact namespaces are immutable and SHA-scoped:

```text
candidate-${GITHUB_SHA}-sdist
candidate-${GITHUB_SHA}-linux-x86_64
candidate-${GITHUB_SHA}-linux-aarch64
candidate-${GITHUB_SHA}-macos-x86_64
candidate-${GITHUB_SHA}-macos-arm64
candidate-${GITHUB_SHA}-windows-amd64
candidate-${GITHUB_SHA}-fallback
candidate-${GITHUB_SHA}-release-set
```

The aggregate set has exactly nine distributions; its manifest is stored
outside the publish directory. Candidate smoke installs exact wheel paths,
uses only the private seam for native parsing, asserts public materialized
routing, and asserts fallback extension absence.

Update `publish.yml` while the gate is still false. It must require a true
source gate, invoke the complete quality/artifact workflow on the tagged SHA,
download only `final-${GITHUB_SHA}-release-set`, repeat phase/revision/hash/
lineage/metadata/Twine/ABI/pip checks, publish only `.whl` and `.tar.gz`, and
retain tag/version/changelog/main-tip validation.

- [ ] **Step 7: Run local packaging verification**

Run in fresh temporary output/extraction directories:

```bash
mx_pack_root="$(mktemp -d)"
mkdir -p "$mx_pack_root/sdist" "$mx_pack_root/source" "$mx_pack_root/fallback" "$mx_pack_root/native"
.venv/bin/pytest tests/packaging -q
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_pack_root/sdist"
tar -xzf "$mx_pack_root"/sdist/*.tar.gz -C "$mx_pack_root/source"
mx_sdist_tree="$(find "$mx_pack_root/source" -mindepth 1 -maxdepth 1 -type d)"
(
  cd "$mx_sdist_tree"
  MESSY_XLSX_BUILD_MODE=fallback /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_pack_root/fallback"
  MESSY_XLSX_BUILD_MODE=native /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_pack_root/native"
)
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_pack_root/native"/*abi3*.whl
.venv/bin/python scripts/smoke_csv_artifact.py --phase candidate --wheel "$mx_pack_root/native"/*abi3*.whl
.venv/bin/python scripts/smoke_csv_artifact.py --phase fallback --wheel "$mx_pack_root/fallback"/*.whl
```

Expected: both exact-sdist build modes, artifact unit tests, native audit, source
inventory, and candidate public/private routing assertions pass. The complete
nine-file `assemble` and `verify` commands run only after the seven platform
jobs have downloaded their exact artifacts into the candidate aggregation job.

- [ ] **Step 8: Commit and obtain candidate workflow acceptance**

```bash
git add build_support.py requirements/native-release.txt setup.py pyproject.toml .gitignore Makefile scripts/release_artifacts.py scripts/check_wheel_resolution.py scripts/smoke_csv_artifact.py tests/packaging .github/workflows/native-abi.yml .github/workflows/native-wheels.yml .github/workflows/native-safety.yml .github/workflows/native-artifacts.yml .github/workflows/test.yml .github/workflows/publish.yml
git commit -m "ci: gate disabled native candidate artifacts"
```

Do not stage `CONTINUE.md`, `.superpowers`, generated C/native files, or
`uv.lock`. Production enablement is blocked until this exact commit is pushed
with user authorization and its full candidate, safety, performance, ABI,
resolver, and artifact workflow succeeds. If any claimed ABI combination
fails, stop and amend the approved design to per-minor wheels before proceeding.

---

### Task 17: Enable Native Routing and Rebuild Final Artifacts

**Files:**
- Modify one functional line: `src/messy_xlsx/parsing/csv_native.py`
- Consume without modification: `.github/workflows/native-artifacts.yml`
- Consume without modification: `.github/workflows/native-wheels.yml`
- Consume without modification: `.github/workflows/native-safety.yml`
- Consume without modification: `.github/workflows/native-performance.yml`

**Interfaces:**
- Changes: `_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False` to `True`.
- Produces: a completely new SHA-scoped final source archive, seven native
  wheels, one fallback wheel, manifest, and verified release set.
- Proves: public built-in native routing, fallback materialization, the runtime
  kill switch, resolver behavior, ABI3 loading, safety, and performance.

- [ ] **Step 1: Verify the exact disabled candidate before editing**

Run:

```bash
git status --short --branch
git show HEAD:src/messy_xlsx/parsing/csv_native.py | rg '_NATIVE_CSV_PRODUCTION_READY.*False'
gh run list --workflow native-artifacts.yml --branch perf/parser-v1 --limit 5
gh run list --workflow native-safety.yml --branch perf/parser-v1 --limit 5
gh run list --workflow native-performance.yml --branch perf/parser-v1 --limit 5
```

Expected: the worktree is clean except intentionally untracked `CONTINUE.md`,
the constant is false, and the candidate, safety, ABI, resolver, artifact, and
authoritative performance gates for `HEAD` all succeeded.

- [ ] **Step 2: Make and inspect the one-line functional change**

Change exactly:

```python
_NATIVE_CSV_PRODUCTION_READY: Final[bool] = True
```

Run:

```bash
git diff -- src/messy_xlsx/parsing/csv_native.py
```

Expected: the diff contains exactly the false-to-true line change. Do not
combine fixes, workflow changes, version changes, or documentation changes
with this commit.

- [ ] **Step 3: Run local public-route, kill-switch, and fallback tests**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_execution_decisions.py tests/native_csv/test_integration.py tests/packaging/test_artifact_smoke_cli.py tests/packaging/test_wheel_resolution.py -q
MESSY_XLSX_DISABLE_NATIVE=1 .venv/bin/pytest tests/native_csv/test_execution_decisions.py -q -k kill_switch
```

Expected: built-in supported routing is native, the kill switch wins, custom
handlers remain authoritative, and unsupported/fallback cases materialize.

- [ ] **Step 4: Commit only the production gate**

```bash
git add src/messy_xlsx/parsing/csv_native.py
git diff --cached --stat
git commit -m "feat: enable production native CSV routing"
```

Expected: the commit changes one functional source line.

- [ ] **Step 5: Build and accept a fresh final matrix**

After pushing with user authorization, the workflow derives `final` from the
source. It must build every wheel from the new final source archive and store
only `final-${GITHUB_SHA}-*` artifacts. Candidate files or manifests are never
merged or reused.

Every exact wheel repeats direct-extension smoke, public native routing on
supported runtimes, fallback behavior, kill-switch behavior, CPython
3.11–3.14 ABI smoke, `abi3audit --strict`, resolver checks, sanitizer/fault
gates, and performance evidence.

If final CI fails, revert the gate to false, fix and verify under a new
disabled candidate revision, then make a fresh one-line enablement commit.
Never patch an enabled revision in place.

---

### Task 18: Remove Superseded Runtime Code and Run Whole-Repository Acceptance

**Files:**
- Create: `src/messy_xlsx/parsing/csv_probe.py`
- Modify: `src/messy_xlsx/parsing/csv_handler.py`
- Modify: `src/messy_xlsx/parsing/csv_streaming.py`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Delete when no production consumer remains: `src/messy_xlsx/parsing/csv_io.py`
- Modify: `tests/native_csv/reference_streaming.py`
- Modify: `tests/test_architecture_contracts.py`
- Modify: `docs/superpowers/plans/2026-07-22-parser-performance-v1.md`
- Modify: `.github/workflows/test.yml`

**Interfaces:**
- Removes: pandas `chunksize` full-pass execution and duplicated installed
  Python framing/filter/footer execution.
- Retains: only small logical-record probe helpers needed by legacy metadata
  inspection in `csv_probe.py`.
- Preserves: the benchmark/differential Python reference under tests.
- Reconciles: parent Task 14 and parent Task 20 with native artifact ownership.

- [ ] **Step 1: Add failing architecture and installed-inventory tests**

```python
def test_installed_csv_streaming_has_one_full_pass_implementation() -> None:
    assert not runtime_imports("messy_xlsx.parsing.csv_io")
    assert not source_contains("src/messy_xlsx/parsing/csv_streaming.py", "chunksize=")
    assert source_contains("src/messy_xlsx/parsing/csv_native.py", "NativeCSVReader")


def test_test_reference_is_not_imported_by_runtime() -> None:
    assert not any_runtime_imports("tests.native_csv.reference_streaming")
```

Add an sdist/wheel inventory assertion that the test reference is absent from
wheels, required `.pyx`/build/probe files are present where appropriate, and
no generated C/native artifact, `CONTINUE.md`, `.superpowers`, or `uv.lock`
leaks into source distributions.

- [ ] **Step 2: Run the architecture tests and confirm the red state**

Run:

```bash
.venv/bin/pytest tests/test_architecture_contracts.py tests/packaging/test_release_artifacts.py -q
```

Expected: the superseded runtime module or pandas full-pass implementation is
still discoverable.

- [ ] **Step 3: Move only retained probe helpers and delete superseded runtime code**

Use `rg` to enumerate every `csv_io` and pandas-chunk consumer. Move the
minimal inspection-only logical-record helpers into `csv_probe.py`, update
`CSVHandler` and inspection imports, and keep the historical streaming
reference only in `tests/native_csv/reference_streaming.py`. Delete
`csv_io.py` only after `rg` and architecture tests prove no production
consumer remains.

Do not alter tokenizer semantics during this cleanup. If cleanup exposes a
behavior difference, add a failing oracle/regression test and fix it under a
new disabled-candidate cycle rather than weakening an assertion.

- [ ] **Step 4: Reconcile the parent plan and release-version boundary**

In `docs/superpowers/plans/2026-07-22-parser-performance-v1.md`:

- mark Task 14 complete only after all native tasks and the final artifact
  matrix pass;
- replace its obsolete pandas-chunk and single-artifact directions with links
  to this completed plan;
- keep the package/module version at `0.10.0` until parent Task 20;
- make Task 20 set `pyproject.toml` and `src/messy_xlsx/__init__.py` to
  `1.0.0`, update changelog/docs, and invoke the complete final artifact matrix
  on that exact release SHA;
- state that the earlier final manifest cannot be reused after the version
  change because filenames, metadata, and hashes change;
- retain explicit user authorization as a prerequisite for creating or
  pushing `v1.0.0`.

- [ ] **Step 5: Run the complete local acceptance gate**

Run:

```bash
mx_accept_root="$(mktemp -d)"
mkdir -p "$mx_accept_root/sdist" "$mx_accept_root/source" "$mx_accept_root/fallback" "$mx_accept_root/native"
.venv/bin/ruff check src/messy_xlsx tests scripts benchmarks
.venv/bin/ruff format --check src/messy_xlsx tests scripts benchmarks
.venv/bin/mypy src/messy_xlsx --ignore-missing-imports
.venv/bin/bandit -q -r src/messy_xlsx
.venv/bin/pytest tests -q --cov=messy_xlsx --cov-report=term-missing --cov-fail-under=75
.venv/bin/mkdocs build --strict --site-dir /tmp/messy-xlsx-native-site
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_accept_root/sdist"
tar -xzf "$mx_accept_root"/sdist/*.tar.gz -C "$mx_accept_root/source"
mx_accept_tree="$(find "$mx_accept_root/source" -mindepth 1 -maxdepth 1 -type d)"
(
  cd "$mx_accept_tree"
  MESSY_XLSX_BUILD_MODE=fallback /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_accept_root/fallback"
  MESSY_XLSX_BUILD_MODE=native /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_accept_root/native"
)
.venv/bin/twine check "$mx_accept_root"/sdist/* "$mx_accept_root"/fallback/* "$mx_accept_root"/native/*
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_accept_root"/native/*abi3*.whl
git diff --check
```

Also rerun the fixed-seed 5,000-case differential fuzz suites, ASan/UBSan,
debug-allocator lifecycle suite, deterministic performance-contract tests,
native/fallback clean-install smoke, and artifact/resolver verification.

Expected: all checks pass with no production pandas-chunk reader, no duplicate
Python full pass, no public API change, and no ownership/memory regression.

- [ ] **Step 6: Obtain independent compatibility, safety, and release-readiness reviews**

Give each reviewer the approved design, this plan, the complete diff, focused
and full test output, sanitizer/fuzz reports, performance JSON, candidate/final
manifests, and CI URLs. Resolve every blocker with a regression test and repeat
the affected gate. Do not mark Task 14 complete on reviewer promises or
partial CI.

- [ ] **Step 7: Commit cleanup and parent-plan reconciliation**

```bash
git add src/messy_xlsx/parsing/csv_probe.py src/messy_xlsx/parsing/csv_handler.py src/messy_xlsx/parsing/csv_streaming.py src/messy_xlsx/parsing/csv_native.py tests/native_csv/reference_streaming.py tests/test_architecture_contracts.py .github/workflows/test.yml docs/superpowers/plans/2026-07-22-parser-performance-v1.md
git add -u src/messy_xlsx/parsing/csv_io.py
git diff --cached --name-only
git commit -m "refactor: retire superseded CSV streaming runtime"
```

Verify that `CONTINUE.md` is still untracked and absent from the staged set.
Do not tag or publish in this task.

---

## Cross-Task Verification and Review Discipline

- [ ] Before each behavior change, demonstrate the named focused test failing
  for the intended reason; after implementation, run the focused green test
  and the adjacent compatibility/lifecycle suite.
- [ ] After every task, produce a review package containing the design section,
  task diff, base/head commits, red/green commands, and results. Run a
  requirements/specification review first and a code-quality review second.
  Resolve all blockers and repeat the affected review before continuing.
- [ ] Keep `_NATIVE_CSV_PRODUCTION_READY = False` through semantic, lifecycle,
  bounds, fuzz, sanitizer, performance, ABI, resolver, and disabled-candidate
  acceptance. The enablement commit changes only that functional line.
- [ ] Record a typed CSV execution decision exactly once per operation after
  successful reader construction. Evidence failure before construction records
  no decision; custom handlers continue to record `CUSTOM_DATAFRAME`.
- [ ] Confirm evidence suppresses warnings, full execution emits each warning
  once, footer evidence reaches physical EOF or routes to materialized
  fallback, generated multi-row headers materialize, and evidence state is
  never reused as full-pass parser state.
- [ ] Confirm replay retains original bytes, the full pass borrows with
  `SourceHandle.open_binary()`, and caller streams are restored at every
  success, failure, early-close, and return-gap boundary.
- [ ] Run `git status --short` before every commit. Preserve unrelated user
  changes, keep `CONTINUE.md` untracked, and never add `.superpowers`,
  generated native files, or `uv.lock`.
- [ ] Run focused tests after every task; run the whole test, Ruff, formatting,
  mypy, Bandit, docs, build, Twine, native ABI, artifact, fuzz, sanitizer, and
  performance gates before declaring the implementation complete.
- [ ] Update this tracker and the parent Task 14 ledger only from verified
  evidence. Parent Task 20 remains responsible for v1.0.0 documentation,
  changelog, version metadata, release-SHA artifact rebuild, and the
  separately authorized tag/publication action.
