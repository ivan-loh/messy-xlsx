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
- [ ] Task 8 — Implement C-mode structural evidence and hard budgets
- [ ] Task 9 — Implement C-mode pandas evidence and physical value conversion
- [ ] Task 10 — Implement Python-engine parsing order
- [ ] Task 11 — Implement footer retention and complete Python-mode evidence
- [ ] Task 12 — Implement exact capability selection, the candidate seam, and native integration
- [ ] Task 13 — Preserve warning, error, encoding, and cleanup semantics
- [ ] Task 14 — Close deterministic bounds and allocation-fault gates
- [ ] Task 15 — Pass differential fuzz and sanitizer gates
- [ ] Task 16 — Remove superseded runtime code
- [ ] Task 17 — Set v1.0.0 metadata and verify the parent handoff
- [ ] Task 18 — Pass the authoritative performance gate
- [ ] Task 19 — Add exact-sdist artifact and provenance tooling
- [ ] Task 20 — Add resolver, candidate-smoke, and publish contracts
- [ ] Task 21 — Wire the seven-wheel and aggregate workflows
- [ ] Task 22 — Build and accept disabled candidate artifacts
- [ ] Task 23 — Enable native routing and rebuild final artifacts
- [ ] Task 24 — Run whole-repository and exact-SHA acceptance

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

Import new modules inside test functions so a missing scaffold is reported as
a normal failing test rather than a collection error. Start with the gate and
decision assertions below; after they fail, add signature-only modules whose
methods raise `NotImplementedError`, then add the behavioral assertions.

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

Expected: pytest collects successfully and the in-test imports or assertions
fail because the execution types, metric field, and gate selection do not
exist.

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
_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False


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
- Create: `tests/packaging/test_ci_run_verifier.py`
- Create: `tests/native_csv/conftest.py`
- Create: `tests/native_csv/test_abi_shell.py`
- Create: `scripts/verify_native_ci.py`
- Create: `scripts/run_native_csv_sanitizers.sh`
- Create: `.github/workflows/native-abi.yml`
- Modify: `pyproject.toml`
- Modify: `.gitignore`

**Interfaces:**
- Produces: native module constants `NATIVE_API_VERSION == 1`,
  `PANDAS_SEMANTIC_VERSION == "3.0.5"`, and a build-time
  `NATIVE_SOURCE_SHA256` matching the checked-in `.pyx`.
- Produces: `NativeCSVTokenizer` with inert constructor, `bind`, `debug_state`, `read_batch`, and idempotent `close` surface.
- Proves: `cp311-abi3` import on CPython 3.11–3.14 before parser implementation.

- [ ] **Step 1: Write failing build-mode and ABI-shell tests**

```python
def test_native_module_handshake_and_initial_state() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")

    assert native.NATIVE_API_VERSION == 1
    assert native.PANDAS_SEMANTIC_VERSION == "3.0.5"
    assert native.NATIVE_SOURCE_SHA256 == sha256(PYX_PATH.read_bytes()).hexdigest()
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
Add recorded-run tests for
`verify_native_ci.py collect --revision --workflow --output` and
`print-run-id`; wrong SHA, nonterminal/nonsuccess conclusions, and missing
required matrix jobs must fail.
Add `test_native_abi_workflow_is_dispatchable_and_reusable`, which parses
`native-abi.yml` and requires `push` on feature branches plus
`workflow_dispatch` and `workflow_call`. The push trigger is the bootstrap:
GitHub cannot manually dispatch a new workflow until that workflow exists on
the default branch.

- [ ] **Step 2: Run the red build tests**

Run:

```bash
uv pip install --python .venv/bin/python "build>=1.3" wheel
.venv/bin/pytest tests/packaging/test_build_support.py tests/packaging/test_build_modes.py tests/packaging/test_ci_run_verifier.py tests/native_csv/test_abi_shell.py -q
```

Expected: pytest collects successfully and the in-test import/build assertions
fail because no extension or setuptools build configuration exists.

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

Pin the runtime dependency to `pandas==3.0.5`. Add `build>=1.3`,
`twine>=6.2`, and `bandit>=1.9` to the `dev` extra in this task, before their
first command use. Put all mode/platform decisions
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
CSV yet. `setup.py` hashes the `.pyx`, passes that value through Cython's
compile-time environment, and the extension exposes it as
`NATIVE_SOURCE_SHA256`. The autouse native-test fixture recomputes the hash and
fails immediately if a stale installed extension is loaded. Put that fixture
in `tests/native_csv/conftest.py` so every later native suite inherits it.

Create the shell version of `scripts/run_native_csv_sanitizers.sh`. It builds
the ABI shell in a fresh extraction with ASan/UBSan and runs only
`test_abi_shell.py`, lifecycle construction/close loops, and the one initial
allocation/reallocation fault path. Task 15 expands the same script to the full
semantic suite and test-owned allocation manifest.

Implement `verify_native_ci.py collect` with a two-hour deadline and a
30-second poll interval. It waits for the requested exact-SHA dispatch to
appear and reach a terminal state, then accepts only `success` with the full
test-owned ABI job matrix. Recorded-fixture tests inject the clock and `gh`
responses, so they cover not-yet-visible, queued, success, terminal failure,
timeout, wrong-SHA, and incomplete-matrix states without sleeping.

- [ ] **Step 4: Build, audit, and test both modes locally**

Run:

```bash
mx_abi_native="$(mktemp -d)"
mx_abi_fallback="$(mktemp -d)"
mx_abi_sdist="$(mktemp -d)"
mx_abi_native_source="$(mktemp -d)"
mx_abi_fallback_source="$(mktemp -d)"
mx_abi_native_venv="$(mktemp -d)"
mx_abi_fallback_venv="$(mktemp -d)"
uv pip install --python .venv/bin/python -e ".[dev]" -r requirements/native-release.txt
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/packaging/test_build_support.py tests/packaging/test_build_modes.py tests/packaging/test_ci_run_verifier.py tests/native_csv/test_abi_shell.py -q
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_abi_sdist"
tar -xzf "$mx_abi_sdist"/*.tar.gz -C "$mx_abi_native_source"
tar -xzf "$mx_abi_sdist"/*.tar.gz -C "$mx_abi_fallback_source"
mx_abi_native_tree="$(find "$mx_abi_native_source" -mindepth 1 -maxdepth 1 -type d)"
mx_abi_fallback_tree="$(find "$mx_abi_fallback_source" -mindepth 1 -maxdepth 1 -type d)"
.venv/bin/python -m venv "$mx_abi_native_venv/venv"
.venv/bin/python -m venv "$mx_abi_fallback_venv/venv"
MESSY_XLSX_BUILD_MODE=native uv pip install \
  --python "$mx_abi_native_venv/venv/bin/python" \
  -e "$mx_abi_native_tree[all]"
MESSY_XLSX_BUILD_MODE=fallback uv pip install \
  --python "$mx_abi_fallback_venv/venv/bin/python" \
  -e "$mx_abi_fallback_tree[all]"
"$mx_abi_native_venv/venv/bin/python" -c "import messy_xlsx._csv_tokenizer as native; assert native.NATIVE_API_VERSION == 1"
"$mx_abi_fallback_venv/venv/bin/python" -c "import importlib.util; assert importlib.util.find_spec('messy_xlsx._csv_tokenizer') is None"
"$mx_abi_fallback_venv/venv/bin/python" -c "import io; from messy_xlsx import MessyWorkbook; from messy_xlsx.parsing.csv_contracts import CSVExecutionKind; source=io.BytesIO(b'a\\n1\\n'); workbook=MessyWorkbook(source, filename='x.csv'); stream=workbook.iter_batches(); list(stream); stream.close(); assert workbook.parse_metrics.last_csv_execution.kind is CSVExecutionKind.MATERIALIZED_FALLBACK; workbook.close()"
(
  cd "$mx_abi_native_tree"
  MESSY_XLSX_BUILD_MODE=native /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_abi_native"
)
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_abi_native"/*abi3*.whl
(
  cd "$mx_abi_fallback_tree"
  MESSY_XLSX_BUILD_MODE=fallback /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_abi_fallback"
)
.venv/bin/python -c "from pathlib import Path; from zipfile import ZipFile; wheel=next(Path('$mx_abi_fallback').glob('*.whl')); names=ZipFile(wheel).namelist(); assert not any('_csv_tokenizer' in name or name.endswith(('.so', '.pyd')) for name in names)"
bash scripts/run_native_csv_sanitizers.sh --shell-only
```

Expected: native build/test passes and audits as `cp311-abi3`; fallback wheel
contains no extension and is `py3-none-any`. Native/fallback editable modes
are built/tested from separate clean source extractions, the fallback wheel has
no `.so`/`.pyd`, and the shell sanitizer/debug smoke passes.

`native-abi.yml` is both dispatchable and reusable. Before Task 3 it must build
the ABI shell on the complete claimed platform matrix: manylinux and
musllinux x86-64/aarch64, macOS x86-64/arm64, and Windows AMD64. Each platform
compiles only `cp311-abi3`; clean jobs then install the exact already-built
artifact by path outside the repository on CPython 3.11, 3.12, 3.13, and 3.14.
It runs `abi3audit==0.0.26 --strict` on all seven wheels and fails on every
empty, skipped, or unsupported matrix leg. Its initial `push` trigger runs the
workflow from the exact feature commit that first introduces it; Task 21 later
removes that bootstrap trigger when `native-artifacts.yml` becomes the sole
push orchestrator.

- [ ] **Step 5: Commit the ABI proof**

```bash
git add build_support.py setup.py setup.cfg MANIFEST.in requirements/native-release.txt pyproject.toml .gitignore src/messy_xlsx/_csv_tokenizer.pyx scripts/verify_native_ci.py scripts/run_native_csv_sanitizers.sh tests/packaging tests/native_csv/conftest.py tests/native_csv/test_abi_shell.py .github/workflows/native-abi.yml
git commit -m "build: prove native CSV stable ABI"
```

If `abi3audit --strict` or any runtime import fails, stop tokenizer work. Amend
the approved design and artifact matrix to per-minor wheels before continuing.

- [ ] **Step 6: Accept the exact-SHA remote ABI proof**

Obtain explicit user authorization before pushing. After the authorized push:

```bash
mx_abi_sha="$(git rev-parse HEAD)"
mx_abi_branch="$(git branch --show-current)"
git fetch origin "$mx_abi_branch"
test "$(git rev-parse "origin/$mx_abi_branch")" = "$mx_abi_sha"
mx_abi_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_abi_sha"
mkdir -p "$mx_abi_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_abi_sha" \
  --workflow native-abi.yml \
  --output "$mx_abi_review_dir/native-abi-run-ledger.json"
```

The authorized push itself creates the exact-SHA bootstrap run; do not call
`gh workflow run` for a workflow that is not yet on the default branch.
`collect` polls for that push-triggered run and then verifies it.
Expected: the exact commit has a completed successful workflow with every
seven-wheel/CPython 3.11–3.14 shell-smoke leg and strict ABI audit. The
complete remote ABI-shell matrix must succeed before Task 3 begins; a local
x86-64 wheel alone is insufficient.

---

### Task 3: Freeze the Engine-Specific Pandas Oracle

**Files:**
- Create: `tests/native_csv/oracle.py`
- Create: `tests/native_csv/fixtures/*.csv.bin`
- Create: `tests/native_csv/test_oracle.py`
- Create: `tests/native_csv/test_differential.py`
- Create: `tests/native_csv/test_fuzz_contract.py`
- Create: `tests/native_csv/fuzz_worker.py`
- Create: `scripts/run_native_csv_fuzz.py`
- Modify: `tests/test_parsing/test_csv_handler.py`
- Modify: `tests/test_edge_cases/test_csv_variations.py`

**Interfaces:**
- Produces: `materialized_oracle(source_case, options) -> OracleResult` from
  exact materialized `CSVHandler`, preserving path/stream source behavior.
- Produces: `assert_oracle_equivalent(actual, expected)` with missing-aware
  comparisons for `NaN`, `pd.NA`, `None`, scalar families, warnings, errors,
  and error context.
- Produces: `structural_oracle(data, options, engine)` for decoded/raw
  tokenizer-stage expectations without pandas physical conversion.
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
    source_kind: OracleSourceKind
    entry_cursor: int | None
    exit_cursor: int | None
    caller_closed: bool | None
    requested_read_sizes: tuple[int, ...]


def materialized_oracle(
    source_case: OracleSourceCase | bytes, options: ParseOptions
) -> OracleResult:
    source_case = OracleSourceCase.seekable(source_case) if isinstance(source_case, bytes) else source_case
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            with source_case.open() as opened:
                frame = CSVHandler().parse(opened.argument, None, options)
        except BaseException as error:
            return oracle_failure(source_case, caught, error)
    return OracleResult(
        tuple(frame.columns),
        tuple(map(tuple, frame.itertuples(index=False, name=None))),
        tuple(tuple(type(value).__qualname__ for value in frame.iloc[:, i]) for i in range(frame.shape[1])),
        captured(caught),
        None,
        *source_case.lifecycle_contract(),
    )
```

`OracleSourceCase` has explicit `PATH`, `SEEKABLE_STREAM`,
`NONSEEKABLE_STREAM`, and `ONE_BYTE_STREAM` variants. Path cases exercise
strict decoding and fallback-encoding retries. Caller-stream variants exercise
ignore/no-retry behavior and assert exact entry-cursor restoration without
closing the caller.
`open()` yields `OpenedOracleSource(argument)` where `argument` is the actual
`Path` for `PATH` and the actual instrumented binary stream for all stream
variants; the second `CSVHandler.parse` argument is always `None` because CSV
has no sheet selector. `lifecycle_contract()` records source kind, exact
entry/exit cursor, caller `closed` state, and requested read sizes on both
success and failure.

Add explicit C/Python variants for LF/CRLF/CR-only, NUL, quote junk,
unterminated quote, embedded newline, short/wide rows, implicit index, blank
rows, malformed footer rows, quote errors before footer removal, skiprows that
bisect multiline records, empty/header-only/all-footer input, fallback encoding,
warning-as-error, and `max_rows + skip_footer`.

`assert_oracle_equivalent` never uses tuple/list equality for values that can
contain missing sentinels. It compares each cell through a missing-kind
classifier, then separately compares nonmissing values, Python scalar family,
columns, warning category/message/order, error class/message/context, and
source kind/cursor/closed/read-size lifecycle.

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
native expected result. Freeze raw structural expectations beside typed oracle
expectations so Tasks 6–8 can test decoded fields without depending on the
Task 9 value adapter.

- [ ] **Step 4: Freeze the deterministic fuzz generator/worker contract**

Before native semantics exist, implement `fuzz_worker.py` generation and
materialized-oracle modes plus
`run_native_csv_fuzz.py --oracle-only`. Contract tests fix both seeds, case
ordering, byte/source-chunk generation, per-case timeout handling, result JSON
schema, and minimized-regression filename hashing. Native C/Python execution
adapters remain explicit `NotImplementedError` branches activated by Tasks 7
and 11.

Run:

```bash
.venv/bin/pytest tests/native_csv/test_fuzz_contract.py -q
.venv/bin/python scripts/run_native_csv_fuzz.py \
  --oracle-only \
  --c-seed 0x0C5A14 \
  --python-seed 0xBADC5EED \
  --examples 100 \
  --timeout 30
```

Expected: deterministic reruns produce identical case/result hashes for both
materialized engines.

- [ ] **Step 5: Run and commit the frozen oracle**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_oracle.py tests/native_csv/test_differential.py tests/native_csv/test_fuzz_contract.py tests/test_parsing/test_csv_handler.py tests/test_edge_cases/test_csv_variations.py -q
```

Then:

```bash
git add tests/native_csv scripts/run_native_csv_fuzz.py tests/test_parsing/test_csv_handler.py tests/test_edge_cases/test_csv_variations.py
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
- Produces: staging-only `NativeStructuralEvidence`, containing no pandas-typed
  rows and never crossing the public reader boundary.
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
`NativeEvidenceReplay`, `NativeStructuralEvidence`, `NativeEvidence`,
`CompiledNativeEvidence`, `PandasValueConverter`,
`ResolvedNativeCSVConfig`, `NativeCSVWarning`, `NativeDebugState`,
`NativeCSVRead`, and `NativeModule`.

The staging type is exact:

```python
@dataclass(frozen=True, slots=True)
class NativeStructuralEvidence:
    status: NativeEvidenceStatus
    pandas_replay: NativeEvidenceReplay
    header_fields: tuple[str | None, ...]
    raw_data_rows: tuple[tuple[str | None, ...], ...]
    physical_lines: tuple[int, ...]
    leading_index_fields: int
    parser_diagnostics: tuple[NativeCSVWarning, ...]
    target_data_rows: int
    records_examined: int
    payload_bytes_examined: int
    cells_examined: int
    replay_bytes_retained: int
    cells_retained: int
    eof: bool
```

It owns original replay bytes and decoded structural rows only. Task 9 or
Task 11 combines it with pandas 3.0.5 and returns the approved
`NativeEvidence`, adding `typed_data_rows` and final hashable
`column_names`. No selector treats `NativeStructuralEvidence` as final schema
evidence.

Pandas compilation has one non-exceptional result type:

```python
@dataclass(frozen=True, slots=True)
class CompiledNativeEvidence:
    evidence: NativeEvidence | None
    config: ResolvedNativeCSVConfig | None
    fallback_reason: CSVExecutionReason | None
```

Construction enforces either both `evidence` and `config` with no fallback
reason, or one fallback reason with both success fields absent. Unsupported
extension dtypes or heterogeneous object evidence return the corresponding
fallback reason; parser, decoder, source, and process failures remain
exceptions.

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

Execute these microcycles in order; do not add the next cycle's tests until the
current one is green and reviewed:

| Cycle | Red test nodes | Minimal implementation | Rebuild/green command |
|---|---|---|---|
| A — states | `test_new_bind_read_close_transitions`, `test_bind_is_one_shot`, `test_nonpositive_request_rejected` | native state enum, constructor, `bind`, request validation, idempotent `close` only | native reinstall; `pytest tests/native_csv/test_native_api.py -q -k "transitions or one_shot or nonpositive"` |
| B — source values | `test_source_protocol_copies_supported_binary_results`, `test_rejects_invalid_memoryview`, `test_over_return_is_terminal`, `test_partial_and_zero_reads` | bounded `read(1..65_536)`, byte/memoryview validation, owned copy, terminal failure | native reinstall; `pytest tests/native_csv/test_native_api.py -q -k "source_protocol or memoryview or over_return or partial"` |
| C — lifecycle/reentrancy | `test_recursive_read_rejected`, `test_observer_debug_snapshot_allowed`, `test_callback_failure_terminal`, `test_partial_construction_and_finalizer` | non-reentrant mutation guard, immutable debug snapshots, one cleanup path, no-throw deallocator | native reinstall; `PYTHONMALLOC=debug pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q` |

After each green command, run the Task 5 specification review on only that
cycle's diff before starting the next cycle.

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
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
PYTHONMALLOC=debug .venv/bin/pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
.venv/bin/python -X dev -m pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
```

Expected: the autouse source-hash fixture proves the loaded extension matches
the current `.pyx`, and all state/source/lifecycle tests pass without debug
allocator output.

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
  metadata and fixed-buffer streaming; it does not produce pandas scalars.

- [ ] **Step 1: Add failing valid-record tests**

Parameterize every fixture across source chunk sizes `1, 2, 7, 65_535,
65_536`, requested rows `1, 2, 3, 127`, and LF/CRLF/CR-only:

```python
@pytest.mark.parametrize("batch_size", [1, 2, 3, 127])
def test_valid_c_records_match_oracle(valid_c_case, batch_size) -> None:
    expected = structural_oracle(
        valid_c_case.data, valid_c_case.options, NativeSemanticEngine.C
    )
    actual = native_structural_rows(valid_c_case.data, valid_c_case.options, batch_size)
    assert_structural_equivalent(actual, expected)
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

- [ ] **Step 3: Implement in three red-green decoder/framer cycles**

Cycle A adds failing one-byte/65,536-byte split tests, then implements fixed
source reads and incremental UTF-8/Latin-1 decoding. Rebuild and run only that
matrix.

Cycle B adds failing delimiter/quote/double-quote/embedded-newline tests, then
implements field-start, unquoted, quoted, and post-quote states. Rebuild and
run only those cases.

Cycle C adds failing CR/LF/CRLF, BOM, UTF-16 code-unit split, missing-final-
terminator, and large-record tests, then implements CR-pending/EOF and decoder
flush behavior. Rebuild and run the complete valid matrix.

Never decode unread prefetched bytes. Grow only the current record with checked
`Py_ssize_t` arithmetic. On record completion, split fields without a second
complete-record copy where possible, then transfer owned Python strings into
the returned read. At the moment the requested accepted row is releasable,
return without another read, framing step, field split, or callback.

- [ ] **Step 4: Run valid C and lifecycle gates**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py -q -k valid
.venv/bin/pytest tests/native_csv/test_native_api.py tests/native_csv/test_lifecycle.py -q
```

Expected: the loaded-extension source hash matches and all valid-record/source
split combinations pass as raw decoded structure.

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
- Modify: `tests/native_csv/test_differential.py`
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
    expected = structural_oracle(data, ParseOptions(), NativeSemanticEngine.C)
    assert_structural_equivalent(native_structural_outcome(data, batch_size), expected)
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

- [ ] **Step 3: Implement three structural-parity red-green cycles**

Cycle A adds failing width, short-row, blank-before-width, and implicit-index
fixtures, then establishes physical width/index fields from the first
pandas-eligible row and right-pads short rows.

Cycle B adds failing later-wide, consecutive/all-bad, and warning-location
fixtures, then emits synchronous `NativeCSVWarning` and discards later wide
rows.

Cycle C adds failing NUL, quote-junk, quote-in-unquoted-field, and unterminated-
EOF fixtures, then implements pandas C transitions from the frozen structural
oracle rather than Python `csv` behavior.

Keep `operation_max_rows` distinct from `source_eof`; reaching the limit marks
the final nonempty read `done=True` and performs no lookahead.

- [ ] **Step 4: Run complete C differential tests**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_tokenizer_c.py tests/native_csv/test_differential.py tests/test_csv_streaming.py -q -k "not python_mode and not footer"
```

Task 7 explicitly enables the C-structural cases in
`tests/native_csv/test_differential.py`; no `xfail` or disabled native branch
remains for those cases. Expected: raw fields, header structure, warnings,
errors, and physical line context match at all batch sizes. Typed value/scalar
parity remains red and is enabled only in Task 9.

- [ ] **Step 5: Commit C structural parity**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv tests/test_csv_streaming.py
git commit -m "feat: match pandas C CSV edge semantics"
```

---

### Task 8: Implement C-Mode Structural Evidence and Hard Budgets

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/csv_contracts.py`
- Modify: `tests/native_csv/test_evidence.py`

**Interfaces:**
- Produces: native
  `_scan_structural_evidence(source, framing, limits) -> NativeStructuralEvidence`.
- Produces: Python
  `scan_c_structural_evidence(source, framing, limits) -> NativeStructuralEvidence`.
- Does not: type values, handle Python/footer semantics, or make a final
  backend decision.

- [ ] **Step 1: Add failing evidence budget/replay tests**

```python
def test_payload_budget_stops_inside_unterminated_record() -> None:
    limits = replace(native_evidence_limits(None), max_payload_bytes_examined=16)
    evidence = scan_c_structural_evidence(
        ScriptedBinarySource(b'a\n"' + b"x" * 1_000_000),
        c_framing_config(),
        limits,
    )
    assert evidence.status is NativeEvidenceStatus.BUDGET_EXHAUSTED
    assert evidence.payload_bytes_examined == 16
    assert evidence.replay_bytes_retained <= limits.max_replay_bytes
    assert not evidence.eof
```

Cover all five budgets independently; fixed-buffer prefetch excluded from
examined bytes; original quoting/terminators preserved in replay; no
reserialization; `COMPLETE`/`SAMPLE_FULL`/`BUDGET_EXHAUSTED`; zero target;
width too large for one row; empty/header-only/all-bad input; fresh borrow
after evidence; and no warning emission during evidence. Assert the result is
`NativeStructuralEvidence` and contains no typed rows or final pandas columns.

- [ ] **Step 2: Run the red evidence suite**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_evidence.py -q -k "c_mode and structural"
```

Expected: missing structural evidence, replay, or budget behavior fails.

- [ ] **Step 3: Implement three structural-evidence red-green cycles**

Cycle A adds one failing test per examined byte/cell/record budget and
implements pre-increment checks. Budget exhaustion may discard an incomplete
current record immediately and never searches for its terminator.

Cycle B adds independent replay-byte and retained-cell failures, then retains
original fragments/terminators without reserialization and enforces both
retention counters.

Cycle C adds status/EOF/zero-target/one-row-too-wide cases, then returns
`COMPLETE`, `SAMPLE_FULL`, or `BUDGET_EXHAUSTED` with the exact invariants.
For C/no-footer mode, return `SAMPLE_FULL` when target rows are fixed.
Evidence cleanup restores the borrow before returning structural evidence.
Final routing is deliberately deferred to Tasks 9 and 12.

- [ ] **Step 4: Run evidence, cursor, and source tests**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_evidence.py tests/test_source_handle.py tests/test_csv_streaming.py -q -k "c_mode and (sample or evidence or cursor or replay)"
```

Expected: the extension hash is current and all C structural evidence states,
counters, replay bytes, and cursor restores are exact.

- [ ] **Step 5: Commit bounded evidence**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx src/messy_xlsx/parsing/csv_contracts.py src/messy_xlsx/parsing/csv_native.py tests/native_csv tests/test_source_handle.py tests/test_csv_streaming.py
git commit -m "feat: add bounded native CSV evidence"
```

---

### Task 9: Implement C-Mode Pandas Evidence and Physical Value Conversion

**Files:**
- Create: `src/messy_xlsx/parsing/csv_value_adapter.py`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Create: `tests/native_csv/test_pandas_evidence.py`
- Create: `tests/native_csv/test_value_adapter.py`
- Modify: `tests/native_csv/test_differential.py`

**Interfaces:**
- Produces:
  `compile_pandas_evidence(structural, framing, options) -> CompiledNativeEvidence`.
- Produces: `PandasCSVValueAdapter(converters, columns).convert(read, row_offset)`.
- Selects: `UNSUPPORTED_EVIDENCE_TYPE` before a public reader for unsupported
  extension or heterogeneous object evidence.

- [ ] **Step 1: Add failing converter-classification tests**

Execute three reviewed microcycles:

| Cycle | Red test nodes | Minimal implementation | Green command |
|---|---|---|---|
| A — pandas classification | `test_pandas_evidence_classifies_exact_scalar_family`, `test_missing_kind_classification`, `test_original_replay_is_authoritative` | replay parse, scalar-family/missing classifier, successful `CompiledNativeEvidence(evidence, config, None)` | `pytest tests/native_csv/test_pandas_evidence.py -q -k "classifies or missing_kind or replay"` |
| B — typed fallback | `test_heterogeneous_object_selects_fallback`, `test_extension_dtype_selects_fallback`, `test_failure_is_not_fallback` | exclusive fallback result for only the two approved unsupported categories | `pytest tests/native_csv/test_pandas_evidence.py -q -k "fallback or failure_is_not"` |
| C — later conversion | `test_each_converter_exact_values`, `test_late_incompatible_value_context`, `test_missing_sentinel_output` | `PandasCSVValueAdapter.convert` and first-incompatible `StreamingTypeError` | `pytest tests/native_csv/test_value_adapter.py -q` |

Run a specification review after each cycle; do not enable Task 9's typed
differential cases until Cycle C passes.

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
    compiled = compile_evidence(csv)
    assert compiled.evidence is not None
    assert compiled.config is not None
    assert compiled.config.value_converters[0] == PandasValueConverter(kind, missing)
```

Add string dtype with float NaN, uint64, object text, default/configured NA
markers, whitespace, quoted/unquoted empties, structural missing fields,
heterogeneous-object fallback, unsupported extension dtype, and original
replay quoting. Footer exclusion is deferred to Task 11 because Task 9 is
C/no-footer only.

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

For supported input, construct the approved `NativeEvidence` by copying all
structural counters/replay/raw rows and adding pandas `typed_data_rows` and
final `column_names`; compile `ResolvedNativeCSVConfig`. For unsupported
extension or heterogeneous object evidence, return
`CompiledNativeEvidence(None, None,
CSVExecutionReason.UNSUPPORTED_EVIDENCE_TYPE)`.
No backend metric is recorded until Task 12 successfully constructs a reader.

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
.venv/bin/pytest tests/native_csv/test_pandas_evidence.py tests/native_csv/test_value_adapter.py tests/native_csv/test_differential.py tests/test_streaming_normalization.py tests/test_arrow_api.py -q -k "c_mode or physical or normalize_false or late"
```

Task 9 enables typed C-mode cases in `test_differential.py`. Expected: exact
supported physical scalars use the missing-aware oracle comparator and
stable-schema failures pass; Python/footer typed cases remain deferred to
Task 11.

- [ ] **Step 5: Commit the value adapter**

```bash
git add src/messy_xlsx/parsing/csv_value_adapter.py src/messy_xlsx/parsing/csv_native.py tests/native_csv/test_pandas_evidence.py tests/native_csv/test_value_adapter.py tests/native_csv/test_differential.py
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
- Tests: an internal `NativeCSVFramingConfig(semantic_engine=PYTHON,
  skip_footer=0)` against `structural_oracle(..., engine=PYTHON)`; no public
  selector can request this otherwise-internal combination.

- [ ] **Step 1: Add failing Python-engine tests without footer assertions**

Execute these reviewed microcycles, rebuilding the extension before each green
command:

| Cycle | Red test nodes | Minimal implementation | Green command |
|---|---|---|---|
| A — physical lines | `test_python_skiprows_can_bisect_multiline_record`, `test_python_cr_only_line_accounting` | Python-engine physical-line skip and CR/LF accounting | native reinstall; `pytest tests/native_csv/test_tokenizer_python.py -q -k "skiprows or line_accounting"` |
| B — parser recovery | `test_python_csv_error_warns_and_discards`, `test_python_unterminated_quote`, `test_python_quote_junk` | Python quote-error transitions and one-record recovery/diagnostic | native reinstall; `pytest tests/native_csv/test_tokenizer_python.py -q -k "csv_error or unterminated or quote_junk"` |
| C — width/stage order | `test_python_width_before_post_skip`, `test_python_blank_before_width`, `test_python_implicit_index` | first eligible physical width/index and blank/post-skip order | native reinstall; `pytest tests/native_csv/test_tokenizer_python.py -q -k "width or blank or implicit_index"` |

Run the Task 10 specification review after each cycle.

```python
def test_python_skiprows_can_bisect_multiline_record() -> None:
    data = b'a,b\n"first\nsecond",x\n3,4\n'
    options = ParseOptions(skip_rows=2, skip_footer=0)
    expected = structural_oracle(data, options, NativeSemanticEngine.PYTHON)
    actual = native_structural_outcome(
        data,
        options,
        semantic_engine=NativeSemanticEngine.PYTHON,
        skip_footer=0,
    )
    assert_structural_equivalent(actual, expected)
```

Cover NUL, quote-junk, quotes in unquoted fields, unterminated quote, CR-only,
blank rows, `csv.Error` recovery, physical-line skiprows, implicit-index width
fixed before later post-processing, and malformed warning metadata. Force the
internal semantic engine while keeping `skip_footer=0`; do not compare final
footer output in this task.

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
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
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

### Task 11: Implement Footer Retention and Complete Python-Mode Evidence

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/csv_contracts.py`
- Modify: `tests/native_csv/test_tokenizer_python.py`
- Create: `tests/native_csv/test_footer.py`
- Modify: `tests/native_csv/test_evidence.py`
- Modify: `tests/native_csv/test_pandas_evidence.py`
- Modify: `tests/native_csv/test_differential.py`
- Modify: `tests/test_csv_streaming.py`

**Interfaces:**
- Produces: successful parsed-row footer deque bounded by `skip_footer`.
- Produces: accepted output bounded by `requested_rows`, with no work after the
  requested row becomes releasable.
- Produces:
  `scan_python_structural_evidence(...) -> NativeStructuralEvidence`, which
  reaches physical EOF or returns `BUDGET_EXHAUSTED`.
- Produces: final `CompiledNativeEvidence` for supported Python/footer input
  using the Task 9 pandas compiler.

- [ ] **Step 1: Add failing footer-order and bound tests**

```python
def test_quote_error_is_discarded_before_footer_removal() -> None:
    data = b"a,b\n1,2\n3,4\n\"bad\n5,6\n"
    options = ParseOptions(skip_footer=1)
    actual = native_outcome(data, options, batch_size=1)
    expected = materialized_oracle(data, options)
    assert_oracle_equivalent(actual, expected)


def test_wide_trailing_row_can_disappear_as_footer_without_warning() -> None:
    data = b"a,b\n1,2\n3,4,5\n"
    outcome = native_outcome(data, ParseOptions(skip_footer=1), batch_size=1)
    assert outcome.rows == ((1, 2),)
    assert outcome.warnings == ()
```

Cover blank/wide rows occupying footer slots, consecutive parser errors,
footer zero/one/all/greater-than-rows, multiline records, header-none
all-footer, batch sizes `1,2,3,127`, and exact successor-row counters. Add
Python/footer evidence cases for each hard budget, original-byte replay,
physical EOF, unsupported pandas evidence, and typed missing-aware oracle
comparison.

- [ ] **Step 2: Run and confirm the footer red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_footer.py tests/native_csv/test_tokenizer_python.py tests/native_csv/test_evidence.py tests/native_csv/test_pandas_evidence.py tests/native_csv/test_differential.py -q -k "footer or python_mode"
```

Expected: footer ordering/bounds, Python structural evidence, or Python typed
evidence assertions fail for the intended missing behavior.

- [ ] **Step 3: Implement footer ownership, then Python evidence, in separate cycles**

Cycle A implements parsed-success footer ownership. Retain at most
`skip_footer` successfully parsed rows. Parser-error records never occupy a
slot; blank or over-wide successfully parsed records do until their later
classification stage. Release an output row only after the required successor
count is known. Stop all source reads, framing, field tokenization, and
callbacks immediately after the requested output becomes releasable.

Cycle B reuses those exact Python/footer transitions for structural evidence.
Unlike C/no-footer evidence, it never returns `SAMPLE_FULL`: it reaches
physical EOF inside every hard limit or returns `BUDGET_EXHAUSTED`.

Cycle C passes the exact header/sample/true-footer replay into Task 9's pandas
compiler and produces `CompiledNativeEvidence`. It enables Python/footer typed
cases in `test_differential.py`; unsupported evidence returns the typed
fallback reason but records no execution decision yet.

- [ ] **Step 4: Run Python/footer differential and architecture tests**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/native_csv/test_footer.py tests/native_csv/test_tokenizer_python.py tests/native_csv/test_evidence.py tests/native_csv/test_pandas_evidence.py tests/native_csv/test_differential.py tests/test_csv_streaming.py -q -k "footer or python_mode"
```

Expected: Python/footer cases match the materialized oracle and every retained
row counter stays within `batch_size + skip_footer`; evidence reaches EOF or
exhausts a named budget, and supported typed cases use the missing-aware
comparator.

- [ ] **Step 5: Commit footer semantics**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx src/messy_xlsx/parsing/csv_native.py src/messy_xlsx/parsing/csv_contracts.py tests/native_csv/test_footer.py tests/native_csv/test_tokenizer_python.py tests/native_csv/test_evidence.py tests/native_csv/test_pandas_evidence.py tests/native_csv/test_differential.py tests/test_csv_streaming.py
git commit -m "feat: add bounded Python-mode CSV footers"
```

---

### Task 12: Implement Exact Capability Selection, the Candidate Seam, and Native Integration

**Files:**
- Modify: `src/messy_xlsx/parsing/csv_native.py`
- Modify: `src/messy_xlsx/parsing/csv_streaming.py`
- Modify: `src/messy_xlsx/workbook.py`
- Modify: `src/messy_xlsx/parsing/contracts.py`
- Create: `tests/native_csv/test_integration.py`
- Create: `tests/native_csv/test_selection.py`
- Modify: `tests/test_csv_streaming.py`
- Modify: `tests/test_arrow_api.py`

**Interfaces:**
- Preserves: `prepare_csv_streaming_reader(source, plan, metrics, *, construction_owner=None) -> PreparedStreamingReader`.
- Produces: internal `NativeCSVReader(StreamingBatchReader)`.
- Produces:
  `resolve_native_capability(registry, handler, *, candidate_token=None) -> NativeCapability`.
- Produces:
  `_run_candidate_artifact_smoke(source, plan, metrics) -> PreparedStreamingReader`,
  whose wrapper supplies the module-owned token internally.
- Produces:
  `_prepare_native_streaming_reader(source, plan, metrics, *,
  candidate_token=None) -> PreparedStreamingReader`, the only constructor that
  accepts the internal token.
- Removes: `_PreownedPandasReader` and pandas `chunksize` from the full pass.
- Retains: current bounded inspection, `PreparedStreamingReader`,
  `_CloseOnceReader`, normalization compilation, physical encoding, and public
  display-name handling.

- [ ] **Step 1: Add failing end-to-end integration tests**

```python
@pytest.mark.parametrize("batch_size", [1, 2, 3, 127])
def test_private_candidate_native_stream_matches_oracle(csv_case, batch_size) -> None:
    with candidate_native_stream(csv_case.data, csv_case.options, batch_size) as stream:
        actual = native_outcome_from_batches(stream)
    assert_oracle_equivalent(
        actual,
        materialized_oracle(csv_case.data, csv_case.options),
    )
```

Cover `normalize=True/False`, all-null columns, duplicate/non-string labels,
dataframe chunks and global `RangeIndex`, path/seekable/nonseekable/one-byte
sources, final nonempty `done`, `max_rows == 0`, early close, stable schema
before return, and records larger than 8 MiB after the evidence sample.

Add a table-driven selector test for every precedence branch:

```text
custom registry/component/handler/parse override -> CUSTOM_SPI
false public gate -> PRODUCTION_GATE_DISABLED
true gate plus exact "1" kill switch -> KILL_SWITCH
unsupported/free-threaded runtime -> UNSUPPORTED_RUNTIME before import
ModuleNotFoundError/ImportError/import-time shared-library OSError -> IMPORT_OR_LOAD_FAILURE
API or pandas semantic mismatch -> HANDSHAKE_MISMATCH
generated multi-header -> MULTI_HEADER_EXACTNESS
structural budget exhaustion -> EVIDENCE_BUDGET
unsupported pandas evidence -> UNSUPPORTED_EVIDENCE_TYPE
eligible, compiled evidence, successful reader -> NATIVE_SELECTED
```

Inject an import spy and prove unsupported runtime never imports. Prove only
`ModuleNotFoundError`, `ImportError`, and import-time shared-library `OSError`
become capability fallback; `ValueError`, `RuntimeError`, `MemoryError`, and
arbitrary execution exceptions propagate.

Test `_run_candidate_artifact_smoke(source, plan, metrics)` without a token
argument. It succeeds by supplying `_CANDIDATE_SMOKE_TOKEN` internally.
Direct adapter construction with a caller-created token fails. The wrapper
rechecks exact built-in ownership, runtime, kill switch, import, handshake,
evidence, and configuration, bypassing only the false production constant.
Public routing in the same test remains materialized.

- [ ] **Step 2: Run the integration red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_selection.py tests/native_csv/test_integration.py tests/test_csv_streaming.py tests/test_arrow_api.py -q -k "csv or capability or candidate"
```

Expected: selector precedence, private-token protection, or private native
integration fails because exact capability loading/integration does not exist
and the full pass still uses pandas chunks.

- [ ] **Step 3: Implement selection, the private seam, and reader integration in separate cycles**

Cycle A implements `resolve_native_capability`. It first verifies exact
built-in registry/detector/handler/component ownership, then the public gate
(unless the identity-equal module token is present), then reads the kill switch
once, checks non-free-threaded CPython 3.11–3.14 before import, narrowly loads
the extension, and validates both handshake constants. `NativeCapability` is
defined in `csv_native.py` as:

```python
@dataclass(frozen=True, slots=True)
class NativeCapability:
    module: NativeModule | None
    fallback_reason: CSVExecutionReason | None
```

Construction enforces exactly one populated field.

Cycle B applies generated-header and compiled-evidence decisions before
full-pass startup. It restores the evidence borrow before constructing any
materialized reader. Evidence parse failure records the existing failure
metric but no CSV execution decision. No native-to-materialized transition is
permitted after full-pass binding.

Cycle C adds
`_prepare_native_streaming_reader(..., candidate_token=None)` and
`_run_candidate_artifact_smoke(source, plan, metrics)`. The wrapper, not its
caller, supplies `_CANDIDATE_SMOKE_TOKEN` to the internal constructor. The
wrapper returns the ordinary frozen `PreparedStreamingReader`; callers inspect
`prepared.reader.execution_decision` and must close `prepared.reader` in
`finally`. A caller-created token is rejected by the constructor. There is no
environment or public selector for the bypass.

Cycle D removes Task 1's gate-mutating autouse fixture. Retained native
characterization tests use `candidate_native_stream`, which calls only
`_run_candidate_artifact_smoke`; public route tests run with the checked-in
false gate and remain materialized.

`prepare_csv_streaming_reader` performs inspection, native evidence, pandas
converter compilation, normalization-plan compilation, and routing while no
full-pass borrow is open. It returns an inert `NativeCSVReader`; first
`read_next_batch()` opens `SourceHandle.open_binary()`, binds the tokenizer,
and converts `NativeCSVRead` columns through `PandasCSVValueAdapter` and the
existing normalization wrapper.

If capability/evidence selection chooses materialized fallback, restore the
evidence borrow first and construct the existing materialized streaming
adapter. Never catch a native execution failure as fallback.

After a reader is constructed successfully, record exactly one immutable
`CSVExecutionDecision`, retain it on
`NativeCSVReader.execution_decision` (and the materialized/custom reader
counterpart), and return the reader. Sequential metrics replace
`last_csv_execution` and increment the per-kind/per-reason count exactly once.

- [ ] **Step 4: Run public, Arrow, and custom-registry gates**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_selection.py tests/native_csv/test_integration.py tests/test_csv_streaming.py tests/test_arrow_api.py tests/test_reader_routing.py tests/test_architecture_contracts.py -q
```

Expected: private native integration passes; public default remains
`PRODUCTION_GATE_DISABLED`; custom registry behavior remains materialized and
authoritative; selector order/catches are exact; each successful reader retains
one immutable decision and failed evidence retains none.

- [ ] **Step 5: Commit native integration**

```bash
git add src/messy_xlsx/parsing/csv_native.py src/messy_xlsx/parsing/csv_streaming.py src/messy_xlsx/parsing/contracts.py src/messy_xlsx/workbook.py tests/native_csv/test_selection.py tests/native_csv/test_integration.py tests/test_csv_streaming.py tests/test_arrow_api.py
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

Execute three reviewed microcycles, forcing a native reinstall before each
green command:

| Cycle | Red test nodes | Minimal implementation | Green command |
|---|---|---|---|
| A — warnings | `test_evidence_warning_is_suppressed`, `test_full_warning_once_stacklevel_three`, `test_warning_promoted_to_error_uses_materialized_format_boundary` | captured evidence diagnostics and one full-pass `warnings.warn(..., stacklevel=3)` boundary | native reinstall; `pytest tests/native_csv/test_failures.py -q -k warning` |
| B — data errors | `test_eager_and_lazy_decode_boundaries`, `test_fallback_encoding_footer_evidence_uses_legacy_terminal_context`, `test_internal_state_error_propagates` | contextual `FormatError` for data/source failures, legacy fallback context, unchanged internal/process failures | native reinstall; `pytest tests/native_csv/test_failures.py -q -k "decode or fallback_encoding or internal_state"` |
| C — cleanup precedence | `test_cleanup_precedence_matrix`, `test_return_gap_restore_failure`, `test_finalizer_and_early_close_restore` | tokenizer-close then borrow-restore order and ordinary/process precedence | native reinstall; `pytest tests/native_csv/test_lifecycle.py tests/test_resource_lifecycle.py tests/test_stream_lifecycle.py -q -k "cleanup or return_gap or finalizer or early_close"` |

Run specification and quality reviews after each cycle.

```python
def test_warning_promoted_to_error_uses_materialized_format_boundary() -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("error", pd.errors.ParserWarning)
        with pytest.raises(FormatError) as caught:
            exhaust_native(b"a,b\n1,2\n3,4,5\n")
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
never closed. Assert every caller-visible parser warning uses the pandas
category/message and `stacklevel=3`. Assert internal native state/configuration
`ValueError`, `RuntimeError`, and invariant failures propagate unchanged
rather than becoming fallback or `FormatError`.

- [ ] **Step 2: Run the failure suite and confirm red**

Run:

```bash
.venv/bin/pytest tests/native_csv/test_failures.py tests/native_csv/test_lifecycle.py tests/test_source_handle.py tests/test_resource_lifecycle.py tests/test_stream_lifecycle.py -q
```

Expected: missing native error mapping/cleanup cases fail.

- [ ] **Step 3: Implement exact adapter translation and cleanup order**

Data-driven native parser/decoder/source failures become contextual
`FormatError`. Internal state/configuration/invariant defects propagate
unchanged and are never treated as capability fallback.
The adapter emits full-pass diagnostics exactly once with
`warnings.warn(message, category, stacklevel=3)`; evidence diagnostics are
captured only.
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
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
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

### Task 14: Close Deterministic Bounds and Allocation-Fault Gates

**Files:**
- Modify: `src/messy_xlsx/_csv_tokenizer.pyx`
- Create: `tests/native_csv/test_bounds.py`
- Create: `tests/native_csv/test_failure_injection.py`
- Create: `tests/native_csv/test_reentrancy.py`

**Interfaces:**
- Produces: observer events `before_source_read`, `after_source_read`,
  `before_warning`, `after_warning`, and `before_return`.
- Produces: `_allocation_sites_for_tests()` and
  `_set_allocation_failure_for_tests(site)`.
- Proves: native safety and literal memory/row bounds before performance work.

- [ ] **Step 1: Add failing deterministic counter and allocation tests**

```python
@dataclass(frozen=True, slots=True)
class ExpectedLedger:
    current_record_active: bool
    current_payload: int
    footer_payload: int
    output_payload: int


def assert_snapshot(
    snapshot: NativeDebugState,
    expected: ExpectedLedger,
    batch_size: int,
    footer: int,
) -> None:
    assert snapshot.output_rows_retained <= batch_size
    assert snapshot.post_output_rows_retained <= footer
    assert snapshot.field_tokenized_successor_rows <= footer
    assert snapshot.undecoded_buffer_bytes <= 65_536
    assert snapshot.current_record_active is expected.current_record_active
    assert snapshot.current_record_payload_bytes == expected.current_payload
    assert snapshot.footer_payload_bytes == expected.footer_payload
    assert snapshot.output_payload_bytes == expected.output_payload
    assert snapshot.logical_payload_bytes == (
        snapshot.current_record_payload_bytes
        + snapshot.footer_payload_bytes
        + snapshot.output_payload_bytes
    )


EXPECTED_PYMEM_SITES = (
    "input_buffer.allocate",
    "record_buffer.allocate",
    "record_buffer.resize",
    "field_offsets.allocate",
    "field_offsets.resize",
    "footer_ring.allocate",
)


@pytest.mark.parametrize("site", EXPECTED_PYMEM_SITES)
def test_every_native_allocation_failure_is_safe(site) -> None:
    set_allocation_failure(site)
    tokenizer = tokenizer_for_fixture()
    with pytest.raises(MemoryError):
        tokenizer.read_batch(1, lambda warning: None)
    tokenizer.close()
```

Each fixture defines an ordered map from observer event ordinal to
`ExpectedLedger`, calculated directly from its checked-in original record
bytes; the implementation's counters are never used to derive expectations.

Assert no source read/tokenization/callback after a requested batch becomes
releasable, oversized full-pass records remain valid, overflow/realloc paths
are checked, recursive mutating calls fail, `debug_state` is allowed
reentrantly, and observer failure is terminal. A source scan and native
registry assertion must prove
`tuple(_allocation_sites_for_tests()) == EXPECTED_PYMEM_SITES`; missing,
renamed, duplicate, or unregistered direct `PyMem_*` calls fail.

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

- [ ] **Step 4: Rebuild and run deterministic bound/fault gates**

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
PYTHONMALLOC=debug .venv/bin/pytest tests/native_csv/test_bounds.py tests/native_csv/test_failure_injection.py tests/native_csv/test_reentrancy.py -q
```

Expected: the loaded extension hash matches; every event snapshot has exact
test-owned payload ledgers; the implementation registry equals the test-owned
allocation manifest; every injected failure unwinds safely.

- [ ] **Step 5: Commit deterministic safety gates**

```bash
git add src/messy_xlsx/_csv_tokenizer.pyx tests/native_csv/test_bounds.py tests/native_csv/test_failure_injection.py tests/native_csv/test_reentrancy.py
git commit -m "test: prove native CSV deterministic bounds"
```

---

### Task 15: Pass Differential Fuzz and Sanitizer Gates

**Files:**
- Create: `tests/native_csv/test_fuzz.py`
- Modify: `tests/native_csv/test_fuzz_contract.py`
- Modify: `tests/native_csv/fuzz_worker.py`
- Add: `tests/native_csv/regressions/*.bin`
- Modify: `scripts/run_native_csv_fuzz.py`
- Modify: `scripts/run_native_csv_sanitizers.sh`
- Create: `.github/workflows/native-safety.yml`
- Modify: `.github/workflows/native-abi.yml`

**Interfaces:**
- Produces: fixed-seed C/Python differential fuzzing with missing-aware oracle
  comparisons and minimized regression fixtures.
- Produces: clean native ASan/UBSan and debug-allocator jobs that prove the
  loaded extension matches the current `.pyx`.

- [ ] **Step 1: Add failing fuzz/sanitizer contract tests**

Test that the driver requires exactly:

```text
C seed = 0x0C5A14
Python seed = 0xBADC5EED
examples per engine = 5,000
timeout = 300 seconds
```

The worker generates valid/malformed bytes, randomized source chunk splits,
all batch sizes, both source ownership classes, and both semantic engines.
Schema-compatible cases compare values with `assert_oracle_equivalent`,
including scalar families, columns, warnings, error class/message/context, and
source lifecycle. Late incompatible values and late path decoding assert their
separate streaming exceptions.

Add workflow-contract tests that fail unless sanitizer commands force a clean
native rebuild and run the complete native test suite.

- [ ] **Step 2: Run the red contract**

```bash
.venv/bin/pytest tests/native_csv/test_fuzz_contract.py -q
```

Expected: assertion failures identify the absent driver, seeds, worker
protocol, regression persistence, or sanitizer configuration.

- [ ] **Step 3: Implement and run seeded differential fuzzing**

Each mismatch is minimized and written to
`tests/native_csv/regressions/<engine>-<sha256>.bin` with an adjacent expected
oracle record before the implementation is changed. Rerunning the same seed
must reproduce the same case ordering and output hashes.

Run:

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/python scripts/run_native_csv_fuzz.py \
  --c-seed 0x0C5A14 \
  --python-seed 0xBADC5EED \
  --examples 5000 \
  --timeout 300
```

Expected: the source-hash guard passes and both 5,000-case suites match the
approved oracle or a named streaming-exception contract.

- [ ] **Step 4: Implement and run clean sanitizer/debug builds**

`scripts/run_native_csv_sanitizers.sh` creates a fresh temporary source
extraction and builds with:

```text
CFLAGS=-O1 -g -fno-omit-frame-pointer -fsanitize=address,undefined -Werror
LDFLAGS=-fsanitize=address,undefined
LD_PRELOAD=$(gcc -print-file-name=libasan.so)
ASAN_OPTIONS=detect_leaks=0:halt_on_error=1
UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1
```

It validates that `gcc -print-file-name=libasan.so` returns an existing
absolute path before setting `LD_PRELOAD`, sets
`MESSY_XLSX_BUILD_MODE=native`, installs the extension freshly, checks
`NATIVE_SOURCE_SHA256`, runs `tests/native_csv`, then repeats lifecycle,
allocation-failure, and reentrancy suites under `PYTHONMALLOC=debug`.

Run:

```bash
bash scripts/run_native_csv_sanitizers.sh
PYTHONMALLOC=debug .venv/bin/pytest tests/native_csv -q
```

Expected: no sanitizer, debug allocator, stale-extension, leak, or lifecycle
failure.

- [ ] **Step 5: Commit fuzz and sanitizer gates**

```bash
git add tests/native_csv/test_fuzz.py tests/native_csv/test_fuzz_contract.py tests/native_csv/fuzz_worker.py tests/native_csv/regressions scripts/run_native_csv_fuzz.py scripts/run_native_csv_sanitizers.sh .github/workflows/native-safety.yml .github/workflows/native-abi.yml
git commit -m "test: fuzz and sanitize native CSV tokenizer"
```

---

### Task 16: Remove Superseded Runtime Code

**Files:**
- Create: `src/messy_xlsx/parsing/csv_probe.py`
- Modify: `src/messy_xlsx/parsing/csv_handler.py`
- Modify: `src/messy_xlsx/parsing/csv_streaming.py`
- Delete: `src/messy_xlsx/parsing/csv_io.py`
- Create: `tests/native_csv/reference_streaming.py`
- Modify: `tests/test_architecture_contracts.py`
- Modify: `tests/packaging/test_build_modes.py`

**Interfaces:**
- Retains: only the small logical-record probe helpers needed by legacy
  metadata inspection in installed `csv_probe.py`.
- Moves: the superseded Python framing/filter/footer implementation under
  tests as a differential/performance reference.
- Proves: installed runtime contains one full-pass implementation and no pandas
  `chunksize` reader.

- [ ] **Step 1: Add failing runtime-architecture and distribution tests**

```python
def test_installed_csv_runtime_has_one_full_pass() -> None:
    assert not runtime_imports("messy_xlsx.parsing.csv_io")
    assert not source_contains("src/messy_xlsx/parsing/csv_streaming.py", "chunksize=")
    assert not source_contains("src/messy_xlsx/parsing/csv_streaming.py", "_PreownedPandasReader")
    assert source_contains("src/messy_xlsx/parsing/csv_native.py", "NativeCSVReader")


def test_reference_implementation_is_test_only() -> None:
    assert not any_runtime_imports("tests.native_csv.reference_streaming")
```

Distribution inventory tests assert the test-only reference is absent from
wheels, `csv_probe.py` is present, and no `csv_io.py` or generated native
source is installed.

- [ ] **Step 2: Run the red architecture tests**

```bash
.venv/bin/pytest tests/test_architecture_contracts.py tests/packaging/test_build_modes.py -q -k "csv or runtime_inventory"
```

Expected: `csv_io.py` and/or a production import still exists.

- [ ] **Step 3: Move the reference and retain only probes**

Use:

```bash
rg -n "csv_io|chunksize|_PreownedPandasReader" src tests
```

Copy the superseded framing/filter/footer code needed for differential and
footer benchmarks to `tests/native_csv/reference_streaming.py`. Move only
inspection-time logical-record helpers still consumed by `CSVHandler` or
bounded metadata inspection into `csv_probe.py`, update imports, then delete
`csv_io.py`.

No tokenizer behavior changes in this task. If the move exposes a difference,
add an oracle regression and fix it while the production gate is still false.

- [ ] **Step 4: Run architecture and semantic regression gates**

```bash
.venv/bin/pytest tests/test_architecture_contracts.py tests/packaging/test_build_modes.py tests/native_csv tests/test_csv_streaming.py tests/compatibility -q
.venv/bin/ruff check src/messy_xlsx/parsing/csv_probe.py src/messy_xlsx/parsing/csv_handler.py src/messy_xlsx/parsing/csv_streaming.py tests/native_csv/reference_streaming.py
```

Expected: one installed native full pass, a test-only Python reference, and no
semantic/lifecycle regression.

- [ ] **Step 5: Commit runtime cleanup**

```bash
git add src/messy_xlsx/parsing/csv_probe.py src/messy_xlsx/parsing/csv_handler.py src/messy_xlsx/parsing/csv_streaming.py tests/native_csv/reference_streaming.py tests/test_architecture_contracts.py tests/packaging/test_build_modes.py
git add -u src/messy_xlsx/parsing/csv_io.py
git commit -m "refactor: retire superseded CSV streaming runtime"
```

---

### Task 17: Set v1.0.0 Metadata and Verify the Parent Handoff

**Files:**
- Modify: `pyproject.toml`
- Modify: `src/messy_xlsx/__init__.py`
- Modify: `tests/test_integration.py`
- Verify or update if task numbers change:
  `docs/superpowers/plans/2026-07-22-parser-performance-v1.md`

**Interfaces:**
- Sets: artifact/package/module version exactly `1.0.0` before candidate
  filenames, metadata, and hashes are created.
- Verifies: parent Task 14 points only to this native implementation and parent Task 20
  with the exact release-SHA artifact rebuild.

- [ ] **Step 1: Add a failing v1 artifact-metadata test**

```python
def test_native_artifact_version_is_v100() -> None:
    assert metadata.version("messy-xlsx") == "1.0.0"
    assert messy_xlsx.__version__ == "1.0.0"
```

Run:

```bash
.venv/bin/pytest tests/test_integration.py -q -k native_artifact_version
```

Expected: assertions fail because both values remain `0.10.0`.

- [ ] **Step 2: Set both versions and verify parent ownership**

Set `pyproject.toml` and `src/messy_xlsx/__init__.py` to `1.0.0`.
Confirm the already-amended parent plan:

- replace Task 14's pandas-chunk `CSVStreamingReader` instructions and file map
  with a link to this native plan and `NativeCSVReader`;
- mark parent Task 14 complete only after this plan's Task 24 passes;
- make parent Task 20 verify the already-set `1.0.0` metadata rather than
  expecting a `0.10.0` failure;
- retain parent Task 20 ownership of README/docs/changelog, release readiness,
  and the full final artifact rebuild on that exact documentation/release SHA;
- state that no earlier manifest survives any metadata/docs source change;
- retain explicit user authorization before creating or pushing `v1.0.0`.

- [ ] **Step 3: Run version and plan-consistency checks**

```bash
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/pytest tests/test_integration.py -q -k "version or exports"
rg -n "pandas.chunk|_PreownedPandasReader|CSVStreamingReader" docs/superpowers/plans/2026-07-22-parser-performance-v1.md || true
git diff --check
```

Expected: version assertions pass and any remaining parent references are
explicit historical/superseded notes rather than executable instructions.

- [ ] **Step 4: Commit v1 artifact metadata and parent reconciliation**

```bash
git add pyproject.toml src/messy_xlsx/__init__.py tests/test_integration.py docs/superpowers/plans/2026-07-22-parser-performance-v1.md
git commit -m "release: set native parser v1 artifact metadata"
```

---

### Task 18: Pass the Authoritative Performance Gate

**Files:**
- Create: `benchmarks/native_csv.py`
- Create: `scripts/run_native_csv_benchmarks.py`
- Modify: `tests/native_csv/reference_streaming.py`
- Create: `tests/test_performance/test_native_csv_contract.py`
- Create: `.github/workflows/native-performance.yml`
- Modify: `.github/workflows/test.yml`

**Interfaces:**
- Produces: deterministic benchmark corpora from seed `0x0C5A14`.
- Produces: machine-readable per-run and aggregate benchmark reports.
- Compares: tokenizer-only no-footer execution with direct pandas C-engine
  parsing.
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
counters, source position at each batch, and a stable output hash. The
authoritative report additionally contains exact `phase`, forty-character
`revision`, installed native-wheel SHA-256, runner-image digest, runner CPU
identity, `python_version == "3.12"`, `pandas_version == "3.0.5"`, and
`thresholds_passed is True`, top-level `workflow_run_id`, and expected
aggregate artifact name; validation rejects any mismatch.

The large logical record must appear after the first 1,000 accepted data rows.
The footer corpus must use `skip_footer=10`, reach physical EOF within every
evidence limit, and record `CSVExecutionKind.NATIVE`. Report-contract tests
parameterize both phases by monkeypatching the gate: candidate isolated metrics
require private native plus public
`MATERIALIZED_FALLBACK/PRODUCTION_GATE_DISABLED`; final metrics require public
native plus kill-switch fallback.

- [ ] **Step 2: Run the contract tests and confirm the red state**

Run:

```bash
.venv/bin/pytest tests/test_performance/test_native_csv_contract.py -q
```

Expected: pytest collects successfully and in-test imports or contract
assertions fail because the deterministic harness and report validator do not
exist.

- [ ] **Step 3: Implement deterministic generation, alternating runs, and validation**

`benchmarks/native_csv.py` generates each corpus without timing file creation.
`scripts/run_native_csv_benchmarks.py` installs or accepts the exact native
wheel path and requires `--phase candidate|final`. It also exposes a
non-benchmarking validation mode:

```bash
python scripts/run_native_csv_benchmarks.py \
  --phase candidate|final \
  --validate-report REPORT.json
```

That mode parses the recorded metadata, hashes, counters, ratios, phase-specific
execution decisions, and thresholds without regenerating a corpus or rerunning
timings. Tokenizer-only no-footer timing invokes the internal tokenizer
extension directly with a precompiled `ResolvedNativeCSVConfig`; immediately
before timing, an untimed Task 12 candidate-wrapper run proves the same corpus
selects `NATIVE/NATIVE_SELECTED`. End-to-end correctness/memory timing uses
the private wrapper and isolated metrics.
The public route is executed outside the timer. It must be disabled fallback
for candidate and native (plus kill-switch fallback) for final.
The harness performs three warmups, then alternates contenders for seven
measured runs. It reports medians and computes geometric means from per-corpus
ratios.

The footer baseline imports only
`tests/native_csv/reference_streaming.py`; installed runtime never imports it.

Reject a performance report unless all of these are true:

```text
each no-footer native/direct-pandas-C median ratio <= 3.0
geometric mean of no-footer native/direct-pandas-C ratios <= 2.0
geometric mean of footer reference/native median ratios >= 4.0
candidate: every private case is NATIVE and every public case is PRODUCTION_GATE_DISABLED
final: every public case is NATIVE and exact "1" produces KILL_SWITCH fallback
every deterministic row, fixed-buffer, and logical-payload bound passes
```

Direct pandas C is not an equivalent footer baseline. Footer reports include
the retained Python reference and end-to-end materialized `CSVHandler`.

- [ ] **Step 4: Add the authoritative and corroborating CI jobs**

`.github/workflows/native-performance.yml` builds or downloads the exact
manylinux native artifact and runs the gate on the dedicated Ubuntu 24.04
x86-64 benchmark runner, its pinned image/CPU identity, and CPython 3.12.
In this task it builds one fallback-mode source archive and the manylinux wheel
from a clean extraction of that archive; Task 21 later changes only the
artifact input wiring so the aggregate candidate/final workflow supplies that
same exact wheel.
Other supported native platforms produce corroborating reports without
relaxing or replacing the authoritative thresholds. Store the JSON report,
corpora manifest, wheel hash, and runner identity as SHA-scoped artifacts.
Name the authoritative report
`$ARTIFACT_PHASE-$GITHUB_SHA-native-csv-performance.json` and its artifact
`$ARTIFACT_PHASE-$GITHUB_SHA-performance`; never publish a generic
first-match filename. The worker asserts `pandas.__version__ == "3.0.5"`
before warmup and records it in the report.

There is no automatic waiver. A threshold miss leaves
`_NATIVE_CSV_PRODUCTION_READY = False` until the design and this plan receive
explicit user-approved amendments.

- [ ] **Step 5: Run the local contract and representative benchmark**

Run:

```bash
.venv/bin/pytest tests/test_performance/test_native_csv_contract.py -q
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python --no-deps --reinstall -e .
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --rows 300000 \
  --seed 0x0C5A14 \
  --warmups 3 \
  --runs 7 \
  --phase candidate \
  --output /tmp/messy-xlsx-native-csv-performance.json
```

Expected: the extension source hash and report validator pass; private
candidate metrics are native while public metrics remain disabled fallback.
Local timing is diagnostic; only the dedicated CI runner decides the timing
gate.

- [ ] **Step 6: Commit the performance gate**

```bash
git add benchmarks/native_csv.py scripts/run_native_csv_benchmarks.py tests/native_csv/reference_streaming.py tests/test_performance/test_native_csv_contract.py .github/workflows/native-performance.yml .github/workflows/test.yml
git commit -m "perf: gate native CSV production routing"
```

---

### Task 19: Add Exact-Sdist Artifact and Provenance Tooling

**Files:**
- Modify: `build_support.py`
- Modify: `requirements/native-release.txt`
- Create: `scripts/release_artifacts.py`
- Modify: `tests/packaging/test_build_support.py`
- Modify: `tests/packaging/test_build_modes.py`
- Create: `tests/packaging/test_release_artifacts.py`
- Modify: `setup.py`
- Modify: `pyproject.toml`
- Modify: `.gitignore`
- Modify: `Makefile`

**Interfaces:**
- Produces: exact native/fallback build modes from one source archive.
- Produces: `release_artifacts.py record|assemble|verify`.
- Proves: exact source-archive lineage, immutable phase/revision/hash metadata,
  and a nine-distribution allowlist independently of GitHub Actions.

- [ ] **Step 1: Write failing build and artifact-provenance tests**

Execute three reviewed microcycles:

| Cycle | Red test nodes | Minimal implementation | Green command |
|---|---|---|---|
| A — build mode | `test_explicit_build_modes`, `test_invalid_build_mode`, `test_supported_default`, `test_fallback_has_no_extension` | pure `resolve_build_mode`, setup delegation, exact wheel purity/tag checks | `pytest tests/packaging/test_build_support.py tests/packaging/test_build_modes.py -q -k "build_mode or supported_default or no_extension"` |
| B — static inventory | `test_rejects_wrong_count_tags_and_contents`, `test_rejects_metadata_drift`, `test_rejects_forbidden_sdist_files` | read-only wheel/sdist inventory and metadata validator | `pytest tests/packaging/test_release_artifacts.py -q -k "count or tags or contents or metadata or forbidden"` |
| C — provenance | `test_record_fragment_schema`, `test_assemble_rejects_lineage_mismatch`, `test_verify_rejects_hash_phase_revision_mismatch`, `test_performance_record_binds_report_revision_phase_wheel_and_runner` | `record`, transactional `assemble`, immutable performance record, and static `verify` subcommands | `pytest tests/packaging/test_release_artifacts.py -q -k "fragment or lineage or hash or phase or revision or performance"` |

Run build/release specification review after each cycle.

Test `build_support.resolve_build_mode()` with explicit `native`, explicit
`fallback`, invalid values, supported CPython, unsupported/free-threaded
runtimes, and supported/unsupported architectures. Test clean native and
fallback editable installations from separate source extractions so a stale
extension cannot satisfy fallback assertions.

Create corrupt synthetic artifacts and manifests that must be rejected for
missing/extra/duplicate files, wrong tags or purity, extension absence or
presence in the wrong variant, cross-wheel `METADATA` drift, missing `.pyx` or
build files in the source archive, forbidden generated native files,
`CONTINUE.md`, `.superpowers`, or `uv.lock`, wrong source-archive lineage,
phase/namespace mismatch, or altered SHA-256 content.

- [ ] **Step 2: Run the tests and confirm the red state**

Run:

```bash
uv pip install --python .venv/bin/python -e ".[dev]" -r requirements/native-release.txt
.venv/bin/pytest \
  tests/packaging/test_build_support.py \
  tests/packaging/test_build_modes.py \
  tests/packaging/test_release_artifacts.py -q
```

Expected: missing build-helper and artifact/provenance assertions fail.

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

Do not weaken the Task 2 build-system or development-tool pins.

- [ ] **Step 4: Implement artifact/provenance tools**

`scripts/release_artifacts.py` records source and wheel hashes, assembles an
allowlisted release set, and verifies phase, revision, nine-file count,
filenames, wheel tags, purity, extension inventory, cross-wheel `METADATA`,
source inventory, exact source-archive lineage, SHA-256 content,
`abi3audit --strict`, `twine check`, and complete per-wheel smoke-record
coverage. It never attempts to install an incompatible cross-platform wheel on
the aggregation host.

The JSON manifest schema is fixed and versioned:

```python
class ArtifactRecord(TypedDict):
    filename: str
    sha256: str
    kind: Literal["sdist", "native-wheel", "fallback-wheel"]
    python_tag: Literal["source", "cp311", "py3"]
    abi_tag: Literal["source", "abi3", "none"]
    platform_tag: str
    source_sdist_sha256: str


class WheelSmokeRecord(TypedDict):
    wheel_sha256: str
    phase: Literal["candidate", "final"]
    python_version: Literal["3.11", "3.12", "3.13", "3.14"]
    platform_tag: str
    extension_present: bool
    pip_check_passed: Literal[True]
    public_decision: str
    private_decision: str | None
    kill_switch_decision: str | None


class PerformanceRecord(TypedDict):
    filename: str
    sha256: str
    phase: Literal["candidate", "final"]
    revision: str
    native_wheel_sha256: str
    runner_image_digest: str
    runner_cpu_identity: str
    python_version: Literal["3.12"]
    pandas_version: Literal["3.0.5"]
    thresholds_passed: Literal[True]
    workflow_run_id: int
    aggregate_artifact_name: str


class ReleaseManifest(TypedDict):
    schema_version: Literal[1]
    phase: Literal["candidate", "final"]
    revision: str
    workflow_run_id: int
    aggregate_artifact_name: str
    package_version: Literal["1.0.0"]
    sdist_sha256: str
    artifacts: list[ArtifactRecord]
    smoke_records: list[WheelSmokeRecord]
    performance: PerformanceRecord
```

`revision` must match `[0-9a-f]{40}`; every digest must match
`[0-9a-f]{64}`. The verifier enforces the exact nine filenames/tags for the
selected phase and requires every `source_sdist_sha256` to equal the manifest
source digest. It also requires the exact wheel-hash/runtime/platform smoke
record matrix and rejects a record whose phase, wheel hash, or expected
candidate/final decision contract differs. The performance record must name
the one authoritative phase/SHA-specific report, match its content hash,
revision, phase, exact manylinux x86-64 wheel hash, pinned runner identity,
Python/pandas versions, successful thresholds, top-level workflow run ID, and
aggregate artifact name.

Each platform builder writes one provenance fragment with
`record --phase --revision --sdist-sha256 --artifact --output`. Aggregation
rejects a fragment before copying any distribution if phase, revision,
package version, source hash, filename, or content hash differs.

- [ ] **Step 5: Run local build/provenance verification**

Run in fresh temporary output/extraction directories:

```bash
mx_pack_root="$(mktemp -d)"
mkdir -p "$mx_pack_root/sdist" "$mx_pack_root/fallback-source" "$mx_pack_root/native-source" "$mx_pack_root/fallback" "$mx_pack_root/native"
uv pip install --python .venv/bin/python -e ".[dev]" -r requirements/native-release.txt
.venv/bin/pytest tests/packaging -q
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_pack_root/sdist"
tar -xzf "$mx_pack_root"/sdist/*.tar.gz -C "$mx_pack_root/fallback-source"
tar -xzf "$mx_pack_root"/sdist/*.tar.gz -C "$mx_pack_root/native-source"
mx_fallback_tree="$(find "$mx_pack_root/fallback-source" -mindepth 1 -maxdepth 1 -type d)"
mx_native_tree="$(find "$mx_pack_root/native-source" -mindepth 1 -maxdepth 1 -type d)"
(
  cd "$mx_fallback_tree"
  MESSY_XLSX_BUILD_MODE=fallback /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_pack_root/fallback"
)
(
  cd "$mx_native_tree"
  MESSY_XLSX_BUILD_MODE=native /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_pack_root/native"
)
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_pack_root/native"/*abi3*.whl
```

Expected: both separate clean exact-sdist build modes, artifact unit tests,
native audit, and source inventory pass. Synthetic nine-file fixtures exercise
`record`, `assemble`, and `verify` with exact smoke and authoritative
performance records in `test_release_artifacts.py`.

- [ ] **Step 6: Commit exact-sdist provenance tooling**

```bash
git add build_support.py requirements/native-release.txt setup.py pyproject.toml .gitignore Makefile scripts/release_artifacts.py tests/packaging/test_build_support.py tests/packaging/test_build_modes.py tests/packaging/test_release_artifacts.py
git commit -m "build: verify exact-sdist artifact provenance"
```

Do not stage `CONTINUE.md`, `.superpowers`, generated C/native files, or
`uv.lock`. The gate remains false.

- [ ] **Step 7: Issue revision-bearing provenance only from the clean commit**

Run from the new committed `HEAD`; do not reuse Step 5 artifacts:

```bash
git diff --quiet
git diff --cached --quiet
mx_provenance_sha="$(git rev-parse HEAD)"
mx_provenance_root="$(mktemp -d)"
mkdir -p "$mx_provenance_root/sdist" "$mx_provenance_root/fallback-source" "$mx_provenance_root/native-source" "$mx_provenance_root/fallback" "$mx_provenance_root/native"
MESSY_XLSX_BUILD_MODE=fallback .venv/bin/python -m build --sdist --outdir "$mx_provenance_root/sdist"
tar -xzf "$mx_provenance_root"/sdist/*.tar.gz -C "$mx_provenance_root/fallback-source"
tar -xzf "$mx_provenance_root"/sdist/*.tar.gz -C "$mx_provenance_root/native-source"
mx_provenance_fallback_tree="$(find "$mx_provenance_root/fallback-source" -mindepth 1 -maxdepth 1 -type d)"
mx_provenance_native_tree="$(find "$mx_provenance_root/native-source" -mindepth 1 -maxdepth 1 -type d)"
(
  cd "$mx_provenance_fallback_tree"
  MESSY_XLSX_BUILD_MODE=fallback /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_provenance_root/fallback"
)
(
  cd "$mx_provenance_native_tree"
  MESSY_XLSX_BUILD_MODE=native /home/ivan/Projects/messy-xlsx/.worktrees/parser-v1/.venv/bin/python -m build --wheel --outdir "$mx_provenance_root/native"
)
mx_provenance_sdist="$(find "$mx_provenance_root/sdist" -maxdepth 1 -type f -name '*.tar.gz')"
mx_provenance_sdist_sha="$(sha256sum "$mx_provenance_sdist" | cut -d ' ' -f 1)"
for mx_distribution in "$mx_provenance_sdist" "$mx_provenance_root"/fallback/*.whl "$mx_provenance_root"/native/*abi3*.whl; do
  mx_fragment_name="$(basename "$mx_distribution").json"
  .venv/bin/python scripts/release_artifacts.py record \
    --phase candidate \
    --revision "$mx_provenance_sha" \
    --sdist-sha256 "$mx_provenance_sdist_sha" \
    --artifact "$mx_distribution" \
    --output "$mx_provenance_root/$mx_fragment_name"
done
```

Expected: every fragment names the clean committed revision and one exact
source digest. A failure requires a new corrective commit and a fresh rebuild;
never relabel an existing distribution.

---

### Task 20: Add Resolver, Candidate-Smoke, and Publish Contracts

**Files:**
- Create: `scripts/check_wheel_resolution.py`
- Create: `scripts/smoke_csv_artifact.py`
- Create: `tests/packaging/test_artifact_smoke_cli.py`
- Create: `tests/packaging/test_wheel_resolution.py`
- Create: `tests/packaging/test_publish_contract.py`
- Modify: `.github/workflows/publish.yml`

**Interfaces:**
- Consumes: Task 12's
  `_run_candidate_artifact_smoke(source, plan, metrics)`.
- Proves: native/fallback resolver behavior without network indexes.
- Requires: publishing only a verified exact-SHA final release set.

- [ ] **Step 1: Write failing seam, resolver, and publish-contract tests**

```python
monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", False)
assert public_csv_decision().reason is CSVExecutionReason.PRODUCTION_GATE_DISABLED
with pytest.raises(PermissionError):
    csv_native._prepare_native_streaming_reader(
        source, plan, metrics, candidate_token=object()
    )
prepared = csv_native._run_candidate_artifact_smoke(source, plan, metrics)
try:
    assert prepared.reader.execution_decision.kind is CSVExecutionKind.NATIVE
finally:
    prepared.reader.close()

monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", True)
assert public_csv_decision().kind is CSVExecutionKind.NATIVE
```

The wrapper has no token parameter and internally supplies the module-owned
token. Add negative cases for custom ownership, kill switch, unsupported
runtime, import failure, handshake mismatch, evidence budget, and unsupported
evidence type.

All smoke and benchmark contract tests parameterize `phase` as `candidate` or
`final`; they do not assume the checked-in constant. Candidate requires
private native plus public disabled fallback. Final requires public native and
the exact `"1"` kill-switch fallback. The source-derived workflow phase selects
which prewritten assertions run, so Task 23 changes only the gate line.

Resolver tests use an isolated temporary wheelhouse with
`--no-index --find-links --no-deps`. They prove native preference on each
supported tag, universal fallback on unsupported/free-threaded tags, and ABI3
wheel preference on future supported platform tags while the runtime guard
materializes.

Publish tests reject a false gate, wrong tag/version/changelog/main tip,
candidate namespace, wrong revision, non-nine-file set, manifest inside the
publish directory, missing/renamed/hash-mismatched authoritative performance
evidence, and any publish allowlist entry other than `.whl` or `.tar.gz`.

- [ ] **Step 2: Run the red contracts**

```bash
uv pip install --python .venv/bin/python -e ".[dev]" -r requirements/native-release.txt
.venv/bin/pytest tests/packaging/test_artifact_smoke_cli.py tests/packaging/test_wheel_resolution.py tests/packaging/test_publish_contract.py -q
```

Expected: assertion failures identify absent smoke/resolver commands or an
unverified publish path.

- [ ] **Step 3: Implement isolated smoke and resolver CLIs**

`scripts/smoke_csv_artifact.py --phase candidate --wheel PATH --wheel-sha256
DIGEST --output RECORD` installs one compatible exact wheel by path outside the
repository, runs `pip check`, and emits one wheel-hash/runtime/platform-bound
smoke record. Native wheels assert extension presence/source handshake,
private native parsing, and public disabled fallback. The universal wheel
asserts extension absence and automatic materialization.

`--phase final` repeats direct extension checks, requires public native routing
on supported runtimes, exercises the kill switch, and retains fallback
behavior on universal/unsupported cases.

`scripts/check_wheel_resolution.py` consumes only a statically verified manifest and
wheelhouse. It cannot download or rebuild artifacts.

- [ ] **Step 4: Make publishing consume only verified final artifacts**

While the source gate is false, update `publish.yml` so a future tagged run:

- requires a true source gate;
- invokes the complete quality/artifact workflow on the tagged SHA;
- downloads only `final-${GITHUB_SHA}-release-set`;
- repeats phase, revision, nine-file, hashes, sdist lineage, metadata parity,
  Twine, ABI3, resolver, `pip check`, and the exact uniquely named
  revision/phase/wheel/runner-bound performance-report verification;
- publishes only the eight wheels and one `.tar.gz`, never the manifest;
- retains exact tag/version/changelog/main-tip validation.

- [ ] **Step 5: Run and commit the contracts**

```bash
.venv/bin/pytest tests/packaging/test_artifact_smoke_cli.py tests/packaging/test_wheel_resolution.py tests/packaging/test_publish_contract.py -q
git add scripts/check_wheel_resolution.py scripts/smoke_csv_artifact.py tests/packaging/test_artifact_smoke_cli.py tests/packaging/test_wheel_resolution.py tests/packaging/test_publish_contract.py .github/workflows/publish.yml
git commit -m "ci: verify native resolver smoke and publishing"
```

---

### Task 21: Wire the Seven-Wheel and Aggregate Workflows

**Files:**
- Create: `tests/packaging/test_workflow_contract.py`
- Modify: `tests/packaging/test_ci_run_verifier.py`
- Modify: `scripts/verify_native_ci.py`
- Create: `.github/workflows/native-wheels.yml`
- Create: `.github/workflows/native-artifacts.yml`
- Modify: `.github/workflows/native-safety.yml`
- Modify: `.github/workflows/native-abi.yml`
- Modify: `.github/workflows/native-performance.yml`
- Modify: `.github/workflows/test.yml`

**Interfaces:**
- Produces: one exact source archive, seven native wheels, one fallback wheel,
  provenance fragments, and one manifest outside the release directory.
- Derives: candidate/final phase only from the source-controlled gate.
- Produces: one exact-SHA top-level `native-artifacts.yml` run whose nested job
  graph contains ABI, safety, performance, wheel, resolver, smoke, and
  aggregate acceptance.

- [ ] **Step 1: Write failing workflow-contract tests**

Implement the workflow in three separate reviewed cycles:

| Cycle | Red test nodes | Minimal workflow change | Green command |
|---|---|---|---|
| A — wheel jobs | `test_exact_seven_wheel_matrix`, `test_each_wheel_has_exact_runtime_smokes`, `test_builds_use_clean_exact_sdist` | source job plus seven exclusive wheel jobs and per-wheel smoke records | `pytest tests/packaging/test_workflow_contract.py -q -k "seven_wheel or runtime_smokes or exact_sdist"` |
| B — aggregate provenance | `test_phase_is_source_derived`, `test_aggregate_requires_nine_all_smoke_and_performance_records`, `test_candidate_final_namespaces_are_disjoint` | transactional fragment/performance download, assemble, and static verify jobs | `pytest tests/packaging/test_workflow_contract.py -q -k "phase or aggregate or namespaces or performance"` |
| C — orchestrator/run verification | `test_push_bootstrap_dispatch_call_and_nested_job_graph`, `test_exact_sha_run_ledger`, `test_artifact_record_requires_exact_run_name_id_and_digest`, `test_acceptance_record_cross_binds_ledger_manifest_performance_and_artifact`, `test_test_workflow_requires_orchestrator` | one push-triggered/reusable top-level orchestrator, exact-SHA verifier/artifact/acceptance records, caller gate | `pytest tests/packaging/test_workflow_contract.py tests/packaging/test_ci_run_verifier.py -q` |

Run packaging specification and workflow-security reviews after each cycle.

Parse the YAML and fail unless all actions are commit-pinned, every claimed
matrix leg is nonempty, official builds set `MESSY_XLSX_BUILD_MODE`
explicitly, Linux uses `CIBW_ENVIRONMENT`, Windows uses AMD64 only, each Linux
libc family has an exclusive selector, ABI smokes install exact artifact paths
outside the repository, and callers cannot request `final`.

Require `native-artifacts.yml` to declare `push` for all branches plus
`workflow_dispatch` and `workflow_call`. The push path bootstraps the workflow
before it exists on the default branch and owns the complete nested graph.
Task 21 removes the temporary standalone `push` trigger from `native-abi.yml`;
reusable wheel, ABI, safety, and performance workflows may then be called only
from the top-level graph. `test.yml` requires the orchestrator result but is
not a second source of artifact acceptance.

Test `verify_native_ci.py collect` with recorded `gh run list --json` fixtures.
It accepts exactly one `--workflow native-artifacts.yml`, rejects
older/wrong SHAs, queued/in-progress/cancelled/skipped/neutral runs, duplicate
ambiguous successful runs without deterministic newest selection, missing
nested jobs, and missing required matrix legs. Its output ledger records the
exact revision, top-level database ID/conclusion, and every required nested
job name/conclusion. `print-revision --ledger PATH` emits only the validated
forty-character revision; tests reject a malformed or internally inconsistent
ledger. `accept --ledger PATH --manifest PATH --performance-report PATH
--artifact-record PATH --output PATH` writes a versioned external acceptance
record only when all four inputs agree on revision, phase, report hash,
top-level run ID, and aggregate artifact name.
`print-revision --acceptance PATH` validates that record and emits its SHA.

The post-upload artifact record is exact:

```python
class WorkflowArtifactRecord(TypedDict):
    workflow_run_id: int
    artifact_id: int
    name: str
    digest: str
    expired: Literal[False]
```

`collect-artifact --ledger PATH --name NAME --output PATH` queries the selected
run's artifact API, requires exactly one unexpired artifact with that name and
a `sha256:[0-9a-f]{64}` digest, and atomically writes this external record. The aggregate
artifact digest cannot live inside its own manifest without a circular hash;
the acceptance record binds that post-upload ID/digest to the manifest/report
run ID and artifact name. Mixed same-SHA runs and altered artifact records are
explicit rejection tests.

- [ ] **Step 2: Run the red workflow contract**

```bash
.venv/bin/pytest tests/packaging/test_workflow_contract.py tests/packaging/test_ci_run_verifier.py -q
```

Expected: assertion failures identify absent wheel/aggregate jobs.

- [ ] **Step 3: Implement the exact seven-wheel matrix**

`.github/workflows/native-wheels.yml` builds the source archive once in
fallback mode. Every wheel uses a separate clean extraction of that archive.
Cibuildwheel 4.1.1's pinned image manifest supplies Linux image digests.

| Family | Runner and architecture | `CIBW_BUILD` | Required platform tag |
|---|---|---|---|
| manylinux x86-64 | `ubuntu-24.04`, x86-64, pinned manylinux2014 | `cp311-manylinux_x86_64` | `manylinux_2_17_x86_64.manylinux2014_x86_64` |
| musllinux x86-64 | `ubuntu-24.04`, x86-64, pinned musllinux 1.2 | `cp311-musllinux_x86_64` | `musllinux_1_2_x86_64` |
| manylinux aarch64 | `ubuntu-24.04-arm`, aarch64, pinned manylinux2014 | `cp311-manylinux_aarch64` | `manylinux_2_17_aarch64.manylinux2014_aarch64` |
| musllinux aarch64 | `ubuntu-24.04-arm`, aarch64, pinned musllinux 1.2 | `cp311-musllinux_aarch64` | `musllinux_1_2_aarch64` |
| macOS x86-64 | `macos-15-intel`, target 10.13 | `cp311-macosx_x86_64` | `macosx_10_13_x86_64` |
| macOS arm64 | `macos-15`, target 11.0 | `cp311-macosx_arm64` | `macosx_11_0_arm64` |
| Windows x86-64 | `windows-2025`, `CIBW_ARCHS_WINDOWS=AMD64` | `cp311-win_amd64` | `win_amd64` |

All are `cp311-abi3`. Cibuildwheel tests CPython 3.11; separate jobs install
each exact already-built wheel on CPython 3.12, 3.13, and 3.14. Do not use
`allow-empty`, floating images, `macos-latest`, Windows `auto`, broad Linux
selectors, or unreviewed `test-skip`.

Inside each compatible platform/runtime job, outside the repository:

```bash
python scripts/smoke_csv_artifact.py \
  --phase "$ARTIFACT_PHASE" \
  --wheel "$EXACT_WHEEL_PATH" \
  --wheel-sha256 "$EXACT_WHEEL_SHA256" \
  --output "$SMOKE_RECORD"
```

The universal fallback wheel has its own compatible CPython 3.11–3.14 smoke
jobs. Aggregation never installs foreign wheels.

- [ ] **Step 4: Implement provenance recording and aggregation**

Every builder receives the immutable source archive/hash and runs:

```bash
python scripts/release_artifacts.py record \
  --phase "$ARTIFACT_PHASE" \
  --revision "$GITHUB_SHA" \
  --sdist-sha256 "$SOURCE_SDIST_SHA256" \
  --artifact "$BUILT_DISTRIBUTION" \
  --output "$PROVENANCE_FRAGMENT"
```

`native-artifacts.yml` derives `$ARTIFACT_PHASE` by parsing the source constant.
Candidate namespaces are:

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

Aggregation runs:

```bash
python scripts/release_artifacts.py assemble \
  --phase "$ARTIFACT_PHASE" \
  --revision "$GITHUB_SHA" \
  --workflow-run-id "$GITHUB_RUN_ID" \
  --aggregate-artifact-name "$ARTIFACT_PHASE-$GITHUB_SHA-release-set" \
  --fragments incoming/fragments \
  --smoke-records incoming/smoke-records \
  --performance-report "incoming/performance/$ARTIFACT_PHASE-$GITHUB_SHA-native-csv-performance.json" \
  --artifacts incoming/distributions \
  --dist release-set \
  --manifest "$ARTIFACT_PHASE-manifest.json"
python scripts/release_artifacts.py verify \
  --phase "$ARTIFACT_PHASE" \
  --revision "$GITHUB_SHA" \
  --workflow-run-id "$GITHUB_RUN_ID" \
  --dist release-set \
  --manifest "$ARTIFACT_PHASE-manifest.json" \
  --performance-report "incoming/performance/$ARTIFACT_PHASE-$GITHUB_SHA-native-csv-performance.json"
python scripts/check_wheel_resolution.py \
  --wheelhouse release-set \
  --manifest "$ARTIFACT_PHASE-manifest.json"
```

The release set contains exactly nine distributions; the manifest remains
outside it. The uniquely named authoritative performance report and manifest
are uploaded beside `release-set/` in the SHA-scoped aggregate artifact, not
inside the publishable distribution directory. Candidate and final namespaces
never merge.

- [ ] **Step 5: Wire safety, performance, and aggregate callers**

`native-safety.yml` runs ASan/UBSan, debug allocator, the test-owned allocation
manifest, reentrancy, and lifecycle stress. `native-performance.yml` consumes
the exact manylinux x86-64 artifact. `test.yml` requires ABI, safety,
performance, wheel, resolver, smoke, and aggregate jobs with no empty leg.
`native-artifacts.yml` is the single top-level exact-SHA orchestrator and has
`push`, `workflow_dispatch`, and `workflow_call` triggers. Its nested graph
calls the reusable workflows and exposes one aggregate conclusion only after
every required job and matrix leg succeeds. Before `native-artifacts.yml`
exists on `main`, the authorized feature-branch push starts it from that exact
commit; manual dispatch is used only after the workflow has landed on the
default branch.

`scripts/verify_native_ci.py collect --revision SHA --workflow
native-artifacts.yml --output ledger.json` invokes `gh run list --commit SHA
--workflow native-artifacts.yml --json`, selects only a completed successful
exact-SHA run, verifies the complete required nested job graph through `gh run
view RUN_ID --json jobs`, and writes the immutable ledger. `print-run-id`
reads that ledger and returns the selected top-level ID; `print-revision`
returns the validated exact SHA from a ledger or acceptance record. `accept`
atomically writes the ledger/manifest/report/artifact cross-bound acceptance
record outside the repository.

- [ ] **Step 6: Run and commit workflow contracts**

```bash
.venv/bin/pytest tests/packaging/test_workflow_contract.py tests/packaging/test_ci_run_verifier.py tests/packaging/test_publish_contract.py -q
git add scripts/verify_native_ci.py tests/packaging/test_workflow_contract.py tests/packaging/test_ci_run_verifier.py .github/workflows/native-wheels.yml .github/workflows/native-artifacts.yml .github/workflows/native-safety.yml .github/workflows/native-abi.yml .github/workflows/native-performance.yml .github/workflows/test.yml
git commit -m "ci: wire native CSV candidate artifact matrix"
```

The gate remains false. Do not stage `CONTINUE.md`, `.superpowers`, generated
native files, or `uv.lock`.

---

### Task 22: Build and Accept Disabled Candidate Artifacts

**Files:**
- Consume without modification: `.github/workflows/native-artifacts.yml`
- Consume without modification: `.github/workflows/native-wheels.yml`
- Consume without modification: `.github/workflows/native-safety.yml`
- Consume without modification: `.github/workflows/native-performance.yml`
- Consume without modification: `scripts/release_artifacts.py`

**Interfaces:**
- Proves: every required workflow conclusion, matrix leg, distribution,
  provenance record, and benchmark belongs to the exact false-gate commit.
- Blocks: the gate edit until the downloaded candidate manifest verifies
  locally.

- [ ] **Step 1: Establish a clean immutable candidate revision**

Run:

```bash
git status --short --branch
git diff --quiet
git diff --cached --quiet
mx_candidate_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_candidate_untracked" || test "$mx_candidate_untracked" = "CONTINUE.md"
mx_candidate_sha="$(git rev-parse HEAD)"
git show "$mx_candidate_sha":src/messy_xlsx/parsing/csv_native.py | rg -F '_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False'
```

Expected: only intentionally untracked `CONTINUE.md` is present and the exact
candidate SHA contains the false gate. If the branch has not been pushed,
obtain explicit user authorization before pushing it; this task does not infer
push authority.

- [ ] **Step 2: Accept the push-triggered orchestrator for the exact remote tip**

After the authorized push, run:

```bash
mx_candidate_sha="$(git rev-parse HEAD)"
mx_candidate_branch="$(git branch --show-current)"
git fetch origin "$mx_candidate_branch"
test "$(git rev-parse "origin/$mx_candidate_branch")" = "$mx_candidate_sha"
mx_candidate_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_candidate_sha"
mkdir -p "$mx_candidate_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_candidate_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_candidate_review_dir/candidate-run-ledger.json"
```

The authorized push starts the newly introduced workflow from the exact
feature commit; do not manually dispatch it before it exists on `main`.
`collect` polls that top-level run to completion within its declared timeout.
The verifier programmatically requires `headSha ==
$mx_candidate_sha`, `status == completed`, `conclusion == success`, and the
complete nested ABI, safety, performance, wheel, resolver, smoke, and aggregate
job graph. Missing, cancelled, skipped, neutral, stale-branch, older-SHA, or
incomplete runs block enablement.

- [ ] **Step 3: Download and independently verify the exact candidate manifest**

Use a self-contained shell so no variable or worktree ledger from Step 2 is
required:

```bash
mx_candidate_sha="$(git rev-parse HEAD)"
mx_candidate_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_candidate_sha"
mkdir -p "$mx_candidate_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_candidate_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_candidate_review_dir/candidate-run-ledger.json"
mx_candidate_run="$(
  .venv/bin/python scripts/verify_native_ci.py print-run-id \
    --ledger "$mx_candidate_review_dir/candidate-run-ledger.json" \
    --workflow native-artifacts.yml
)"
mx_candidate_artifact_record="$mx_candidate_review_dir/candidate-artifact.json"
.venv/bin/python scripts/verify_native_ci.py collect-artifact \
  --ledger "$mx_candidate_review_dir/candidate-run-ledger.json" \
  --name "candidate-$mx_candidate_sha-release-set" \
  --output "$mx_candidate_artifact_record"
mx_candidate_download="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-candidate-$mx_candidate_sha-XXXXXX")"
gh run download "$mx_candidate_run" \
  --name "candidate-$mx_candidate_sha-release-set" \
  --dir "$mx_candidate_download"
mx_candidate_performance="$mx_candidate_download/candidate-$mx_candidate_sha-native-csv-performance.json"
test -f "$mx_candidate_performance"
.venv/bin/python scripts/release_artifacts.py verify \
  --phase candidate \
  --revision "$mx_candidate_sha" \
  --workflow-run-id "$mx_candidate_run" \
  --dist "$mx_candidate_download/release-set" \
  --manifest "$mx_candidate_download/candidate-manifest.json" \
  --performance-report "$mx_candidate_performance"
.venv/bin/python scripts/check_wheel_resolution.py \
  --wheelhouse "$mx_candidate_download/release-set" \
  --manifest "$mx_candidate_download/candidate-manifest.json"
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --phase candidate \
  --validate-report "$mx_candidate_performance"
sha256sum "$mx_candidate_download/candidate-manifest.json" "$mx_candidate_performance"
.venv/bin/python scripts/verify_native_ci.py accept \
  --ledger "$mx_candidate_review_dir/candidate-run-ledger.json" \
  --manifest "$mx_candidate_download/candidate-manifest.json" \
  --performance-report "$mx_candidate_performance" \
  --artifact-record "$mx_candidate_artifact_record" \
  --output "$mx_candidate_review_dir/candidate-acceptance.json"
```

Expected: schema version, phase, exact revision, v1.0.0 metadata, nine-file
allowlist, seven ABI3 audits, source-archive lineage, hashes, resolver matrix,
per-wheel private native/public-disabled/fallback smoke records, safety legs,
and benchmark report all verify.

- [ ] **Step 4: Record the candidate acceptance checkpoint**

Attach the one top-level workflow ID, its complete nested-job ledger, candidate
manifest SHA-256, performance-report SHA-256, aggregate artifact ID/digest,
`candidate-acceptance.json`, and local verification output to the external task
review package. No source file or Git commit changes in this task. If any
claimed ABI combination fails, stop and amend the approved design to per-minor
wheels before tokenizer work continues.

---

### Task 23: Enable Native Routing and Rebuild Final Artifacts

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

Re-collect and re-download the accepted candidate in one shell; do not rely on
Task 22 shell state:

```bash
git status --short --branch
git diff --quiet
git diff --cached --quiet
mx_candidate_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_candidate_untracked" || test "$mx_candidate_untracked" = "CONTINUE.md"
mx_candidate_sha="$(git rev-parse HEAD)"
git show "$mx_candidate_sha":src/messy_xlsx/parsing/csv_native.py | rg -F '_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False'
mx_candidate_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_candidate_sha"
mkdir -p "$mx_candidate_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_candidate_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_candidate_review_dir/candidate-run-ledger.json"
mx_candidate_run="$(
  .venv/bin/python scripts/verify_native_ci.py print-run-id \
    --ledger "$mx_candidate_review_dir/candidate-run-ledger.json" \
    --workflow native-artifacts.yml
)"
mx_candidate_artifact_record="$mx_candidate_review_dir/candidate-artifact.json"
.venv/bin/python scripts/verify_native_ci.py collect-artifact \
  --ledger "$mx_candidate_review_dir/candidate-run-ledger.json" \
  --name "candidate-$mx_candidate_sha-release-set" \
  --output "$mx_candidate_artifact_record"
mx_candidate_download="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-candidate-recheck-$mx_candidate_sha-XXXXXX")"
gh run download "$mx_candidate_run" \
  --name "candidate-$mx_candidate_sha-release-set" \
  --dir "$mx_candidate_download"
mx_candidate_performance="$mx_candidate_download/candidate-$mx_candidate_sha-native-csv-performance.json"
test -f "$mx_candidate_performance"
.venv/bin/python scripts/release_artifacts.py verify \
  --phase candidate \
  --revision "$mx_candidate_sha" \
  --workflow-run-id "$mx_candidate_run" \
  --dist "$mx_candidate_download/release-set" \
  --manifest "$mx_candidate_download/candidate-manifest.json" \
  --performance-report "$mx_candidate_performance"
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --phase candidate \
  --validate-report "$mx_candidate_performance"
.venv/bin/python scripts/verify_native_ci.py accept \
  --ledger "$mx_candidate_review_dir/candidate-run-ledger.json" \
  --manifest "$mx_candidate_download/candidate-manifest.json" \
  --performance-report "$mx_candidate_performance" \
  --artifact-record "$mx_candidate_artifact_record" \
  --output "$mx_candidate_review_dir/candidate-recheck-acceptance.json"
```

Expected: the worktree is clean except intentionally untracked `CONTINUE.md`,
the constant is false, the downloaded Task 22 manifest verifies against this
exact SHA, and the review package contains successful exact-SHA IDs for every
candidate, safety, ABI, resolver, artifact, test, and performance gate.

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

Obtain user authorization before pushing the enablement commit. After the
authorized push, the workflow derives `final` from the
source. It must build every wheel from the new final source archive and store
only `final-${GITHUB_SHA}-*` artifacts. Candidate files or manifests are never
merged or reused.

Every exact wheel repeats direct-extension smoke, public native routing on
supported runtimes, fallback behavior, kill-switch behavior, CPython
3.11–3.14 ABI smoke, `abi3audit --strict`, resolver checks, sanitizer/fault
gates, and performance evidence.

Bind acceptance to the exact gate-true SHA:

```bash
mx_final_sha="$(git rev-parse HEAD)"
mx_final_branch="$(git branch --show-current)"
git fetch origin "$mx_final_branch"
test "$(git rev-parse "origin/$mx_final_branch")" = "$mx_final_sha"
mx_final_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_final_sha"
mkdir -p "$mx_final_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_final_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_final_review_dir/final-run-ledger.json"
mx_final_artifact_run="$(
  .venv/bin/python scripts/verify_native_ci.py print-run-id \
    --ledger "$mx_final_review_dir/final-run-ledger.json" \
    --workflow native-artifacts.yml
)"
mx_final_artifact_record="$mx_final_review_dir/final-artifact.json"
.venv/bin/python scripts/verify_native_ci.py collect-artifact \
  --ledger "$mx_final_review_dir/final-run-ledger.json" \
  --name "final-$mx_final_sha-release-set" \
  --output "$mx_final_artifact_record"
mx_final_download="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-final-$mx_final_sha-XXXXXX")"
gh run download "$mx_final_artifact_run" \
  --name "final-$mx_final_sha-release-set" \
  --dir "$mx_final_download"
mx_final_performance="$mx_final_download/final-$mx_final_sha-native-csv-performance.json"
test -f "$mx_final_performance"
.venv/bin/python scripts/release_artifacts.py verify \
  --phase final \
  --revision "$mx_final_sha" \
  --workflow-run-id "$mx_final_artifact_run" \
  --dist "$mx_final_download/release-set" \
  --manifest "$mx_final_download/final-manifest.json" \
  --performance-report "$mx_final_performance"
.venv/bin/python scripts/check_wheel_resolution.py \
  --wheelhouse "$mx_final_download/release-set" \
  --manifest "$mx_final_download/final-manifest.json"
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --phase final \
  --validate-report "$mx_final_performance"
sha256sum "$mx_final_download/final-manifest.json" "$mx_final_performance"
.venv/bin/python scripts/verify_native_ci.py accept \
  --ledger "$mx_final_review_dir/final-run-ledger.json" \
  --manifest "$mx_final_download/final-manifest.json" \
  --performance-report "$mx_final_performance" \
  --artifact-record "$mx_final_artifact_record" \
  --output "$mx_final_review_dir/final-acceptance.json"
```

The authorized push starts the top-level final run. The verifier selects
`$mx_final_artifact_run` only from a completed successful
exact-SHA row and verifies the complete nested job graph. Store the ledger and final
manifest/report hashes in the review package.

If final CI fails, revert the gate to false, fix and verify under a new
disabled candidate revision, then make a fresh one-line enablement commit.
Never patch an enabled revision in place.

---

### Task 24: Run Whole-Repository and Exact-SHA Acceptance

**Files:**
- Consume without modification: the Task 23 exact final revision
- Consume without modification: the final manifest, performance report,
  sanitizer/fuzz output, full test output, and workflow run IDs

**Interfaces:**
- Proves: the complete source tree, tests, documentation, native/fallback
  builds, and final release set are all derived from one unchanged gate-true
  SHA.
- Produces: independent compatibility, native-safety, and release-readiness
  approvals without changing that SHA.

- [ ] **Step 1: Freeze and reverify the exact final revision**

Re-collect and re-download the final set in one self-contained shell:

```bash
git status --short --branch
git diff --quiet
git diff --cached --quiet
mx_accept_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_accept_untracked" || test "$mx_accept_untracked" = "CONTINUE.md"
mx_accept_sha="$(git rev-parse HEAD)"
git show "$mx_accept_sha":src/messy_xlsx/parsing/csv_native.py | rg -F '_NATIVE_CSV_PRODUCTION_READY: Final[bool] = True'
mx_accept_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_accept_sha"
mkdir -p "$mx_accept_review_dir"
.venv/bin/python scripts/verify_native_ci.py collect \
  --revision "$mx_accept_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_accept_review_dir/final-run-ledger.json"
mx_accept_run="$(
  .venv/bin/python scripts/verify_native_ci.py print-run-id \
    --ledger "$mx_accept_review_dir/final-run-ledger.json" \
    --workflow native-artifacts.yml
)"
mx_accept_artifact_record="$mx_accept_review_dir/final-artifact.json"
.venv/bin/python scripts/verify_native_ci.py collect-artifact \
  --ledger "$mx_accept_review_dir/final-run-ledger.json" \
  --name "final-$mx_accept_sha-release-set" \
  --output "$mx_accept_artifact_record"
mx_accept_download="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-final-accept-$mx_accept_sha-XXXXXX")"
gh run download "$mx_accept_run" \
  --name "final-$mx_accept_sha-release-set" \
  --dir "$mx_accept_download"
mx_accept_performance="$mx_accept_download/final-$mx_accept_sha-native-csv-performance.json"
test -f "$mx_accept_performance"
.venv/bin/python scripts/release_artifacts.py verify \
  --phase final \
  --revision "$mx_accept_sha" \
  --workflow-run-id "$mx_accept_run" \
  --dist "$mx_accept_download/release-set" \
  --manifest "$mx_accept_download/final-manifest.json" \
  --performance-report "$mx_accept_performance"
.venv/bin/python scripts/run_native_csv_benchmarks.py \
  --phase final \
  --validate-report "$mx_accept_performance"
.venv/bin/python scripts/verify_native_ci.py accept \
  --ledger "$mx_accept_review_dir/final-run-ledger.json" \
  --manifest "$mx_accept_download/final-manifest.json" \
  --performance-report "$mx_accept_performance" \
  --artifact-record "$mx_accept_artifact_record" \
  --output "$mx_accept_review_dir/final-acceptance.json"
```

Expected: only intentionally untracked `CONTINUE.md` is present, HEAD is the
accepted gate-true SHA, and the final manifest verifies against it.

- [ ] **Step 2: Synchronize the validation environment**

Run:

```bash
mx_accept_head="$(git rev-parse HEAD)"
mx_accept_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_accept_head"
mx_accept_sha="$(
  .venv/bin/python scripts/verify_native_ci.py print-revision \
    --acceptance "$mx_accept_review_dir/final-acceptance.json"
)"
test "$mx_accept_head" = "$mx_accept_sha"
git diff --quiet
git diff --cached --quiet
mx_accept_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_accept_untracked" || test "$mx_accept_untracked" = "CONTINUE.md"
mx_accept_sync_root="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-final-sync-$mx_accept_sha-XXXXXX")"
mkdir -p "$mx_accept_sync_root/source"
git archive "$mx_accept_sha" | tar -x -C "$mx_accept_sync_root/source"
uv pip install --python .venv/bin/python \
  -e "$mx_accept_sync_root/source[dev,docs,all]" \
  -r "$mx_accept_sync_root/source/requirements/native-release.txt"
MESSY_XLSX_BUILD_MODE=native uv pip install --python .venv/bin/python \
  --no-deps --reinstall -e "$mx_accept_sync_root/source"
```

Expected: `build`, Twine, Bandit, pinned native tools, docs tools, and all
optional runtime dependencies are present from the exact accepted extraction;
the native source-hash guard passes before the acceptance suite starts.

- [ ] **Step 3: Run the complete local acceptance gate**

Run:

```bash
mx_accept_head="$(git rev-parse HEAD)"
mx_accept_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_accept_head"
mx_accept_sha="$(
  .venv/bin/python scripts/verify_native_ci.py print-revision \
    --acceptance "$mx_accept_review_dir/final-acceptance.json"
)"
test "$mx_accept_head" = "$mx_accept_sha"
git diff --quiet
git diff --cached --quiet
mx_accept_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_accept_untracked" || test "$mx_accept_untracked" = "CONTINUE.md"
mx_accept_root="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-final-local-$mx_accept_sha-XXXXXX")"
mkdir -p "$mx_accept_root/source" "$mx_accept_root/sdist" "$mx_accept_root/fallback-source" "$mx_accept_root/native-source" "$mx_accept_root/fallback" "$mx_accept_root/native"
git archive "$mx_accept_sha" | tar -x -C "$mx_accept_root/source"
mx_accept_venv="$(pwd)/.venv"
MESSY_XLSX_BUILD_MODE=native uv pip install --python "$mx_accept_venv/bin/python" \
  --no-deps --reinstall -e "$mx_accept_root/source"
(
  cd "$mx_accept_root/source"
  "$mx_accept_venv/bin/ruff" check src/messy_xlsx tests scripts benchmarks
  "$mx_accept_venv/bin/ruff" format --check src/messy_xlsx tests scripts benchmarks
  "$mx_accept_venv/bin/mypy" src/messy_xlsx --ignore-missing-imports
  "$mx_accept_venv/bin/bandit" -q -r src/messy_xlsx
  "$mx_accept_venv/bin/pytest" tests -q --cov=messy_xlsx --cov-report=term-missing --cov-fail-under=75
  "$mx_accept_venv/bin/mkdocs" build --strict --site-dir "$mx_accept_root/site"
  MESSY_XLSX_BUILD_MODE=fallback "$mx_accept_venv/bin/python" -m build --sdist --outdir "$mx_accept_root/sdist"
)
tar -xzf "$mx_accept_root"/sdist/*.tar.gz -C "$mx_accept_root/fallback-source"
tar -xzf "$mx_accept_root"/sdist/*.tar.gz -C "$mx_accept_root/native-source"
mx_accept_fallback_tree="$(find "$mx_accept_root/fallback-source" -mindepth 1 -maxdepth 1 -type d)"
mx_accept_native_tree="$(find "$mx_accept_root/native-source" -mindepth 1 -maxdepth 1 -type d)"
(
  cd "$mx_accept_fallback_tree"
  MESSY_XLSX_BUILD_MODE=fallback "$mx_accept_venv/bin/python" -m build --wheel --outdir "$mx_accept_root/fallback"
)
(
  cd "$mx_accept_native_tree"
  MESSY_XLSX_BUILD_MODE=native "$mx_accept_venv/bin/python" -m build --wheel --outdir "$mx_accept_root/native"
)
"$mx_accept_venv/bin/twine" check "$mx_accept_root"/sdist/* "$mx_accept_root"/fallback/* "$mx_accept_root"/native/*
uvx --from abi3audit==0.0.26 abi3audit --strict "$mx_accept_root"/native/*abi3*.whl
test "$(git rev-parse HEAD)" = "$mx_accept_sha"
git diff --quiet
git diff --cached --quiet
```

- [ ] **Step 4: Run exact native safety, performance, smoke, and artifact commands**

Run this self-contained gate:

```bash
mx_accept_head="$(git rev-parse HEAD)"
mx_accept_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_accept_head"
mx_accept_sha="$(
  .venv/bin/python scripts/verify_native_ci.py print-revision \
    --acceptance "$mx_accept_review_dir/final-acceptance.json"
)"
test "$mx_accept_head" = "$mx_accept_sha"
git diff --quiet
git diff --cached --quiet
mx_accept_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_accept_untracked" || test "$mx_accept_untracked" = "CONTINUE.md"
mx_accept_exact_root="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-final-gate-source-$mx_accept_sha-XXXXXX")"
mkdir -p "$mx_accept_exact_root/source"
git archive "$mx_accept_sha" | tar -x -C "$mx_accept_exact_root/source"
mx_accept_venv="$(pwd)/.venv"
MESSY_XLSX_BUILD_MODE=native uv pip install --python "$mx_accept_venv/bin/python" \
  --no-deps --reinstall -e "$mx_accept_exact_root/source"
mx_accept_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_accept_sha"
mkdir -p "$mx_accept_review_dir"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/verify_native_ci.py" collect \
  --revision "$mx_accept_sha" \
  --workflow native-artifacts.yml \
  --output "$mx_accept_review_dir/final-run-ledger.json"
mx_accept_run="$(
  "$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/verify_native_ci.py" print-run-id \
    --ledger "$mx_accept_review_dir/final-run-ledger.json" \
    --workflow native-artifacts.yml
)"
mx_accept_artifact_record="$mx_accept_review_dir/final-artifact.json"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/verify_native_ci.py" collect-artifact \
  --ledger "$mx_accept_review_dir/final-run-ledger.json" \
  --name "final-$mx_accept_sha-release-set" \
  --output "$mx_accept_artifact_record"
mx_accept_download="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-final-gates-$mx_accept_sha-XXXXXX")"
gh run download "$mx_accept_run" \
  --name "final-$mx_accept_sha-release-set" \
  --dir "$mx_accept_download"
mx_accept_performance_report="$mx_accept_download/final-$mx_accept_sha-native-csv-performance.json"
test -f "$mx_accept_performance_report"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/release_artifacts.py" verify \
  --phase final \
  --revision "$mx_accept_sha" \
  --workflow-run-id "$mx_accept_run" \
  --dist "$mx_accept_download/release-set" \
  --manifest "$mx_accept_download/final-manifest.json" \
  --performance-report "$mx_accept_performance_report"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/check_wheel_resolution.py" \
  --wheelhouse "$mx_accept_download/release-set" \
  --manifest "$mx_accept_download/final-manifest.json"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/run_native_csv_benchmarks.py" \
  --phase final \
  --validate-report "$mx_accept_performance_report"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/verify_native_ci.py" accept \
  --ledger "$mx_accept_review_dir/final-run-ledger.json" \
  --manifest "$mx_accept_download/final-manifest.json" \
  --performance-report "$mx_accept_performance_report" \
  --artifact-record "$mx_accept_artifact_record" \
  --output "$mx_accept_review_dir/final-acceptance.json"
(
  cd "$mx_accept_exact_root/source"
  "$mx_accept_venv/bin/python" scripts/run_native_csv_fuzz.py \
  --c-seed 0x0C5A14 \
  --python-seed 0xBADC5EED \
  --examples 5000 \
  --timeout 300
  bash scripts/run_native_csv_sanitizers.sh
  PYTHONMALLOC=debug "$mx_accept_venv/bin/pytest" tests/native_csv -q
  "$mx_accept_venv/bin/pytest" tests/test_performance/test_native_csv_contract.py -q
)
mx_accept_native_wheel="$(
  find "$mx_accept_download/release-set" -maxdepth 1 -type f \
    -name '*manylinux_2_17_x86_64.manylinux2014_x86_64.whl' -print -quit
)"
mx_accept_fallback_wheel="$(
  find "$mx_accept_download/release-set" -maxdepth 1 -type f \
    -name '*-py3-none-any.whl' -print -quit
)"
test -n "$mx_accept_native_wheel"
test -n "$mx_accept_fallback_wheel"
mx_accept_native_sha="$(sha256sum "$mx_accept_native_wheel" | cut -d' ' -f1)"
mx_accept_fallback_sha="$(sha256sum "$mx_accept_fallback_wheel" | cut -d' ' -f1)"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/smoke_csv_artifact.py" \
  --phase final \
  --wheel "$mx_accept_native_wheel" \
  --wheel-sha256 "$mx_accept_native_sha" \
  --output "$mx_accept_review_dir/local-native-smoke.json"
"$mx_accept_venv/bin/python" "$mx_accept_exact_root/source/scripts/smoke_csv_artifact.py" \
  --phase final \
  --wheel "$mx_accept_fallback_wheel" \
  --wheel-sha256 "$mx_accept_fallback_sha" \
  --output "$mx_accept_review_dir/local-fallback-smoke.json"
test "$(git rev-parse HEAD)" = "$mx_accept_sha"
git diff --quiet
git diff --cached --quiet
```

Expected: all checks pass with no production pandas-chunk reader, no duplicate
Python full pass, no public API change, and no ownership/memory regression.

- [ ] **Step 5: Obtain independent compatibility, safety, and release-readiness reviews**

Give each reviewer the approved design, this plan, the complete diff, focused
and full test output, sanitizer/fuzz reports, performance JSON, candidate/final
manifests, CI URLs, `final-acceptance.json`, and the exact-SHA workflow ledger.
Include the post-upload artifact ID/digest record. Each approval names that
accepted revision, workflow run ID, and artifact digest. Resolve every blocker
with a regression test and repeat the affected gate. Do not mark Task 14
complete on reviewer promises or partial CI.

- [ ] **Step 6: Close the exact-SHA acceptance checkpoint without source changes**

Attach full local output, exact workflow IDs, candidate/final manifests,
performance JSON, fuzz seeds/results, sanitizer logs, clean-install smoke, and
all three independent approvals to the implementation review package. Run:

```bash
mx_accept_head="$(git rev-parse HEAD)"
mx_accept_review_dir="${TMPDIR:-/tmp}/messy-xlsx-native-review-$mx_accept_head"
mx_accept_sha="$(
  .venv/bin/python scripts/verify_native_ci.py print-revision \
    --acceptance "$mx_accept_review_dir/final-acceptance.json"
)"
mx_accept_ledger_sha="$(
  .venv/bin/python scripts/verify_native_ci.py print-revision \
    --ledger "$mx_accept_review_dir/final-run-ledger.json"
)"
test "$mx_accept_head" = "$mx_accept_sha"
test "$mx_accept_ledger_sha" = "$mx_accept_sha"
git show "$mx_accept_sha":src/messy_xlsx/parsing/csv_native.py | rg -F '_NATIVE_CSV_PRODUCTION_READY: Final[bool] = True'
git diff --quiet
git diff --cached --quiet
mx_accept_untracked="$(git ls-files --others --exclude-standard)"
test -z "$mx_accept_untracked" || test "$mx_accept_untracked" = "CONTINUE.md"
git status --short --branch
```

Expected: HEAD is still the independently accepted gate-true revision and only
intentionally untracked `CONTINUE.md` is present.

Do not edit, commit, tag, or publish in this task. Parent Task 20 later updates
README/docs/changelog and must rebuild/reverify a new final set on that exact
release SHA before any separately authorized `v1.0.0` tag.

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
  no decision; custom handlers keep `ReaderDecision.backend ==
  CUSTOM_DATAFRAME` and record `CSVExecutionKind.CUSTOM_SPI`.
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
- [ ] After every `.pyx` edit, force a native reinstall before green tests and
  require `NATIVE_SOURCE_SHA256` to equal the checked-in `.pyx` hash. A stale
  extension invalidates the test run.
- [ ] Update this tracker and the parent Task 14 ledger only from verified
  evidence. For Tasks 22–24, record progress in the external review package
  without creating an intervening Git commit; parent Task 20 incorporates the
  final checkbox/ledger update into its release-documentation commit and then
  rebuilds the complete final matrix on that new exact SHA. Parent Task 20
  remains responsible for v1.0.0 documentation, changelog, verifying version
  metadata, that release-SHA rebuild, and the separately authorized
  tag/publication action.
