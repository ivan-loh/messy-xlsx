# Architecture Refactor Plan

This is the single source of truth for simplifying and speeding up `messy-xlsx`
without changing its supported behavior. Keep status, decisions, measurements, and
follow-up notes in this file.

Last updated: 2026-07-21

## Objectives

- Preserve public behavior for XLSX/XLSM, legacy XLS, CSV/TSV, paths, and buffers.
- Make every public parsing path use the same configuration, detection, parsing,
  normalization, and error contracts.
- Reduce repeated file reads, workbook opens, DataFrame copies, and per-cell work.
- Turn extension points such as handler registration and formula backends into real,
  testable seams.
- Keep each slice independently reviewable, reversible, and releasable.

## Non-goals

- Removing supported formats or configuration options.
- Changing normalized output merely to simplify implementation.
- Replacing pandas, openpyxl, fastexcel, or xlrd in one large rewrite.
- Removing a public API without a separately approved deprecation cycle.
- Combining unrelated cleanup into a performance or correctness slice.

## Current baseline

| Measure | Baseline |
|---|---:|
| Full suite | 880 passed, 0 skipped |
| Expected warnings | 10 |
| Statement coverage | 79% (last measured) |
| Ruff | Clean across repository |
| Ruff formatting | Clean across repository |
| Mypy | Clean across `src/` |
| Full-suite time | About 67 seconds |
| XLSX, 1,000 rows | About 25.5 ms |
| XLSX, 1,000 rows, no normalization | About 3.9 ms |
| CSV, 1,000 rows | About 23.4 ms |
| XLSX with merged cells, 100 rows | About 17.0 ms |
| Sheet-name lookup | About 0.40 ms after benchmark warm-up |

Benchmark values are local reference points, not universal promises. Compare changes
on the same machine, Python version, dependency versions, and fixtures. Record medians
and distributions rather than relying on one run.

## Working protocol

Every slice follows the same characterization-first flow:

1. Mark exactly one slice `IN PROGRESS` in the progress table.
2. Write or identify a test that protects the behavior being changed.
3. Demonstrate the test fails for a real defect, or passes as a characterization lock
   before a behavior-preserving refactor.
4. Make the smallest production change that completes the slice.
5. Run focused tests while iterating.
6. Run the full test, lint, format, typing, and diff-integrity gates.
7. Run before/after benchmarks for any hot-path slice.
8. Update this file with the result, measurements, decisions, and next slice.

No slice is complete until its acceptance criteria and global gates pass.

## Global gates

Run these before completing every slice:

```bash
pytest -o addopts='' -q
ruff check .
ruff format --check .
mypy src
git diff --check
```

Additional gates:

- Public-output refactor: run API parity and integration workflow tests.
- Parser/backend refactor: run path and `BytesIO` tests for every affected format.
- Performance refactor: run `pytest tests/test_performance -v --benchmark-only` before
  and after the change.
- Documentation or API change: run `mkdocs build --strict`.
- Packaging change: build wheel and sdist, then test installation in a clean environment.

## Progress

Status values: `DONE`, `IN PROGRESS`, `PLANNED`, `BLOCKED`, or `DEFERRED`.

| ID | Slice | Status | Depends on | Completion evidence |
|---|---|---|---|---|
| S00 | Baseline, test dependencies, and quality gates | DONE | — | `xlwt` installed; 723 tests; 0 skipped; lint/type checks clean |
| S01 | Public correctness and documentation alignment | DONE | S00 | API exports, `column_count`, immutable table config, dependency and formula docs fixed |
| S02 | Multi-sheet pipeline convergence and initial fast path | DONE | S01 | Multi-sheet parsing delegates to `MessyWorkbook`; safe fastexcel path enabled |
| S03 | Architecture characterization suite | DONE | S02 | API parity, path/buffer, multi-sheet, cache, formula, registry, XLS contracts added |
| S04 | Cache, formula boundary, and registry correctness | DONE | S03 | Config-aware cache; true cached formula values; injectable compatible handlers |
| S05 | Resource lifecycle and convenience-function closure | DONE | S04 | 26 lifecycle contracts; 749 full-suite tests; no target-FD growth |
| S06 | Pure parse-plan compilation | DONE | S05 | 72 parse-decision contracts; 821 full-suite tests; differential parity sweep clean |
| S07 | Unified source and buffer abstraction | DONE | S05 | 58 source contracts; 880 full-suite tests; path/buffer/read-once parity; no material benchmark regression |
| S08 | Workbook session and repeated-open reduction | IN PROGRESS | S06, S07 | Open/load-count characterization is next |
| S09 | Shared header-detection contract | PLANNED | S06 | — |
| S10 | Structure analyzer decomposition | PLANNED | S08, S09 | — |
| S11 | Thin multi-sheet policy layer | PLANNED | S09, S10 | — |
| S12 | Explicit handler/backend strategy | PLANNED | S06, S08 | — |
| S13 | CSV read-once and metadata-detection consolidation | PLANNED | S07, S12 | — |
| S14 | Legacy XLS parity and optional-capability boundary | PLANNED | S09, S12 | — |
| S15 | Normalization pipeline pass and copy reduction | PLANNED | S06 | — |
| S16 | Formula backend adapters and buffer support | PLANNED | S07, S08 | — |
| S17 | Bulk cell/range access | PLANNED | S08, S16 | — |
| S18 | Typed table and structure models | PLANNED | S10 | — |
| S19 | Cache policy, identity, and concurrency hardening | PLANNED | S07, S10 | — |
| S20 | Error taxonomy and diagnostics | PLANNED | S12, S13, S14, S16 | — |
| S21 | Public API typing, deprecations, and lifecycle polish | PLANNED | S11, S18, S20 | — |
| S22 | CI matrix, performance budgets, and release hardening | PLANNED | S21 | — |

## Dependency map

```text
Completed foundation
S00 -> S01 -> S02 -> S03 -> S04
                            |
                            v
Safety and architecture seams
S05 -> S06 -----------+-----------------> S09 -----------+
  |                   |                    |              |
  +----> S07 ---------+----> S08 ----------+----> S10 ----+----> S11
                         |                 |              |
                         +----> S12 -------+----> S14     +----> S18
                         |          |
                         |          +----> S13
                         |
                         +----> S16 ----> S17

Data-path optimization
S06 -----------------------> S15

Cross-cutting hardening
S07 + S10 -----------------> S19
S12 + S13 + S14 + S16 ----> S20
S11 + S18 + S20 ----------> S21 ----> S22
```

## Slice details

### S00 — Baseline, test dependencies, and quality gates

Goal: establish a trustworthy starting point before architecture changes.

Completed work:

- Installed and recorded `xlwt` as a development dependency.
- Activated legacy XLS generation tests instead of skipping them.
- Established the full-suite, coverage, lint, formatting, typing, and benchmark baseline.
- Fixed the arbitrary-text property test so leading `=` strings are stored as text,
  rather than accidentally serialized by openpyxl as formulas.

Acceptance criteria:

- Full suite has no unexpected skip.
- All supported format fixtures execute locally.
- Baseline commands and measurements are recorded in this file.

### S01 — Public correctness and documentation alignment

Goal: remove known public-contract drift before internal refactoring.

Completed work:

- Added the documented `SheetInfo.column_count` alias.
- Exported `read_excel`, `read_excel_tables`, and `analyze_structure` through `__all__`.
- Prevented `MessyTable.to_dataframe()` from mutating caller-owned `SheetConfig`.
- Aligned README dependencies with packaging metadata.
- Documented DataFrame formula behavior separately from cell-level formula evaluation.
- Carried multi-sheet normalization options through `SheetConfig`.

Acceptance criteria:

- README examples execute.
- Public export tests protect `__all__`.
- Caller configuration is unchanged after table parsing.
- Documentation and package requirements agree.

### S02 — Multi-sheet pipeline convergence and initial fast path

Goal: eliminate the largest duplicate public parsing pipeline.

Completed work:

- Routed multi-sheet reads through `MessyWorkbook`.
- Removed duplicate column cleaning, missing-value handling, number normalization, and
  type-consistency implementations from `MultiSheetParser`.
- Reused structure detection for OOXML and a shared legacy header helper for XLS.
- Enabled fastexcel when analysis confirms merged and hidden-cell handling is unnecessary.

Acceptance criteria:

- Single-sheet and multi-sheet outputs agree for equivalent configurations.
- XLS, XLSX, path, and buffer behavior remain covered.
- Performance does not regress on clean XLSX files.

### S03 — Architecture characterization suite

Goal: freeze intended behavior at public and architectural boundaries.

Completed work:

- Added parity tests for convenience, workbook, sheet, and multi-sheet APIs.
- Added path-versus-buffer parity tests.
- Added cache-configuration isolation tests.
- Added formula expression, missing-cache, and populated-cache contracts.
- Added custom registry injection and compatible/incompatible fallback contracts.
- Added direct formula mode contracts.

Acceptance criteria:

- Tests fail when any protected boundary is intentionally broken.
- Tests compare values, columns, dtypes, row order, and error shape where relevant.

### S04 — Cache, formula boundary, and registry correctness

Goal: fix defects exposed by S03 and establish the first real extension seam.

Completed work:

- Included analysis-affecting header patterns in structure-cache identity.
- Applied workbook `SheetConfig` when `get_structure()` is called.
- Separated formula expressions from cached results with formula-preserving and
  data-only workbook views.
- Lazily initialized formula evaluation only for formula cells.
- Closed both workbook views during cleanup.
- Allowed caller-supplied handlers and detectors through `HandlerRegistry`.
- Restricted fallback to handlers declaring support for the detected format.

Acceptance criteria:

- Cached formula values and formula expressions are never conflated.
- Different detection configurations do not share invalid cached structure.
- Custom handlers are usable without module-global mutation.

### S05 — Resource lifecycle and convenience-function closure

Goal: guarantee deterministic cleanup on success and failure.

Completed work:

- Wrapped `read_excel`, `read_excel_tables`, and `analyze_structure` operations in
  `MessyWorkbook` context managers.
- Made `MessyWorkbook.close()` idempotent, best-effort across both formula workbook views,
  and compatible with the existing primary-error precedence.
- Made XLSX sheet-name fallback and XLS `ExcelFile` acquisition exception-safe.
- Closed library-created CSV text streams on successful and failed parse/detection paths.
- Kept caller-owned XLSX, XLS, and CSV buffers open and reusable.
- Updated README and getting-started examples to demonstrate context-managed ownership.

Verification:

- Added 26 deterministic lifecycle contracts covering success, failure, close errors,
  idempotency, handler/analyzer transients, caller buffers, and target-file descriptors.
- Full suite: 749 passed, 0 skipped, with 10 expected third-party/parser warnings.
- Repeated convenience calls showed no target-file descriptor growth on Linux.
- Ruff, formatting, mypy, and diff-integrity gates passed.

Deferred by design:

- Exact buffer-position restoration and non-seekable input policy were completed in S07.
- Workbook/session reuse and open-count reduction remain in S08.
- CSV acquisition/read consolidation remains in S13.
- Formula-backend temporary resource ownership remains in S16.

Acceptance criteria:

- [x] Every identified library-owned resource has deterministic ownership and cleanup.
- [x] Convenience-function output is unchanged.
- [x] Failure-path resource tests pass.
- [x] Caller-owned buffers remain open and reusable.
- [x] Repeated convenience calls do not grow target-file descriptors.

### S06 — Pure parse-plan compilation

Goal: separate configuration decisions from I/O and DataFrame transformation.

Completed work:

- Added a private frozen, slotted `ParsePlan` and pure compiler from configuration,
  optional structure evidence, and format to stable handler/normalization projections.
- Kept analysis and source rewinding in `MessyWorkbook`, while moving header/footer,
  backend-hint, formula-data, locale, and normalization decisions out of orchestration.
- Removed `_apply_structure_detection()` and `_locale_to_separators()` from the workbook,
  including the manual reconstruction of every `SheetConfig` field.
- Kept `ParseOptions` as the handler boundary and return a fresh projection for every parse.
- Snapshotted configuration containers without changing identity-sensitive arbitrary
  condition payloads.
- Unified the sanitize/rename path while preserving the existing transformation order and
  the legacy `normalize=False` row-filter bypass.
- Made `MessySheet.structure` honor workbook detection configuration and made
  `MessyTable.to_dataframe()` inherit workbook configuration by default.
- Corrected multi-table footer math to ignore hidden rows outside the data region.
- Prevented automatic sheet-global footer evidence from truncating a caller-bounded range;
  an explicit caller `skip_footer` still wins, and an empty range keeps full-sheet behavior.

Verification:

- Added 72 parse-decision and public-boundary contracts covering every header mode and
  fallback, threshold equality, hidden rows, footer precedence, ranges, formula-data mode,
  backend hints, text auto-header behavior, locale/separator rules, normalization toggles,
  plan freezing, fresh projections, collection snapshots, and enum/string parity.
- A differential sweep across valid header, footer, hidden-content, locale, separator, and
  backend-hint combinations found no unintended behavior drift.
- Focused architecture/configuration/parser/integration gate: 242 passed.
- Full suite: 821 passed, 0 skipped, with 10 expected third-party/parser warnings.
- Ruff, formatting, mypy, and diff-integrity gates passed.

Deferred by design:

- CSV manual-mode header convergence and multi-row headers remain in S09.
- Range/backend capability semantics beyond global-footer safety remain in S12 and S18.
- The legacy `normalize=False` row-filter bypass and partial-separator application defect
  remain in S15.
- Source position, ownership, and non-seekable behavior were completed in S07.

Acceptance criteria:

- [x] Plan compilation has no file or DataFrame access.
- [x] Workbook orchestration is visibly shorter and has one configuration-decision boundary.
- [x] Existing public outputs and errors remain identical except for characterized bug fixes.
- [x] Caller configuration and structure inputs are not mutated.
- [x] Every `SheetConfig` field has an explicit compiler, analysis-only, or deferred disposition.

### S07 — Unified source and buffer abstraction

Goal: centralize path/buffer identity, rewinding, byte access, and ownership.

Completed work:

- Added one private `SourceHandle` for normalized path identity, explicit or stream-carried
  names, caller ownership, repeatable binary borrows, stable identity, and lifecycle state.
- Borrowed every seekable stream from byte zero and restored the cursor position received by
  each operation on success and parser failure; source restoration errors remain actionable.
- Snapshotted non-seekable streams once, rejected already-consumed read-once streams, and
  retained caller ownership while providing repeatable internal views.
- Kept paths as paths, memoized complete immutable bytes only for byte-only backends, and
  created detached streams only for persistent openpyxl views.
- Normalized `bytearray` and `memoryview` reads only at strict backend boundaries, including
  streams whose zero-length reads return ordinary `bytes`.
- Migrated workbook orchestration, detection, analysis, registry routing, and built-in XLSX,
  XLS, and CSV handlers to the shared source boundary.
- Gave handler retries and legacy custom detectors, handlers, and registry subclasses fresh
  raw borrows; extensions can explicitly opt into the internal handle contract.
- Kept cached-value structure analysis separate from formula-expression detection so
  uncached formulas cannot change structural evidence.
- Replaced repeated read-only formula cell access with one sequential scan and bounded CSV
  stream validation to 1 KiB.

Verification:

- Added 58 direct source contracts and ran 94 source/structure/lifecycle contracts covering
  paths, `BytesIO`, named and invalid names, non-zero cursors, non-seekable sources, repeated
  operations, acquisition failures, fallbacks, legacy extensions, and persistent views.
- Exercised XLSX, XLS, and CSV across seekable/read-once inputs and `bytes`, `bytearray`, and
  `memoryview` reads, including success, parser failure, and simultaneous consumer/restore
  failures.
- Deterministic counters protect one-time snapshots, memoized byte reads, bounded validation,
  fresh fallback borrows, cursor restoration, and caller ownership.
- Full suite: 880 passed, 0 skipped, with 10 expected third-party/parser warnings in 66.77s.
- Ruff, formatting, mypy, diff-integrity, and strict MkDocs gates passed.
- Official benchmarks remained within run-to-run noise; forced cold 1,000-row analysis was
  43.33ms from a path and 42.87ms from a seekable buffer.

Deferred by design:

- Workbook sessions and repeated-open reduction remain in S08.
- Explicit backend selection remains in S12.
- CSV read-once parsing and metadata consolidation remain in S13.
- External formula-engine buffer and temporary-resource adapters remain in S16.
- Source/cache identity and concurrency policy remain in S19.
- Legacy helper deprecation and public input typing remain in S21.
- Cold and buffer benchmark budgets remain in S08 and S22.

Acceptance criteria:

- [x] No ad hoc source-type checks remain outside the source module and adapters.
- [x] Buffers behave consistently across formats and public APIs.
- [x] A source is copied only when required by backend limitations.

### S08 — Workbook session and repeated-open reduction

Goal: reduce OOXML detection, validation, analysis, and parsing from separate opens to a
coordinated session where backend constraints allow it.

Current issue:

- Workbook construction may detect format, read sheet names, validate, analyze structure,
  parse data, and load formula views through separate readers.
- Multi-sheet analysis can parse raw data and then parse each selected sheet again.

Test first:

- Instrument open/load counts for construction, one-sheet parse, all-sheet parse,
  structure access, and formula cell access.
- Preserve path/buffer parity, merged cells, hidden cells, macros, and formulas.
- Benchmark cold and warm single-sheet and multi-sheet workflows.

Implementation direction:

- Add a library-owned `WorkbookSession` coordinating metadata and open workbook views.
- Reuse safe metadata and workbook handles; do not force one backend to emulate another.
- Preserve fastexcel for clean-data fast paths and openpyxl for formatting-sensitive paths.

Acceptance criteria:

- Open counts are explicitly tested and lower on common XLSX workflows.
- Relevant benchmark median improves by at least 20%, or the slice records why a smaller
  measured improvement is the safe limit.
- Memory and file-descriptor use remain bounded.

### S09 — Shared header-detection contract

Goal: make header decisions consistent without pretending all formats expose identical
raw information.

Current issue:

- OOXML structure analysis, legacy XLS DataFrame detection, and CSV metadata detection use
  different scoring implementations and coordinate conventions.
- Header row numbering switches between zero-based and one-based values at boundaries.

Test first:

- A shared corpus covering metadata, sparse headers, numeric-looking headers, dates,
  multi-row headers, blank rows, patterns, hidden rows, and low-confidence fallback.
- Run semantically equivalent XLSX, XLS, and CSV fixtures through public APIs.

Implementation direction:

- Define `HeaderCandidate` and `HeaderDecision` internal models with explicit coordinate
  systems and confidence.
- Extract shared row profiling and scoring.
- Keep format adapters for merged cells, hidden rows, and CSV metadata evidence.

Acceptance criteria:

- Equivalent inputs select equivalent headers across formats where capabilities match.
- Format-specific evidence remains supported.
- Existing public row numbering stays compatible.

### S10 — Structure analyzer decomposition

Goal: replace one large analyzer with independently testable detectors while retaining a
single workbook scan strategy.

Current issue:

- `StructureAnalyzer` owns region, merged, hidden, header, metadata, table, locale, blank,
  sparse, formula, and footer detection.
- Several methods catch broad exceptions because worksheet capability boundaries are not
  explicit.

Test first:

- Component-level tests for each detector plus end-to-end `StructureInfo` snapshots.
- Malformed, empty, sparse, large, hidden, merged, multi-table, and formula-heavy sheets.
- Assert analyzer does not rescan beyond documented limits.

Implementation direction:

- Extract pure detectors operating on a shared sampled sheet profile.
- Gather cell evidence once where practical.
- Compose results into `StructureInfo` in one orchestrator.

Acceptance criteria:

- Each detector has one responsibility and explicit input/output types.
- Broad exception handling is reduced to adapter boundaries.
- Analyzer output and performance remain stable or improve.

### S11 — Thin multi-sheet policy layer

Goal: leave `MultiSheetParser` responsible only for sheet selection policy.

Current issue:

- Although parsing now delegates to `MessyWorkbook`, multi-sheet analysis still performs a
  raw DataFrame parse to classify every sheet and then reparses selected sheets.
- Pivot, empty, minimum-size, and caller-filter policy is mixed with acquisition.

Test first:

- Protect selection order, explicit sheet lists, pivot handling, empty/small sheets,
  custom filters, per-sheet errors, and all current option mappings.
- Count parses and workbook opens for all-sheet workflows.

Implementation direction:

- Introduce a pure `SheetSelectionPolicy` consuming sheet summaries.
- Reuse structure/session summaries rather than parsing complete sheets for classification
  when enough evidence exists.
- Keep compatibility wrappers for `MultiSheetParser`, `read_all_sheets`, and `analyze_excel`.

Acceptance criteria:

- Selected outputs and skip reasons remain compatible.
- A selected sheet is not fully parsed twice in the normal path.
- Multi-sheet code contains no normalization or column-cleaning implementation.

### S12 — Explicit handler/backend strategy

Goal: make format routing and fast/fidelity backend selection declarative and observable.

Current issue:

- `XLSXHandler` computes backend eligibility from several `ParseOptions` flags.
- Registry fallback handles errors but does not expose why a backend was selected.
- Handler methods independently detect, validate, list sheets, and parse.

Test first:

- Backend selection matrix for merge strategies, hidden content, ranges, formulas, macros,
  paths, and buffers.
- Compatible custom fallback and fatal-exception propagation.
- Exact error context when every compatible handler fails.

Implementation direction:

- Add a small internal capability/selection model such as `BackendRequirements` and
  `BackendDecision`.
- Separate format-handler selection from backend selection inside a handler.
- Preserve custom handler registration and priority.

Acceptance criteria:

- Backend decisions are pure and unit-tested.
- Fast paths never silently drop requested fidelity.
- Error context lists only handlers/backends actually attempted.

### S13 — CSV read-once and metadata-detection consolidation

Goal: avoid duplicated path/text implementations and repeated CSV reads.

Current issue:

- Metadata detection has separate path and text methods with nearly identical logic.
- Encoding detection, delimiter detection, metadata sampling, validation, and parsing may
  read the source separately.
- Encoding fallback repeats parse setup and has duplicated exception paths.

Test first:

- UTF-8 BOM, UTF-16, Latin-1, Windows-1252, delimiters, quoted fields, malformed rows,
  metadata, buffers, empty files, and inconsistent columns.
- Instrument bytes read and parser attempts.

Implementation direction:

- Build one decoded sample/input preparation step.
- Share one metadata profiler for path and buffer content.
- Centralize `read_csv` option construction and encoding attempts.

Acceptance criteria:

- Path and buffer outputs/errors match.
- Common CSV paths read/sample once before pandas parsing.
- Existing malformed-row warning behavior is deliberately preserved or explicitly revised.

### S14 — Legacy XLS parity and optional-capability boundary

Goal: keep XLS support optional in production while fully tested in development.

Current issue:

- XLS behavior relies on pandas/xlrd fallback paths and a format-specific header helper.
- Optional dependency errors and fallback semantics can be clearer.

Test first:

- Public convenience, workbook, sheet, multi-sheet, path, buffer, multi-row header,
  skip/footer/range limitations, corruption, and missing-`xlrd` behavior.
- Parity fixtures paired with XLSX where capabilities match.

Implementation direction:

- Define the XLS capability boundary explicitly.
- Use the shared header contract from S09.
- Consolidate xlrd opening and cleanup through the handler/session boundary.

Acceptance criteria:

- Development suite runs XLS tests without skip.
- Production install without `xls` extra fails with one actionable error.
- Supported XLS results match equivalent XLSX results.

### S15 — Normalization pipeline pass and copy reduction

Goal: reduce whole-DataFrame passes and copies while preserving exact normalized output.

Current issue:

- Individual normalizers copy DataFrames or Series independently.
- Missing values, whitespace, numbers, dates, type coercion, sanitization, renames, regex
  drops, and condition drops can each traverse the data.
- A thousands-only separator override is overwritten when the number normalizer detects a
  decimal separator; partial explicit separator configuration is therefore ineffective.
- `normalize=False` also bypasses regex and condition row filters because of the current
  early-return boundary.

Test first:

- Golden DataFrame tests with values, columns, dtypes, null representations, and indexes.
- Mutation tests proving caller input remains unchanged where promised.
- Benchmarks for wide, tall, string-heavy, numeric-heavy, and mixed-type frames.

Implementation direction:

- Make ownership explicit: one defensive copy at the pipeline boundary when needed.
- Fuse compatible column operations and avoid no-op copies.
- Compile row-drop predicates once and apply one final mask.
- Keep normalizers independently testable.

Acceptance criteria:

- Golden outputs and Arrow/BigQuery compatibility remain unchanged.
- Relevant normalization benchmark improves by at least 20% without excessive peak memory.
- No public caller-owned DataFrame or config mutation is introduced.

### S16 — Formula backend adapters and buffer support

Goal: isolate optional evaluators and make every formula mode predictable for paths and
buffers.

Current issue:

- `FormulaEngine` directly imports and coordinates optional `xlcalculator` and `formulas`
  libraries.
- External evaluators are path-oriented; buffer loading currently has no complete adapter.
- Optional backend branches remain among the least-covered code.

Test first:

- Disabled, cached-only, fallback, and always-evaluate modes with present/missing caches.
- Paths and buffers, supported and unsupported functions, circular references, depth,
  backend load failure, and resource cleanup.
- Contract tests using fake adapters; optional integration tests when extras are installed.

Implementation direction:

- Introduce a `FormulaBackend` protocol and one adapter per optional library.
- Let the source/session layer provide a temporary path only when a backend requires it.
- Make backend priority and fallback explicit.

Acceptance criteria:

- Core formula mode tests do not depend on optional third-party libraries.
- Buffer evaluation works or returns the documented configured fallback.
- Temporary resources are deterministic and removed on success/failure.

### S17 — Bulk cell/range access

Goal: avoid repeated per-cell metadata scans and evaluator setup for row/range iteration.

Current issue:

- `MessySheet.iter_rows()` and range access call `get_cell()` for every cell.
- Merged-range and hidden-state checks are repeated per cell.
- Formula and workbook state are accessed through a scalar-only interface.

Test first:

- Scalar versus bulk parity for values, formulas, merged flags, hidden flags, data types,
  formatting, ranges, and bounds.
- Benchmarks for 1,000-cell and 100,000-cell reads.

Implementation direction:

- Add an internal bulk cell reader using precomputed merged/hidden indexes.
- Keep scalar `get_cell()` as a wrapper over the same logic.
- Batch formula-cache access without changing evaluation semantics.

Acceptance criteria:

- Scalar and bulk results are identical.
- Large range iteration has materially lower time complexity and measured runtime.
- Memory remains bounded through iterator/chunk behavior.

### S18 — Typed table and structure models

Goal: remove internal dict round-trips while preserving serialization compatibility.

Current issue:

- Structure analysis creates `TableInfo`, converts it to dictionaries, and `MessyTable`
  converts dictionaries back to `TableInfo`.
- Dict keys are accessed directly in workbook planning logic.

Test first:

- Protect public `StructureInfo.table_ranges` serialization and existing dict access.
- Protect table counts, range strings, DataFrame extraction, and model equality.

Implementation direction:

- Store typed table models internally.
- Expose a compatibility serialization property or explicit `to_dict()` boundary.
- Add precise types for table ranges and drop untyped internal dictionaries.

Acceptance criteria:

- No internal TableInfo → dict → TableInfo cycle remains.
- Existing serialized output remains available and documented.
- Type checks catch invalid table-range fields.

### S19 — Cache policy, identity, and concurrency hardening

Goal: make cache correctness explicit for all analysis inputs and source types.

Current issue:

- Structure caching is global and path/mtime based.
- Only currently used analysis variation is represented; future options could be omitted.
- Buffers are intentionally uncached, and cache ownership is not configurable per workbook.

Test first:

- Configuration fingerprints, rapid file replacement, same-size/same-mtime changes,
  symlinks, invalidation, LRU eviction, concurrency, and custom cache injection.
- Confirm no cached mutable result leaks caller mutation.

Implementation direction:

- Create an immutable analysis-options fingerprint.
- Consider stable file identity using resolved path plus high-resolution stat attributes.
- Permit explicit cache injection or disabling without changing defaults.
- Keep buffer caching opt-in and content-identity based if added.

Acceptance criteria:

- Every analysis-affecting input participates in cache identity.
- Cache behavior is deterministic under concurrent access.
- Disabling or replacing the cache is supported through a documented seam.

### S20 — Error taxonomy and diagnostics

Goal: make failures actionable without broad catches hiding programmer defects.

Current issue:

- Some boundaries catch broad `Exception` and return fallback values such as `Sheet1`.
- Multi-sheet errors, backend attempts, and optional-dependency failures use different
  shapes and levels of detail.

Test first:

- Permission, missing file, corruption, unsupported format, missing dependency, invalid
  sheet, parser failure, evaluator failure, and partial multi-sheet success.
- Protect error type, message, context, exception chaining, and fatal exception propagation.

Implementation direction:

- Define which adapter exceptions map to each library exception.
- Preserve causes with `raise ... from ...`.
- Replace silent fallback with explicit diagnostics where compatibility permits.
- Use structured attempt records internally and serialize at the public error boundary.

Acceptance criteria:

- Programmer errors are not swallowed by fallback chains.
- Users can identify file, sheet, format, backend, operation, and attempted recovery.
- Partial multi-sheet failures retain structured per-sheet information.

### S21 — Public API typing, deprecations, and lifecycle polish

Goal: make the supported API smaller, clearer, and harder to misuse without breaking it.

Current issue:

- Convenience functions accept narrower path types than `MessyWorkbook`.
- Some internal extensibility is public by accident while useful seams are undocumented.
- Legacy names such as `col_count` and compatibility dict shapes need an explicit policy.

Test first:

- Runtime API parity and static typing examples for paths, buffers, configs, registries,
  return unions, and context managers.
- Warning tests for any proposed deprecation.

Implementation direction:

- Introduce shared public input aliases and overloads where useful.
- Document ownership and context-manager behavior.
- Add compatibility aliases before any deprecation warning.
- Do not remove deprecated behavior in this refactor program.

Acceptance criteria:

- Public signatures reflect supported runtime inputs.
- Documentation examples type-check and execute.
- Any deprecation has a replacement, warning test, changelog entry, and release timeline.

### S22 — CI matrix, performance budgets, and release hardening

Goal: make the simplified architecture safe to maintain after the refactor program.

Test first:

- Validate the package on supported Python versions and dependency extras.
- Test minimum and current compatible dependency sets where practical.
- Smoke-test built wheel and sdist rather than only the source checkout.

Implementation direction:

- CI matrix for Python 3.11–3.14 and core, XLS, formula, and all extras.
- Separate fast required tests from scheduled large/generated/performance suites.
- Add benchmark comparison reporting with agreed regression thresholds.
- Enforce strict docs build and package metadata validation.

Acceptance criteria:

- Wheel and sdist install and pass smoke tests in clean environments.
- Every optional capability has at least one CI job.
- A statistically meaningful hot-path regression over 10% is reported and blocks release
  unless explicitly accepted.
- Release notes summarize behavior guarantees, deprecations, and measured improvements.

## Decision log

Append decisions here; do not rewrite history after a slice is complete.

| Date | Slice | Decision | Reason |
|---|---|---|---|
| 2026-07-21 | S02 | Multi-sheet parsing delegates to `MessyWorkbook` | Prevent public parsing paths from drifting |
| 2026-07-21 | S04 | Handler fallback is limited to compatible handlers | Prevent binary Excel data from reaching unrelated parsers |
| 2026-07-21 | S04 | Formula expressions and cached values use separate workbook views | They are distinct Excel concepts and drive different modes |
| 2026-07-21 | S04 | Structure cache varies by analysis configuration | Cached analysis must represent its inputs |
| 2026-07-21 | S05 | Library-owned resources close deterministically; caller buffers remain caller-owned | Explicit ownership prevents success/failure leaks without changing source semantics |
| 2026-07-21 | S06 | Compile parse decisions into a private frozen plan and keep handler projections fresh | One pure decision boundary prevents configuration drift without changing handler APIs |
| 2026-07-21 | S06 | Bounded ranges ignore automatic sheet-global footer evidence; explicit footer settings still win | Global row evidence cannot be safely applied to local coordinates |
| 2026-07-21 | S06 | Preserve arbitrary drop-condition payload identity inside the frozen plan structure | Recursive freezing or copying can change user-defined comparison behavior |
| 2026-07-21 | S07 | Borrow seekable streams from byte zero and restore each borrow's entry cursor; snapshot non-seekable streams once without taking ownership | One repeatable source policy makes success and failure behavior consistent |
| 2026-07-21 | S07 | Keep paths as paths and create memoized, normalized, or detached byte views only when a backend requires them | Preserve native fast paths and avoid unconditional copies |
| 2026-07-21 | S07 | Give legacy extensions fresh raw borrows unless they explicitly opt into `SourceHandle` | Preserve detector, handler, and registry-subclass compatibility |
| 2026-07-21 | S07 | Use cached values for structure and a separate sequential expression scan for formulas | Formula-only cells must not change structural evidence, and read-only random cell access is prohibitively slow |
| 2026-07-21 | Plan | Resource lifecycle precedes session reuse | Optimization needs explicit ownership to remain safe |
| 2026-07-21 | Plan | Public removals are outside this refactor program | Simplification must preserve current capability |

## Performance log

Add one row for every performance-sensitive slice. Store the exact command and environment
notes in the result column when they materially affect comparison.

| Date | Slice | Scenario | Before | After | Result |
|---|---|---|---:|---:|---|
| 2026-07-21 | Baseline | XLSX 1,000 rows | — | ~25.6 ms | Reference only |
| 2026-07-21 | Baseline | XLSX 1,000 rows, no normalization | — | ~3.9 ms | Reference only |
| 2026-07-21 | Baseline | CSV 1,000 rows | — | ~23.2 ms | Reference only |
| 2026-07-21 | Baseline | Merged XLSX, 100 rows | — | ~17.0 ms | Reference only |
| 2026-07-21 | S07 | Official XLSX, 1,000 rows | 25.6 ms | 25.5 ms | No material regression; system Python 3.14.6, pandas 3.0.0 |
| 2026-07-21 | S07 | Official CSV, 1,000 rows | 23.2 ms | 23.4 ms | No material regression; same full-suite benchmark run |
| 2026-07-21 | S07 | Cold forced structure analysis, path | 304.3 ms | 43.33 ms | Sequential read-only formula scan removed repeated XML reparsing |
| 2026-07-21 | S07 | Cold forced structure analysis, buffer | 299.0 ms | 42.87 ms | 0.989× path time; caller cursor and ownership preserved |
| 2026-07-21 | S07 | CSV buffer, 300,000 rows | 155.9 ms path | 182.1 ms buffer | 17% and 2.79× traced-memory overhead deferred to S13 |

## Risks to watch throughout

- Coordinate drift between zero-based DataFrame rows and one-based Excel rows.
- Fast paths silently losing merged, hidden, range, macro, or formula fidelity.
- Caller buffers being closed, consumed, or left at surprising positions.
- Cached structure being reused across incompatible configurations or file revisions.
- Normalization optimization changing dtypes, null representation, row order, or indexes.
- Optional dependencies changing import-time behavior for core-only installations.
- Convenience and multi-sheet APIs drifting from `MessyWorkbook` again.
- Benchmarks measuring warm cache when the intended target is cold-start performance.
