# Parser Performance and Bounded-Memory Architecture Design

**Status:** Approved design for implementation planning

**Target release:** v1.0.0

**Date:** 2026-07-22

**Selected approach:** Hybrid OOXML manifest and Arrow batch pipeline

## Summary

messy-xlsx v1.0.0 will replace repeated whole-workbook analysis with a
workbook-level OOXML manifest, bounded structure sampling, explicit backend
routing, and a canonical Arrow `RecordBatch` pipeline. Ordinary XLSX and XLSM
files will use fastexcel without first loading the complete workbook through
openpyxl. Openpyxl remains an explicit compatibility fallback for behavior that
cannot be preserved by fastexcel and coordinate-aware Arrow transforms.

All existing public APIs remain functional throughout v1.x. Materialized
DataFrame entry points become legacy compatibility APIs and emit a
`LegacyAPIWarning`, derived from `DeprecationWarning`. Their signatures, return
types, values, columns, dtypes, indexes, exceptions, and parsing defaults remain
the compatibility authority.

New additive APIs expose Arrow tables, bounded Arrow batches, pandas chunks, and
sequential multi-sheet results. Arrow `RecordBatch` is the canonical internal
and streaming representation.

All improvements ship in one coordinated v1.0.0 release. Development remains
divided into independently verified milestones, but no partial public release is
planned.

## Motivation and measured baseline

Profiling shows that structure analysis, rather than data parsing or
normalization, dominates runtime and memory.

On the existing 100,000-row `sales_transactions.xlsx` sample:

| Configuration | Elapsed time | Peak process RSS |
|---|---:|---:|
| Current default | 9.99 s | 627 MB |
| Current default without normalization | 9.04 s | 577 MB |
| Auto-detection disabled, openpyxl path | 8.02 s | 563 MB |
| Structure analysis followed by fastexcel | 9.73 s | 628 MB |
| Fastexcel-compatible normalized parse | 1.56 s | 227 MB |
| Fastexcel-compatible raw parse | 0.85 s | 200 MB |

For the workbook without merged or hidden cells, the normalized
fastexcel-compatible result had identical values, columns, and dtypes to the
current default result.

A profile across three maintained large samples attributed:

- 27.9 of 30.2 parsing seconds to structure analysis.
- Six complete `openpyxl.load_workbook()` calls to three sheets.
- 8.6 seconds to a second workbook view used only for formula detection.
- 0.70 seconds to fastexcel parsing.
- 1.64 seconds to normalization.

Multi-sheet instrumentation on a three-sheet workbook observed:

| API | openpyxl loads | sheet parses | structure analyses |
|---|---:|---:|---:|
| `MessyWorkbook.to_dataframes()` | 6 | 3 | 3 |
| `read_all_sheets()` | 9 | 6 | 3 |

For a generated 300,000-row, 20 MB CSV:

| Input | Elapsed time | Peak process RSS |
|---|---:|---:|
| Path, normalized | 1.58 s | 267 MB |
| Seekable buffer, normalized | 1.68 s | 352 MB |
| Path, raw | 0.64 s | 231 MB |
| Seekable buffer, raw | 0.72 s | 349 MB |

The buffer path currently retains complete bytes, a decoded string, a
`StringIO` view, and the output DataFrame simultaneously.

## Goals

1. Preserve all existing public behavior while materially improving ordinary
   XLSX/XLSM parsing speed and peak memory.
2. Build workbook metadata once and reuse it across every sheet and parse.
3. Parse each selected sheet once.
4. Make Arrow `RecordBatch` the canonical backend and streaming boundary.
5. Add bounded-memory row and multi-sheet APIs.
6. Retain openpyxl as an explicit, resource-safe compatibility fallback.
7. Remove avoidable whole-file copies for seekable CSV and binary streams.
8. Reduce normalization copies and repeated per-column scans.
9. Make backend decisions, fallback reasons, and performance contracts
   observable and testable.
10. Ship all approved work as v1.0.0 after complete compatibility and package
    verification.

## Non-goals

- Removing any existing public API during v1.x.
- Establishing an automatic removal date for legacy APIs.
- Supporting XLSB.
- Implementing a complete Excel calculation engine.
- Replacing fastexcel with a custom cell-value OOXML parser.
- Recalculating formulas during DataFrame or Arrow parsing.
- Guaranteeing identical performance numbers on every machine.
- Making parallel sheet parsing the default. Sequential bounded-memory behavior
  takes priority over throughput that multiplies memory.
- Adding a persistent disk cache for parsed worksheet data.

## Considered approaches

### Selected: hybrid manifest and Arrow pipeline

Read OOXML metadata once, use fastexcel for ordinary data, apply coordinate-aware
features in Arrow, and retain a streaming openpyxl fallback. This provides most
of the demonstrated performance benefit without reimplementing Excel cell
decoding.

### Rejected: fully Arrow-native OOXML engine

Using fastexcel plus custom OOXML logic for all formulas, values, styles, and
edge cases offers the highest theoretical performance. It also creates a
partial Excel engine and carries unacceptable compatibility risk for v1.0.0.

### Rejected: conservative openpyxl retrofit

Reusing openpyxl views and reducing copies is lower risk, but it cannot remove
the dominant full-workbook object graph or approach the measured fastexcel
speed and memory profile.

## Architectural overview

```text
SourceHandle
    |
    v
Format detection
    |
    +--> OOXML WorkbookManifest
    |       - sheets and dimensions
    |       - merged ranges
    |       - hidden rows/columns
    |       - formula presence
    |       - styles and locale evidence
    |
    v
Bounded structure sampling
    |
    v
Immutable ParsePlan
    |
    +--> FastexcelBatchReader ------+
    |                               |
    +--> OpenpyxlFallbackReader ----+--> Arrow RecordBatch stream
    |                               |
    +--> CSVBatchReader ------------+
    |                               |
    +--> LegacyXlsAdapter ----------+
                                    |
                    +---------------+----------------+
                    |                                |
          LegacyDataFrameAdapter              New streaming APIs
          Arrow table -> pandas               RecordBatch iterator
          compatibility normalization         pandas chunk adapter
```

## Component boundaries

### `SourceHandle`

`SourceHandle` remains the source-ownership authority.

- Paths remain paths for backends that open them directly.
- Caller-owned streams are never closed.
- Seekable caller streams are borrowed at byte zero and restored after use.
- Non-seekable streams retain one unavoidable immutable snapshot.
- Seekable streams are not copied into a permanent byte cache unless a backend
  demonstrably requires immutable bytes.
- Prefix inspection must not cause complete-source caching.
- Nested active borrows of one caller-owned seekable stream are rejected with a
  clear `RuntimeError`.

### `WorkbookManifest`

`WorkbookManifest` is an immutable workbook-level representation built once per
OOXML source. It stores metadata, never complete worksheet values.

Workbook metadata includes:

- Workbook type and relationships.
- Ordered sheet names, relationship targets, visibility, and dimensions.
- Shared-string table presence and counts, but never the shared-string values;
  cell text remains the data reader's responsibility.
- Style metadata required for structural decisions.
- Date system and relevant number-format evidence.

Per-sheet metadata includes:

- Declared and observed dimensions.
- Merged ranges.
- Hidden rows and columns.
- Formula presence and sampled formula coordinates.
- Candidate data range.
- Worksheet XML target.

The manifest parses ZIP/XML entries incrementally. Formula detection scans
worksheet XML for formula tags instead of opening a second openpyxl workbook.
Manifest cache keys for paths include the resolved path, file size, and
nanosecond modification time. Streams use workbook-local source identity and
are not cached globally.

### `StructureSampler`

`StructureSampler` combines the manifest with bounded fastexcel samples. It
detects:

- Header position and confidence.
- Metadata rows.
- Footer boundaries.
- Multiple-table separators.
- Locale evidence.
- Sparse columns.
- Pivot-like sheet indicators.

Sampling is bounded by explicit row and column limits. Exact row and column
counts come from fastexcel `total_height` and `width` or manifest dimensions,
not a discarded complete DataFrame.

One sampler instance analyzes every requested sheet through one workbook
reader. Results are cached by sheet and header-pattern variant inside
`MessyWorkbook`, including for buffer-backed workbooks.

### `ParsePlan`

`ParsePlan` remains immutable and gains explicit fields for:

- Selected backend and decision reason.
- Original worksheet coordinate projection.
- Batch size.
- Output representation.
- Manifest merge and hidden-coordinate transforms.
- Formula representation.
- Normalization schema and coercion policy.
- Required global transformations.
- Legacy compatibility mode.

Configuration is compiled once per sheet without mutating caller-owned objects.

### `BackendRouter`

`BackendRouter` selects a backend before parsing data. Routing is capability
based rather than exception driven.

Fastexcel is preferred for:

- Cached formula values.
- Ordinary cell values.
- Hidden-row and hidden-column filtering using manifest coordinates.
- Merged-cell `fill`, `skip`, and `first_only` transformations.
- Rectangular cell ranges that can be pushed down or projected.
- Row and column limits.

Openpyxl is selected for:

- Formula-expression mode when fastexcel cannot preserve expressions.
- Unsupported cell representations.
- Workbooks classified as incompatible by a concrete capability check.

Tests and internal diagnostics may override routing explicitly, but v1.0.0 does
not add a public backend-selection setting. Public callers receive capability-
based routing so implementation engines remain replaceable without an API
compatibility commitment.

Legacy XLS remains optional and uses `xlrd` row windows to build bounded Arrow
batches without an intermediate complete DataFrame. The `xlrd` workbook model
and source overhead are format-level constraints and are documented separately
from batch memory. Custom `HandlerRegistry` implementations remain supported
through a legacy DataFrame-to-Arrow adapter and therefore do not receive the
built-in backends' bounded-memory guarantee.

### `BatchReader`

All built-in backends implement one internal protocol that yields
`pyarrow.RecordBatch` objects. Batches carry private original-row coordinate
metadata until coordinate-sensitive transformations finish.

The reader contract requires:

- Stable schema for one stream.
- Configurable bounded batch size.
- Deterministic cleanup on exhaustion, error, or explicit close.
- No complete Python list-of-lists between the backend and Arrow.
- Preservation of row order.
- Structured backend errors.

### Coordinate transforms

Coordinate transforms execute before header removal and row filters so public
row semantics remain unchanged.

- Hidden rows and columns are removed using manifest indexes.
- Merge strategies apply from manifest ranges.
- Merge anchors carry across batch boundaries.
- Cell ranges are projected from original worksheet coordinates.
- Header, footer, and maximum-row rules operate on the transformed coordinate
  stream.

The openpyxl fallback uses `read_only=True` wherever its required behavior
allows. It does not unmerge or mutate a workbook. Merge behavior is applied by
the same coordinate-transform layer used by fastexcel.

### `LegacyDataFrameAdapter`

The legacy adapter is the compatibility authority for existing DataFrame APIs.
It materializes Arrow batches, converts once to pandas, and applies
behavior-equivalent whole-sheet normalization.

The adapter is explicitly named and documented as legacy so new implementation
work does not extend the materialized path accidentally.

### Streaming normalization

An immutable `NormalizationPlan` is compiled from caller type hints, structural
evidence, and the same bounded schema sample used by existing type inference.

Arrow compute functions handle transformations with equivalent semantics.
Operations without an exact Arrow equivalent use one-column pandas adapters,
not a complete batch copy.

The plan provides:

- Stable numeric, temporal, identifier, text, and mixed-column types.
- Locale-aware number rules.
- Date parsing rules.
- Missing-value rules.
- Type-consistency rules.
- Column names and renames.
- Row-filter predicates.

Late incompatible values follow the compiled coercion policy instead of
changing the schema mid-stream.

### `BatchStream`

`BatchStream` is a public closable iterator and context manager returned by
batch APIs.

```python
with workbook.iter_batches("Sales") as batches:
    for batch in batches:
        process(batch)
```

- Exhaustion closes it automatically.
- `close()` and context-manager exit provide deterministic early cleanup.
- `MessyWorkbook.close()` closes remaining child streams.
- A stream cannot outlive its workbook.
- Caller-stream cursors are restored when it closes.
- A top-level convenience function owns its internal workbook until the
  returned stream closes.

### Cell access

Random `get_cell()` retains its lazy compatibility workbook because it exposes
formatting and formula details not present in ordinary Arrow batches.

It gains:

- Indexed merged ranges rather than scanning every range for every cell.
- Cached hidden-row and hidden-column sets.
- Shared source bytes for expression and cached-value views when a detached
  stream is unavoidable.
- Separation from DataFrame parsing so cell APIs never force openpyxl into the
  ordinary data path.

`MessySheet.iter_rows()` uses the batch reader when formula and formatting modes
permit it. Formula-evaluating or formatting-sensitive iteration uses the
compatibility accessor automatically.

## Public API design

### Legacy materialized APIs

The following remain available with their current signatures and behavior:

```python
read_excel(...)
read_excel_tables(...)
read_all_sheets(...)

MessyWorkbook.to_dataframe(...)
MessyWorkbook.to_dataframes(...)
MessySheet.to_dataframe(...)
MessyTable.to_dataframe(...)
```

Each call emits `LegacyAPIWarning`, a subclass of `DeprecationWarning`, with a
correct caller-facing `stacklevel`. Standard Python warning filters hide it by
default, while tests, IDEs, and users who enable deprecation warnings can detect
it.

A direct public legacy invocation emits exactly one warning. Legacy adapters
suppress duplicate warnings from any lower-level legacy entry point they call.

Commitments:

- No legacy API is removed or has its signature changed during v1.x.
- Deprecation establishes no automatic removal date.
- Removal requires a separate v2.0 design decision.
- Existing outputs and exceptions remain the compatibility authority.
- Analysis, configuration, structure, cell access, model, and enum APIs are not
  marked legacy.

### New Arrow and bounded-memory APIs

```python
table = workbook.to_arrow(sheet="Sales", config=config)

with workbook.iter_batches(
    sheet="Sales",
    batch_size=65_536,
    config=config,
) as batches:
    for batch in batches:
        process(batch)

with workbook.iter_dataframe_chunks(
    sheet="Sales",
    batch_size=65_536,
    config=config,
) as chunks:
    for chunk in chunks:
        process(chunk)

for result in workbook.iter_sheets(config=config):
    if result.error is not None:
        handle_error(result.error)
    else:
        process(result.name, result.dataframe)
```

Top-level conveniences mirror the workbook methods:

```python
table = read_excel_arrow("data.xlsx", sheet="Sales")

with read_excel_batches("data.xlsx", sheet="Sales") as batches:
    for batch in batches:
        process(batch)
```

`SheetResult` is a frozen public model with:

- `name: str`
- `dataframe: pandas.DataFrame | None`
- `error: SheetError | None`

`iter_sheets()` processes one sheet at a time and never retains previous
DataFrames internally. It is a new API and does not emit a legacy warning even
though each yielded result may contain one materialized DataFrame.

### Streaming semantic boundary

Existing materialized APIs preserve all whole-sheet transformations exactly.

Streaming APIs apply stream-safe transformations and make these semantics
explicit:

- Declared all-null columns are retained because they cannot be removed before
  a true stream ends.
- Batch schemas remain stable for the iterator lifetime.
- Pandas chunks receive monotonically increasing `RangeIndex` values.
- Row order and retained-column values match materialized parsing.
- Concatenating chunks is deterministic, but all-null-column retention may
  differ from legacy materialized output.

`to_arrow()` is materialized and applies global transformations, including
all-null-column removal.

## Multi-sheet behavior

`MessyWorkbook.to_dataframes()` retains its existing return shape and error
behavior. Because it returns a dictionary, it necessarily retains all returned
DataFrames.

The implementation will:

1. Build one manifest.
2. Open one fastexcel workbook reader where supported.
3. Sample each candidate sheet without full parsing.
4. Apply sheet filters before data parsing.
5. Parse each selected sheet exactly once.

`iter_sheets()` provides bounded multi-sheet memory by releasing each
sheet-local parse before advancing after the caller releases the yielded result.

`read_all_sheets()` retains its selection behavior and becomes a legacy adapter
over the shared multi-sheet planning path. It no longer performs a full raw
parse followed by a final parse.

## CSV and text parsing

Path-backed CSV parsing continues to inspect bounded prefixes and pass paths to
pandas.

For seekable streams:

1. Borrow the stream at byte zero.
2. Read bounded prefixes for encoding and delimiter detection.
3. Restore or rewind within the active borrow.
4. Pass the binary stream directly to pandas with an encoding.
5. Avoid complete `bytes`, decoded `str`, and `StringIO` copies.

For non-seekable streams, `SourceHandle` creates one snapshot. Parsers use a
seekable view over that snapshot without decoding the complete file eagerly.

CSV batch APIs use pandas chunking and convert each chunk to Arrow with a stable
schema. `skip_footer`, which requires the Python engine today, is implemented
with a bounded trailing-row buffer so it does not force complete
materialization.

## Normalization and memory model

### Legacy normalization

Legacy materialized APIs preserve:

- Whitespace normalization.
- Locale-aware number conversion.
- Date detection and conversion.
- Missing-value handling.
- Mixed-type coercion.
- Column sanitization and renaming.
- Regex and condition row removal.
- Empty-row and all-null-column removal.
- Final index behavior.

The materialized compatibility pipeline owns its DataFrame and mutates it
internally. It makes at most one intentional full-frame copy rather than one
copy per normalization stage.

Column transforms are planned once to avoid repeatedly scanning all object and
string columns. Regex row filtering builds one bounded column-at-a-time boolean
mask. Multiple drop conditions are combined and applied with one filter and one
index reset.

### Streaming memory invariants

- No Python list-of-lists exists between a built-in reader and Arrow.
- At most one input batch, one transformed batch, and bounded scratch state are
  live per iterator.
- Metadata caches contain no complete worksheet values.
- Seekable streams are not permanently snapshotted.
- Non-seekable streams retain one source-sized snapshot only.
- Multi-sheet iterators retain one yielded sheet result at a time internally.
- Materialized APIs remain bounded below by the size of their required output.

### Optional Arrow-backed pandas dtypes

The new chunk APIs use Arrow-backed pandas dtypes where doing so preserves the
declared schema. Legacy APIs retain current pandas dtypes. A future default-dtype
change is outside this design and would require a separate compatibility
decision.

## Backend fallback and errors

Fallback is explicit and classified.

- Backend capability decisions occur before data parsing.
- Fastexcel retries through openpyxl only for classified compatibility errors.
- Permission errors, missing files, invalid configuration, `MemoryError`, and
  caller-stream ownership failures are never silently retried.
- If a fallback fails, structured error context retains both the selected
  backend decision and the original failure.
- Existing public exception types remain unchanged for legacy APIs.
- New APIs use the same exception hierarchy.
- Batch iteration errors close resources before propagating.
- Per-sheet iteration converts parse errors into `SheetResult.error` without
  swallowing process-level failures such as `MemoryError`.

Debug logging records the manifest cache result, backend selection, fallback
reason, row batches, and cleanup path without logging cell values or sensitive
source contents.

## Initialization and cache improvements

Workbook initialization currently performs separate sheet-name and validation
backend opens. The new inspection path combines format validation, sheet
discovery, and manifest construction where possible.

Workbook-local caches include:

- One manifest.
- One structure result per sheet and header-pattern variant.
- One normalization plan per compatible configuration.
- Hidden-coordinate sets and merged-range indexes.
- Lazily created cell-access workbook views.

The global path structure cache remains bounded and uses robust identity fields.
No complete DataFrame, Arrow table, or batch is stored in a global cache.

## Compatibility requirements

Before changing a parsing path, characterization tests must capture:

- Values, columns, dtypes, indexes, and null representations.
- Exception types, messages where public, and structured context.
- Warning behavior.
- Header and footer decisions.
- Merge strategies.
- Hidden row and column handling.
- Cached formulas and expression preservation.
- Empty cached formula values.
- Single and multi-row headers.
- Metadata and multiple-table behavior.
- Cell ranges.
- XLSX, XLSM, XLS, CSV, and TSV behavior.
- Path, seekable stream, bytearray/memoryview stream, and non-seekable stream
  behavior.
- Caller cursor restoration and ownership.
- Malformed and partially corrupted files.
- Custom handler behavior.
- Single-sheet and multi-sheet output equivalence.

The maintained sample corpus and generated messy-workbook corpus serve as the
primary compatibility suite. Development-only differential tests may retain a
reference copy of the v0.10.0 engine under tests; it is not shipped in the
runtime package.

## Test strategy

### Characterization and golden tests

- Add stable hashes and schema records for maintained samples.
- Use direct DataFrame equality for focused fixtures.
- Record exception and warning contracts.
- Verify legacy APIs against the v0.10.0 reference behavior.

### Property-based tests

Generate combinations of:

- Batch boundaries.
- Merge ranges crossing batches.
- Hidden coordinates around headers and footers.
- Header offsets.
- Missing and mixed values.
- Cursor positions and injected stream failures.
- Malformed OOXML metadata.

### Architectural tests

Regular CI asserts deterministic implementation properties:

- Ordinary XLSX parsing performs zero complete openpyxl workbook loads.
- One manifest is built per workbook.
- Each selected sheet is parsed once.
- Multi-sheet analysis creates no discarded complete DataFrames.
- Seekable CSV streams create no decoded whole-file copy.
- Batch schemas remain stable.
- Early stream closure restores caller cursors.
- Legacy APIs expose no Arrow-specific implementation details.
- Legacy calls emit `LegacyAPIWarning` with a caller-facing stack location.

### Resource-failure tests

Inject failures during:

- Manifest creation.
- Backend initialization.
- First, middle, and final batch production.
- Normalization.
- Fallback.
- Stream restoration.
- Workbook and batch-stream cleanup.

Tests verify deterministic cleanup and primary-exception preservation.

## Performance test strategy

Performance workloads run in isolated subprocesses and record elapsed time,
peak RSS, result shape, schema, and stable output hash.

Stable architectural properties run in the normal CI matrix. Wall-clock and RSS
thresholds run in a dedicated Linux performance job because hosted runner and
operating-system variability makes them unsuitable as universal matrix gates.

### XLSX acceptance target

For the current 100,000-row sample on the current development machine:

- Default parsing completes in under 2 seconds.
- Peak RSS remains under 300 MB.
- Runtime is at least four times faster than the recorded current default.
- Peak RSS is at least 50% lower than the recorded current default.
- Normalized legacy output is exactly equal to the compatibility baseline.

### CSV acceptance target

For the 300,000-row workload:

- Seekable-buffer peak RSS remains within 20% of path-backed parsing.
- Raw and normalized paths have separate baselines.
- Batch parsing remains within a bounded batch-size-dependent memory envelope.

### Multi-sheet acceptance target

- One manifest per workbook.
- Zero discarded complete analysis DataFrames.
- One data parse per selected sheet.
- Ordinary OOXML sheets require zero complete openpyxl loads.
- `iter_sheets()` peak retained output is one sheet.

## Documentation and migration

v1.0.0 documentation will include:

- A legacy materialized API section.
- The meaning and filtering of `LegacyAPIWarning`.
- Arrow and batch API examples.
- Deterministic cleanup examples using context managers.
- Multi-sheet bounded-memory examples.
- Stream ownership and cursor guarantees.
- Backend-routing and compatibility behavior.
- Formula cached-value versus expression behavior.
- Performance-oriented configuration guidance.
- A migration table mapping each legacy API to additive alternatives.

The internal design document is not added to public MkDocs navigation.

## Delivery and release gates

All milestones land in one coordinated v1.0.0 development line:

1. Performance and compatibility contract.
2. OOXML workbook manifest.
3. Bounded structure sampler.
4. Backend router and default fastexcel path.
5. Coordinate-aware Arrow transforms.
6. Streaming openpyxl fallback.
7. Workbook and multi-sheet reuse.
8. CSV stream optimization.
9. Arrow, batch, pandas-chunk, and sheet-result APIs.
10. Normalization consolidation.
11. Cell-access and row-iteration indexes.
12. Integration, documentation, and packaging.

The v1.0.0 tag requires:

- All existing and new tests passing.
- Python 3.11 through 3.14 on Linux, macOS, and Windows.
- Documentation, lint, formatting, typing, and security checks.
- Wheel and source-distribution builds.
- Twine validation.
- Clean installation and smoke tests for wheel and source distribution.
- Optional formula and XLS dependencies tested.
- Benchmark report against v0.10.0.
- Updated changelog, API reference, migration guide, and version metadata.
- No unresolved compatibility failure.

Performance improvements never override correctness or resource-ownership
failures.

## Implementation-planning constraints

The implementation plan following this design must:

- Use test-driven development for every behavior-changing slice.
- Name exact files, tests, commands, and expected failures.
- Keep individual tasks small enough for focused review and commits.
- Establish compatibility tests before replacing each old path.
- Benchmark after each performance-sensitive milestone.
- Preserve safe rollback points until the replacement passes the full corpus.
- Avoid unrelated refactoring.
- Treat the approved architecture and v1.0.0 compatibility commitments as hard
  constraints.

## Final acceptance criteria

The design is complete when v1.0.0 provides all of the following:

1. Existing public APIs remain available and behavior-compatible throughout
   v1.x.
2. Legacy materialized APIs emit `LegacyAPIWarning` and are clearly documented.
3. New Arrow, batch, pandas-chunk, and sequential sheet APIs are public and
   typed.
4. Ordinary XLSX/XLSM parsing avoids complete openpyxl workbook loads.
5. Multi-sheet parsing performs one analysis and one parse per selected sheet.
6. Seekable CSV streams avoid complete bytes-to-text duplication.
7. Streaming memory is bounded by batch size plus documented source snapshots.
8. Caller-owned streams are never closed and their cursors are restored.
9. The maintained compatibility corpus remains green.
10. The measured XLSX, CSV, and multi-sheet targets are met.
11. The package passes the complete v1.0.0 CI, build, and smoke-test gates.
