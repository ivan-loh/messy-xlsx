# Native CSV Tokenizer Design

**Status:** Approved design for the parser-performance v1.0.0 work

**Date:** 2026-07-26

**Parent plan:** `docs/superpowers/plans/2026-07-22-parser-performance-v1.md`

## Summary

Task 14 replaces CSV, TSV, and TXT materialized batch adapters with bounded
streaming. The initial implementation demonstrated that pandas' public chunk
API cannot simultaneously preserve whole-file C-parser behavior, stop at a
public batch boundary, and retain only `batch_size + skip_footer` rows.

In particular, an excess-field row at the start of a pandas chunk is silently
truncated instead of being rejected, and the C reader may tokenize tens of
thousands of short records before returning a one-row chunk. A Python
`csv.reader` validator stays bounded but differs from pandas C behavior for
malformed quotes, NUL bytes, CR-only input, and first-row implicit indexes.

The selected solution is a focused Cython tokenizer that reproduces the
materialized pandas 3.0.5 CSV contract while enforcing literal streaming
bounds. Python continues to own source management, schema compilation, value
normalization, Arrow conversion, and public lifecycle behavior. Official
platform wheels contain the native extension. If it is unavailable before an
operation begins, the operation uses the exact materialized pandas fallback.

## Goals

1. Preserve the existing public API and the materialized CSV values, columns,
   warnings, and error classes.
2. Produce stable Arrow schemas before a public stream is returned.
3. Yield at most `batch_size` accepted rows per public batch.
4. Retain at most `batch_size + skip_footer` completed logical records, plus
   the current record and a fixed byte buffer.
5. Do not tokenize a future logical record after the requested public batch is
   full.
6. Preserve exact caller-stream cursor and ownership behavior.
7. Keep full-pass memory proportional to the public batch, footer, current
   record, and field payloads rather than source size.
8. Accelerate quoted, CRLF, multiline, and malformed CSV workloads that are
   slow under the Python framing implementation.
9. Ship supported native wheels for CPython 3.11 through 3.14 on Linux, macOS,
   and Windows.
10. Preserve installation and exact behavior on unsupported environments
    through a materialized fallback wheel.

## Non-goals

1. The native layer will not infer numeric, boolean, temporal, Arrow, or public
   display types.
2. It will not depend on pandas, NumPy, or PyArrow C APIs.
3. It will not expose a new public tokenizer API.
4. It will not replace spreadsheet parsing, normalization, or source
   management.
5. It will not switch from native to materialized parsing after consuming
   source bytes or emitting output.
6. It will not support free-threaded `cp313t` or `cp314t` builds in v1.0.0.
7. It will not attempt to make arbitrarily large logical records consume
   constant bytes; one current record and retained footer records are inherent
   payload costs.

## Global constraints

- Python versions: CPython 3.11, 3.12, 3.13, and 3.14.
- Pandas runtime range for v1.0.0: `pandas>=3.0.5,<3.1`.
- Primary semantic oracle: pandas 3.0.5 materialized `read_csv` behavior.
- Public release target: messy-xlsx v1.0.0.
- Native ABI floor: CPython 3.11 Stable ABI (`cp311-abi3`).
- Native implementation language: Cython 3.1 or a later verified compatible
  Cython 3.x release.
- Native build backend: `setuptools.build_meta`.
- Native wheel builder: cibuildwheel.
- No pandas, NumPy, or PyArrow headers in the extension.
- No complete source `bytes`, decoded `str`, `StringIO`, or DataFrame copy in
  the native streaming route.
- No unbounded `read()` request against caller-owned streams.
- No native-to-fallback transition after operation selection.
- `CONTINUE.md` remains untracked and outside all commits.
- No `uv.lock` is added.

## Approaches considered

### Cython with the Stable ABI — selected

Cython provides C-speed state-machine loops while keeping ownership and error
paths readable to Python maintainers. The Python 3.11 Limited API supports the
memoryview facilities required by this design. Stable-ABI support must be
verified on every claimed Python version and platform.

### Rust with PyO3

Rust provides strong memory safety and a good state-machine model, but it adds
Cargo, rustfmt, clippy, Rust security auditing, additional caches, and a second
release ecosystem. That cost is not justified for one internal tokenizer.

### Handwritten CPython C

Handwritten C could produce small binaries with no code-generation dependency,
but manual reference counting and buffer ownership make parser, callback, and
exception paths substantially harder to review and more likely to crash the
interpreter.

## Architecture

```text
SourceHandle
  ├─ bounded inspection and native footer-aware evidence
  │    └─ stable physical and normalization schema
  │
  └─ first public batch request
       └─ NativeCSVTokenizer
            ├─ fixed undecoded byte buffer
            ├─ one current logical record
            ├─ raw footer deque ≤ skip_footer
            └─ accepted output ≤ batch_size
                 ↓
          PandasCSVValueAdapter
                 ↓
          existing Arrow physical encoder
                 ↓
          existing normalization pipeline
                 ↓
          public RecordBatch

Native module unavailable before construction
  └─ existing materialized pandas compatibility reader
```

The native extension is an internal component named
`messy_xlsx._csv_tokenizer`. It owns record framing and every decision that
affects which raw fields reach Python. It returns textual field values and
structural missing values; Python maps pandas-compatible NA spellings and
passes values through the existing stable-schema Arrow pipeline.

The materialized fallback is selected once, before a public operation is
returned. Missing native support is not an ordinary parse failure and does not
trigger retry after startup.

## Component boundaries

### Native configuration

Python supplies an immutable configuration equivalent to:

```python
@dataclass(frozen=True, slots=True)
class NativeCSVConfig:
    encoding: str
    encoding_errors: Literal["strict", "ignore"]
    delimiter: str
    quote_char: str
    double_quote: bool
    skip_initial_space: bool
    skip_blank_lines: bool
    expected_fields: int
    data_start_records: int
    skip_footer: int
    max_rows: int | None
    source_description: str
```

`expected_fields` is the named data-column width after Python header
compilation. `data_start_records` is the number of complete raw logical
records protected as metadata and header evidence. The tokenizer does not
regenerate column names or reinterpret multi-row headers.

### Native tokenizer

The internal interface is equivalent to:

```python
class NativeCSVTokenizer:
    def __init__(self, config: NativeCSVConfig) -> None: ...
    def bind(self, source: BinaryIO) -> None: ...
    def read_batch(
        self,
        max_rows: int,
        on_bad_row: Callable[[BadRow], None],
    ) -> NativeCSVRead: ...
    def close(self) -> None: ...
```

Construction and `bind()` are inert and do not read source data. The first
`read_batch()` begins the full pass. `read_batch(N)` returns zero through `N`
accepted rows and stops tokenizing as soon as `N` are ready.

`NativeCSVRead` contains column-major `str | None` values, its row count, EOF
state, and debug counters. `None` at this boundary means a structurally
missing/padded field. An empty field remains `""` until the Python value
adapter applies pandas NA rules.

`on_bad_row` is synchronous. This avoids an unbounded warning queue when many
malformed records precede the next accepted row. A callback exception makes
the tokenizer terminal and releases its native state during close.

### Python native adapter

`csv_native.py` owns:

- native-module selection;
- construction of `NativeCSVConfig`;
- binding the active `SourceHandle` borrow;
- warning translation to `pandas.errors.ParserWarning`;
- translation of ordinary native failures to `FormatError`;
- propagation of process failures;
- close ordering and metrics;
- conversion of `NativeCSVRead` into the existing physical reader contract.

The adapter never closes a caller-owned stream. It closes native allocations
before the `SourceHandle` borrow restores the caller's entry cursor.

### Pandas-compatible value adapter

`csv_value_adapter.py` maps the union of pandas 3.0.5 default NA strings and
the configured `na_values` to `None` using exact string comparisons.

- Quoted and unquoted empty fields follow the pandas oracle.
- Surrounding whitespace prevents a match unless the configured parsing
  behavior removes that whitespace.
- Structurally missing short-row fields are already `None`.
- All other textual lexemes remain unchanged.

Numeric, boolean, date, all-null, sanitization, physical Arrow, and display
name decisions remain in the existing normalization pipeline.

## Tokenization contract

The native layer reproduces pandas C behavior rather than an RFC-only or
Python `_csv.reader` interpretation.

It owns:

- incremental decoding with strict and ignore modes;
- UTF BOM handling and chunk-split code units;
- arbitrary supported one-character delimiters;
- quote, doubled-quote, embedded newline, CRLF, CR-only, and EOF behavior;
- pandas-compatible NUL and post-quote-junk behavior;
- blank logical-record classification;
- protected metadata/header records;
- first nonblank data-row implicit leading-index inference;
- right-padding short rows;
- rejecting later rows wider than the established physical width;
- raw logical footer retention;
- accepted-row `max_rows`;
- synchronous bad-row line information;
- discarding inferred index fields while returning named data fields.

Critical oracle cases include:

- The first data row may establish one or more implicit leading index fields.
- A later excess-field row is rejected regardless of public batch position.
- `b'a,b\r x ,q"z,q"z\r'` matches whole-file pandas C output exactly.
- `1,x\x00junk` is accepted as `1,x`, while `2,y\x00,z` is classified using
  the pandas C NUL and delimiter rules.
- A trailing blank or malformed raw record consumes a footer slot without
  emitting a warning.
- Footer retention happens before blank and malformed filtering.

## Streaming and memory invariants

At every native callback and public return:

- accepted output records retained are at most `batch_size`;
- completed footer records retained are at most `skip_footer`;
- completed records retained are at most `batch_size + skip_footer`;
- exactly one additional current logical record may be under construction;
- the undecoded read buffer has a fixed configured capacity;
- `tokenized_ahead_records == 0` when a batch is returned;
- no record after the requested accepted batch has been decoded or classified.

Byte memory is:

```text
O(
  fixed read buffer
  + current logical record payload
  + retained footer payloads
  + accepted output field payloads
)
```

It is independent of total source size. Records larger than the schema sample
budget are valid in the full pass. A future plan may spool individual
oversized records, but v1.0.0 retains the current record because the output
must eventually represent its payload.

## Sampling and stable schema

Public streams still compile configuration and stable schemas before return.
Native bounded evidence uses the same framing, footer, blank, malformed, and
implicit-index rules as the full pass.

The evidence pass may examine:

```text
data_start_records + sample_rows + skip_footer
```

logical records, subject to the existing schema byte and cell budgets. Footer
lookahead bytes count toward the byte budget. It does not create a complete
source copy.

If exact footer-aware evidence cannot complete within the configured schema
budgets, the operation selects the materialized fallback before returning a
public stream. It does not infer from records that may all belong to the
footer. If EOF proves that every data record is removed, the empty result
schema matches the materialized pandas-to-Arrow oracle.

The full row pass starts from a fresh borrow on first iteration. Evidence
never shares partially consumed tokenizer state with output.

## Row limits and footer order

Processing order is:

```text
raw logical records
  → protect metadata/header records
  → retain final skip_footer records
  → classify blanks and malformed rows
  → establish/discard implicit leading index fields
  → count accepted rows against max_rows
  → return named data fields
```

`max_rows` counts accepted rows. The established low-level
`max_rows + skip_footer` pandas error is rejected before native tokenizer
startup, preserving the materialized error contract.

## Encoding behavior

Path inputs use strict decoding during bounded evidence. If a decode error
occurs within evidence, Python performs the established one-time Latin-1
fallback and recomputes metadata/header evidence. Caller-owned streams retain
their existing ignore/error policy and do not retry.

After the public stream is returned, no encoding fallback is allowed. A late
decode error is lazy, contextual, and terminal. Bytes fetched into the fixed
buffer but not yet decoded do not raise until demand reaches them.

The tokenizer may use Python incremental codecs internally for uncommon
encodings. Common UTF-8, Latin-1, and UTF-16 paths may receive specialized
Cython loops, provided their oracle behavior is identical.

## Errors and lifecycle

- Ordinary parser, decoder, callback, and I/O failures become contextual
  `FormatError` at the Python adapter boundary.
- `MemoryError`, `KeyboardInterrupt`, `SystemExit`, and any exception chain
  containing an established process failure propagate unchanged.
- Process-failure classification happens before decode-fallback
  classification.
- Cleanup failures never replace an active process failure.
- Any runtime failure makes native tokenizer state terminal.
- There is no retry or implementation switch after source consumption.
- `close()` is idempotent and releases native allocations on success,
  ordinary failure, process failure, downstream normalization failure, early
  close, and exhaustion.
- The native tokenizer closes before the source borrow.
- The existing Task 13 weak ownership, finalizer, return-gap cleanup, and
  one-active-operation behavior remains authoritative.
- EOF may accompany a final nonempty batch. The borrow may close after that
  batch is fully constructed.

## Native selection and fallback

Native capability is resolved once before public operation construction.

Official supported wheels must contain and load the extension. A missing
extension fails that wheel's smoke test. Runtime fallback catches only native
import/load unavailability. It never catches arbitrary execution exceptions.

Fallback behavior:

- use the existing materialized pandas compatibility reader;
- preserve values, dtypes, warnings, errors, and public API;
- do not claim native streaming memory guarantees;
- expose the selected backend through existing internal metrics/debug state;
- allow deterministic selection with
  `MESSY_XLSX_DISABLE_NATIVE=1`.

The project publishes a universal `py3-none-any` fallback wheel for unsupported
architectures and interpreters. Compatible CPython installations prefer the
platform `cp311-abi3` wheel.

## Build and release design

The project switches from Hatchling to `setuptools.build_meta` while retaining
PEP 621 metadata and `src` package discovery. The v1.0.0 pandas dependency is
`pandas>=3.0.5,<3.1`, keeping native and fallback semantics inside the reviewed
pandas 3.0 compatibility line.

Build requirements include a narrowly verified Cython 3.x range and
setuptools. The extension defines:

```text
Py_LIMITED_API = 0x030B0000
py_limited_api = true
wheel tag = cp311-abi3
```

The extension does not include pandas, NumPy, or PyArrow headers.

cibuildwheel produces:

- Linux x86-64 and aarch64 `cp311-abi3` wheels;
- macOS x86-64 and arm64 `cp311-abi3` wheels;
- Windows x86-64 `cp311-abi3` wheels.

Each wheel is installed and smoke-tested on CPython 3.11, 3.12, 3.13, and
3.14 where the runner supports that interpreter. Smoke tests assert that the
native module, not the fallback, is active.

The source distribution is built once and contains the `.pyx` source. Its
isolated build installs Cython and requires a platform compiler. The universal
fallback wheel provides compiler-free installation on unsupported systems.

Release jobs merge native wheels, the universal fallback wheel, and the
source distribution; run `twine check`; smoke-test native and fallback
selection; then publish the single verified artifact set.

## Verification strategy

### Differential oracle tests

Every semantic fixture compares the materialized `CSVHandler` and native
tokenizer paths using pandas 3.0.5 across public batch sizes 1, 2, 3, and 127.
Assertions cover values, columns, warnings, error class, and error context.

The matrix includes:

- LF, CRLF, CR-only, and missing final terminators;
- quoted delimiters, escaped quotes, multiline quotes, and unterminated EOF;
- UTF-8, UTF-16, BOM splits, code-unit splits, Latin-1 fallback, and late
  errors;
- exact, short, empty, implicit-index, later-wide, consecutive-bad, and
  all-bad rows;
- quote junk, quotes inside unquoted fields, and NUL positions;
- zero, small, all-row, blank, malformed, multiline, and large footers;
- pandas default and configured NA values, whitespace, missing trailing
  fields, and all-null columns;
- path, seekable, nonseekable, and one-byte-read sources;
- records larger than 8 MiB.

### Differential fuzzing

Hypothesis generates valid and malformed byte streams and randomized chunk
splits. It compares accepted values, columns, warnings, physical line
information, and error classes against pandas 3.0.5. The native and materialized
routes must agree before production routing is enabled.

### Lifecycle and failure tests

Tests cover inert construction, first-read startup, early close, exhaustion,
warning-callback failure, decoder/parser/I/O/process failures, cleanup
failures, exact cursor restoration, no caller close, finalizer cleanup, and
active-operation release.

Chained process-failure tests prove that decode or cleanup context never masks
`MemoryError`, `KeyboardInterrupt`, or `SystemExit`.

### Deterministic bound tests

Debug state exposes:

- `output_rows_retained`;
- `footer_records_retained`;
- `current_record_active`;
- `undecoded_buffer_bytes`;
- `tokenized_ahead_records`.

Tests assert the invariants after every callback and batch return. These are
architecture tests rather than timing-based proxies.

### Performance tests

Benchmarks cover:

- clean unquoted LF;
- quoted LF;
- CRLF;
- multiline quoted fields;
- sparse malformed rows;
- `batch_size=1`;
- large logical records;
- nonzero footer retention.

Tokenizer-only throughput targets:

- no benchmark slower than 3.0 times direct pandas C;
- geometric-mean slowdown at most 2.0 times direct pandas C.

End-to-end reports include throughput, peak retained rows, retained payload
bytes, and source position at each public batch. Relative timings are recorded
in the performance report; deterministic CI gates enforce route selection and
memory invariants instead of fragile wall-clock thresholds.

## Delivery stages

1. Freeze the pandas 3.0.5 oracle with differential fixtures and fuzzing.
2. Add the internal ABI, deterministic test double, lifecycle adapter, and
   materialized fallback selection.
3. Implement valid-record framing, decoding, quoting, delimiters, line endings,
   BOMs, blanks, and large records.
4. Implement pandas C compatibility for implicit indexes, short rows, NULs,
   quote junk, excess fields, warnings, and malformed fuzz cases.
5. Implement raw footer retention, accepted-row limits, all-footer schemas,
   and literal memory counters.
6. Add the Python NA adapter and integrate native textual batches with the
   existing Arrow/normalization pipeline.
7. Enable production routing only after all semantic, lifecycle, and bound
   gates pass.
8. Add setuptools, Stable-ABI builds, cibuildwheel, native wheel smoke tests,
   the universal fallback wheel, and merged release artifacts.
9. Run performance gates, full repository verification, independent native
   safety review, and final compatibility review.

## Acceptance criteria

The design is complete when:

1. Native and materialized outputs agree across the fixed oracle matrix and
   differential fuzzing.
2. Every public native batch contains at most `batch_size` accepted rows.
3. Deterministic counters prove the completed-record and tokenized-ahead
   bounds.
4. Late errors remain demand-driven and process failures remain unmasked.
5. Caller streams are never closed and their exact entry cursors are restored.
6. Custom registry handlers remain authoritative and materialized.
7. Official native wheels load on supported Python/platform combinations.
8. The universal fallback wheel selects materialized pandas behavior.
9. Full tests, Ruff, formatting, mypy, security checks, docs, wheel smoke
   tests, and source-distribution smoke tests pass.
10. Performance targets are met or any miss is explicitly reviewed before
    production routing.
