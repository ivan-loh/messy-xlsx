# Native CSV Tokenizer Design

**Status:** Independently reviewed; pending user approval

**Date:** 2026-07-26

**Revised after independent review:** 2026-07-27

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
`CSVHandler` structural parsing contract under pandas 3.0.5 while enforcing
literal streaming bounds. That contract is engine-dependent: it uses pandas C
semantics when `skip_footer == 0` and pandas Python semantics when
`skip_footer > 0`. Python continues to own source management, schema
compilation, pandas-compatible physical value conversion, Arrow conversion,
and public lifecycle behavior. Official platform wheels contain the native
extension. If native execution is ineligible or unavailable before the full
row pass begins, the operation uses the exact materialized pandas fallback.
The already-approved bounded stable-schema rule and the late path-decoding
rule are the only native-streaming compatibility exceptions.

## Goals

1. Preserve the existing public API and the materialized `CSVHandler` values,
   physical scalar types, columns, warnings, and error classes for
   schema-compatible streams, subject only to the explicitly documented
   bounded stable-schema and late path-decoding exceptions.
2. Produce stable Arrow schemas before a public stream is returned.
3. Yield at most `batch_size` accepted rows per public batch.
4. Retain at most `batch_size + skip_footer` completed logical records, plus
   the current record and a fixed byte buffer.
5. Stop reading and parsing immediately after the requested public batch
   becomes releasable; footer mode may retain at most `skip_footer`
   already-parsed successor rows.
6. Preserve exact caller-stream cursor and ownership behavior.
7. Keep full-pass memory proportional to the public batch, footer, current
   record, and field payloads rather than source size.
8. Accelerate quoted, CRLF, multiline, and malformed CSV workloads that are
   slow under the Python framing implementation.
9. Ship supported native wheels for CPython 3.11 through 3.14 on Linux, macOS,
   and Windows.
10. Preserve installation and exact behavior on unsupported environments and
    future CPython versions through a materialized fallback path.

## Non-goals

1. The native layer will not infer numeric, boolean, temporal, Arrow, or public
   display types.
2. It will not depend on pandas, NumPy, or PyArrow C APIs.
3. It will not expose a new public tokenizer API.
4. It will not replace spreadsheet parsing, normalization, or source
   management.
5. It will not switch from native to materialized parsing after the full-pass
   borrow begins or after emitting output. A restored bounded-evidence borrow
   may select materialized fallback.
6. It will not support free-threaded `cp313t` or `cp314t` builds in v1.0.0.
7. It will not attempt to make arbitrarily large logical records consume
   constant bytes; one current record and retained footer records are inherent
   payload costs.

## Global constraints

- Python versions: CPython 3.11, 3.12, 3.13, and 3.14.
- Pandas runtime for v1.0.0: `pandas==3.0.5`.
- Primary semantic oracle: pandas 3.0.5 materialized `read_csv` behavior.
- Public release target: messy-xlsx v1.0.0.
- Native ABI floor: CPython 3.11 Stable ABI (`cp311-abi3`).
- Native implementation language/build pins: `Cython==3.2.9`,
  `setuptools==83.0.0`, `cibuildwheel==4.1.1`, and `abi3audit==0.0.26`.
- Native build backend: `setuptools.build_meta`.
- Native wheel builder: cibuildwheel.
- No pandas, NumPy, or PyArrow headers in the extension.
- No unbounded complete-source `bytes`, decoded `str`, `StringIO`, or DataFrame
  copy in the native full pass. Bounded evidence may equal an entire small
  source only while remaining inside every hard evidence budget.
- No unbounded `read()` request against caller-owned streams.
- A bounded evidence pass may select materialized fallback after restoring its
  borrow. No native-to-fallback transition is allowed after the full-pass
  borrow begins.
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
  ├─ capability selection
  ├─ bounded evidence borrow
  │    ├─ COMPLETE → compile stable schemas
  │    ├─ eligible no-footer SAMPLE_FULL → compile stable schemas
  │    └─ incomplete footer, any generated multi-header, unsupported evidence,
  │         or BUDGET_EXHAUSTED → restore borrow → materialized fallback
  │
  └─ first public batch request
       └─ open full-pass borrow → bind NativeCSVTokenizer
            ├─ fixed undecoded byte buffer
            ├─ one current logical record
            ├─ parsed footer deque ≤ skip_footer
            └─ accepted output ≤ batch_size
                 ↓
          PandasCSVValueAdapter
                 ↓
          existing Arrow physical encoder
                 ↓
          existing normalization pipeline
                 ↓
          public RecordBatch

Native ineligible/unavailable before the full-pass borrow
  └─ existing materialized pandas compatibility reader
```

The native extension is an internal component named
`messy_xlsx._csv_tokenizer`. It owns record framing and every decision that
affects which raw fields reach Python. It returns textual field values and
structural missing values; Python maps them to the sampled pandas physical
scalar types and passes those values through the existing stable-schema Arrow
pipeline.

Capability selection happens before source access. A bounded evidence pass may
select materialized fallback only after its borrow is closed and the source is
restored/replayable, before a public reader exists. The no-switch rule begins
when the full-pass borrow opens on first iteration.

## Component boundaries

### Native evidence and resolved configuration

Configuration is deliberately split so schema evidence does not depend on
values that evidence itself must discover:

```python
class NativeSemanticEngine(Enum):
    C = "c"
    PYTHON = "python"


@dataclass(frozen=True, slots=True)
class NativeEvidenceLimits:
    requested_data_rows: int
    max_records_examined: int
    max_payload_bytes_examined: int
    max_cells_examined: int
    max_replay_bytes: int
    max_retained_cells: int


@dataclass(frozen=True, slots=True)
class CSVHeaderPlan:
    pandas_header_mode: Literal["named", "none"]
    pandas_skiprows: int
    post_parse_skip_rows: int
    generated_header_rows: int
    header_present: bool


@dataclass(frozen=True, slots=True)
class NativeCSVFramingConfig:
    encoding: str
    encoding_errors: Literal["strict", "ignore"]
    delimiter: str
    quote_char: str
    double_quote: bool
    skip_initial_space: bool
    skip_blank_lines: bool
    semantic_engine: NativeSemanticEngine
    header_plan: CSVHeaderPlan
    skip_footer: int
    source_description: str
    fallback_encoding_selected: bool


class NativeEvidenceStatus(Enum):
    COMPLETE = "complete"
    SAMPLE_FULL = "sample_full"
    BUDGET_EXHAUSTED = "budget_exhausted"


@dataclass(frozen=True, slots=True)
class NativeEvidence:
    status: NativeEvidenceStatus
    pandas_replay: NativeEvidenceReplay
    raw_data_rows: tuple[tuple[str | None, ...], ...]
    typed_data_rows: tuple[tuple[object, ...], ...]
    column_names: tuple[Hashable, ...]
    physical_lines: tuple[int, ...]
    leading_index_fields: int
    parser_diagnostics: tuple["NativeCSVWarning", ...]
    target_data_rows: int
    records_examined: int
    payload_bytes_examined: int
    cells_examined: int
    replay_bytes_retained: int
    cells_retained: int
    eof: bool


@dataclass(frozen=True, slots=True)
class ResolvedNativeCSVConfig:
    framing: NativeCSVFramingConfig
    expected_fields: int
    leading_index_fields: int
    operation_max_rows: int | None
    value_converters: tuple[PandasValueConverter, ...]
    column_names: tuple[Hashable, ...]
    bad_line_policy: Literal["warn", "error"]
```

`PandasValueConverter` is an internal immutable descriptor. Its value kind is
one of `INT64`, `UINT64`, `FLOAT64`, `BOOL`, `TEXT`, `OBJECT_INTEGER`,
`OBJECT_BOOLEAN`, or `OBJECT_TEXT`; its missing representation is one of
`FLOAT_NAN`, `PANDAS_NA`, or `NONE`. Object columns whose nonmissing evidence
contains heterogeneous scalar families are not assigned a lossy generic
converter: they select materialized fallback before a public stream exists.
Arbitrary-precision Python integers use `OBJECT_INTEGER`; boolean-with-missing
uses `OBJECT_BOOLEAN`; object-backed text uses `OBJECT_TEXT`.

The Python evidence adapter—not the native tokenizer—is the scalar authority.
It gives pandas 3.0.5 a bounded replay of the exact original header, accepted
sample records, and any retained footer records chosen by the native evidence
scanner, then invokes the corresponding `CSVHandler` parsing/value kwargs
while capturing warnings. The replay preserves original field bytes, quoting,
and record terminators; it is not a reserialization of decoded field strings.
For no-footer input this bounded pandas frame defines the approved stable
streaming schema. Footer input requires the native scanner to reach EOF inside
the work budgets so it can choose the true retained footer, but the replay
still omits discarded middle records. Generated multi-header input always
selects materialized fallback in v1.0.0 because its rendered names may depend
on whole-column pandas inference. The adapter records both native raw lexemes
and pandas-typed scalar rows, classifies each pandas `Series` and its nonmissing
Python scalar families, and compiles exact converters. Unsupported pandas
extension dtypes or heterogeneous object evidence also select materialized
fallback. This removes any circular requirement to convert text before the
authoritative typed sample exists.

Production limits are exact:

```python
NativeEvidenceLimits(
    requested_data_rows=(
        1_000 if operation_max_rows is None else min(1_000, operation_max_rows)
    ),
    max_records_examined=1_000_000,
    max_payload_bytes_examined=256 * 1024**2,
    max_cells_examined=16_000_000,
    max_replay_bytes=8 * 1024**2,
    max_retained_cells=1_000_000,
)
```

For `operation_max_rows == 0`, `requested_data_rows` is zero. Once named data
width is established, the scanner sets returned `target_data_rows` to the
lesser of that request and the rows that fit `max_retained_cells`. The first
three limits bound work performed while scanning; the final two independently
bound retained original replay bytes and retained decoded cells. They are hard
budgets, not alternate sample targets. Construction rejects negative targets
or limits.

`scan_evidence(source, framing, limits)` returns a `NativeEvidenceStatus`,
raw and pandas-typed data rows, final column names, physical-line metadata,
parser outcomes, inferred-index evidence, and exact budget consumption.
`SAMPLE_FULL` means the requested resolved
`target_data_rows` accepted data sample was obtained without EOF after all
header, row, malformed-record, and footer stages and after the column schema is
fixed. `COMPLETE` means EOF was observed, even when it supplies fewer than
`target_data_rows`.
`BUDGET_EXHAUSTED` is not EOF and routes to materialized fallback after the
evidence borrow is restored.

Evidence callbacks collect diagnostics but emit no caller-visible warnings.
Returned `target_data_rows` counts accepted post-header data rows. Header, blank,
malformed, and footer rows may require more logical records; if exact evidence
cannot be obtained inside every budget, the status is
`BUDGET_EXHAUSTED`.

`NativeEvidenceReplay` is an internal immutable sequence of original byte
fragments plus stage metadata sufficient to present the selected header,
sample, and footer records to pandas without changing their quoting or line
endings. Its retained bytes count inside `max_replay_bytes`.

`NativeEvidence` owns immutable Python copies independent of native buffers.
`eof` is true only with `status == COMPLETE`; `SAMPLE_FULL` and
`BUDGET_EXHAUSTED` are never presented as EOF. Budget checks are incremental:
before examining the next decoded byte, examining or retaining the next cell,
retaining the next replay byte, or examining the next completed record, the
scanner verifies the corresponding limit. It
may return `BUDGET_EXHAUSTED` with an incomplete current record and need not
find that record's terminator. Partial-record state is discarded; retained
replay bytes/cells and examined bytes/cells/records never exceed their
respective limits. Fixed-buffer prefetched but unread bytes are excluded from
`payload_bytes_examined` and are neither decoded nor parsed. If
`requested_data_rows > 0` but one complete data row cannot fit after width
resolution, the result is `BUDGET_EXHAUSTED`, not a zero-row `SAMPLE_FULL`.

Python constructs `CSVHeaderPlan` from `ParseOptions` and resolved metadata
skipping before evidence. The Python evidence adapter then combines native
structural evidence, its exact pandas frame, and that plan to compile column
names, value converters, and `ResolvedNativeCSVConfig`. The tokenizer does not
invent a simpler protected-record model.

### Native tokenizer

The internal interface is equivalent to:

```python
class NativeCSVTokenizer:
    def __init__(self, config: ResolvedNativeCSVConfig) -> None: ...
    def bind(self, source: BinaryIO) -> None: ...
    @property
    def debug_state(self) -> NativeDebugState: ...
    def _set_debug_observer_for_tests(
        self,
        observer: Callable[[str, NativeDebugState], None] | None,
    ) -> None: ...
    def read_batch(
        self,
        requested_rows: int,
        on_warning: Callable[[NativeCSVWarning], None],
    ) -> NativeCSVRead: ...
    def close(self) -> None: ...
```

Preparation opens no full-pass borrow and never calls `bind()`. The first
`read_batch()` enters the long-lived `SourceHandle` borrow, binds once, and
begins reading. `requested_rows` must be positive. A call returns one through
`requested_rows` rows, or zero rows only with `done=True`. The separate
`operation_max_rows` limit counts accepted rows over the complete operation.
`operation_max_rows == 0` returns an already-complete reader and never opens or
binds a full-pass borrow.

`NativeCSVRead` contains independently owned column-major `list[str | None]`
values, its row count, `done`, `source_eof`, and an immutable
`NativeDebugState`. `done` means the operation is complete because source EOF
was observed or `operation_max_rows` was reached. `source_eof` reports only
physical EOF. A final nonempty batch carries `done=True` when the operation
limit is reached; without lookahead, a batch ending exactly at physical EOF may
carry `done=False` until the next call observes EOF. `None` means a structurally
missing/padded field. Empty fields remain `""` until the Python value adapter
applies pandas rules.

`NativeCSVWarning` is immutable and contains a kind (`"field_count"` or
`"parser_error"`), one-based physical line, one-based logical record, optional
expected/actual field counts, and the exact pandas-compatible warning message.
`NativeDebugState` contains output rows, retained footer rows, current-record
activity, unread buffer bytes, retained field-tokenized successor rows, and
logical payload-byte counters and allocation bytes. The read-only
`debug_state` property is explicitly safe while `READING`, including when
called reentrantly from a test source's `read()` method or the warning
callback; all mutating/re-entrant methods remain forbidden. This provides a
snapshot from inside every source-read and warning callback. The internal
test-only observer receives immutable snapshots immediately before and after
each Python callback and immediately before a public return. It defaults to
`None`, is never used by production routing, and does not recursively observe
its own invocation. An observer failure follows the ordinary callback-failure
terminal path.

The source protocol is exact: native code calls only `source.read(size)` with
`1 <= size <= READ_BUFFER_CAPACITY`; accepts `bytes`, `bytearray`, or
one-dimensional C-contiguous byte-format `memoryview`; accepts partial reads;
treats a zero-length result as EOF; rejects a result longer than `size` before
copying; and never calls `seek`, `tell`, or `close`. Mutable or exported inputs
are copied into owned fixed storage while the GIL is held before another
callback or any `nogil` loop. Non-contiguous, multidimensional, or non-byte
memoryviews raise a contextual source `TypeError`. Returned values own their
Python objects independently of native buffers.

`on_warning` is synchronous. This avoids an unbounded warning queue when many
malformed records precede the next accepted row. A callback exception makes
the tokenizer terminal and releases its native state during close. The adapter
translates an ordinary warning-emission failure, including a `ParserWarning`
promoted to an exception by the caller's warning filter, through the same
contextual `FormatError` boundary as materialized `CSVHandler`; process
failures propagate unchanged.

### Python native adapter

`csv_native.py` owns:

- native-module selection;
- bounded evidence routing and construction of `ResolvedNativeCSVConfig`;
- binding the active `SourceHandle` borrow;
- warning translation to `pandas.errors.ParserWarning`;
- translation of ordinary native failures to `FormatError`;
- propagation of process failures;
- close ordering and metrics;
- conversion of `NativeCSVRead` into the existing physical reader contract.

The adapter never closes a caller-owned stream. It closes native allocations
before the `SourceHandle` borrow restores the caller's entry cursor.

### Pandas-compatible value adapter

`csv_value_adapter.py` receives the compiled `PandasValueConverter` for every
named column. It first maps the union of pandas 3.0.5 default NA strings and
configured `na_values` to the converter's missing representation using exact
string comparisons, then converts compatible lexemes to the sampled pandas
scalar representation.

- Quoted and unquoted empty fields follow the pandas oracle.
- Surrounding whitespace prevents a match unless the configured parsing
  behavior removes that whitespace.
- Structurally missing short-row fields are already `None`.
- Signed/unsigned and arbitrary-precision integers, floats, booleans,
  strings/object text, boolean-with-missing, and missing-promoted numeric
  columns match the supported pandas 3.0.5 physical scalar families.
- Actual named CSV header cells and discarded implicit-index fields bypass
  data conversion.
- Generated multi-header rows are ordinary pandas data rows before
  `_generate_column_names`; in v1.0.0 the materialized `CSVHandler`, not the
  native value adapter, performs their authoritative pandas physical conversion
  and name rendering.
- A later non-null lexeme incompatible with the fixed sampled type raises
  contextual `StreamingTypeError` at its absolute accepted-row offset.

Date inference, public normalization, all-null preservation, sanitization,
Arrow encoding, and display-name decisions remain in the existing
normalization pipeline. `normalize=False` receives the pandas-compatible
physical scalars produced here.

Bounded evidence deliberately fixes the physical schema before the stream is
returned. A later incompatible lexeme raises `StreamingTypeError`; native
streaming does not reproduce pandas whole-file late dtype widening or its
associated `DtypeWarning`. This is the already-approved stable-schema
compatibility exception from the parent parser-performance design.

## Tokenization contract

The normative oracle is `CSVHandler` under `pandas==3.0.5` with its exact
kwargs. The semantic engine is C when `skip_footer == 0` and Python when
`skip_footer > 0`. References to C behavior apply only to the no-footer
branch.

Both semantic modes own:

- incremental decoding with strict and ignore modes;
- UTF BOM handling and chunk-split code units;
- arbitrary supported one-character delimiters;
- quote, doubled-quote, embedded newline, CRLF, CR-only, and EOF behavior;
- mode-specific NUL, malformed-quote, and post-quote-junk behavior;
- the exact `CSVHeaderPlan` stage order;
- mode-specific blank-row classification;
- first-surviving-row width and implicit-index decisions;
- right-padding short rows;
- rejecting later rows wider than the established physical width;
- Python-engine footer retention;
- accepted-row `max_rows`;
- synchronous bad-row line information;
- discarding inferred index fields while returning named data fields.

Mode-specific rules are normative:

- C mode has no footer. It reproduces pandas C NUL truncation, quote-junk,
  unterminated-EOF, implicit-index, warning, and bad-line behavior.
- Python mode first parses physical input with pandas Python semantics.
  `csv.Error` records warn and are discarded before footer removal. The final
  `skip_footer` successfully parsed rows are then removed; at that stage blank
  and over-wide rows may still occupy footer slots. Blank removal and
  field-width validation follow the oracle order.
- In Python mode, `pandas_skiprows` counts physical input lines and may bisect
  a multiline quoted record, exactly as pandas 3.0.5 does.
- Implicit-index width is established by the first pandas-eligible data row
  before `post_parse_skip_rows` and generated-header consumption, even when a
  later stage removes that row.
- In multi-header mode, footer and pandas row limits run before
  `post_parse_skip_rows` and `generated_header_rows`.
- With `pandas_header_mode == "none"`, the first surviving row establishes
  data width and remains in the data sequence; footer or row-limit stages may
  still remove it. Header-based implicit-index inference is disabled.

Critical oracle cases include:

- The first data row may establish one or more implicit leading index fields.
- A later excess-field row is rejected regardless of public batch position.
- In C mode, `b'a,b\r x ,q"z,q"z\r'` matches whole-file pandas C output.
- `1,x\x00junk` is accepted as `1,x`, while `2,y\x00,z` is classified using
  C-mode pandas NUL and delimiter rules.
- In Python mode, a trailing wide row may occupy a footer slot and disappear
  without a warning.
- In Python mode, a trailing quote-error row warns and is discarded before
  footer removal, so an earlier successfully parsed row becomes the footer.
- NUL, malformed quote, CR-only, blank, and implicit-index fixtures are
  tested in both semantic modes because their results can differ.

## Streaming and memory invariants

Terms used by the counters are exact:

- **framed:** the complete raw record boundary is known;
- **field-tokenized:** raw content has been split/interpreted into fields;
- **classified:** blank, parser-error, width, and implicit-index rules have
  been applied;
- **accepted:** the row survives all engine/header/footer rules and counts
  against operation/public row limits.

At every native callback and public return:

- accepted output records retained are at most `batch_size`;
- successfully parsed footer rows retained are at most `skip_footer`;
- completed records retained are at most `batch_size + skip_footer`;
- exactly one additional current logical record may be under construction;
- the undecoded read buffer has fixed
  `READ_BUFFER_CAPACITY == 64 * 1024` bytes;
- after the Nth accepted output row becomes releasable, the tokenizer performs
  no additional source read, record framing, field tokenization, or callback;
- at return, `post_output_rows_retained <= skip_footer`;
- field-tokenized successful successor rows retained in the footer are at most
  `skip_footer`; malformed parser-error records may be consumed and discarded
  while making an output row releasable, but retain no completed-row payload;
- the source cursor may be ahead only by unread bytes already fetched into the
  fixed buffer, and those bytes have caused no decoding or parser callback.

`NativeDebugState` exposes exact logical ownership counters for
`current_record_payload_bytes`, `footer_payload_bytes`,
`output_payload_bytes`, and `undecoded_buffer_bytes`. Their sum is the native
retained payload at that observation point. Returned output objects transfer
ownership out of tokenizer state before the next observation.

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
must eventually represent its payload. `native_allocation_bytes` and measured
process/Python peaks are report-only because Python object and allocator
overhead is platform-dependent; acceptance asserts the exact logical
row/buffer/payload ownership bounds rather than a fictitious cross-platform
total-allocation formula.

## Sampling and stable schema

Public streams still compile configuration and stable schemas before return.
Native bounded evidence uses the same semantic-engine, header-plan, footer,
blank, malformed, implicit-index, and physical-value rules as the full pass.
It runs inside
`NativeEvidenceLimits(requested_data_rows, max_records_examined,
max_payload_bytes_examined, max_cells_examined, max_replay_bytes,
max_retained_cells)`. Footer lookahead work and retained values count toward
every applicable budget. It
does not create an unbounded source copy; a complete small-source replay is
permitted only within every hard budget. It never converts budget exhaustion
into synthetic EOF.

If exact footer/header evidence cannot complete within all limits,
`BUDGET_EXHAUSTED` selects the materialized fallback after evidence cleanup
and cursor restoration, before a public stream exists. It does not infer from
rows that might all be removed by later stages. `COMPLETE` pins the legacy
empty, blank-only, header-only, duplicate/unnamed-header, all-bad, and
all-footer schema or error behavior.

When `skip_footer > 0`, reaching `target_data_rows` is not sufficient for
`SAMPLE_FULL`: the evidence scanner continues to physical EOF or a hard budget
so it can identify the true footer. Only `COMPLETE` is native-eligible in that
mode.

The full row pass starts from a fresh borrow on first iteration. Evidence
never shares partially consumed tokenizer state with output.

## Header and stage order

`CSVHeaderPlan` reproduces the existing two materialized branches.

For `header_rows <= 1`:

- `pandas_header_mode` is `"named"` when `header_rows == 1`, otherwise
  `"none"`;
- the resolved metadata skip count is applied as pandas `skiprows`;
- in Python mode, `skiprows` counts physical lines and may bisect a multiline
  quoted record;
- footer, bad-line, blank, index, and accepted-row behavior follows the chosen
  semantic engine;
- when the mode is `"none"`, generated names are `col_0`, `col_1`, and so on.

For `header_rows > 1`:

- the pandas-equivalent pass uses `header=0` and `skiprows=0`;
- pandas footer and row-limit behavior occurs first;
- implicit-index inference occurs on the first pandas-eligible data row before
  messy-xlsx post-processing;
- only then does messy-xlsx apply `post_parse_skip_rows == options.skip_rows`;
- the next `generated_header_rows == header_rows` accepted, pandas-typed
  DataFrame rows are consumed to generate names;
- too few surviving rows preserves the current materialized exception instead
  of manufacturing an empty schema.

The tokenizer and adapter keep these phases distinct. They do not convert
generated multi-header rows into protected raw records. Because pandas renders
their values using physical types inferred from the whole column,
`generated_header_rows > 0` selects materialized fallback in v1.0.0 rather than
guessing column names. The engine-specific multi-header behavior remains in
the differential oracle so a future bounded solution cannot weaken it.

## Row limits and footer order

C mode has no footer and counts `operation_max_rows` after blank and bad-line
removal. Python mode with a footer follows the engine-specific parsing/footer
order above. The established low-level `max_rows + skip_footer` pandas error is
rejected before evidence or tokenizer startup.

## Encoding behavior

Path inputs use strict decoding during bounded evidence. If a decode error
occurs within evidence, Python performs the established one-time Latin-1
fallback and recomputes metadata/header evidence. Caller-owned streams retain
their existing ignore/error policy and do not retry.

After the public stream is returned, no encoding fallback is allowed. A late
decode error is lazy, contextual, and terminal. Bytes fetched into the fixed
buffer but not yet decoded do not raise until demand reaches them.

This late path behavior is an explicit compatibility exception: materialized
`CSVHandler` may reach EOF, restart the whole path as Latin-1, and reinterpret
earlier bytes. A streaming reader cannot revise already emitted rows. The
exception is documented for v1.0.0 and has dedicated compatibility tests.

When bounded path evidence selects the legacy Latin-1 fallback, the recomputed
evidence and resolved configuration also select the legacy fallback bad-line
policy. The existing fallback reader omits `on_bad_lines="warn"`, so its
default error behavior is part of the oracle rather than being silently
normalized to the ordinary warning route. Any ordinary parser failure during
recomputed fallback evidence, its pandas replay, or the native full pass raises
the legacy terminal `FormatError("Cannot read CSV with any encoding")` with the
established `attempted_formats` context and the parser failure chained as its
cause. Process failures propagate unchanged.

The tokenizer may use Python incremental codecs internally for uncommon
encodings. Common UTF-8, Latin-1, and UTF-16 paths may receive specialized
Cython loops, provided their oracle behavior is identical.

## Native safety, errors, and lifecycle

The native state machine is:

```text
NEW → BOUND → READING → BOUND
                   └──→ TERMINAL → CLOSED
NEW/BOUND ───────────────────────→ CLOSED
```

`bind()` is one-shot. `read_batch()` is non-reentrant. Recursive
`bind()`, `read_batch()`, or `close()` calls from `source.read()` or
`on_warning()` raise a stable internal state error without changing ownership.

The GIL is held for every Python call and every Python-object allocation. A
`nogil` loop may touch only exclusively owned native storage. No pointer into
an exported or resizable buffer crosses a Python callback.

Raw native allocations use `PyMem_Malloc`, `PyMem_Realloc`, and `PyMem_Free`.
Every `Py_ssize_t`/`size_t` addition and multiplication is checked before
allocation. Reallocation uses a temporary pointer; each allocation has one
owner and one complete unwind path. `__dealloc__` is no-throw and releases all
remaining ownership. Allocation-failure injection covers every allocation
site and transition.

Failure translation is:

| Source | Public behavior |
|---|---|
| Unsupported runtime/configuration or API-version mismatch | Native ineligible before full-pass startup; materialized route plus backend reason metric |
| Invalid internal state/configuration supplied after eligibility | Stable internal exception; do not hide a programming defect with fallback |
| Evidence budget exhaustion | Restore evidence borrow, then select materialized route |
| Evidence parser/decoder/I/O failure | Eager contextual `FormatError` before public return |
| Full-pass parser/decoder/I/O failure | Lazy contextual `FormatError`, terminal |
| Parser failure after path fallback encoding was selected, in evidence/replay/full pass | Legacy `Cannot read CSV with any encoding` `FormatError` plus `attempted_formats`; process failures unchanged |
| Native warning event | Exact message via `pandas.errors.ParserWarning`, `stacklevel=3` |
| Ordinary warning-emission callback failure | Contextual `FormatError` with the callback failure as cause, terminal |
| Downstream physical incompatibility | Preserve contextual `StreamingTypeError` |
| `MemoryError`, `KeyboardInterrupt`, `SystemExit`, or chained process failure | Propagate unchanged |
| Cleanup failure during another failure | Preserve the active primary/process failure |

Evidence diagnostics never emit public warnings. Full-pass warnings emit once.
Merely prefetched, undecoded bytes cause no error or callback.

`close()` is idempotent and releases native allocations on success, ordinary
failure, process failure, downstream normalization failure, early close, and
exhaustion. The tokenizer closes before the source borrow. A cursor-restoration
failure prevents a batch from being returned, including a `done` final
nonempty batch. The existing Task 13 weak ownership, finalizer, return-gap
cleanup, and one-active-operation behavior remains authoritative.

Task 13 cleanup precedence remains exact: a process-level cleanup failure
replaces an ordinary primary failure; an ordinary cleanup failure never
replaces an active primary; and an already-active process failure remains
authoritative.

## Native selection and fallback

Backend observation uses a typed decision separate from the existing
format-level `ReaderDecision`:

```python
class CSVExecutionKind(Enum):
    NATIVE = "csv_native"
    MATERIALIZED_FALLBACK = "csv_materialized_fallback"
    CUSTOM_SPI = "custom_dataframe"


class CSVExecutionReason(Enum):
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

Exactly one decision is recorded after capability routing and, when applicable,
evidence routing, before a successful reader is returned. The reader retains
its immutable decision.
`ParseMetrics.last_csv_execution` is replaced on each successful CSV operation,
and per-kind/per-reason counters are incremented, so sequential operations do
not leave an ambiguous untyped string. A failed evidence parse increments the
existing failure metric but returns no reader and records no execution
decision. Custom registry execution preserves
`ReaderDecision.backend == CUSTOM_DATAFRAME` and separately records
`CSVExecutionKind.CUSTOM_SPI`; it never mislabels a custom handler as the
built-in pandas fallback.

The production gate is a source-controlled internal constant in
`csv_native.py`:

```python
_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False  # candidate revision
```

The final enablement revision changes only this functional source line to
`True`; documentation and generated release metadata may change separately.
Candidate public built-in CSV operations therefore select
`MATERIALIZED_FALLBACK/PRODUCTION_GATE_DISABLED`.

Candidate artifact smoke uses only the private
`csv_native._run_candidate_artifact_smoke(...)` entry point. Internally it
passes a module-owned `_CANDIDATE_SMOKE_TOKEN` to the adapter constructor,
rechecks exact built-in configuration, runtime, import, and handshake
eligibility, bypasses only the production-ready constant, and never
participates in workbook/public routing. It records `NATIVE/NATIVE_SELECTED` in
its isolated test metrics. No environment variable or public selector can
activate this bypass.

Native eligibility is resolved in this order:

1. The exact built-in `HandlerRegistry`, unchanged built-in components, and
   exact built-in `CSVHandler` must own the detected format. Any registry
   subclass, detector override, handler replacement/subclass, component
   mutation, or `parse` override selects the materialized compatibility SPI.
2. `_NATIVE_CSV_PRODUCTION_READY` must be true; otherwise public execution
   selects the materialized route with `PRODUCTION_GATE_DISABLED`.
3. `MESSY_XLSX_DISABLE_NATIVE` is read once for the operation. Only the exact
   string `"1"` disables native execution; all other and unset values are
   treated as enabled.
4. Runtime must be non-free-threaded CPython 3.11 through 3.14. This guard runs
   before importing the extension. Future CPython versions may install an ABI3
   wheel but execute its bundled materialized path.
5. The extension must import and report
   `NATIVE_API_VERSION == 1` and
   `PANDAS_SEMANTIC_VERSION == "3.0.5"`.

Official supported wheels must contain and load the extension. A missing
extension fails that wheel's smoke test. Capability fallback catches only
`ModuleNotFoundError`, `ImportError`, shared-library `OSError`, a disabled
environment setting, unsupported runtime, or version-handshake mismatch. It
never catches arbitrary exceptions from native execution.

Fallback behavior:

- use the existing materialized pandas compatibility reader;
- preserve values, dtypes, warnings, errors, and public API;
- do not claim native streaming memory guarantees;
- expose the typed execution decision and reason through the reader and
  backend metrics;
- allow deterministic selection with
  `MESSY_XLSX_DISABLE_NATIVE=1`.

The project publishes a universal `py3-none-any` fallback wheel. Compatible
CPython installations prefer the platform `cp311-abi3` wheel. Unsupported and
future runtimes may still install that preferred artifact, so the runtime
guard—not pip tag preference—selects materialized execution.

Emergency rollback uses the environment kill switch immediately, followed by
yanking/replacing affected native wheels. It never retries an operation whose
full-pass borrow has begun.

## Build and release design

The project switches from Hatchling to `setuptools.build_meta` while retaining
PEP 621 metadata and `src` package discovery. The v1.0.0 pandas dependency is
exactly `pandas==3.0.5`.

Build requirements pin `Cython==3.2.9` and `setuptools==83.0.0`. A build-only
`MESSY_XLSX_BUILD_MODE` is either `native` or `fallback`; any other value is a
build error. Official builds always set it explicitly. When absent in a local
source/editable build, it defaults to `native` only on supported non-free-
threaded CPython 3.11–3.14 and to `fallback` elsewhere.

Native mode declares the extension and fails closed on any compilation error:

```text
Py_LIMITED_API = 0x030B0000
Extension(..., py_limited_api=True)
[bdist_wheel] py_limited_api = cp311
wheel tag = cp311-abi3
```

Fallback mode uses `ext_modules=[]`, contains no extension/generated native
artifact, and must produce `py3-none-any` with `Root-Is-Purelib: true`.
The extension does not include pandas, NumPy, or PyArrow headers.

PEP 660 editable installs default to native mode on supported CPython and fail
closed. `MESSY_XLSX_BUILD_MODE=fallback` explicitly creates a fallback
editable install. CI parses CSV and asserts backend selection for both modes.

`cibuildwheel==4.1.1` uses `CIBW_BUILD=cp311-*` to compile ABI3 exactly once
per platform/architecture. The build jobs are deterministic:

- manylinux x86-64 runs on an x86-64 Linux runner with
  `CIBW_ARCHS_LINUX=x86_64` and
  `CIBW_MANYLINUX_X86_64_IMAGE=manylinux2014`;
- manylinux aarch64 runs on a native arm64 Linux runner with
  `CIBW_ARCHS_LINUX=aarch64` and
  `CIBW_MANYLINUX_AARCH64_IMAGE=manylinux2014`;
- musllinux x86-64 and aarch64 run on the corresponding native architecture
  with explicit `musllinux_1_2` images;
- macOS x86-64 uses deployment target 10.13 and macOS arm64 uses target 11.0
  on native architecture runners;
- Windows sets `CIBW_ARCHS_WINDOWS=AMD64` and produces only x86-64.

These jobs produce manylinux_2_17 and musllinux_1_2 x86-64/aarch64, macOS
x86-64/arm64, and Windows x86-64 `cp311-abi3` wheels.

Free-threaded and all other architectures are excluded and use materialized
fallback. Cibuildwheel's build-environment test runs on CPython 3.11. Separate
`abi3-smoke` jobs install the exact already-built artifact by path and exercise
every applicable wheel family on CPython 3.12, 3.13, and 3.14; this verifies
later ABI3 runtimes without rebuilding duplicate wheel filenames. Apple arm64
tests run on an Apple-silicon runner. No claimed combination may use
`allow-empty` or an unreviewed `test-skip`.

The source distribution is built once and contains the `.pyx` source. Its
isolated build installs Cython and requires a platform compiler. The universal
fallback wheel requires no compiler for messy-xlsx itself on unsupported
systems; dependency wheels retain their own published platform requirements.

Every native and fallback wheel is built from a clean extraction of that exact
source archive. The release set contains exactly seven native wheels, one
universal wheel, and one source archive. Release jobs enforce the nine-artifact
filename/count allowlist, cross-wheel `METADATA` parity, sdist source inventory,
wheel content and tag checks, `abi3audit==0.0.26 --strict`, `twine check`, and
`pip check`.
Smoke installations use direct artifact paths or an isolated
`--no-index --find-links` wheelhouse outside the source tree.

Candidate and final outputs use separate immutable artifact namespaces because
their filenames and version tags coincide. Manifests record the SHA-256 of
each source archive and wheel; final jobs verify every final wheel was built
from the final source archive and never merge a candidate artifact. Resolver
tests also prove unsupported-platform and free-threaded tags select the
universal fallback wheel.

Candidate native-wheel smoke tests assert extension presence and direct
internal-adapter CSV parsing on all supported runtimes while the production
gate remains disabled. Final native-wheel smoke tests repeat those assertions
and require public built-in CSV operations to select native. Fallback wheel
tests assert extension absence and automatic materialized parsing. Final native
wheels also test the runtime kill switch. Resolver/tag tests prove native
preference on supported tags and the runtime guard on unsupported future
CPython.

## Verification strategy

### Differential oracle tests

Every semantic fixture compares the materialized `CSVHandler` and native
tokenizer paths using `pandas==3.0.5` across public batch sizes 1, 2, 3, and
127. Every NUL, quote, malformed, CR-only, blank, header, and implicit-index
fixture runs once in C mode and once in Python/footer mode.
For schema-compatible fixtures, assertions cover values, physical scalar
families, columns, warnings, error class, and error context.

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
- integer, missing-promoted numeric, float, boolean, object/string, overflow,
  and late physical-type incompatibility values with `normalize=False`;
- `header_rows` zero, one, and greater than one; physical-line `skiprows`;
  post-parse skipping; generated headers; duplicate/unnamed headers; too few
  surviving header rows; and first-row implicit indexes;
- path, seekable, nonseekable, and one-byte-read sources;
- every custom registry override (`parse`, `detect_format`, `validate`,
  `get_sheet_names`), registry subclasses, detector replacements, handler
  replacements/subclasses, and component mutation;
- records larger than 8 MiB.

Dedicated exception fixtures assert:

- bounded-schema streams yield the successful schema-compatible prefix, then
  raise `StreamingTypeError` at the first incompatible accepted-row offset
  instead of reproducing whole-file widening or `DtypeWarning`;
- a late strict path decode failure remains lazy and terminal instead of
  restarting already-emitted rows as Latin-1;
- warning-as-error filters become contextual `FormatError`;
- a fallback-encoding parser failure during footer evidence/pandas replay uses
  the legacy terminal message and `attempted_formats` context;
- `max_rows` completion before physical EOF, including zero;
- object overflow, boolean-with-missing, and heterogeneous-evidence fallback;
- header-none all-footer input, implicit-index establishment by a row removed
  later, and generated multi-header type rendering.

### Differential fuzzing

Hypothesis generates valid and malformed byte streams and randomized chunk
splits. Schema-compatible cases compare accepted physical scalars, columns,
warnings, physical line information, and error classes against pandas 3.0.5.
C and Python modes each run at least 5,000 generated examples using fixed CI seeds
`0x0C5A14` and `0xBADC5EED`. Every discovered mismatch is minimized and added
to a checked-in byte-fixture regression corpus. Subprocess cases impose an
explicit timeout so crashes and hangs leave durable reproducers. Native and
materialized routes must agree for schema-compatible input before production
routing is enabled. Generated late-incompatible cases separately assert the
stable-prefix/first-`StreamingTypeError` contract; generated late path-decode
cases separately assert the documented lazy exception.

### Lifecycle and failure tests

Tests cover inert construction, first-read startup, early close, exhaustion,
warning-callback failure, decoder/parser/I/O/process failures, cleanup
failures, exact cursor restoration, no caller close, finalizer cleanup, and
active-operation release.

Chained process-failure tests prove that decode or cleanup context never masks
`MemoryError`, `KeyboardInterrupt`, or `SystemExit`.

Native safety gates include:

- Linux ASan and UBSan builds running the full native regression corpus;
- `PYTHONMALLOC=debug` constructor/destructor and lifecycle stress loops;
- C compiler warnings treated as errors for project-generated native code;
- `abi3audit --strict` on every native wheel;
- allocation-failure injection at every `PyMem_*` site;
- callback/source-read reentrancy at every native state transition;
- repeated partial-construction, close, and no-throw `__dealloc__` tests.

### Deterministic bound tests

Debug state exposes:

- `output_rows_retained`;
- `post_output_rows_retained`;
- `current_record_active`;
- `undecoded_buffer_bytes`;
- `field_tokenized_successor_rows`;
- `current_record_payload_bytes`;
- `footer_payload_bytes`;
- `output_payload_bytes`;
- `native_allocation_bytes`.

Tests read the permitted snapshot accessor from source-read and warning
callbacks and assert row, buffer, and logical payload invariants at every
callback and batch return. `native_allocation_bytes` is reported and
cross-checked against allocator hooks but is not treated as a portable
Python-process memory formula. These are architecture tests rather than
timing-based proxies.

### Performance tests

The checked-in benchmark harness uses fixed 300,000-row corpora generated from
seed `0x0C5A14`. Each comparison performs three warmups and seven alternating
measured runs, reports the median, and computes a geometric mean across
corpora. The harness records Python, pandas, native API, compiler, platform,
CPU, and build-mode metadata.

Routing performance thresholds are authoritative on the project's dedicated
Ubuntu 24.04 x86-64 benchmark runner, its pinned image/CPU identity, CPython
3.12, and the installed manylinux native wheel. Results from other supported
platforms are required release evidence but are corroborating rather than
threshold-gating.

Benchmarks cover:

- clean unquoted LF;
- quoted LF;
- CRLF;
- multiline quoted fields;
- sparse malformed rows;
- `batch_size=1`;
- large logical records;
- nonzero footer retention.

Every generated corpus stays below 310,000 logical records, 48 MiB of examined
payload, and 4.8 million examined cells. Its selected header/sample/footer
replay stays below 2 MiB and 100,000 retained cells. The large logical record
appears after the first 1,000 accepted data rows so it exercises full-pass
growth rather than evidence fallback. The footer corpus uses
`skip_footer == 10`, reaches EOF inside the same limits, and must record
`CSVExecutionKind.NATIVE`. Every performance case asserts the native decision
before timing; unexpected fallback invalidates the run instead of silently
benchmarking materialization.

No-footer tokenizer-only cases compare with direct pandas C. Footer cases
compare with both the retained Python streaming reference implementation and
end-to-end materialized `CSVHandler`; they do not label direct pandas C as an
equivalent footer baseline.

Production routing requires:

- no no-footer case slower than 3.0 times direct pandas C;
- no-footer geometric-mean slowdown at most 2.0 times direct pandas C;
- footer geometric-mean speedup at least 4.0 times the retained Python
  streaming reference;
- end-to-end native row, fixed-buffer, and logical payload ownership satisfying
  the deterministic invariants for every corpus.

End-to-end reports include throughput, peak retained rows, logical retained
payload bytes, native allocation bytes, measured process/Python peak memory,
and source position at each public batch. Relative timings are recorded in the
performance report. There is no automatic performance waiver: a miss keeps
production routing disabled until the spec and user-approved plan are amended.

## Supersession and migration

This design supersedes the pandas-chunk full-pass implementation originally
specified in parent Task 14. The parent plan is amended in the same design
revision and the native implementation receives its own task-by-task plan.

The current uncommitted Task 14 work is classified as follows:

- bounded prefix inspection, source borrowing, process-failure fixes, schema
  compilation, lifecycle integration, custom-registry eligibility, and
  compatible tests are retained;
- `csv_streaming.py` remains the Python adapter shell, but its pandas
  full-pass chunk reader is replaced;
- `csv_io.py` framing/filter/footer code is retained only as a differential
  reference during development, moved under tests before production routing,
  and removed from the installed runtime if no remaining production consumer
  exists;
- characterization and adversarial tests are retained, while tests that
  encoded known parity tradeoffs are corrected against the engine-specific
  oracle;
- workbook native routing remains disabled until the semantic, lifecycle,
  safety, bound, wheel, and performance gates pass;
- packaging and release expansion is part of Task 14A and must complete before
  Task 14 is marked done.

## Delivery stages

0. Prove ABI feasibility first: build the exact extension-type, typed
   memoryview, Python `read`, callback, allocation, and cleanup skeleton as
   `cp311-abi3`; run it on every supported runtime/platform; and pass
   `abi3audit --strict`. If Stable ABI fails, use per-minor CPython 3.11–3.14
   wheels and amend the artifact matrix before tokenizer work continues.
1. Freeze the engine-specific pandas 3.0.5 oracle with differential fixtures,
   seeded fuzzing, and the minimized regression corpus.
2. Add the internal API/version handshake, explicit state machine,
   deterministic test double, lifecycle adapter, evidence statuses,
   backend metrics, and materialized selection.
3. Implement C-mode valid-record framing, decoding, quoting, line endings,
   BOMs, blanks, large records, and fixed read-buffer behavior.
4. Implement C-mode compatibility for headers, implicit indexes, short rows,
   NULs, quote junk, excess fields, warnings, and malformed fuzz cases.
5. Implement Python-mode parsing order, physical-line skipping, quote errors,
   footer retention, accepted-row limits, deterministic memory counters, and
   the exact materialized routing decision for generated multi-headers.
6. Add the pandas physical-value adapter and integrate native batches with the
   existing Arrow/normalization pipeline.
7. Pass semantic, lifecycle, safety, bound, and performance gates while default
   production routing remains disabled.
8. Complete candidate native/fallback PEP 660 builds, exact-sdist dual wheel
   builds, cibuildwheel matrices, wheel audits, direct internal-adapter smoke,
   ABI3 runtime smoke, and resolver tests with production routing disabled.
9. Only after the candidate artifact matrix passes, enable exact built-in
   native routing, build a new final sdist and all final wheels from that exact
   revision, and rerun the complete artifact, public-routing, kill-switch,
   resolver, and runtime matrix. Only this second verified set is releasable.
10. Run full repository verification, independent native memory-safety review,
    final compatibility review, and release readiness review.

## Acceptance criteria

The design is complete when:

1. Native and materialized outputs agree across the schema-compatible
   engine-specific oracle matrix and differential fuzzing; late incompatible
   values and late path decoding satisfy their separately documented streaming
   exception contracts.
2. Every public native batch contains at most `batch_size` accepted rows.
3. Deterministic counters prove output/footer/current-record/fixed-buffer and
   logical payload ownership bounds and immediate stopping after the requested
   batch becomes releasable; platform-dependent allocation totals are measured
   and reported.
4. Late errors remain demand-driven and process failures remain unmasked.
5. Caller streams are never closed and their exact entry cursors are restored.
6. Custom registry handlers remain authoritative and materialized.
7. Stable-ABI feasibility or its documented per-minor contingency is proven
   before tokenizer implementation.
8. Official native wheels load on every claimed Python/platform combination.
9. Native and universal wheels select the correct runtime behavior, including
   future-Python guards and the environment kill switch.
10. Full tests, Ruff, formatting, mypy, security checks, docs, wheel smoke
   tests, and source-distribution smoke tests pass.
11. All native safety, performance, and disabled candidate-wheel gates pass
    before default production routing is enabled; the complete final-wheel
    matrix passes again before release.
