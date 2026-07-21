# Configuration

## `SheetConfig`

`SheetConfig` controls how one sheet is parsed into a DataFrame. Pass it to
`MessyWorkbook`, `to_dataframe()`, `to_dataframes()`, or a sheet/table wrapper.
The `read_excel()` convenience function accepts the same fields as keyword
arguments.

```python
from messy_xlsx import SheetConfig

config = SheetConfig(
    # Rows and range
    skip_rows=0,
    header_rows=1,
    skip_footer=0,
    cell_range=None,

    # Structure and headers
    auto_detect=True,
    header_detection_mode="smart",
    header_confidence_threshold=0.7,
    header_fallback="first_row",
    multi_row_headers=False,
    header_patterns=None,

    # Worksheet features
    include_hidden=False,
    merge_strategy="fill",
    evaluate_formulas=True,

    # Normalization
    normalize=True,
    normalize_dates=True,
    normalize_numbers=True,
    normalize_whitespace=True,
    use_extended_missing_list=False,
    preserve_types=True,
    ensure_type_consistency=True,
    decimal_separator=None,
    thousands_separator=None,
    locale=None,

    # Columns and row filters
    sanitize_column_names=True,
    column_renames={},
    type_hints={},
    drop_regex=None,
    drop_conditions=[],
)
```

Every parse compiles the configuration into an independent parse plan. The
library does not mutate a caller-owned `SheetConfig`, including when the same
configuration is reused to read a detected table.

### Rows and ranges

| Parameter | Type | Default | Behavior |
|---|---|---:|---|
| `skip_rows` | `int` | `0` | Rows before the header to discard. Must be non-negative. |
| `header_rows` | `int` | `1` | Rows used to build column names. `0` starts with positional integer columns (the default sanitizer turns these into `col_0`, `col_1`, and so on); values greater than one join header levels with `_`. |
| `skip_footer` | `int` | `0` | Rows to discard from the bottom. Must be non-negative. An explicit value takes precedence over a detected footer. |
| `cell_range` | `str \| None` | `None` | Current-sheet A1 range such as `"A1:F100"`. Range parsing is supported for OOXML workbooks (`.xlsx`/`.xlsm`), not CSV or legacy `.xls`. |

For a deterministic range parse, disable structure detection and state the
header policy explicitly:

```python
config = SheetConfig(
    auto_detect=False,
    cell_range="B4:F100",
    header_rows=1,
)
```

### Structure and header detection

| Parameter | Type | Default | Behavior |
|---|---|---:|---|
| `auto_detect` | `bool` | `True` | Enables OOXML structure analysis and text-file metadata/header detection. Set to `False` to use the row settings exactly. |
| `header_detection_mode` | `HeaderDetectionMode \| str` | `"smart"` | Controls how OOXML header evidence is applied. |
| `header_confidence_threshold` | `float` | `0.7` | Minimum accepted OOXML header confidence, inclusive from `0.0` through `1.0`. |
| `header_fallback` | `HeaderFallback \| str` | `"first_row"` | Action when `"auto"` mode has no sufficiently confident OOXML header. |
| `header_patterns` | `list[str] \| None` | `None` | Additional regular expressions used as evidence by OOXML header analysis. Matching is case-insensitive. |
| `multi_row_headers` | `bool` | `False` | Compatibility field. It does not independently enable multi-row headers in v0.10.0; use `header_rows`, or allow detection to supply its detected header-row count. |

OOXML header modes behave as follows:

| Mode | Behavior |
|---|---|
| `"smart"` | Uses a sufficiently confident detected header when `skip_rows == 0`; otherwise retains the configured row settings. |
| `"auto"` | Uses a sufficiently confident detected header, otherwise applies `header_fallback`. |
| `"manual"` | Retains `skip_rows` and `header_rows`. Use `auto_detect=False` as well when structure analysis itself is not needed. |

`header_fallback` is applied by OOXML `"auto"` mode:

| Fallback | Behavior |
|---|---|
| `"first_row"` | Uses the first row as one header row. |
| `"none"` | Uses no header row. |
| `"error"` | Raises `StructureError` with the detected and required confidence. |

CSV, TSV, and TXT files use their text metadata/header detector when
`auto_detect=True`; the OOXML confidence, mode, fallback, and pattern options do
not govern that detector. Legacy `.xls` parsing uses the configured row values.

### Hidden and merged cells

| Parameter | Type | Default | Behavior |
|---|---|---:|---|
| `include_hidden` | `bool` | `False` | Includes hidden rows and columns when `True`; excludes them by default. |
| `merge_strategy` | `MergeStrategy \| str` | `"fill"` | Controls merged OOXML cells. |

| Strategy | Behavior |
|---|---|
| `"fill"` | Copies the top-left value into every cell in the merged range. |
| `"skip"` | Leaves the worksheet's native values unchanged (non-anchor merged cells are normally empty) and enables the fastest safe reader where possible. |
| `"first_only"` | Keeps only the top-left value and clears the other cells in the range. |

### Formula cells in DataFrames

`SheetConfig.evaluate_formulas` controls the OOXML workbook view used for a
DataFrame parse:

- `True` (default) reads the cached result last saved by Excel or another
  spreadsheet application.
- `False` preserves formula expressions such as `"=A2+B2"`.

Neither setting recalculates a workbook. Cached results can be `None` when the
producer did not save them. `FormulaConfig`, described below, applies to
single-cell access (`get_cell()`), not whole-DataFrame parsing.

### Normalization

| Parameter | Type | Default | Behavior |
|---|---|---:|---|
| `normalize` | `bool` | `True` | Runs the normalization pipeline. In v0.10.0, setting this to `False` also bypasses `drop_regex` and `drop_conditions`; column sanitization and renaming still run. |
| `normalize_whitespace` | `bool` | `True` | Strips and collapses whitespace in text cells. |
| `normalize_numbers` | `bool` | `True` | Converts number-like text, including currency and accounting notation. |
| `normalize_dates` | `bool` | `True` | Converts supported date text and likely Excel serial-date columns. |
| `use_extended_missing_list` | `bool` | `False` | Also treats ambiguous markers such as `-`, `.`, and `?` as missing. Conservative markers such as `N/A`, `null`, and `missing` are always recognized. |
| `preserve_types` | `bool` | `True` | Uses type-appropriate nulls (`None`, `NaN`, or `pd.NA`) during missing-value normalization. |
| `ensure_type_consistency` | `bool` | `True` | Converts mixed object columns to a consistent string representation for Arrow/BigQuery compatibility. |
| `type_hints` | `dict[str, str]` | `{}` | Overrides semantic inference by column. Common hints are `VARCHAR`, `DECIMAL`, and `TIMESTAMP`. |
| `decimal_separator` | `str \| None` | `None` | Explicit decimal separator. |
| `thousands_separator` | `str \| None` | `None` | Explicit thousands separator. |
| `locale` | `str \| None` | `None` | Locale convention such as `"de_DE"` used to derive separators when neither separator is explicit. Use `"auto"` to leave separator detection to the number normalizer. |

The individual normalization switches have an effect only while
`normalize=True`. Explicit number separators take precedence over locale-derived
ones.

### Columns and row filters

| Parameter | Type | Default | Behavior |
|---|---|---:|---|
| `sanitize_column_names` | `bool` | `True` | Produces BigQuery-compatible names and disambiguates duplicates. |
| `column_renames` | `dict[str, str]` | `{}` | Renames columns after sanitization. Keys therefore refer to sanitized names when sanitization is enabled. |
| `drop_regex` | `str \| None` | `None` | Drops a row when the regular expression matches the string form of any non-null value. Invalid patterns raise `re.error` during parsing. |
| `drop_conditions` | `list[dict[str, Any]]` | `[]` | For each `{"column": name, "value": value}`, drops rows whose named column equals the value. Conditions run in order and refer to final, renamed column names. Missing columns are ignored. |

```python
config = SheetConfig(
    column_renames={"order_status": "status"},
    drop_regex=r"(?i)^grand total$",
    drop_conditions=[{"column": "status", "value": "void"}],
)
```

## Enum values

`SheetConfig` accepts either its `StrEnum` members or their string values and
coerces strings during construction:

```python
from messy_xlsx import MergeStrategy, SheetConfig

assert SheetConfig(merge_strategy="fill").merge_strategy is MergeStrategy.FILL
```

| Enum | Members and values |
|---|---|
| `MergeStrategy` | `FILL="fill"`, `SKIP="skip"`, `FIRST_ONLY="first_only"` |
| `HeaderDetectionMode` | `SMART="smart"`, `AUTO="auto"`, `MANUAL="manual"` |
| `HeaderFallback` | `FIRST_ROW="first_row"`, `NONE="none"`, `ERROR="error"` |
| `DataType` | `EMPTY="empty"`, `TEXT="text"`, `NUMBER="number"`, `BOOLEAN="boolean"`, `DATE="date"`, `ERROR="error"`, `FORMULA="formula"` |
| `FormatType` | `XLSX="xlsx"`, `XLSM="xlsm"`, `XLSB="xlsb"`, `XLS="xls"`, `CSV="csv"`, `TSV="tsv"`, `TXT="txt"`, `UNKNOWN="unknown"` |

## `FormulaConfig`

`FormulaConfig` controls formula handling for OOXML single-cell access through
`MessyWorkbook.get_cell()`, `get_cell_by_ref()`, `MessySheet.get_cell()`, and
sheet indexing. It does not change DataFrame parsing.

```python
from messy_xlsx import (
    CircularRefStrategy,
    FormulaConfig,
    FormulaEvaluationMode,
    MessyWorkbook,
)

formula_config = FormulaConfig(
    mode=FormulaEvaluationMode.CACHED_WITH_FALLBACK,
    circular_strategy=CircularRefStrategy.ERROR,
    max_iterations=100,
    max_depth=1000,
    unsupported_value="#UNSUPPORTED",
    raise_on_unsupported=False,
)

with MessyWorkbook("book.xlsx", formula_config=formula_config) as workbook:
    cell = workbook.get_cell("Sheet1", row=2, col=3)
```

Unlike the `SheetConfig` enums, formula configuration uses regular `Enum`
classes; pass enum members rather than raw strings.

### Evaluation modes

| Mode | Value | Behavior of `CellValue.value` for a formula cell |
|---|---|---|
| `DISABLED` | `"disabled"` | Preserves the formula expression; no cached lookup or live evaluation is attempted. |
| `CACHED_ONLY` | `"cached_only"` | Returns only the workbook's last-saved cached value. |
| `CACHED_WITH_FALLBACK` | `"fallback"` | Default. Returns a non-null cached value; otherwise attempts live evaluation. |
| `ALWAYS_EVALUATE` | `"evaluate"` | Attempts live evaluation even when a cached value exists. |

Live evaluation is optional and requires an installed formula backend from the
`formulas` extra. It is loaded for path-backed workbooks. Formula coverage is
limited by those backends; this library is not a complete Excel calculation
engine. When no backend can produce a value, `unsupported_value` is returned,
or `UnsupportedFunctionError` is raised when `raise_on_unsupported=True`.

`CellValue.formula` retains the original expression independently of the
resolved `CellValue.value`.

### Circular references and limits

| Setting | Default | Behavior |
|---|---:|---|
| `circular_strategy` | `CircularRefStrategy.ERROR` | `ERROR` raises `CircularReferenceError`; `RETURN_CACHED` uses a cached value or raises when none exists; `ITERATE` uses an evaluator-cached value, then a workbook-cached value, then `0`. |
| `max_depth` | `1000` | Raises `FormulaError` when the library evaluation stack exceeds this depth. Must be at least `1`. |
| `max_iterations` | `100` | Validated compatibility limit for iterative evaluation. The v0.10.0 engine does not run its own iterative recalculation loop. Must be at least `1`. |

## Construction validation

Construction validates enum strings and numeric bounds that have explicit
constraints:

```python
from messy_xlsx import FormulaConfig, SheetConfig

# Each raises ValueError:
SheetConfig(skip_rows=-1)
SheetConfig(header_rows=-1)
SheetConfig(skip_footer=-1)
SheetConfig(header_confidence_threshold=1.5)
SheetConfig(merge_strategy="invalid")
FormulaConfig(max_iterations=0)
FormulaConfig(max_depth=0)
```

Other fields are not exhaustively schema-validated at construction time; for
example, invalid regular expressions fail when parsing reaches the row-filter
step.
