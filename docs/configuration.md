# Configuration

## SheetConfig

The primary configuration object for controlling how sheets are parsed.

```python
from messy_xlsx import SheetConfig

config = SheetConfig(
    # Row handling
    skip_rows=0,
    header_rows=1,
    skip_footer=0,
    cell_range=None,

    # Header detection
    auto_detect=True,
    header_detection_mode="smart",
    header_confidence_threshold=0.7,
    header_fallback="first_row",

    # Merged cells
    merge_strategy="fill",
    include_hidden=False,

    # Normalization
    normalize=True,
    normalize_dates=True,
    normalize_numbers=True,
    normalize_whitespace=True,
    sanitize_column_names=True,

    # Post-processing
    column_renames={},
    type_hints={},
    drop_regex=None,
    drop_conditions=[],

    # Locale & formulas
    locale=None,
    evaluate_formulas=True,
)
```

### Row Handling

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `skip_rows` | `int` | `0` | Number of rows to skip at the top |
| `header_rows` | `int` | `1` | Number of header rows (0 = no header) |
| `skip_footer` | `int` | `0` | Number of rows to skip at the bottom |
| `cell_range` | `str \| None` | `None` | Excel range to read (e.g., `"A1:F100"`) |

### Header Detection

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `auto_detect` | `bool` | `True` | Enable automatic structure detection |
| `header_detection_mode` | `HeaderDetectionMode` | `"smart"` | Detection strategy |
| `header_confidence_threshold` | `float` | `0.7` | Minimum confidence (0.0–1.0) |
| `header_fallback` | `HeaderFallback` | `"first_row"` | Fallback when detection fails |

#### Header Detection Modes

| Mode | Behavior |
|------|----------|
| `"smart"` | Use detection unless user explicitly set `skip_rows` |
| `"auto"` | Always trust detection if confidence meets threshold |
| `"manual"` | Ignore detection, use config values only |

#### Header Fallbacks

| Fallback | Behavior |
|----------|----------|
| `"first_row"` | Use the first row as header |
| `"none"` | No header row (generate `col_0`, `col_1`, ...) |
| `"error"` | Raise `StructureError` if detection fails |

### Merge Strategies

| Strategy | Behavior |
|----------|----------|
| `"fill"` | Fill all cells in merged range with the top-left value |
| `"skip"` | Leave merged cells as-is (fastest) |
| `"first_only"` | Keep value in top-left cell, set others to `None` |

### Normalization

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `normalize` | `bool` | `True` | Master switch for all normalization |
| `normalize_dates` | `bool` | `True` | Convert date-like columns to `datetime` |
| `normalize_numbers` | `bool` | `True` | Convert number-like strings to numeric |
| `normalize_whitespace` | `bool` | `True` | Strip and collapse whitespace |
| `sanitize_column_names` | `bool` | `True` | BigQuery-compatible column names |

## Enum Types

All string-based configuration values have corresponding enum types. Both forms
are accepted — raw strings are automatically coerced to enums:

```python
from messy_xlsx import SheetConfig, MergeStrategy

# These are equivalent:
config1 = SheetConfig(merge_strategy="fill")
config2 = SheetConfig(merge_strategy=MergeStrategy.FILL)

# Enum values compare equal to strings:
assert MergeStrategy.FILL == "fill"  # True
```

### Available Enums

| Enum | Values |
|------|--------|
| `MergeStrategy` | `FILL`, `SKIP`, `FIRST_ONLY` |
| `HeaderDetectionMode` | `SMART`, `AUTO`, `MANUAL` |
| `HeaderFallback` | `FIRST_ROW`, `NONE`, `ERROR` |
| `DataType` | `EMPTY`, `TEXT`, `NUMBER`, `BOOLEAN`, `DATE`, `ERROR`, `FORMULA` |
| `FormatType` | `XLSX`, `XLSM`, `XLSB`, `XLS`, `CSV`, `TSV`, `TXT`, `UNKNOWN` |

## FormulaConfig

Controls formula evaluation behavior:

```python
from messy_xlsx.formulas import FormulaConfig, FormulaEvaluationMode

config = FormulaConfig(
    mode=FormulaEvaluationMode.CACHED_ONLY,
    max_iterations=100,
    max_depth=1000,
    raise_on_unsupported=False,
    unsupported_value=None,
)
```

### Formula Evaluation Modes

| Mode | Behavior |
|------|----------|
| `DISABLED` | Ignore all formulas |
| `CACHED_ONLY` | Use cached (last-saved) values only |
| `CACHED_WITH_FALLBACK` | Use cached values, evaluate if missing |
| `ALWAYS_EVALUATE` | Always evaluate formulas |

## Input Validation

All configuration dataclasses validate inputs on construction:

```python
# These raise ValueError:
SheetConfig(skip_rows=-1)
SheetConfig(header_confidence_threshold=1.5)
SheetConfig(merge_strategy="invalid")
FormulaConfig(max_iterations=0)
```
