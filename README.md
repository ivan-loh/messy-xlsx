# messy-xlsx

[![Tests](https://github.com/ivan-loh/messy-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/ivan-loh/messy-xlsx/actions/workflows/test.yml)
[![PyPI version](https://badge.fury.io/py/messy-xlsx.svg)](https://badge.fury.io/py/messy-xlsx)

Parse messy Excel files (XLSX, XLS, CSV) to clean pandas DataFrames with
intelligent structure detection, merged cell handling, and type normalization.

## Install

```bash
pip install messy-xlsx

# Optional: formula evaluation
pip install messy-xlsx[formulas]

# Optional: legacy .xls support
pip install messy-xlsx[xls]

# Everything
pip install messy-xlsx[all]
```

## Quick Start

```python
from messy_xlsx import MessyWorkbook, SheetConfig, read_excel

# Quick read
df = read_excel("data.xlsx")

# With options
df = read_excel("data.xlsx", sheet="Sheet1", skip_rows=2, normalize=False)

# Workbook API
with MessyWorkbook("data.xlsx") as wb:
    df = wb.to_dataframe(sheet="Sheet1")
    all_dfs = wb.to_dataframes()  # All sheets
    structure = wb.get_structure()

# From bytes (S3, cloud storage)
import io
with MessyWorkbook(io.BytesIO(content), filename="data.xlsx") as wb:
    df = wb.to_dataframe()
```

`messy-xlsx` does not close caller-owned binary streams. For each library
operation, a seekable stream is borrowed from byte zero and restored to the
cursor position that operation received, including when parsing fails. Supply a
non-seekable stream before any bytes have been consumed; it is read once into an
internal snapshot, remains open, and leaves the original exhausted. `filename=`
supplies or overrides `.name` for diagnostics and extension fallback.

## Configuration

```python
from messy_xlsx import SheetConfig, MergeStrategy, HeaderDetectionMode

config = SheetConfig(
    # Row handling
    skip_rows=0,
    header_rows=1,
    skip_footer=0,
    cell_range=None,                       # "A1:F100"

    # Detection
    auto_detect=True,
    header_detection_mode="smart",         # or HeaderDetectionMode.SMART
    header_confidence_threshold=0.7,

    # Parsing
    merge_strategy="fill",                 # or MergeStrategy.FILL
    include_hidden=False,

    # Normalization
    normalize=True,
    normalize_dates=True,
    normalize_numbers=True,
    normalize_whitespace=True,
    sanitize_column_names=True,            # BigQuery-compatible names

    # DataFrame formula cells: cached results (True) or expressions (False)
    evaluate_formulas=True,
)

with MessyWorkbook("data.xlsx", sheet_config=config) as wb:
    df = wb.to_dataframe()
```

All string-based config values accept both raw strings and enum types:

```python
from messy_xlsx import MergeStrategy

# These are equivalent:
SheetConfig(merge_strategy="fill")
SheetConfig(merge_strategy=MergeStrategy.FILL)

# Enums compare equal to strings:
assert MergeStrategy.FILL == "fill"  # True
```

Invalid values raise `ValueError` at construction time:

```python
SheetConfig(skip_rows=-1)              # ValueError
SheetConfig(merge_strategy="banana")   # ValueError
```

## Multi-Sheet

```python
from messy_xlsx import read_all_sheets, analyze_excel

# Read all sheets
results = read_all_sheets("data.xlsx")
for name, df in results.items():
    print(f"{name}: {len(df)} rows")

# Analyze without loading
info = analyze_excel("data.xlsx")
for sheet in info:
    print(f"{sheet.name}: {sheet.row_count} rows, {sheet.column_count} cols")
```

## Output

Output is compatible with BigQuery/Arrow. Column names are sanitized by default
and mixed-type columns are coerced to strings.

## Dependencies

- Python >= 3.11
- fastexcel >= 0.19
- openpyxl >= 3.1.5
- pandas >= 3.0
- numpy >= 2.4
- pyarrow >= 23.0

Optional:
- formulas (formula evaluation fallback for cell access)
- xlrd (XLS support)

## Development

```bash
# Install with dev dependencies
make install

# Run tests, lint, type check
make ci

# Run benchmarks
make benchmark

# Serve documentation locally
make docs
```

## License

MIT
