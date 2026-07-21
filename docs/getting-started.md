# Getting Started

## Installation

```bash
# Core package
pip install messy-xlsx

# With formula-evaluation fallback for individual cell access
pip install messy-xlsx[formulas]

# With legacy .xls support
pip install messy-xlsx[xls]

# Everything
pip install messy-xlsx[all]
```

## Basic Usage

### Quick Read

The simplest way to read an XLSX, XLSM, XLS, CSV, or TSV file:

```python
from messy_xlsx import read_excel

df = read_excel("data.xlsx")

# SheetConfig fields can also be supplied as keyword arguments.
sales = read_excel("data.xlsx", sheet="Sales", skip_rows=2, normalize=False)
```

### Workbook API

For more control, use the `MessyWorkbook` context manager:

```python
from messy_xlsx import MessyWorkbook

with MessyWorkbook("data.xlsx") as wb:
    # Parse a single sheet
    df = wb.to_dataframe(sheet="Sheet1")

    # Attempt every sheet; failed sheets are skipped by default
    all_dfs = wb.to_dataframes()

    # Or retain structured details for any failed sheets
    all_dfs, errors = wb.to_dataframes(include_errors=True)

    # Inspect the detected structure
    structure = wb.get_structure()
    print(f"Header at row {structure.header_row}")
    print(f"Data rows: {structure.data_start_row}-{structure.data_end_row}")
```

### Reading from Bytes

When loading from cloud storage (S3, GCS, etc.):

```python
import io
from messy_xlsx import MessyWorkbook

content = download_from_s3("bucket", "data.xlsx")
with MessyWorkbook(io.BytesIO(content), filename="data.xlsx") as wb:
    df = wb.to_dataframe()
```

`messy-xlsx` never closes a caller-owned binary stream. It reads seekable streams
from byte zero and restores their original cursor, including after a failure.
Non-seekable streams must be supplied before any bytes are consumed; they are
read once into an internal snapshot and remain open but exhausted. Use
`filename=` to provide a useful name when the stream has no `.name` attribute.

## Configuration

Pass a `SheetConfig` to customize parsing:

```python
from messy_xlsx import MessyWorkbook, SheetConfig

config = SheetConfig(
    skip_rows=2,              # Skip metadata rows
    header_rows=1,            # Number of header rows
    merge_strategy="fill",    # How to handle merged cells
    normalize=True,           # Clean dates, numbers, whitespace
    sanitize_column_names=True,  # BigQuery-compatible names
)

with MessyWorkbook("data.xlsx", sheet_config=config) as wb:
    df = wb.to_dataframe()
```

See [Configuration](configuration.md) for all options.

## Multi-Sheet Processing

```python
from messy_xlsx import read_all_sheets, analyze_excel

# Select likely data sheets, skipping empty and pivot-like sheets by default
results = read_all_sheets("data.xlsx")
for name, df in results.items():
    print(f"{name}: {len(df)} rows")

# Inspect the selection metadata used by read_all_sheets()
info = analyze_excel("data.xlsx")
for sheet in info:
    print(f"{sheet.name}: {sheet.row_count} rows, {sheet.column_count} cols")
```

`read_all_sheets()` and `analyze_excel()` accept filesystem paths to XLSX,
XLSM, or XLS workbooks. Use `MessyWorkbook.to_dataframes()` when the input is a
buffer, when every sheet should be attempted, or when structured per-sheet
errors are needed.

## Error Handling

messy-xlsx uses a structured exception hierarchy:

```python
from messy_xlsx import MessyWorkbook
from messy_xlsx.exceptions import (
    FileError,
    FormatError,
    MessyXlsxError,
    StructureError,
)

try:
    with MessyWorkbook("data.xlsx") as wb:
        df = wb.to_dataframe()
except FileError as e:
    print(f"File problem: {e}")
except FormatError as e:
    print(f"Format problem: {e}")
except StructureError as e:
    print(f"Structure problem: {e}")
except MessyXlsxError as e:
    print(e.to_dict())  # Structured logging payload
```

All library exceptions derive from `MessyXlsxError` and expose `.to_dict()`.
