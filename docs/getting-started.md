# Getting Started

## Installation

```bash
# Core package
pip install messy-xlsx

# With formula evaluation support
pip install messy-xlsx[formulas]

# With legacy .xls support
pip install messy-xlsx[xls]

# Everything
pip install messy-xlsx[all]
```

## Basic Usage

### Quick Read

The simplest way to read an Excel file:

```python
from messy_xlsx import read_excel

df = read_excel("data.xlsx")
```

### Workbook API

For more control, use the `MessyWorkbook` context manager:

```python
from messy_xlsx import MessyWorkbook

with MessyWorkbook("data.xlsx") as wb:
    # Parse a single sheet
    df = wb.to_dataframe(sheet="Sheet1")

    # Parse all sheets
    all_dfs = wb.to_dataframes()

    # Inspect structure without parsing
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

`messy-xlsx` does not close caller-owned binary streams. For each library
operation, a seekable stream is borrowed from byte zero and restored to the
cursor position that operation received, including when parsing fails. Supply a
non-seekable stream before any bytes have been consumed; it is read once into an
internal snapshot, remains open, and leaves the original exhausted. `filename=`
supplies or overrides `.name` for diagnostics and extension fallback.

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

# Read all sheets at once
results = read_all_sheets("data.xlsx")
for name, df in results.items():
    print(f"{name}: {len(df)} rows")

# Analyze without loading data
info = analyze_excel("data.xlsx")
for sheet in info:
    print(f"{sheet.name}: {sheet.row_count} rows, {sheet.column_count} cols")
```

## Error Handling

messy-xlsx uses a structured exception hierarchy:

```python
from messy_xlsx import MessyWorkbook
from messy_xlsx.exceptions import FileError, FormatError, StructureError

try:
    with MessyWorkbook("data.xlsx") as wb:
        df = wb.to_dataframe()
except FileError as e:
    print(f"File problem: {e}")
except FormatError as e:
    print(f"Format problem: {e}")
except StructureError as e:
    print(f"Structure problem: {e}")
```

All exceptions have a `.to_dict()` method for structured logging:

```python
except MessyXlsxError as e:
    log_structured(e.to_dict())
```
