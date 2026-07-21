# messy-xlsx

**Parse messy spreadsheet files into clean pandas DataFrames.**

messy-xlsx handles the real-world Excel files that other parsers struggle with:
metadata rows, merged cells, mixed types, European number formats, hidden
rows/columns, and multi-row headers.

## Features

- **Smart structure detection** — automatically finds where data starts, skipping
  report headers and metadata rows.
- **Merged cell handling** — fill, skip, or keep only the first value.
- **Type normalization** — dates, numbers, and whitespace cleaned automatically.
- **Formula handling** — choose saved values or expressions in DataFrames, with
  optional evaluation fallback for individual cells.
- **Multi-format support** — XLSX, XLSM, XLS, CSV, and TSV.
- **BigQuery-ready output** — column names sanitized for warehouse compatibility.
- **Stream support** — parse caller-owned `BytesIO` and other binary streams
  without taking ownership.

## Quick Start

```bash
pip install messy-xlsx
```

```python
from messy_xlsx import read_excel

df = read_excel("data.xlsx")
```

See [Getting Started](getting-started.md) for more examples and
[Configuration](configuration.md) for all available options.

For an Excel workbook with several sheets, choose between two workflows:

```python
from messy_xlsx import MessyWorkbook, read_all_sheets

# Select likely data sheets; empty and pivot-like sheets are skipped by default.
selected = read_all_sheets("report.xlsx")

# Attempt every sheet and retain structured information about failures.
with MessyWorkbook("report.xlsx") as workbook:
    all_sheets, errors = workbook.to_dataframes(include_errors=True)
```

`read_all_sheets()` accepts filesystem paths to XLSX, XLSM, or XLS workbooks.
Use `MessyWorkbook` for buffers, CSV/TSV files, and unfiltered all-sheet reads.

## Requirements

- Python >= 3.11
- fastexcel >= 0.19
- numpy >= 2.4
- openpyxl >= 3.1.5
- pandas >= 3.0
- pyarrow >= 23.0

Legacy XLS support requires the optional `xlrd >= 2.0.2` dependency. Individual
cell formula evaluation can use the optional `formulas >= 1.3.4` dependency;
DataFrame parsing reads saved formula results by default and does not recalculate
the workbook. See [Installation and basic usage](getting-started.md#installation).
