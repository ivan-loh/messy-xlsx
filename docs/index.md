# messy-xlsx

**Parse messy Excel files to clean pandas DataFrames.**

messy-xlsx handles the real-world Excel files that other parsers struggle with:
metadata rows, merged cells, mixed types, European number formats, hidden
rows/columns, and multi-row headers.

## Features

- **Smart structure detection** — automatically finds where data starts, skipping
  report headers and metadata rows.
- **Merged cell handling** — fill, skip, or keep only the first value.
- **Type normalization** — dates, numbers, and whitespace cleaned automatically.
- **Formula evaluation** — optional evaluation via xlcalculator / formulas libraries.
- **Multi-format support** — XLSX, XLSM, XLS, CSV, TSV.
- **BigQuery-ready output** — column names sanitized for warehouse compatibility.
- **Streaming support** — parse from `BytesIO` objects (S3, cloud storage).

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

## Requirements

- Python >= 3.10
- pandas >= 2.0
- openpyxl >= 3.1
- fastexcel >= 0.11
- numpy >= 1.24
