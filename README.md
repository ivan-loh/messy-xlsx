# messy-xlsx

[![Tests](https://github.com/ivan-loh/messy-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/ivan-loh/messy-xlsx/actions/workflows/test.yml)
[![PyPI version](https://badge.fury.io/py/messy-xlsx.svg)](https://badge.fury.io/py/messy-xlsx)

Parse Excel files (XLSX, XLS, CSV) to pandas DataFrames with structure detection and normalization.

## Install

```bash
pip install messy-xlsx

# Optional: formula evaluation
pip install messy-xlsx[formulas]

# Optional: legacy .xls support
pip install messy-xlsx[xls]
```

## Usage

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
wb = MessyWorkbook(io.BytesIO(content), filename="data.xlsx")
```

## Configuration

```python
config = SheetConfig(
    # Row handling
    skip_rows=0,
    header_rows=1,
    skip_footer=0,
    cell_range=None,              # "A1:F100"

    # Detection
    auto_detect=True,
    header_detection_mode="smart", # "smart", "auto", "manual"
    header_confidence_threshold=0.7,

    # Parsing
    merge_strategy="fill",        # "fill", "skip", "first_only"
    include_hidden=False,
    locale="auto",                # "auto", "en_US", "de_DE"

    # Normalization
    normalize=True,
    normalize_dates=True,
    normalize_numbers=True,
    normalize_whitespace=True,

    # Formulas
    evaluate_formulas=True,
)

wb = MessyWorkbook("data.xlsx", sheet_config=config)
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

Output is compatible with BigQuery/Arrow. Mixed-type columns are coerced to strings.

## Dependencies

- Python >= 3.10
- fastexcel >= 0.11
- openpyxl >= 3.1
- pandas >= 2.0
- numpy >= 1.24

Optional:
- formulas, xlcalculator (formula evaluation)
- xlrd (XLS support)

## License

MIT
