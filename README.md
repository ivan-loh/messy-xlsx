# messy-xlsx

[![Tests](https://github.com/ivan-loh/messy-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/ivan-loh/messy-xlsx/actions/workflows/test.yml)
[![PyPI version](https://badge.fury.io/py/messy-xlsx.svg)](https://badge.fury.io/py/messy-xlsx)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

Python library for parsing Excel files with structure detection and normalization.

## Installation

```bash
pip install messy-xlsx

# With formula evaluation
pip install messy-xlsx[formulas]

# With XLS support
pip install messy-xlsx[xls]
```

## Quick Start

```python
import messy_xlsx
from messy_xlsx import MessyWorkbook, SheetConfig
import io

# Basic usage
df = messy_xlsx.read_excel("data.xlsx")

# Multi-sheet workbook
wb = MessyWorkbook("data.xlsx")
df = wb.to_dataframe(sheet="Sheet1")
dfs = wb.to_dataframes()  # All sheets

# Read from BytesIO (cloud storage, S3, etc.)
with open("data.xlsx", "rb") as f:
    content = f.read()
buffer = io.BytesIO(content)
wb = MessyWorkbook(buffer, filename="data.xlsx")
df = wb.to_dataframe()

# Structure analysis
structure = wb.get_structure()
print(f"Header row: {structure.header_row}")
print(f"Tables detected: {structure.num_tables}")
print(f"Locale: {structure.detected_locale}")
```

## Features

**Structure Detection**
- Auto-detects header rows with confidence scoring
- Identifies multiple tables per sheet
- Detects merged cells, hidden rows/columns
- Locale detection (US vs European number formats)

**Format Support**
- XLSX/XLSM (Office Open XML)
- XLS (legacy Excel)
- CSV/TSV with delimiter detection

**Normalization Pipeline**
- Locale-aware number parsing (1,234.56 vs 1.234,56)
- Date normalization from multiple formats
- Whitespace cleaning
- Missing value standardization (NA, null, -, etc.)
- Semantic type inference

**Formula Evaluation** (optional)
- Integrates formulas/xlcalculator libraries
- Fallback to cached values
- Configurable evaluation modes

## Configuration

```python
from messy_xlsx import MessyWorkbook, SheetConfig

config = SheetConfig(
    skip_rows        = 2,
    header_rows      = 1,
    auto_detect      = True,
    merge_strategy   = "fill",
    locale           = "auto",
    evaluate_formulas = True,

    # Header detection (auto-enabled by default)
    header_detection_mode = "smart",  # "auto" | "manual" | "smart"
    header_confidence_threshold = 0.7,
    header_patterns  = [r".*name.*", r".*date.*"],  # Boost confidence
)

wb = MessyWorkbook("data.xlsx", sheet_config=config)
df = wb.to_dataframe()
```

### Header Detection Modes

**smart (default)** - Uses detection unless user explicitly overrides
```python
config = SheetConfig(auto_detect=True)  # Headers detected automatically
```

**auto** - Always trust detection if confidence >= threshold
```python
config = SheetConfig(
    header_detection_mode="auto",
    header_confidence_threshold=0.8  # Only use if 80%+ confident
)
```

**manual** - Ignore detection, use explicit values
```python
config = SheetConfig(
    skip_rows=5,
    header_rows=2,
    header_detection_mode="manual"
)
```

## API Reference

### MessyWorkbook

```python
# From file path
wb = MessyWorkbook("data.xlsx", sheet_config=None, formula_config=None)

# From file-like object (BytesIO, S3 stream, etc.)
wb = MessyWorkbook(buffer, sheet_config=None, filename="data.xlsx")

# Properties and methods
wb.sheet_names                    # List of sheet names
wb.file_path                      # Path or None if from buffer
wb.source                         # The file path or buffer
wb.get_sheet(name)                # Get MessySheet object
wb.to_dataframe(sheet=None)       # Convert sheet to DataFrame
wb.to_dataframes()                # Convert all sheets
wb.get_structure(sheet=None)      # Get StructureInfo
wb.get_cell_by_ref("Sheet1!A1")   # Get cell with formula eval
```

### SheetConfig

```python
SheetConfig(
    skip_rows        = 0,
    header_rows      = 1,
    skip_footer      = 0,
    cell_range       = None,           # "A1:F100"
    auto_detect      = True,
    include_hidden   = False,
    merge_strategy   = "fill",         # "fill", "skip", "first_only"
    locale           = "auto",         # "en_US", "de_DE", "auto"
    evaluate_formulas = True,

    # Header detection
    header_detection_mode = "smart",   # "auto", "manual", "smart"
    header_confidence_threshold = 0.7, # 0.0-1.0
    header_fallback  = "first_row",    # "first_row", "none", "error"
    header_patterns  = None,           # List of regex patterns

    # Normalization controls
    normalize        = True,           # Master switch for all normalization
    normalize_dates  = True,           # Convert date columns to datetime
    normalize_numbers = True,          # Parse number strings to numeric
    normalize_whitespace = True,       # Clean whitespace in text columns
)
```

### Disable Normalization

To get raw data without type conversion (useful for ETL pipelines where the destination handles schema):

```python
config = SheetConfig(normalize=False)
wb = MessyWorkbook("data.xlsx", sheet_config=config)
df = wb.to_dataframe()  # All columns as object dtype
```

Or disable specific normalizations:

```python
config = SheetConfig(
    normalize_dates=False,   # Keep dates as strings
    normalize_numbers=False, # Keep numbers as strings
)
```

### StructureInfo

```python
structure = wb.get_structure()

structure.header_row              # Detected header row index
structure.header_confidence       # 0.0-1.0 confidence score
structure.num_tables             # Number of tables detected
structure.detected_locale        # "en_US" or "de_DE"
structure.merged_ranges          # List of merged cell ranges
structure.has_formulas           # Boolean
```

## Architecture

```
messy_xlsx/
├── detection/
│   ├── format_detector.py       # Binary signatures, ZIP analysis
│   ├── structure_analyzer.py    # Headers, tables, merged cells
│   └── locale_detector.py       # Number format detection
├── parsing/
│   ├── xlsx_handler.py          # XLSX/XLSM
│   ├── xls_handler.py           # XLS
│   ├── csv_handler.py           # CSV/TSV with dialect detection
│   └── handler_registry.py      # Format routing
├── normalization/
│   ├── pipeline.py              # Orchestration
│   ├── numbers.py               # Locale-aware number parsing
│   ├── dates.py                 # Date normalization
│   ├── whitespace.py            # Whitespace cleaning
│   ├── missing_values.py        # NA standardization
│   └── type_inference.py        # Semantic type detection
├── formulas/
│   ├── config.py                # Configuration
│   └── engine.py                # External library integration
├── workbook.py                  # MessyWorkbook class
├── sheet.py                     # MessySheet class
└── models.py                    # Data structures
```

## Testing

Tested on 33 real-world Excel files including:
- 100,000 row spreadsheets (5MB)
- Multi-table sheets
- European number formats
- Formula-heavy workbooks
- Merged cells and complex layouts

100% success rate on test suite.

## Dependencies

**Required:**
- Python >= 3.10
- openpyxl >= 3.1
- pandas >= 2.0
- numpy >= 1.24

**Optional:**
- formulas >= 1.2 (formula evaluation)
- xlcalculator >= 0.4 (lightweight formula eval)
- xlrd >= 2.0 (XLS support)

## License

MIT
