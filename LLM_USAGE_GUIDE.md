# messy-xlsx: LLM Usage Guide

## Overview for LLMs

This library parses messy Excel files into clean pandas DataFrames with intelligent structure detection and data normalization.

## Core Capabilities

### What This Library Does

1. **Format Detection** - Automatically identifies XLSX, XLS, CSV, TSV files
2. **Structure Analysis** - Detects headers, merged cells, multiple tables, locale
3. **Data Normalization** - Cleans whitespace, parses numbers (US/European), handles dates
4. **Formula Evaluation** - Optional integration with external libraries
5. **DataFrame Output** - Returns standard pandas DataFrame

### What Makes It Special

- Handles messy real-world Excel files (merged cells, hidden rows, metadata)
- Auto-detects US (1,234.56) vs European (1.234,56) number formats
- Finds multiple tables on same sheet (separated by blank rows)
- Preserves leading zeros in ID columns
- Evaluates formulas when possible

## Quick Reference

### Import and Basic Usage

```python
import messy_xlsx

df = messy_xlsx.read_excel("file.xlsx")
```

### Main Classes

| Class | Purpose |
|-------|---------|
| `MessyWorkbook` | Main entry point for file parsing |
| `MessySheet` | Represents a single sheet |
| `MessyTable` | Represents a table within a sheet |
| `SheetConfig` | Configuration for parsing behavior |
| `FormulaConfig` | Configuration for formula evaluation |
| `StructureInfo` | Results from structure detection |
| `CellValue` | Individual cell with metadata |

### Key Methods

```python
wb = MessyWorkbook("file.xlsx")

wb.sheet_names                    list[str]
wb.format_type                    str
wb.get_sheet(name)                MessySheet
wb.to_dataframe(sheet, config)    DataFrame
wb.to_dataframes()                dict[str, DataFrame]
wb.get_structure(sheet)           StructureInfo
wb.get_cell(sheet, row, col)      CellValue
wb.get_cell_by_ref(ref)           CellValue
```

## Common LLM Tasks

### Task 1: Parse Excel File

**User Request:** "Parse this Excel file"

**Code:**
```python
import messy_xlsx

wb = messy_xlsx.MessyWorkbook("file.xlsx")
df = wb.to_dataframe()

print(f"Parsed {len(df)} rows × {len(df.columns)} columns")
print(f"Columns: {list(df.columns)}")
print(df.head())
```

### Task 2: Handle Messy File

**User Request:** "This file has metadata rows at the top"

**Code:**
```python
from messy_xlsx import MessyWorkbook, SheetConfig

config = SheetConfig(
    skip_rows   = 3,
    auto_detect = True,
)

wb = MessyWorkbook("messy_file.xlsx", sheet_config=config)
df = wb.to_dataframe()
```

### Task 3: Analyze Structure First

**User Request:** "What's the structure of this file?"

**Code:**
```python
import messy_xlsx

structure = messy_xlsx.analyze_structure("file.xlsx")

print(f"Header row: {structure.header_row} (confidence: {structure.header_confidence})")
print(f"Data region: rows {structure.data_start_row}-{structure.data_end_row}")
print(f"Tables: {structure.num_tables}")
print(f"Number format: {structure.detected_locale}")
print(f"Merged cells: {len(structure.merged_ranges)}")
```

### Task 4: Multi-Table Sheet

**User Request:** "This sheet has multiple tables"

**Code:**
```python
wb = messy_xlsx.MessyWorkbook("file.xlsx")
sheet = wb.get_sheet("Data")

tables = messy_xlsx.read_excel_tables("file.xlsx", sheet="Data")

for i, table_df in enumerate(tables):
    print(f"Table {i}: {len(table_df)} rows")
```

### Task 5: European Number Format

**User Request:** "Numbers use comma as decimal separator"

**Code:**
```python
config = SheetConfig(locale="de_DE")
wb = MessyWorkbook("european_data.xlsx", sheet_config=config)
df = wb.to_dataframe()
```

### Task 6: Access Individual Cells

**User Request:** "Get the value from cell A1"

**Code:**
```python
wb = MessyWorkbook("file.xlsx")
cell = wb.get_cell_by_ref("Sheet1!A1")

print(f"Value: {cell.value}")
if cell.is_formula:
    print(f"Formula: {cell.formula}")
```

## Structure Detection Details

### StructureInfo Fields

```python
structure = wb.get_structure("Sheet1")

structure.data_start_row          int
structure.data_end_row            int
structure.header_row              int | None
structure.header_confidence       float (0.0-1.0)
structure.num_tables              int
structure.table_ranges            list[dict]
structure.detected_locale         str ("en_US" or "de_DE")
structure.decimal_separator       str ("." or ",")
structure.thousands_separator     str ("," or ".")
structure.merged_ranges           list[tuple]
structure.hidden_rows             list[int]
structure.has_formulas            bool
structure.suggested_skip_rows     int
structure.suggested_skip_footer   int
```

## Configuration Options

### SheetConfig

```python
SheetConfig(
    skip_rows         = 0,
    header_rows       = 1,
    skip_footer       = 0,
    cell_range        = None,
    column_renames    = {},
    type_hints        = {},
    auto_detect       = True,
    include_hidden    = False,
    merge_strategy    = "fill",
    locale            = None,
    evaluate_formulas = True,
    drop_regex        = None,
)
```

### FormulaConfig

```python
FormulaConfig(
    mode                 = FormulaEvaluationMode.CACHED_WITH_FALLBACK,
    circular_strategy    = CircularRefStrategy.ERROR,
    max_iterations       = 100,
    max_depth            = 1000,
    unsupported_value    = "#UNSUPPORTED",
    raise_on_unsupported = False,
)
```

## Error Handling

### Exception Hierarchy

```python
MessyXlsxError
├── FileError
├── FormatError
├── StructureError
├── NormalizationError
└── FormulaError
    ├── CircularReferenceError
    └── UnsupportedFunctionError
```

### Handling Errors

```python
from messy_xlsx import MessyWorkbook, FileError, FormatError

try:
    wb = MessyWorkbook("file.xlsx")
    df = wb.to_dataframe()
except FileError as e:
    print(f"File error: {e}")
except FormatError as e:
    print(f"Format error: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Performance Characteristics

### Tested Performance

| File Size | Rows | Time | Status |
|-----------|------|------|--------|
| < 100 KB | 1-1,000 | < 1s | Fast |
| 100 KB - 1 MB | 1K-10K | 1-5s | Good |
| 1 MB - 10 MB | 10K-100K | 5-15s | Acceptable |

### Optimizations

- First 10,000 rows analyzed for structure
- Sampled blank row detection for large files
- LRU cache for repeated analysis
- values_only iteration for speed

### Limits

- Max recommended file size: 50 MB
- Structure analysis limited to 10,000 rows
- No support for encrypted files
- No VBA macro support

## Troubleshooting

### Issue: Slow Performance

**Solution:** Disable auto-detection for very large files

```python
config = SheetConfig(auto_detect=False)
wb = MessyWorkbook("huge_file.xlsx", sheet_config=config)
```

### Issue: Wrong Number Format

**Solution:** Explicitly set locale

```python
config = SheetConfig(locale="de_DE")
```

### Issue: Can't Find Header

**Solution:** Manually specify header row

```python
config = SheetConfig(header_rows=2, skip_rows=1)
```

### Issue: Formula Errors

**Solution:** Configure formula handling

```python
from messy_xlsx import FormulaConfig, FormulaEvaluationMode

formula_config = FormulaConfig(
    mode = FormulaEvaluationMode.CACHED_ONLY,
)
```

## Testing Status

**Validated on 33 real-world files:**
- Success rate: 100%
- Range: 9 to 100,000 rows
- Formats: XLSX with various complexities
- Features tested: merged cells, formulas, European formats, multi-tables

## Code Examples

### Example 1: Financial Data

```python
import messy_xlsx

wb = messy_xlsx.MessyWorkbook("general_ledger.xlsx")
df = wb.to_dataframe()

print(df.columns)
# Output: ['entry_id', 'date', 'account_code', 'account_name',
#          'description', 'debit', 'credit', 'reference', 'posted_by']

print(f"Total entries: {len(df)}")
# Output: Total entries: 1000
```

### Example 2: Manufacturing Data

```python
wb = messy_xlsx.MessyWorkbook("job_orders.xlsx")

structure = wb.get_structure()
print(f"Detected {structure.num_tables} tables")
print(f"Header at row {structure.header_row}")

df = wb.to_dataframe()
print(f"Jobs: {len(df)}")
# Output: Jobs: 660
```

### Example 3: Large Customer Database

```python
wb = messy_xlsx.MessyWorkbook("customers.xlsx")

df = wb.to_dataframe()
print(f"Customers: {len(df):,}")
# Output: Customers: 50,000

print(df.dtypes)
```

## Dependencies

**Required:**
- openpyxl >= 3.1
- pandas >= 2.0
- numpy >= 1.24

**Optional:**
- formulas >= 1.2 (formula evaluation)
- xlcalculator >= 0.4 (lightweight formula eval)
- xlrd >= 2.0 (legacy XLS support)

## Summary

**When to use messy-xlsx:**
- Messy Excel files with complex layouts
- Need automatic header detection
- European number formats
- Multiple tables per sheet
- Formula evaluation required

**When to use pandas.read_excel():**
- Clean, simple Excel files
- Performance critical (no structure detection overhead)
- Don't need advanced features
