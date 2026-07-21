# API Reference

## Core workbook

::: messy_xlsx.workbook.MessyWorkbook
    options:
      members:
        - __init__
        - to_dataframe
        - to_dataframes
        - get_sheet
        - get_structure
        - get_cell
        - get_cell_by_ref
        - sheet_names
        - format_type
        - close

## Sheet and table access

`MessyWorkbook.get_sheet()` returns a `MessySheet`. Detected tables are exposed
as `MessyTable` objects through `MessySheet.tables`.

::: messy_xlsx.sheet.MessySheet
    options:
      members:
        - name
        - structure
        - tables
        - has_multiple_tables
        - to_dataframe
        - get_cell
        - iter_rows
        - __getitem__

::: messy_xlsx.sheet.MessyTable
    options:
      members:
        - start_row
        - end_row
        - row_count
        - column_count
        - to_dataframe

## Convenience functions

::: messy_xlsx.read_excel

::: messy_xlsx.read_excel_tables

::: messy_xlsx.analyze_structure

## Multi-sheet API

::: messy_xlsx.read_all_sheets

::: messy_xlsx.analyze_excel

::: messy_xlsx.MultiSheetParser

::: messy_xlsx.MultiSheetOptions

## Configuration

::: messy_xlsx.SheetConfig

::: messy_xlsx.FormulaConfig

## Models

::: messy_xlsx.StructureInfo

::: messy_xlsx.models.TableInfo

::: messy_xlsx.FormatInfo

::: messy_xlsx.CellValue

::: messy_xlsx.SheetError

::: messy_xlsx.SheetInfo

## Enums

::: messy_xlsx.MergeStrategy

::: messy_xlsx.HeaderDetectionMode

::: messy_xlsx.HeaderFallback

::: messy_xlsx.DataType

::: messy_xlsx.FormatType

::: messy_xlsx.FormulaEvaluationMode

::: messy_xlsx.CircularRefStrategy

## Utilities

### `sanitize_column_name(name, max_length=300)`

Sanitizes any value into a lowercase, BigQuery-compatible column name. It is
available directly from `messy_xlsx` and returns `"unnamed"` for an empty or
null-like header.

```python
from messy_xlsx import sanitize_column_name

assert sanitize_column_name("Order Total ($)") == "order_total"
```

## Exceptions

All library exceptions derive from `MessyXlsxError`.

::: messy_xlsx.MessyXlsxError

::: messy_xlsx.FileError

::: messy_xlsx.FormatError

::: messy_xlsx.StructureError

::: messy_xlsx.NormalizationError

::: messy_xlsx.FormulaError

::: messy_xlsx.CircularReferenceError

::: messy_xlsx.UnsupportedFunctionError
