# API Reference

## Core

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

## Convenience Functions

::: messy_xlsx.read_excel

::: messy_xlsx.read_all_sheets

::: messy_xlsx.analyze_excel

::: messy_xlsx.read_excel_tables

::: messy_xlsx.analyze_structure

## Configuration

::: messy_xlsx.models.SheetConfig

::: messy_xlsx.formulas.config.FormulaConfig

## Models

::: messy_xlsx.models.StructureInfo

::: messy_xlsx.models.FormatInfo

::: messy_xlsx.models.CellValue

::: messy_xlsx.models.SheetError

## Enums

::: messy_xlsx.enums.MergeStrategy

::: messy_xlsx.enums.HeaderDetectionMode

::: messy_xlsx.enums.HeaderFallback

::: messy_xlsx.enums.DataType

::: messy_xlsx.enums.FormatType

## Exceptions

::: messy_xlsx.exceptions.MessyXlsxError

::: messy_xlsx.exceptions.FileError

::: messy_xlsx.exceptions.FormatError

::: messy_xlsx.exceptions.StructureError

::: messy_xlsx.exceptions.NormalizationError

::: messy_xlsx.exceptions.FormulaError

::: messy_xlsx.exceptions.CircularReferenceError

::: messy_xlsx.exceptions.UnsupportedFunctionError
