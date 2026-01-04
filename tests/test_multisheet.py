#!/usr/bin/env python3
"""Demo: Multi-sheet workbook handling."""

import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, "src")

from messy_xlsx import MessyWorkbook

print("=" * 80)
print("MULTI-SHEET WORKBOOK DEMO")
print("=" * 80)

file = Path("tests/samples/financial_statements.xlsx")

with MessyWorkbook(file) as wb:
    print(f"\nFile: {file.name}")
    print(f"Format: {wb.format_type}")
    print(f"\nSheets found: {len(wb.sheet_names)}")
    for i, name in enumerate(wb.sheet_names, 1):
        print(f"  {i}. {name}")

    print("\n" + "-" * 80)
    print("OPTION 1: Parse all sheets")
    print("-" * 80)

    dfs = wb.to_dataframes()

    for sheet_name, df in dfs.items():
        print(f"\n{sheet_name}:")
        print(f"  Shape: {len(df)} rows × {len(df.columns)} columns")
        print(f"  Columns: {list(df.columns)[:5]}")

    print("\n" + "-" * 80)
    print("OPTION 2: Parse specific sheet")
    print("-" * 80)

    sheet = wb.get_sheet("Income Statement")
    df = sheet.to_dataframe()

    print(f"\nSheet: {sheet.name}")
    print(f"Rows: {len(df)}")
    print(f"Columns: {list(df.columns)}")
    print(f"\nFirst 3 rows:")
    print(df.head(3))

    print("\n" + "-" * 80)
    print("OPTION 3: Access by sheet name in to_dataframe()")
    print("-" * 80)

    df = wb.to_dataframe(sheet="Balance Sheet")
    print(f"\nBalance Sheet: {len(df)} rows")

print("\n" + "=" * 80)
