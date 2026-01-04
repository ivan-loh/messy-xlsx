#!/usr/bin/env python3
"""Quick demo of messy-xlsx parsing sample files."""

import sys
import warnings
from pathlib import Path

warnings.filterwarnings('ignore')
sys.path.insert(0, str(Path(__file__).parent / "src"))

from messy_xlsx import MessyWorkbook

# Test a few representative files
test_files = [
    "accounts_receivable.xlsx",
    "budget_vs_actuals.xlsx",
    "job_orders.xlsx",
    "general_ledger.xlsx",
]

samples_dir = Path("tests/samples")

print("=" * 100)
print("MESSY-XLSX DEMO - Parsing Sample Files")
print("=" * 100)

for filename in test_files:
    filepath = samples_dir / filename

    if not filepath.exists():
        print(f"\n✗ {filename} not found")
        continue

    print(f"\n📄 {filename}")
    print("-" * 100)

    try:
        with MessyWorkbook(filepath) as wb:
            # Show workbook info
            print(f"   Format: {wb.format_type}")
            print(f"   Sheets: {', '.join(wb.sheet_names)}")

            # Get structure
            structure = wb.get_structure()
            print(f"   Header: Row {structure.header_row} (confidence: {structure.header_confidence:.2f})")
            print(f"   Locale: {structure.detected_locale} ({structure.decimal_separator} decimal)")
            print(f"   Tables: {structure.num_tables}")
            print(f"   Merged cells: {len(structure.merged_ranges)}")
            print(f"   Has formulas: {structure.has_formulas}")

            # Parse to DataFrame
            df = wb.to_dataframe()
            print(f"\n   ✓ Parsed: {len(df):,} rows × {len(df.columns)} columns")
            print(f"   Columns: {', '.join(df.columns[:5])}" + ("..." if len(df.columns) > 5 else ""))

            # Show sample data
            print(f"\n   Sample data (first 3 rows):")
            print("   " + str(df.head(3).to_string()).replace("\n", "\n   "))

    except Exception as e:
        print(f"   ✗ Error: {e}")

print("\n" + "=" * 100)
print("Demo complete!")
print("=" * 100)
