#!/usr/bin/env python3
"""Quick test of optimized performance."""

import sys
import warnings
import time
from pathlib import Path

warnings.filterwarnings('ignore')
sys.path.insert(0, str(Path(__file__).parent / "src"))

from messy_xlsx import MessyWorkbook

# Test the problematic large file
test_files = [
    ("accounts_receivable.xlsx", "small"),
    ("cost_analysis.xlsx", "medium"),
    ("customers.xlsx", "large"),
]

samples_dir = Path("tests/samples")

print("=" * 100)
print("PERFORMANCE TEST - Optimized Structure Analysis")
print("=" * 100)

for filename, size_cat in test_files:
    filepath = samples_dir / filename

    if not filepath.exists():
        print(f"\n✗ {filename} not found")
        continue

    file_size_mb = filepath.stat().st_size / (1024 * 1024)
    print(f"\n📄 {filename} ({size_cat}, {file_size_mb:.2f} MB)")
    print("-" * 100)

    try:
        start = time.time()

        with MessyWorkbook(filepath) as wb:
            # Get structure (this is what was slow)
            structure = wb.get_structure()
            structure_time = time.time() - start

            # Parse to DataFrame
            parse_start = time.time()
            df = wb.to_dataframe()
            parse_time = time.time() - parse_start

            total_time = time.time() - start

            print(f"   ⏱️  Structure analysis: {structure_time:.2f}s")
            print(f"   ⏱️  DataFrame parsing: {parse_time:.2f}s")
            print(f"   ⏱️  Total time: {total_time:.2f}s")
            print(f"   ✓  Parsed: {len(df):,} rows × {len(df.columns)} columns")
            print(f"   📊 Detection:")
            print(f"      - Header row: {structure.header_row}")
            print(f"      - Tables: {structure.num_tables}")
            print(f"      - Data region: rows {structure.data_start_row}-{structure.data_end_row}")

    except Exception as e:
        print(f"   ✗ Error: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 100)
