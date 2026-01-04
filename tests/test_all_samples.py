#!/usr/bin/env python3
"""
Comprehensive test of messy-xlsx on all sample files.

Tests:
- Format detection
- Structure analysis
- DataFrame parsing
- Error handling
"""

import sys
from pathlib import Path
import traceback
from datetime import datetime
import warnings

# Suppress pandas warnings for cleaner output
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=UserWarning)

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from messy_xlsx import MessyWorkbook, read_excel
from messy_xlsx.exceptions import MessyXlsxError


def test_file(file_path: Path) -> dict:
    """Test a single file and return results."""
    result = {
        "file": file_path.name,
        "size_kb": file_path.stat().st_size / 1024,
        "status": "unknown",
        "format": None,
        "sheets": [],
        "rows": 0,
        "columns": 0,
        "structure": {},
        "error": None,
    }

    try:
        # Open workbook
        with MessyWorkbook(file_path) as wb:
            result["format"] = wb.format_type
            result["sheets"] = wb.sheet_names.copy()

            # Try first sheet
            sheet_name = wb.sheet_names[0]

            # Get structure
            try:
                structure = wb.get_structure(sheet_name)
                result["structure"] = {
                    "header_row": structure.header_row,
                    "header_confidence": structure.header_confidence,
                    "num_tables": structure.num_tables,
                    "has_formulas": structure.has_formulas,
                    "merged_cells": len(structure.merged_ranges),
                    "hidden_rows": len(structure.hidden_rows),
                    "detected_locale": structure.detected_locale,
                }
            except Exception as e:
                result["structure"]["error"] = str(e)

            # Try to parse to DataFrame
            try:
                df = wb.to_dataframe(sheet_name)
                result["rows"] = len(df)
                result["columns"] = len(df.columns)
                result["status"] = "success"
            except Exception as e:
                result["status"] = "parse_failed"
                result["error"] = str(e)

    except MessyXlsxError as e:
        result["status"] = "messy_error"
        result["error"] = str(e)
    except Exception as e:
        result["status"] = "unexpected_error"
        result["error"] = str(e)
        result["traceback"] = traceback.format_exc()

    return result


def main():
    """Test all sample files."""
    samples_dir = Path(__file__).parent / "tests" / "samples"

    if not samples_dir.exists():
        print(f"Error: {samples_dir} does not exist")
        return 1

    files = sorted(samples_dir.glob("*.xlsx"))

    if not files:
        print(f"No .xlsx files found in {samples_dir}")
        return 1

    print(f"Testing {len(files)} Excel files...\n")
    print("=" * 100)

    results = []
    success_count = 0
    failed_count = 0

    for i, file_path in enumerate(files, 1):
        print(f"\n[{i}/{len(files)}] Testing: {file_path.name}", flush=True)
        print("-" * 100, flush=True)

        result = test_file(file_path)
        results.append(result)

        # Print summary
        print(f"  Status: {result['status']}")
        print(f"  Format: {result['format']}")
        print(f"  Size: {result['size_kb']:.1f} KB")
        print(f"  Sheets: {len(result['sheets'])} - {result['sheets'][:3]}")

        if result['status'] == 'success':
            print(f"  ✓ Parsed: {result['rows']:,} rows × {result['columns']} columns")
            success_count += 1

            # Show structure info
            if result['structure']:
                s = result['structure']
                print(f"  Structure:")
                if 'header_row' in s:
                    print(f"    - Header row: {s['header_row']} (confidence: {s.get('header_confidence', 0):.2f})")
                if 'num_tables' in s:
                    print(f"    - Tables: {s['num_tables']}")
                if 'has_formulas' in s:
                    print(f"    - Has formulas: {s['has_formulas']}")
                if 'merged_cells' in s:
                    print(f"    - Merged cells: {s['merged_cells']}")
                if 'detected_locale' in s:
                    print(f"    - Locale: {s['detected_locale']}")
        else:
            print(f"  ✗ Failed: {result['error']}")
            failed_count += 1

            if 'traceback' in result and '--verbose' in sys.argv:
                print(f"\n  Traceback:")
                print("  " + "\n  ".join(result['traceback'].split('\n')))

    # Final summary
    print("\n" + "=" * 100)
    print(f"\nFINAL SUMMARY")
    print("=" * 100)
    print(f"Total files tested: {len(files)}")
    print(f"✓ Successful: {success_count} ({success_count/len(files)*100:.1f}%)")
    print(f"✗ Failed: {failed_count} ({failed_count/len(files)*100:.1f}%)")

    # Breakdown by status
    print("\nStatus breakdown:")
    status_counts = {}
    for result in results:
        status = result['status']
        status_counts[status] = status_counts.get(status, 0) + 1

    for status, count in sorted(status_counts.items()):
        print(f"  {status}: {count}")

    # Show failures
    if failed_count > 0:
        print("\nFailed files:")
        for result in results:
            if result['status'] != 'success':
                print(f"  - {result['file']}: {result['status']} - {result['error'][:80]}")

    # Show largest successfully parsed files
    print("\nLargest successfully parsed files:")
    successful = [r for r in results if r['status'] == 'success']
    for result in sorted(successful, key=lambda r: r['size_kb'], reverse=True)[:5]:
        print(f"  - {result['file']}: {result['size_kb']:.1f} KB, {result['rows']:,} rows")

    print("\n" + "=" * 100)

    return 0 if failed_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
