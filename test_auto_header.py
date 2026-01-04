#!/usr/bin/env python3
"""Test auto header detection with new configuration options."""

import sys
from pathlib import Path

sys.path.insert(0, "src")

from messy_xlsx import MessyWorkbook, SheetConfig

print("=" * 80)
print("AUTO HEADER DETECTION TEST")
print("=" * 80)

# Test file with messy structure
test_file = Path("tests/samples/budget_vs_actuals.xlsx")

if not test_file.exists():
    print(f"Test file not found: {test_file}")
    sys.exit(1)

print(f"\nTest file: {test_file.name}\n")

# ============================================================================
# Test 1: Default behavior (smart mode)
# ============================================================================
print("-" * 80)
print("Test 1: Default (smart mode)")
print("-" * 80)

config1 = SheetConfig(
    auto_detect=True,
    # header_detection_mode="smart" is default
)

with MessyWorkbook(test_file, sheet_config=config1) as wb:
    structure = wb.get_structure()
    print(f"Detected header row: {structure.header_row}")
    print(f"Header confidence: {structure.header_confidence:.2f}")
    print(f"Suggested skip_rows: {structure.suggested_skip_rows}")

    df = wb.to_dataframe()
    print(f"DataFrame shape: {df.shape}")
    print(f"Columns: {list(df.columns)[:5]}")

# ============================================================================
# Test 2: Auto mode with high confidence threshold
# ============================================================================
print("\n" + "-" * 80)
print("Test 2: Auto mode (high confidence threshold)")
print("-" * 80)

config2 = SheetConfig(
    auto_detect=True,
    header_detection_mode="auto",
    header_confidence_threshold=0.8,
)

with MessyWorkbook(test_file, sheet_config=config2) as wb:
    structure = wb.get_structure()
    print(f"Detected header row: {structure.header_row}")
    print(f"Header confidence: {structure.header_confidence:.2f}")
    print(f"Threshold: 0.8")

    df = wb.to_dataframe()
    print(f"DataFrame shape: {df.shape}")
    print(f"Columns: {list(df.columns)[:5]}")

# ============================================================================
# Test 3: Pattern-based header detection
# ============================================================================
print("\n" + "-" * 80)
print("Test 3: Pattern-based detection")
print("-" * 80)

config3 = SheetConfig(
    auto_detect=True,
    header_detection_mode="auto",
    header_patterns=[r".*budget.*", r".*actual.*", r".*variance.*"],
)

with MessyWorkbook(test_file, sheet_config=config3) as wb:
    structure = wb.get_structure()
    print(f"Detected header row: {structure.header_row}")
    print(f"Header confidence (with patterns): {structure.header_confidence:.2f}")

    df = wb.to_dataframe()
    print(f"DataFrame shape: {df.shape}")
    print(f"Columns: {list(df.columns)[:5]}")

# ============================================================================
# Test 4: Manual mode (override detection)
# ============================================================================
print("\n" + "-" * 80)
print("Test 4: Manual mode (override)")
print("-" * 80)

config4 = SheetConfig(
    skip_rows=2,
    header_rows=1,
    header_detection_mode="manual",
)

with MessyWorkbook(test_file, sheet_config=config4) as wb:
    df = wb.to_dataframe()
    print(f"Manual skip_rows: 2")
    print(f"DataFrame shape: {df.shape}")
    print(f"Columns: {list(df.columns)[:5]}")

print("\n" + "=" * 80)
print("✓ All tests completed successfully!")
print("=" * 80)
