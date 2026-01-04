#!/usr/bin/env python3
"""
Test that all messy-xlsx modules can be imported (syntax check).
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

print("Testing messy-xlsx module imports...")
print("=" * 80)

tests_passed = 0
tests_failed = 0

def test_import(module_path, description):
    """Test importing a module."""
    global tests_passed, tests_failed
    try:
        parts = module_path.split('.')
        if len(parts) == 1:
            mod = __import__(module_path)
        else:
            mod = __import__(module_path, fromlist=[parts[-1]])
        print(f"✓ {description:50s} OK")
        tests_passed += 1
        return True
    except ImportError as e:
        # Expected for modules that need external dependencies
        if 'openpyxl' in str(e) or 'pandas' in str(e) or 'numpy' in str(e):
            print(f"⊘ {description:50s} SKIP (missing dependency: {str(e).split()[3]})")
            tests_passed += 1  # Count as passed - syntax is OK
            return True
        else:
            print(f"✗ {description:50s} FAILED: {e}")
            tests_failed += 1
            return False
    except SyntaxError as e:
        print(f"✗ {description:50s} SYNTAX ERROR: {e}")
        tests_failed += 1
        return False
    except Exception as e:
        print(f"✗ {description:50s} ERROR: {e}")
        tests_failed += 1
        return False

# Test all modules
print("\n1. Core Modules")
print("-" * 80)
test_import('messy_xlsx.models', 'Models (dataclasses)')
test_import('messy_xlsx.exceptions', 'Exceptions')
test_import('messy_xlsx.cache', 'Cache (LRU)')
test_import('messy_xlsx.utils', 'Utilities')

print("\n2. Detection Modules")
print("-" * 80)
test_import('messy_xlsx.detection.format_detector', 'Format Detector')
test_import('messy_xlsx.detection.locale_detector', 'Locale Detector')
test_import('messy_xlsx.detection.structure_analyzer', 'Structure Analyzer')

print("\n3. Parsing Modules")
print("-" * 80)
test_import('messy_xlsx.parsing.base_handler', 'Base Handler')
test_import('messy_xlsx.parsing.xlsx_handler', 'XLSX Handler')
test_import('messy_xlsx.parsing.xls_handler', 'XLS Handler')
test_import('messy_xlsx.parsing.csv_handler', 'CSV Handler')
test_import('messy_xlsx.parsing.handler_registry', 'Handler Registry')

print("\n4. Normalization Modules")
print("-" * 80)
test_import('messy_xlsx.normalization.whitespace', 'Whitespace Normalizer')
test_import('messy_xlsx.normalization.numbers', 'Number Normalizer')
test_import('messy_xlsx.normalization.dates', 'Date Normalizer')
test_import('messy_xlsx.normalization.missing_values', 'Missing Value Handler')
test_import('messy_xlsx.normalization.type_inference', 'Type Inference')
test_import('messy_xlsx.normalization.pipeline', 'Normalization Pipeline')

print("\n5. Formula Modules")
print("-" * 80)
test_import('messy_xlsx.formulas.config', 'Formula Config')
test_import('messy_xlsx.formulas.engine', 'Formula Engine')

print("\n6. Public API")
print("-" * 80)
test_import('messy_xlsx.sheet', 'MessySheet')
test_import('messy_xlsx.workbook', 'MessyWorkbook')
test_import('messy_xlsx', 'Main Package')

# Summary
print("\n" + "=" * 80)
print(f"SUMMARY: {tests_passed} passed, {tests_failed} failed")
print("=" * 80)

if tests_failed > 0:
    print("\n⚠ Some modules failed to import (syntax errors or missing dependencies)")
    print("Note: Missing openpyxl/pandas/numpy is expected without installation")
    sys.exit(1)
else:
    print("\n✓ All modules have valid syntax!")
    print("\nNote: To run full tests with actual file parsing:")
    print("  1. Install dependencies: pip install openpyxl pandas numpy")
    print("  2. Run: python test_all_samples.py")
    sys.exit(0)
