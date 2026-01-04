#!/usr/bin/env python3
"""Profile which detection methods are slow."""

import sys
import warnings
import time
from pathlib import Path

warnings.filterwarnings('ignore')
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Monkey-patch to add timing
import messy_xlsx.detection.structure_analyzer as sa_module

original_analyze = sa_module.StructureAnalyzer.analyze

def timed_analyze(self, file_path, sheet, force=False):
    """Wrapped analyze with timing for each step."""
    import openpyxl

    file_path = Path(file_path)
    if not force:
        cached = self.cache.get(file_path, sheet)
        if cached:
            print("   [CACHE HIT]")
            return cached

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb[sheet]

    timings = {}

    try:
        t = time.time()
        data_region = self._detect_data_region(ws)
        timings['data_region'] = time.time() - t

        t = time.time()
        merged_ranges = self._detect_merged_cells(ws)
        timings['merged_cells'] = time.time() - t

        t = time.time()
        hidden_rows, hidden_columns = self._detect_hidden_content(ws)
        timings['hidden_content'] = time.time() - t

        t = time.time()
        header_info = self._detect_headers(ws, data_region, merged_ranges)
        timings['headers'] = time.time() - t

        t = time.time()
        metadata_rows = self._detect_metadata_rows(ws, data_region, header_info)
        timings['metadata'] = time.time() - t

        t = time.time()
        tables = self._detect_multiple_tables(ws, data_region, header_info)
        timings['tables'] = time.time() - t

        t = time.time()
        locale_info = self.locale_detector.detect(ws, data_region)
        timings['locale'] = time.time() - t

        t = time.time()
        blank_rows = self._detect_blank_rows(ws, data_region)
        timings['blank_rows'] = time.time() - t

        t = time.time()
        has_formulas = self._detect_formulas(ws, data_region)
        timings['formulas'] = time.time() - t

        # Print timing breakdown
        print(f"   Timing breakdown:")
        for step, duration in sorted(timings.items(), key=lambda x: -x[1]):
            print(f"      {step:20s}: {duration:6.2f}s")

    finally:
        wb.close()

    return original_analyze(self, file_path, sheet, force)

sa_module.StructureAnalyzer.analyze = timed_analyze

from messy_xlsx import MessyWorkbook

filepath = Path("tests/samples/cost_analysis.xlsx")
print(f"Testing: {filepath.name}\n")

with MessyWorkbook(filepath) as wb:
    wb.get_structure()
