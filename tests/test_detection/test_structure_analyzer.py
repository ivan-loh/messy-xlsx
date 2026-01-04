"""Unit tests for StructureAnalyzer."""

import pytest
from messy_xlsx import MessyWorkbook
from messy_xlsx.detection import StructureAnalyzer
from messy_xlsx.cache import StructureCache


class TestStructureAnalyzer:
    """Test structure analysis functionality."""

    def test_analyze_simple_structure(self, sample_xlsx):
        """Test analyzing simple file structure."""
        with MessyWorkbook(sample_xlsx) as wb:
            structure = wb.get_structure("Data")

            assert structure is not None
            assert structure.header_row is not None
            assert structure.num_tables >= 1
            assert structure.detected_locale in ["en_US", "de_DE", "unknown"]

    def test_header_detection(self, sample_xlsx):
        """Test header row detection."""
        with MessyWorkbook(sample_xlsx) as wb:
            structure = wb.get_structure("Data")

            assert structure.header_row == 1
            assert structure.header_confidence >= 0.5

    def test_detect_messy_structure(self, messy_xlsx):
        """Test detecting messy file structure."""
        with MessyWorkbook(messy_xlsx) as wb:
            structure = wb.get_structure("Report")

            # Should detect metadata rows before data
            assert structure.header_row > 1

    def test_detect_data_region(self, sample_xlsx):
        """Test data region detection."""
        with MessyWorkbook(sample_xlsx) as wb:
            structure = wb.get_structure("Data")

            assert structure.data_start_row >= 1
            assert structure.data_end_row >= structure.data_start_row
            assert structure.data_start_col >= 1
            assert structure.data_end_col >= structure.data_start_col

    def test_detect_formulas(self, messy_xlsx):
        """Test formula detection."""
        with MessyWorkbook(messy_xlsx) as wb:
            structure = wb.get_structure("Report")

            # messy_xlsx has formulas
            assert isinstance(structure.has_formulas, bool)

    def test_merged_cell_detection(self, merged_cells_xlsx):
        """Test merged cell detection."""
        with MessyWorkbook(merged_cells_xlsx) as wb:
            structure = wb.get_structure("Data")

            assert len(structure.merged_ranges) > 0

    def test_multi_table_detection(self, multi_table_xlsx):
        """Test multiple table detection."""
        with MessyWorkbook(multi_table_xlsx) as wb:
            structure = wb.get_structure("Data")

            # Should detect 2 tables separated by blank rows
            assert structure.num_tables >= 1

    def test_locale_detection(self, european_xlsx):
        """Test locale detection."""
        with MessyWorkbook(european_xlsx) as wb:
            structure = wb.get_structure("Data")

            # Should detect decimal/thousands separators
            assert structure.decimal_separator in [".", ","]
            assert structure.thousands_separator in [".", ",", " ", ""]

    def test_structure_caching(self, sample_xlsx):
        """Test that structure analysis results are cached."""
        cache = StructureCache()

        with MessyWorkbook(sample_xlsx) as wb:
            # First call
            structure1 = wb.get_structure("Data")

            # Second call should use cache
            structure2 = wb.get_structure("Data")

            assert structure1.header_row == structure2.header_row
            assert structure1.num_tables == structure2.num_tables
