"""Unit tests for XLSXHandler."""

import pytest
from messy_xlsx.parsing import XLSXHandler, ParseOptions


class TestXLSXHandler:
    """Test XLSX file parsing."""

    def test_can_handle_xlsx(self):
        """Test handler recognizes XLSX format."""
        handler = XLSXHandler()

        assert handler.can_handle("xlsx") is True
        assert handler.can_handle("xlsm") is True
        assert handler.can_handle("xls") is False
        assert handler.can_handle("csv") is False

    def test_parse_simple_xlsx(self, sample_xlsx):
        """Test parsing simple XLSX file."""
        handler = XLSXHandler()
        options = ParseOptions()

        df = handler.parse(sample_xlsx, "Data", options)

        assert df is not None
        assert len(df) == 3
        assert list(df.columns) == ["Name", "Age", "City"]

    def test_get_sheet_names(self, sample_xlsx):
        """Test getting sheet names."""
        handler = XLSXHandler()

        sheet_names = handler.get_sheet_names(sample_xlsx)

        assert len(sheet_names) == 1
        assert "Data" in sheet_names

    def test_validate_file(self, sample_xlsx):
        """Test file validation."""
        handler = XLSXHandler()

        is_valid, error = handler.validate(sample_xlsx)

        assert is_valid is True
        assert error is None

    def test_handle_merged_cells(self, merged_cells_xlsx):
        """Test handling merged cells."""
        handler = XLSXHandler()
        options = ParseOptions(merge_strategy="fill")

        df = handler.parse(merged_cells_xlsx, "Data", options)

        assert df is not None
        assert len(df) > 0

    def test_skip_rows_config(self, messy_xlsx):
        """Test skip_rows configuration."""
        handler = XLSXHandler()
        options = ParseOptions(skip_rows=3)  # Skip metadata rows

        df = handler.parse(messy_xlsx, "Report", options)

        assert df is not None

    def test_parse_formulas(self, messy_xlsx):
        """Test parsing files with formulas."""
        handler = XLSXHandler()
        options = ParseOptions(data_only=False)

        df = handler.parse(messy_xlsx, "Report", options)

        # Should return data (cached values or formulas)
        assert df is not None
