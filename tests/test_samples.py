"""Parametrized tests for all sample Excel files."""

from pathlib import Path

import pytest

from messy_xlsx import MessyWorkbook
from messy_xlsx.exceptions import MessyXlsxError


# Get all sample files
SAMPLES_DIR = Path(__file__).parent / "samples"
SAMPLE_FILES = sorted(SAMPLES_DIR.glob("*.xlsx"))


@pytest.mark.parametrize("sample_file", SAMPLE_FILES, ids=lambda f: f.name)
class TestSampleFiles:
    """Test all sample files can be processed."""

    def test_can_open_workbook(self, sample_file):
        """Test that workbook can be opened."""
        with MessyWorkbook(sample_file) as wb:
            assert wb is not None
            assert len(wb.sheet_names) > 0

    def test_format_detection(self, sample_file):
        """Test that format is correctly detected."""
        with MessyWorkbook(sample_file) as wb:
            assert wb.format_type in ["xlsx", "xlsm", "xls", "csv"]

    def test_structure_detection(self, sample_file):
        """Test that structure can be analyzed."""
        with MessyWorkbook(sample_file) as wb:
            sheet_name = wb.sheet_names[0]
            structure = wb.get_structure(sheet_name)

            assert structure is not None
            assert structure.data_start_row >= 1
            assert structure.data_end_row >= structure.data_start_row
            assert structure.num_tables >= 0
            assert structure.detected_locale in ["en_US", "de_DE", "unknown"]

    def test_can_parse_to_dataframe(self, sample_file):
        """Test that file can be parsed to DataFrame."""
        with MessyWorkbook(sample_file) as wb:
            sheet_name = wb.sheet_names[0]
            df = wb.to_dataframe(sheet_name)

            assert df is not None
            assert len(df.columns) > 0

    def test_multi_sheet_parsing(self, sample_file):
        """Test parsing all sheets in workbook."""
        with MessyWorkbook(sample_file) as wb:
            if len(wb.sheet_names) > 1:
                dfs = wb.to_dataframes()
                assert len(dfs) == len(wb.sheet_names)
                for sheet_name, df in dfs.items():
                    assert sheet_name in wb.sheet_names
                    assert df is not None


@pytest.mark.parametrize("sample_file", SAMPLE_FILES, ids=lambda f: f.name)
def test_file_processing_no_errors(sample_file):
    """Test that each file processes without errors."""
    try:
        with MessyWorkbook(sample_file) as wb:
            # Should not raise exceptions
            _ = wb.sheet_names
            _ = wb.format_type

            # Try first sheet
            if wb.sheet_names:
                sheet_name = wb.sheet_names[0]
                _ = wb.get_structure(sheet_name)
                df = wb.to_dataframe(sheet_name)
                assert df is not None

    except MessyXlsxError as e:
        pytest.fail(f"MessyXlsxError: {e}")
    except Exception as e:
        pytest.fail(f"Unexpected error: {e}")


def test_all_samples_exist():
    """Test that sample files directory exists and has files."""
    assert SAMPLES_DIR.exists(), f"Samples directory not found: {SAMPLES_DIR}"
    assert len(SAMPLE_FILES) > 0, "No sample files found"
    assert len(SAMPLE_FILES) == 32, f"Expected 32 sample files, found {len(SAMPLE_FILES)}"
