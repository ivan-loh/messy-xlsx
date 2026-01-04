"""Tests for CSV format variations and edge cases."""

import pytest
from messy_xlsx import MessyWorkbook, SheetConfig


class TestCSVDelimiters:
    """Test various CSV delimiter types."""

    def test_pipe_delimiter(self, temp_dir):
        """Test pipe-delimited CSV."""
        csv_file = temp_dir / "pipe.csv"
        csv_file.write_text("Name|Age|City\nAlice|30|NYC\nBob|25|LA\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) == 2
            assert len(df.columns) == 3

    def test_semicolon_delimiter(self, temp_dir):
        """Test semicolon-delimited CSV (European)."""
        csv_file = temp_dir / "semicolon.csv"
        csv_file.write_text("Name;Age;City\nAlice;30;NYC\nBob;25;LA\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) == 2
            assert len(df.columns) == 3

    def test_tab_delimiter(self, temp_dir):
        """Test tab-delimited (TSV)."""
        csv_file = temp_dir / "tabs.tsv"
        csv_file.write_text("Name\tAge\tCity\nAlice\t30\tNYC\nBob\t25\tLA\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) == 2
            assert len(df.columns) == 3

    def test_space_delimiter(self, temp_dir):
        """Test space-delimited CSV."""
        csv_file = temp_dir / "space.csv"
        csv_file.write_text("Name Age City\nAlice 30 NYC\nBob 25 LA\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) >= 1


class TestCSVQuoting:
    """Test CSV quoting and escaping."""

    def test_quoted_fields_with_commas(self, temp_dir):
        """Test quoted fields containing commas."""
        csv_file = temp_dir / "quoted.csv"
        csv_file.write_text('Name,Description\n"Smith, John","A person"\n')

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) == 1
            assert df.iloc[0]["Name"] == "Smith, John"

    def test_quoted_fields_with_newlines(self, temp_dir):
        """Test quoted fields containing newlines."""
        csv_file = temp_dir / "newlines.csv"
        csv_file.write_text('Name,Address\n"Alice","123 Main St\nApt 4"\n')

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) >= 1


class TestCSVEncodings:
    """Test different CSV encodings."""

    def test_utf8_with_bom(self, temp_dir):
        """Test UTF-8 with BOM marker."""
        csv_file = temp_dir / "utf8_bom.csv"
        csv_file.write_bytes(b"\xef\xbb\xbfName,Age\nAlice,30\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) == 1

    def test_latin1_encoding(self, temp_dir):
        """Test Latin-1 encoding."""
        csv_file = temp_dir / "latin1.csv"
        csv_file.write_text("Name,Value\nCafé,100\n", encoding="latin-1")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) >= 1

    def test_windows1252_encoding(self, temp_dir):
        """Test Windows-1252 encoding."""
        csv_file = temp_dir / "win1252.csv"
        csv_file.write_text("Name,Value\nTest,100\n", encoding="windows-1252")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) >= 1


class TestCSVInconsistent:
    """Test CSV files with inconsistencies."""

    def test_inconsistent_column_counts(self, temp_dir):
        """Test CSV with varying column counts per row."""
        csv_file = temp_dir / "inconsistent.csv"
        csv_file.write_text("A,B,C\n1,2,3\n4,5\n6,7,8,9\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            # Should handle gracefully
            assert len(df) >= 1

    def test_empty_lines_in_csv(self, temp_dir):
        """Test CSV with empty lines."""
        csv_file = temp_dir / "empty_lines.csv"
        csv_file.write_text("A,B\n1,2\n\n3,4\n\n")

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()
            assert len(df) >= 2
