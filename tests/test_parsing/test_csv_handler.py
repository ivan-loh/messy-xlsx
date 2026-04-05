"""Unit tests for CSVHandler."""

from messy_xlsx.parsing import CSVHandler, ParseOptions


class TestCSVHandler:
    """Test CSV file parsing."""

    def test_can_handle_csv(self):
        """Test handler recognizes CSV format."""
        handler = CSVHandler()

        assert handler.can_handle("csv") is True
        assert handler.can_handle("tsv") is True
        assert handler.can_handle("xlsx") is False

    def test_parse_simple_csv(self, temp_dir):
        """Test parsing simple CSV file."""
        csv_file = temp_dir / "test.csv"
        csv_file.write_text("Name,Age,City\nAlice,30,NYC\nBob,25,LA\n")

        handler = CSVHandler()
        options = ParseOptions()

        df = handler.parse(csv_file, "Sheet1", options)

        assert df is not None
        assert len(df) == 2
        assert list(df.columns) == ["Name", "Age", "City"]

    def test_detect_delimiter(self, temp_dir):
        """Test delimiter detection."""
        # Tab-separated
        tsv_file = temp_dir / "test.tsv"
        tsv_file.write_text("Name\tAge\tCity\nAlice\t30\tNYC\n")

        handler = CSVHandler()
        options = ParseOptions()

        df = handler.parse(tsv_file, "Sheet1", options)

        assert df is not None
        assert len(df.columns) == 3

    def test_encoding_detection(self, temp_dir):
        """Test UTF-8 encoding."""
        csv_file = temp_dir / "test.csv"
        csv_file.write_text("Name,Value\nTest,123\n", encoding="utf-8")

        handler = CSVHandler()
        options = ParseOptions()

        df = handler.parse(csv_file, "Sheet1", options)

        assert df is not None

    def test_get_sheet_names_csv(self, temp_dir):
        """Test CSV always returns single sheet."""
        csv_file = temp_dir / "test.csv"
        csv_file.write_text("A,B\n1,2\n")

        handler = CSVHandler()
        sheet_names = handler.get_sheet_names(csv_file)

        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"


class TestCSVMetadataDetectionIntegration:
    """Test CSV metadata/header detection through MessyWorkbook."""

    def test_messy_csv_metadata_skipped(self, temp_dir):
        """Messy CSV with title/metadata rows should auto-detect the real header."""
        from messy_xlsx import MessyWorkbook

        csv_file = temp_dir / "messy_report.csv"
        # Metadata rows must have same delimiter so parser reads all columns
        csv_file.write_text(
            "Report: Sales,\nGenerated: 2024-01-01,\nName,Amount\nAlice,10\nBob,20\n"
        )

        with MessyWorkbook(csv_file) as wb:
            df = wb.to_dataframe()

            # auto_detect_header is now wired through — metadata rows should be skipped
            assert len(df.columns) == 2
            assert len(df) == 2

    def test_auto_detect_header_passed_to_csv(self, temp_dir):
        """Verify auto_detect_header flag reaches CSVHandler from MessyWorkbook."""
        from messy_xlsx import MessyWorkbook, SheetConfig

        csv_file = temp_dir / "simple.csv"
        csv_file.write_text("Name,Amount\nAlice,10\nBob,20\n")

        # With auto_detect enabled, CSVHandler should receive auto_detect_header=True
        with MessyWorkbook(csv_file, sheet_config=SheetConfig(auto_detect=True)) as wb:
            df = wb.to_dataframe()
            assert len(df) == 2
            assert "Name" in df.columns or "name" in df.columns
