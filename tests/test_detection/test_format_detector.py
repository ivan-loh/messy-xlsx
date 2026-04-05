"""Unit tests for FormatDetector."""

from pathlib import Path

import pytest

from messy_xlsx.detection import FormatDetector
from messy_xlsx.exceptions import FormatError


class TestFormatDetector:
    """Test format detection functionality."""

    def test_detect_xlsx_format(self, sample_xlsx):
        """Test XLSX format detection."""
        detector = FormatDetector()
        info = detector.detect(sample_xlsx)

        assert info.format_type == "xlsx"
        assert info.confidence >= 0.9

    def test_detect_from_path_string(self, sample_xlsx):
        """Test detection works with string path."""
        detector = FormatDetector()
        info = detector.detect(str(sample_xlsx))

        assert info.format_type == "xlsx"

    def test_detect_nonexistent_file(self):
        """Test error handling for nonexistent file."""
        detector = FormatDetector()

        with pytest.raises(FormatError):
            detector.detect(Path("/nonexistent/file.xlsx"))

    def test_binary_signature_detection(self, sample_xlsx):
        """Test binary signature matching."""
        FormatDetector()

        # Read first 8 bytes
        with open(sample_xlsx, "rb") as f:
            signature = f.read(8)

        # XLSX files start with PK (ZIP signature)
        assert signature[:2] == b"PK"

    def test_confidence_scoring(self, sample_xlsx):
        """Test confidence scores are within valid range."""
        detector = FormatDetector()
        info = detector.detect(sample_xlsx)

        assert 0.0 <= info.confidence <= 1.0

    def test_xlsb_rejected_at_workbook_level(self, temp_dir):
        """XLSB format should be detected but rejected with a clear error."""
        import zipfile

        # Create a minimal fake .xlsb file (ZIP with xl/workbook.bin)
        xlsb_path = temp_dir / "test.xlsb"
        with zipfile.ZipFile(xlsb_path, "w") as zf:
            zf.writestr("xl/workbook.bin", b"\x00" * 100)
            zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types></Types>')

        # FormatDetector should detect it as xlsb
        detector = FormatDetector()
        info = detector.detect(xlsb_path)
        assert info.format_type == "xlsb"

        # But MessyWorkbook should reject it with a clear error
        from messy_xlsx import MessyWorkbook

        with pytest.raises(FormatError, match="XLSB.*not supported"):
            MessyWorkbook(xlsb_path)
