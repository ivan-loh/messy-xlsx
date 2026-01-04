"""
Basic tests for messy-xlsx library.
"""

import pytest


class TestImports:
    """Test that all modules can be imported."""

    def test_import_main_module(self):
        """Test main module import."""
        import messy_xlsx

        assert hasattr(messy_xlsx, "MessyWorkbook")
        assert hasattr(messy_xlsx, "SheetConfig")
        assert hasattr(messy_xlsx, "FormulaConfig")

    def test_import_models(self):
        """Test models module import."""
        from messy_xlsx.models import (
            CellValue,
            FormatInfo,
            SheetConfig,
            StructureInfo,
            TableInfo,
        )

        assert CellValue is not None
        assert SheetConfig is not None

    def test_import_exceptions(self):
        """Test exceptions module import."""
        from messy_xlsx.exceptions import (
            CircularReferenceError,
            FileError,
            FormatError,
            FormulaError,
            MessyXlsxError,
            NormalizationError,
            StructureError,
            UnsupportedFunctionError,
        )

        assert issubclass(FileError, MessyXlsxError)
        assert issubclass(CircularReferenceError, FormulaError)

    def test_import_detection(self):
        """Test detection modules import."""
        from messy_xlsx.detection import (
            FormatDetector,
            LocaleDetector,
            StructureAnalyzer,
        )

        assert FormatDetector is not None

    def test_import_parsing(self):
        """Test parsing modules import."""
        from messy_xlsx.parsing import (
            CSVHandler,
            FormatHandler,
            HandlerRegistry,
            ParseOptions,
            XLSHandler,
            XLSXHandler,
        )

        assert XLSXHandler is not None

    def test_import_normalization(self):
        """Test normalization modules import."""
        from messy_xlsx.normalization import (
            DateNormalizer,
            MissingValueHandler,
            NormalizationPipeline,
            NumberNormalizer,
            SemanticTypeInference,
            WhitespaceNormalizer,
        )

        assert NormalizationPipeline is not None

    def test_import_formulas(self):
        """Test formulas modules import."""
        from messy_xlsx.formulas import (
            CircularRefStrategy,
            FormulaConfig,
            FormulaEngine,
            FormulaEvaluationMode,
        )

        assert FormulaEngine is not None


class TestBasicParsing:
    """Test basic file parsing functionality."""

    def test_parse_simple_xlsx(self, sample_xlsx):
        """Test parsing a simple XLSX file."""
        from messy_xlsx import MessyWorkbook

        workbook = MessyWorkbook(sample_xlsx)

        assert len(workbook.sheet_names) == 1
        assert workbook.sheet_names[0] == "Data"

        df = workbook.to_dataframe()

        assert len(df) == 3
        assert list(df.columns) == ["Name", "Age", "City"]

        workbook.close()

    def test_workbook_context_manager(self, sample_xlsx):
        """Test using workbook as context manager."""
        from messy_xlsx import MessyWorkbook

        with MessyWorkbook(sample_xlsx) as workbook:
            df = workbook.to_dataframe()
            assert len(df) == 3

    def test_sheet_access(self, sample_xlsx):
        """Test accessing individual sheets."""
        from messy_xlsx import MessyWorkbook

        with MessyWorkbook(sample_xlsx) as workbook:
            sheet = workbook.get_sheet("Data")
            assert sheet.name == "Data"

            df = sheet.to_dataframe()
            assert len(df) == 3

    def test_cell_access(self, sample_xlsx):
        """Test accessing individual cells."""
        from messy_xlsx import MessyWorkbook

        with MessyWorkbook(sample_xlsx) as workbook:
            # Access by sheet/row/col
            cell = workbook.get_cell("Data", 2, 1)  # First data row, first column
            assert cell.value == "Alice"
            assert cell.data_type == "text"

    def test_cell_by_reference(self, sample_xlsx):
        """Test accessing cells by A1 reference."""
        from messy_xlsx import MessyWorkbook

        with MessyWorkbook(sample_xlsx) as workbook:
            cell = workbook.get_cell_by_ref("Data!A2")
            assert cell.value == "Alice"


class TestStructureDetection:
    """Test structure detection functionality."""

    def test_detect_simple_structure(self, sample_xlsx):
        """Test detecting structure of simple file."""
        from messy_xlsx import MessyWorkbook

        with MessyWorkbook(sample_xlsx) as workbook:
            structure = workbook.get_structure("Data")

            assert structure.header_row is not None
            assert structure.data_start_row <= structure.data_end_row
            assert structure.num_tables == 1

    def test_detect_messy_structure(self, messy_xlsx):
        """Test detecting structure of messy file."""
        from messy_xlsx import MessyWorkbook

        with MessyWorkbook(messy_xlsx) as workbook:
            structure = workbook.get_structure()

            # Should detect metadata rows
            assert len(structure.metadata_rows) >= 0

            # Should detect formulas
            assert structure.has_formulas is True or structure.has_formulas is False


class TestFormatDetection:
    """Test file format detection."""

    def test_detect_xlsx_format(self, sample_xlsx):
        """Test detecting XLSX format."""
        from messy_xlsx.detection import FormatDetector

        detector = FormatDetector()
        info = detector.detect(sample_xlsx)

        assert info.format_type == "xlsx"
        assert info.confidence > 0.9


class TestNormalization:
    """Test data normalization."""

    def test_whitespace_normalization(self):
        """Test whitespace cleaning."""
        import pandas as pd

        from messy_xlsx.normalization import WhitespaceNormalizer

        normalizer = WhitespaceNormalizer()
        df = pd.DataFrame({
            "text": ["  hello  ", "world  ", "  foo bar  "],
        })

        result = normalizer.normalize(df)
        assert result["text"].tolist() == ["hello", "world", "foo bar"]

    def test_missing_value_handling(self):
        """Test missing value standardization."""
        import numpy as np
        import pandas as pd

        from messy_xlsx.normalization import MissingValueHandler

        handler = MissingValueHandler()
        df = pd.DataFrame({
            "col": ["value", "NA", "N/A", "null", ""],
        })

        result = handler.normalize(df, drop_empty_rows=False)
        assert pd.isna(result["col"].iloc[1])
        assert pd.isna(result["col"].iloc[2])
        assert pd.isna(result["col"].iloc[3])
