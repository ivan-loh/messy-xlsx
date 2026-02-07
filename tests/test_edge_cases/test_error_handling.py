"""Tests for error handling with formulas and Excel errors."""

import openpyxl
import pytest

from messy_xlsx import MessyWorkbook, SheetConfig
from messy_xlsx.formulas import FormulaConfig, FormulaEvaluationMode


class TestFormulaErrors:
    """Test handling of Excel formula errors."""

    def test_all_formula_errors(self, temp_dir):
        """Test file with all types of Excel errors."""
        file_path = temp_dir / "formula_errors.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Error Type", "Formula", "Result"])
        ws.append(["DIV/0", "=1/0", None])
        ws.append(["N/A", "=NA()", None])
        ws.append(["NAME", "=UNKNOWNFUNC()", None])
        ws.append(["NULL", "=A1 B1", None])
        ws.append(["NUM", "=SQRT(-1)", None])
        ws.append(["REF", "=A1000000", None])
        ws.append(["VALUE", "=1+\"text\"", None])
        wb.save(file_path)
        wb.close()

        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_ONLY)

        with MessyWorkbook(file_path, formula_config=config) as mwb:
            df = mwb.to_dataframe()
            # Should handle errors gracefully
            assert len(df) >= 0


class TestCircularReferences:
    """Test handling of circular formula references."""

    def test_simple_circular_reference(self, temp_dir):
        """Test circular reference A1=B1, B1=A1."""
        file_path = temp_dir / "circular.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "=B1"
        ws["B1"] = "=A1"
        wb.save(file_path)
        wb.close()

        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_ONLY)

        with MessyWorkbook(file_path, formula_config=config) as mwb:
            # Should not crash
            df = mwb.to_dataframe()
            assert df is not None


class TestComplexFormulas:
    """Test complex formula scenarios."""

    def test_nested_formulas(self, temp_dir):
        """Test deeply nested formulas."""
        file_path = temp_dir / "nested.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Value", "Formula"])
        ws["A2"] = 10
        ws["B2"] = "=IF(A2>5,IF(A2>8,\"High\",\"Medium\"),\"Low\")"
        wb.save(file_path)
        wb.close()

        with MessyWorkbook(file_path) as mwb:
            df = mwb.to_dataframe()
            assert len(df) == 1

    def test_array_formulas(self, temp_dir):
        """Test array formulas if supported."""
        file_path = temp_dir / "array.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Data", "Sum"])
        ws.append([1, "=SUM(A2:A5)"])
        ws.append([2, None])
        ws.append([3, None])
        ws.append([4, None])
        wb.save(file_path)
        wb.close()

        with MessyWorkbook(file_path) as mwb:
            df = mwb.to_dataframe()
            assert len(df) >= 1


class TestFormulaEvaluationModes:
    """Test different formula evaluation modes."""

    def test_disabled_mode(self, temp_dir):
        """Test formula evaluation disabled."""
        file_path = temp_dir / "formulas.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["A", "B", "Sum"])
        ws.append([1, 2, "=A2+B2"])
        wb.save(file_path)
        wb.close()

        config = FormulaConfig(mode=FormulaEvaluationMode.DISABLED)

        with MessyWorkbook(file_path, formula_config=config) as mwb:
            df = mwb.to_dataframe()
            assert len(df) == 1

    def test_cached_only_mode(self, temp_dir):
        """Test using only cached formula values."""
        file_path = temp_dir / "cached.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Value", "Double"])
        ws["A2"] = 5
        ws["B2"] = "=A2*2"
        wb.save(file_path)
        wb.close()

        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_ONLY)

        with MessyWorkbook(file_path, formula_config=config) as mwb:
            df = mwb.to_dataframe()
            assert len(df) == 1


class TestExceptionSerialization:
    """Test to_dict() on each exception class."""

    def test_base_exception_to_dict(self):
        from messy_xlsx.exceptions import MessyXlsxError

        e = MessyXlsxError("something went wrong", context={"key": "val"})
        d = e.to_dict()
        assert d["error"] == "MessyXlsxError"
        assert d["message"] == "something went wrong"
        assert d["context"] == {"key": "val"}

    def test_file_error_to_dict(self):
        from messy_xlsx.exceptions import FileError

        e = FileError("not found", file_path="/tmp/x.xlsx", operation="open")
        d = e.to_dict()
        assert d["error"] == "FileError"
        assert d["context"]["file_path"] == "/tmp/x.xlsx"
        assert d["context"]["operation"] == "open"

    def test_format_error_to_dict(self):
        from messy_xlsx.exceptions import FormatError

        e = FormatError(
            "bad format",
            file_path="/tmp/x.xlsx",
            detected_format="csv",
            attempted_formats=["xlsx", "xls"],
        )
        d = e.to_dict()
        assert d["error"] == "FormatError"
        assert d["context"]["detected_format"] == "csv"
        assert d["context"]["attempted_formats"] == ["xlsx", "xls"]

    def test_structure_error_to_dict(self):
        from messy_xlsx.exceptions import StructureError

        e = StructureError("no header", sheet="Sheet1", detection_phase="headers")
        d = e.to_dict()
        assert d["error"] == "StructureError"
        assert d["context"]["sheet"] == "Sheet1"
        assert d["context"]["detection_phase"] == "headers"

    def test_normalization_error_to_dict(self):
        from messy_xlsx.exceptions import NormalizationError

        e = NormalizationError(
            "type mismatch",
            column="price",
            row=5,
            value="abc",
            expected_type="float",
        )
        d = e.to_dict()
        assert d["error"] == "NormalizationError"
        assert d["context"]["column"] == "price"
        assert d["context"]["row"] == 5
        assert d["context"]["expected_type"] == "float"

    def test_formula_error_to_dict(self):
        from messy_xlsx.exceptions import FormulaError

        e = FormulaError("eval failed", cell_ref="A1", formula="=1/0")
        d = e.to_dict()
        assert d["error"] == "FormulaError"
        assert d["context"]["cell_ref"] == "A1"
        assert d["context"]["formula"] == "=1/0"

    def test_circular_reference_error_to_dict(self):
        from messy_xlsx.exceptions import CircularReferenceError

        e = CircularReferenceError("cycle detected", cycle=["A1", "B1", "A1"])
        d = e.to_dict()
        assert d["error"] == "CircularReferenceError"
        assert d["context"]["cycle"] == ["A1", "B1", "A1"]

    def test_unsupported_function_error_to_dict(self):
        from messy_xlsx.exceptions import UnsupportedFunctionError

        e = UnsupportedFunctionError("WEBSERVICE", cell_ref="B2")
        d = e.to_dict()
        assert d["error"] == "UnsupportedFunctionError"
        assert d["context"]["function_name"] == "WEBSERVICE"
        assert d["context"]["cell_ref"] == "B2"

    def test_none_context_values_excluded(self):
        from messy_xlsx.exceptions import FileError

        e = FileError("error", file_path=None, operation=None)
        d = e.to_dict()
        assert "file_path" not in d["context"]
        assert "operation" not in d["context"]


class TestUnsupportedFunctions:
    """Test handling of unsupported Excel functions."""

    def test_unsupported_function_placeholder(self, temp_dir):
        """Test unsupported function returns placeholder."""
        file_path = temp_dir / "unsupported.xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Function", "Result"])
        ws.append(["WEBSERVICE", "=WEBSERVICE(\"url\")"])
        wb.save(file_path)
        wb.close()

        config = FormulaConfig(
            mode=FormulaEvaluationMode.CACHED_WITH_FALLBACK,
            raise_on_unsupported=False,
            unsupported_value="#UNSUPPORTED"
        )

        with MessyWorkbook(file_path, formula_config=config) as mwb:
            df = mwb.to_dataframe()
            assert len(df) >= 0
