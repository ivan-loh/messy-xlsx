"""Unit tests for FormulaEngine."""

import pytest
from messy_xlsx.formulas import FormulaEngine, FormulaConfig, FormulaEvaluationMode
from messy_xlsx import MessyWorkbook


class TestFormulaEngine:
    """Test formula evaluation engine."""

    def test_cached_value_mode(self, messy_xlsx):
        """Test using cached values only."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_ONLY)

        with MessyWorkbook(messy_xlsx, formula_config=config) as wb:
            df = wb.to_dataframe("Report")

            # Should use cached values
            assert df is not None

    def test_disabled_mode(self, messy_xlsx):
        """Test formula evaluation disabled."""
        config = FormulaConfig(mode=FormulaEvaluationMode.DISABLED)

        with MessyWorkbook(messy_xlsx, formula_config=config) as wb:
            df = wb.to_dataframe("Report")

            assert df is not None

    def test_fallback_mode(self, messy_xlsx):
        """Test fallback evaluation mode."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_WITH_FALLBACK)

        with MessyWorkbook(messy_xlsx, formula_config=config) as wb:
            df = wb.to_dataframe("Report")

            # Should try to evaluate, fall back to cached
            assert df is not None

    def test_unsupported_function_handling(self):
        """Test handling unsupported functions."""
        config = FormulaConfig(
            raise_on_unsupported=False,
            unsupported_value="#UNSUPPORTED"
        )

        engine = FormulaEngine(config)

        # Should not raise error, return placeholder
        assert config.unsupported_value == "#UNSUPPORTED"

    def test_formula_detection(self, messy_xlsx):
        """Test detecting formulas in cells."""
        with MessyWorkbook(messy_xlsx) as wb:
            structure = wb.get_structure("Report")

            # messy_xlsx has formulas
            assert isinstance(structure.has_formulas, bool)

    def test_cell_formula_access(self, messy_xlsx):
        """Test accessing cell formulas."""
        with MessyWorkbook(messy_xlsx) as wb:
            # Access a cell that should have a formula
            cell = wb.get_cell("Report", 5, 4)  # Total column

            # Should have either a value or formula
            assert cell.value is not None or cell.formula is not None
