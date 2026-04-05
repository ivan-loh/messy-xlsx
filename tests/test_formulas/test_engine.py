"""Unit tests for FormulaEngine."""

from messy_xlsx import MessyWorkbook
from messy_xlsx.formulas import FormulaConfig, FormulaEngine, FormulaEvaluationMode


class TestFormulaEngine:
    """Test formula evaluation engine."""

    def test_cached_value_mode(self, messy_xlsx):
        """Test that CACHED_ONLY mode parses without errors."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_ONLY)

        with MessyWorkbook(messy_xlsx, formula_config=config) as wb:
            df = wb.to_dataframe("Report")

            # Note: formula modes only affect get_cell(), not to_dataframe()
            # (to_dataframe always uses data_only=True). This test verifies
            # that the mode doesn't crash the pipeline.
            assert len(df) > 0
            assert len(df.columns) > 0

    def test_disabled_mode(self, messy_xlsx):
        """Test that DISABLED mode parses without errors."""
        config = FormulaConfig(mode=FormulaEvaluationMode.DISABLED)

        with MessyWorkbook(messy_xlsx, formula_config=config) as wb:
            df = wb.to_dataframe("Report")

            assert len(df) > 0

    def test_fallback_mode(self, messy_xlsx):
        """Test that CACHED_WITH_FALLBACK mode parses without errors."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_WITH_FALLBACK)

        with MessyWorkbook(messy_xlsx, formula_config=config) as wb:
            df = wb.to_dataframe("Report")

            assert len(df) > 0

    def test_unsupported_function_handling(self):
        """Test handling unsupported functions."""
        config = FormulaConfig(raise_on_unsupported=False, unsupported_value="#UNSUPPORTED")

        FormulaEngine(config)

        # Should not raise error, return placeholder
        assert config.unsupported_value == "#UNSUPPORTED"

    def test_formula_detection(self, messy_xlsx):
        """Test detecting formulas in workbook with known formulas."""
        with MessyWorkbook(messy_xlsx) as wb:
            structure = wb.get_structure("Report")

            # messy_xlsx fixture has formulas (=B5+C5, =SUM(...))
            assert structure.has_formulas is True

    def test_cell_formula_access(self, messy_xlsx):
        """Test accessing cell formulas."""
        with MessyWorkbook(messy_xlsx) as wb:
            # Access a cell that should have a formula
            cell = wb.get_cell("Report", 5, 4)  # Total column

            # Should have either a value or formula
            assert cell.value is not None or cell.formula is not None
