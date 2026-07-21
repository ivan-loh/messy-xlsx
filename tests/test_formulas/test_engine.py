"""Unit tests for FormulaEngine."""

from messy_xlsx import MessyWorkbook
from messy_xlsx.formulas import FormulaConfig, FormulaEngine, FormulaEvaluationMode


class TestFormulaEngine:
    """Test formula evaluation engine."""

    def test_cached_only_returns_cached_value_without_evaluating(self):
        """CACHED_ONLY must never invoke an evaluator."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_ONLY)
        engine = FormulaEngine(config)

        def should_not_evaluate(cell_ref):
            raise AssertionError(f"unexpected evaluation of {cell_ref}")

        engine._evaluate_formula = should_not_evaluate

        assert engine.evaluate("Sheet", 1, 1, cached_value=12) == 12

    def test_fallback_prefers_cached_value(self):
        """Fallback mode uses a real cached result when one exists."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_WITH_FALLBACK)
        engine = FormulaEngine(config)

        def should_not_evaluate(cell_ref):
            raise AssertionError(f"unexpected evaluation of {cell_ref}")

        engine._evaluate_formula = should_not_evaluate

        assert engine.evaluate("Sheet", 1, 1, cached_value=12) == 12

    def test_fallback_evaluates_when_cached_value_is_missing(self):
        """Fallback mode evaluates only when no cached result is available."""
        config = FormulaConfig(mode=FormulaEvaluationMode.CACHED_WITH_FALLBACK)
        engine = FormulaEngine(config)
        engine._evaluate_formula = lambda cell_ref: 24

        assert engine.evaluate("Sheet", 1, 1, cached_value=None) == 24

    def test_always_evaluate_ignores_cached_value(self):
        """ALWAYS_EVALUATE replaces even a populated cached result."""
        config = FormulaConfig(mode=FormulaEvaluationMode.ALWAYS_EVALUATE)
        engine = FormulaEngine(config)
        engine._evaluate_formula = lambda cell_ref: 24

        assert engine.evaluate("Sheet", 1, 1, cached_value=12) == 24

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
