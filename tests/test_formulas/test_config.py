"""Unit tests for FormulaConfig."""

import pytest
from messy_xlsx.formulas import FormulaConfig, FormulaEvaluationMode, CircularRefStrategy


class TestFormulaConfig:
    """Test formula configuration."""

    def test_default_config(self):
        """Test default configuration values."""
        config = FormulaConfig()

        assert config.mode == FormulaEvaluationMode.CACHED_WITH_FALLBACK
        assert config.circular_strategy == CircularRefStrategy.ERROR
        assert config.raise_on_unsupported is False

    def test_custom_mode(self):
        """Test custom evaluation mode."""
        config = FormulaConfig(mode=FormulaEvaluationMode.ALWAYS_EVALUATE)

        assert config.mode == FormulaEvaluationMode.ALWAYS_EVALUATE

    def test_circular_ref_strategy(self):
        """Test circular reference strategies."""
        config = FormulaConfig(circular_strategy=CircularRefStrategy.RETURN_CACHED)

        assert config.circular_strategy == CircularRefStrategy.RETURN_CACHED

    def test_unsupported_value(self):
        """Test custom unsupported value."""
        config = FormulaConfig(unsupported_value="N/A")

        assert config.unsupported_value == "N/A"

    def test_raise_on_unsupported(self):
        """Test raise on unsupported flag."""
        config = FormulaConfig(raise_on_unsupported=True)

        assert config.raise_on_unsupported is True

    def test_max_iterations(self):
        """Test max iterations configuration."""
        config = FormulaConfig(max_iterations=50)

        assert config.max_iterations == 50

    def test_max_depth(self):
        """Test max depth configuration."""
        config = FormulaConfig(max_depth=500)

        assert config.max_depth == 500
