"""Tests for input validation on configuration dataclasses."""

import pytest

from messy_xlsx.formulas.config import FormulaConfig
from messy_xlsx.models import SheetConfig


class TestSheetConfigValidation:
    """Test bounds validation on SheetConfig fields."""

    def test_rejects_negative_skip_rows(self) -> None:
        with pytest.raises(ValueError, match="skip_rows must be >= 0"):
            SheetConfig(skip_rows=-1)

    def test_rejects_negative_header_rows(self) -> None:
        with pytest.raises(ValueError, match="header_rows must be >= 0"):
            SheetConfig(header_rows=-5)

    def test_rejects_negative_skip_footer(self) -> None:
        with pytest.raises(ValueError, match="skip_footer must be >= 0"):
            SheetConfig(skip_footer=-1)

    def test_rejects_threshold_below_zero(self) -> None:
        with pytest.raises(ValueError, match="header_confidence_threshold"):
            SheetConfig(header_confidence_threshold=-0.1)

    def test_rejects_threshold_above_one(self) -> None:
        with pytest.raises(ValueError, match="header_confidence_threshold"):
            SheetConfig(header_confidence_threshold=1.5)

    def test_accepts_zero_skip_rows(self) -> None:
        config = SheetConfig(skip_rows=0)
        assert config.skip_rows == 0

    def test_accepts_zero_header_rows(self) -> None:
        config = SheetConfig(header_rows=0)
        assert config.header_rows == 0

    def test_accepts_zero_threshold(self) -> None:
        config = SheetConfig(header_confidence_threshold=0.0)
        assert config.header_confidence_threshold == 0.0

    def test_accepts_one_threshold(self) -> None:
        config = SheetConfig(header_confidence_threshold=1.0)
        assert config.header_confidence_threshold == 1.0

    def test_accepts_valid_mid_threshold(self) -> None:
        config = SheetConfig(header_confidence_threshold=0.5)
        assert config.header_confidence_threshold == 0.5

    def test_accepts_positive_values(self) -> None:
        config = SheetConfig(skip_rows=10, header_rows=2, skip_footer=3)
        assert config.skip_rows == 10
        assert config.header_rows == 2
        assert config.skip_footer == 3

    def test_default_values_pass_validation(self) -> None:
        # Should not raise
        config = SheetConfig()
        assert config.skip_rows == 0
        assert config.header_rows == 1
        assert config.skip_footer == 0
        assert config.header_confidence_threshold == 0.7


class TestFormulaConfigValidation:
    """Test bounds validation on FormulaConfig fields."""

    def test_rejects_zero_max_iterations(self) -> None:
        with pytest.raises(ValueError, match="max_iterations must be >= 1"):
            FormulaConfig(max_iterations=0)

    def test_rejects_negative_max_iterations(self) -> None:
        with pytest.raises(ValueError, match="max_iterations must be >= 1"):
            FormulaConfig(max_iterations=-10)

    def test_rejects_zero_max_depth(self) -> None:
        with pytest.raises(ValueError, match="max_depth must be >= 1"):
            FormulaConfig(max_depth=0)

    def test_rejects_negative_max_depth(self) -> None:
        with pytest.raises(ValueError, match="max_depth must be >= 1"):
            FormulaConfig(max_depth=-1)

    def test_accepts_valid_config(self) -> None:
        config = FormulaConfig(max_iterations=50, max_depth=500)
        assert config.max_iterations == 50
        assert config.max_depth == 500

    def test_default_values_pass_validation(self) -> None:
        config = FormulaConfig()
        assert config.max_iterations == 100
        assert config.max_depth == 1000
