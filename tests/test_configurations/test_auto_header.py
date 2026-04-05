"""Tests for auto header detection with configuration options.

Converted from the repo-root test_auto_header.py script into proper pytest tests.
"""

from pathlib import Path

import pytest

from messy_xlsx import MessyWorkbook, SheetConfig

SAMPLES_DIR = Path(__file__).parent.parent / "samples"
BUDGET_FILE = SAMPLES_DIR / "budget_vs_actuals.xlsx"


@pytest.fixture
def budget_file():
    """Get the budget_vs_actuals sample file."""
    if not BUDGET_FILE.exists():
        pytest.skip(f"Sample file not found: {BUDGET_FILE}")
    return BUDGET_FILE


class TestAutoHeaderDetection:
    """Test header detection across all modes."""

    def test_smart_mode_default(self, budget_file):
        """Smart mode (default) should detect headers with reasonable confidence."""
        config = SheetConfig(auto_detect=True)

        with MessyWorkbook(budget_file, sheet_config=config) as wb:
            structure = wb.get_structure()

            assert structure.header_row is not None
            assert structure.header_confidence >= 0.5

            df = wb.to_dataframe()
            assert len(df) > 0
            assert len(df.columns) > 0

    def test_auto_mode_high_confidence(self, budget_file):
        """Auto mode with high confidence threshold should still find headers."""
        config = SheetConfig(
            auto_detect=True,
            header_detection_mode="auto",
            header_confidence_threshold=0.8,
        )

        with MessyWorkbook(budget_file, sheet_config=config) as wb:
            structure = wb.get_structure()

            assert structure.header_row is not None

            df = wb.to_dataframe()
            assert len(df) > 0

    def test_pattern_based_detection(self, budget_file):
        """Pattern-based detection should boost confidence for matching headers."""
        config = SheetConfig(
            auto_detect=True,
            header_detection_mode="auto",
            header_patterns=[r".*budget.*", r".*actual.*", r".*variance.*"],
        )

        with MessyWorkbook(budget_file, sheet_config=config) as wb:
            structure = wb.get_structure()

            # Patterns should match, header row should be found
            assert structure.header_row is not None
            assert structure.header_confidence > 0.0

            df = wb.to_dataframe()
            assert len(df) > 0

    def test_manual_mode_override(self, budget_file):
        """Manual mode should use exact skip_rows/header_rows from config."""
        config = SheetConfig(
            skip_rows=2,
            header_rows=1,
            header_detection_mode="manual",
        )

        with MessyWorkbook(budget_file, sheet_config=config) as wb:
            df = wb.to_dataframe()

            assert len(df) > 0
            assert len(df.columns) > 0
