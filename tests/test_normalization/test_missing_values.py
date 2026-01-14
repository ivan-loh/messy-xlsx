"""Unit tests for MissingValueHandler."""

import numpy as np
import pandas as pd
import pytest

from messy_xlsx.normalization import MissingValueHandler


class TestMissingValueHandler:
    """Test missing value handling."""

    def test_standardize_na_patterns(self):
        """Test standardizing various NA patterns."""
        handler = MissingValueHandler()

        df = pd.DataFrame({
            "col": ["NA", "N/A", "null", "NULL", "value"]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        assert pd.isna(result["col"].iloc[0])
        assert pd.isna(result["col"].iloc[1])
        assert pd.isna(result["col"].iloc[2])
        assert pd.isna(result["col"].iloc[3])
        assert result["col"].iloc[4] == "value"

    def test_extended_missing_values(self):
        """Test extended list with ambiguous values like dash."""
        handler = MissingValueHandler(use_extended_list=True)

        df = pd.DataFrame({
            "col": ["-", ".", "?", "value"]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        assert pd.isna(result["col"].iloc[0])  # dash
        assert pd.isna(result["col"].iloc[1])  # dot
        assert pd.isna(result["col"].iloc[2])  # question mark
        assert result["col"].iloc[3] == "value"

    def test_dash_not_converted_by_default(self):
        """Test that dash is NOT converted by default (conservative behavior)."""
        handler = MissingValueHandler()

        df = pd.DataFrame({
            "col": ["-", "value"]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        # Dash should remain as-is with default settings
        assert result["col"].iloc[0] == "-"
        assert result["col"].iloc[1] == "value"

    def test_empty_string_handling(self):
        """Test handling empty strings."""
        handler = MissingValueHandler()

        df = pd.DataFrame({
            "col": ["", "  ", "value"]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        # Empty strings should be converted to NA
        assert pd.isna(result["col"].iloc[0])
        assert pd.isna(result["col"].iloc[1])

    def test_drop_empty_rows(self):
        """Test dropping rows that are entirely empty."""
        handler = MissingValueHandler()

        df = pd.DataFrame({
            "a": ["NA", "value", "NA"],
            "b": ["NA", "data", "NA"]
        })

        result = handler.normalize(df, drop_empty_rows=True)

        # Should keep only the row with actual values
        assert len(result) == 1

    def test_preserve_zero_values(self):
        """Test that zero values are not treated as missing."""
        handler = MissingValueHandler()

        df = pd.DataFrame({
            "number": [0, 1, 2]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        assert result["number"].iloc[0] == 0

    def test_none_values(self):
        """Test handling None values."""
        handler = MissingValueHandler()

        df = pd.DataFrame({
            "col": [None, "value", None]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        assert pd.isna(result["col"].iloc[0])
        assert pd.isna(result["col"].iloc[2])

    def test_preserve_types_no_mixed_types(self):
        """Test that preserve_types prevents mixed str/float columns."""
        handler = MissingValueHandler(preserve_types=True)

        df = pd.DataFrame({
            "col": ["value", "NA", "another"]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        # All values should be either str or None, not float
        types = set(type(v).__name__ for v in result["col"])
        assert "float" not in types, "Should not have float values in string column"

    def test_legacy_mode_uses_nan(self):
        """Test that preserve_types=False uses np.nan (legacy behavior)."""
        handler = MissingValueHandler(preserve_types=False)

        df = pd.DataFrame({
            "col": ["value", "NA", "another"]
        })

        result = handler.normalize(df, drop_empty_rows=False)

        # With legacy mode, NA becomes np.nan (a float)
        assert isinstance(result["col"].iloc[1], float)
        assert np.isnan(result["col"].iloc[1])
