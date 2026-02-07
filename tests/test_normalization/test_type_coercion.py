"""Tests for TypeCoercionNormalizer."""

import numpy as np
import pandas as pd
import pytest

from messy_xlsx.normalization.type_coercion import TypeCoercionNormalizer


class TestTypeCoercionNormalizer:
    """Test type coercion for BigQuery/Arrow compatibility."""

    def setup_method(self):
        self.normalizer = TypeCoercionNormalizer()

    def test_all_string_column_unchanged(self):
        """All-string object column should pass through unchanged."""
        df = pd.DataFrame({"a": ["x", "y", "z"]})
        result = self.normalizer.normalize(df)
        assert list(result["a"]) == ["x", "y", "z"]

    def test_all_numeric_column_unchanged(self):
        """All-numeric object column should pass through unchanged."""
        df = pd.DataFrame({"a": pd.array([1, 2, 3], dtype=object)})
        result = self.normalizer.normalize(df)
        assert list(result["a"]) == [1, 2, 3]

    def test_mixed_string_int_coerced_to_string(self):
        """Mixed string+int column should be coerced to all strings."""
        df = pd.DataFrame({"a": pd.array(["foo", 42, "bar"], dtype=object)})
        result = self.normalizer.normalize(df)
        assert list(result["a"]) == ["foo", "42", "bar"]

    def test_mixed_string_float_coerced_to_string(self):
        """Mixed string+float column should be coerced to all strings."""
        df = pd.DataFrame({"a": pd.array(["foo", 3.14, "bar"], dtype=object)})
        result = self.normalizer.normalize(df)
        assert list(result["a"]) == ["foo", "3.14", "bar"]

    def test_nan_preserved_as_null(self):
        """NaN values should be preserved as null after coercion."""
        df = pd.DataFrame({"a": pd.array(["foo", np.nan, 42], dtype=object)})
        result = self.normalizer.normalize(df)
        assert result["a"].iloc[0] == "foo"
        assert pd.isna(result["a"].iloc[1])
        assert result["a"].iloc[2] == "42"

    def test_none_preserved_as_null(self):
        """None values should be preserved as null after coercion."""
        df = pd.DataFrame({"a": pd.array(["foo", None, 42], dtype=object)})
        result = self.normalizer.normalize(df)
        assert result["a"].iloc[0] == "foo"
        assert pd.isna(result["a"].iloc[1])
        assert result["a"].iloc[2] == "42"

    def test_empty_column(self):
        """All-NaN column should pass through unchanged."""
        df = pd.DataFrame({"a": pd.array([None, None, None], dtype=object)})
        result = self.normalizer.normalize(df)
        assert result["a"].isna().all()

    def test_non_object_dtype_skipped(self):
        """Non-object dtype columns should not be touched."""
        df = pd.DataFrame({"a": [1, 2, 3], "b": [1.0, 2.0, 3.0]})
        result = self.normalizer.normalize(df)
        assert result["a"].dtype in (np.int64, int)
        assert result["b"].dtype in (np.float64, float)

    def test_single_value_column(self):
        """Single-value object column should pass through unchanged."""
        df = pd.DataFrame({"a": pd.array(["only"], dtype=object)})
        result = self.normalizer.normalize(df)
        assert list(result["a"]) == ["only"]

    def test_multiple_columns_independent(self):
        """Each column should be coerced independently."""
        df = pd.DataFrame({
            "pure_str": pd.array(["a", "b"], dtype=object),
            "mixed": pd.array(["x", 1], dtype=object),
            "nums": [10, 20],
        })
        result = self.normalizer.normalize(df)
        assert list(result["pure_str"]) == ["a", "b"]
        assert list(result["mixed"]) == ["x", "1"]
        assert list(result["nums"]) == [10, 20]

    def test_numpy_integer_types(self):
        """np.integer subtypes should be grouped with int."""
        df = pd.DataFrame({"a": pd.array([np.int32(1), np.int64(2), np.int32(3)], dtype=object)})
        result = self.normalizer.normalize(df)
        # All ints => single type => no coercion
        assert all(isinstance(v, (int, np.integer)) for v in result["a"])

    def test_numpy_floating_types(self):
        """np.floating subtypes should be grouped with float."""
        df = pd.DataFrame({"a": pd.array([np.float32(1.0), np.float64(2.0)], dtype=object)})
        result = self.normalizer.normalize(df)
        # All floats => single type => no coercion
        assert all(isinstance(v, (float, np.floating)) for v in result["a"])

    def test_float_nan_not_counted_as_type(self):
        """float('nan') should not add 'float' to the type set."""
        df = pd.DataFrame({"a": pd.array(["x", float("nan"), "y"], dtype=object)})
        result = self.normalizer.normalize(df)
        # Only strings + nan => single type => no coercion
        assert result["a"].iloc[0] == "x"
        assert result["a"].iloc[2] == "y"

    def test_does_not_mutate_input(self):
        """Normalize should not modify the original DataFrame."""
        df = pd.DataFrame({"a": pd.array(["foo", 42], dtype=object)})
        original_values = list(df["a"])
        self.normalizer.normalize(df)
        assert list(df["a"]) == original_values
