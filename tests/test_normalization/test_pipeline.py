"""Unit tests for NormalizationPipeline."""

import pandas as pd
import pytest

from messy_xlsx.normalization import NormalizationPipeline


class TestNormalizationPipeline:
    """Test normalization pipeline."""

    def test_pipeline_execution(self):
        """Test that all normalization steps execute."""
        pipeline = NormalizationPipeline(decimal_separator=".", thousands_separator=",")

        df = pd.DataFrame({
            "text": ["  hello  ", "world  "],
            "amount": ["$1,234.56", "$2,345.67"]
        })

        result = pipeline.normalize(df)

        # Whitespace should be cleaned
        assert result["text"].iloc[0] == "hello"

        # Numbers should be normalized
        assert pd.api.types.is_numeric_dtype(result["amount"])

    def test_pipeline_order(self):
        """Test that normalization steps run in correct order."""
        pipeline = NormalizationPipeline(decimal_separator=".", thousands_separator=",")

        df = pd.DataFrame({
            "messy": ["  $1,234.56  ", "  NA  ", "  $2,345.67  "]
        })

        result = pipeline.normalize(df)

        # Should handle whitespace, then numbers, then missing values
        # Note: Rows with only missing values after conversion are dropped
        assert len(result) == 2  # NA row was dropped
        assert not pd.isna(result["messy"].iloc[0])
        assert not pd.isna(result["messy"].iloc[1])

    def test_empty_dataframe(self):
        """Test handling empty DataFrame."""
        pipeline = NormalizationPipeline()

        df = pd.DataFrame()

        result = pipeline.normalize(df)

        assert len(result) == 0

    def test_preserve_column_types(self):
        """Test that appropriate column types are preserved."""
        pipeline = NormalizationPipeline()

        df = pd.DataFrame({
            "int_col": [1, 2, 3],
            "float_col": [1.5, 2.5, 3.5],
            "str_col": ["a", "b", "c"]
        })

        result = pipeline.normalize(df)

        assert pd.api.types.is_numeric_dtype(result["int_col"])
        assert pd.api.types.is_numeric_dtype(result["float_col"])
