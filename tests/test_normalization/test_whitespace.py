"""Unit tests for WhitespaceNormalizer."""

import pandas as pd
import pytest

from messy_xlsx.normalization import WhitespaceNormalizer


class TestWhitespaceNormalizer:
    """Test whitespace normalization."""

    def test_strip_leading_trailing(self):
        """Test stripping leading and trailing whitespace."""
        normalizer = WhitespaceNormalizer()

        df = pd.DataFrame({
            "text": ["  hello  ", "  world", "foo  "]
        })

        result = normalizer.normalize(df)

        assert result["text"].iloc[0] == "hello"
        assert result["text"].iloc[1] == "world"
        assert result["text"].iloc[2] == "foo"

    def test_collapse_internal_whitespace(self):
        """Test collapsing internal whitespace."""
        normalizer = WhitespaceNormalizer()

        df = pd.DataFrame({
            "text": ["hello    world", "foo  bar"]
        })

        result = normalizer.normalize(df)

        assert result["text"].iloc[0] == "hello world"
        assert result["text"].iloc[1] == "foo bar"

    def test_remove_nbsp(self):
        """Test removing non-breaking spaces."""
        normalizer = WhitespaceNormalizer()

        df = pd.DataFrame({
            "text": ["hello\xa0world"]
        })

        result = normalizer.normalize(df)

        assert result["text"].iloc[0] == "hello world"

    def test_preserve_numeric_values(self):
        """Test that numeric values are not affected."""
        normalizer = WhitespaceNormalizer()

        df = pd.DataFrame({
            "number": [123, 456, 789]
        })

        result = normalizer.normalize(df)

        assert result["number"].iloc[0] == 123

    def test_empty_strings(self):
        """Test handling empty strings."""
        normalizer = WhitespaceNormalizer()

        df = pd.DataFrame({
            "text": ["", "  ", "hello"]
        })

        result = normalizer.normalize(df)

        assert result["text"].iloc[0] == ""
        assert result["text"].iloc[1] == ""
        assert result["text"].iloc[2] == "hello"
