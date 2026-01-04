"""Data normalization pipeline."""

# ============================================================================
# Imports
# ============================================================================

import pandas as pd

from messy_xlsx.normalization.dates import DateNormalizer
from messy_xlsx.normalization.missing_values import MissingValueHandler
from messy_xlsx.normalization.numbers import NumberNormalizer
from messy_xlsx.normalization.type_inference import SemanticTypeInference
from messy_xlsx.normalization.whitespace import WhitespaceNormalizer


# ============================================================================
# Core
# ============================================================================

class NormalizationPipeline:
    """Orchestrate data normalization steps."""

    def __init__(
        self,
        decimal_separator: str | None = None,
        thousands_separator: str | None = None,
        extra_missing_values: list[str] | None = None,
        preserve_linebreaks: bool = False,
    ):
        """Initialize pipeline."""
        self.whitespace       = WhitespaceNormalizer()
        self.numbers          = NumberNormalizer(decimal_separator, thousands_separator)
        self.dates            = DateNormalizer()
        self.missing          = MissingValueHandler(extra_missing_values)
        self.type_inference   = SemanticTypeInference()
        self.preserve_linebreaks = preserve_linebreaks

    def normalize(
        self,
        df: pd.DataFrame,
        semantic_hints: dict[str, str] | None = None,
        skip_steps: list[str] | None = None,
    ) -> pd.DataFrame:
        """Apply full normalization pipeline."""
        skip_steps = skip_steps or []

        if semantic_hints is None:
            semantic_hints = self.type_inference.infer_types(df)

        if "whitespace" not in skip_steps:
            df = self.whitespace.normalize(df, self.preserve_linebreaks)

        if "numbers" not in skip_steps:
            df = self.numbers.normalize(df, semantic_hints)

        if "dates" not in skip_steps:
            df = self.dates.normalize(df, semantic_hints)

        if "missing" not in skip_steps:
            df = self.missing.normalize(df)

        return df

    def analyze(self, df: pd.DataFrame) -> dict:
        """Analyze DataFrame without modifying it."""
        type_hints = self.type_inference.infer_types(df)

        warnings = self.type_inference.detect_type_contamination(df, type_hints)

        decimal_sep, thousands_sep = self.numbers._detect_locale(df)

        return {
            "type_hints": type_hints,
            "warnings": warnings,
            "detected_locale": {
                "decimal_separator": decimal_sep,
                "thousands_separator": thousands_sep,
            },
            "columns": len(df.columns),
            "rows": len(df),
        }
