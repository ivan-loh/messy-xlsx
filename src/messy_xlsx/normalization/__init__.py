"""Data normalization pipeline."""

# ============================================================================
# Imports
# ============================================================================

from messy_xlsx.normalization.arrow_pipeline import (
    ArrowNormalizationOperation,
    NormalizedStreamingReader,
)
from messy_xlsx.normalization.dates import DateNormalizer
from messy_xlsx.normalization.missing_values import MissingValueHandler
from messy_xlsx.normalization.numbers import NumberNormalizer
from messy_xlsx.normalization.pipeline import NormalizationPipeline
from messy_xlsx.normalization.plan import (
    ColumnNormalization,
    NormalizationPlan,
    NormalizationSample,
    RowCondition,
    compile_normalization_plan,
)
from messy_xlsx.normalization.type_coercion import TypeCoercionNormalizer
from messy_xlsx.normalization.type_inference import SemanticTypeInference
from messy_xlsx.normalization.whitespace import WhitespaceNormalizer

# ============================================================================
# Exports
# ============================================================================

__all__ = [
    "ArrowNormalizationOperation",
    "ColumnNormalization",
    "DateNormalizer",
    "MissingValueHandler",
    "NormalizationPipeline",
    "NormalizationPlan",
    "NormalizationSample",
    "NormalizedStreamingReader",
    "NumberNormalizer",
    "RowCondition",
    "SemanticTypeInference",
    "TypeCoercionNormalizer",
    "WhitespaceNormalizer",
    "compile_normalization_plan",
]
