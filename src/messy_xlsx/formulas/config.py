"""Formula evaluation configuration."""

# ============================================================================
# Imports
# ============================================================================

from dataclasses import dataclass
from enum import Enum
from typing import Any


# ============================================================================
# Enums
# ============================================================================

class FormulaEvaluationMode(Enum):
    """How to handle formula evaluation."""

    DISABLED             = "disabled"
    CACHED_ONLY          = "cached_only"
    CACHED_WITH_FALLBACK = "fallback"
    ALWAYS_EVALUATE      = "evaluate"


class CircularRefStrategy(Enum):
    """How to handle circular references."""

    ERROR         = "error"
    RETURN_CACHED = "cached"
    ITERATE       = "iterate"


# ============================================================================
# Models
# ============================================================================

@dataclass
class FormulaConfig:
    """Configuration for formula evaluation."""

    mode: FormulaEvaluationMode               = FormulaEvaluationMode.CACHED_WITH_FALLBACK
    circular_strategy: CircularRefStrategy    = CircularRefStrategy.ERROR
    max_iterations: int                       = 100
    max_depth: int                            = 1000
    unsupported_value: Any                    = "#UNSUPPORTED"
    raise_on_unsupported: bool                = False
