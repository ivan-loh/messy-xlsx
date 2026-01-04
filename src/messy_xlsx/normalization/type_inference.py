"""Semantic type inference from column names."""

# ============================================================================
# Imports
# ============================================================================

import re

import pandas as pd


# ============================================================================
# Config
# ============================================================================

NUMERIC_PATTERNS = [
    r".*amount.*",
    r".*price.*",
    r".*cost.*",
    r".*revenue.*",
    r".*total.*",
    r".*sum.*",
    r".*value.*",
    r".*balance.*",
    r".*qty.*",
    r".*quantity.*",
    r".*count.*",
    r".*sales.*",
    r".*profit.*",
    r".*expense.*",
    r".*fee.*",
    r".*charge.*",
    r".*rate.*",
    r".*percent.*",
    r".*margin.*",
    r".*weight.*",
    r".*height.*",
    r".*width.*",
    r".*length.*",
    r".*size.*",
    r".*tax.*",
    r".*discount.*",
]

TEXT_ID_PATTERNS = [
    r"^id$",
    r".*_id$",
    r".*id$",
    r".*number$",
    r".*code$",
    r".*ref$",
    r".*key$",
    r"^sku$",
    r".*sku$",
    r".*zip.*",
    r".*postal.*",
    r"account.*code.*",
    r".*batch.*",
    r"tracking.*",
    r"invoice.*number.*",
    r"order.*number.*",
    r"customer.*id.*",
    r"employee.*id.*",
    r"product.*id.*",
    r".*phone.*",
    r".*fax.*",
    r".*ssn.*",
    r".*ein.*",
    r".*iban.*",
    r".*swift.*",
]

DATE_PATTERNS = [
    r".*date.*",
    r".*time.*",
    r".*timestamp.*",
    r".*when.*",
    r".*created.*",
    r".*modified.*",
    r".*updated.*",
    r".*due.*",
    r".*start.*",
    r".*end.*",
    r".*period.*",
    r".*year.*",
    r".*month.*",
    r".*day.*",
    r".*born.*",
    r".*birth.*",
    r".*expir.*",
    r".*valid.*",
]


# ============================================================================
# Core
# ============================================================================

class SemanticTypeInference:
    """Infer column types from semantic patterns in names."""

    def infer_types(self, df: pd.DataFrame) -> dict[str, str]:
        """Infer types for all columns based on names."""
        hints = {}

        for col in df.columns:
            inferred = self._infer_from_name(str(col))
            if inferred:
                hints[col] = inferred

        return hints

    def _infer_from_name(self, col_name: str) -> str | None:
        """Infer type from column name."""
        col_lower = col_name.lower().strip()

        for pattern in DATE_PATTERNS:
            if re.match(pattern, col_lower):
                return "TIMESTAMP"

        for pattern in TEXT_ID_PATTERNS:
            if re.match(pattern, col_lower):
                return "VARCHAR"

        for pattern in NUMERIC_PATTERNS:
            if re.match(pattern, col_lower):
                return "DECIMAL"

        return None

    def detect_type_contamination(
        self,
        df: pd.DataFrame,
        hints: dict[str, str] | None = None,
    ) -> list[dict]:
        """Detect when inferred type doesn't match actual data type."""
        warnings = []
        hints    = hints or self.infer_types(df)

        for col, expected_type in hints.items():
            if col not in df.columns:
                continue

            actual_dtype = str(df[col].dtype)

            if expected_type == "DECIMAL":
                if "datetime" in actual_dtype or "timestamp" in actual_dtype:
                    warnings.append({
                        "column": col,
                        "expected": "numeric",
                        "actual": actual_dtype,
                        "suggestion": f'type_hints: {{"{col}": "DECIMAL"}}',
                    })

            elif expected_type == "VARCHAR":
                if "int" in actual_dtype or "float" in actual_dtype:
                    warnings.append({
                        "column": col,
                        "expected": "text",
                        "actual": actual_dtype,
                        "suggestion": f'type_hints: {{"{col}": "VARCHAR"}}',
                    })

            elif expected_type == "TIMESTAMP":
                if "int" in actual_dtype or "float" in actual_dtype:
                    warnings.append({
                        "column": col,
                        "expected": "datetime",
                        "actual": actual_dtype,
                        "suggestion": f'type_hints: {{"{col}": "TIMESTAMP"}}',
                    })

        return warnings
