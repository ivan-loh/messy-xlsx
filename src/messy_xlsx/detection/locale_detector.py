"""Locale detection for number format parsing."""

# ============================================================================
# Imports
# ============================================================================

import re
from dataclasses import dataclass

import openpyxl


# ============================================================================
# Configuration
# ============================================================================

EUROPEAN_FORMAT_PATTERNS = [
    r"#\.##0,",
    r"0,00",
    r"\[\$€",
]

COMMA_DECIMAL_PATTERN = re.compile(r"\d,\d{2}(?!\d)")
DOT_DECIMAL_PATTERN = re.compile(r"\d\.\d{2}(?!\d)")
DOT_THOUSANDS_PATTERN = re.compile(r"\d\.\d{3}")
COMMA_THOUSANDS_PATTERN = re.compile(r"\d,\d{3}")


# ============================================================================
# Models
# ============================================================================

@dataclass
class LocaleInfo:
    """Detected locale information."""

    locale: str
    decimal_separator: str
    thousands_separator: str
    confidence: float


# ============================================================================
# Locale Detector
# ============================================================================

class LocaleDetector:
    """Detect number format locale from cell values and formats."""

    def detect(self, ws: openpyxl.worksheet.worksheet.Worksheet, data_region: dict | None = None) -> LocaleInfo:
        """Detect locale from worksheet."""
        format_codes = []
        text_values = []

        if data_region:
            start_row = data_region.get("start_row", 1)
            end_row = min(data_region.get("end_row", 50), start_row + 50)
            start_col = data_region.get("start_col", 1)
            end_col = min(data_region.get("end_col", 20), start_col + 20)
        else:
            start_row, end_row = 1, 50
            start_col, end_col = 1, 20

        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                try:
                    cell = ws.cell(row, col)

                    if cell.number_format and cell.number_format != "General":
                        format_codes.append(cell.number_format)

                    if isinstance(cell.value, str):
                        text_values.append(cell.value)

                except Exception:
                    continue

        european_from_formats = self._check_format_codes(format_codes)

        if european_from_formats:
            return LocaleInfo(
                locale="de_DE",
                decimal_separator=",",
                thousands_separator=".",
                confidence=0.9,
            )

        european_from_values, value_confidence = self._check_text_values(text_values)

        if european_from_values:
            return LocaleInfo(
                locale="de_DE",
                decimal_separator=",",
                thousands_separator=".",
                confidence=value_confidence,
            )

        return LocaleInfo(
            locale="en_US",
            decimal_separator=".",
            thousands_separator=",",
            confidence=0.7,
        )

    def _check_format_codes(self, format_codes: list[str]) -> bool:
        """Check if format codes indicate European locale."""
        for code in format_codes:
            for pattern in EUROPEAN_FORMAT_PATTERNS:
                if re.search(pattern, code):
                    return True
        return False

    def _check_text_values(self, text_values: list[str]) -> tuple[bool, float]:
        """Check text values for locale patterns."""
        comma_decimal_count = 0
        dot_decimal_count = 0
        dot_thousands_count = 0
        comma_thousands_count = 0

        for value in text_values:
            if COMMA_DECIMAL_PATTERN.search(value):
                comma_decimal_count += 1
            if DOT_DECIMAL_PATTERN.search(value):
                dot_decimal_count += 1
            if DOT_THOUSANDS_PATTERN.search(value):
                dot_thousands_count += 1
            if COMMA_THOUSANDS_PATTERN.search(value):
                comma_thousands_count += 1

        european_score = comma_decimal_count + dot_thousands_count
        us_score = dot_decimal_count + comma_thousands_count

        total = european_score + us_score
        if total == 0:
            return False, 0.5

        if european_score > us_score:
            confidence = min(0.9, 0.5 + (european_score - us_score) / total)
            return True, confidence

        return False, 0.5

    def detect_from_samples(self, samples: list[str]) -> LocaleInfo:
        """Detect locale from text samples only."""
        is_european, confidence = self._check_text_values(samples)

        if is_european:
            return LocaleInfo(
                locale="de_DE",
                decimal_separator=",",
                thousands_separator=".",
                confidence=confidence,
            )

        return LocaleInfo(
            locale="en_US",
            decimal_separator=".",
            thousands_separator=",",
            confidence=0.7,
        )
