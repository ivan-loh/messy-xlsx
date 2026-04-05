"""Date normalization for DataFrame columns."""

# ============================================================================
# Imports
# ============================================================================

import re

import pandas as pd

# ============================================================================
# Config
# ============================================================================

EXCEL_EPOCH = "1899-12-30"
EXCEL_DATE_MIN = 1
EXCEL_DATE_MAX = 60000

# Pre-compile all patterns at module level for performance
_DATE_COLUMN_PATTERNS = [
    re.compile(p)
    for p in [
        r"(?i)date",
        r"(?i)time",
        r"(?i)timestamp",
        r"(?i)created",
        r"(?i)updated",
        r"(?i)modified",
        r"(?i)born",
        r"(?i)expired?",
        r"(?i)due",
        r"(?i)start",
        r"(?i)end",
        r"(?i)period",
        r"(?i)day",
        r"(?i)month",
        r"(?i)year",
    ]
]

_NON_DATE_COLUMN_PATTERNS = [
    re.compile(p)
    for p in [
        r"(?i)^year$",
        r"(?i)^month$",
        r"(?i)^day$",
        r"(?i)^fiscal.?year$",
        r"(?i)count",
        r"(?i)total",
        r"(?i)sum",
        r"(?i)qty",
        r"(?i)quantity",
        r"(?i)amount",
        r"(?i)number",
        r"(?i)num",
        r"(?i)id$",
        r"(?i)_id$",
        r"(?i)unique",
        r"(?i)transactions?",
        r"(?i)customers?",
        r"(?i)users?",
        r"(?i)orders?",
        r"(?i)items?",
        r"(?i)units?",
        r"(?i)price",
        r"(?i)cost",
        r"(?i)revenue",
        r"(?i)sales",
        r"(?i)score",
        r"(?i)rating",
        r"(?i)rank",
        r"(?i)index",
        r"(?i)age",
        r"(?i)size",
        r"(?i)length",
        r"(?i)width",
        r"(?i)height",
        r"(?i)weight",
        r"(?i)percent",
        r"(?i)rate",
        r"(?i)ratio",
    ]
]

# Common date formats to try (ordered by likelihood)
_COMMON_DATE_FORMATS = [
    "%Y-%m-%d",  # 2024-01-15
    "%d/%m/%Y",  # 15/01/2024
    "%m/%d/%Y",  # 01/15/2024
    "%Y/%m/%d",  # 2024/01/15
    "%d-%m-%Y",  # 15-01-2024
    "%m-%d-%Y",  # 01-15-2024
    "%d.%m.%Y",  # 15.01.2024
    "%Y-%m-%d %H:%M:%S",  # 2024-01-15 10:30:00
    "%d/%m/%Y %H:%M:%S",  # 15/01/2024 10:30:00
    "%m/%d/%Y %H:%M:%S",  # 01/15/2024 10:30:00
    "%Y-%m-%dT%H:%M:%S",  # 2024-01-15T10:30:00 (ISO)
    "%B %d, %Y",  # January 15, 2024
    "%b %d, %Y",  # Jan 15, 2024
    "%d %B %Y",  # 15 January 2024
    "%d %b %Y",  # 15 Jan 2024
]


# ============================================================================
# Core
# ============================================================================


class DateNormalizer:
    """Normalize dates with multiple format support."""

    def normalize(
        self,
        df: pd.DataFrame,
        semantic_hints: dict[str, str] | None = None,
    ) -> pd.DataFrame:
        """Normalize dates in DataFrame."""
        df = df.copy()
        semantic_hints = semantic_hints or {}

        for idx, col in enumerate(df.columns):
            series = df.iloc[:, idx]
            col_name = str(col)
            # Skip if semantic hint says not a date
            if col in semantic_hints:
                hint = semantic_hints[col].upper()
                if any(
                    t in hint for t in ["DECIMAL", "NUMERIC", "INTEGER", "FLOAT", "VARCHAR", "TEXT"]
                ):
                    continue
                # Explicitly marked as timestamp - always convert
                if "TIMESTAMP" in hint or "DATE" in hint:
                    if self._is_numeric_date_candidate(series):
                        df.isetitem(idx, self._convert_excel_dates(series))
                    elif self._looks_like_text_dates(series, col_name):
                        df.isetitem(idx, self._convert_text_dates(series))
                    continue

            # For numeric columns, only convert if column name suggests it's a date
            if self._is_numeric_date_candidate(series):
                if self._column_name_suggests_date(col_name):
                    df.isetitem(idx, self._convert_excel_dates(series))
            # For text columns, check column name for ambiguous cases
            elif self._looks_like_text_dates(series, col_name):
                df.isetitem(idx, self._convert_text_dates(series))

        return df

    def _column_name_suggests_date(self, col_name: str) -> bool:
        """Check if column name suggests it contains dates."""
        # First check if name suggests NON-date (more specific patterns)
        for pattern in _NON_DATE_COLUMN_PATTERNS:
            if pattern.search(col_name):
                return False

        # Then check if name suggests date
        return any(pattern.search(col_name) for pattern in _DATE_COLUMN_PATTERNS)

    def _is_numeric_date_candidate(self, series: pd.Series) -> bool:
        """Check if column could be Excel serial dates (numeric check only)."""
        if not pd.api.types.is_numeric_dtype(series):
            return False

        sample = series.dropna()
        if len(sample) == 0:
            return False

        in_range = (sample >= EXCEL_DATE_MIN) & (sample <= EXCEL_DATE_MAX)
        is_integer = sample % 1 == 0

        return bool((in_range & is_integer).mean() > 0.8)

    def _looks_like_text_dates(self, series: pd.Series, col_name: str = "") -> bool:
        """Check if column contains text dates."""
        # Accept both object dtype and StringDtype
        is_string_type = series.dtype == object or isinstance(series.dtype, pd.StringDtype)
        if not is_string_type:
            return False

        sample = series.dropna().head(20)
        if len(sample) == 0:
            return False

        # Reject columns where most values are purely numeric (int/float in object dtype)
        # These are NOT dates — they're numbers that happen to be in an object column
        numeric_count = sum(
            1 for v in sample if isinstance(v, (int, float)) and not isinstance(v, bool)
        )
        if numeric_count > len(sample) * 0.5:
            return False

        str_sample = sample.astype(str)

        # Reject columns where most string values look purely numeric
        # (e.g., "1000", "2100", "850000.0") — pd.to_datetime parses these as years
        _NUMERIC_LIKE = re.compile(r"^[+-]?[\d,.\s]+%?$")
        numeric_str_count = sum(1 for v in str_sample if _NUMERIC_LIKE.match(v.strip()))
        if numeric_str_count > len(str_sample) * 0.5:
            return False

        # Try to detect a specific date format from sample first.
        # If values clearly match a known format (e.g., %Y-%m-%d), trust the format
        # regardless of column name.
        detected_format = self._detect_date_format(str_sample)
        if detected_format:
            return True

        # Fallback to mixed format detection — only if column name suggests dates.
        # This prevents bare strings from being aggressively parsed as dates.
        if not self._column_name_suggests_date(col_name):
            return False

        try:
            parsed = pd.to_datetime(str_sample, errors="coerce", format="mixed")
            return bool(parsed.notna().sum() > len(str_sample) * 0.5)
        except Exception:
            return False

    def _column_name_suggests_non_date(self, col_name: str) -> bool:
        """Check if column name suggests it does NOT contain dates."""
        return any(pattern.search(col_name) for pattern in _NON_DATE_COLUMN_PATTERNS)

    def _detect_date_format(self, sample: pd.Series) -> str | None:
        """Try to detect a consistent date format from sample."""
        for fmt in _COMMON_DATE_FORMATS:
            try:
                parsed = pd.to_datetime(sample, format=fmt, errors="coerce")
                # If >80% parse successfully, we found the format
                if parsed.notna().sum() > len(sample) * 0.8:
                    return fmt
            except Exception:
                continue
        return None

    def _convert_excel_dates(self, series: pd.Series) -> pd.Series:
        """Convert Excel serial dates to datetime."""
        try:
            return pd.to_datetime(
                series,
                unit="D",
                origin=EXCEL_EPOCH,
                errors="coerce",
            )
        except Exception:
            return series

    def _convert_text_dates(self, series: pd.Series) -> pd.Series:
        """Convert text dates to datetime."""
        # First try to detect a consistent format
        sample = series.dropna().head(20).astype(str)
        detected_format = self._detect_date_format(sample)

        if detected_format:
            # Use the detected format - much faster than format="mixed"
            try:
                return pd.to_datetime(series, format=detected_format, errors="coerce")
            except Exception:
                pass

        # Fallback to mixed format (slower but handles varied formats)
        try:
            return pd.to_datetime(series, errors="coerce", format="mixed")
        except Exception:
            return series
