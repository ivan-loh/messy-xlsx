"""Header detection for already-parsed tabular data."""

import re

import pandas as pd

METADATA_PATTERNS = [
    re.compile(r"printed\s*(date|on|by)", re.I),
    re.compile(r"page\s*[-:]?\s*\d+\s*(of|/)\s*\d+", re.I),
    re.compile(r"generated\s*(on|by|at)", re.I),
    re.compile(r"report\s*(date|name|title)", re.I),
    re.compile(r"^(as\s+of|date|run\s+date)\s*:", re.I),
]


def detect_header_row(df: pd.DataFrame, max_scan_rows: int = 10) -> int:
    """Return the zero-based header row for a raw DataFrame."""
    for row_idx in range(min(max_scan_rows, len(df))):
        row_values = df.iloc[row_idx].dropna().astype(str).tolist()
        if not row_values:
            continue
        if any(pattern.search(value) for value in row_values for pattern in METADATA_PATTERNS):
            continue
        if len(row_values) < 2:
            continue
        if any(re.match(r"^\d{4}-\d{2}-\d{2}", value) for value in row_values):
            continue
        if any(re.fullmatch(r"[\d,.]+", value) for value in row_values):
            continue
        return row_idx
    return 0
