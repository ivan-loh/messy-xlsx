"""Utility functions for messy-xlsx."""

# ============================================================================
# Imports
# ============================================================================

import re
from typing import Any


# ============================================================================
# Cell Reference Functions
# ============================================================================

def cell_ref_to_coords(ref: str) -> tuple[str | None, int, int]:
    """Parse an Excel cell reference to (sheet, row, col)."""
    sheet = None

    if "!" in ref:
        sheet_part, cell_part = ref.rsplit("!", 1)
        sheet = sheet_part.strip("'\"[]")
    else:
        cell_part = ref

    cell_part = cell_part.replace("$", "")

    match = re.match(r"^([A-Za-z]+)(\d+)$", cell_part)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")

    col_letters, row_str = match.groups()
    col = column_letter_to_index(col_letters)
    row = int(row_str)

    return sheet, row, col


def coords_to_cell_ref(row: int, col: int, sheet: str | None = None) -> str:
    """Convert coordinates to Excel cell reference."""
    col_letter = column_index_to_letter(col)
    cell_ref = f"{col_letter}{row}"

    if sheet:
        if re.search(r"[\s!']", sheet):
            sheet = f"'{sheet}'"
        return f"{sheet}!{cell_ref}"

    return cell_ref


def parse_range(range_str: str) -> tuple[str | None, int, int, int, int]:
    """Parse Excel range notation."""
    sheet = None

    if "!" in range_str:
        sheet_part, range_part = range_str.rsplit("!", 1)
        sheet = sheet_part.strip("'\"[]")
    else:
        range_part = range_str

    if ":" not in range_part:
        raise ValueError(f"Invalid range (missing ':'): {range_str}")

    start, end = range_part.split(":")
    _, start_row, start_col = cell_ref_to_coords(start)
    _, end_row, end_col = cell_ref_to_coords(end)

    return sheet, start_row, start_col, end_row, end_col


# ============================================================================
# Column Conversion Functions
# ============================================================================

def column_letter_to_index(letters: str) -> int:
    """Convert column letters to 1-based index."""
    result = 0
    for char in letters.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def column_index_to_letter(index: int) -> str:
    """Convert 1-based column index to letters."""
    result = []
    while index > 0:
        index -= 1
        result.append(chr(ord("A") + index % 26))
        index //= 26
    return "".join(reversed(result))


# ============================================================================
# String Processing Functions
# ============================================================================

# BigQuery reserved words that require quoting in queries
# https://cloud.google.com/bigquery/docs/reference/standard-sql/lexical#reserved_keywords
BIGQUERY_RESERVED_WORDS = frozenset({
    "all", "and", "any", "array", "as", "asc", "assert_rows_modified", "at",
    "between", "by", "case", "cast", "collate", "contains", "create", "cross",
    "cube", "current", "default", "define", "desc", "distinct", "else", "end",
    "enum", "escape", "except", "exclude", "exists", "extract", "false", "fetch",
    "following", "for", "from", "full", "group", "grouping", "groups", "hash",
    "having", "if", "ignore", "in", "inner", "intersect", "interval", "into",
    "is", "join", "lateral", "left", "like", "limit", "lookup", "merge", "natural",
    "new", "no", "not", "null", "nulls", "of", "on", "or", "order", "outer",
    "over", "partition", "preceding", "proto", "qualify", "range", "recursive",
    "respect", "right", "rollup", "rows", "select", "set", "some", "struct",
    "tablesample", "then", "to", "treat", "true", "unbounded", "union", "unnest",
    "using", "when", "where", "window", "with", "within",
})


def _camel_to_snake(name: str) -> str:
    """
    Convert camelCase or PascalCase to snake_case.

    Examples:
        firstName -> first_name
        XMLParser -> xml_parser
        getHTTPResponse -> get_http_response
        already_snake -> already_snake
    """
    # Insert underscore before uppercase letters that follow lowercase letters
    result = re.sub(r"([a-z])([A-Z])", r"\1_\2", name)
    # Insert underscore before uppercase letters that are followed by lowercase
    # (handles cases like XMLParser -> xml_parser)
    result = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1_\2", result)
    return result


def sanitize_column_name(name: Any, max_length: int = 300) -> str:
    """
    Sanitize a value for use as a BigQuery-compatible column name.

    Transforms headers to match: ^[a-zA-Z_][a-zA-Z0-9_]*$
    Converts to lowercase for consistency.

    Features:
        - Converts camelCase/PascalCase to snake_case
        - Replaces spaces and special characters with underscores
        - Prefixes reserved words with 'col_' to avoid quoting in queries
        - Handles names starting with digits

    Args:
        name: The column name to sanitize (any type, will be converted to string)
        max_length: Maximum length for the resulting name (default 300 for BigQuery)

    Returns:
        A sanitized, lowercase column name safe for BigQuery
    """
    if name is None:
        return "unnamed"

    name_str = str(name).strip()

    if not name_str or name_str.lower() == "nan":
        return "unnamed"

    # Convert camelCase/PascalCase to snake_case before lowercasing
    result = _camel_to_snake(name_str)

    # Lowercase for consistency
    result = result.lower()

    # Replace non-ASCII and special chars with underscore
    result = re.sub(r"[^a-z0-9_]", "_", result)

    # Collapse consecutive underscores
    result = re.sub(r"_+", "_", result)

    # Strip leading/trailing underscores
    result = result.strip("_")

    # Ensure starts with letter or underscore (not digit)
    if result and result[0].isdigit():
        result = f"col_{result}"

    # Handle empty result
    if not result:
        return "unnamed"

    # Prefix reserved words to avoid needing backticks in queries
    if result in BIGQUERY_RESERVED_WORDS:
        result = f"col_{result}"

    # Truncate if too long
    if len(result) > max_length:
        result = result[:max_length].rstrip("_")

    return result


def flatten(nested: Any) -> list[Any]:
    """Flatten nested iterables into a single list."""
    result = []

    def _flatten(item: Any) -> None:
        if isinstance(item, (str, bytes)):
            result.append(item)
        elif hasattr(item, "__iter__"):
            for sub in item:
                _flatten(sub)
        else:
            result.append(item)

    _flatten(nested)
    return result
