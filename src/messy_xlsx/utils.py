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

def sanitize_column_name(name: Any) -> str:
    """Sanitize a value for use as a column name."""
    if name is None:
        return "unnamed"

    name_str = str(name).strip()

    if not name_str or name_str.lower() == "nan":
        return "unnamed"

    name_str = re.sub(r"[^\w\s-]", "_", name_str)
    name_str = re.sub(r"[\s-]+", "_", name_str)
    name_str = name_str.strip("_")

    if name_str and name_str[0].isdigit():
        name_str = f"col_{name_str}"

    return name_str or "unnamed"


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
