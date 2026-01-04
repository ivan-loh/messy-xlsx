"""Base class for format handlers."""

# ============================================================================
# Imports
# ============================================================================

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO

import pandas as pd


# ============================================================================
# Type Aliases
# ============================================================================

FileSource = Path | BinaryIO


# ============================================================================
# Models
# ============================================================================

@dataclass
class ParseOptions:
    """Options for parsing a file."""

    encoding: str = "utf-8"

    skip_rows: int = 0
    header_rows: int = 1
    skip_footer: int = 0

    na_values: list[str] = field(default_factory=list)

    parse_dates: bool = True

    chunk_size: int | None = None

    data_only: bool = True

    preserve_formatting: bool = False

    merge_strategy: str = "fill"

    ignore_hidden: bool = False

    max_rows: int | None = None

    cell_range: str | None = None


# ============================================================================
# Core
# ============================================================================

class FormatHandler(ABC):
    """Abstract base class for format handlers."""

    @abstractmethod
    def can_handle(self, format_type: str) -> bool:
        """Check if this handler can process the given format."""
        ...

    @abstractmethod
    def parse(
        self,
        file_source: FileSource,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Parse a file and return a DataFrame."""
        ...

    @abstractmethod
    def get_sheet_names(self, file_source: FileSource) -> list[str]:
        """Get list of sheet names in file."""
        ...

    @abstractmethod
    def validate(self, file_source: FileSource) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        ...

    def _apply_row_limits(
        self,
        df: pd.DataFrame,
        options: ParseOptions,
    ) -> pd.DataFrame:
        """Apply skip_rows, skip_footer, and max_rows to DataFrame."""
        if options.skip_rows > 0:
            df = df.iloc[options.skip_rows:]

        if options.skip_footer > 0:
            df = df.iloc[:-options.skip_footer]

        if options.max_rows is not None:
            df = df.iloc[:options.max_rows]

        return df.reset_index(drop=True)

    def _generate_column_names(
        self,
        df: pd.DataFrame,
        header_rows: int,
    ) -> tuple[pd.DataFrame, list[str]]:
        """Generate column names from header rows."""
        if header_rows == 0:
            columns = [f"col_{i}" for i in range(len(df.columns))]
            return df, columns

        if header_rows == 1:
            headers = df.iloc[0].tolist()
            columns = []
            for i, h in enumerate(headers):
                if h is None or (isinstance(h, float) and pd.isna(h)):
                    columns.append(f"col_{i}")
                else:
                    columns.append(str(h))
            data_df = df.iloc[1:]
            return data_df, columns

        header_rows_data = df.iloc[:header_rows]
        data_df          = df.iloc[header_rows:]

        columns = []
        for col_idx in range(len(df.columns)):
            parts = []
            for row_idx in range(header_rows):
                val = header_rows_data.iloc[row_idx, col_idx]
                if val is not None and not (isinstance(val, float) and pd.isna(val)):
                    val_str = str(val).strip()
                    if val_str and val_str.lower() != "nan":
                        parts.append(val_str)

            if parts:
                columns.append("__".join(parts))
            else:
                columns.append(f"col_{col_idx}")

        return data_df, columns
