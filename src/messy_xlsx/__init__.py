"""messy-xlsx: A Python library for parsing messy Excel files."""

# ============================================================================
# Imports
# ============================================================================

import atexit as _atexit
import warnings as _warnings
import weakref as _weakref
from collections import deque as _deque
from collections.abc import Iterator as _Iterator
from pathlib import Path
from typing import Any, BinaryIO, cast

import pandas as pd
import pyarrow as pa

from messy_xlsx._fallback_signals import _exception_traceback
from messy_xlsx.enums import (
    DataType,
    FormatType,
    HeaderDetectionMode,
    HeaderFallback,
    MergeStrategy,
)
from messy_xlsx.exceptions import (
    CircularReferenceError,
    FileError,
    FormatError,
    FormulaError,
    MessyXlsxError,
    NormalizationError,
    StreamingTypeError,
    StructureError,
    UnsupportedFunctionError,
)
from messy_xlsx.formulas.config import (
    CircularRefStrategy,
    FormulaConfig,
    FormulaEvaluationMode,
)
from messy_xlsx.models import (
    CellValue,
    FormatInfo,
    SheetConfig,
    SheetError,
    SheetResult,
    StructureInfo,
)
from messy_xlsx.multi_sheet import (
    MultiSheetOptions,
    MultiSheetParser,
    SheetInfo,
    analyze_excel,
    read_all_sheets,
)
from messy_xlsx.parsing.streams import BatchStream, DataFrameChunkStream, SheetStream
from messy_xlsx.parsing.streams import _run_cleanups as _run_stream_cleanups
from messy_xlsx.sheet import MessyTable as _MessyTable
from messy_xlsx.utils import sanitize_column_name
from messy_xlsx.warnings import LegacyAPIWarning
from messy_xlsx.warnings import warn_legacy as _warn_legacy
from messy_xlsx.workbook import MessyWorkbook

_MESSY_TABLE_TO_DATAFRAME = _MessyTable.to_dataframe
_MESSY_WORKBOOK_TO_DATAFRAME = MessyWorkbook.to_dataframe
_MAX_TOP_LEVEL_ORPHAN_OWNERS = 64


class _TopLevelBatchOwner:
    """Independent child/workbook cleanup state for one convenience stream."""

    def __init__(self) -> None:
        self.workbook: Any | None = None
        self.child: Any | None = None
        self.finalizer: _weakref.finalize | None = None

    def attach_workbook(self, workbook: Any) -> None:
        if self.workbook is not None:
            raise RuntimeError("top-level workbook owner is already attached")
        self.workbook = workbook

    def attach_child(self, child: Any) -> None:
        if self.child is not None:
            raise RuntimeError("top-level child owner is already attached")
        self.child = child

    def close_child(self) -> None:
        child = self.child
        if child is None:
            return
        child.close()
        self.child = None
        self._detach_if_complete()

    def close_workbook(self) -> None:
        workbook = self.workbook
        if workbook is None:
            return
        workbook.close()
        self.workbook = None
        self._detach_if_complete()

    def close(self) -> None:
        cleanups: list[tuple[str, Any]] = []
        if self.child is not None:
            cleanups.append(("owned child stream cleanup", self.close_child))
        if self.workbook is not None:
            cleanups.append(("owned workbook cleanup", self.close_workbook))
        _run_stream_cleanups(cleanups)
        self._detach_if_complete()

    @property
    def pending(self) -> bool:
        return self.child is not None or self.workbook is not None

    def _detach_if_complete(self) -> None:
        if self.pending:
            return
        finalizer = self.finalizer
        if finalizer is not None:
            finalizer.detach()
            self.finalizer = None


class _OwnedBatchSource(_Iterator[pa.RecordBatch]):
    """Delegate iteration while the independent owner controls child close."""

    def __init__(self, owner: _TopLevelBatchOwner) -> None:
        self._owner = owner

    def __iter__(self) -> "_OwnedBatchSource":
        return self

    def __next__(self) -> pa.RecordBatch:
        child = self._owner.child
        if child is None:
            raise StopIteration
        return next(child)

    def close(self) -> None:
        self._owner.close_child()


_TOP_LEVEL_ORPHAN_OWNERS: _deque[_TopLevelBatchOwner] = _deque()


def _queue_top_level_orphan(owner: _TopLevelBatchOwner) -> None:
    if not owner.pending or any(candidate is owner for candidate in _TOP_LEVEL_ORPHAN_OWNERS):
        return
    if len(_TOP_LEVEL_ORPHAN_OWNERS) >= _MAX_TOP_LEVEL_ORPHAN_OWNERS:
        _drain_top_level_orphans()
    if len(_TOP_LEVEL_ORPHAN_OWNERS) < _MAX_TOP_LEVEL_ORPHAN_OWNERS:
        _TOP_LEVEL_ORPHAN_OWNERS.append(owner)


def _drain_top_level_orphans() -> None:
    pending = tuple(_TOP_LEVEL_ORPHAN_OWNERS)
    _TOP_LEVEL_ORPHAN_OWNERS.clear()
    for owner in pending:
        try:
            owner.close()
        except BaseException:
            if len(_TOP_LEVEL_ORPHAN_OWNERS) < _MAX_TOP_LEVEL_ORPHAN_OWNERS:
                _TOP_LEVEL_ORPHAN_OWNERS.append(owner)


def _finalize_top_level_owner(owner: _TopLevelBatchOwner) -> None:
    try:
        owner.close()
    except BaseException:
        _queue_top_level_orphan(owner)


_atexit.register(_drain_top_level_orphans)

# ============================================================================
# Package Metadata
# ============================================================================

__version__ = "0.10.0"

__all__ = [
    "BatchStream",
    "CellValue",
    "CircularRefStrategy",
    "CircularReferenceError",
    "DataFrameChunkStream",
    "DataType",
    "FileError",
    "FormatError",
    "FormatInfo",
    "FormatType",
    "FormulaConfig",
    "FormulaError",
    "FormulaEvaluationMode",
    "HeaderDetectionMode",
    "HeaderFallback",
    "LegacyAPIWarning",
    "MergeStrategy",
    "MessyWorkbook",
    "MessyXlsxError",
    "MultiSheetOptions",
    "MultiSheetParser",
    "NormalizationError",
    "SheetConfig",
    "SheetError",
    "SheetInfo",
    "SheetResult",
    "SheetStream",
    "StreamingTypeError",
    "StructureError",
    "StructureInfo",
    "UnsupportedFunctionError",
    "analyze_excel",
    "analyze_structure",
    "read_all_sheets",
    "read_excel",
    "read_excel_arrow",
    "read_excel_batches",
    "read_excel_tables",
    "sanitize_column_name",
]


# ============================================================================
# Convenience Functions
# ============================================================================


def _workbook_to_dataframe_compat(
    workbook: MessyWorkbook,
    sheet: str | None,
) -> pd.DataFrame:
    to_dataframe = workbook.to_dataframe
    if (
        getattr(to_dataframe, "__func__", None) is _MESSY_WORKBOOK_TO_DATAFRAME
        and getattr(to_dataframe, "__self__", None) is workbook
    ):
        return workbook._to_dataframe_compat(sheet=sheet)
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore", LegacyAPIWarning)
        return to_dataframe(sheet=sheet)


def _table_to_dataframe_compat(table: _MessyTable) -> pd.DataFrame:
    to_dataframe = table.to_dataframe
    if (
        getattr(to_dataframe, "__func__", None) is _MESSY_TABLE_TO_DATAFRAME
        and getattr(to_dataframe, "__self__", None) is table
    ):
        compat = getattr(table, "_to_dataframe_compat", None)
        assert callable(compat)
        return cast(pd.DataFrame, compat())
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore", LegacyAPIWarning)
        return to_dataframe()


def read_excel(file_path: str, sheet: str | None = None, **config_kwargs: Any) -> pd.DataFrame:
    """Quick function to read an Excel file to a pandas DataFrame."""
    _warn_legacy("read_excel")
    config = SheetConfig(**config_kwargs) if config_kwargs else None
    with MessyWorkbook(file_path, sheet_config=config) as wb:
        return _workbook_to_dataframe_compat(wb, sheet)


def read_excel_arrow(
    file_path_or_buffer: str | Path | BinaryIO,
    sheet: str | None = None,
    config: SheetConfig | None = None,
    filename: str | None = None,
) -> pa.Table:
    """Read one sheet into a materialized Arrow table."""
    _validate_arrow_config(config)
    with MessyWorkbook(file_path_or_buffer, filename=filename) as workbook:
        return workbook.to_arrow(sheet, config)


def read_excel_batches(
    file_path_or_buffer: str | Path | BinaryIO,
    sheet: str | None = None,
    batch_size: int = 65_536,
    config: SheetConfig | None = None,
    filename: str | None = None,
) -> BatchStream:
    """Open one owned workbook and return its closable batch stream."""
    _validate_batch_size(batch_size)
    _validate_arrow_config(config)
    _drain_top_level_orphans()
    owner = _TopLevelBatchOwner()
    stream: BatchStream | None = None

    try:
        workbook_type = MessyWorkbook
        if isinstance(workbook_type, type):
            workbook = workbook_type.__new__(workbook_type)
            owner.attach_workbook(workbook)
            workbook_type.__init__(
                workbook,
                file_path_or_buffer,
                filename=filename,
            )
        else:
            workbook = workbook_type(file_path_or_buffer, filename=filename)
            owner.attach_workbook(workbook)
        child = workbook.iter_batches(sheet, batch_size, config)
        owner.attach_child(child)
        display_names = getattr(child, "_display_names", tuple(child.schema.names))
        stream = BatchStream(
            _OwnedBatchSource(owner),
            child.schema,
            owner.close_workbook,
        )
        stream._display_names = display_names
        owner.finalizer = _weakref.finalize(
            stream,
            _finalize_top_level_owner,
            owner,
        )
        return stream
    except BaseException as error:
        if stream is not None:
            cleanups = [("partially constructed batch stream cleanup", stream.close)]
        else:
            cleanups = [("top-level batch owner cleanup", owner.close)]
        _run_stream_cleanups(
            cleanups,
            primary_error=error,
            primary_traceback=_exception_traceback(error),
        )
        raise


def _validate_arrow_config(config: object) -> None:
    if config is not None and not isinstance(config, SheetConfig):
        raise TypeError("config must be a SheetConfig or None")


def _validate_batch_size(batch_size: object) -> None:
    if not isinstance(batch_size, int) or isinstance(batch_size, bool) or batch_size < 1:
        raise ValueError("batch_size must be >= 1")


def read_excel_tables(file_path: str, sheet: str | None = None) -> list[pd.DataFrame]:
    """Read all detected tables from a sheet."""
    _warn_legacy("read_excel_tables")
    with MessyWorkbook(file_path) as wb:
        sheet_obj = wb.get_sheet(sheet or wb.sheet_names[0])
        return [_table_to_dataframe_compat(table) for table in sheet_obj.tables]


def analyze_structure(file_path: str, sheet: str | None = None) -> StructureInfo:
    """Analyze and return structure info without loading data."""
    with MessyWorkbook(file_path) as wb:
        return wb.get_structure(sheet or wb.sheet_names[0])
