"""MessyWorkbook - Main entry point for parsing Excel files."""

# ============================================================================
# Imports
# ============================================================================

import logging
import re
import weakref
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from types import TracebackType
from typing import Any, BinaryIO, Generic, TypeVar

import openpyxl
import pandas as pd

from messy_xlsx._fallback_signals import (
    _blocks_backend_retry,
    _contains_process_failure,
    _exception_traceback,
    _FallbackBlockReason,
    _mark_fallback_blocked,
)
from messy_xlsx._source import BackendSource, SourceHandle, describe_source
from messy_xlsx.cache import get_structure_cache
from messy_xlsx.detection.structure_analyzer import StructureAnalyzer
from messy_xlsx.exceptions import FileError, FormatError
from messy_xlsx.formulas.config import FormulaConfig, FormulaEvaluationMode
from messy_xlsx.formulas.engine import FormulaEngine
from messy_xlsx.models import CellValue, SheetConfig, SheetError, StructureInfo
from messy_xlsx.normalization.pipeline import NormalizationPipeline
from messy_xlsx.ooxml.manifest import ManifestReader
from messy_xlsx.ooxml.models import IntervalIndex, SheetManifest
from messy_xlsx.parsing.coordinates import CoordinateTransform
from messy_xlsx.parsing.fastexcel_session import FastexcelSession
from messy_xlsx.parsing.handler_registry import HandlerRegistry
from messy_xlsx.parsing.parse_plan import (
    ParsePlan,
    compile_parse_plan,
    requires_structure_analysis,
)
from messy_xlsx.parsing.streams import _close_if_present, _run_cleanups
from messy_xlsx.parsing.xlsx_handler import (
    XLSXHandler,
    _is_fastexcel_materialized_plan,
)
from messy_xlsx.sheet import MessySheet
from messy_xlsx.warnings import warn_legacy

# ============================================================================
# Core
# ============================================================================

logger = logging.getLogger(__name__)
_NO_BUILTIN_MATERIALIZATION = object()


class _ActiveOperationError(RuntimeError):
    """Identify a rejected concurrent or re-entrant workbook operation."""


def _operation_error(message: str) -> _ActiveOperationError:
    error = _ActiveOperationError(message)
    _mark_fallback_blocked(
        error,
        _FallbackBlockReason.CONFIGURATION,
    )
    return error


_OwnedResourceT = TypeVar("_OwnedResourceT")


class _CloseOnceOwner(Generic[_OwnedResourceT]):
    """Proxy one closeable reader through a shared close-once boundary."""

    def __init__(self, resource: _OwnedResourceT) -> None:
        self._resource: _OwnedResourceT | None = resource
        self._closed = False

    def __getattr__(self, name: str) -> Any:
        resource = self._resource
        if resource is None:
            raise AttributeError(name)
        return getattr(resource, name)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        resource = self._resource
        self._resource = None
        if resource is not None:
            _close_if_present(resource)


class _StreamOperationLease:
    """Weak workbook lease used while constructing and owning one child stream."""

    def __init__(self, workbook: "MessyWorkbook", token: object) -> None:
        self._workbook_ref = weakref.ref(workbook)
        self._token = token
        self._partial: _CloseOnceOwner[Any] | None = None
        self._stream_ref: weakref.ReferenceType[Any] | None = None
        self._bound = False
        self._released = False

    def own(self, partial: _OwnedResourceT) -> _CloseOnceOwner[_OwnedResourceT]:
        """Register a partially opened reader for construction-failure cleanup."""
        if self._released or self._bound or self._partial is not None:
            raise RuntimeError("Stream operation lease already owns a resource")
        owner = _CloseOnceOwner(partial)
        self._partial = owner
        return owner

    def bind(self, stream: Any) -> Any:
        """Register a stream while retaining lease-owned reader cleanup."""
        if self._released or self._bound:
            raise RuntimeError("Stream operation lease is no longer available")
        stream_ref = weakref.ref(stream)
        self._stream_ref = stream_ref
        workbook = self._workbook_ref()
        if workbook is None:
            raise _operation_error("MessyWorkbook is closed")
        workbook._register_stream(self._token, stream)
        self._bound = True
        return stream

    def release(self) -> None:
        """Release the matching reservation at most once."""
        if self._released:
            return
        self._released = True
        partial = self._partial
        self._partial = None
        self._stream_ref = None
        workbook = self._workbook_ref()
        cleanups: list[tuple[str, Any]] = []
        if partial is not None:
            cleanups.append(("owned stream reader cleanup", lambda: _close_if_present(partial)))
        if workbook is not None:
            cleanups.append(
                ("stream reservation release", lambda: workbook._end_operation(self._token))
            )
        _run_cleanups(cleanups)

    def __enter__(self) -> "_StreamOperationLease":
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        del exc_type
        if self._bound and exc_value is None:
            return

        stream = self._stream_ref() if self._stream_ref is not None else None
        cleanups: list[tuple[str, Any]] = []
        if stream is not None:
            cleanups.append(
                ("partially registered stream cleanup", lambda: _close_if_present(stream))
            )
        cleanups.append(("stream reservation release", self.release))
        _run_cleanups(
            cleanups,
            primary_error=exc_value,
            primary_traceback=traceback,
        )


class MessyWorkbook:
    """Main entry point for parsing Excel files."""

    def __init__(
        self,
        file_path_or_buffer: str | Path | BinaryIO,
        sheet_config: SheetConfig | None = None,
        formula_config: FormulaConfig | None = None,
        filename: str | None = None,
        registry: HandlerRegistry | None = None,
    ):
        """Open an Excel file for parsing.

        Args:
            file_path_or_buffer: Path to file, or file-like object (BytesIO, etc.)
            sheet_config: Configuration for parsing sheets
            formula_config: Configuration for formula evaluation
            filename: Optional filename hint when using file-like objects (for format detection)
            registry: Optional format-handler registry for custom parsing behavior
        """
        self._closed = False
        self._active_operation_token: object | None = None
        self._active_stream: Any | None = None
        self._sheets: dict[str, MessySheet] = {}
        self._formula_loaded = False
        self._wb: openpyxl.Workbook | None = None
        self._cached_wb: openpyxl.Workbook | None = None
        self._wb_source: BinaryIO | None = None
        self._cached_wb_source: BinaryIO | None = None
        self._fastexcel_session: FastexcelSession | None = None
        self._manifest_reader: ManifestReader | None = None

        self._sheet_config = sheet_config or SheetConfig()
        self._formula_config = formula_config or FormulaConfig()

        self._registry = registry if registry is not None else HandlerRegistry()
        self._analyzer = StructureAnalyzer(get_structure_cache())
        self._formula_engine = FormulaEngine(self._formula_config)

        try:
            self._source_handle = SourceHandle(file_path_or_buffer, filename=filename)
        except Exception as e:
            file_desc = describe_source(file_path_or_buffer, filename)
            raise FormatError(
                f"Cannot read from file object: {e}",
                file_path=file_desc,
            ) from e
        try:
            self._initialize_source()
        except BaseException as error:
            self._close(
                primary_error=error,
                primary_traceback=_exception_traceback(error),
            )
            raise

    def _begin_operation(self) -> object:
        """Reserve the workbook for one parse or child stream."""
        if getattr(self, "_closed", False):
            raise _operation_error("MessyWorkbook is closed")
        if getattr(self, "_active_operation_token", None) is not None:
            raise _operation_error("MessyWorkbook already has an active parse or stream")
        token = object()
        self._active_operation_token = token
        return token

    def _end_operation(self, token: object) -> None:
        """Release only the exact current token; stale callbacks are harmless."""
        if getattr(self, "_active_operation_token", None) is not token:
            return
        self._active_operation_token = None
        self._active_stream = None

    def _register_stream(self, token: object, stream: Any) -> None:
        """Register a child only while its exact operation token is current."""
        if getattr(self, "_closed", False):
            raise _operation_error("MessyWorkbook is closed")
        if getattr(self, "_active_operation_token", None) is not token:
            raise _operation_error("MessyWorkbook stream reservation is no longer active")
        if getattr(self, "_active_stream", None) is not None:
            raise _operation_error("MessyWorkbook already has an active parse or stream")
        self._active_stream = stream

    def _stream_operation(self) -> _StreamOperationLease:
        """Reserve a future child stream with construction-failure cleanup."""
        token = self._begin_operation()
        try:
            return _StreamOperationLease(self, token)
        except BaseException:
            self._end_operation(token)
            raise

    @contextmanager
    def _registry_source(self) -> Iterator[SourceHandle | BackendSource]:
        """Adapt the internal handle for legacy registry subclasses."""
        accepts_handle = bool(type(self._registry).__dict__.get("_accepts_source_handle", False))
        if accepts_handle:
            yield self._source_handle
            return
        with self._source_handle.open_legacy() as source:
            yield source

    def _initialize_source(self) -> None:
        """Detect, inspect, and validate the source without taking ownership."""

        if self._source_handle.path is not None and not self._source_handle.path.exists():
            raise FileError(
                f"File not found: {self._source_handle.path}",
                file_path=str(self._source_handle.path),
            )

        with self._registry_source() as source:
            self._format_info = self._registry.detect_format(
                source,
                filename=self._source_handle.filename_hint,
            )

        if self._format_info.format_type == "unknown":
            file_desc = self._source_handle.description
            raise FormatError(
                f"Unknown file format: {file_desc}",
                file_path=str(file_desc),
            )

        if self._format_info.format_type == "xlsb":
            file_desc = self._source_handle.description
            raise FormatError(
                "XLSB (Excel Binary) format is not supported. "
                "Please convert the file to XLSX format.",
                file_path=str(file_desc),
                detected_format="xlsb",
            )

        # Validate extension matches detected format for Excel files
        # This catches files with .xlsx extension but different content
        if self._source_handle.path is not None:
            file_ext = self._source_handle.path.suffix.lower()
            excel_extensions = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}
            if file_ext in excel_extensions and self._format_info.format_type not in (
                "xlsx",
                "xlsm",
                "xls",
                "xltx",
                "xltm",
            ):
                raise FormatError(
                    f"File extension {file_ext} suggests Excel format, but content is {self._format_info.format_type}",
                    file_path=str(self._source_handle.path),
                    detected_format=self._format_info.format_type,
                )

        # Get sheet names and validate file is readable
        with self._registry_source() as source:
            self._sheet_names = self._registry.get_sheet_names(
                source,
            )

        # Validate that the file is actually readable (not just format-detected)
        # This catches corrupted files that pass format detection but can't be opened
        if self._format_info.format_type in ("xlsx", "xlsm", "xltx", "xltm", "xls"):
            with self._registry_source() as source:
                is_valid, error = self._registry.validate(
                    source,
                    self._format_info.format_type,
                )
            if not is_valid:
                file_desc = self._source_handle.description
                raise FormatError(
                    f"File appears corrupted or invalid: {error}",
                    file_path=str(file_desc),
                    detected_format=self._format_info.format_type,
                )

    @property
    def file_path(self) -> Path | None:
        """Path to the Excel file, or None if reading from buffer."""
        return self._source_handle.path

    @property
    def source(self) -> Path | BinaryIO:
        """The source file path or buffer."""
        return self._source_handle.original

    @property
    def sheet_names(self) -> list[str]:
        """List of sheet names in the workbook."""
        return self._sheet_names.copy()

    @property
    def format_type(self) -> str:
        """Detected file format (xlsx, xls, csv, etc.)."""
        return self._format_info.format_type

    def get_sheet(self, name: str | None = None) -> MessySheet:
        """Get a sheet by name."""
        if name is None:
            name = self._sheet_names[0]

        if name not in self._sheet_names:
            file_desc = self._source_handle.description
            raise FormatError(
                f"Sheet '{name}' not found",
                file_path=str(file_desc),
            )

        if name not in self._sheets:
            self._sheets[name] = MessySheet(self, name)

        return self._sheets[name]

    def to_dataframe(
        self,
        sheet: str | None = None,
        config: SheetConfig | None = None,
    ) -> pd.DataFrame:
        """Convert a sheet to a pandas DataFrame."""
        warn_legacy("MessyWorkbook.to_dataframe")
        return self._to_dataframe_compat(sheet, config)

    def _to_dataframe_compat(
        self,
        sheet: str | None = None,
        config: SheetConfig | None = None,
    ) -> pd.DataFrame:
        sheet_name = sheet or self._sheet_names[0]
        return self._parse_sheet(sheet_name, config)

    def to_dataframes(
        self,
        config: SheetConfig | None = None,
        include_errors: bool = False,
    ) -> dict[str, pd.DataFrame] | tuple[dict[str, pd.DataFrame], list[SheetError]]:
        """Convert all sheets to DataFrames.

        Args:
            config: Optional sheet configuration.
            include_errors: If True, return a tuple of (results, errors) instead
                of just results. Each error contains structured information about
                which sheet failed and why.

        Returns:
            If include_errors is False (default): dict mapping sheet name to DataFrame.
            If include_errors is True: tuple of (results_dict, errors_list).
        """
        warn_legacy("MessyWorkbook.to_dataframes")
        return self._to_dataframes_compat(config, include_errors)

    def _to_dataframes_compat(
        self,
        config: SheetConfig | None = None,
        include_errors: bool = False,
    ) -> dict[str, pd.DataFrame] | tuple[dict[str, pd.DataFrame], list[SheetError]]:
        token = self._begin_operation()
        try:
            result = {}
            errors: list[SheetError] = []
            for name in self._sheet_names:
                try:
                    result[name] = self._parse_sheet_unreserved(name, config)
                except _ActiveOperationError:
                    raise
                except Exception as e:
                    if _contains_process_failure(e):
                        raise
                    logger.warning("Failed to parse sheet %r, skipping", name, exc_info=True)
                    if include_errors:
                        context = {}
                        if hasattr(e, "context"):
                            context = e.context
                        errors.append(
                            SheetError(
                                sheet_name=name,
                                error_type=type(e).__name__,
                                message=str(e),
                                context=context,
                            )
                        )
            if include_errors:
                return result, errors
            return result
        finally:
            self._end_operation(token)

    def get_structure(self, sheet: str | None = None) -> StructureInfo:
        """Get detected structure for a sheet."""
        sheet_name = sheet or self._sheet_names[0]
        return self._analyze_structure(sheet_name, self._sheet_config)

    def get_cell(
        self,
        sheet: str,
        row: int,
        col: int,
    ) -> CellValue:
        """Get a single cell value."""
        self._ensure_workbook()

        if self._wb is None:
            raise FileError("Workbook not loaded — call _ensure_workbook() first")
        ws = self._wb[sheet]
        cell = ws.cell(row, col)

        resolved_value = cell.value

        formula = None
        is_formula = False
        if hasattr(cell, "data_type") and cell.data_type == "f":
            is_formula = True
            if (
                hasattr(cell, "value")
                and isinstance(cell.value, str)
                and cell.value.startswith("=")
            ):
                formula = cell.value

        if is_formula and self._formula_config.mode != FormulaEvaluationMode.DISABLED:
            cached_value = self._get_cached_cell_value(sheet, row, col)
            resolved_value = cached_value
            self._ensure_formula_engine()
            try:
                resolved_value = self._formula_engine.evaluate(sheet, row, col, cached_value)
            except (ValueError, TypeError, KeyError) as e:
                logger.debug(
                    "Formula evaluation failed for cell (%s, %d, %d): %s", sheet, row, col, e
                )

        data_type = self._get_data_type(resolved_value)

        is_merged = self._is_cell_merged(ws, row, col)

        is_hidden = self._is_cell_hidden(ws, row, col)

        return CellValue(
            value=resolved_value,
            formula=formula,
            is_merged=is_merged,
            is_hidden=is_hidden,
            data_type=data_type,
            original_format=cell.number_format if hasattr(cell, "number_format") else None,
        )

    def get_cell_by_ref(self, ref: str) -> CellValue:
        """Get a cell by A1-style reference."""
        from messy_xlsx.utils import cell_ref_to_coords

        sheet, row, col = cell_ref_to_coords(ref)
        sheet = sheet or self._sheet_names[0]
        return self.get_cell(sheet, row, col)

    def _parse_sheet(
        self,
        sheet: str,
        config: SheetConfig | None = None,
    ) -> pd.DataFrame:
        """Parse a sheet to DataFrame with normalization."""
        token = self._begin_operation()
        try:
            return self._parse_sheet_unreserved(sheet, config)
        finally:
            self._end_operation(token)

    def _parse_sheet_unreserved(
        self,
        sheet: str,
        config: SheetConfig | None = None,
    ) -> pd.DataFrame:
        """Parse one sheet while the caller owns the workbook reservation."""
        config = config or self._sheet_config
        format_type = self.format_type

        structure = None
        if requires_structure_analysis(config, format_type):
            structure = self._analyze_structure(sheet, config)
        plan = compile_parse_plan(config, structure, format_type)

        built_in = self._parse_builtin_materialized(sheet, format_type, plan)
        if built_in is _NO_BUILTIN_MATERIALIZATION:
            with self._registry_source() as source:
                df = self._registry.parse(
                    source,
                    sheet=sheet,
                    options=plan.to_parse_options(),
                    format_type=format_type,
                )
        else:
            assert isinstance(built_in, pd.DataFrame)
            df = built_in

        if plan.normalize:
            pipeline = NormalizationPipeline(
                decimal_separator=plan.decimal_separator,
                thousands_separator=plan.thousands_separator,
                use_extended_missing_list=plan.use_extended_missing_list,
                preserve_types=plan.preserve_types,
            )

            df = pipeline.normalize(
                df,
                semantic_hints=plan.thaw_type_hints(),
                skip_steps=list(plan.skip_normalization_steps),
            )

        # Sanitize column names if requested
        if plan.sanitize_column_names:
            df = self._sanitize_columns(df)

        # Apply user renames (user overrides take precedence)
        if plan.column_renames:
            df = df.rename(columns=plan.thaw_column_renames())

        # Preserve the legacy behavior where disabling normalization also
        # bypasses row filters. S15 owns any change to that public contract.
        if not plan.normalize:
            return df

        # Drop rows matching regex pattern
        if plan.drop_regex and not df.empty:
            pattern = re.compile(plan.drop_regex)
            mask = df.apply(
                lambda row: any(
                    bool(pattern.search(str(v)))
                    for v in row
                    if v is not None and not (isinstance(v, float) and pd.isna(v))
                ),
                axis=1,
            )
            df = df[~mask].reset_index(drop=True)

        # Drop rows matching column-value conditions
        if plan.drop_conditions and not df.empty:
            for col, value in plan.thaw_drop_conditions():
                if col is not None and col in df.columns:
                    df = df[df[col] != value].reset_index(drop=True)

        return df

    def _parse_builtin_materialized(
        self,
        sheet: str,
        format_type: str,
        plan: ParsePlan,
    ) -> pd.DataFrame | object:
        """Use the bound-plan seam only for the untouched built-in XLSX stack."""
        if format_type not in {"xlsx", "xlsm"}:
            return _NO_BUILTIN_MATERIALIZATION
        if type(self._registry) is not HandlerRegistry:
            return _NO_BUILTIN_MATERIALIZATION
        if not self._registry._uses_builtin_components():
            return _NO_BUILTIN_MATERIALIZATION
        handler = self._registry.get_handler(format_type)
        if type(handler) is not XLSXHandler:
            return _NO_BUILTIN_MATERIALIZATION
        if not _is_fastexcel_materialized_plan(plan):
            return _NO_BUILTIN_MATERIALIZATION
        transform: CoordinateTransform | None = None
        coordinate_features = (
            plan.merge_strategy != "skip" or plan.ignore_hidden or bool(plan.cell_range)
        )
        if coordinate_features:
            if not self._coordinate_range_is_supported(plan):
                return _NO_BUILTIN_MATERIALIZATION
            if sheet not in self._sheet_names:
                return _NO_BUILTIN_MATERIALIZATION
            manifest = self._get_sheet_manifest(sheet)
            if not self._manifest_supports_coordinate_plan(manifest, plan):
                return _NO_BUILTIN_MATERIALIZATION
            transform = CoordinateTransform.from_manifest(manifest)
        try:
            return handler._parse_materialized_plan(
                self._source_handle,
                sheet,
                plan,
                self._get_fastexcel_session,
                transform=transform,
            )
        except Exception as error:
            if _blocks_backend_retry(error):
                raise

        file_desc = self._source_handle.description
        name = self._source_handle.path.name if self._source_handle.path is not None else file_desc
        raise FormatError(
            f"All handlers failed for {name}",
            file_path=file_desc,
            detected_format=format_type,
            attempted_formats=[type(handler).__name__],
        )

    def _get_fastexcel_session(self) -> FastexcelSession:
        """Return the workbook-owned session shared by eligible sheet reads."""
        session = self._fastexcel_session
        if session is None:
            session = FastexcelSession(self._source_handle)
            self._fastexcel_session = session
        return session

    def _get_sheet_manifest(self, sheet: str) -> SheetManifest:
        reader = self._manifest_reader
        if reader is None:
            reader = ManifestReader(self._source_handle)
            self._manifest_reader = reader
        return reader.sheet(sheet)

    @staticmethod
    def _coordinate_range_is_supported(plan: ParsePlan) -> bool:
        if not plan.cell_range:
            return True
        try:
            CoordinateTransform(
                hidden_rows=IntervalIndex(()),
                hidden_columns=IntervalIndex(()),
                merged_ranges=(),
            ).open(plan)
        except ValueError:
            return False
        return True

    @staticmethod
    def _manifest_supports_coordinate_plan(
        manifest: SheetManifest,
        plan: ParsePlan,
    ) -> bool:
        if not plan.ignore_hidden or bool(plan.cell_range):
            return True
        if any(interval.start != interval.end for interval in manifest.hidden_columns.intervals):
            return False
        if manifest.observed_max_col == 0:
            return True
        intervals = manifest.hidden_columns.intervals
        return not (
            len(intervals) == 1
            and intervals[0].start == 1
            and intervals[0].end >= manifest.observed_max_col
        )

    def _analyze_structure(self, sheet: str, config: SheetConfig | None = None) -> StructureInfo:
        """Analyze sheet structure."""
        header_patterns = config.header_patterns if config else None
        return self._analyzer.analyze(
            self._source_handle,
            sheet,
            header_patterns=header_patterns,
        )

    def _sanitize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Sanitize column names for BigQuery compatibility."""
        from .utils import sanitize_column_name

        new_columns = []
        seen: dict[str, int] = {}

        for col in df.columns:
            clean = sanitize_column_name(col)

            # Handle duplicates by appending counter
            if clean in seen:
                seen[clean] += 1
                clean = f"{clean}_{seen[clean]}"
            else:
                seen[clean] = 0

            new_columns.append(clean)

        df.columns = new_columns
        return df

    def _ensure_formula_engine(self) -> None:
        """Lazily load formula engine on first get_cell() call."""
        if self._formula_loaded:
            return
        self._formula_loaded = True

        if (
            self._formula_config.mode != FormulaEvaluationMode.DISABLED
            and self._formula_engine.is_available
        ):
            try:
                if self._source_handle.path is not None:
                    self._formula_engine.load_workbook(self._source_handle.path)
            except (OSError, ValueError, TypeError) as e:
                logger.debug("Formula engine load failed: %s", e)

    def _ensure_workbook(self) -> None:
        """Ensure openpyxl workbook is loaded."""
        if self._wb is None:
            source: Path | BinaryIO
            owned_source: BinaryIO | None
            if self._source_handle.path is None:
                owned_source = self._source_handle.detached_binary()
                source = owned_source
            else:
                owned_source = None
                source = self._source_handle.path
            try:
                self._wb = openpyxl.load_workbook(
                    source,
                    read_only=False,
                    data_only=False,
                )
            except BaseException:
                if owned_source is not None:
                    owned_source.close()
                raise
            self._wb_source = owned_source

    def _get_cached_cell_value(self, sheet: str, row: int, col: int) -> Any:
        """Read a formula's cached result from a data-only workbook view."""
        if self._cached_wb is None:
            source: Path | BinaryIO
            owned_source: BinaryIO | None
            if self._source_handle.path is None:
                owned_source = self._source_handle.detached_binary()
                source = owned_source
            else:
                owned_source = None
                source = self._source_handle.path
            try:
                self._cached_wb = openpyxl.load_workbook(
                    source,
                    read_only=False,
                    data_only=True,
                )
            except BaseException:
                if owned_source is not None:
                    owned_source.close()
                raise
            self._cached_wb_source = owned_source

        return self._cached_wb[sheet].cell(row, col).value

    def _get_data_type(self, value: Any) -> str:
        """Determine data type string for a value."""
        if value is None:
            return "empty"
        if isinstance(value, bool):
            return "boolean"
        if isinstance(value, (int, float)):
            return "number"
        if isinstance(value, str):
            if value.startswith("#") and value.endswith("!"):
                return "error"
            return "text"
        if hasattr(value, "date"):
            return "date"
        return "text"

    def _is_cell_merged(self, ws: Any, row: int, col: int) -> bool:
        """Check if cell is part of a merged range."""
        try:
            for merged_range in ws.merged_cells.ranges:
                if (
                    merged_range.min_row <= row <= merged_range.max_row
                    and merged_range.min_col <= col <= merged_range.max_col
                ):
                    return True
        except (AttributeError, TypeError):
            pass
        return False

    def _is_cell_hidden(self, ws: Any, row: int, col: int) -> bool:
        """Check if cell is in a hidden row or column."""
        try:
            if row in ws.row_dimensions and ws.row_dimensions[row].hidden:
                return True
            from openpyxl.utils import get_column_letter

            col_letter = get_column_letter(col)
            if col_letter in ws.column_dimensions and ws.column_dimensions[col_letter].hidden:
                return True
        except (AttributeError, TypeError):
            pass
        return False

    def close(self) -> None:
        """Close the workbook and release resources."""
        self._close()

    def _close(
        self,
        *,
        primary_error: BaseException | None = None,
        primary_traceback: TracebackType | None = None,
    ) -> None:
        """Close once, preserving operations unless process cleanup must win."""
        if getattr(self, "_closed", False):
            return
        self._closed = True

        active_stream = getattr(self, "_active_stream", None)
        fastexcel_session = getattr(self, "_fastexcel_session", None)
        workbook = getattr(self, "_wb", None)
        cached_workbook = getattr(self, "_cached_wb", None)
        workbook_source = getattr(self, "_wb_source", None)
        cached_workbook_source = getattr(self, "_cached_wb_source", None)
        source_handle = getattr(self, "_source_handle", None)
        self._active_stream = None
        self._active_operation_token = None
        self._fastexcel_session = None
        self._manifest_reader = None
        self._wb = None
        self._cached_wb = None
        self._wb_source = None
        self._cached_wb_source = None

        cleanups: list[tuple[str, Any]] = []
        if active_stream is not None:
            invalidate = getattr(active_stream, "invalidate_from_owner", None)
            if callable(invalidate):
                cleanups.append(("active stream invalidation", invalidate))
        if fastexcel_session is not None:
            cleanups.append(
                ("fastexcel session cleanup", lambda: _close_if_present(fastexcel_session))
            )
        if workbook is not None:
            cleanups.append(("workbook cleanup", lambda: _close_if_present(workbook)))
        if cached_workbook is not None:
            cleanups.append(("cached workbook cleanup", lambda: _close_if_present(cached_workbook)))
        if workbook_source is not None:
            cleanups.append(("workbook source cleanup", lambda: _close_if_present(workbook_source)))
        if cached_workbook_source is not None:
            cleanups.append(
                (
                    "cached workbook source cleanup",
                    lambda: _close_if_present(cached_workbook_source),
                )
            )
        if source_handle is not None:
            cleanups.append(("source handle cleanup", lambda: _close_if_present(source_handle)))
        _run_cleanups(
            cleanups,
            primary_error=primary_error,
            primary_traceback=primary_traceback,
        )

    def __enter__(self) -> "MessyWorkbook":
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        del exc_type
        if isinstance(exc_val, BaseException):
            self._close(primary_error=exc_val, primary_traceback=exc_tb)
            return
        self.close()

    def __repr__(self) -> str:
        name = (
            self._source_handle.path.name
            if self._source_handle.path is not None
            else self._source_handle.description
        )
        return f"MessyWorkbook({name!r}, sheets={self._sheet_names})"
