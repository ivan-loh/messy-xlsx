"""Resource-ownership contracts for public convenience APIs."""

from __future__ import annotations

import io
from pathlib import Path
from types import SimpleNamespace
from typing import Any, ClassVar

import pandas as pd
import pytest

import messy_xlsx as api
import messy_xlsx.detection.structure_analyzer as structure_analyzer_module
import messy_xlsx.parsing.csv_handler as csv_handler_module
import messy_xlsx.parsing.xls_handler as xls_handler_module
import messy_xlsx.parsing.xlsx_handler as xlsx_handler_module
from messy_xlsx._source import SourceHandle
from messy_xlsx._spool import DEFAULT_MEMORY_LIMIT
from messy_xlsx.detection import StructureAnalyzer
from messy_xlsx.exceptions import FormatError
from messy_xlsx.parsing import (
    CSVHandler,
    MetadataRowDetector,
    ParseOptions,
    XLSHandler,
    XLSXHandler,
)


class _ExpectedFailure(RuntimeError):
    """Failure injected after a workbook has been constructed."""


class _CloseTracker:
    def __init__(self, name: str, fail_on_close: bool = False):
        self.name = name
        self.fail_on_close = fail_on_close
        self.close_calls = 0
        self.closed = False

    def close(self) -> None:
        self.close_calls += 1
        self.closed = True
        if self.fail_on_close:
            raise _ExpectedFailure(f"{self.name} close")


class _BaseCloseTracker:
    def __init__(self, error: BaseException):
        self.error = error
        self.close_calls = 0

    def close(self) -> None:
        self.close_calls += 1
        raise self.error


class _TrackedStringIO(io.StringIO):
    instances: ClassVar[list[_TrackedStringIO]] = []

    def __init__(self, value: str):
        super().__init__(value)
        type(self).instances.append(self)


class _TrackingTable:
    def __init__(self, workbook: _TrackingWorkbook):
        self._workbook = workbook

    def to_dataframe(self) -> pd.DataFrame:
        self._workbook.raise_if_requested("parse")
        return pd.DataFrame({"value": [1]})


class _TrackingSheet:
    def __init__(self, workbook: _TrackingWorkbook):
        self.tables = [_TrackingTable(workbook)]


class _TrackingWorkbook:
    """Small fake exposing only the contract used by convenience functions."""

    instances: ClassVar[list[_TrackingWorkbook]] = []
    fail_at: ClassVar[str | None] = None

    def __init__(self, *_args: Any, **_kwargs: Any):
        self.closed = False
        self.close_calls = 0
        type(self).instances.append(self)

    @classmethod
    def reset(cls, fail_at: str | None = None) -> None:
        cls.instances = []
        cls.fail_at = fail_at

    @property
    def sheet_names(self) -> list[str]:
        return ["Data"]

    def raise_if_requested(self, stage: str) -> None:
        if type(self).fail_at == stage:
            raise _ExpectedFailure(stage)

    def to_dataframe(self, **_kwargs: Any) -> pd.DataFrame:
        self.raise_if_requested("parse")
        return pd.DataFrame({"value": [1]})

    def get_sheet(self, _name: str) -> _TrackingSheet:
        return _TrackingSheet(self)

    def get_structure(self, _name: str) -> object:
        self.raise_if_requested("analysis")
        return object()

    def close(self) -> None:
        self.close_calls += 1
        self.closed = True

    def __enter__(self) -> _TrackingWorkbook:
        return self

    def __exit__(self, *_args: Any) -> None:
        self.close()


def _invoke(api_name: str, source: object) -> object:
    if api_name == "read_excel":
        return api.read_excel(source)  # type: ignore[arg-type]
    if api_name == "read_excel_tables":
        return api.read_excel_tables(source)  # type: ignore[arg-type]
    if api_name == "analyze_structure":
        return api.analyze_structure(source)  # type: ignore[arg-type]
    raise AssertionError(f"Unhandled API: {api_name}")


@pytest.mark.parametrize(
    "api_name",
    ["read_excel", "read_excel_tables", "analyze_structure"],
)
def test_convenience_function_closes_owned_workbook_on_success(
    monkeypatch: pytest.MonkeyPatch,
    api_name: str,
) -> None:
    _TrackingWorkbook.reset()
    monkeypatch.setattr(api, "MessyWorkbook", _TrackingWorkbook)

    _invoke(api_name, "owned.xlsx")

    assert len(_TrackingWorkbook.instances) == 1
    workbook = _TrackingWorkbook.instances[0]
    assert workbook.closed is True
    assert workbook.close_calls == 1


@pytest.mark.parametrize(
    ("api_name", "failure_stage"),
    [
        ("read_excel", "parse"),
        ("read_excel_tables", "parse"),
        ("analyze_structure", "analysis"),
    ],
)
def test_convenience_function_closes_owned_workbook_on_failure(
    monkeypatch: pytest.MonkeyPatch,
    api_name: str,
    failure_stage: str,
) -> None:
    _TrackingWorkbook.reset(fail_at=failure_stage)
    monkeypatch.setattr(api, "MessyWorkbook", _TrackingWorkbook)

    with pytest.raises(_ExpectedFailure, match=failure_stage):
        _invoke(api_name, "owned.xlsx")

    assert len(_TrackingWorkbook.instances) == 1
    workbook = _TrackingWorkbook.instances[0]
    assert workbook.closed is True
    assert workbook.close_calls == 1


@pytest.mark.parametrize(
    "api_name",
    ["read_excel", "read_excel_tables", "analyze_structure"],
)
def test_convenience_function_leaves_caller_buffer_open_and_reusable(
    sample_xlsx: Any,
    api_name: str,
) -> None:
    content = sample_xlsx.read_bytes()
    source = io.BytesIO(content)

    _invoke(api_name, source)

    assert source.closed is False
    source.seek(0)
    assert source.read(4) == content[:4]

    source.seek(0)
    _invoke(api_name, source)
    assert source.closed is False


def test_parse_failure_does_not_transfer_ownership_of_caller_buffer(
    monkeypatch: pytest.MonkeyPatch,
    sample_xlsx: Any,
) -> None:
    real_workbook = api.MessyWorkbook

    class FailingWorkbook(real_workbook):
        def to_dataframe(self, **_kwargs: Any) -> pd.DataFrame:
            # Open the lazy openpyxl archive before failing so cleanup exercises
            # a real library-owned resource backed by a caller-owned stream.
            self._ensure_workbook()
            raise _ExpectedFailure("parse")

    source = io.BytesIO(sample_xlsx.read_bytes())
    monkeypatch.setattr(api, "MessyWorkbook", FailingWorkbook)

    with pytest.raises(_ExpectedFailure, match="parse"):
        api.read_excel(source)  # type: ignore[arg-type]

    assert source.closed is False
    source.seek(0)
    with real_workbook(source, filename=sample_xlsx.name) as workbook:
        result = workbook.to_dataframe()

    assert not result.empty
    assert source.closed is False


def test_csv_caller_buffer_remains_open_and_readable() -> None:
    content = b"Name,Value\nA,1\n"
    source = io.BytesIO(content)

    api.read_excel(source)  # type: ignore[arg-type]

    assert source.closed is False
    source.seek(0)
    assert source.read() == content


def test_legacy_xls_caller_buffer_remains_open_and_readable() -> None:
    xlwt = pytest.importorskip("xlwt")
    legacy_workbook = xlwt.Workbook()
    sheet = legacy_workbook.add_sheet("Data")
    sheet.write(0, 0, "Name")
    sheet.write(0, 1, "Value")
    sheet.write(1, 0, "A")
    sheet.write(1, 1, 1)
    source = io.BytesIO()
    legacy_workbook.save(source)
    content = source.getvalue()
    source.seek(0)

    api.read_excel(source)  # type: ignore[arg-type]

    assert source.closed is False
    source.seek(0)
    assert source.read() == content


def test_source_handle_close_removes_spill_and_is_idempotent() -> None:
    source = io.BytesIO(b"x" * (DEFAULT_MEMORY_LIMIT + 1))
    source.seek(3)
    handle = SourceHandle(source)

    with handle.open_path_or_bytes() as backend:
        assert isinstance(backend, Path)
        spill_path = backend
        assert spill_path.exists()

    handle.close()
    handle.close()

    assert not spill_path.exists()
    assert source.tell() == 3
    assert source.closed is False


@pytest.mark.parametrize("failing_resource", ["primary", "cached"])
def test_workbook_close_attempts_and_clears_both_resources_when_close_raises(
    failing_resource: str,
) -> None:
    workbook = object.__new__(api.MessyWorkbook)
    primary = _CloseTracker("primary", fail_on_close=failing_resource == "primary")
    cached = _CloseTracker("cached", fail_on_close=failing_resource == "cached")
    workbook._wb = primary  # type: ignore[assignment]
    workbook._cached_wb = cached  # type: ignore[assignment]

    # Preserve the existing close-error contract: do not silently swallow the
    # resource's exception, but still finish best-effort cleanup first.
    with pytest.raises(_ExpectedFailure, match=f"{failing_resource} close"):
        workbook.close()

    assert primary.close_calls == 1
    assert cached.close_calls == 1
    assert workbook._wb is None
    assert workbook._cached_wb is None


def test_workbook_close_preserves_primary_error_when_both_resources_raise() -> None:
    workbook = object.__new__(api.MessyWorkbook)
    primary = _CloseTracker("primary", fail_on_close=True)
    cached = _CloseTracker("cached", fail_on_close=True)
    workbook._wb = primary  # type: ignore[assignment]
    workbook._cached_wb = cached  # type: ignore[assignment]

    with pytest.raises(_ExpectedFailure, match="primary close"):
        workbook.close()

    assert primary.close_calls == 1
    assert cached.close_calls == 1
    assert workbook._wb is None
    assert workbook._cached_wb is None


def test_workbook_close_preserves_session_error_before_later_resource_error() -> None:
    workbook = object.__new__(api.MessyWorkbook)
    session = _CloseTracker("fastexcel session", fail_on_close=True)
    primary = _CloseTracker("primary", fail_on_close=True)
    workbook._fastexcel_session = session  # type: ignore[assignment]
    workbook._wb = primary  # type: ignore[assignment]
    workbook._cached_wb = None

    with pytest.raises(_ExpectedFailure, match="fastexcel session close"):
        workbook.close()

    assert session.close_calls == 1
    assert primary.close_calls == 1
    assert workbook._fastexcel_session is None
    assert workbook._wb is None


def test_workbook_close_is_idempotent_for_both_resources() -> None:
    workbook = object.__new__(api.MessyWorkbook)
    primary = _CloseTracker("primary")
    cached = _CloseTracker("cached")
    workbook._wb = primary  # type: ignore[assignment]
    workbook._cached_wb = cached  # type: ignore[assignment]

    workbook.close()
    workbook.close()

    assert primary.close_calls == 1
    assert cached.close_calls == 1
    assert workbook._wb is None
    assert workbook._cached_wb is None


def test_workbook_lifecycle_exists_before_initialization_and_preserves_failure(
    monkeypatch: pytest.MonkeyPatch,
    sample_xlsx: Path,
) -> None:
    primary = ValueError("initialization failure")
    cleanup = _CloseTracker("initialization", fail_on_close=True)
    observed: dict[str, object] = {}

    def fail_initialization(workbook: api.MessyWorkbook) -> None:
        observed["closed"] = workbook._closed
        observed["token"] = workbook._active_operation_token
        observed["stream"] = workbook._active_stream
        workbook._wb = cleanup  # type: ignore[assignment]
        raise primary

    monkeypatch.setattr(api.MessyWorkbook, "_initialize_source", fail_initialization)

    with pytest.raises(ValueError) as captured:
        api.MessyWorkbook(sample_xlsx)

    assert captured.value is primary
    assert observed == {"closed": False, "token": None, "stream": None}
    assert cleanup.close_calls == 1
    assert primary.__dict__["backend_context"]["cleanup_failure"] == {"type": "_ExpectedFailure"}


def test_workbook_context_preserves_body_error_under_ordinary_cleanup() -> None:
    workbook = object.__new__(api.MessyWorkbook)
    cleanup = _CloseTracker("workbook", fail_on_close=True)
    body_error = ValueError("body failure")
    workbook._wb = cleanup  # type: ignore[assignment]
    workbook._cached_wb = None

    with pytest.raises(ValueError) as captured, workbook:
        raise body_error

    assert captured.value is body_error
    assert cleanup.close_calls == 1
    assert body_error.__dict__["backend_context"]["cleanup_failure"] == {"type": "_ExpectedFailure"}
    assert "workbook close" not in " ".join(getattr(body_error, "__notes__", ()))


def test_workbook_context_process_cleanup_wins_over_ordinary_body() -> None:
    workbook = object.__new__(api.MessyWorkbook)
    body_error = ValueError("body failure")
    cleanup_error = MemoryError("process cleanup")
    cleanup = _BaseCloseTracker(cleanup_error)
    workbook._fastexcel_session = cleanup  # type: ignore[assignment]
    workbook._wb = None
    workbook._cached_wb = None

    with pytest.raises(MemoryError) as captured, workbook:
        raise body_error

    assert captured.value is cleanup_error
    assert cleanup.close_calls == 1
    assert cleanup_error.__dict__["backend_context"]["operation_failure"] == {"type": "ValueError"}


def test_workbook_context_keeps_first_process_failure_and_attempts_cleanup() -> None:
    workbook = object.__new__(api.MessyWorkbook)
    body_error = KeyboardInterrupt("process body")
    cleanup_error = MemoryError("later process cleanup")
    session = _BaseCloseTracker(cleanup_error)
    workbook_resource = _CloseTracker("workbook")
    workbook._fastexcel_session = session  # type: ignore[assignment]
    workbook._wb = workbook_resource  # type: ignore[assignment]
    workbook._cached_wb = None

    with pytest.raises(KeyboardInterrupt) as captured, workbook:
        raise body_error

    assert captured.value is body_error
    assert session.close_calls == 1
    assert workbook_resource.close_calls == 1
    assert body_error.__dict__["backend_context"]["cleanup_failure"] == {"type": "MemoryError"}


def test_structure_analyzer_closes_workbook_when_detection_helper_raises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class AnalyzerWorkbook(_CloseTracker):
        sheetnames: ClassVar[list[str]] = ["Data"]

        def __getitem__(self, _sheet: str) -> object:
            return object()

    loaded_workbook = AnalyzerWorkbook("structure analysis")
    analyzer = StructureAnalyzer()

    def fail_detection(_worksheet: object) -> dict[str, int]:
        raise _ExpectedFailure("data-region detection")

    monkeypatch.setattr(
        structure_analyzer_module.openpyxl,
        "load_workbook",
        lambda *_args, **_kwargs: loaded_workbook,
    )
    monkeypatch.setattr(analyzer, "_detect_data_region", fail_detection)

    with pytest.raises(_ExpectedFailure, match="data-region detection"):
        analyzer.analyze(io.BytesIO(b"content supplied to fake loader"), "Data")

    assert loaded_workbook.close_calls == 1
    assert loaded_workbook.closed is True


def test_xlsx_openpyxl_parse_closes_workbook_when_worksheet_reading_raises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class ParseWorkbook(_CloseTracker):
        sheetnames: ClassVar[list[str]] = ["Data"]

        def __getitem__(self, _sheet: str) -> object:
            return object()

    loaded_workbook = ParseWorkbook("xlsx parse")
    handler = XLSXHandler()

    def fail_worksheet_read(_worksheet: object, _options: ParseOptions) -> list[list[Any]]:
        raise _ExpectedFailure("worksheet read")

    monkeypatch.setattr(
        xlsx_handler_module.openpyxl,
        "load_workbook",
        lambda *_args, **_kwargs: loaded_workbook,
    )
    monkeypatch.setattr(handler, "_read_worksheet", fail_worksheet_read)

    options = ParseOptions(merge_strategy="skip", ignore_hidden=True)
    with pytest.raises(_ExpectedFailure, match="worksheet read"):
        handler.parse(Path("book.xlsx"), "Data", options)

    assert loaded_workbook.close_calls == 1
    assert loaded_workbook.closed is True


def test_xlsx_sheet_name_fallback_closes_workbook_when_sheetnames_raise(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class SheetNameFailureWorkbook(_CloseTracker):
        @property
        def sheetnames(self) -> list[str]:
            raise _ExpectedFailure("sheetnames")

    loaded_workbook = SheetNameFailureWorkbook("xlsx")

    def fail_fastexcel(_source: object) -> None:
        raise RuntimeError("force openpyxl fallback")

    monkeypatch.setattr(xlsx_handler_module.fastexcel, "read_excel", fail_fastexcel)
    monkeypatch.setattr(
        xlsx_handler_module.openpyxl,
        "load_workbook",
        lambda *_args, **_kwargs: loaded_workbook,
    )

    with pytest.raises(FormatError, match="Cannot read sheet names"):
        XLSXHandler().get_sheet_names(Path("book.xlsx"))

    assert loaded_workbook.close_calls == 1
    assert loaded_workbook.closed is True


def test_xls_sheet_name_failure_closes_constructed_excel_file(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class SheetNameFailureExcelFile(_CloseTracker):
        @property
        def sheet_names(self) -> list[str]:
            raise _ExpectedFailure("sheet names")

    excel_file = SheetNameFailureExcelFile("xls")
    monkeypatch.setattr(
        xls_handler_module.pd,
        "ExcelFile",
        lambda *_args, **_kwargs: excel_file,
    )

    # XLS keeps its established compatibility fallback when sheet discovery
    # fails, while still releasing the successfully constructed ExcelFile.
    assert XLSHandler().get_sheet_names(Path("book.xls")) == ["Sheet1"]
    assert excel_file.close_calls == 1
    assert excel_file.closed is True


def test_xls_validation_attempts_close_and_reports_close_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    excel_file = _CloseTracker("xls validation", fail_on_close=True)
    monkeypatch.setattr(
        xls_handler_module.pd,
        "ExcelFile",
        lambda *_args, **_kwargs: excel_file,
    )

    is_valid, error = XLSHandler().validate(Path("book.xls"))

    assert is_valid is False
    assert error == "xls validation close"
    assert excel_file.close_calls == 1
    assert excel_file.closed is True


def _track_csv_string_streams(monkeypatch: pytest.MonkeyPatch) -> None:
    _TrackedStringIO.instances = []
    monkeypatch.setattr(
        csv_handler_module,
        "io",
        SimpleNamespace(StringIO=_TrackedStringIO),
    )


def test_csv_parse_closes_library_created_text_stream_on_success(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _track_csv_string_streams(monkeypatch)

    result = CSVHandler().parse(
        io.BytesIO(b"Name,Value\nA,1\n"),
        None,
        ParseOptions(auto_detect_header=False),
    )

    assert not result.empty
    assert _TrackedStringIO.instances
    assert all(stream.closed for stream in _TrackedStringIO.instances)


def _count_descriptors_resolving_to(path: Path) -> int:
    target = path.resolve()
    count = 0
    for descriptor in Path("/proc/self/fd").iterdir():
        try:
            if descriptor.resolve(strict=True) == target:
                count += 1
        except OSError:
            # Descriptors may disappear while /proc is being traversed.
            continue
    return count


def test_repeated_read_excel_does_not_grow_descriptors_for_target(
    sample_xlsx: Path,
) -> None:
    proc_descriptors = Path("/proc/self/fd")
    if not proc_descriptors.is_dir():
        pytest.skip("target-specific descriptor inspection requires /proc/self/fd")

    # Warm lazy imports and the selected parser backend before measuring.
    api.read_excel(sample_xlsx)
    before = _count_descriptors_resolving_to(sample_xlsx)

    for _ in range(10):
        api.read_excel(sample_xlsx)

    after = _count_descriptors_resolving_to(sample_xlsx)
    assert after == before


def test_csv_parse_closes_library_created_text_stream_on_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _track_csv_string_streams(monkeypatch)

    def fail_read_csv(*_args: Any, **_kwargs: Any) -> None:
        raise ValueError("injected CSV parse failure")

    monkeypatch.setattr(csv_handler_module.pd, "read_csv", fail_read_csv)

    with pytest.raises(FormatError, match="Cannot parse CSV file"):
        CSVHandler().parse(
            io.BytesIO(b"Name,Value\nA,1\n"),
            None,
            ParseOptions(auto_detect_header=False),
        )

    assert _TrackedStringIO.instances
    assert all(stream.closed for stream in _TrackedStringIO.instances)


def test_csv_metadata_detection_closes_library_created_text_stream(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _track_csv_string_streams(monkeypatch)

    MetadataRowDetector().detect_skip_rows_from_text(
        "Report:,\nName,Value\nA,1\nB,2\n",
        ",",
    )

    assert _TrackedStringIO.instances
    assert all(stream.closed for stream in _TrackedStringIO.instances)


def test_csv_metadata_failure_closes_library_created_text_stream(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _track_csv_string_streams(monkeypatch)

    def fail_read_csv(*_args: Any, **_kwargs: Any) -> None:
        raise pd.errors.ParserError("injected metadata failure")

    monkeypatch.setattr(csv_handler_module.pd, "read_csv", fail_read_csv)

    assert MetadataRowDetector().detect_skip_rows_from_text("A,B\n1,2\n", ",") == 0
    assert _TrackedStringIO.instances
    assert all(stream.closed for stream in _TrackedStringIO.instances)
