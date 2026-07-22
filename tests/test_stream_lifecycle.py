"""Contracts for deterministic one-shot parser stream ownership."""

from __future__ import annotations

import gc
import io
import weakref
from collections.abc import Iterator
from contextlib import ExitStack
from typing import Any

import pandas as pd
import pyarrow as pa
import pytest

from messy_xlsx import MessyWorkbook
from messy_xlsx.parsing.handler_registry import HandlerRegistry
from messy_xlsx.parsing.streams import BatchStream, DataFrameChunkStream, SheetStream


class _ClosableIterator(Iterator[Any]):
    def __init__(
        self,
        values: list[Any] | None = None,
        *,
        iteration_error: BaseException | None = None,
        close_error: BaseException | None = None,
        events: list[str] | None = None,
    ) -> None:
        self._values = iter(values or [])
        self._iteration_error = iteration_error
        self._close_error = close_error
        self._events = events
        self.read_calls = 0
        self.close_calls = 0

    def __iter__(self) -> _ClosableIterator:
        return self

    def __next__(self) -> Any:
        self.read_calls += 1
        if self._iteration_error is not None:
            error = self._iteration_error
            self._iteration_error = None
            raise error
        return next(self._values)

    def close(self) -> None:
        self.close_calls += 1
        if self._events is not None:
            self._events.append("source")
        if self._close_error is not None:
            raise self._close_error


class _InvalidatedChild:
    def __init__(
        self,
        events: list[str],
        error: BaseException | None = None,
    ) -> None:
        self._events = events
        self._error = error
        self.calls = 0

    def invalidate_from_owner(self) -> None:
        self.calls += 1
        self._events.append("child")
        if self._error is not None:
            raise self._error


class _EventResource:
    def __init__(
        self,
        name: str,
        events: list[str],
        error: BaseException | None = None,
    ) -> None:
        self.name = name
        self._events = events
        self._error = error
        self.calls = 0

    def close(self) -> None:
        self.calls += 1
        self._events.append(self.name)
        if self._error is not None:
            raise self._error


class _SecondCloseFailureReader:
    def __init__(self) -> None:
        self._batches = iter([_batch()])
        self.close_calls = 0

    def read_next_batch(self) -> pa.RecordBatch | None:
        return next(self._batches, None)

    def close(self) -> None:
        self.close_calls += 1
        if self.close_calls > 1:
            raise AssertionError("reader closed more than once")


def _batch() -> pa.RecordBatch:
    return pa.record_batch([[1, 2]], names=["value"])


def _assert_sanitized_cleanup(error: BaseException, cleanup_type: str) -> None:
    assert error.__dict__["backend_context"]["cleanup_failure"] == {"type": cleanup_type}
    assert "secret" not in " ".join(getattr(error, "__notes__", ()))


def test_batch_stream_is_one_shot_and_close_is_idempotent() -> None:
    released: list[bool] = []
    source = _ClosableIterator([_batch()])
    stream = BatchStream(source, _batch().schema, lambda: released.append(True))

    assert iter(stream) is stream
    assert next(stream).equals(_batch())
    with pytest.raises(StopIteration):
        next(stream)
    stream.close()
    stream.close()

    assert source.close_calls == 1
    assert released == [True]


def test_stream_construction_does_not_read_and_schema_is_persistent_read_only() -> None:
    schema = pa.schema([("value", pa.int64())])
    source = _ClosableIterator()
    stream = BatchStream(source, schema, lambda: None)

    assert source.read_calls == 0
    assert stream.schema is schema
    with pytest.raises(AttributeError):
        stream.schema = pa.schema([])  # type: ignore[misc]

    stream.close()
    assert stream.schema is schema


def test_empty_exhaustion_closes_and_keeps_schema() -> None:
    schema = pa.schema([("value", pa.int64())])
    source = _ClosableIterator()
    released: list[bool] = []
    stream = BatchStream(source, schema, lambda: released.append(True))

    with pytest.raises(StopIteration):
        next(stream)

    assert source.close_calls == 1
    assert released == [True]
    assert stream.schema is schema


def test_backend_error_closes_once_and_preserves_the_exact_error() -> None:
    primary = ValueError("iteration failed")
    source = _ClosableIterator(iteration_error=primary)
    released: list[bool] = []
    stream = BatchStream(source, pa.schema([]), lambda: released.append(True))

    with pytest.raises(ValueError) as captured:
        next(stream)

    assert captured.value is primary
    assert source.close_calls == 1
    assert released == [True]
    assert stream.schema == pa.schema([])
    with pytest.raises(StopIteration):
        next(stream)


@pytest.mark.parametrize(
    ("stream_type", "item"),
    [
        (DataFrameChunkStream, pd.DataFrame({"value": [1]})),
        (SheetStream, object()),
    ],
)
def test_typed_wrappers_share_the_one_shot_lifecycle(
    stream_type: type[DataFrameChunkStream] | type[SheetStream],
    item: object,
) -> None:
    released: list[bool] = []
    stream = stream_type(iter([item]), lambda: released.append(True))

    assert next(stream) is item
    with pytest.raises(StopIteration):
        next(stream)
    assert released == [True]


def test_explicit_close_stops_but_owner_invalidation_remains_distinct() -> None:
    explicit = BatchStream(_ClosableIterator([_batch()]), _batch().schema, lambda: None)
    explicit.close()
    with pytest.raises(StopIteration):
        next(explicit)

    invalidated = BatchStream(
        _ClosableIterator([_batch()]),
        _batch().schema,
        lambda: None,
    )
    invalidated.invalidate_from_owner()
    with pytest.raises(RuntimeError, match="MessyWorkbook is closed"):
        next(invalidated)


def test_owner_invalidation_state_wins_even_when_cleanup_fails() -> None:
    cleanup_error = OSError("secret cleanup detail")
    stream = BatchStream(
        _ClosableIterator(close_error=cleanup_error),
        pa.schema([]),
        lambda: None,
    )

    with pytest.raises(OSError) as captured:
        stream.invalidate_from_owner()
    assert captured.value is cleanup_error

    with pytest.raises(RuntimeError, match="MessyWorkbook is closed"):
        next(stream)


def test_source_and_release_cleanup_are_both_attempted_and_only_once() -> None:
    source_error = OSError("secret source close")
    release_error = RuntimeError("secret release close")
    events: list[str] = []
    source = _ClosableIterator(close_error=source_error, events=events)

    def release() -> None:
        events.append("release")
        raise release_error

    stream = BatchStream(source, pa.schema([]), release)

    with pytest.raises(OSError) as captured:
        stream.close()
    stream.close()

    assert captured.value is source_error
    assert events == ["source", "release"]
    _assert_sanitized_cleanup(source_error, "RuntimeError")
    assert stream._source is None
    assert stream._close_callback is None


def test_callback_process_cleanup_wins_over_earlier_ordinary_source_cleanup() -> None:
    source_error = OSError("secret source close")
    callback_error = MemoryError("secret callback close")
    events: list[str] = []
    source = _ClosableIterator(close_error=source_error, events=events)

    def release() -> None:
        events.append("release")
        raise callback_error

    stream = BatchStream(source, pa.schema([]), release)

    with pytest.raises(MemoryError) as captured:
        stream.close()

    assert captured.value is callback_error
    assert events == ["source", "release"]
    assert callback_error.__dict__["backend_context"]["operation_failure"] == {"type": "OSError"}


def test_first_process_cleanup_wins_while_later_callback_is_attempted() -> None:
    source_error = MemoryError("secret source close")
    callback_error = SystemExit("secret callback close")
    events: list[str] = []
    source = _ClosableIterator(close_error=source_error, events=events)

    def release() -> None:
        events.append("release")
        raise callback_error

    stream = BatchStream(source, pa.schema([]), release)

    with pytest.raises(MemoryError) as captured:
        stream.close()

    assert captured.value is source_error
    assert events == ["source", "release"]
    _assert_sanitized_cleanup(source_error, "SystemExit")


def test_iteration_error_keeps_precedence_over_ordinary_cleanup() -> None:
    primary = ValueError("primary iteration")
    cleanup_error = OSError("secret cleanup detail")
    source = _ClosableIterator(iteration_error=primary, close_error=cleanup_error)
    stream = BatchStream(source, pa.schema([]), lambda: None)

    with pytest.raises(ValueError) as captured:
        next(stream)

    assert captured.value is primary
    _assert_sanitized_cleanup(primary, "OSError")


def test_process_cleanup_wins_over_iteration_error() -> None:
    primary = ValueError("primary iteration")
    cleanup_error = BaseExceptionGroup(
        "cleanup",
        [OSError("secret ordinary"), KeyboardInterrupt("secret process")],
    )
    source = _ClosableIterator(iteration_error=primary, close_error=cleanup_error)
    stream = BatchStream(source, pa.schema([]), lambda: None)

    with pytest.raises(BaseExceptionGroup) as captured:
        next(stream)

    assert captured.value is cleanup_error
    assert cleanup_error.__dict__["backend_context"]["operation_failure"] == {"type": "ValueError"}
    assert "secret" not in " ".join(getattr(cleanup_error, "__notes__", ()))


def test_process_iteration_error_remains_the_first_process_winner() -> None:
    primary = KeyboardInterrupt("primary process")
    cleanup_error = MemoryError("later process cleanup")
    events: list[str] = []
    source = _ClosableIterator(
        iteration_error=primary,
        close_error=cleanup_error,
        events=events,
    )

    def release() -> None:
        events.append("release")

    stream = BatchStream(source, pa.schema([]), release)

    with pytest.raises(KeyboardInterrupt) as captured:
        next(stream)

    assert captured.value is primary
    assert events == ["source", "release"]
    _assert_sanitized_cleanup(primary, "MemoryError")


def test_context_body_error_survives_ordinary_cleanup() -> None:
    body_error = ValueError("body")
    cleanup_error = OSError("secret cleanup")
    stream = BatchStream(
        _ClosableIterator(close_error=cleanup_error),
        pa.schema([]),
        lambda: None,
    )

    with pytest.raises(ValueError) as captured, stream:
        raise body_error

    assert captured.value is body_error
    _assert_sanitized_cleanup(body_error, "OSError")


def test_context_process_cleanup_wins_over_body_error() -> None:
    body_error = ValueError("body")
    cleanup_error = MemoryError("secret cleanup")
    stream = BatchStream(
        _ClosableIterator(close_error=cleanup_error),
        pa.schema([]),
        lambda: None,
    )

    with pytest.raises(MemoryError) as captured, stream:
        raise body_error

    assert captured.value is cleanup_error
    assert cleanup_error.__dict__["backend_context"]["operation_failure"] == {"type": "ValueError"}


def test_foreign_and_stale_tokens_cannot_release_a_live_operation(sample_xlsx: Any) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        first = workbook._begin_operation()
        workbook._end_operation(object())
        assert workbook._active_operation_token is first

        workbook._end_operation(first)
        second = workbook._begin_operation()
        workbook._end_operation(first)
        assert workbook._active_operation_token is second
        workbook._end_operation(second)


def test_workbook_rejects_closed_and_second_active_operations(sample_xlsx: Any) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    token = workbook._begin_operation()
    with pytest.raises(RuntimeError, match="active parse or stream"):
        workbook._begin_operation()
    workbook._end_operation(token)
    workbook.close()
    with pytest.raises(RuntimeError, match="MessyWorkbook is closed"):
        workbook._begin_operation()


@pytest.mark.parametrize("termination", ["exhaust", "error", "explicit"])
def test_stream_termination_releases_matching_workbook_token_once(
    sample_xlsx: Any,
    monkeypatch: pytest.MonkeyPatch,
    termination: str,
) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    real_end = workbook._end_operation
    released: list[object] = []

    def record_end(token: object) -> None:
        released.append(token)
        real_end(token)

    monkeypatch.setattr(workbook, "_end_operation", record_end)
    iteration_error = ValueError("backend") if termination == "error" else None
    source = _ClosableIterator(
        [] if termination != "explicit" else [_batch()],
        iteration_error=iteration_error,
    )
    with workbook._stream_operation() as lease:
        stream = BatchStream(source, pa.schema([]), lease.release)
        lease.bind(stream)

    if termination == "error":
        with pytest.raises(ValueError, match="backend"):
            next(stream)
    elif termination == "exhaust":
        with pytest.raises(StopIteration):
            next(stream)
    else:
        stream.close()
        stream.close()

    assert len(released) == 1
    assert workbook._active_operation_token is None
    assert workbook._active_stream is None
    workbook.close()


def test_existing_materialized_calls_reject_an_active_operation(sample_xlsx: Any) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        token = workbook._begin_operation()
        try:
            with pytest.raises(RuntimeError, match="active parse or stream"):
                workbook.to_dataframe()
            with pytest.raises(RuntimeError, match="active parse or stream"):
                workbook._parse_sheet("Data")
            with pytest.raises(RuntimeError, match="active parse or stream"):
                workbook.to_dataframes()
        finally:
            workbook._end_operation(token)


def test_reentrant_custom_parse_error_is_not_swallowed_by_to_dataframes(
    sample_xlsx: Any,
) -> None:
    class ReentrantRegistry(HandlerRegistry):
        workbook: MessyWorkbook

        def parse(self, *_args: Any, **_kwargs: Any) -> pd.DataFrame:
            return self.workbook.to_dataframe()

    registry = ReentrantRegistry()
    with MessyWorkbook(sample_xlsx, registry=registry) as workbook:
        registry.workbook = workbook
        with pytest.raises(RuntimeError, match="active parse or stream"):
            workbook.to_dataframes(include_errors=True)


def test_wrapped_reentrant_custom_parse_cause_is_not_swallowed_by_to_dataframes(
    sample_xlsx: Any,
) -> None:
    class WrappedReentrantRegistry(HandlerRegistry):
        workbook: MessyWorkbook
        wrapped: ValueError | None = None

        def parse(self, *_args: Any, **_kwargs: Any) -> pd.DataFrame:
            try:
                return self.workbook.to_dataframe()
            except RuntimeError as error:
                self.wrapped = ValueError("wrapped reentrant parse")
                raise self.wrapped from error

    registry = WrappedReentrantRegistry()
    with MessyWorkbook(sample_xlsx, registry=registry) as workbook:
        registry.workbook = workbook
        with pytest.raises(ValueError, match="wrapped reentrant parse") as captured:
            workbook.to_dataframes(include_errors=True)

    assert captured.value is registry.wrapped
    assert isinstance(captured.value.__cause__, RuntimeError)


def test_grouped_reentrant_custom_parse_error_is_not_swallowed_by_to_dataframes(
    sample_xlsx: Any,
) -> None:
    class GroupedReentrantRegistry(HandlerRegistry):
        workbook: MessyWorkbook
        wrapped: ExceptionGroup | None = None

        def parse(self, *_args: Any, **_kwargs: Any) -> pd.DataFrame:
            try:
                return self.workbook.to_dataframe()
            except RuntimeError as error:
                self.wrapped = ExceptionGroup(
                    "grouped reentrant parse",
                    [ValueError("ordinary"), error],
                )
                raise self.wrapped from None

    registry = GroupedReentrantRegistry()
    with MessyWorkbook(sample_xlsx, registry=registry) as workbook:
        registry.workbook = workbook
        with pytest.raises(ExceptionGroup, match="grouped reentrant parse") as captured:
            workbook.to_dataframes(include_errors=True)

    assert captured.value is registry.wrapped


def test_process_failure_is_not_swallowed_by_to_dataframes(sample_xlsx: Any) -> None:
    failure = MemoryError("parse process failure")

    class ProcessFailureRegistry(HandlerRegistry):
        def parse(self, *_args: Any, **_kwargs: Any) -> pd.DataFrame:
            raise failure

    with (
        MessyWorkbook(sample_xlsx, registry=ProcessFailureRegistry()) as workbook,
        pytest.raises(MemoryError) as captured,
    ):
        workbook.to_dataframes(include_errors=True)

    assert captured.value is failure


@pytest.mark.parametrize(
    "failure",
    [PermissionError("sheet permission"), FileNotFoundError("sheet missing")],
)
@pytest.mark.parametrize("include_errors", [False, True])
def test_ordinary_blocked_sheet_failure_remains_aggregated_for_legacy_callers(
    sample_xlsx: Any,
    failure: Exception,
    include_errors: bool,
) -> None:
    class OrdinaryFailureRegistry(HandlerRegistry):
        def parse(self, *_args: Any, **_kwargs: Any) -> pd.DataFrame:
            raise failure

    with MessyWorkbook(sample_xlsx, registry=OrdinaryFailureRegistry()) as workbook:
        result = workbook.to_dataframes(include_errors=include_errors)

    if not include_errors:
        assert result == {}
        return
    frames, errors = result
    assert frames == {}
    assert len(errors) == 1
    assert errors[0].sheet_name == "Data"
    assert errors[0].error_type == type(failure).__name__


def test_nested_process_failure_tree_is_not_aggregated_by_to_dataframes(
    sample_xlsx: Any,
) -> None:
    failure = ExceptionGroup(
        "outer",
        [ValueError("ordinary"), ExceptionGroup("inner", [MemoryError("process")])],
    )

    class ProcessTreeRegistry(HandlerRegistry):
        def parse(self, *_args: Any, **_kwargs: Any) -> pd.DataFrame:
            raise failure

    with (
        MessyWorkbook(sample_xlsx, registry=ProcessTreeRegistry()) as workbook,
        pytest.raises(ExceptionGroup) as captured,
    ):
        workbook.to_dataframes(include_errors=True)

    assert captured.value is failure


def test_failed_materialized_parse_releases_reservation_for_the_next_parse(
    sample_xlsx: Any,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        original = workbook._parse_sheet_unreserved
        primary = ValueError("parse failed")
        calls = 0

        def fail_once(sheet: str, config: Any = None) -> pd.DataFrame:
            nonlocal calls
            calls += 1
            if calls == 1:
                raise primary
            return original(sheet, config)

        monkeypatch.setattr(workbook, "_parse_sheet_unreserved", fail_once)
        with pytest.raises(ValueError) as captured:
            workbook._parse_sheet("Data")

        assert captured.value is primary
        assert workbook._active_operation_token is None
        assert not workbook._parse_sheet("Data").empty


def test_stale_lease_callback_and_registration_cannot_replace_a_new_child(
    sample_xlsx: Any,
) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    with workbook._stream_operation() as first_lease:
        first = BatchStream(_ClosableIterator(), pa.schema([]), first_lease.release)
        first_lease.bind(first)
    stale_token = first_lease._token
    first.close()

    with workbook._stream_operation() as second_lease:
        second = BatchStream(_ClosableIterator(), pa.schema([]), second_lease.release)
        second_lease.bind(second)

    first_lease.release()
    with pytest.raises(RuntimeError, match="reservation is no longer active"):
        workbook._register_stream(stale_token, first)

    assert workbook._active_operation_token is second_lease._token
    assert workbook._active_stream is second
    second.close()
    workbook.close()


def test_failed_stale_bind_closes_created_child_without_releasing_new_operation(
    sample_xlsx: Any,
) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    lease = workbook._stream_operation()
    source = _ClosableIterator([_batch()])
    stream = BatchStream(source, _batch().schema, lease.release)
    workbook._end_operation(lease._token)
    new_token = workbook._begin_operation()
    new_child = object()
    workbook._register_stream(new_token, new_child)

    with pytest.raises(RuntimeError, match="reservation is no longer active"), lease:
        lease.bind(stream)

    assert source.close_calls == 1
    assert workbook._active_operation_token is new_token
    assert workbook._active_stream is new_child
    workbook._end_operation(new_token)
    workbook.close()


def test_failed_stale_bind_closes_unstarted_adapter_and_distinct_owned_reader_once(
    sample_xlsx: Any,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    real_end = workbook._end_operation
    releases: list[object] = []

    def record_end(token: object) -> None:
        releases.append(token)
        real_end(token)

    monkeypatch.setattr(workbook, "_end_operation", record_end)
    lease = workbook._stream_operation()
    reader = _EventResource("reader", [])
    owned_reader = lease.own(reader)

    def reader_batches() -> Iterator[pa.RecordBatch]:
        try:
            yield _batch()
        finally:
            owned_reader.close()

    stream = BatchStream(reader_batches(), _batch().schema, lease.release)
    workbook._end_operation(lease._token)
    new_token = workbook._begin_operation()
    new_child = object()
    workbook._register_stream(new_token, new_child)
    releases.clear()

    with pytest.raises(RuntimeError, match="reservation is no longer active"), lease:
        lease.bind(stream)

    assert reader.calls == 1
    assert releases == [lease._token]
    assert workbook._active_operation_token is new_token
    assert workbook._active_stream is new_child
    workbook._end_operation(new_token)
    workbook.close()


def test_explicit_close_before_first_adapter_read_closes_owned_reader_once(
    sample_xlsx: Any,
) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    reader = _EventResource("reader", [])

    with workbook._stream_operation() as lease:
        owned_reader = lease.own(reader)

        def reader_batches() -> Iterator[pa.RecordBatch]:
            try:
                yield _batch()
            finally:
                owned_reader.close()

        stream = BatchStream(reader_batches(), _batch().schema, lease.release)
        lease.bind(stream)

    stream.close()
    stream.close()

    assert reader.calls == 1
    assert workbook._active_operation_token is None
    assert workbook._active_stream is None
    workbook.close()


@pytest.mark.parametrize("termination", ["exhaust", "early"])
def test_started_adapter_and_lease_share_one_underlying_reader_close(
    sample_xlsx: Any,
    termination: str,
) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    reader = _SecondCloseFailureReader()

    with workbook._stream_operation() as lease:
        owned_reader = lease.own(reader)

        def reader_batches() -> Iterator[pa.RecordBatch]:
            try:
                while True:
                    batch = owned_reader.read_next_batch()
                    if batch is None:
                        return
                    yield batch
            finally:
                owned_reader.close()

        stream = BatchStream(reader_batches(), _batch().schema, lease.release)
        lease.bind(stream)

    assert next(stream).equals(_batch())
    if termination == "exhaust":
        with pytest.raises(StopIteration):
            next(stream)
    else:
        stream.close()

    assert reader.close_calls == 1
    assert workbook._active_operation_token is None
    workbook.close()


def test_parent_close_invalidates_first_and_attempts_every_resource() -> None:
    workbook = object.__new__(MessyWorkbook)
    events: list[str] = []
    child_error = OSError("secret child close")
    process_error = MemoryError("secret session close")
    child = _InvalidatedChild(events, child_error)
    session = _EventResource("session", events, process_error)
    primary = _EventResource("workbook", events)
    cached = _EventResource("cached", events)
    workbook_source = _EventResource("workbook-source", events)
    cached_source = _EventResource("cached-source", events)
    source_handle = _EventResource("source-handle", events)
    workbook._closed = False
    workbook._active_operation_token = object()
    workbook._active_stream = child
    workbook._fastexcel_session = session
    workbook._manifest_reader = object()
    workbook._wb = primary
    workbook._cached_wb = cached
    workbook._wb_source = workbook_source
    workbook._cached_wb_source = cached_source
    workbook._source_handle = source_handle

    with pytest.raises(MemoryError) as captured:
        workbook.close()

    assert captured.value is process_error
    assert events == [
        "child",
        "session",
        "workbook",
        "cached",
        "workbook-source",
        "cached-source",
        "source-handle",
    ]
    assert process_error.__dict__["backend_context"]["operation_failure"] == {"type": "OSError"}
    assert workbook._active_operation_token is None
    assert workbook._active_stream is None

    workbook.close()
    assert child.calls == 1
    assert session.calls == 1
    assert source_handle.calls == 1


def test_object_new_workbook_close_tolerates_all_new_fields_being_absent() -> None:
    workbook = object.__new__(MessyWorkbook)
    workbook.close()
    workbook.close()


def test_failed_stream_construction_closes_partial_and_restores_caller_cursor(
    sample_xlsx: Any,
) -> None:
    source = io.BytesIO(sample_xlsx.read_bytes())
    source.seek(7)
    workbook = MessyWorkbook(source, filename=sample_xlsx.name)
    entry = source.tell()
    stack = ExitStack()

    with (
        pytest.raises(ValueError, match="construction"),
        workbook._stream_operation() as lease,
    ):
        lease.own(stack)
        borrowed = stack.enter_context(workbook._source_handle.open_binary())
        assert borrowed is source
        assert source.tell() == 0
        raise ValueError("construction")

    assert workbook._active_operation_token is None
    assert source.tell() == entry
    assert source.closed is False
    assert not workbook.to_dataframe().empty
    workbook.close()


def test_workbook_and_stream_are_collectible_after_close(sample_xlsx: Any) -> None:
    workbook = MessyWorkbook(sample_xlsx)
    with workbook._stream_operation() as lease:
        stream = BatchStream(_ClosableIterator(), pa.schema([]), lease.release)
        lease.bind(stream)

    workbook_ref = weakref.ref(workbook)
    stream_ref = weakref.ref(stream)
    stream.close()
    workbook.close()
    del stream
    del lease
    del workbook
    gc.collect()

    assert stream_ref() is None
    assert workbook_ref() is None


def test_stream_releases_source_and_callback_payload_references() -> None:
    class Callback:
        def __call__(self) -> None:
            return None

    source = _ClosableIterator()
    callback = Callback()
    source_ref = weakref.ref(source)
    callback_ref = weakref.ref(callback)
    stream = BatchStream(source, pa.schema([]), callback)

    stream.close()
    del source
    del callback
    gc.collect()

    assert source_ref() is None
    assert callback_ref() is None


def test_task9_keeps_stream_types_off_the_package_root() -> None:
    import messy_xlsx
    import messy_xlsx.parsing.streams as streams

    assert streams.__all__ == ["BatchStream", "DataFrameChunkStream", "SheetStream"]
    assert "BatchStream" not in messy_xlsx.__all__
    assert "DataFrameChunkStream" not in messy_xlsx.__all__
    assert "SheetStream" not in messy_xlsx.__all__
