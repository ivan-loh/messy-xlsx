"""Shared deterministic lifecycle for one-shot parser result streams."""

from __future__ import annotations

import weakref
from collections.abc import Callable, Iterator, Sequence
from types import TracebackType
from typing import Any, Generic, Self, TypeVar

import pandas as pd
import pyarrow as pa

from messy_xlsx._fallback_signals import (
    _attach_cleanup_failure,
    _attach_operation_failure,
    _contains_process_failure,
    _exception_traceback,
    _raise_with_traceback,
    _safe_add_note,
    _type_name,
)
from messy_xlsx.models import SheetResult

T = TypeVar("T")
_OwnerT = TypeVar("_OwnerT")
_Cleanup = tuple[str, Callable[[], object]]

__all__ = ["BatchStream", "DataFrameChunkStream", "SheetStream"]


def _run_cleanups(
    cleanups: Sequence[_Cleanup],
    *,
    primary_error: BaseException | None = None,
    primary_traceback: TracebackType | None = None,
) -> bool:
    """Attempt every cleanup and apply the shared process-failure policy."""
    winner = primary_error
    winner_traceback = primary_traceback
    process_cleanup_won = primary_error is not None and _contains_process_failure(primary_error)
    cleanup_failed = False

    for label, cleanup in cleanups:
        try:
            cleanup()
        except BaseException as cleanup_error:
            cleanup_failed = True
            cleanup_traceback = _exception_traceback(cleanup_error)
            if winner is None:
                winner = cleanup_error
                winner_traceback = cleanup_traceback
                process_cleanup_won = _contains_process_failure(cleanup_error)
                _attach_cleanup_failure(cleanup_error, cleanup_error)
                _safe_add_note(
                    cleanup_error,
                    f"{label} failed: {_type_name(cleanup_error)}",
                )
                continue

            cleanup_is_process_failure = _contains_process_failure(
                cleanup_error,
                exclude=winner,
            )
            if cleanup_is_process_failure and not process_cleanup_won:
                _attach_operation_failure(cleanup_error, winner)
                _attach_cleanup_failure(cleanup_error, cleanup_error)
                _safe_add_note(
                    cleanup_error,
                    f"operation also failed: {_type_name(winner)}",
                )
                winner = cleanup_error
                winner_traceback = cleanup_traceback
                process_cleanup_won = True
                continue

            _attach_cleanup_failure(winner, cleanup_error)
            _safe_add_note(
                winner,
                f"{label} also failed: {_type_name(cleanup_error)}",
            )

    if winner is None or winner is primary_error:
        return cleanup_failed
    _raise_with_traceback(winner, winner_traceback)
    return cleanup_failed


def _close_if_present(owner: _OwnerT) -> _OwnerT | None:
    """Close an owner and return typed ``None`` without a trace-event gap."""
    close = getattr(owner, "close", None)
    return (close(), None)[1] if callable(close) else None


class _StreamCleanupState:
    """Own stream resources without retaining the public stream object."""

    def __init__(
        self,
        source: Iterator[Any],
        close_callback: Callable[[], object],
    ) -> None:
        self.source: Iterator[Any] | None = source
        self.close_callback: Callable[[], object] | None = close_callback

    @property
    def done(self) -> bool:
        return self.source is None and self.close_callback is None

    def close_source(self) -> None:
        source = self.source
        if source is None:
            return
        defer_retry = getattr(source, "_defer_process_retry_to_owner", None)
        if callable(defer_retry) and defer_retry():
            return
        try:
            self.source = _close_if_present(source)
        except BaseException as error:
            if not _contains_process_failure(error):
                self.source = None
            raise

    def release_stream(self) -> None:
        close_callback = self.close_callback
        if close_callback is None:
            return
        try:
            self.close_callback = close_callback if close_callback() is False else None
        except BaseException as error:
            if not _contains_process_failure(error):
                self.close_callback = None
            raise

    def close(
        self,
        *,
        primary_error: BaseException | None = None,
        primary_traceback: TracebackType | None = None,
    ) -> None:
        cleanups: list[_Cleanup] = []
        if self.source is not None:
            cleanups.append(("stream source cleanup", self.close_source))
        if self.close_callback is not None:
            cleanups.append(("stream release callback", self.release_stream))
        _run_cleanups(
            cleanups,
            primary_error=primary_error,
            primary_traceback=primary_traceback,
        )


def _finalize_stream_cleanup(state: _StreamCleanupState) -> None:
    """Best-effort cleanup for a stream lost at a public return boundary."""
    try:
        state.close()
    except BaseException:
        # Finalizers cannot propagate. Retryable state remains workbook-owned
        # until the next operation or parent close attempts cleanup again.
        pass


class _ResultStream(Generic[T], Iterator[T]):
    """One-shot result iterator with deterministic non-masking cleanup."""

    def __init__(self, source: Iterator[T], close_callback: Callable[[], object]) -> None:
        self._cleanup_state = _StreamCleanupState(source, close_callback)
        self._abandonment_finalizer = weakref.finalize(
            self,
            _finalize_stream_cleanup,
            self._cleanup_state,
        )
        self._closed = False
        self._owner_invalidated = False

    @property
    def _source(self) -> Iterator[T] | None:
        return self._cleanup_state.source

    @_source.setter
    def _source(self, source: Iterator[T] | None) -> None:
        self._cleanup_state.source = source

    @property
    def _close_callback(self) -> Callable[[], object] | None:
        return self._cleanup_state.close_callback

    @_close_callback.setter
    def _close_callback(self, callback: Callable[[], object] | None) -> None:
        self._cleanup_state.close_callback = callback

    def __iter__(self) -> Self:
        return self

    def __next__(self) -> T:
        if self._owner_invalidated:
            raise RuntimeError("MessyWorkbook is closed")
        source = self._source
        if self._closed or source is None:
            raise StopIteration
        try:
            return next(source)
        except StopIteration:
            self.close()
            raise
        except BaseException as error:
            self._close(primary_error=error, primary_traceback=_exception_traceback(error))
            raise

    def close(self) -> None:
        self._close()

    def _close(
        self,
        *,
        primary_error: BaseException | None = None,
        primary_traceback: TracebackType | None = None,
    ) -> None:
        if self._closed and self._cleanup_state.done:
            return
        self._closed = True
        self._cleanup_state.close(
            primary_error=primary_error,
            primary_traceback=primary_traceback,
        )
        if self._cleanup_state.done:
            self._abandonment_finalizer.detach()

    def _close_source(self) -> None:
        self._cleanup_state.close_source()

    def _release_stream(self) -> None:
        self._cleanup_state.release_stream()

    def invalidate_from_owner(self) -> None:
        try:
            self._owner_invalidated = True
            self.close()
        finally:
            self._owner_invalidated = True

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        del exc_type
        if exc_value is None:
            self.close()
            return
        self._close(primary_error=exc_value, primary_traceback=traceback)


class BatchStream(_ResultStream[pa.RecordBatch]):
    """One-shot stream of Arrow batches with a stable, persistent schema."""

    def __init__(
        self,
        source: Iterator[pa.RecordBatch],
        schema: pa.Schema,
        close_callback: Callable[[], object],
    ) -> None:
        super().__init__(source, close_callback)
        self._schema = schema
        self._display_names: tuple[object, ...] = tuple(schema.names)

    @property
    def schema(self) -> pa.Schema:
        return self._schema


class DataFrameChunkStream(_ResultStream[pd.DataFrame]):
    """One-shot stream of pandas chunks with deterministic cleanup."""


class SheetStream(_ResultStream[SheetResult]):
    """One-shot stream of ordered sheet results with deterministic cleanup."""
