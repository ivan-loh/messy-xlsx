"""Shared deterministic lifecycle for one-shot parser result streams."""

from __future__ import annotations

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


class _ResultStream(Generic[T], Iterator[T]):
    """One-shot result iterator with deterministic non-masking cleanup."""

    def __init__(self, source: Iterator[T], close_callback: Callable[[], object]) -> None:
        self._source: Iterator[T] | None = source
        self._close_callback: Callable[[], object] | None = close_callback
        self._closed = False
        self._owner_invalidated = False

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
        if self._closed and self._source is None and self._close_callback is None:
            return
        self._closed = True

        cleanups: list[_Cleanup] = []
        if self._source is not None:
            cleanups.append(("stream source cleanup", self._close_source))
        if self._close_callback is not None:
            cleanups.append(("stream release callback", self._release_stream))
        _run_cleanups(
            cleanups,
            primary_error=primary_error,
            primary_traceback=primary_traceback,
        )

    def _close_source(self) -> None:
        source = self._source
        if source is None:
            return
        defer_retry = getattr(source, "_defer_process_retry_to_owner", None)
        if callable(defer_retry) and defer_retry():
            return
        try:
            self._source = _close_if_present(source)
        except BaseException as error:
            if not _contains_process_failure(error):
                self._source = None
            raise

    def _release_stream(self) -> None:
        close_callback = self._close_callback
        if close_callback is None:
            return
        try:
            self._close_callback = close_callback if close_callback() is False else None
        except BaseException as error:
            if not _contains_process_failure(error):
                self._close_callback = None
            raise

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


class SheetStream(_ResultStream[Any]):
    """One-shot stream of ordered sheet results with deterministic cleanup."""
