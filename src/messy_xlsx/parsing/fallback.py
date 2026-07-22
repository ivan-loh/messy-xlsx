"""Transactional, classified parser-backend fallback coordination."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from types import TracebackType
from typing import Any

from messy_xlsx._fallback_signals import (
    _attach_cleanup_failure,
    _attach_operation_failure,
    _blocks_backend_retry,
    _contains_process_failure,
    _exception_traceback,
    _failure_summary,
    _merge_backend_context,
    _raise_with_traceback,
    _safe_add_note,
    _type_name,
)
from messy_xlsx.parsing.contracts import ParseMetrics

_MISSING_SPECIAL = object()


@dataclass(slots=True)
class _OpenedReader:
    owner: Any
    reader: Any
    entered: bool
    exit_method: Callable[[Any, Any, Any], Any] | None = None


@dataclass(slots=True)
class _Attempt:
    value: Any = None
    error: BaseException | None = None
    traceback: TracebackType | None = None
    cleanup_failed: bool = False
    suppressed: bool = False


class FallbackCoordinator:
    """Run a primary backend and one safe compatibility fallback."""

    def __init__(
        self,
        is_compatibility_error: Callable[[Exception], bool],
        *,
        metrics: ParseMetrics | None = None,
    ) -> None:
        self._is_compatibility_error = is_compatibility_error
        self._metrics = metrics

    def materialize(
        self,
        primary_factory: Callable[[], Any],
        fallback_factory: Callable[[], Any],
    ) -> Any:
        """Materialize once, falling back only after complete primary cleanup."""
        primary = self._materialized_attempt(primary_factory)
        if primary.error is None:
            self._record_materialization()
            return primary.value

        self._record_failure()
        if primary.cleanup_failed or not self._can_fallback(primary.error):
            _raise_with_traceback(primary.error, primary.traceback)

        primary_summary = _failure_summary(primary.error)
        fallback = self._materialized_attempt(fallback_factory)
        if fallback.error is not None:
            self._record_failure()
            _attach_primary_failure(fallback.error, primary_summary)
            _raise_with_traceback(fallback.error, fallback.traceback)

        self._record_materialization()
        return fallback.value

    def batches(  # noqa: C901
        self,
        primary_factory: Callable[[], Any],
        fallback_factory: Callable[[], Any],
    ) -> Iterator[Any]:
        """Yield batches with at most one pre-output compatibility fallback."""
        opened: _OpenedReader | None = None
        using_fallback = False
        yielded = False
        primary_summary: dict[str, str] | None = None

        try:
            while True:
                if opened is None:
                    factory = fallback_factory if using_fallback else primary_factory
                    opened, attempt = _open_reader(factory, inspect_schema=True)
                    if attempt.suppressed:
                        self._record_streaming_pass()
                        return
                    if attempt.error is not None:
                        primary_summary = self._handle_stream_failure(
                            attempt,
                            yielded=yielded,
                            using_fallback=using_fallback,
                            primary_summary=primary_summary,
                        )
                        using_fallback = True
                        continue

                assert opened is not None
                try:
                    batch = opened.reader.read_next_batch()
                except BaseException as error:
                    traceback = _exception_traceback(error)
                    closed = _close_reader(opened, error, traceback)
                    opened = None
                    if closed.suppressed:
                        self._record_streaming_pass()
                        return
                    primary_summary = self._handle_stream_failure(
                        closed,
                        yielded=yielded,
                        using_fallback=using_fallback,
                        primary_summary=primary_summary,
                    )
                    using_fallback = True
                    continue

                if batch is None:
                    closed = _close_reader(opened, None, None)
                    opened = None
                    if closed.error is not None:
                        self._record_failure()
                        if using_fallback and primary_summary is not None:
                            _attach_primary_failure(closed.error, primary_summary)
                        _raise_with_traceback(closed.error, closed.traceback)
                    self._record_streaming_pass()
                    return

                yielded = True
                yield batch
        except GeneratorExit as error:
            if opened is not None:
                closed = _close_reader(
                    opened,
                    error,
                    _exception_traceback(error),
                    cleanup_overrides=True,
                )
                opened = None
                if closed.suppressed:
                    return
                assert closed.error is not None
                if closed.error is not error:
                    self._record_failure()
                    if using_fallback and primary_summary is not None:
                        _attach_primary_failure(closed.error, primary_summary)
                    _raise_with_traceback(closed.error, closed.traceback)
            raise
        except BaseException as error:
            if opened is not None:
                closed = _close_reader(opened, error, _exception_traceback(error))
                opened = None
                if closed.suppressed:
                    return
                assert closed.error is not None
                if closed.cleanup_failed:
                    self._record_failure()
                if using_fallback and primary_summary is not None:
                    _attach_primary_failure(closed.error, primary_summary)
                _raise_with_traceback(closed.error, closed.traceback)
            raise
        finally:
            if opened is not None:
                closed = _close_reader(opened, None, None)
                opened = None
                if closed.error is not None:
                    self._record_failure()
                    _raise_with_traceback(closed.error, closed.traceback)

    def _materialized_attempt(self, factory: Callable[[], Any]) -> _Attempt:
        opened, attempt = _open_reader(factory, inspect_schema=False)
        if attempt.error is not None:
            return attempt
        assert opened is not None

        value: Any = None
        error: BaseException | None = None
        traceback: TracebackType | None = None
        try:
            value = opened.reader.read_table()
        except BaseException as caught:
            error = caught
            traceback = _exception_traceback(caught)

        closed = _close_reader(opened, error, traceback)
        if closed.error is not None:
            return closed
        return _Attempt(value=value)

    def _can_fallback(self, error: BaseException) -> bool:
        if _blocks_backend_retry(error):
            return False
        if not isinstance(error, Exception):
            return False
        try:
            return self._is_compatibility_error(error)
        except BaseException as classifier_error:
            _merge_backend_context(
                classifier_error,
                primary_failure=_failure_summary(error),
                classifier_failure=_failure_summary(classifier_error),
            )
            _safe_add_note(
                classifier_error,
                f"primary backend failed: {_type_name(error)}",
            )
            raise

    def _may_retry_stream(
        self,
        attempt: _Attempt,
        *,
        yielded: bool,
        using_fallback: bool,
    ) -> bool:
        return (
            not yielded
            and not using_fallback
            and not attempt.cleanup_failed
            and attempt.error is not None
            and self._can_fallback(attempt.error)
        )

    def _handle_stream_failure(
        self,
        attempt: _Attempt,
        *,
        yielded: bool,
        using_fallback: bool,
        primary_summary: dict[str, str] | None,
    ) -> dict[str, str]:
        """Return a primary summary when retry is safe, otherwise propagate."""
        assert attempt.error is not None
        self._record_failure()
        if using_fallback and primary_summary is not None:
            _attach_primary_failure(attempt.error, primary_summary)
        if self._may_retry_stream(
            attempt,
            yielded=yielded,
            using_fallback=using_fallback,
        ):
            return _failure_summary(attempt.error)
        return _raise_with_traceback(attempt.error, attempt.traceback)

    def _record_failure(self) -> None:
        if self._metrics is not None:
            self._metrics.failed_attempts += 1

    def _record_materialization(self) -> None:
        if self._metrics is not None:
            self._metrics.full_materializations += 1

    def _record_streaming_pass(self) -> None:
        if self._metrics is not None:
            self._metrics.streaming_passes += 1


def _open_reader(
    factory: Callable[[], Any],
    *,
    inspect_schema: bool,
) -> tuple[_OpenedReader | None, _Attempt]:
    """Create and initialize one reader, closing a partially entered owner."""
    try:
        owner = factory()
    except BaseException as error:
        return None, _Attempt(error=error, traceback=_exception_traceback(error))

    try:
        exit_method = _bind_special(owner, "__exit__")
        enter = _bind_special(owner, "__enter__")
    except BaseException as error:
        opened = _OpenedReader(owner=owner, reader=owner, entered=False)
        return None, _close_reader(opened, error, _exception_traceback(error))

    has_exit = exit_method is not _MISSING_SPECIAL
    has_enter = enter is not _MISSING_SPECIAL
    if has_exit != has_enter or (has_enter and (not callable(enter) or not callable(exit_method))):
        protocol_error = TypeError(
            f"{_type_name(owner)} has an incomplete context manager protocol"
        )
        opened = _OpenedReader(owner=owner, reader=owner, entered=False)
        return None, _close_reader(
            opened,
            protocol_error,
            _exception_traceback(protocol_error),
        )

    if callable(enter):
        try:
            reader = enter()
        except BaseException as error:
            opened = _OpenedReader(owner=owner, reader=owner, entered=False)
            return None, _close_reader(opened, error, _exception_traceback(error))
        opened = _OpenedReader(
            owner=owner,
            reader=reader,
            entered=True,
            exit_method=exit_method,
        )
    else:
        opened = _OpenedReader(owner=owner, reader=owner, entered=False)

    if inspect_schema:
        try:
            if _declares_schema(opened.reader):
                schema = opened.reader.schema
                del schema
        except BaseException as error:
            return None, _close_reader(opened, error, _exception_traceback(error))

    return opened, _Attempt()


def _declares_schema(reader: Any) -> bool:
    """Avoid requiring the future protocol field from legacy test doubles."""
    class_declares = any("schema" in vars(owner) for owner in type(reader).__mro__)
    if class_declares:
        return True
    try:
        instance_state = object.__getattribute__(reader, "__dict__")
    except AttributeError:
        instance_state = {}
    return class_declares or "schema" in instance_state


def _bind_special(owner: Any, name: str) -> Any:
    """Bind a special method from the owner type/MRO, ignoring instance state."""
    owner_type = type(owner)
    for value_type in owner_type.__mro__:
        namespace = vars(value_type)
        if name not in namespace:
            continue
        descriptor = namespace[name]
        binder = getattr(type(descriptor), "__get__", None)
        if binder is None:
            return descriptor
        return binder(descriptor, owner, owner_type)
    return _MISSING_SPECIAL


def _close_reader(
    opened: _OpenedReader,
    primary_error: BaseException | None,
    primary_traceback: TracebackType | None,
    *,
    cleanup_overrides: bool = False,
) -> _Attempt:
    """Close once, keeping an ordinary cleanup failure attached to its cause."""
    try:
        if opened.entered:
            assert opened.exit_method is not None
            exit_result = opened.exit_method(
                type(primary_error) if primary_error is not None else None,
                primary_error,
                primary_traceback,
            )
            suppressed = (
                primary_error is not None
                and _is_suppressible_parse_failure(primary_error)
                and bool(exit_result)
            )
        else:
            suppressed = False
            close = getattr(opened.owner, "close", None)
            if callable(close):
                close()
    except BaseException as cleanup_error:
        cleanup_traceback = _exception_traceback(cleanup_error)
        if primary_error is None:
            _safe_add_note(
                cleanup_error,
                f"reader cleanup failed: {_type_name(cleanup_error)}",
            )
            _attach_cleanup_failure(cleanup_error, cleanup_error)
            return _Attempt(
                error=cleanup_error,
                traceback=cleanup_traceback,
                cleanup_failed=True,
            )
        if cleanup_overrides or _cleanup_takes_precedence(cleanup_error, primary_error):
            _safe_add_note(
                cleanup_error,
                f"backend operation also failed: {_type_name(primary_error)}",
            )
            _attach_operation_failure(cleanup_error, primary_error)
            _attach_cleanup_failure(cleanup_error, cleanup_error)
            return _Attempt(
                error=cleanup_error,
                traceback=cleanup_traceback,
                cleanup_failed=True,
            )
        _safe_add_note(
            primary_error,
            f"reader cleanup also failed: {_type_name(cleanup_error)}",
        )
        _attach_cleanup_failure(primary_error, cleanup_error)
        return _Attempt(
            error=primary_error,
            traceback=primary_traceback,
            cleanup_failed=True,
        )

    if primary_error is not None and not suppressed:
        return _Attempt(error=primary_error, traceback=primary_traceback)
    return _Attempt(suppressed=suppressed and primary_error is not None)


def _cleanup_takes_precedence(
    error: BaseException,
    primary_error: BaseException,
) -> bool:
    """Return whether teardown must replace an operation failure."""
    return _contains_process_failure(error, exclude=primary_error)


def _is_suppressible_parse_failure(error: BaseException) -> bool:
    """Limit context suppression to ordinary, recoverable parse failures."""
    return isinstance(error, Exception) and not _contains_process_failure(error)


def _attach_primary_failure(
    fallback_error: BaseException,
    primary_summary: dict[str, str],
) -> None:
    fallback_summary = _failure_summary(fallback_error)
    _merge_backend_context(
        fallback_error,
        primary_failure=dict(primary_summary),
        fallback_failure=fallback_summary,
    )
    _safe_add_note(
        fallback_error,
        f"primary backend failed: {primary_summary['type']}",
    )
