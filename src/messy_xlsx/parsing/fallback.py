"""Transactional, classified parser-backend fallback coordination."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from types import TracebackType
from typing import Any, NoReturn

from messy_xlsx.parsing.contracts import ParseMetrics

_CONFIGURATION_MARKERS = (
    "batch_size",
    "invalid config",
    "invalid configuration",
)
_OWNERSHIP_MARKERS = (
    "active borrow",
    "caller-owned",
    "cursor restoration",
    "source ownership",
)
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
                    traceback = error.__traceback__
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
                    error.__traceback__,
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
                closed = _close_reader(opened, error, error.__traceback__)
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
            traceback = caught.__traceback__

        closed = _close_reader(opened, error, traceback)
        if closed.error is not None:
            return closed
        return _Attempt(value=value)

    def _can_fallback(self, error: BaseException) -> bool:
        if not isinstance(error, Exception):
            return False
        if isinstance(error, (PermissionError, FileNotFoundError, MemoryError)):
            return False

        message = str(error).lower()
        if isinstance(error, ValueError) and any(
            marker in message for marker in _CONFIGURATION_MARKERS
        ):
            return False
        if isinstance(error, RuntimeError) and any(
            marker in message for marker in _OWNERSHIP_MARKERS
        ):
            return False
        try:
            return self._is_compatibility_error(error)
        except BaseException as classifier_error:
            _merge_backend_context(
                classifier_error,
                primary_failure=_failure_summary(error),
                classifier_failure=_failure_summary(classifier_error),
            )
            classifier_error.add_note(f"primary backend failed: {type(error).__name__}")
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
        raise attempt.error.with_traceback(attempt.traceback)

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
        return None, _Attempt(error=error, traceback=error.__traceback__)

    try:
        exit_method = _bind_special(owner, "__exit__")
        enter = _bind_special(owner, "__enter__")
    except BaseException as error:
        opened = _OpenedReader(owner=owner, reader=owner, entered=False)
        return None, _close_reader(opened, error, error.__traceback__)

    has_exit = exit_method is not _MISSING_SPECIAL
    has_enter = enter is not _MISSING_SPECIAL
    if has_exit != has_enter or (has_enter and (not callable(enter) or not callable(exit_method))):
        protocol_error = TypeError(
            f"{type(owner).__name__} has an incomplete context manager protocol"
        )
        opened = _OpenedReader(owner=owner, reader=owner, entered=False)
        return None, _close_reader(
            opened,
            protocol_error,
            protocol_error.__traceback__,
        )

    if callable(enter):
        try:
            reader = enter()
        except BaseException as error:
            opened = _OpenedReader(owner=owner, reader=owner, entered=False)
            return None, _close_reader(opened, error, error.__traceback__)
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
            return None, _close_reader(opened, error, error.__traceback__)

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
            suppressed = bool(
                opened.exit_method(
                    type(primary_error) if primary_error is not None else None,
                    primary_error,
                    primary_traceback,
                )
            )
        else:
            suppressed = False
            close = getattr(opened.owner, "close", None)
            if callable(close):
                close()
    except BaseException as cleanup_error:
        cleanup_traceback = cleanup_error.__traceback__
        if primary_error is None:
            cleanup_error.add_note(f"reader cleanup failed: {type(cleanup_error).__name__}")
            _attach_cleanup_failure(cleanup_error, cleanup_error)
            return _Attempt(
                error=cleanup_error,
                traceback=cleanup_traceback,
                cleanup_failed=True,
            )
        if cleanup_overrides or _cleanup_takes_precedence(cleanup_error):
            cleanup_error.add_note(f"backend operation also failed: {type(primary_error).__name__}")
            _attach_operation_failure(cleanup_error, primary_error)
            _attach_cleanup_failure(cleanup_error, cleanup_error)
            return _Attempt(
                error=cleanup_error,
                traceback=cleanup_traceback,
                cleanup_failed=True,
            )
        primary_error.add_note(f"reader cleanup also failed: {type(cleanup_error).__name__}")
        _attach_cleanup_failure(primary_error, cleanup_error)
        return _Attempt(
            error=primary_error,
            traceback=primary_traceback,
            cleanup_failed=True,
        )

    if primary_error is not None and not suppressed:
        return _Attempt(error=primary_error, traceback=primary_traceback)
    return _Attempt(suppressed=suppressed and primary_error is not None)


def _cleanup_takes_precedence(error: BaseException) -> bool:
    """Return whether teardown must replace an operation failure."""
    return isinstance(error, (MemoryError, KeyboardInterrupt, SystemExit)) or not isinstance(
        error,
        Exception,
    )


def _failure_summary(error: BaseException) -> dict[str, str]:
    """Return useful backend context without source paths or cell values."""
    return {"type": type(error).__name__}


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
    fallback_error.add_note(f"primary backend failed: {primary_summary['type']}")


def _attach_cleanup_failure(
    primary_error: BaseException,
    cleanup_error: BaseException,
) -> None:
    _merge_backend_context(
        primary_error,
        cleanup_failure=_failure_summary(cleanup_error),
    )


def _attach_operation_failure(
    cleanup_error: BaseException,
    operation_error: BaseException,
) -> None:
    _merge_backend_context(
        cleanup_error,
        operation_failure=_failure_summary(operation_error),
    )


def _merge_backend_context(
    error: BaseException,
    **updates: dict[str, str],
) -> None:
    """Merge only type summaries produced by this coordinator."""
    context: dict[str, dict[str, str]] = {}
    existing = getattr(error, "backend_context", None)
    if isinstance(existing, dict):
        for name in (
            "primary_failure",
            "fallback_failure",
            "operation_failure",
            "cleanup_failure",
            "classifier_failure",
        ):
            summary = existing.get(name)
            if (
                isinstance(summary, dict)
                and set(summary) == {"type"}
                and isinstance(summary["type"], str)
            ):
                context[name] = {"type": summary["type"]}
    context.update({name: dict(summary) for name, summary in updates.items()})
    error.backend_context = context  # type: ignore[attr-defined]


def _raise_with_traceback(
    error: BaseException,
    traceback: TracebackType | None,
) -> NoReturn:
    raise error.with_traceback(traceback)
