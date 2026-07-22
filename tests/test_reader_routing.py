"""Capability and extension-aware backend routing contracts."""

from __future__ import annotations

from types import MethodType
from typing import Any

import pyarrow as pa
import pytest

from messy_xlsx.detection.format_detector import FormatDetector
from messy_xlsx.parsing.contracts import BackendKind, OutputMode, ParseMetrics
from messy_xlsx.parsing.csv_handler import CSVHandler
from messy_xlsx.parsing.fallback import FallbackCoordinator
from messy_xlsx.parsing.handler_registry import HandlerRegistry
from messy_xlsx.parsing.router import BackendRouter, WorkbookContext
from messy_xlsx.parsing.xls_handler import XLSHandler
from messy_xlsx.parsing.xlsx_handler import XLSXHandler


@pytest.mark.parametrize("format_type", ["xlsx", "xlsm", "xltx", "xltm"])
@pytest.mark.parametrize(
    ("mode", "evaluate_formulas", "expected"),
    [
        (OutputMode.MATERIALIZED, True, BackendKind.FASTEXCEL),
        (OutputMode.MATERIALIZED, False, BackendKind.OPENPYXL_COMPAT),
        (OutputMode.STREAMING, True, BackendKind.OPENPYXL_STREAMING),
        (OutputMode.STREAMING, False, BackendKind.OPENPYXL_STREAMING),
    ],
)
def test_ooxml_routing_matrix(
    format_type: str,
    mode: OutputMode,
    evaluate_formulas: bool,
    expected: BackendKind,
) -> None:
    context = WorkbookContext(
        format_type=format_type,
        output_mode=mode,
        evaluate_formulas=evaluate_formulas,
        has_custom_registry=False,
    )

    assert BackendRouter().select(context).backend is expected


@pytest.mark.parametrize("format_type", ["csv", "tsv", "txt"])
@pytest.mark.parametrize("mode", list(OutputMode))
def test_text_routing_uses_the_chunk_reader(
    format_type: str,
    mode: OutputMode,
) -> None:
    context = WorkbookContext(
        format_type=format_type,
        output_mode=mode,
        evaluate_formulas=True,
        has_custom_registry=False,
    )

    assert BackendRouter().select(context).backend is BackendKind.CSV_STREAMING


@pytest.mark.parametrize("mode", list(OutputMode))
def test_xls_routing_uses_the_row_reader(mode: OutputMode) -> None:
    context = WorkbookContext(
        format_type="xls",
        output_mode=mode,
        evaluate_formulas=True,
        has_custom_registry=False,
    )

    assert BackendRouter().select(context).backend is BackendKind.XLS_STREAMING


@pytest.mark.parametrize(
    "format_type",
    ["xlsx", "xlsm", "xltx", "xltm", "csv", "tsv", "txt", "xls"],
)
@pytest.mark.parametrize("mode", list(OutputMode))
def test_custom_components_always_route_through_the_dataframe_spi(
    format_type: str,
    mode: OutputMode,
) -> None:
    context = WorkbookContext(
        format_type=format_type,
        output_mode=mode,
        evaluate_formulas=True,
        has_custom_registry=True,
    )

    assert BackendRouter().select(context).backend is BackendKind.CUSTOM_DATAFRAME


def test_only_the_untouched_exact_builtin_registry_is_fast_path_eligible() -> None:
    registry = HandlerRegistry()

    assert registry._uses_builtin_components() is True


def test_constructor_injection_is_not_mistaken_for_the_builtin_registry() -> None:
    registry = HandlerRegistry(handlers=[XLSXHandler(), XLSHandler(), CSVHandler()])
    detector_registry = HandlerRegistry(detector=FormatDetector())

    assert registry._uses_builtin_components() is False
    assert detector_registry._uses_builtin_components() is False


def test_registry_and_component_subclasses_are_not_fast_path_eligible() -> None:
    class RegistrySubclass(HandlerRegistry):
        pass

    class HandlerSubclass(XLSXHandler):
        pass

    class DetectorSubclass(FormatDetector):
        pass

    subclassed_registry = RegistrySubclass()
    handler_registry = HandlerRegistry(handlers=[HandlerSubclass(), XLSHandler(), CSVHandler()])
    detector_registry = HandlerRegistry(detector=DetectorSubclass())

    assert subclassed_registry._uses_builtin_components() is False
    assert handler_registry._uses_builtin_components() is False
    assert detector_registry._uses_builtin_components() is False


def test_later_handler_list_mutations_force_the_compatibility_path() -> None:
    registries = [HandlerRegistry() for _ in range(3)]

    registries[0].register_handler(CSVHandler())
    registries[1].handlers.reverse()
    registries[2].handlers[0] = XLSXHandler()

    assert all(not registry._uses_builtin_components() for registry in registries)


def test_instance_level_registry_and_handler_overrides_force_compatibility() -> None:
    registry_override = HandlerRegistry()
    handler_override = HandlerRegistry()

    registry_override.parse = MethodType(  # type: ignore[method-assign]
        lambda self, *_args, **_kwargs: None,
        registry_override,
    )
    handler_override.handlers[0].parse = MethodType(  # type: ignore[method-assign]
        lambda self, *_args, **_kwargs: None,
        handler_override.handlers[0],
    )

    assert registry_override._uses_builtin_components() is False
    assert handler_override._uses_builtin_components() is False


def test_mutating_a_nested_builtin_handler_component_forces_compatibility() -> None:
    registry = HandlerRegistry()
    csv_handler = registry.handlers[2]
    assert isinstance(csv_handler, CSVHandler)

    csv_handler._detector = object()  # type: ignore[assignment]

    assert registry._uses_builtin_components() is False


def test_class_level_handler_override_before_registry_construction_is_custom(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        XLSXHandler,
        "parse",
        lambda self, *_args, **_kwargs: None,
    )

    assert HandlerRegistry()._uses_builtin_components() is False


def test_class_level_detector_override_before_registry_construction_is_custom(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        FormatDetector,
        "detect",
        lambda self, *_args, **_kwargs: None,
    )

    assert HandlerRegistry()._uses_builtin_components() is False


def test_class_level_registry_override_before_construction_is_custom(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        HandlerRegistry,
        "parse",
        lambda self, *_args, **_kwargs: None,
    )

    assert HandlerRegistry()._uses_builtin_components() is False


def test_router_rejects_unsupported_formats_before_any_backend_factory_exists() -> None:
    context = WorkbookContext(
        format_type="unknown",
        output_mode=OutputMode.MATERIALIZED,
        evaluate_formulas=True,
        has_custom_registry=False,
    )

    with pytest.raises(ValueError, match="Unsupported format"):
        BackendRouter().select(context)


class _CompatibilityFailure(Exception):
    pass


def _is_compatibility_failure(error: Exception) -> bool:
    return isinstance(error, _CompatibilityFailure)


class _MaterializedReaderFake:
    def __init__(
        self,
        events: list[str],
        name: str,
        *,
        result: pa.Table | None = None,
        enter_error: BaseException | None = None,
        read_error: BaseException | None = None,
        close_error: BaseException | None = None,
    ) -> None:
        self.events = events
        self.name = name
        self.result = result or pa.table({"value": [1]})
        self.enter_error = enter_error
        self.read_error = read_error
        self.close_error = close_error
        self.close_calls = 0

    def __enter__(self) -> _MaterializedReaderFake:
        self.events.append(f"{self.name}-enter")
        if self.enter_error is not None:
            raise self.enter_error
        return self

    def read_table(self) -> pa.Table:
        self.events.append(f"{self.name}-read")
        if self.read_error is not None:
            raise self.read_error
        return self.result

    def close(self) -> None:
        self.close_calls += 1
        self.events.append(f"{self.name}-close")
        if self.close_error is not None:
            raise self.close_error

    def __exit__(
        self,
        _error_type: type[BaseException] | None,
        _error: BaseException | None,
        _traceback: object,
    ) -> bool:
        self.close()
        return False


class _StreamingReaderFake:
    def __init__(
        self,
        events: list[str],
        name: str,
        outcomes: list[pa.RecordBatch | BaseException | None],
        *,
        schema_error: BaseException | None = None,
        close_error: BaseException | None = None,
    ) -> None:
        self.events = events
        self.name = name
        self.outcomes = list(outcomes)
        self.schema_error = schema_error
        self.close_error = close_error
        self.close_calls = 0

    @property
    def schema(self) -> pa.Schema:
        self.events.append(f"{self.name}-schema")
        if self.schema_error is not None:
            raise self.schema_error
        return pa.schema([("value", pa.int64())])

    def read_next_batch(self) -> pa.RecordBatch | None:
        self.events.append(f"{self.name}-read")
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, BaseException):
            raise outcome
        return outcome

    def close(self) -> None:
        self.close_calls += 1
        self.events.append(f"{self.name}-close")
        if self.close_error is not None:
            raise self.close_error


def _batch(value: int) -> pa.RecordBatch:
    return pa.record_batch({"value": [value]})


def test_materialized_fallback_closes_primary_before_opening_fallback() -> None:
    events: list[str] = []
    primary = _MaterializedReaderFake(
        events,
        "primary",
        read_error=_CompatibilityFailure("unsupported representation"),
    )
    fallback = _MaterializedReaderFake(
        events,
        "fallback",
        result=pa.table({"value": [2]}),
    )

    result = FallbackCoordinator(_is_compatibility_failure).materialize(
        lambda: (events.append("primary-factory"), primary)[1],
        lambda: (events.append("fallback-factory"), fallback)[1],
    )

    assert result.to_pydict() == {"value": [2]}
    assert events.index("primary-close") < events.index("fallback-factory")
    assert primary.close_calls == fallback.close_calls == 1


def test_materialized_reader_is_closed_when_context_initialization_fails() -> None:
    events: list[str] = []
    primary = _MaterializedReaderFake(
        events,
        "primary",
        enter_error=_CompatibilityFailure("cannot initialize"),
    )
    fallback = _MaterializedReaderFake(events, "fallback")

    FallbackCoordinator(_is_compatibility_failure).materialize(
        lambda: primary,
        lambda: fallback,
    )

    assert primary.close_calls == fallback.close_calls == 1
    assert events.index("primary-close") < events.index("fallback-enter")


@pytest.mark.parametrize(
    "error",
    [
        PermissionError("denied"),
        FileNotFoundError("missing"),
        MemoryError("capacity"),
        ValueError("invalid configuration"),
        RuntimeError("SourceHandle already has an active borrow"),
    ],
)
def test_unclassified_failures_never_materialize_a_fallback(error: Exception) -> None:
    fallback_calls = 0

    def fallback_factory() -> _MaterializedReaderFake:
        nonlocal fallback_calls
        fallback_calls += 1
        return _MaterializedReaderFake([], "fallback")

    with pytest.raises(type(error), match=str(error)):
        FallbackCoordinator(_is_compatibility_failure).materialize(
            lambda: _MaterializedReaderFake([], "primary", read_error=error),
            fallback_factory,
        )

    assert fallback_calls == 0


def test_cleanup_failure_blocks_materialized_fallback_and_stays_attached() -> None:
    fallback_calls = 0
    primary_error = _CompatibilityFailure("unsupported")
    primary = _MaterializedReaderFake(
        [],
        "primary",
        read_error=primary_error,
        close_error=OSError("restore failed"),
    )

    def fallback_factory() -> _MaterializedReaderFake:
        nonlocal fallback_calls
        fallback_calls += 1
        return _MaterializedReaderFake([], "fallback")

    with pytest.raises(_CompatibilityFailure) as captured:
        FallbackCoordinator(_is_compatibility_failure).materialize(
            lambda: primary,
            fallback_factory,
        )

    assert captured.value is primary_error
    assert fallback_calls == 0
    assert primary.close_calls == 1
    assert any("cleanup" in note and "restore failed" in note for note in captured.value.__notes__)


def test_materialized_fallback_failure_has_sanitized_structured_context() -> None:
    secret = "/customers/acme/private-payroll.xlsx cell salary=999999"
    metrics = ParseMetrics()

    with pytest.raises(RuntimeError, match="fallback failed") as captured:
        FallbackCoordinator(
            _is_compatibility_failure,
            metrics=metrics,
        ).materialize(
            lambda: _MaterializedReaderFake(
                [],
                "primary",
                read_error=_CompatibilityFailure(secret),
            ),
            lambda: _MaterializedReaderFake(
                [],
                "fallback",
                read_error=RuntimeError("fallback failed"),
            ),
        )

    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "primary_failure": {"type": "_CompatibilityFailure"},
        "fallback_failure": {"type": "RuntimeError"},
    }
    assert secret not in repr(captured.value.backend_context)  # type: ignore[attr-defined]
    assert secret not in repr(captured.value.__notes__)
    assert metrics.failed_attempts == 2
    assert metrics.full_materializations == 0


def test_materialized_metrics_count_one_successful_result_and_failed_attempt() -> None:
    metrics = ParseMetrics()

    FallbackCoordinator(_is_compatibility_failure, metrics=metrics).materialize(
        lambda: _MaterializedReaderFake(
            [],
            "primary",
            read_error=_CompatibilityFailure("unsupported"),
        ),
        lambda: _MaterializedReaderFake([], "fallback"),
    )

    assert metrics.failed_attempts == 1
    assert metrics.full_materializations == 1


def test_streaming_fallback_is_allowed_before_the_first_observable_batch() -> None:
    events: list[str] = []
    primary = _StreamingReaderFake(
        events,
        "primary",
        [_CompatibilityFailure("unsupported")],
    )
    fallback = _StreamingReaderFake(events, "fallback", [_batch(2), None])

    batches = list(
        FallbackCoordinator(_is_compatibility_failure).batches(
            lambda: (events.append("primary-factory"), primary)[1],
            lambda: (events.append("fallback-factory"), fallback)[1],
        )
    )

    assert [batch.to_pydict() for batch in batches] == [{"value": [2]}]
    assert events.index("primary-close") < events.index("fallback-factory")
    assert primary.close_calls == fallback.close_calls == 1


def test_streaming_schema_failure_closes_primary_before_fallback() -> None:
    events: list[str] = []
    primary = _StreamingReaderFake(
        events,
        "primary",
        [],
        schema_error=_CompatibilityFailure("schema unsupported"),
    )
    fallback = _StreamingReaderFake(events, "fallback", [None])

    assert (
        list(
            FallbackCoordinator(_is_compatibility_failure).batches(
                lambda: primary,
                lambda: (events.append("fallback-factory"), fallback)[1],
            )
        )
        == []
    )
    assert events.index("primary-close") < events.index("fallback-factory")
    assert primary.close_calls == fallback.close_calls == 1


def test_streaming_never_restarts_after_the_first_yield() -> None:
    fallback_calls = 0
    primary = _StreamingReaderFake(
        [],
        "primary",
        [_batch(1), _CompatibilityFailure("late incompatibility")],
    )

    def fallback_factory() -> _StreamingReaderFake:
        nonlocal fallback_calls
        fallback_calls += 1
        return _StreamingReaderFake([], "fallback", [None])

    stream = FallbackCoordinator(_is_compatibility_failure).batches(
        lambda: primary,
        fallback_factory,
    )
    assert next(stream).to_pydict() == {"value": [1]}
    with pytest.raises(_CompatibilityFailure, match="late incompatibility"):
        next(stream)

    assert fallback_calls == 0
    assert primary.close_calls == 1


def test_streaming_cleanup_failure_blocks_an_unsafe_fallback() -> None:
    fallback_calls = 0
    primary_error = _CompatibilityFailure("unsupported")
    primary = _StreamingReaderFake(
        [],
        "primary",
        [primary_error],
        close_error=OSError("cursor restore failed"),
    )

    def fallback_factory() -> _StreamingReaderFake:
        nonlocal fallback_calls
        fallback_calls += 1
        return _StreamingReaderFake([], "fallback", [None])

    with pytest.raises(_CompatibilityFailure) as captured:
        list(
            FallbackCoordinator(_is_compatibility_failure).batches(
                lambda: primary,
                fallback_factory,
            )
        )

    assert captured.value is primary_error
    assert fallback_calls == 0
    assert primary.close_calls == 1
    assert any("cleanup" in note for note in captured.value.__notes__)


def test_streaming_exhaustion_and_generator_close_each_close_exactly_once() -> None:
    metrics = ParseMetrics()
    exhausted = _StreamingReaderFake([], "exhausted", [_batch(1), None])
    interrupted = _StreamingReaderFake([], "interrupted", [_batch(2), None])
    coordinator = FallbackCoordinator(_is_compatibility_failure, metrics=metrics)

    assert len(list(coordinator.batches(lambda: exhausted, pytest.fail))) == 1
    stream = coordinator.batches(lambda: interrupted, pytest.fail)
    assert next(stream).to_pydict() == {"value": [2]}
    stream.close()

    assert exhausted.close_calls == interrupted.close_calls == 1
    assert metrics.streaming_passes == 1


def test_closing_an_unstarted_stream_does_not_initialize_a_backend() -> None:
    factory_calls = 0

    def primary_factory() -> _StreamingReaderFake:
        nonlocal factory_calls
        factory_calls += 1
        return _StreamingReaderFake([], "primary", [None])

    stream = FallbackCoordinator(_is_compatibility_failure).batches(
        primary_factory,
        pytest.fail,
    )
    stream.close()

    assert factory_calls == 0


def test_generator_close_cleanup_failure_is_counted_and_propagated() -> None:
    metrics = ParseMetrics()
    reader = _StreamingReaderFake(
        [],
        "primary",
        [_batch(1), None],
        close_error=OSError("close failed"),
    )
    stream = FallbackCoordinator(
        _is_compatibility_failure,
        metrics=metrics,
    ).batches(lambda: reader, pytest.fail)
    assert next(stream).to_pydict() == {"value": [1]}

    with pytest.raises(OSError, match="close failed"):
        stream.close()

    assert reader.close_calls == 1
    assert metrics.failed_attempts == 1
    assert metrics.streaming_passes == 0


@pytest.mark.parametrize("error", [KeyboardInterrupt(), SystemExit(2)])
def test_process_level_failures_are_never_swallowed_or_retried(
    error: BaseException,
) -> None:
    fallback_calls = 0

    def fallback_factory() -> Any:
        nonlocal fallback_calls
        fallback_calls += 1
        pytest.fail("process-level failure must not retry")

    reader = _StreamingReaderFake([], "primary", [error])
    with pytest.raises(type(error)):
        list(
            FallbackCoordinator(_is_compatibility_failure).batches(
                lambda: reader,
                fallback_factory,
            )
        )

    assert fallback_calls == 0
    assert reader.close_calls == 1
