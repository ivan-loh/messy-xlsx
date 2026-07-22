"""Capability and extension-aware backend routing contracts."""

from __future__ import annotations

from types import MethodType
from typing import Any

import pyarrow as pa
import pytest

import messy_xlsx.parsing.handler_registry as handler_registry_module
from messy_xlsx.detection.format_detector import FormatDetector
from messy_xlsx.parsing.contracts import BackendKind, OutputMode, ParseMetrics
from messy_xlsx.parsing.csv_handler import CSVHandler, MetadataRowDetector
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


def test_nested_csv_detector_instance_override_forces_compatibility() -> None:
    registry = HandlerRegistry()
    csv_handler = registry.handlers[2]
    assert isinstance(csv_handler, CSVHandler)

    csv_handler._detector.detect_skip_rows_from_text = MethodType(  # type: ignore[method-assign]
        lambda self, *_args, **_kwargs: 0,
        csv_handler._detector,
    )

    assert registry._uses_builtin_components() is False


def test_nested_component_container_state_and_cycles_force_compatibility() -> None:
    registry = HandlerRegistry()
    csv_handler = registry.handlers[2]
    assert isinstance(csv_handler, CSVHandler)
    cycle: list[object] = []
    cycle.append(cycle)
    detector_state = vars(csv_handler._detector)
    assert detector_state == {}
    detector_state["routing_state"] = {
        "rules": ["metadata", {"enabled": True}],
        "cycle": cycle,
    }

    assert [registry._uses_builtin_components() for _ in range(3)] == [False] * 3


def test_private_handler_behavior_override_and_restore_is_detected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    registry = HandlerRegistry()

    with monkeypatch.context() as patch:
        patch.setattr(
            CSVHandler,
            "_detect_delimiter_from_text",
            lambda self, _sample: "|",
        )
        assert registry._uses_builtin_components() is False

    assert registry._uses_builtin_components() is True


def test_private_handler_behavior_flag_override_is_detected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    registry = HandlerRegistry()

    with monkeypatch.context() as patch:
        patch.setattr(CSVHandler, "_accepts_source_handle", False)
        assert registry._uses_builtin_components() is False

    assert registry._uses_builtin_components() is True


def test_contextmanager_wrapper_closure_mutation_and_restore_is_detected() -> None:
    registry = HandlerRegistry()
    wrapper = vars(HandlerRegistry)["_source_handle"]
    assert wrapper.__closure__ is not None
    closure_cell = wrapper.__closure__[0]
    original = closure_cell.cell_contents

    def replacement(*_args: object, **_kwargs: object) -> object:
        return object()

    try:
        closure_cell.cell_contents = replacement
        assert registry._uses_builtin_components() is False
    finally:
        closure_cell.cell_contents = original

    assert registry._uses_builtin_components() is True


def test_private_descriptor_override_before_construction_and_restore_is_detected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with monkeypatch.context() as patch:
        patch.setattr(
            HandlerRegistry,
            "_handler_accepts_source_handle",
            staticmethod(lambda _handler: False),
        )
        registry = HandlerRegistry()
        assert registry._uses_builtin_components() is False

    assert registry._uses_builtin_components() is True


def test_nested_detector_class_override_is_detected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    registry = HandlerRegistry()
    monkeypatch.setattr(
        MetadataRowDetector,
        "_score_as_metadata",
        lambda self, _profile, _consensus: 0.0,
    )

    assert registry._uses_builtin_components() is False


def test_added_raising_descriptor_is_detected_without_invocation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = 0

    class RaisingDescriptor:
        def __get__(self, _instance: object, _owner: type[object]) -> object:
            nonlocal calls
            calls += 1
            raise AssertionError("descriptor must not be invoked while routing")

    monkeypatch.setattr(
        CSVHandler,
        "_routing_probe",
        RaisingDescriptor(),
        raising=False,
    )

    registry = HandlerRegistry()
    assert [registry._uses_builtin_components() for _ in range(3)] == [False] * 3
    assert calls == 0


def test_multiple_ordinary_builtin_registries_remain_fast_path_eligible() -> None:
    registries = [HandlerRegistry(), HandlerRegistry(), HandlerRegistry()]

    assert all(registry._uses_builtin_components() for registry in registries for _ in range(3))


@pytest.mark.parametrize(
    "fingerprinter",
    [
        pytest.param(
            lambda: handler_registry_module._CompositionFingerprinter(include_identity=True),
            id="composition",
        ),
        pytest.param(
            handler_registry_module._BehaviorFingerprinter,
            id="behavior",
        ),
    ],
)
@pytest.mark.parametrize(
    "value",
    [
        pytest.param([0] * 1_000_000, id="million-primitives"),
        pytest.param(b"x" * 1_000_000, id="large-bytes"),
    ],
)
def test_fingerprint_budget_counts_primitive_edges_and_bulk_bytes(
    fingerprinter: Any,
    value: object,
) -> None:
    with pytest.raises(RuntimeError, match="fingerprint budget"):
        fingerprinter().token(value)


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


class _FatalCleanup(BaseException):
    pass


class _ExplodingTruth:
    def __init__(self) -> None:
        self.calls = 0

    def __bool__(self) -> bool:
        self.calls += 1
        raise AssertionError("process-level exit result must not be truth-tested")


class _HostileExceptionMeta(type):
    def __getattribute__(cls, name: str) -> object:
        if name == "__name__":
            raise AssertionError("diagnostics must bypass metaclass name lookup")
        return type.__getattribute__(cls, name)


class _HostileDiagnosticError(RuntimeError, metaclass=_HostileExceptionMeta):
    def __getattribute__(self, name: str) -> object:
        if name in {
            "__dict__",
            "__traceback__",
            "add_note",
            "backend_context",
            "with_traceback",
        }:
            raise AssertionError(f"diagnostics must bypass {name}")
        return BaseException.__getattribute__(self, name)

    def __setattr__(self, name: str, value: object) -> None:
        if name == "backend_context":
            raise AssertionError("diagnostics must bypass backend_context assignment")
        BaseException.__setattr__(self, name, value)

    def __str__(self) -> str:
        raise AssertionError("diagnostics must not stringify backend failures")

    def add_note(self, note: str) -> None:
        del note
        raise AssertionError("diagnostics must bypass add_note overrides")

    def with_traceback(self, traceback: object) -> BaseException:
        del traceback
        raise AssertionError("diagnostics must bypass with_traceback overrides")


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
        suppress: object = False,
    ) -> None:
        self.events = events
        self.name = name
        self.result = result or pa.table({"value": [1]})
        self.enter_error = enter_error
        self.read_error = read_error
        self.close_error = close_error
        self.suppress = suppress
        self.close_calls = 0
        self.exit_triples: list[
            tuple[type[BaseException] | None, BaseException | None, object]
        ] = []
        self.exit_traceback_matches_error: list[bool] = []

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
        error_type: type[BaseException] | None,
        error: BaseException | None,
        traceback: object,
    ) -> object:
        self.exit_triples.append((error_type, error, traceback))
        self.exit_traceback_matches_error.append(error is None or traceback is error.__traceback__)
        self.close()
        return self.suppress


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


class _ContextStreamingReaderFake(_StreamingReaderFake):
    def __init__(
        self,
        events: list[str],
        name: str,
        outcomes: list[pa.RecordBatch | BaseException | None],
        *,
        schema_error: BaseException | None = None,
        close_error: BaseException | None = None,
        suppress: object = False,
    ) -> None:
        super().__init__(
            events,
            name,
            outcomes,
            schema_error=schema_error,
            close_error=close_error,
        )
        self.suppress = suppress
        self.enter_calls = 0
        self.exit_triples: list[
            tuple[type[BaseException] | None, BaseException | None, object]
        ] = []
        self.exit_traceback_matches_error: list[bool] = []

    def __enter__(self) -> _ContextStreamingReaderFake:
        self.enter_calls += 1
        return self

    def __exit__(
        self,
        error_type: type[BaseException] | None,
        error: BaseException | None,
        traceback: object,
    ) -> object:
        self.exit_triples.append((error_type, error, traceback))
        self.exit_traceback_matches_error.append(error is None or traceback is error.__traceback__)
        self.close()
        return self.suppress


class _RaisingSchemaDiscoveryReader:
    def __init__(self) -> None:
        self.enter_calls = 0
        self.exit_triples: list[
            tuple[type[BaseException] | None, BaseException | None, object]
        ] = []

    @property
    def __dict__(self) -> dict[str, object]:
        raise RuntimeError("schema discovery failed")

    def __enter__(self) -> _RaisingSchemaDiscoveryReader:
        self.enter_calls += 1
        return self

    def __exit__(
        self,
        error_type: type[BaseException] | None,
        error: BaseException | None,
        traceback: object,
    ) -> bool:
        self.exit_triples.append((error_type, error, traceback))
        return False

    def read_next_batch(self) -> pa.RecordBatch | None:
        pytest.fail("schema discovery failure must happen before reading")


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
    assert primary.exit_triples == []
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
        close_error=OSError("secret restore path"),
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
    assert any("cleanup" in note and "OSError" in note for note in captured.value.__notes__)
    assert "secret" not in repr(captured.value.__notes__)


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


def test_hostile_fallback_diagnostics_never_mask_the_exact_backend_failure() -> None:
    fallback_error = _HostileDiagnosticError("fallback failed")

    class HostileFailureReader:
        def read_table(self) -> pa.Table:
            raise fallback_error

        def close(self) -> None:
            pass

    with pytest.raises(RuntimeError) as captured:
        FallbackCoordinator(_is_compatibility_failure).materialize(
            lambda: _MaterializedReaderFake(
                [],
                "primary",
                read_error=_CompatibilityFailure("unsupported"),
            ),
            HostileFailureReader,
        )

    assert captured.value is fallback_error
    state = BaseException.__getattribute__(fallback_error, "__dict__")
    assert state["backend_context"] == {
        "primary_failure": {"type": "_CompatibilityFailure"},
        "fallback_failure": {"type": "_HostileDiagnosticError"},
    }
    notes = BaseException.__getattribute__(fallback_error, "__notes__")
    assert notes == ["primary backend failed: _CompatibilityFailure"]
    traceback = BaseException.__getattribute__(fallback_error, "__traceback__")
    frame_names: list[str] = []
    while traceback is not None:
        frame_names.append(traceback.tb_frame.f_code.co_name)
        traceback = traceback.tb_next
    assert "read_table" in frame_names


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


def test_context_special_methods_ignore_instance_shadows() -> None:
    reader = _MaterializedReaderFake([], "primary")
    reader.__enter__ = lambda: pytest.fail("instance __enter__ must be ignored")  # type: ignore[method-assign]
    reader.__exit__ = lambda *_args: pytest.fail("instance __exit__ must be ignored")  # type: ignore[method-assign]

    result = FallbackCoordinator(_is_compatibility_failure).materialize(
        lambda: reader,
        pytest.fail,
    )

    assert result.to_pydict() == {"value": [1]}
    assert reader.close_calls == 1


def test_context_special_method_descriptors_bind_exit_before_enter() -> None:
    events: list[str] = []

    class SpecialDescriptor:
        def __init__(self, name: str) -> None:
            self.name = name

        def __get__(self, instance: object, _owner: type[object]) -> object:
            events.append(f"bind-{self.name}")
            if self.name == "enter":
                return lambda: instance
            return lambda *_args: events.append("exit") or False

    class DescriptorReader:
        __enter__ = SpecialDescriptor("enter")
        __exit__ = SpecialDescriptor("exit")

        def read_table(self) -> pa.Table:
            return pa.table({"value": [1]})

    result = FallbackCoordinator(_is_compatibility_failure).materialize(
        DescriptorReader,
        pytest.fail,
    )

    assert result.to_pydict() == {"value": [1]}
    assert events == ["bind-exit", "bind-enter", "exit"]


def test_present_but_noncallable_context_special_methods_are_rejected() -> None:
    class InvalidContextReader:
        __enter__ = None
        __exit__ = None

        def read_table(self) -> pa.Table:
            pytest.fail("invalid context protocol must fail before reading")

    with pytest.raises(TypeError, match="incomplete context manager protocol"):
        FallbackCoordinator(_is_compatibility_failure).materialize(
            InvalidContextReader,
            pytest.fail,
        )


def test_context_exit_is_bound_before_enter_can_mutate_its_class() -> None:
    events: list[str] = []

    class MutatingExitReader:
        def __enter__(self) -> MutatingExitReader:
            type(self).__exit__ = replacement_exit
            return self

        def read_table(self) -> pa.Table:
            return pa.table({"value": [1]})

        def __exit__(self, *_args: object) -> bool:
            events.append("original-exit")
            return False

    def replacement_exit(_self: object, *_args: object) -> bool:
        events.append("replacement-exit")
        return False

    original_exit = MutatingExitReader.__exit__
    try:
        FallbackCoordinator(_is_compatibility_failure).materialize(
            MutatingExitReader,
            pytest.fail,
        )
    finally:
        MutatingExitReader.__exit__ = original_exit

    assert events == ["original-exit"]


def test_materialized_exit_suppression_is_a_success_without_fallback() -> None:
    metrics = ParseMetrics()
    reader = _MaterializedReaderFake(
        [],
        "primary",
        read_error=_CompatibilityFailure("suppressed"),
        suppress=True,
    )

    result = FallbackCoordinator(
        _is_compatibility_failure,
        metrics=metrics,
    ).materialize(lambda: reader, pytest.fail)

    assert result is None
    assert reader.close_calls == 1
    assert reader.exit_triples[0][0] is _CompatibilityFailure
    assert reader.exit_triples[0][1] is reader.read_error
    assert metrics == ParseMetrics(full_materializations=1)


@pytest.mark.parametrize(
    "error",
    [
        MemoryError("capacity"),
        KeyboardInterrupt("interrupt"),
        SystemExit("exit"),
        _FatalCleanup("fatal"),
    ],
)
def test_materialized_exit_cannot_suppress_process_level_failure(
    error: BaseException,
) -> None:
    metrics = ParseMetrics()
    reader = _MaterializedReaderFake(
        [],
        "primary",
        read_error=error,
        suppress=True,
    )
    fallback_calls = 0

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    with pytest.raises(type(error)) as captured:
        FallbackCoordinator(
            _is_compatibility_failure,
            metrics=metrics,
        ).materialize(lambda: reader, fallback_factory)

    assert captured.value is error
    assert reader.close_calls == 1
    assert fallback_calls == 0
    assert metrics == ParseMetrics(failed_attempts=1)


def test_process_level_exit_result_is_not_truth_tested() -> None:
    error = MemoryError("capacity")
    exit_result = _ExplodingTruth()
    reader = _MaterializedReaderFake(
        [],
        "primary",
        read_error=error,
        suppress=exit_result,
    )

    with pytest.raises(MemoryError) as captured:
        FallbackCoordinator(_is_compatibility_failure).materialize(
            lambda: reader,
            pytest.fail,
        )

    assert captured.value is error
    assert exit_result.calls == 0


@pytest.mark.parametrize(
    "cleanup_error",
    [
        MemoryError("secret-memory"),
        KeyboardInterrupt("secret-key"),
        SystemExit("secret-exit"),
        _FatalCleanup("secret-fatal"),
    ],
)
def test_process_level_cleanup_failure_wins_over_materialized_operation_error(
    cleanup_error: BaseException,
) -> None:
    operation_error = _CompatibilityFailure("secret-operation")
    metrics = ParseMetrics()
    reader = _MaterializedReaderFake(
        [],
        "primary",
        read_error=operation_error,
        close_error=cleanup_error,
    )

    with pytest.raises(type(cleanup_error)) as captured:
        FallbackCoordinator(
            _is_compatibility_failure,
            metrics=metrics,
        ).materialize(lambda: reader, pytest.fail)

    assert captured.value is cleanup_error
    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "operation_failure": {"type": "_CompatibilityFailure"},
        "cleanup_failure": {"type": type(cleanup_error).__name__},
    }
    assert "secret" not in repr(captured.value.__notes__)
    assert "secret" not in repr(captured.value.backend_context)  # type: ignore[attr-defined]
    assert metrics == ParseMetrics(failed_attempts=1)


@pytest.mark.parametrize(
    "cleanup_error",
    [
        OSError("secret-os"),
        MemoryError("secret-memory"),
        KeyboardInterrupt("secret-key"),
        SystemExit("secret-exit"),
        _FatalCleanup("secret-fatal"),
    ],
)
def test_cleanup_failure_without_an_operation_error_wins_and_is_sanitized(
    cleanup_error: BaseException,
) -> None:
    metrics = ParseMetrics()
    reader = _MaterializedReaderFake([], "primary", close_error=cleanup_error)

    with pytest.raises(type(cleanup_error)) as captured:
        FallbackCoordinator(
            _is_compatibility_failure,
            metrics=metrics,
        ).materialize(lambda: reader, pytest.fail)

    assert captured.value is cleanup_error
    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "cleanup_failure": {"type": type(cleanup_error).__name__}
    }
    assert "secret" not in repr(getattr(captured.value, "__notes__", []))
    assert "secret" not in repr(captured.value.backend_context)  # type: ignore[attr-defined]
    assert metrics == ParseMetrics(failed_attempts=1)


@pytest.mark.parametrize(
    "classifier_error",
    [
        RuntimeError("secret-runtime"),
        MemoryError("secret-memory"),
        KeyboardInterrupt("secret-key"),
        SystemExit("secret-exit"),
    ],
)
def test_classifier_failure_propagates_exactly_after_primary_cleanup(
    classifier_error: BaseException,
) -> None:
    metrics = ParseMetrics()
    reader = _MaterializedReaderFake(
        [],
        "primary",
        read_error=_CompatibilityFailure("secret-primary"),
    )
    fallback_calls = 0

    def classifier(_error: Exception) -> bool:
        raise classifier_error

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    with pytest.raises(type(classifier_error)) as captured:
        FallbackCoordinator(classifier, metrics=metrics).materialize(
            lambda: reader,
            fallback_factory,
        )

    assert captured.value is classifier_error
    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "primary_failure": {"type": "_CompatibilityFailure"},
        "classifier_failure": {"type": type(classifier_error).__name__},
    }
    assert "secret" not in repr(captured.value.__notes__)
    assert fallback_calls == 0
    assert reader.close_calls == 1
    assert metrics == ParseMetrics(failed_attempts=1)


def test_streaming_classifier_failure_propagates_after_exact_once_cleanup() -> None:
    classifier_error = RuntimeError("secret-classifier")
    reader = _StreamingReaderFake(
        [],
        "primary",
        [_CompatibilityFailure("secret-primary")],
    )
    metrics = ParseMetrics()
    fallback_calls = 0

    def classifier(_error: Exception) -> bool:
        raise classifier_error

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    with pytest.raises(RuntimeError) as captured:
        list(
            FallbackCoordinator(classifier, metrics=metrics).batches(
                lambda: reader,
                fallback_factory,
            )
        )

    assert captured.value is classifier_error
    assert reader.close_calls == 1
    assert fallback_calls == 0
    assert metrics == ParseMetrics(failed_attempts=1)
    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "primary_failure": {"type": "_CompatibilityFailure"},
        "classifier_failure": {"type": "RuntimeError"},
    }
    assert "secret" not in repr(captured.value.__notes__)


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
def test_excluded_errors_are_never_sent_to_the_classifier(error: BaseException) -> None:
    def classifier(_error: Exception) -> bool:
        pytest.fail("excluded failure reached classifier")

    with pytest.raises(type(error)):
        FallbackCoordinator(classifier).materialize(
            lambda: _MaterializedReaderFake([], "primary", read_error=error),
            pytest.fail,
        )


def test_fallback_cleanup_context_merges_without_leaking_messages() -> None:
    primary_error = _CompatibilityFailure("secret-primary")
    fallback_error = RuntimeError("secret-fallback")
    cleanup_error = OSError("secret-cleanup")

    with pytest.raises(RuntimeError) as captured:
        FallbackCoordinator(_is_compatibility_failure).materialize(
            lambda: _MaterializedReaderFake([], "primary", read_error=primary_error),
            lambda: _MaterializedReaderFake(
                [],
                "fallback",
                read_error=fallback_error,
                close_error=cleanup_error,
            ),
        )

    assert captured.value is fallback_error
    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "cleanup_failure": {"type": "OSError"},
        "primary_failure": {"type": "_CompatibilityFailure"},
        "fallback_failure": {"type": "RuntimeError"},
    }
    assert "secret" not in repr(captured.value.__notes__)
    assert "secret" not in repr(captured.value.backend_context)  # type: ignore[attr-defined]


def test_fallback_exit_only_failure_preserves_primary_and_cleanup_context() -> None:
    cleanup_error = OSError("secret-cleanup")

    with pytest.raises(OSError) as captured:
        FallbackCoordinator(_is_compatibility_failure).materialize(
            lambda: _MaterializedReaderFake(
                [],
                "primary",
                read_error=_CompatibilityFailure("secret-primary"),
            ),
            lambda: _MaterializedReaderFake(
                [],
                "fallback",
                close_error=cleanup_error,
            ),
        )

    assert captured.value is cleanup_error
    assert captured.value.backend_context == {  # type: ignore[attr-defined]
        "cleanup_failure": {"type": "OSError"},
        "primary_failure": {"type": "_CompatibilityFailure"},
        "fallback_failure": {"type": "OSError"},
    }
    assert "secret" not in repr(captured.value.__notes__)


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


def test_schema_declaration_probe_failure_exits_entered_reader_exactly_once() -> None:
    reader = _RaisingSchemaDiscoveryReader()
    metrics = ParseMetrics()
    fallback_calls = 0

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    with pytest.raises(RuntimeError, match="schema discovery failed"):
        list(
            FallbackCoordinator(
                _is_compatibility_failure,
                metrics=metrics,
            ).batches(lambda: reader, fallback_factory)
        )

    assert reader.enter_calls == 1
    assert len(reader.exit_triples) == 1
    assert reader.exit_triples[0][0] is RuntimeError
    assert fallback_calls == 0
    assert metrics == ParseMetrics(failed_attempts=1)


@pytest.mark.parametrize("failure_point", ["schema", "read"])
def test_streaming_exit_suppression_ends_cleanly_without_fallback(
    failure_point: str,
) -> None:
    error = _CompatibilityFailure("suppressed")
    metrics = ParseMetrics()
    reader = _ContextStreamingReaderFake(
        [],
        "primary",
        [error] if failure_point == "read" else [],
        schema_error=error if failure_point == "schema" else None,
        suppress=True,
    )

    batches = list(
        FallbackCoordinator(
            _is_compatibility_failure,
            metrics=metrics,
        ).batches(lambda: reader, pytest.fail)
    )

    assert batches == []
    assert reader.enter_calls == reader.close_calls == 1
    assert len(reader.exit_triples) == 1
    assert reader.exit_triples[0][0] is _CompatibilityFailure
    assert reader.exit_triples[0][1] is error
    assert metrics == ParseMetrics(streaming_passes=1)


@pytest.mark.parametrize("failure_point", ["schema", "read"])
@pytest.mark.parametrize(
    "error",
    [
        MemoryError("capacity"),
        KeyboardInterrupt("interrupt"),
        SystemExit("exit"),
        _FatalCleanup("fatal"),
    ],
)
def test_streaming_exit_cannot_suppress_process_level_failure(
    error: BaseException,
    failure_point: str,
) -> None:
    metrics = ParseMetrics()
    reader = _ContextStreamingReaderFake(
        [],
        "primary",
        [error] if failure_point == "read" else [],
        schema_error=error if failure_point == "schema" else None,
        suppress=True,
    )
    fallback_calls = 0

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    with pytest.raises(type(error)) as captured:
        list(
            FallbackCoordinator(
                _is_compatibility_failure,
                metrics=metrics,
            ).batches(lambda: reader, fallback_factory)
        )

    assert captured.value is error
    assert reader.close_calls == 1
    assert len(reader.exit_triples) == 1
    assert fallback_calls == 0
    assert metrics == ParseMetrics(failed_attempts=1)


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
        close_error=OSError("secret cursor restore path"),
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
    assert "secret" not in repr(captured.value.__notes__)


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


def test_context_streaming_exhaustion_receives_the_none_exit_triple_once() -> None:
    reader = _ContextStreamingReaderFake([], "primary", [_batch(1), None])

    assert (
        len(
            list(
                FallbackCoordinator(_is_compatibility_failure).batches(lambda: reader, pytest.fail)
            )
        )
        == 1
    )

    assert reader.close_calls == 1
    assert reader.exit_triples == [(None, None, None)]


def test_context_streaming_early_close_receives_generator_exit_without_pass_metric() -> None:
    metrics = ParseMetrics()
    reader = _ContextStreamingReaderFake([], "primary", [_batch(1), None])
    stream = FallbackCoordinator(
        _is_compatibility_failure,
        metrics=metrics,
    ).batches(lambda: reader, pytest.fail)

    assert next(stream).to_pydict() == {"value": [1]}
    stream.close()

    assert reader.close_calls == 1
    assert len(reader.exit_triples) == 1
    assert reader.exit_triples[0][0] is GeneratorExit
    assert isinstance(reader.exit_triples[0][1], GeneratorExit)
    assert metrics == ParseMetrics()


def test_context_streaming_late_failure_receives_exact_triple_and_never_retries() -> None:
    error = _CompatibilityFailure("late")
    reader = _ContextStreamingReaderFake([], "primary", [_batch(1), error])
    fallback_calls = 0

    def fallback_factory() -> object:
        nonlocal fallback_calls
        fallback_calls += 1
        return object()

    stream = FallbackCoordinator(_is_compatibility_failure).batches(
        lambda: reader,
        fallback_factory,
    )
    assert next(stream).to_pydict() == {"value": [1]}
    with pytest.raises(_CompatibilityFailure) as captured:
        next(stream)

    assert captured.value is error
    assert reader.close_calls == 1
    assert len(reader.exit_triples) == 1
    error_type, received_error, _traceback = reader.exit_triples[0]
    assert error_type is _CompatibilityFailure
    assert received_error is error
    assert reader.exit_traceback_matches_error == [True]
    assert fallback_calls == 0


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
