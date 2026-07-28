"""CSV execution routing and decision contracts."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import pyarrow as pa
import pytest

from messy_xlsx import MessyWorkbook, SheetConfig
from messy_xlsx.parsing.handler_registry import HandlerRegistry


@pytest.mark.parametrize(
    ("production_ready", "disable_native", "expected_reason"),
    [
        (False, None, "production_gate_disabled"),
        (False, "1", "production_gate_disabled"),
        (True, None, None),
        (True, "0", None),
        (True, "true", None),
        (True, "1", "kill_switch"),
    ],
)
def test_csv_capability_reason_uses_gate_then_exact_kill_switch(
    production_ready: bool,
    disable_native: str | None,
    expected_reason: str | None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from messy_xlsx.parsing import csv_native
    from messy_xlsx.parsing.csv_contracts import CSVExecutionReason

    monkeypatch.setattr(
        csv_native,
        "_NATIVE_CSV_PRODUCTION_READY",
        production_ready,
    )
    if disable_native is None:
        monkeypatch.delenv("MESSY_XLSX_DISABLE_NATIVE", raising=False)
    else:
        monkeypatch.setenv("MESSY_XLSX_DISABLE_NATIVE", disable_native)

    reason = csv_native.capability_reason()
    if expected_reason is None:
        assert reason is None
    else:
        assert reason is CSVExecutionReason(expected_reason)


def test_record_csv_execution_sequences_last_decision_and_typed_counts() -> None:
    from messy_xlsx.parsing.contracts import ParseMetrics
    from messy_xlsx.parsing.csv_contracts import (
        CSVExecutionDecision,
        CSVExecutionKind,
        CSVExecutionReason,
    )

    metrics = ParseMetrics()

    first = metrics.record_csv_execution(
        CSVExecutionKind.NATIVE,
        CSVExecutionReason.NATIVE_SELECTED,
    )
    second = metrics.record_csv_execution(
        CSVExecutionKind.MATERIALIZED_FALLBACK,
        CSVExecutionReason.PRODUCTION_GATE_DISABLED,
    )
    third = metrics.record_csv_execution(
        CSVExecutionKind.MATERIALIZED_FALLBACK,
        CSVExecutionReason.PRODUCTION_GATE_DISABLED,
    )

    assert first == CSVExecutionDecision(
        operation_id=1,
        kind=CSVExecutionKind.NATIVE,
        reason=CSVExecutionReason.NATIVE_SELECTED,
    )
    assert second.operation_id == 2
    assert third.operation_id == 3
    assert metrics.csv_operation_sequence == 3
    assert metrics.last_csv_execution is third
    assert metrics.csv_execution_counts == {
        (CSVExecutionKind.NATIVE, CSVExecutionReason.NATIVE_SELECTED): 1,
        (
            CSVExecutionKind.MATERIALIZED_FALLBACK,
            CSVExecutionReason.PRODUCTION_GATE_DISABLED,
        ): 2,
    }


@pytest.fixture
def custom_csv_registry() -> HandlerRegistry:
    """Return a registry whose CSV result proves the custom SPI was selected."""

    class CustomCSVRegistry(HandlerRegistry):
        def parse(self, *_args: object, **_kwargs: object) -> pd.DataFrame:
            return pd.DataFrame({"a": [1]})

    return CustomCSVRegistry()


def _assert_csv_execution_unrecorded(workbook: MessyWorkbook) -> None:
    metrics = workbook.parse_metrics
    assert metrics.csv_operation_sequence == 0
    assert metrics.last_csv_execution is None
    assert metrics.csv_execution_counts == {}


def _track_materialized_reader_closes(
    monkeypatch: pytest.MonkeyPatch,
) -> list[object]:
    import messy_xlsx.parsing.materialized_streaming as materialized_streaming

    close_calls: list[object] = []
    real_close = materialized_streaming.PublicSchemaReader.close

    def record_close(reader: object) -> None:
        close_calls.append(reader)
        real_close(reader)  # type: ignore[arg-type]

    monkeypatch.setattr(
        materialized_streaming.PublicSchemaReader,
        "close",
        record_close,
    )
    return close_calls


def test_candidate_public_csv_route_is_materialized(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from messy_xlsx.parsing import csv_native
    from messy_xlsx.parsing.csv_contracts import (
        CSVExecutionDecision,
        CSVExecutionKind,
        CSVExecutionReason,
    )

    monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", False)
    source = tmp_path / "rows.csv"
    source.write_text("a,b\n1,2\n", encoding="utf-8")

    with MessyWorkbook(source) as workbook:
        with workbook.iter_batches() as stream:
            assert pa.Table.from_batches(list(stream)).to_pydict() == {
                "a": [1],
                "b": [2],
            }
        assert workbook.parse_metrics.last_csv_execution == CSVExecutionDecision(
            operation_id=1,
            kind=CSVExecutionKind.MATERIALIZED_FALLBACK,
            reason=CSVExecutionReason.PRODUCTION_GATE_DISABLED,
        )


def test_candidate_reader_construction_failure_records_no_csv_decision(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import messy_xlsx.workbook as workbook_module
    from messy_xlsx.parsing import csv_native

    def fail_reader_construction(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError("candidate reader construction failed")

    monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", False)
    monkeypatch.setattr(
        workbook_module,
        "prepare_materialized_streaming_reader",
        fail_reader_construction,
    )

    with MessyWorkbook(io.BytesIO(b"a\n1\n"), filename="candidate.csv") as workbook:
        with pytest.raises(RuntimeError, match="candidate reader construction failed"):
            workbook.iter_batches(config=SheetConfig(auto_detect=False))
        _assert_csv_execution_unrecorded(workbook)


def test_custom_csv_keeps_custom_backend_and_csv_decision(
    custom_csv_registry: HandlerRegistry,
) -> None:
    from messy_xlsx.parsing.csv_contracts import CSVExecutionKind, CSVExecutionReason

    with MessyWorkbook(
        io.BytesIO(b"a\n1\n"),
        filename="x.csv",
        registry=custom_csv_registry,
    ) as workbook:
        with workbook.iter_batches() as stream:
            list(stream)
        assert workbook.parse_metrics.last_csv_execution.kind is CSVExecutionKind.CUSTOM_SPI
        assert (
            workbook.parse_metrics.last_csv_execution.reason
            is CSVExecutionReason.CUSTOM_SPI
        )


def test_custom_reader_construction_failure_records_no_csv_decision(
    custom_csv_registry: HandlerRegistry,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import messy_xlsx.workbook as workbook_module

    def fail_reader_construction(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError("custom reader construction failed")

    monkeypatch.setattr(
        workbook_module,
        "prepare_materialized_streaming_reader",
        fail_reader_construction,
    )

    with MessyWorkbook(
        io.BytesIO(b"a\n1\n"),
        filename="custom.csv",
        registry=custom_csv_registry,
    ) as workbook:
        with pytest.raises(RuntimeError, match="custom reader construction failed"):
            workbook.iter_batches(config=SheetConfig(auto_detect=False))
        _assert_csv_execution_unrecorded(workbook)


@pytest.mark.parametrize("route", ["candidate", "custom"])
def test_materialized_csv_recording_failure_closes_owned_reader_once(
    route: str,
    custom_csv_registry: HandlerRegistry,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from messy_xlsx.parsing import csv_native
    from messy_xlsx.parsing.contracts import ParseMetrics

    close_calls = _track_materialized_reader_closes(monkeypatch)
    failure = MemoryError(f"{route} CSV metric recording failed")

    def fail_recording(
        _metrics: ParseMetrics,
        _kind: object,
        _reason: object,
    ) -> None:
        raise failure

    monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", False)
    monkeypatch.setattr(ParseMetrics, "record_csv_execution", fail_recording)
    registry = custom_csv_registry if route == "custom" else None

    with MessyWorkbook(
        io.BytesIO(b"a\n1\n"),
        filename=f"{route}.csv",
        registry=registry,
    ) as workbook:
        with pytest.raises(MemoryError) as captured:
            workbook.iter_batches(
                batch_size=1,
                config=SheetConfig(auto_detect=False),
            )
        assert captured.value is failure
        assert len(close_calls) == 1

    assert len(close_calls) == 1


def test_materialized_csv_recording_failure_without_owner_rolls_back_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from messy_xlsx.parsing import csv_native
    from messy_xlsx.parsing.contracts import ParseMetrics

    close_calls = _track_materialized_reader_closes(monkeypatch)
    failure = MemoryError("unowned CSV metric recording failed")

    def fail_recording(
        _metrics: ParseMetrics,
        _kind: object,
        _reason: object,
    ) -> None:
        raise failure

    monkeypatch.setattr(csv_native, "_NATIVE_CSV_PRODUCTION_READY", False)
    monkeypatch.setattr(ParseMetrics, "record_csv_execution", fail_recording)

    with MessyWorkbook(io.BytesIO(b"a\n1\n"), filename="unowned.csv") as workbook:
        with pytest.raises(MemoryError) as captured:
            workbook._prepare_streaming_operation(
                None,
                1,
                SheetConfig(auto_detect=False),
            )
        assert captured.value is failure
        assert len(close_calls) == 1

    assert len(close_calls) == 1
