"""CSV execution routing and decision contracts."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import pyarrow as pa
import pytest

from messy_xlsx import MessyWorkbook
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
