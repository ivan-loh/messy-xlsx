"""Internal contracts shared by parser backends and orchestration."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Protocol

import pyarrow as pa

from messy_xlsx.parsing.csv_contracts import (
    CSVExecutionDecision,
    CSVExecutionKind,
    CSVExecutionReason,
)


class OutputMode(StrEnum):
    """Internal output representation requested from a parser backend."""

    MATERIALIZED = "materialized"
    STREAMING = "streaming"


class BackendKind(StrEnum):
    """Backend capability selected for one parse."""

    FASTEXCEL = "fastexcel"
    OPENPYXL_COMPAT = "openpyxl_compat"
    OPENPYXL_STREAMING = "openpyxl_streaming"
    CSV_STREAMING = "csv_streaming"
    XLS_STREAMING = "xls_streaming"
    CUSTOM_DATAFRAME = "custom_dataframe"


@dataclass(frozen=True, slots=True)
class ReaderDecision:
    """Selected backend plus a stable diagnostic explanation."""

    backend: BackendKind
    reason: str


@dataclass(slots=True)
class ParseMetrics:
    """Workbook-scoped counters for bounded and full-value parser work."""

    manifest_builds: int = 0
    sample_reads: int = 0
    full_materializations: int = 0
    streaming_passes: int = 0
    failed_attempts: int = 0
    csv_operation_sequence: int = 0
    last_csv_execution: CSVExecutionDecision | None = None
    csv_execution_counts: dict[
        tuple[CSVExecutionKind, CSVExecutionReason],
        int,
    ] = field(default_factory=dict)

    def record_csv_execution(
        self,
        kind: CSVExecutionKind,
        reason: CSVExecutionReason,
    ) -> CSVExecutionDecision:
        """Record and return one CSV execution decision."""
        self.csv_operation_sequence += 1
        decision = CSVExecutionDecision(self.csv_operation_sequence, kind, reason)
        self.last_csv_execution = decision
        key = (kind, reason)
        self.csv_execution_counts[key] = self.csv_execution_counts.get(key, 0) + 1
        return decision


class MaterializedArrowReader(Protocol):
    """A backend that produces one complete Arrow table."""

    def read_table(self) -> pa.Table:
        """Materialize the requested sheet using the bound parse plan."""
        ...


class StreamingBatchReader(Protocol):
    """A backend that incrementally produces schema-stable Arrow batches."""

    @property
    def schema(self) -> pa.Schema:
        """Return the stable schema for this reader."""
        ...

    def read_next_batch(self) -> pa.RecordBatch | None:
        """Return the next batch, or ``None`` after exhaustion."""
        ...

    def close(self) -> None:
        """Release backend and source resources deterministically."""
        ...
