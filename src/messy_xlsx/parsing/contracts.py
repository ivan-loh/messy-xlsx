"""Internal contracts shared by parser backends and orchestration."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from typing import TYPE_CHECKING, Protocol

import pyarrow as pa

if TYPE_CHECKING:
    from messy_xlsx.parsing.parse_plan import ParsePlan


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


class MaterializedArrowReader(Protocol):
    """A backend that produces one complete Arrow table."""

    def read_table(self, plan: ParsePlan) -> pa.Table:
        """Materialize the requested sheet according to *plan*."""
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
