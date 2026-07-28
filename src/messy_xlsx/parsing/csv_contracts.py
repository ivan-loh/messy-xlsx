"""Typed execution contracts for CSV backend selection."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum


class CSVExecutionKind(StrEnum):
    """Observable CSV execution backend."""

    NATIVE = "csv_native"
    MATERIALIZED_FALLBACK = "csv_materialized_fallback"
    CUSTOM_SPI = "custom_dataframe"


class CSVExecutionReason(StrEnum):
    """Stable explanation for one CSV execution decision."""

    NATIVE_SELECTED = "native_selected"
    CUSTOM_SPI = "custom_spi"
    PRODUCTION_GATE_DISABLED = "production_gate_disabled"
    KILL_SWITCH = "kill_switch"
    UNSUPPORTED_RUNTIME = "unsupported_runtime"
    IMPORT_OR_LOAD_FAILURE = "import_or_load_failure"
    HANDSHAKE_MISMATCH = "handshake_mismatch"
    EVIDENCE_BUDGET = "evidence_budget"
    MULTI_HEADER_EXACTNESS = "multi_header_exactness"
    UNSUPPORTED_EVIDENCE_TYPE = "unsupported_evidence_type"


@dataclass(frozen=True, slots=True)
class CSVExecutionDecision:
    """One workbook-scoped CSV execution selection."""

    operation_id: int
    kind: CSVExecutionKind
    reason: CSVExecutionReason
