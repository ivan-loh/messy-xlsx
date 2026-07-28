"""Source-controlled gate for the candidate native CSV backend."""

from __future__ import annotations

import os
from typing import Final

from messy_xlsx.parsing.csv_contracts import CSVExecutionReason

_NATIVE_CSV_PRODUCTION_READY: Final[bool] = False


def capability_reason() -> CSVExecutionReason | None:
    """Return why the candidate native CSV backend cannot be selected."""
    if not _NATIVE_CSV_PRODUCTION_READY:
        return CSVExecutionReason.PRODUCTION_GATE_DISABLED
    if os.environ.get("MESSY_XLSX_DISABLE_NATIVE") == "1":
        return CSVExecutionReason.KILL_SWITCH
    return None
