"""Focused regressions for retryable materialized-operation leases."""

from __future__ import annotations

import sys
from pathlib import Path
from types import FrameType

import pandas as pd
import pytest

import messy_xlsx
import messy_xlsx.workbook as workbook_module


@pytest.mark.parametrize("entry", ["to_arrow", "_to_dataframes_compat", "_parse_sheet"])
def test_body_start_return_interruption_does_not_poison_next_materialized_operation(
    sample_xlsx: Path,
    monkeypatch: pytest.MonkeyPatch,
    entry: str,
) -> None:
    workbook = messy_xlsx.MessyWorkbook(sample_xlsx)
    frame = pd.DataFrame({"value": [1]})
    monkeypatch.setattr(
        workbook,
        "_parse_sheet_unreserved",
        lambda *_args, **_kwargs: frame,
    )
    real_end = workbook._end_operation
    end_calls = 0

    def fail_twice(token: object) -> None:
        nonlocal end_calls
        end_calls += 1
        if end_calls <= 2:
            raise MemoryError("materialized release interrupted")
        real_end(token)

    target_code = workbook_module._MaterializedOperationLease._body_started.__code__
    interrupted = False

    def interrupt_body_start_return(
        frame: FrameType,
        event: str,
        _arg: object,
    ) -> object:
        nonlocal interrupted
        if frame.f_code is target_code and event == "return" and not interrupted:
            interrupted = True
            sys.settrace(None)
            frame.f_trace = None
            raise MemoryError("materialized body start return interrupted")
        return interrupt_body_start_return

    monkeypatch.setattr(workbook, "_end_operation", fail_twice)
    sys.settrace(interrupt_body_start_return)
    try:
        with pytest.raises(
            MemoryError,
            match="materialized body start return interrupted",
        ):
            if entry == "to_arrow":
                workbook.to_arrow()
            elif entry == "_to_dataframes_compat":
                workbook._to_dataframes_compat()
            else:
                workbook._parse_sheet("Data")
    finally:
        sys.settrace(None)

    try:
        assert interrupted
        assert end_calls == 2
        retained_lease = workbook._active_materialized_lease
        assert retained_lease is not None
        assert retained_lease._body_active is False
        assert workbook._active_operation_token is retained_lease._token

        assert workbook.to_arrow().to_pydict() == {"value": [1]}
        assert end_calls == 4
        assert workbook._active_operation_token is None
        assert workbook._active_materialized_lease is None
    finally:
        monkeypatch.setattr(workbook, "_end_operation", real_end)
        workbook.close()


@pytest.mark.parametrize("entry", ["to_arrow", "_to_dataframes_compat", "_parse_sheet"])
def test_body_complete_entry_interruption_leaves_lease_retryable(
    sample_xlsx: Path,
    monkeypatch: pytest.MonkeyPatch,
    entry: str,
) -> None:
    workbook = messy_xlsx.MessyWorkbook(sample_xlsx)
    frame = pd.DataFrame({"value": [1]})
    monkeypatch.setattr(
        workbook,
        "_parse_sheet_unreserved",
        lambda *_args, **_kwargs: frame,
    )
    real_end = workbook._end_operation
    end_calls = 0

    def fail_twice(token: object) -> None:
        nonlocal end_calls
        end_calls += 1
        if end_calls <= 2:
            raise MemoryError("materialized release interrupted")
        real_end(token)

    target_code = workbook_module._MaterializedOperationLease._body_complete.__code__
    interrupted = False

    def interrupt_body_complete_entry(
        frame: FrameType,
        event: str,
        _arg: object,
    ) -> object:
        nonlocal interrupted
        if frame.f_code is target_code and event == "call" and not interrupted:
            interrupted = True
            sys.settrace(None)
            frame.f_trace = None
            raise MemoryError("materialized body completion interrupted")
        return interrupt_body_complete_entry

    monkeypatch.setattr(workbook, "_end_operation", fail_twice)
    sys.settrace(interrupt_body_complete_entry)
    try:
        with pytest.raises(
            MemoryError,
            match="materialized body completion interrupted",
        ):
            if entry == "to_arrow":
                workbook.to_arrow()
            elif entry == "_to_dataframes_compat":
                workbook._to_dataframes_compat()
            else:
                workbook._parse_sheet("Data")
    finally:
        sys.settrace(None)

    try:
        assert interrupted
        assert end_calls == 2
        retained_lease = workbook._active_materialized_lease
        assert retained_lease is not None
        assert retained_lease._body_active is False
        assert workbook._active_operation_token is retained_lease._token

        assert workbook.to_arrow().to_pydict() == {"value": [1]}
        assert end_calls == 4
        assert workbook._active_operation_token is None
        assert workbook._active_materialized_lease is None
    finally:
        monkeypatch.setattr(workbook, "_end_operation", real_end)
        workbook.close()
