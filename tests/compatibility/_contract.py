from __future__ import annotations

import hashlib
import json
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pandas as pd


def _label(value: object) -> dict[str, str]:
    return {"type": type(value).__qualname__, "repr": repr(value)}


def frame_contract(frame: pd.DataFrame) -> dict[str, Any]:
    normalized = frame.astype(object).where(frame.notna(), None)
    records = normalized.to_dict(orient="split")
    payload = json.dumps(records, default=str, ensure_ascii=False, sort_keys=True)
    return {
        "shape": list(frame.shape),
        "columns": [_label(value) for value in frame.columns],
        "index": [_label(value) for value in frame.index],
        "dtypes": [str(dtype) for dtype in frame.dtypes],
        "value_sha256": hashlib.sha256(payload.encode("utf-8")).hexdigest(),
    }


def exception_contract(callable_object: Callable[[], object]) -> dict[str, Any]:
    try:
        callable_object()
    except Exception as error:
        context = dict(getattr(error, "context", {}))
        message = str(error)
        if "file_path" in context:
            original = str(context["file_path"])
            normalized = Path(original).name
            context["file_path"] = normalized
            message = message.replace(original, normalized)
        return {
            "type": f"{type(error).__module__}.{type(error).__qualname__}",
            "message": message,
            "context": context,
        }
    raise AssertionError("expected callable to raise")
