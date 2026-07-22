"""Native allocation layouts must not be projected as plain Python state."""

from __future__ import annotations

from io import BufferedReader, BytesIO

import pytest

from messy_xlsx import SheetConfig
from messy_xlsx.parsing.parse_plan import compile_parse_plan


class _BufferedReaderValue(BufferedReader):
    def __init__(self, label: str) -> None:
        super().__init__(BytesIO(b"hidden native state"))
        self.label = label

    def __eq__(self, other: object) -> bool:
        return isinstance(other, _BufferedReaderValue) and self.label == other.label

    def __hash__(self) -> int:
        return hash(self.label)


def test_c_backed_base_without_own_allocator_is_rejected_before_thaw() -> None:
    payload = _BufferedReaderValue("visible")
    try:
        with pytest.raises(TypeError, match="opaque mutable configuration value"):
            compile_parse_plan(
                SheetConfig(
                    auto_detect=False,
                    type_hints={"payload": payload},  # type: ignore[dict-item]
                ),
                None,
                "xlsx",
            )
    finally:
        payload.close()
