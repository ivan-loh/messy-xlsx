"""One reusable fastexcel workbook context for bounded samples and materialization."""

from __future__ import annotations

from typing import Any

import fastexcel
import pandas as pd

from messy_xlsx.detection.structure_sampler import SampleWindow, StructureEvidence


class FastexcelSession:
    """Own exactly one fastexcel reader and its source backend context."""

    def __init__(self, source: Any) -> None:
        self._context = source.open_path_or_bytes()
        backend = self._context.__enter__()
        try:
            self._reader = fastexcel.read_excel(backend)
            self._sheet_names = tuple(self._reader.sheet_names)
        except BaseException:
            self._context.__exit__(None, None, None)
            raise
        self._closed = False

    @property
    def sheet_names(self) -> tuple[str, ...]:
        """Return workbook sheet names captured from the one reader."""
        return self._sheet_names

    def sample_windows(
        self,
        sheet: str,
        windows: tuple[SampleWindow, ...],
        max_column: int,
    ) -> StructureEvidence:
        """Read bounded windows using integer skip_rows and n_rows only."""
        self._ensure_open()
        frames: list[pd.DataFrame] = []
        rows: list[int] = []
        for window in windows:
            options: dict[str, Any] = {
                "header_row": None,
                "skip_rows": window.start_row - 1,
                "n_rows": window.n_rows,
                "schema_sample_rows": min(1_000, window.n_rows),
                "eager": True,
            }
            if max_column > 0:
                options["use_columns"] = list(range(max_column))
            batch = self._reader.load_sheet(sheet, **options)
            frame = batch.to_pandas()
            frame.index = pd.RangeIndex(
                window.start_row,
                window.start_row + len(frame),
                name="worksheet_row",
            )
            rows.extend(int(row) for row in frame.index)
            frames.append(frame)
        values = pd.concat(frames, axis=0) if frames else pd.DataFrame()
        return StructureEvidence(tuple(rows), values)

    def materialize(self, sheet: str) -> Any:
        """Perform one whole-sheet fastexcel materialization for the caller."""
        self._ensure_open()
        return self._reader.load_sheet(
            sheet,
            header_row=None,
            schema_sample_rows=1_000,
            dtype_coercion="coerce",
            eager=True,
        )

    def close(self) -> None:
        """Close the source backend deterministically and idempotently."""
        if self._closed:
            return
        self._closed = True
        self._reader = None
        self._context.__exit__(None, None, None)

    def _ensure_open(self) -> None:
        if self._closed:
            raise RuntimeError("FastexcelSession is closed")

    def __enter__(self) -> FastexcelSession:
        self._ensure_open()
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        self.close()
