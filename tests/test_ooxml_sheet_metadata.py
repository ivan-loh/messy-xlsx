"""Tests for lazy, bounded worksheet metadata indexing."""

from __future__ import annotations

from dataclasses import FrozenInstanceError

import openpyxl
import pytest

from messy_xlsx._source import SourceHandle
from messy_xlsx.ooxml.manifest import ManifestReader
from messy_xlsx.ooxml.models import Interval, IntervalIndex, MergeRange


def test_interval_index_normalizes_ranges_without_expanding_cells() -> None:
    index = IntervalIndex(
        (
            Interval(10, 12),
            Interval(2, 4),
            Interval(4, 9),
            Interval(20, 20),
        )
    )

    assert index.intervals == (Interval(2, 12), Interval(20, 20))
    assert index.contains(2)
    assert index.contains(11)
    assert not index.contains(13)
    assert index.contains(20)
    assert len(index.intervals) == 2


@pytest.mark.parametrize("start,end", [(0, 1), (-1, 3), (3, 2)])
def test_interval_rejects_invalid_one_based_bounds(start: int, end: int) -> None:
    with pytest.raises(ValueError, match="invalid one-based interval"):
        Interval(start, end)


def test_interval_models_are_immutable() -> None:
    interval = Interval(2, 4)

    with pytest.raises(FrozenInstanceError):
        interval.start = 3  # type: ignore[misc]


@pytest.fixture
def metadata_xlsx(tmp_path):
    path = tmp_path / "metadata.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "anchor"
    worksheet["D9"] = "tail"
    worksheet["D9"].number_format = "#.##0,00"
    worksheet.merge_cells("B2:C3")
    worksheet.row_dimensions[2].hidden = True
    worksheet.row_dimensions[3].hidden = True
    worksheet.row_dimensions[5].hidden = True
    worksheet.column_dimensions.group("B", "D", hidden=True)
    worksheet.column_dimensions["F"].hidden = True
    second = workbook.create_sheet("Other")
    second["C4"] = "other"
    workbook.save(path)
    workbook.close()
    return path


def test_sheet_xml_is_loaded_once_and_lazily_per_sheet(metadata_xlsx) -> None:
    opened: list[str] = []
    with SourceHandle(metadata_xlsx) as source:
        reader = ManifestReader(source, on_member_open=opened.append)

        assert not any(name.startswith("xl/worksheets/") for name in opened)
        first = reader.sheet("Data")
        assert reader.sheet("Data") is first
        assert sum(name.startswith("xl/worksheets/") for name in opened) == 1

        reader.sheet("Other")
        assert sum(name.startswith("xl/worksheets/") for name in opened) == 2


def test_sheet_metadata_compacts_hidden_ranges_and_indexes_dimensions(metadata_xlsx) -> None:
    with SourceHandle(metadata_xlsx) as source:
        sheet = ManifestReader(source).sheet("Data")

    assert sheet.name == "Data"
    assert sheet.declared_dimension == (1, 1, 9, 4)
    assert (sheet.observed_max_row, sheet.observed_max_col) == (9, 4)
    assert sheet.hidden_rows.intervals == (Interval(2, 3), Interval(5, 5))
    assert sheet.hidden_columns.intervals == (Interval(2, 4), Interval(6, 6))
    assert sheet.merged_ranges == (MergeRange(2, 2, 3, 3),)
    assert sheet.number_format_codes == ("#.##0,00",)
    assert not hasattr(sheet, "values")


def test_formula_presence_is_exact_but_coordinate_samples_are_capped(tmp_path) -> None:
    path = tmp_path / "formulas.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    for row in range(1, 301):
        worksheet.cell(row=row, column=8, value=f"=A{row}+1")
    workbook.save(path)
    workbook.close()

    with SourceHandle(path) as source:
        sheet = ManifestReader(source).sheet("Data")

    assert sheet.has_formulas is True
    assert len(sheet.formula_samples) == 256
    assert sheet.formula_samples[0] == "H1"
    assert sheet.formula_samples[-1] == "H256"
    assert (sheet.observed_max_row, sheet.observed_max_col) == (300, 8)
