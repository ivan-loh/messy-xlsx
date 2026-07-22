"""Tests for lazy, bounded worksheet metadata indexing."""

from __future__ import annotations

import os
from collections.abc import Callable
from dataclasses import FrozenInstanceError
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
import pytest

from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.manifest import ManifestReader
from messy_xlsx.ooxml.models import Interval, IntervalIndex, MergeRange


def _rewrite_member(path: Path, member: str, transform: Callable[[bytes], bytes]) -> None:
    replacement = path.with_suffix(".replacement.xlsx")
    with ZipFile(path) as source, ZipFile(replacement, "w", ZIP_DEFLATED) as target:
        for info in source.infolist():
            content = source.read(info.filename)
            target.writestr(info, transform(content) if info.filename == member else content)
    replacement.replace(path)


def _replace_required(content: bytes, old: bytes, new: bytes) -> bytes:
    assert old in content
    return content.replace(old, new)


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


def test_excel_upper_coordinate_is_accepted(tmp_path) -> None:
    path = tmp_path / "upper-edge.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "value"
    workbook.save(path)
    workbook.close()
    _rewrite_member(
        path,
        "xl/worksheets/sheet1.xml",
        lambda xml: _replace_required(
            _replace_required(xml, b'ref="A1:A1"', b'ref="XFD1048576:XFD1048576"'),
            b'r="A1"',
            b'r="XFD1048576"',
        ),
    )

    with SourceHandle(path) as source:
        sheet = ManifestReader(source).sheet("Sheet")

    assert sheet.declared_dimension == (1_048_576, 16_384, 1_048_576, 16_384)
    assert (sheet.observed_max_row, sheet.observed_max_col) == (1_048_576, 16_384)


@pytest.mark.parametrize(
    ("setup", "old", "new", "coordinate"),
    [
        ("cell", b'r="A1"', b'r="XFE1"', "XFE1"),
        ("cell", b'r="A1"', b'r="A1048577"', "A1048577"),
        ("dimension", b'ref="A1:A1"', b'ref="A1:XFE1"', "A1:XFE1"),
        ("merge", b'ref="A1:B1"', b'ref="XFE1:XFE1"', "XFE1:XFE1"),
        ("hidden_row", b'r="1" hidden="1"', b'r="1048577" hidden="1"', "1048577"),
        (
            "hidden_col",
            b'customWidth="1" min="1" max="1"',
            b'customWidth="1" min="16385" max="16385"',
            "16385",
        ),
    ],
)
def test_out_of_bounds_coordinates_are_rejected_with_member_context(
    tmp_path: Path,
    setup: str,
    old: bytes,
    new: bytes,
    coordinate: str,
) -> None:
    path = tmp_path / f"{setup}.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet["A1"] = "value"
    if setup == "merge":
        worksheet.merge_cells("A1:B1")
    elif setup == "hidden_row":
        worksheet.row_dimensions[1].hidden = True
    elif setup == "hidden_col":
        worksheet.column_dimensions["A"].hidden = True
    workbook.save(path)
    workbook.close()
    _rewrite_member(
        path,
        "xl/worksheets/sheet1.xml",
        lambda xml: _replace_required(xml, old, new),
    )

    with (
        SourceHandle(path) as source,
        pytest.raises(FormatError, match="coordinate") as raised,
    ):
        ManifestReader(source).sheet("Sheet")

    assert raised.value.context["member"] == "xl/worksheets/sheet1.xml"
    assert coordinate in str(raised.value.context)


def test_path_manifest_reader_rejects_equal_size_replacement_with_restored_mtime(
    tmp_path: Path,
) -> None:
    path = tmp_path / "stable.xlsx"
    workbook = openpyxl.Workbook()
    workbook.save(path)
    workbook.close()

    with SourceHandle(path) as source:
        reader = ManifestReader(source)
        original = path.stat()
        replacement = tmp_path / "replacement.xlsx"
        replacement.write_bytes(path.read_bytes())
        os.utime(replacement, ns=(original.st_atime_ns, original.st_mtime_ns))
        replacement.replace(path)
        os.utime(path, ns=(original.st_atime_ns, original.st_mtime_ns))

        with pytest.raises(FormatError, match="changed") as raised:
            reader.sheet("Sheet")

    assert raised.value.context["file_path"] == str(path)


def test_path_manifest_reader_rechecks_identity_after_lazy_parse(tmp_path: Path) -> None:
    path = tmp_path / "race.xlsx"
    workbook = openpyxl.Workbook()
    workbook.save(path)
    workbook.close()

    def replace_after_member_open(_member: str) -> None:
        replacement = tmp_path / "race-replacement.xlsx"
        replacement.write_bytes(path.read_bytes())
        replacement.replace(path)

    with SourceHandle(path) as source:
        reader = ManifestReader(source, on_member_open=replace_after_member_open)
        with pytest.raises(FormatError, match="changed"):
            reader.sheet("Sheet")
