"""Tests for XLS file handler."""

import io

import pytest

from messy_xlsx.parsing.base_handler import ParseOptions
from messy_xlsx.parsing.xls_handler import XLSHandler

xlrd = pytest.importorskip("xlrd")


def _create_xls(data, sheet_name="Sheet1", header=True):
    """Create an XLS file in memory using xlwt."""
    xlwt = pytest.importorskip("xlwt")
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)

    start_row = 0
    if header and data:
        for col_idx, val in enumerate(data[0]):
            ws.write(0, col_idx, val)
        start_row = 1
        rows = data[1:]
    else:
        rows = data

    for row_idx, row in enumerate(rows, start=start_row):
        for col_idx, val in enumerate(row):
            ws.write(row_idx, col_idx, val)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _create_xls_file(path, data, sheet_name="Sheet1", header=True):
    """Create an XLS file on disk."""
    xlwt = pytest.importorskip("xlwt")
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)

    start_row = 0
    if header and data:
        for col_idx, val in enumerate(data[0]):
            ws.write(0, col_idx, val)
        start_row = 1
        rows = data[1:]
    else:
        rows = data

    for row_idx, row in enumerate(rows, start=start_row):
        for col_idx, val in enumerate(row):
            ws.write(row_idx, col_idx, val)

    wb.save(str(path))


class TestXLSHandler:
    """Test XLS format handler."""

    def setup_method(self):
        self.handler = XLSHandler()

    def test_can_handle(self):
        assert self.handler.can_handle("xls") is True
        assert self.handler.can_handle("xlsx") is False
        assert self.handler.can_handle("csv") is False

    def test_basic_parse(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        data = [
            ["Name", "Age", "City"],
            ["Alice", 30, "NY"],
            ["Bob", 25, "LA"],
        ]
        _create_xls_file(path, data)

        options = ParseOptions(header_rows=1, skip_rows=0)
        df = self.handler.parse(path, "Sheet1", options)
        assert len(df) == 2
        assert list(df.columns) == ["Name", "Age", "City"]

    def test_skip_rows(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        # Row 0: metadata, Row 1: header, Rows 2+: data
        data = [
            ["Report Title", "", ""],
            ["Name", "Age", "City"],
            ["Alice", 30, "NY"],
            ["Bob", 25, "LA"],
        ]
        _create_xls_file(path, data, header=False)

        options = ParseOptions(header_rows=1, skip_rows=1)
        df = self.handler.parse(path, "Sheet1", options)
        assert len(df) == 2
        assert "Name" in df.columns

    def test_skip_footer(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        data = [
            ["Name", "Value"],
            ["A", 10],
            ["B", 20],
            ["Total", 30],
        ]
        _create_xls_file(path, data)

        options = ParseOptions(header_rows=1, skip_footer=1)
        df = self.handler.parse(path, "Sheet1", options)
        assert len(df) == 2

    def test_max_rows(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        data = [
            ["Name", "Value"],
            ["A", 1],
            ["B", 2],
            ["C", 3],
            ["D", 4],
        ]
        _create_xls_file(path, data)

        options = ParseOptions(header_rows=1, max_rows=2)
        df = self.handler.parse(path, "Sheet1", options)
        assert len(df) == 2

    def test_header_rows_zero(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        data = [
            ["A", 1],
            ["B", 2],
        ]
        _create_xls_file(path, data, header=False)

        options = ParseOptions(header_rows=0)
        df = self.handler.parse(path, "Sheet1", options)
        assert all(c.startswith("col_") for c in df.columns)

    def test_multi_row_headers(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        data = [
            ["Category", "Sales", "Sales"],
            ["", "Q1", "Q2"],
            ["Product A", 100, 200],
            ["Product B", 300, 400],
        ]
        _create_xls_file(path, data, header=False)

        options = ParseOptions(header_rows=2)
        df = self.handler.parse(path, "Sheet1", options)
        assert len(df) == 2

    def test_get_sheet_names(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        _create_xls_file(path, [["A"], [1]], sheet_name="MySheet")

        names = self.handler.get_sheet_names(path)
        assert "MySheet" in names

    def test_validate(self, temp_dir):
        pytest.importorskip("xlwt")
        path = temp_dir / "test.xls"
        _create_xls_file(path, [["A"], [1]])

        is_valid, error = self.handler.validate(path)
        assert is_valid is True
        assert error is None

    def test_file_like_object(self):
        pytest.importorskip("xlwt")
        data = [
            ["Name", "Value"],
            ["A", 1],
            ["B", 2],
        ]
        buf = _create_xls(data)

        options = ParseOptions(header_rows=1)
        df = self.handler.parse(buf, None, options)
        assert len(df) == 2
