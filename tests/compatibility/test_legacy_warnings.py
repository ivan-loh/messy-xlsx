import warnings
from types import MethodType

import messy_xlsx as api
import messy_xlsx.multi_sheet as multi_sheet_module
from messy_xlsx import (
    LegacyAPIWarning,
    MessyWorkbook,
    MultiSheetParser,
    read_all_sheets,
    read_excel,
    read_excel_tables,
)
from messy_xlsx.parsing.base_handler import ParseOptions
from messy_xlsx.parsing.handler_registry import HandlerRegistry
from messy_xlsx.parsing.xlsx_handler import XLSXHandler
from messy_xlsx.sheet import MessyTable


def _legacy_records(callable_object):
    with warnings.catch_warnings(record=True) as records:
        warnings.simplefilter("always", LegacyAPIWarning)
        callable_object()
    return [record for record in records if record.category is LegacyAPIWarning]


def test_read_excel_emits_one_caller_facing_warning(sample_xlsx) -> None:
    records = _legacy_records(lambda: read_excel(str(sample_xlsx)))
    assert len(records) == 1
    assert records[0].filename == __file__
    assert str(records[0].message) == (
        "read_excel is a legacy materialized API retained through messy-xlsx v1.x"
    )
    assert issubclass(LegacyAPIWarning, DeprecationWarning)


def test_read_excel_tables_emits_one_caller_facing_warning(sample_xlsx) -> None:
    records = _legacy_records(lambda: read_excel_tables(str(sample_xlsx)))
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_all_sheets_suppresses_nested_parse_all_warning(sample_xlsx) -> None:
    records = _legacy_records(lambda: read_all_sheets(sample_xlsx))
    assert len(records) == 1
    assert records[0].filename == __file__


def test_direct_multi_sheet_methods_warn_once(sample_xlsx) -> None:
    parser = MultiSheetParser(sample_xlsx)
    parse_all_records = _legacy_records(parser.parse_all)
    parse_sheet_records = _legacy_records(lambda: parser.parse_sheet("Data"))

    assert len(parse_all_records) == 1
    assert parse_all_records[0].filename == __file__
    assert len(parse_sheet_records) == 1
    assert parse_sheet_records[0].filename == __file__


def test_direct_workbook_methods_warn_once(sample_xlsx) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        dataframe_records = _legacy_records(workbook.to_dataframe)
        dataframes_records = _legacy_records(workbook.to_dataframes)

    assert len(dataframe_records) == 1
    assert dataframe_records[0].filename == __file__
    assert len(dataframes_records) == 1
    assert dataframes_records[0].filename == __file__


def test_direct_sheet_and_table_methods_warn_once(sample_xlsx) -> None:
    with MessyWorkbook(sample_xlsx) as workbook:
        sheet = workbook.get_sheet("Data")
        table = sheet.tables[0]

        sheet_records = _legacy_records(sheet.to_dataframe)
        table_records = _legacy_records(table.to_dataframe)

    assert len(sheet_records) == 1
    assert sheet_records[0].filename == __file__
    assert len(table_records) == 1
    assert table_records[0].filename == __file__


def test_extension_spi_parse_calls_remain_warning_free(sample_xlsx) -> None:
    registry = HandlerRegistry()
    registry_records = _legacy_records(
        lambda: registry.parse(sample_xlsx, sheet="Data", format_type="xlsx")
    )
    handler_records = _legacy_records(
        lambda: XLSXHandler().parse(sample_xlsx, "Data", ParseOptions())
    )

    assert registry_records == []
    assert handler_records == []


def test_read_excel_honors_override_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    class OverridingWorkbook(MessyWorkbook):
        called = False

        def to_dataframe(self, sheet=None, config=None):
            type(self).called = True
            return super().to_dataframe(sheet, config)

    monkeypatch.setattr(api, "MessyWorkbook", OverridingWorkbook)

    records = _legacy_records(lambda: read_excel(str(sample_xlsx)))

    assert OverridingWorkbook.called is True
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_excel_tables_honors_override_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    original = MessyTable.to_dataframe
    called = False

    def overridden(table, config=None):
        nonlocal called
        called = True
        return original(table, config)

    monkeypatch.setattr(MessyTable, "to_dataframe", overridden)

    records = _legacy_records(lambda: read_excel_tables(str(sample_xlsx)))

    assert called is True
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_all_sheets_honors_override_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    class OverridingParser(MultiSheetParser):
        called = False

        def parse_all(self):
            type(self).called = True
            return super().parse_all()

    monkeypatch.setattr(multi_sheet_module, "MultiSheetParser", OverridingParser)

    records = _legacy_records(lambda: read_all_sheets(sample_xlsx))

    assert OverridingParser.called is True
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_excel_honors_instance_override_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    called = False

    def overridden(workbook, sheet=None, config=None):
        nonlocal called
        called = True
        return MessyWorkbook.to_dataframe(workbook, sheet, config)

    class InstanceOverriddenWorkbook(MessyWorkbook):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.to_dataframe = MethodType(overridden, self)

    monkeypatch.setattr(api, "MessyWorkbook", InstanceOverriddenWorkbook)

    records = _legacy_records(lambda: read_excel(str(sample_xlsx)))

    assert called is True
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_excel_tables_honors_instance_override_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    original_init = MessyTable.__init__
    called = False

    def overridden(table, config=None):
        nonlocal called
        called = True
        return MessyTable.to_dataframe(table, config)

    def instance_overriding_init(table, *args, **kwargs):
        original_init(table, *args, **kwargs)
        table.to_dataframe = MethodType(overridden, table)

    monkeypatch.setattr(MessyTable, "__init__", instance_overriding_init)

    records = _legacy_records(lambda: read_excel_tables(str(sample_xlsx)))

    assert called is True
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_all_sheets_honors_instance_override_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    called = False

    def overridden(parser):
        nonlocal called
        called = True
        return MultiSheetParser.parse_all(parser)

    class InstanceOverriddenParser(MultiSheetParser):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.parse_all = MethodType(overridden, self)

    monkeypatch.setattr(
        multi_sheet_module,
        "MultiSheetParser",
        InstanceOverriddenParser,
    )

    records = _legacy_records(lambda: read_all_sheets(sample_xlsx))

    assert called is True
    assert len(records) == 1
    assert records[0].filename == __file__


def test_read_excel_honors_dynamic_getattribute_dispatch_without_duplicate_warning(
    monkeypatch, sample_xlsx
) -> None:
    called = False

    def dynamic_to_dataframe(workbook, sheet=None, config=None):
        nonlocal called
        called = True
        return MessyWorkbook.to_dataframe(workbook, sheet, config)

    class DynamicWorkbook(MessyWorkbook):
        def __getattribute__(self, name):
            if name == "to_dataframe":
                return MethodType(dynamic_to_dataframe, self)
            return super().__getattribute__(name)

    monkeypatch.setattr(api, "MessyWorkbook", DynamicWorkbook)

    records = _legacy_records(lambda: read_excel(str(sample_xlsx)))

    assert called is True
    assert len(records) == 1
    assert records[0].filename == __file__
