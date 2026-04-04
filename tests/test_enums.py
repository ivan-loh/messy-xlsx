"""Tests for enum types and backward compatibility."""

import pytest

from messy_xlsx.enums import (
    DataType,
    FormatType,
    HeaderDetectionMode,
    HeaderFallback,
    MergeStrategy,
)
from messy_xlsx.models import CellValue, FormatInfo, SheetConfig


class TestStrEnumBackwardCompat:
    """Ensure StrEnum values compare equal to their raw string counterparts."""

    def test_merge_strategy_equals_string(self) -> None:
        assert MergeStrategy.FILL == "fill"
        assert MergeStrategy.SKIP == "skip"
        assert MergeStrategy.FIRST_ONLY == "first_only"

    def test_header_detection_mode_equals_string(self) -> None:
        assert HeaderDetectionMode.SMART == "smart"
        assert HeaderDetectionMode.AUTO == "auto"
        assert HeaderDetectionMode.MANUAL == "manual"

    def test_header_fallback_equals_string(self) -> None:
        assert HeaderFallback.FIRST_ROW == "first_row"
        assert HeaderFallback.NONE == "none"
        assert HeaderFallback.ERROR == "error"

    def test_data_type_equals_string(self) -> None:
        assert DataType.EMPTY == "empty"
        assert DataType.TEXT == "text"
        assert DataType.NUMBER == "number"
        assert DataType.BOOLEAN == "boolean"
        assert DataType.DATE == "date"
        assert DataType.ERROR == "error"
        assert DataType.FORMULA == "formula"

    def test_format_type_equals_string(self) -> None:
        assert FormatType.XLSX == "xlsx"
        assert FormatType.CSV == "csv"
        assert FormatType.UNKNOWN == "unknown"


class TestSheetConfigEnumCoercion:
    """Ensure SheetConfig accepts both raw strings and enum values."""

    def test_accepts_raw_string_merge_strategy(self) -> None:
        config = SheetConfig(merge_strategy="fill")
        assert config.merge_strategy == MergeStrategy.FILL
        assert config.merge_strategy == "fill"

    def test_accepts_enum_merge_strategy(self) -> None:
        config = SheetConfig(merge_strategy=MergeStrategy.SKIP)
        assert config.merge_strategy == MergeStrategy.SKIP

    def test_accepts_raw_string_header_mode(self) -> None:
        config = SheetConfig(header_detection_mode="auto")
        assert config.header_detection_mode == HeaderDetectionMode.AUTO

    def test_accepts_raw_string_header_fallback(self) -> None:
        config = SheetConfig(header_fallback="none")
        assert config.header_fallback == HeaderFallback.NONE

    def test_rejects_invalid_merge_strategy(self) -> None:
        with pytest.raises(ValueError, match="banana"):
            SheetConfig(merge_strategy="banana")

    def test_rejects_invalid_header_mode(self) -> None:
        with pytest.raises(ValueError, match="turbo"):
            SheetConfig(header_detection_mode="turbo")

    def test_rejects_invalid_header_fallback(self) -> None:
        with pytest.raises(ValueError, match="crash"):
            SheetConfig(header_fallback="crash")

    def test_default_values_are_enums(self) -> None:
        config = SheetConfig()
        assert isinstance(config.merge_strategy, MergeStrategy)
        assert isinstance(config.header_detection_mode, HeaderDetectionMode)
        assert isinstance(config.header_fallback, HeaderFallback)


class TestCellValueEnumCoercion:
    """Ensure CellValue accepts both raw strings and DataType enums."""

    def test_accepts_raw_string(self) -> None:
        cell = CellValue(value=42, data_type="number")
        assert cell.data_type == DataType.NUMBER
        assert cell.data_type == "number"

    def test_accepts_enum(self) -> None:
        cell = CellValue(value="hello", data_type=DataType.TEXT)
        assert cell.data_type == DataType.TEXT

    def test_default_is_empty(self) -> None:
        cell = CellValue(value=None)
        assert cell.data_type == DataType.EMPTY

    def test_rejects_invalid_type(self) -> None:
        with pytest.raises(ValueError):
            CellValue(value=None, data_type="nonexistent")


class TestFormatInfoEnumCoercion:
    """Ensure FormatInfo accepts both raw strings and FormatType enums."""

    def test_accepts_raw_string(self) -> None:
        info = FormatInfo(format_type="xlsx")
        assert info.format_type == FormatType.XLSX
        assert info.format_type == "xlsx"

    def test_accepts_enum(self) -> None:
        info = FormatInfo(format_type=FormatType.CSV)
        assert info.format_type == FormatType.CSV

    def test_rejects_invalid_format(self) -> None:
        with pytest.raises(ValueError):
            FormatInfo(format_type="pdf")


class TestEnumExports:
    """Ensure enums are properly exported from the package."""

    def test_enums_importable_from_package(self) -> None:
        from messy_xlsx import (
            DataType,
            FormatType,
            HeaderDetectionMode,
            HeaderFallback,
            MergeStrategy,
        )

        assert DataType.TEXT == "text"
        assert FormatType.XLSX == "xlsx"
        assert MergeStrategy.FILL == "fill"
        assert HeaderDetectionMode.SMART == "smart"
        assert HeaderFallback.FIRST_ROW == "first_row"
