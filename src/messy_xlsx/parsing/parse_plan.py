"""Pure compilation of sheet configuration into an immutable parse plan.

This module deliberately contains no source, workbook, worksheet, or DataFrame
access.  It resolves configuration and optional structure evidence into stable
projections consumed by the existing parsing and normalization layers.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, fields, is_dataclass
from enum import Enum
from inspect import isbuiltin, isfunction, ismethoddescriptor
from itertools import pairwise
from typing import Any, Final

from messy_xlsx.enums import (
    FormatType,
    HeaderDetectionMode,
    HeaderFallback,
    MergeStrategy,
)
from messy_xlsx.exceptions import StructureError
from messy_xlsx.models import SheetConfig, StructureInfo
from messy_xlsx.parsing.base_handler import ParseOptions
from messy_xlsx.parsing.contracts import OutputMode

_STRUCTURE_FORMATS: Final = frozenset({"xlsx", "xlsm", "xltx", "xltm"})
_TEXT_FORMATS: Final = frozenset({"csv", "tsv", "txt"})

_COMMA_DECIMAL_DOT_THOUSANDS: Final = frozenset(
    {
        "de",
        "nl",
        "it",
        "es",
        "pt",
        "el",
        "tr",
        "id",
    }
)
_COMMA_DECIMAL_SPACE_THOUSANDS: Final = frozenset(
    {
        "fr",
        "sv",
        "nb",
        "nn",
        "fi",
        "pl",
        "cs",
        "sk",
        "hu",
        "ro",
        "bg",
        "hr",
        "sl",
        "sr",
        "da",
        "ru",
        "uk",
    }
)


@dataclass(frozen=True, slots=True)
class FrozenDataclassValue:
    """Deterministic immutable projection of a supported dataclass value."""

    dataclass_type: type[Any]
    field_values: tuple[tuple[str, Any], ...]


@dataclass(frozen=True, slots=True)
class _FrozenMapping:
    items: tuple[tuple[Any, Any], ...]


@dataclass(frozen=True, slots=True)
class _FrozenList:
    items: tuple[Any, ...]


@dataclass(frozen=True, slots=True)
class _FrozenTuple:
    items: tuple[Any, ...]


@dataclass(frozen=True, slots=True)
class _FrozenSet:
    items: tuple[Any, ...]


@dataclass(frozen=True, slots=True)
class _FrozenFrozenSet:
    items: tuple[Any, ...]


@dataclass(frozen=True, slots=True)
class _FrozenBytearray:
    value: bytes


@dataclass(frozen=True, slots=True)
class _FrozenMemoryview:
    value: bytes


@dataclass(frozen=True, slots=True)
class _FrozenFloat:
    hexadecimal: str


@dataclass(frozen=True, slots=True)
class _FrozenComplex:
    real: _FrozenFloat
    imaginary: _FrozenFloat


@dataclass(frozen=True, slots=True)
class _FrozenObject:
    object_type: type[Any]
    attribute_values: tuple[tuple[str, Any], ...]


@dataclass(frozen=True, slots=True)
class ConfigSnapshot:
    """Deep immutable snapshot of collection-valued sheet configuration."""

    type_hints: tuple[tuple[Any, Any], ...]
    column_renames: tuple[tuple[Any, Any], ...]
    drop_conditions: tuple[tuple[Any, Any], ...]

    @classmethod
    def from_config(cls, config: SheetConfig) -> ConfigSnapshot:
        """Copy every consumed configuration collection recursively."""
        return cls(
            type_hints=_freeze_mapping(config.type_hints),
            column_renames=_freeze_mapping(config.column_renames),
            drop_conditions=tuple(
                (
                    _freeze(condition.get("column")),
                    _freeze(condition.get("value")),
                )
                for condition in config.drop_conditions
            ),
        )

    def thaw_type_hints(self) -> dict[Any, Any]:
        """Return a fresh deep mutable projection for legacy normalization."""
        return {_thaw(key): _thaw(value) for key, value in self.type_hints}

    def thaw_column_renames(self) -> dict[Any, Any]:
        """Return a fresh deep mutable projection for pandas rename."""
        return {_thaw(key): _thaw(value) for key, value in self.column_renames}

    def thaw_drop_conditions(self) -> list[tuple[Any, Any]]:
        """Return fresh comparison values for legacy row filtering."""
        return [(_thaw(column), _thaw(value)) for column, value in self.drop_conditions]


@dataclass(frozen=True, slots=True)
class ParsePlan:
    """Frozen internal decisions for one sheet parse.

    Configuration values are projected to immutable tagged snapshots. Legacy
    consumers receive fresh deep projections through the ``thaw_*`` methods.
    """

    # Projection for format handlers.
    skip_rows: int
    header_rows: int
    skip_footer: int
    merge_strategy: MergeStrategy | str
    ignore_hidden: bool
    cell_range: str | None
    data_only: bool
    auto_detect_header: bool

    # Projection for normalization.
    normalize: bool
    decimal_separator: str | None
    thousands_separator: str | None
    use_extended_missing_list: bool
    preserve_types: bool
    type_hints: tuple[tuple[Any, Any], ...]
    skip_normalization_steps: tuple[str, ...]

    # Projection for post-processing.
    sanitize_column_names: bool
    column_renames: tuple[tuple[Any, Any], ...]
    drop_regex: str | None
    drop_conditions: tuple[tuple[Any, Any], ...]

    # Projection for backend orchestration. Defaults preserve the established
    # three-argument compiler and direct ParsePlan construction behavior.
    output_mode: OutputMode = OutputMode.MATERIALIZED
    batch_size: int | None = None

    def to_parse_options(self) -> ParseOptions:
        """Return a fresh mutable projection for the existing handler API."""
        return ParseOptions(
            skip_rows=self.skip_rows,
            header_rows=self.header_rows,
            skip_footer=self.skip_footer,
            merge_strategy=self.merge_strategy,
            ignore_hidden=self.ignore_hidden,
            cell_range=self.cell_range,
            data_only=self.data_only,
            auto_detect_header=self.auto_detect_header,
        )

    def thaw_type_hints(self) -> dict[Any, Any]:
        """Return fresh type-hint values in their legacy container kinds."""
        return {_thaw(key): _thaw(value) for key, value in self.type_hints}

    def thaw_column_renames(self) -> dict[Any, Any]:
        """Return fresh column rename values for the legacy pandas path."""
        return {_thaw(key): _thaw(value) for key, value in self.column_renames}

    def thaw_drop_conditions(self) -> list[tuple[Any, Any]]:
        """Return fresh condition values for the legacy pandas path."""
        return [(_thaw(column), _thaw(value)) for column, value in self.drop_conditions]


def requires_structure_analysis(
    config: SheetConfig,
    format_type: FormatType | str,
) -> bool:
    """Return whether current behavior requires OOXML structure evidence."""
    return config.auto_detect and format_type in _STRUCTURE_FORMATS


def compile_parse_plan(
    config: SheetConfig,
    structure: StructureInfo | None,
    format_type: FormatType | str,
    output_mode: OutputMode = OutputMode.MATERIALIZED,
    batch_size: int | None = None,
) -> ParsePlan:
    """Compile configuration and optional structure evidence without mutation or I/O."""
    output_mode = OutputMode(output_mode)
    batch_size = _validated_batch_size(output_mode, batch_size)
    snapshot = ConfigSnapshot.from_config(config)

    use_structure = requires_structure_analysis(config, format_type)
    if use_structure and structure is None:
        raise ValueError("StructureInfo is required for auto-detected OOXML parsing")

    active_structure = structure if use_structure else None

    skip_rows = config.skip_rows
    header_rows = config.header_rows
    skip_footer = config.skip_footer
    merge_strategy = config.merge_strategy
    ignore_hidden = not config.include_hidden
    locale = config.locale

    if active_structure is not None:
        skip_rows, header_rows = _resolve_header_rows(config, active_structure)
        skip_footer = _resolve_skip_footer(config, active_structure)
        locale = config.locale or active_structure.detected_locale

        if not active_structure.merged_ranges:
            merge_strategy = "skip"
        if not active_structure.hidden_rows and not active_structure.hidden_columns:
            ignore_hidden = False

    decimal_separator, thousands_separator = _resolve_separators(config, locale)

    skipped_steps: list[str] = []
    if not config.normalize_whitespace:
        skipped_steps.append("whitespace")
    if not config.normalize_numbers:
        skipped_steps.append("numbers")
    if not config.normalize_dates:
        skipped_steps.append("dates")
    if not config.ensure_type_consistency:
        skipped_steps.append("type_coercion")

    return ParsePlan(
        skip_rows=skip_rows,
        header_rows=header_rows,
        skip_footer=skip_footer,
        merge_strategy=merge_strategy,
        ignore_hidden=ignore_hidden,
        cell_range=config.cell_range,
        data_only=config.evaluate_formulas,
        auto_detect_header=config.auto_detect and format_type in _TEXT_FORMATS,
        normalize=config.normalize,
        decimal_separator=decimal_separator,
        thousands_separator=thousands_separator,
        use_extended_missing_list=config.use_extended_missing_list,
        preserve_types=config.preserve_types,
        type_hints=snapshot.type_hints,
        skip_normalization_steps=tuple(skipped_steps),
        sanitize_column_names=config.sanitize_column_names,
        column_renames=snapshot.column_renames,
        drop_regex=config.drop_regex,
        drop_conditions=snapshot.drop_conditions,
        output_mode=output_mode,
        batch_size=batch_size,
    )


def _validated_batch_size(
    mode: OutputMode,
    batch_size: int | None,
) -> int | None:
    """Reject an unusable streaming window before any backend work can start."""
    if mode is OutputMode.STREAMING and (type(batch_size) is not int or batch_size < 1):
        raise ValueError("batch_size must be >= 1 for streaming output")
    return batch_size


def _freeze_mapping(value: Mapping[Any, Any]) -> tuple[tuple[Any, Any], ...]:
    """Return a deterministic immutable mapping projection."""
    frozen_items = [(_freeze(key), _freeze(item)) for key, item in value.items()]
    return _order_mapping_items(frozen_items)


def _freeze(value: Any, active: set[int] | None = None) -> Any:  # noqa: C901
    """Recursively snapshot supported mutable values without retaining aliases."""
    if active is None:
        active = set()
    if isinstance(value, Enum):
        _validate_enum(value)
        return value
    if value is None or isinstance(value, (bool, int, str, bytes)):
        return value
    if isinstance(value, float):
        return _FrozenFloat(value.hex())
    if isinstance(value, complex):
        return _FrozenComplex(
            _FrozenFloat(value.real.hex()),
            _FrozenFloat(value.imag.hex()),
        )
    if (
        isinstance(value, type)
        or isfunction(value)
        or isbuiltin(value)
        or ismethoddescriptor(value)
    ):
        return value

    identity = id(value)
    if identity in active:
        raise TypeError("cyclic configuration values are not supported")
    active.add(identity)
    try:
        if type(value) is dict:
            items = [(_freeze(key, active), _freeze(item, active)) for key, item in value.items()]
            return _FrozenMapping(_order_mapping_items(items))
        if type(value) is list:
            return _FrozenList(tuple(_freeze(item, active) for item in value))
        if type(value) is tuple:
            return _FrozenTuple(tuple(_freeze(item, active) for item in value))
        if type(value) is set:
            set_items = tuple(_freeze(item, active) for item in value)
            return _FrozenSet(_order_set_items(set_items))
        if type(value) is frozenset:
            frozenset_items = tuple(_freeze(item, active) for item in value)
            return _FrozenFrozenSet(_order_set_items(frozenset_items))
        if type(value) is bytearray:
            return _FrozenBytearray(bytes(value))
        if type(value) is memoryview:
            return _FrozenMemoryview(value.tobytes())
        if is_dataclass(value) and not isinstance(value, type):
            return FrozenDataclassValue(
                dataclass_type=type(value),
                field_values=tuple(
                    (field.name, _freeze(getattr(value, field.name), active))
                    for field in fields(value)
                ),
            )

        attributes = _object_attributes(value)
        if attributes:
            return _FrozenObject(
                object_type=type(value),
                attribute_values=tuple(
                    (name, _freeze(item, active)) for name, item in sorted(attributes.items())
                ),
            )
        try:
            hash(value)
        except TypeError as error:
            raise TypeError(
                f"unsupported mutable configuration value: {type(value).__name__}"
            ) from error
        return value
    finally:
        active.remove(identity)


def _thaw(value: Any) -> Any:  # noqa: C901
    """Return a fresh legacy value from a tagged immutable snapshot."""
    if isinstance(value, _FrozenMapping):
        return {_thaw(key): _thaw(item) for key, item in value.items}
    if isinstance(value, _FrozenList):
        return [_thaw(item) for item in value.items]
    if isinstance(value, _FrozenTuple):
        return tuple(_thaw(item) for item in value.items)
    if isinstance(value, _FrozenSet):
        return {_thaw(item) for item in value.items}
    if isinstance(value, _FrozenFrozenSet):
        return frozenset(_thaw(item) for item in value.items)
    if isinstance(value, _FrozenBytearray):
        return bytearray(value.value)
    if isinstance(value, _FrozenMemoryview):
        return memoryview(bytearray(value.value))
    if isinstance(value, _FrozenFloat):
        return float.fromhex(value.hexadecimal)
    if isinstance(value, _FrozenComplex):
        return complex(_thaw(value.real), _thaw(value.imaginary))
    if isinstance(value, FrozenDataclassValue):
        instance = object.__new__(value.dataclass_type)
        for name, item in value.field_values:
            object.__setattr__(instance, name, _thaw(item))
        return instance
    if isinstance(value, _FrozenObject):
        instance = object.__new__(value.object_type)
        for name, item in value.attribute_values:
            object.__setattr__(instance, name, _thaw(item))
        return instance
    return value


def _order_mapping_items(items: list[tuple[Any, Any]]) -> tuple[tuple[Any, Any], ...]:
    """Order frozen mapping entries and reject ambiguous structural keys."""
    keyed_items = [(_stable_token(key), key, item) for key, item in items]
    keyed_items.sort(key=lambda item: item[0])
    for previous, current in pairwise(keyed_items):
        if previous[0] == current[0] and previous[1] != current[1]:
            raise TypeError("configuration mapping keys cannot be ordered deterministically")
    return tuple((key, item) for _, key, item in keyed_items)


def _order_set_items(items: tuple[Any, ...]) -> tuple[Any, ...]:
    """Order frozen set entries and reject ambiguous structural values."""
    keyed_items = sorted((_stable_token(item), item) for item in items)
    for previous, current in pairwise(keyed_items):
        if previous[0] == current[0] and previous[1] != current[1]:
            raise TypeError("configuration set values cannot be ordered deterministically")
    return tuple(item for _, item in keyed_items)


def _stable_token(value: Any) -> tuple[Any, ...]:  # noqa: C901
    """Build a deterministic structural ordering token without repr or ids."""
    if value is None:
        return ("none",)
    if isinstance(value, Enum):
        value_type = type(value)
        return ("enum", value_type.__module__, value_type.__qualname__, value.name)
    if isinstance(value, bool):
        return ("bool", int(value))
    if isinstance(value, int):
        return ("int", str(value))
    if isinstance(value, str):
        return ("str", value)
    if isinstance(value, bytes):
        return ("bytes", value.hex())
    if isinstance(value, _FrozenFloat):
        return ("float", value.hexadecimal)
    if isinstance(value, _FrozenComplex):
        return ("complex", _stable_token(value.real), _stable_token(value.imaginary))
    if isinstance(value, _FrozenMapping):
        return (
            "mapping",
            tuple((_stable_token(key), _stable_token(item)) for key, item in value.items),
        )
    if isinstance(value, _FrozenList):
        return ("list", tuple(_stable_token(item) for item in value.items))
    if isinstance(value, _FrozenTuple):
        return ("tuple", tuple(_stable_token(item) for item in value.items))
    if isinstance(value, _FrozenSet):
        return ("set", tuple(_stable_token(item) for item in value.items))
    if isinstance(value, _FrozenFrozenSet):
        return ("frozenset", tuple(_stable_token(item) for item in value.items))
    if isinstance(value, _FrozenBytearray):
        return ("bytearray", value.value.hex())
    if isinstance(value, _FrozenMemoryview):
        return ("memoryview", value.value.hex())
    if isinstance(value, _FrozenObject):
        return (
            "object",
            value.object_type.__module__,
            value.object_type.__qualname__,
            tuple((name, _stable_token(item)) for name, item in value.attribute_values),
        )
    if isinstance(value, FrozenDataclassValue):
        return (
            "dataclass",
            value.dataclass_type.__module__,
            value.dataclass_type.__qualname__,
            tuple((name, _stable_token(item)) for name, item in value.field_values),
        )
    if (
        isinstance(value, type)
        or isfunction(value)
        or isbuiltin(value)
        or ismethoddescriptor(value)
    ):
        return ("symbol", value.__module__, value.__qualname__)
    raise TypeError(
        f"configuration value cannot be ordered deterministically: {type(value).__name__}"
    )


def _validate_enum(value: Enum) -> None:
    """Allow only state-free enum members whose payload is deeply immutable."""
    state_names = set(vars(value)) - {"_name_", "_value_", "__objclass__", "_sort_order_"}
    if state_names or not _is_deeply_immutable(value.value):
        raise TypeError(f"mutable Enum configuration value: {type(value).__name__}.{value.name}")


def _is_deeply_immutable(value: Any) -> bool:
    if value is None or isinstance(value, (bool, int, float, complex, str, bytes, type)):
        return True
    if isinstance(value, tuple | frozenset):
        return all(_is_deeply_immutable(item) for item in value)
    return isfunction(value) or isbuiltin(value) or ismethoddescriptor(value)


def _object_attributes(value: Any) -> dict[str, Any]:
    """Collect mutable instance state without consulting value representations."""
    attributes = dict(vars(value)) if hasattr(value, "__dict__") else {}
    for value_type in type(value).__mro__:
        slots = value_type.__dict__.get("__slots__", ())
        if isinstance(slots, str):
            slots = (slots,)
        for name in slots:
            if name not in {"__dict__", "__weakref__"} and hasattr(value, name):
                attributes.setdefault(name, getattr(value, name))
    return attributes


def _resolve_header_rows(
    config: SheetConfig,
    structure: StructureInfo,
) -> tuple[int, int]:
    """Resolve row offsets while preserving existing header-mode semantics."""
    skip_rows = config.skip_rows
    header_rows = config.header_rows
    header_is_usable = (
        structure.header_row is not None
        and structure.header_confidence >= config.header_confidence_threshold
    )

    if config.header_detection_mode == HeaderDetectionMode.AUTO:
        if header_is_usable:
            skip_rows, header_rows = _detected_header_rows(config, structure)
        elif config.header_fallback == HeaderFallback.FIRST_ROW:
            skip_rows = 0
            header_rows = 1
        elif config.header_fallback == HeaderFallback.NONE:
            header_rows = 0
        elif config.header_fallback == HeaderFallback.ERROR:
            raise StructureError(
                f"No header detected with sufficient confidence "
                f"(found: {structure.header_confidence:.2f}, "
                f"required: {config.header_confidence_threshold:.2f})"
            )
    elif (
        config.header_detection_mode == HeaderDetectionMode.SMART
        and config.skip_rows == 0
        and header_is_usable
    ):
        skip_rows, header_rows = _detected_header_rows(config, structure)

    return skip_rows, header_rows


def _detected_header_rows(
    config: SheetConfig,
    structure: StructureInfo,
) -> tuple[int, int]:
    """Convert a detected one-based header into handler row offsets."""
    if structure.header_row is None:
        raise ValueError("A detected header row is required")

    skip_rows = max(0, structure.header_row - 1)
    if structure.hidden_rows and not config.include_hidden:
        hidden_before_header = sum(row < structure.header_row for row in structure.hidden_rows)
        skip_rows = max(0, skip_rows - hidden_before_header)

    return skip_rows, structure.header_rows_count


def _resolve_skip_footer(config: SheetConfig, structure: StructureInfo) -> int:
    """Resolve footer trimming with caller and detected-table precedence."""
    skip_footer = structure.suggested_skip_footer
    if config.skip_footer > 0:
        return config.skip_footer

    # A caller-selected range already constrains the parse. Sheet-global
    # detected footer evidence cannot be mapped safely into that local range.
    if config.cell_range:
        return 0

    if structure.num_tables > 1 and structure.table_ranges:
        first_table_end = structure.table_ranges[0]["end_row"]
        rows_after_first_table = max(0, structure.data_end_row - first_table_end)

        if structure.hidden_rows and not config.include_hidden:
            hidden_after = sum(
                first_table_end < row <= structure.data_end_row for row in structure.hidden_rows
            )
            rows_after_first_table = max(0, rows_after_first_table - hidden_after)

        if rows_after_first_table > 0:
            skip_footer = max(skip_footer, rows_after_first_table)

    return skip_footer


def _resolve_separators(
    config: SheetConfig,
    locale: str | None,
) -> tuple[str | None, str | None]:
    """Resolve explicit or locale-derived number separators."""
    decimal_separator = config.decimal_separator
    thousands_separator = config.thousands_separator
    if decimal_separator is None and thousands_separator is None and locale and locale != "auto":
        return _locale_to_separators(locale)
    return decimal_separator, thousands_separator


def _locale_to_separators(locale: str) -> tuple[str, str]:
    """Convert an existing locale convention to number separators."""
    lang = locale.split("_")[0].lower()
    country = locale.split("_")[1].upper() if "_" in locale else ""

    if lang == "de" and country == "CH":
        return ".", "'"
    if lang in _COMMA_DECIMAL_DOT_THOUSANDS:
        return ",", "."
    if lang in _COMMA_DECIMAL_SPACE_THOUSANDS:
        return ",", " "
    return ".", ","
