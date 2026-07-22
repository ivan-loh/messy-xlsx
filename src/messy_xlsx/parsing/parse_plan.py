"""Pure compilation of sheet configuration into an immutable parse plan.

This module deliberately contains no source, workbook, worksheet, or DataFrame
access.  It resolves configuration and optional structure evidence into stable
projections consumed by the existing parsing and normalization layers.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field, is_dataclass
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal
from enum import Enum
from inspect import isbuiltin, isfunction
from itertools import pairwise
from types import MethodDescriptorType, WrapperDescriptorType
from typing import Any, Final, TypeVar, cast
from uuid import UUID

from messy_xlsx._fallback_signals import (
    _FallbackBlockReason,
    _mark_fallback_blocked,
)
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
_PY_TPFLAGS_IMMUTABLETYPE: Final = 1 << 8
_ENUM_SNAPSHOT_ERROR: Final = "Enum values require drop-condition identity semantics"
_ChoiceEnum = TypeVar("_ChoiceEnum", bound=Enum)

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

    type_identity: int
    type_module: str
    type_qualname: str
    dataclass_type: type[Any] = field(compare=False, hash=False, repr=False)
    attribute_values: tuple[tuple[str, Any], ...]


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
    format: str
    shape: tuple[int, ...]
    strides: tuple[int, ...]
    suboffsets: tuple[int, ...]
    itemsize: int
    ndim: int
    readonly: bool


@dataclass(frozen=True, slots=True)
class _FrozenFloat:
    hexadecimal: str


@dataclass(frozen=True, slots=True)
class _FrozenComplex:
    real: _FrozenFloat
    imaginary: _FrozenFloat


@dataclass(frozen=True, slots=True)
class _FrozenDecimal:
    text: str


@dataclass(frozen=True, slots=True)
class _FrozenDate:
    year: int
    month: int
    day: int


@dataclass(frozen=True, slots=True)
class _FrozenTimedelta:
    days: int
    seconds: int
    microseconds: int


@dataclass(frozen=True, slots=True)
class _FrozenTimezone:
    offset: _FrozenTimedelta
    name: str


@dataclass(frozen=True, slots=True)
class _FrozenDatetime:
    year: int
    month: int
    day: int
    hour: int
    minute: int
    second: int
    microsecond: int
    timezone: _FrozenTimezone | None
    fold: int


@dataclass(frozen=True, slots=True)
class _FrozenTime:
    hour: int
    minute: int
    second: int
    microsecond: int
    timezone: _FrozenTimezone | None
    fold: int


@dataclass(frozen=True, slots=True)
class _FrozenUUID:
    integer: int


@dataclass(frozen=True, slots=True)
class _FrozenTypeReference:
    type_identity: int
    type_module: str
    type_qualname: str
    reference: type[Any] = field(compare=False, hash=False, repr=False)


@dataclass(frozen=True, slots=True)
class _FrozenIdentityReference:
    type_identity: int
    type_module: str
    type_qualname: str
    identity: int
    reference: Any = field(compare=False, hash=False, repr=False)


@dataclass(frozen=True, slots=True)
class _FrozenObject:
    type_identity: int
    type_module: str
    type_qualname: str
    object_type: type[Any] = field(compare=False, hash=False, repr=False)
    attribute_values: tuple[tuple[str, Any], ...]


@dataclass(frozen=True, slots=True)
class ConfigSnapshot:
    """Deep immutable snapshot of collection-valued sheet configuration."""

    type_hints: tuple[tuple[Any, Any], ...]
    column_renames: tuple[tuple[Any, Any], ...]
    drop_conditions: tuple[tuple[Any, Any], ...]

    @classmethod
    def from_config(cls, config: SheetConfig) -> ConfigSnapshot:
        """Copy typed built-in configuration collections recursively.

        ``SheetConfig`` exposes concrete ``dict``/``list`` fields. Requiring
        those exact container types keeps snapshotting independent of
        caller-controlled collection hooks.
        """
        if type(config.type_hints) is not dict:
            raise TypeError("type_hints must be an exact dict")
        if type(config.column_renames) is not dict:
            raise TypeError("column_renames must be an exact dict")
        if type(config.drop_conditions) is not list:
            raise TypeError("drop_conditions must be an exact list")
        for index, condition in enumerate(config.drop_conditions):
            if type(condition) is not dict:
                raise TypeError(f"drop_conditions[{index}] must be an exact dict")
        return cls(
            type_hints=_freeze_mapping(config.type_hints),
            column_renames=_freeze_mapping(config.column_renames),
            drop_conditions=tuple(
                (
                    _freeze(condition.get("column")),
                    _freeze(
                        condition.get("value"),
                        preserve_identity_reference=True,
                    ),
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
class _ValidatedConfigProjection:
    """Exact immutable scalar projection used while compiling a plan."""

    skip_rows: int
    header_rows: int
    skip_footer: int
    cell_range: str | None
    auto_detect: bool
    include_hidden: bool
    merge_strategy: MergeStrategy
    locale: str | None
    evaluate_formulas: bool
    drop_regex: str | None
    header_detection_mode: HeaderDetectionMode
    header_confidence_threshold: float
    header_fallback: HeaderFallback
    normalize: bool
    normalize_dates: bool
    normalize_numbers: bool
    normalize_whitespace: bool
    use_extended_missing_list: bool
    preserve_types: bool
    ensure_type_consistency: bool
    decimal_separator: str | None
    thousands_separator: str | None
    sanitize_column_names: bool
    snapshot: ConfigSnapshot


@dataclass(frozen=True, slots=True)
class ParsePlan:
    """Frozen internal decisions for one sheet parse.

    Configuration values are projected to immutable tagged snapshots. Legacy
    consumers receive fresh deep projections through the ``thaw_*`` methods.
    Drop operands with exact object identity equality/hash use immutable
    identity tokens that thaw to the exact semantic reference.
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
    output_mode, batch_size, projected = _validated_plan_inputs(
        config,
        output_mode,
        batch_size,
    )

    use_structure = projected.auto_detect and format_type in _STRUCTURE_FORMATS
    if use_structure and structure is None:
        error = ValueError("StructureInfo is required for auto-detected OOXML parsing")
        raise _mark_fallback_blocked(
            error,
            _FallbackBlockReason.CONFIGURATION,
        )

    active_structure = structure if use_structure else None

    skip_rows = projected.skip_rows
    header_rows = projected.header_rows
    skip_footer = projected.skip_footer
    merge_strategy = projected.merge_strategy
    ignore_hidden = not projected.include_hidden
    locale = projected.locale

    if active_structure is not None:
        skip_rows, header_rows = _configured_header_rows(projected, active_structure)
        skip_footer = _resolve_skip_footer(projected, active_structure)
        locale = projected.locale or active_structure.detected_locale

        if not active_structure.merged_ranges:
            merge_strategy = MergeStrategy.SKIP
        if not active_structure.hidden_rows and not active_structure.hidden_columns:
            ignore_hidden = False

    decimal_separator, thousands_separator = _resolve_separators(projected, locale)

    skipped_steps: list[str] = []
    if not projected.normalize_whitespace:
        skipped_steps.append("whitespace")
    if not projected.normalize_numbers:
        skipped_steps.append("numbers")
    if not projected.normalize_dates:
        skipped_steps.append("dates")
    if not projected.ensure_type_consistency:
        skipped_steps.append("type_coercion")

    try:
        return ParsePlan(
            skip_rows=_project_nonnegative_int(skip_rows, "skip_rows"),
            header_rows=_project_nonnegative_int(header_rows, "header_rows"),
            skip_footer=_project_nonnegative_int(skip_footer, "skip_footer"),
            merge_strategy=_project_choice(
                merge_strategy,
                MergeStrategy,
                "merge_strategy",
            ),
            ignore_hidden=_project_bool(ignore_hidden, "ignore_hidden"),
            cell_range=projected.cell_range,
            data_only=projected.evaluate_formulas,
            auto_detect_header=projected.auto_detect and format_type in _TEXT_FORMATS,
            normalize=projected.normalize,
            decimal_separator=_project_optional_str(
                decimal_separator,
                "decimal_separator",
            ),
            thousands_separator=_project_optional_str(
                thousands_separator,
                "thousands_separator",
            ),
            use_extended_missing_list=projected.use_extended_missing_list,
            preserve_types=projected.preserve_types,
            type_hints=projected.snapshot.type_hints,
            skip_normalization_steps=tuple(skipped_steps),
            sanitize_column_names=projected.sanitize_column_names,
            column_renames=projected.snapshot.column_renames,
            drop_regex=projected.drop_regex,
            drop_conditions=projected.snapshot.drop_conditions,
            output_mode=output_mode,
            batch_size=batch_size,
        )
    except (TypeError, ValueError) as error:
        _mark_fallback_blocked(error, _FallbackBlockReason.CONFIGURATION)
        raise


def _validated_plan_inputs(
    config: SheetConfig,
    output_mode: OutputMode,
    batch_size: int | None,
) -> tuple[OutputMode, int | None, _ValidatedConfigProjection]:
    """Validate and snapshot caller configuration with a typed failure signal."""
    try:
        mode = OutputMode(output_mode)
        size = _validated_batch_size(mode, batch_size)
        projected = _ValidatedConfigProjection(
            skip_rows=_project_nonnegative_int(config.skip_rows, "skip_rows"),
            header_rows=_project_nonnegative_int(config.header_rows, "header_rows"),
            skip_footer=_project_nonnegative_int(config.skip_footer, "skip_footer"),
            cell_range=_project_optional_str(config.cell_range, "cell_range"),
            auto_detect=_project_bool(config.auto_detect, "auto_detect"),
            include_hidden=_project_bool(config.include_hidden, "include_hidden"),
            merge_strategy=_project_choice(
                config.merge_strategy,
                MergeStrategy,
                "merge_strategy",
            ),
            locale=_project_optional_str(config.locale, "locale"),
            evaluate_formulas=_project_bool(
                config.evaluate_formulas,
                "evaluate_formulas",
            ),
            drop_regex=_project_optional_str(config.drop_regex, "drop_regex"),
            header_detection_mode=_project_choice(
                config.header_detection_mode,
                HeaderDetectionMode,
                "header_detection_mode",
            ),
            header_confidence_threshold=_project_confidence_threshold(
                config.header_confidence_threshold
            ),
            header_fallback=_project_choice(
                config.header_fallback,
                HeaderFallback,
                "header_fallback",
            ),
            normalize=_project_bool(config.normalize, "normalize"),
            normalize_dates=_project_bool(config.normalize_dates, "normalize_dates"),
            normalize_numbers=_project_bool(
                config.normalize_numbers,
                "normalize_numbers",
            ),
            normalize_whitespace=_project_bool(
                config.normalize_whitespace,
                "normalize_whitespace",
            ),
            use_extended_missing_list=_project_bool(
                config.use_extended_missing_list,
                "use_extended_missing_list",
            ),
            preserve_types=_project_bool(config.preserve_types, "preserve_types"),
            ensure_type_consistency=_project_bool(
                config.ensure_type_consistency,
                "ensure_type_consistency",
            ),
            decimal_separator=_project_optional_str(
                config.decimal_separator,
                "decimal_separator",
            ),
            thousands_separator=_project_optional_str(
                config.thousands_separator,
                "thousands_separator",
            ),
            sanitize_column_names=_project_bool(
                config.sanitize_column_names,
                "sanitize_column_names",
            ),
            snapshot=ConfigSnapshot.from_config(config),
        )
    except (TypeError, ValueError) as error:
        _mark_fallback_blocked(error, _FallbackBlockReason.CONFIGURATION)
        raise
    return mode, size, projected


def _project_nonnegative_int(value: object, field_name: str) -> int:
    """Copy an integer-like scalar to an owned exact ``int``."""
    if not isinstance(value, int) or isinstance(value, bool):
        raise TypeError(f"{field_name} must be an int")
    projected = value if type(value) is int else int.__int__(value)
    if projected < 0:
        raise ValueError(f"{field_name} must be >= 0, got {projected}")
    return projected


def _project_bool(value: object, field_name: str) -> bool:
    """Require the exact immutable boolean type without truthiness coercion."""
    if type(value) is not bool:
        raise TypeError(f"{field_name} must be a bool")
    return value


def _project_str(value: object, field_name: str) -> str:
    """Copy a string subclass to an owned exact ``str`` value."""
    if not isinstance(value, str):
        raise TypeError(f"{field_name} must be a str")
    return value if type(value) is str else str.__str__(value)


def _project_optional_str(value: object, field_name: str) -> str | None:
    if value is None:
        return None
    return _project_str(value, field_name)


def _project_choice(
    value: object,
    choice_type: type[_ChoiceEnum],
    field_name: str,
) -> _ChoiceEnum:
    """Canonicalize a string choice to the exact project Enum member."""
    text = _project_str(value, field_name)
    return choice_type(text)


def _project_confidence_threshold(value: object) -> float:
    """Copy the header threshold to an exact bounded float."""
    if isinstance(value, bool):
        raise TypeError("header_confidence_threshold must be a float")
    if type(value) is float:
        projected = value
    elif isinstance(value, float):
        projected = float.__float__(value)
    elif isinstance(value, int):
        integer = value if type(value) is int else int.__int__(value)
        projected = float(integer)
    else:
        raise TypeError("header_confidence_threshold must be a float")
    if not 0.0 <= projected <= 1.0:
        raise ValueError(
            f"header_confidence_threshold must be between 0.0 and 1.0, got {projected}"
        )
    return projected


def _configured_header_rows(
    config: _ValidatedConfigProjection,
    structure: StructureInfo,
) -> tuple[int, int]:
    """Resolve header configuration with a typed failure signal."""
    try:
        return _resolve_header_rows(config, structure)
    except (StructureError, ValueError) as error:
        _mark_fallback_blocked(error, _FallbackBlockReason.CONFIGURATION)
        raise


def _validated_batch_size(
    mode: OutputMode,
    batch_size: int | None,
) -> int | None:
    """Reject any invalid retained window before backend work can start."""
    if batch_size is None:
        if mode is OutputMode.STREAMING:
            raise ValueError("batch_size must be >= 1 for streaming output")
        return None
    if not isinstance(batch_size, int) or isinstance(batch_size, bool):
        raise ValueError("batch_size must be >= 1")
    size = int.__int__(batch_size)
    if size < 1:
        raise ValueError("batch_size must be >= 1")
    return size


def _freeze_mapping(value: Mapping[Any, Any]) -> tuple[tuple[Any, Any], ...]:
    """Return a deterministic immutable mapping projection."""
    frozen_items = [(_freeze(key), _freeze(item)) for key, item in value.items()]
    return _order_mapping_items(frozen_items)


def _freeze(  # noqa: C901
    value: Any,
    active: set[int] | None = None,
    *,
    preserve_identity_reference: bool = False,
) -> Any:
    """Recursively snapshot supported mutable values without retaining aliases."""
    if active is None:
        active = set()
    if isinstance(value, Enum):
        if preserve_identity_reference:
            return _freeze_identity_reference(value)
        raise TypeError(_ENUM_SNAPSHOT_ERROR)
    if value is None or type(value) in {bool, int, str, bytes}:
        return value
    if type(value) is float:
        return _FrozenFloat(value.hex())
    if type(value) is complex:
        return _FrozenComplex(
            _FrozenFloat(value.real.hex()),
            _FrozenFloat(value.imag.hex()),
        )
    if type(value) is Decimal:
        return _FrozenDecimal(str(value))
    if type(value) is datetime:
        frozen_timezone = _freeze_stdlib_timezone(value.tzinfo, "datetime")
        return _FrozenDatetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            frozen_timezone,
            value.fold,
        )
    if type(value) is date:
        return _FrozenDate(value.year, value.month, value.day)
    if type(value) is time:
        return _FrozenTime(
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            _freeze_stdlib_timezone(value.tzinfo, "time"),
            value.fold,
        )
    if type(value) is timedelta:
        return _FrozenTimedelta(value.days, value.seconds, value.microseconds)
    if type(value) is UUID:
        return _FrozenUUID(value.int)
    if isinstance(value, type):
        type_module, type_qualname = _safe_type_description(value)
        return _FrozenTypeReference(
            type_identity=id(value),
            type_module=type_module,
            type_qualname=type_qualname,
            reference=value,
        )
    if isfunction(value) or isbuiltin(value) or _is_method_descriptor(value):
        return value
    is_dataclass_instance = is_dataclass(value) and not isinstance(value, type)
    uses_identity_semantics = _uses_legacy_identity_semantics(value)
    if preserve_identity_reference and uses_identity_semantics:
        return _freeze_identity_reference(value)

    identity = id(value)
    if identity in active:
        raise TypeError("cyclic configuration values are not supported")
    active.add(identity)
    try:
        if type(value) is dict:
            items = [
                (
                    _freeze(
                        key,
                        active,
                        preserve_identity_reference=preserve_identity_reference,
                    ),
                    _freeze(
                        item,
                        active,
                        preserve_identity_reference=preserve_identity_reference,
                    ),
                )
                for key, item in value.items()
            ]
            return _FrozenMapping(_order_mapping_items(items))
        if type(value) is list:
            return _FrozenList(
                tuple(
                    _freeze(
                        item,
                        active,
                        preserve_identity_reference=preserve_identity_reference,
                    )
                    for item in value
                )
            )
        if type(value) is tuple:
            return _FrozenTuple(
                tuple(
                    _freeze(
                        item,
                        active,
                        preserve_identity_reference=preserve_identity_reference,
                    )
                    for item in value
                )
            )
        if type(value) is set:
            set_items = tuple(
                _freeze(
                    item,
                    active,
                    preserve_identity_reference=preserve_identity_reference,
                )
                for item in value
            )
            return _FrozenSet(_order_set_items(set_items))
        if type(value) is frozenset:
            frozenset_items = tuple(
                _freeze(
                    item,
                    active,
                    preserve_identity_reference=preserve_identity_reference,
                )
                for item in value
            )
            return _FrozenFrozenSet(_order_set_items(frozenset_items))
        if type(value) is bytearray:
            return _FrozenBytearray(bytes(value))
        if type(value) is memoryview:
            return _freeze_memoryview(value)
        if is_dataclass_instance:
            if not _is_safely_reconstructable_python_object(value):
                raise TypeError(f"opaque mutable configuration value: {type(value).__name__}")
            attributes = _object_attributes(value)
            value_type = type(value)
            type_module, type_qualname = _safe_type_description(value_type)
            return FrozenDataclassValue(
                type_identity=id(value_type),
                type_module=type_module,
                type_qualname=type_qualname,
                dataclass_type=value_type,
                attribute_values=tuple(
                    (
                        name,
                        _freeze(
                            item,
                            active,
                            preserve_identity_reference=preserve_identity_reference,
                        ),
                    )
                    for name, item in sorted(attributes.items())
                ),
            )
        if uses_identity_semantics:
            raise TypeError(f"opaque mutable configuration value: {type(value).__name__}")

        attributes = _object_attributes(value)
        if attributes:
            if not _is_safely_reconstructable_python_object(value):
                raise TypeError(f"opaque mutable configuration value: {type(value).__name__}")
            value_type = type(value)
            type_module, type_qualname = _safe_type_description(value_type)
            return _FrozenObject(
                type_identity=id(value_type),
                type_module=type_module,
                type_qualname=type_qualname,
                object_type=value_type,
                attribute_values=tuple(
                    (
                        name,
                        _freeze(
                            item,
                            active,
                            preserve_identity_reference=preserve_identity_reference,
                        ),
                    )
                    for name, item in sorted(attributes.items())
                ),
            )
        raise TypeError(f"unsupported mutable configuration value: {type(value).__name__}")
    finally:
        active.remove(identity)


def _freeze_identity_reference(value: Any) -> _FrozenIdentityReference:
    """Capture stable identity metadata without virtual class/member access."""
    value_type = type(value)
    type_module, type_qualname = _safe_type_description(value_type)
    return _FrozenIdentityReference(
        type_identity=id(value_type),
        type_module=type_module,
        type_qualname=type_qualname,
        identity=id(value),
        reference=value,
    )


def _freeze_memoryview(value: memoryview) -> _FrozenMemoryview:
    """Snapshot a reconstructable C-contiguous view and its equality metadata."""
    if not value.c_contiguous:
        raise TypeError("unsupported non-contiguous memoryview configuration value")
    if value.shape is None or value.strides is None or value.suboffsets is None:
        raise TypeError("unsupported memoryview configuration value")
    frozen = _FrozenMemoryview(
        value=value.tobytes(),
        format=value.format,
        shape=tuple(value.shape),
        strides=tuple(value.strides),
        suboffsets=tuple(value.suboffsets),
        itemsize=value.itemsize,
        ndim=value.ndim,
        readonly=value.readonly,
    )
    try:
        reconstructed = _thaw_memoryview(frozen)
    except (TypeError, ValueError) as error:
        raise TypeError("unsupported memoryview configuration value") from error
    try:
        metadata = (
            reconstructed.format,
            reconstructed.shape,
            reconstructed.strides,
            reconstructed.suboffsets,
            reconstructed.itemsize,
            reconstructed.ndim,
            reconstructed.readonly,
        )
        expected = (
            frozen.format,
            frozen.shape,
            frozen.strides,
            frozen.suboffsets,
            frozen.itemsize,
            frozen.ndim,
            frozen.readonly,
        )
        if metadata != expected:
            raise TypeError("unsupported memoryview configuration value")
    finally:
        reconstructed.release()
    return frozen


def _freeze_stdlib_timezone(
    value: object,
    scalar_name: str,
) -> _FrozenTimezone | None:
    """Snapshot only the immutable stdlib fixed-offset timezone type."""
    if value is None:
        return None
    if type(value) is not timezone:
        raise TypeError(f"unsupported mutable configuration value: {scalar_name}")
    offset = value.utcoffset(None)
    name = value.tzname(None)
    if offset is None or name is None:
        raise TypeError(f"unsupported mutable configuration value: {scalar_name}")
    return _FrozenTimezone(
        _FrozenTimedelta(offset.days, offset.seconds, offset.microseconds),
        name,
    )


def _thaw_stdlib_timezone(value: _FrozenTimezone | None) -> timezone | None:
    if value is None:
        return None
    offset = value.offset
    return timezone(
        timedelta(offset.days, offset.seconds, offset.microseconds),
        value.name,
    )


def _thaw_memoryview(value: _FrozenMemoryview) -> memoryview:
    backing: bytes | bytearray
    backing = value.value if value.readonly else bytearray(value.value)
    shape = None if not value.value and value.ndim == 1 and value.shape == (0,) else value.shape
    if shape is None:
        return cast(
            memoryview,
            memoryview(backing).cast(value.format),  # type: ignore[call-overload]
        )
    return cast(
        memoryview,
        memoryview(backing).cast(  # type: ignore[call-overload]
            value.format,
            shape=value.shape,
        ),
    )


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
        return _thaw_memoryview(value)
    if isinstance(value, _FrozenFloat):
        return float.fromhex(value.hexadecimal)
    if isinstance(value, _FrozenComplex):
        return complex(_thaw(value.real), _thaw(value.imaginary))
    if isinstance(value, _FrozenDecimal):
        return Decimal(value.text)
    if isinstance(value, _FrozenDate):
        return date(value.year, value.month, value.day)
    if isinstance(value, _FrozenTimedelta):
        return timedelta(value.days, value.seconds, value.microseconds)
    if isinstance(value, _FrozenDatetime):
        return datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            tzinfo=_thaw_stdlib_timezone(value.timezone),
            fold=value.fold,
        )
    if isinstance(value, _FrozenTime):
        return time(
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            tzinfo=_thaw_stdlib_timezone(value.timezone),
            fold=value.fold,
        )
    if isinstance(value, _FrozenUUID):
        return UUID(int=value.integer)
    if isinstance(value, _FrozenTypeReference):
        return value.reference
    if isinstance(value, _FrozenIdentityReference):
        return value.reference
    if isinstance(value, FrozenDataclassValue):
        instance = object.__new__(value.dataclass_type)
        for name, item in value.attribute_values:
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
        if previous[0] == current[0]:
            raise TypeError("configuration mapping keys cannot be ordered deterministically")
    return tuple((key, item) for _, key, item in keyed_items)


def _order_set_items(items: tuple[Any, ...]) -> tuple[Any, ...]:
    """Order frozen set entries and reject ambiguous structural values."""
    keyed_items = sorted((_stable_token(item), item) for item in items)
    for previous, current in pairwise(keyed_items):
        if previous[0] == current[0]:
            raise TypeError("configuration set values cannot be ordered deterministically")
    return tuple(item for _, item in keyed_items)


def _stable_token(value: Any) -> tuple[Any, ...]:  # noqa: C901
    """Build a deterministic structural ordering token without repr or ids."""
    if value is None:
        return ("none",)
    if isinstance(value, Enum):
        value_type = type(value)
        return ("enum", value_type.__module__, value_type.__qualname__, value.name)
    if type(value) is bool:
        return ("bool", int(value))
    if type(value) is int:
        return ("int", str(value))
    if type(value) is str:
        return ("str", value)
    if type(value) is bytes:
        return ("bytes", value.hex())
    if isinstance(value, _FrozenFloat):
        return ("float", value.hexadecimal)
    if isinstance(value, _FrozenComplex):
        return ("complex", _stable_token(value.real), _stable_token(value.imaginary))
    if isinstance(value, _FrozenDecimal):
        return ("decimal", value.text)
    if isinstance(value, _FrozenDate):
        return ("date", value.year, value.month, value.day)
    if isinstance(value, _FrozenTimedelta):
        return ("timedelta", value.days, value.seconds, value.microseconds)
    if isinstance(value, _FrozenDatetime):
        frozen_timezone = value.timezone
        timezone_token: tuple[Any, ...]
        if frozen_timezone is None:
            timezone_token = ("naive",)
        else:
            timezone_token = (
                "timezone",
                _stable_token(frozen_timezone.offset),
                frozen_timezone.name,
            )
        return (
            "datetime",
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            timezone_token,
            value.fold,
        )
    if isinstance(value, _FrozenTime):
        frozen_timezone = value.timezone
        timezone_token = (
            ("naive",)
            if frozen_timezone is None
            else (
                "timezone",
                _stable_token(frozen_timezone.offset),
                frozen_timezone.name,
            )
        )
        return (
            "time",
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            timezone_token,
            value.fold,
        )
    if isinstance(value, _FrozenUUID):
        return ("uuid", value.integer)
    if isinstance(value, _FrozenTypeReference):
        return (
            "type",
            value.type_module,
            value.type_qualname,
            value.type_identity,
        )
    if isinstance(value, _FrozenIdentityReference):
        return (
            "identity",
            value.type_module,
            value.type_qualname,
            value.type_identity,
            value.identity,
        )
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
        return (
            "memoryview",
            value.value.hex(),
            value.format,
            value.shape,
            value.strides,
            value.suboffsets,
            value.itemsize,
            value.ndim,
            value.readonly,
        )
    if isinstance(value, _FrozenObject):
        return (
            "object",
            value.type_module,
            value.type_qualname,
            value.type_identity,
            tuple((name, _stable_token(item)) for name, item in value.attribute_values),
        )
    if isinstance(value, FrozenDataclassValue):
        return (
            "dataclass",
            value.type_module,
            value.type_qualname,
            value.type_identity,
            tuple((name, _stable_token(item)) for name, item in value.attribute_values),
        )
    if isfunction(value) or isbuiltin(value) or _is_method_descriptor(value):
        return ("symbol", value.__module__, value.__qualname__)
    raise TypeError(
        f"configuration value cannot be ordered deterministically: {type(value).__name__}"
    )


def _object_attributes(value: Any) -> dict[str, Any]:
    """Collect mutable instance state without consulting value representations."""
    attributes = dict(vars(value)) if hasattr(value, "__dict__") else {}
    for value_type in type(value).__mro__:
        slots = value_type.__dict__.get("__slots__", ())
        if isinstance(slots, str):
            slots = (slots,)
        for name in slots:
            storage_name = _mangled_slot_name(value_type, name)
            if storage_name not in {"__dict__", "__weakref__"} and hasattr(
                value,
                storage_name,
            ):
                attributes.setdefault(storage_name, getattr(value, storage_name))
    return attributes


def _mangled_slot_name(owner: type[Any], name: str) -> str:
    """Return the storage name Python assigns to a slot declared by *owner*."""
    if name.startswith("__") and not name.endswith("__"):
        owner_name = owner.__name__.lstrip("_")
        if owner_name:
            return f"_{owner_name}{name}"
    return name


def _uses_legacy_identity_semantics(value: Any) -> bool:
    """Recognize the default identity comparison contract used by sentinels."""
    value_type = type(value)
    return (
        _raw_mro_attribute(value_type, "__eq__") is object.__eq__
        and _raw_mro_attribute(value_type, "__hash__") is object.__hash__
    )


def _safe_type_description(value_type: type[Any]) -> tuple[str, str]:
    """Capture type labels without invoking a caller-controlled metaclass."""
    try:
        module = type.__getattribute__(value_type, "__module__")
    except BaseException:
        module = "<unknown>"
    try:
        qualname = type.__getattribute__(value_type, "__qualname__")
    except BaseException:
        qualname = "<unknown>"
    return (
        module if isinstance(module, str) else "<unknown>",
        qualname if isinstance(qualname, str) else "<unknown>",
    )


def _is_safely_reconstructable_python_object(value: Any) -> bool:
    """Reject C-backed values whose complete hidden state cannot be captured."""
    return _has_object_allocation_layout(type(value), set())


def _has_object_allocation_layout(
    value_type: type[Any],
    seen: set[type[Any]],
) -> bool:
    """Accept Python allocation overrides only over object-layout base types."""
    if value_type is object or value_type in seen:
        return True
    seen.add(value_type)

    try:
        flags = type.__getattribute__(value_type, "__flags__")
    except BaseException:
        return False
    if type(flags) is not int or flags & _PY_TPFLAGS_IMMUTABLETYPE:
        return False

    namespace = vars(value_type)
    if "__new__" in namespace:
        allocator = namespace["__new__"]
        if isinstance(allocator, staticmethod):
            allocator = allocator.__func__
        if allocator is not object.__new__ and not isfunction(allocator):
            return False

    return all(_has_object_allocation_layout(base, seen) for base in value_type.__bases__)


def _raw_mro_attribute(value_type: type[Any], name: str) -> Any:
    for owner in value_type.__mro__:
        namespace = vars(owner)
        if name in namespace:
            return namespace[name]
    return None


def _is_method_descriptor(value: Any) -> bool:
    return isinstance(value, (MethodDescriptorType, WrapperDescriptorType))


def _resolve_header_rows(
    config: _ValidatedConfigProjection,
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
    config: _ValidatedConfigProjection,
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


def _resolve_skip_footer(
    config: _ValidatedConfigProjection,
    structure: StructureInfo,
) -> int:
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
    config: _ValidatedConfigProjection,
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
