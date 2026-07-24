"""Collision-safe physical provenance carried through raw string batches."""

from __future__ import annotations

import json
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal
from typing import Literal, cast
from zoneinfo import ZoneInfo

import pandas as pd
import pyarrow as pa

_PREFIX = "\x1eMESSY_XLSX_PHYSICAL_V1:"
_TEXT = "s"
_BOOL = "b"
_INT = "i"
_FLOAT = "f"
_DATETIME = "d"
_DATE = "a"
_TIME = "t"
_TIMEDELTA = "u"
_PANDAS_TIMESTAMP = "p"
_PANDAS_TIMEDELTA = "q"
_BYTES = "y"
_DECIMAL = "m"
_KNOWN_CODES = frozenset(
    {
        _TEXT,
        _BOOL,
        _INT,
        _FLOAT,
        _DATETIME,
        _DATE,
        _TIME,
        _TIMEDELTA,
        _PANDAS_TIMESTAMP,
        _PANDAS_TIMEDELTA,
        _BYTES,
        _DECIMAL,
    }
)
_TEMPORAL_UNITS = ("ns", "us", "ms", "s")
_NANOSECONDS_PER_UNIT = {
    "s": 1_000_000_000,
    "ms": 1_000_000,
    "us": 1_000,
    "ns": 1,
}
_NANOSECONDS_PER_DAY = 86_400_000_000_000
_INT64_MIN = -(1 << 63)
_INT64_MAX = (1 << 63) - 1
_PANDAS_TIMESTAMP_BASE = next(base for base in pd.Timestamp.__mro__ if "asm8" in base.__dict__)
_PANDAS_TIMEDELTA_BASE = next(base for base in pd.Timedelta.__mro__ if "asm8" in base.__dict__)
_PANDAS_TIMESTAMP_ASM8 = _PANDAS_TIMESTAMP_BASE.__dict__["asm8"]
_PANDAS_TIMESTAMP_UNIT = _PANDAS_TIMESTAMP_BASE.__dict__["unit"]
_PANDAS_TIMESTAMP_FOLD = _PANDAS_TIMESTAMP_BASE.__dict__["fold"]
_PANDAS_TIMESTAMP_COMPONENTS = tuple(
    _PANDAS_TIMESTAMP_BASE.__dict__[name]
    for name in (
        "year",
        "month",
        "day",
        "hour",
        "minute",
        "second",
        "microsecond",
        "nanosecond",
    )
)
_PANDAS_TIMEDELTA_ASM8 = _PANDAS_TIMEDELTA_BASE.__dict__["asm8"]
_PANDAS_TIMEDELTA_UNIT = _PANDAS_TIMEDELTA_BASE.__dict__["unit"]


class UnsupportedPhysicalValueError(TypeError):
    """A scalar cannot be carried through the collision-safe raw channel."""


@dataclass(frozen=True, slots=True)
class PandasTemporalPayload:
    """Hook-free physical storage for one pandas temporal scalar."""

    family: Literal["timestamp", "duration"]
    raw: int
    unit: Literal["s", "ms", "us", "ns"]
    timezone: timezone | ZoneInfo | None = field(
        default=None,
        compare=False,
        hash=False,
        repr=False,
    )
    timezone_descriptor: tuple[object, ...] | None = None
    timezone_identity: int | None = None
    fold: int = 0
    utc_offset_nanoseconds: int = field(
        default=0,
        compare=False,
        hash=False,
        repr=False,
    )


def pandas_temporal_payload(value: object) -> PandasTemporalPayload | None:
    """Snapshot pandas temporals without invoking subclass property hooks."""
    if isinstance(value, PandasTemporalPayload):
        return value
    if value is pd.NaT:
        raise UnsupportedPhysicalValueError("pandas NaT is a missing value")
    value_type = type(value)
    mro = type.__getattribute__(value_type, "__mro__")
    if any(base is pd.Timestamp for base in mro):
        unit = _PANDAS_TIMESTAMP_UNIT.__get__(value, pd.Timestamp)
        raw = int(_PANDAS_TIMESTAMP_ASM8.__get__(value, pd.Timestamp).view("i8"))
        fold = _PANDAS_TIMESTAMP_FOLD.__get__(value, pd.Timestamp)
        timezone_value = datetime.tzinfo.__get__(value, datetime)
        if unit not in _NANOSECONDS_PER_UNIT or type(fold) is not int:
            raise UnsupportedPhysicalValueError("unsupported pandas timestamp payload")
        descriptor = _trusted_timezone_descriptor(timezone_value)
        components = tuple(
            descriptor.__get__(value, pd.Timestamp) for descriptor in _PANDAS_TIMESTAMP_COMPONENTS
        )
        if any(type(component) is not int for component in components):
            raise UnsupportedPhysicalValueError("unsupported pandas timestamp components")
        year, month, day, hour, minute, second, microsecond, nanosecond = cast(
            "tuple[int, int, int, int, int, int, int, int]",
            components,
        )
        local_nanoseconds = (
            proleptic_days_from_civil(year, month, day) * _NANOSECONDS_PER_DAY
            + ((hour * 60 + minute) * 60 + second) * 1_000_000_000
            + microsecond * 1_000
            + nanosecond
        )
        utc_offset_nanoseconds = local_nanoseconds - raw * _NANOSECONDS_PER_UNIT[unit]
        return PandasTemporalPayload(
            family="timestamp",
            raw=raw,
            unit=cast("Literal['s', 'ms', 'us', 'ns']", unit),
            timezone=timezone_value,
            timezone_descriptor=descriptor,
            timezone_identity=id(timezone_value) if type(timezone_value) is ZoneInfo else None,
            fold=fold,
            utc_offset_nanoseconds=utc_offset_nanoseconds,
        )
    if any(base is pd.Timedelta for base in mro):
        unit = _PANDAS_TIMEDELTA_UNIT.__get__(value, pd.Timedelta)
        raw = int(_PANDAS_TIMEDELTA_ASM8.__get__(value, pd.Timedelta).view("i8"))
        if unit not in _NANOSECONDS_PER_UNIT:
            raise UnsupportedPhysicalValueError("unsupported pandas timedelta payload")
        return PandasTemporalPayload(
            family="duration",
            raw=raw,
            unit=cast("Literal['s', 'ms', 'us', 'ns']", unit),
        )
    return None


def temporal_payload(value: object) -> PandasTemporalPayload | None:
    """Return inert storage for a pandas or trusted exact stdlib temporal."""
    pandas_payload = pandas_temporal_payload(value)
    if pandas_payload is not None:
        return pandas_payload
    if type(value) is timedelta:
        duration = cast(timedelta, value)
        return PandasTemporalPayload(
            family="duration",
            raw=((duration.days * 86_400 + duration.seconds) * 1_000_000 + duration.microseconds),
            unit="us",
        )
    if type(value) is not datetime:
        return None
    timestamp = cast(datetime, value)
    timezone_value = datetime.tzinfo.__get__(timestamp, datetime)
    try:
        descriptor = _trusted_timezone_descriptor(timezone_value)
    except UnsupportedPhysicalValueError:
        return None
    year = datetime.year.__get__(timestamp, datetime)
    month = datetime.month.__get__(timestamp, datetime)
    adjusted_year = year - (1 if month <= 2 else 0)
    era = adjusted_year // 400
    year_of_era = adjusted_year - era * 400
    shifted_month = month + (-3 if month > 2 else 9)
    day_of_year = (153 * shifted_month + 2) // 5 + datetime.day.__get__(timestamp, datetime) - 1
    days = (
        era * 146_097
        + year_of_era * 365
        + year_of_era // 4
        - year_of_era // 100
        + day_of_year
        - 719_468
    )
    raw = (
        days * 86_400_000_000
        + (
            (
                datetime.hour.__get__(timestamp, datetime) * 60
                + datetime.minute.__get__(timestamp, datetime)
            )
            * 60
            + datetime.second.__get__(timestamp, datetime)
        )
        * 1_000_000
        + datetime.microsecond.__get__(timestamp, datetime)
    )
    timezone_identity = None
    if type(timezone_value) is timezone:
        raw -= cast(int, cast(tuple[object, ...], descriptor)[1])
    elif type(timezone_value) is ZoneInfo:
        offset = ZoneInfo.utcoffset(cast(ZoneInfo, timezone_value), timestamp)
        if type(offset) is not timedelta:
            return None
        raw -= (offset.days * 86_400 + offset.seconds) * 1_000_000 + offset.microseconds
        timezone_identity = id(timezone_value)
    return PandasTemporalPayload(
        family="timestamp",
        raw=raw,
        unit="us",
        timezone=cast("timezone | ZoneInfo | None", timezone_value),
        timezone_descriptor=descriptor,
        timezone_identity=timezone_identity,
        fold=datetime.fold.__get__(timestamp, datetime),
    )


def convert_temporal_raw(
    raw: int,
    source_unit: str,
    target_unit: str,
) -> int | None:
    """Convert an integer temporal payload exactly inside signed int64."""
    source_factor = _NANOSECONDS_PER_UNIT.get(source_unit)
    target_factor = _NANOSECONDS_PER_UNIT.get(target_unit)
    if source_factor is None or target_factor is None:
        return None
    total_nanoseconds = raw * source_factor
    converted, remainder = divmod(total_nanoseconds, target_factor)
    if remainder or converted < _INT64_MIN or converted > _INT64_MAX:
        return None
    return converted


def common_temporal_arrow_type(
    payloads: list[PandasTemporalPayload],
) -> pa.DataType | None:
    """Choose the finest exact common Arrow unit for one semantic family."""
    if not payloads:
        return None
    family = payloads[0].family
    timezone_key = pandas_temporal_arrow_timezone(payloads[0])
    if any(
        payload.family != family or pandas_temporal_arrow_timezone(payload) != timezone_key
        for payload in payloads[1:]
    ):
        return None
    has_submicrosecond_value = any(
        (payload.raw * _NANOSECONDS_PER_UNIT[payload.unit]) % 1_000 for payload in payloads
    )
    candidate_units = _TEMPORAL_UNITS if has_submicrosecond_value else ("us", "ms", "s")
    for unit in candidate_units:
        if all(
            convert_temporal_raw(payload.raw, payload.unit, unit) is not None
            for payload in payloads
        ):
            if family == "timestamp":
                return pa.timestamp(unit, tz=timezone_key)
            return pa.duration(unit)
    return None


def arrow_temporal_array(
    payloads: list[PandasTemporalPayload | None],
    target_type: pa.DataType,
    *,
    require_timezone_match: bool = True,
) -> pa.Array:
    """Construct an Arrow temporal array only from inert integer buffers."""
    if not (pa.types.is_timestamp(target_type) or pa.types.is_duration(target_type)):
        raise TypeError("target type must be an Arrow timestamp or duration")
    target_unit = cast("pa.TimestampType | pa.DurationType", target_type).unit
    expected_family = "timestamp" if pa.types.is_timestamp(target_type) else "duration"
    converted: list[int | None] = []
    for payload in payloads:
        if payload is None:
            converted.append(None)
            continue
        if payload.family != expected_family:
            raise TypeError("pandas temporal family does not match Arrow target")
        if (
            require_timezone_match
            and pa.types.is_timestamp(target_type)
            and (
                pandas_temporal_arrow_timezone(payload) != cast("pa.TimestampType", target_type).tz
            )
        ):
            raise TypeError("pandas timestamp timezone does not match Arrow target")
        raw = convert_temporal_raw(payload.raw, payload.unit, target_unit)
        if raw is None:
            raise OverflowError("pandas temporal value cannot fit the fixed Arrow unit")
        converted.append(raw)
    storage = pa.array(converted, type=pa.int64())
    return pa.Array.from_buffers(
        target_type,
        len(storage),
        storage.buffers(),
        null_count=storage.null_count,
    )


def pandas_temporal_from_payload(payload: PandasTemporalPayload) -> pd.Timestamp | pd.Timedelta:
    """Reconstruct the exact base pandas scalar represented by a payload."""
    if payload.family == "duration":
        return pd.Timedelta(payload.raw, unit=payload.unit)
    return pd.Timestamp(
        payload.raw,
        unit=payload.unit,
        tz=payload.timezone,
    )


def pandas_temporal_arrow_timezone(payload: PandasTemporalPayload) -> str | None:
    """Return the stable Arrow timezone metadata for a trusted payload."""
    descriptor = payload.timezone_descriptor
    if descriptor is None:
        return None
    if descriptor[:1] == ("zoneinfo",):
        return cast(str, descriptor[1])
    if descriptor[:1] != ("fixed",):
        raise UnsupportedPhysicalValueError("unsupported pandas timestamp timezone")
    total_microseconds = cast(int, descriptor[1])
    if total_microseconds == 0:
        return "UTC"
    total_minutes = abs(total_microseconds) // 60_000_000
    sign = "+" if total_microseconds >= 0 else "-"
    hours, minutes = divmod(total_minutes, 60)
    return f"{sign}{hours:02d}:{minutes:02d}"


def pandas_timestamp_label_text(value: object) -> str:
    """Render one pandas timestamp label without pandas or datetime conversion."""
    payload = pandas_temporal_payload(value)
    if payload is None or payload.family != "timestamp":
        raise UnsupportedPhysicalValueError("value is not a pandas timestamp")
    return pandas_timestamp_payload_label_text(payload)


def pandas_timestamp_payload_label_text(payload: PandasTemporalPayload) -> str:
    """Render readable pandas-like text directly from inert integer storage."""
    if payload.family != "timestamp":
        raise UnsupportedPhysicalValueError("payload is not a pandas timestamp")
    local_nanoseconds = (
        payload.raw * _NANOSECONDS_PER_UNIT[payload.unit] + payload.utc_offset_nanoseconds
    )
    days, nanoseconds_of_day = divmod(local_nanoseconds, _NANOSECONDS_PER_DAY)
    year, month, day = civil_from_proleptic_days(days)
    seconds_of_day, nanoseconds = divmod(nanoseconds_of_day, 1_000_000_000)
    hour, seconds_of_hour = divmod(seconds_of_day, 3_600)
    minute, second = divmod(seconds_of_hour, 60)
    year_text = f"-{abs(year):03d}" if year < 0 else f"{year:04d}"
    text = f"{year_text}-{month:02d}-{day:02d} {hour:02d}:{minute:02d}:{second:02d}"
    if nanoseconds:
        text = f"{text}.{nanoseconds:09d}".rstrip("0")
    if payload.timezone_descriptor is not None:
        text = f"{text}{_format_utc_offset(payload.utc_offset_nanoseconds)}"
    return text


def proleptic_days_from_civil(year: int, month: int, day: int) -> int:
    """Return proleptic Gregorian days from 1970-01-01 for any integer year."""
    adjusted_year = year - (1 if month <= 2 else 0)
    era = adjusted_year // 400
    year_of_era = adjusted_year - era * 400
    shifted_month = month + (-3 if month > 2 else 9)
    day_of_year = (153 * shifted_month + 2) // 5 + day - 1
    day_of_era = year_of_era * 365 + year_of_era // 4 - year_of_era // 100 + day_of_year
    return era * 146_097 + day_of_era - 719_468


def civil_from_proleptic_days(days: int) -> tuple[int, int, int]:
    """Invert :func:`proleptic_days_from_civil` for any integer day."""
    shifted_days = days + 719_468
    era = shifted_days // 146_097
    day_of_era = shifted_days - era * 146_097
    year_of_era = (
        day_of_era - day_of_era // 1_460 + day_of_era // 36_524 - day_of_era // 146_096
    ) // 365
    year = year_of_era + era * 400
    day_of_year = day_of_era - (365 * year_of_era + year_of_era // 4 - year_of_era // 100)
    month_prime = (5 * day_of_year + 2) // 153
    day = day_of_year - (153 * month_prime + 2) // 5 + 1
    month = month_prime + (3 if month_prime < 10 else -9)
    year += 1 if month <= 2 else 0
    return year, month, day


def _format_utc_offset(offset_nanoseconds: int) -> str:
    sign = "+" if offset_nanoseconds >= 0 else "-"
    seconds, nanoseconds = divmod(abs(offset_nanoseconds), 1_000_000_000)
    hours, seconds_of_hour = divmod(seconds, 3_600)
    minutes, seconds = divmod(seconds_of_hour, 60)
    text = f"{sign}{hours:02d}:{minutes:02d}"
    if seconds or nanoseconds:
        text = f"{text}:{seconds:02d}"
        if nanoseconds:
            text = f"{text}.{nanoseconds:09d}".rstrip("0")
    return text


def _decode_boolean(payload: str) -> bool:
    if payload == "True":
        return True
    if payload == "False":
        return False
    raise ValueError("invalid internal boolean provenance")


def _decode_timedelta(payload: str) -> timedelta:
    return timedelta(microseconds=int(payload))


def _encode_timedelta(value: object) -> str:
    duration = cast(timedelta, value)
    microseconds = (duration.days * 86_400 + duration.seconds) * 1_000_000 + duration.microseconds
    return str(microseconds)


def _encode_bytearray(value: object) -> str:
    return bytes(cast(bytearray, value)).hex()


def _encode_memoryview(value: object) -> str:
    return cast(memoryview, value).tobytes().hex()


def _encode_pandas_timestamp(value: object) -> str:
    temporal = pandas_temporal_payload(value)
    if temporal is None or temporal.family != "timestamp":
        raise UnsupportedPhysicalValueError("unsupported pandas timestamp payload")
    return json.dumps(
        [
            temporal.unit,
            temporal.raw,
            _timezone_descriptor_payload(temporal.timezone_descriptor),
            temporal.fold,
        ],
        ensure_ascii=True,
        separators=(",", ":"),
    )


def _decode_pandas_timestamp(payload: str) -> pd.Timestamp:
    try:
        decoded = json.loads(payload)
    except (TypeError, ValueError) as error:
        raise ValueError("invalid internal pandas timestamp provenance") from error
    if (
        type(decoded) is not list
        or len(decoded) not in {3, 4}
        or type(decoded[0]) is not str
        or decoded[0] not in {"s", "ms", "us", "ns"}
        or type(decoded[1]) is not int
    ):
        raise ValueError("invalid internal pandas timestamp provenance")
    timezone_value = _decode_trusted_timezone_payload(decoded[2])
    return pd.Timestamp(decoded[1], unit=decoded[0], tz=timezone_value)


def _decode_pandas_timedelta(payload: str) -> pd.Timedelta:
    try:
        decoded = json.loads(payload)
        if (
            type(decoded) is list
            and len(decoded) == 2
            and type(decoded[0]) is str
            and decoded[0] in _NANOSECONDS_PER_UNIT
            and type(decoded[1]) is int
        ):
            return pd.Timedelta(decoded[1], unit=decoded[0])
        nanoseconds = int(payload)
    except (TypeError, ValueError) as error:
        raise ValueError("invalid internal pandas timedelta provenance") from error
    return pd.Timedelta(nanoseconds, unit="ns")


def _trusted_timezone_descriptor(value: object) -> tuple[object, ...] | None:
    if value is None:
        return None
    if type(value) is ZoneInfo:
        key = ZoneInfo.__dict__["key"].__get__(cast(ZoneInfo, value), ZoneInfo)
        if type(key) is not str:
            raise UnsupportedPhysicalValueError("unsafe ZoneInfo key")
        return ("zoneinfo", key)
    if type(value) is timezone:
        timezone_value = cast(timezone, value)
        offset = timezone.utcoffset(timezone_value, None)
        name = timezone.tzname(timezone_value, None)
        if type(offset) is not timedelta or type(name) is not str:
            raise UnsupportedPhysicalValueError("unsafe fixed timezone payload")
        total_microseconds = (
            offset.days * 86_400 + offset.seconds
        ) * 1_000_000 + offset.microseconds
        if total_microseconds % 60_000_000:
            raise UnsupportedPhysicalValueError(
                "sub-minute timestamp timezone offsets are not losslessly representable"
            )
        return ("fixed", total_microseconds, name)
    raise UnsupportedPhysicalValueError("unsupported pandas timestamp timezone")


def _trusted_timezone_payload(value: object) -> list[object] | None:
    return _timezone_descriptor_payload(_trusted_timezone_descriptor(value))


def _timezone_descriptor_payload(
    descriptor: tuple[object, ...] | None,
) -> list[object] | None:
    return None if descriptor is None else list(descriptor)


def _decode_trusted_timezone_payload(value: object) -> timezone | ZoneInfo | None:
    if value is None:
        return None
    if type(value) is not list or len(value) < 2 or type(value[0]) is not str:
        raise ValueError("invalid internal pandas timestamp timezone")
    if value[0] == "zoneinfo" and len(value) == 2 and type(value[1]) is str:
        return ZoneInfo(value[1])
    if value[0] == "fixed" and len(value) == 3 and type(value[1]) is int and type(value[2]) is str:
        return timezone(timedelta(microseconds=value[1]), value[2])
    raise ValueError("invalid internal pandas timestamp timezone")


_DECODERS: dict[str, Callable[[str], object]] = {
    _BOOL: _decode_boolean,
    _INT: int,
    _FLOAT: float.fromhex,
    _DATETIME: datetime.fromisoformat,
    _DATE: date.fromisoformat,
    _TIME: time.fromisoformat,
    _TIMEDELTA: _decode_timedelta,
    _PANDAS_TIMESTAMP: _decode_pandas_timestamp,
    _PANDAS_TIMEDELTA: _decode_pandas_timedelta,
    _BYTES: bytes.fromhex,
    _DECIMAL: Decimal,
}

_ENCODERS: dict[type[object], tuple[str, Callable[[object], str]]] = {
    bool: (_BOOL, lambda value: str(cast(bool, value))),
    int: (_INT, lambda value: str(cast(int, value))),
    float: (_FLOAT, lambda value: cast(float, value).hex()),
    datetime: (_DATETIME, lambda value: cast(datetime, value).isoformat()),
    date: (_DATE, lambda value: cast(date, value).isoformat()),
    time: (_TIME, lambda value: cast(time, value).isoformat()),
    timedelta: (_TIMEDELTA, _encode_timedelta),
    pd.Timestamp: (_PANDAS_TIMESTAMP, _encode_pandas_timestamp),
    pd.Timedelta: (
        _PANDAS_TIMEDELTA,
        lambda value: json.dumps(
            [
                cast(PandasTemporalPayload, pandas_temporal_payload(value)).unit,
                cast(PandasTemporalPayload, pandas_temporal_payload(value)).raw,
            ],
            separators=(",", ":"),
        ),
    ),
    bytes: (_BYTES, lambda value: cast(bytes, value).hex()),
    bytearray: (_BYTES, _encode_bytearray),
    memoryview: (_BYTES, _encode_memoryview),
    Decimal: (_DECIMAL, lambda value: str(cast(Decimal, value))),
}
_EXACT_SUPPORTED_PHYSICAL_TYPES = tuple(_ENCODERS)
_CANONICAL_PHYSICAL_BASES = (
    str,
    bytes,
    bool,
    int,
    float,
    datetime,
    date,
    time,
    timedelta,
    Decimal,
)
_SIMPLE_PHYSICAL_FAMILIES: dict[type[object], tuple[str, None]] = {
    bool: ("bool", None),
    int: ("integer", None),
    float: ("floating", None),
    bytes: ("binary", None),
    bytearray: ("binary", None),
    memoryview: ("binary", None),
    date: ("date", None),
    timedelta: ("duration", None),
    pd.Timedelta: ("duration", None),
    Decimal: ("decimal", None),
}
_SIMPLE_ARROW_FAMILIES: tuple[
    tuple[Callable[[pa.DataType], bool], str],
    ...,
] = (
    (pa.types.is_boolean, "bool"),
    (pa.types.is_integer, "integer"),
    (pa.types.is_floating, "floating"),
    (pa.types.is_string, "string"),
    (pa.types.is_large_string, "string"),
    (pa.types.is_binary, "binary"),
    (pa.types.is_large_binary, "binary"),
    (pa.types.is_date, "date"),
    (pa.types.is_duration, "duration"),
    (pa.types.is_decimal, "decimal"),
)


def encode_physical_value(value: object) -> str | None:
    """Encode one openpyxl scalar without losing original-text provenance."""
    if value is None:
        return None
    ensure_supported_physical_value(value)
    temporal = pandas_temporal_payload(value)
    if temporal is not None:
        code = _PANDAS_TIMESTAMP if temporal.family == "timestamp" else _PANDAS_TIMEDELTA
        if temporal.family == "timestamp":
            payload = _encode_pandas_timestamp(temporal)
        else:
            payload = json.dumps([temporal.unit, temporal.raw], separators=(",", ":"))
        return f"{_PREFIX}{code}{payload}"
    value_type = type(value)
    if isinstance(value, str):
        text = value if value_type is str else str.__str__(value)
        return f"{_PREFIX}{_TEXT}{text}" if text.startswith(_PREFIX) else text
    encoder = _ENCODERS.get(value_type)
    if encoder is None:
        raise UnsupportedPhysicalValueError("unsupported physical scalar type")
    code, encode_payload = encoder
    return f"{_PREFIX}{code}{encode_payload(value)}"


def ensure_supported_physical_value(value: object) -> None:
    """Reject Python scalars whose semantics Arrow cannot preserve."""
    if value is None:
        return
    if pandas_temporal_payload(value) is not None:
        return
    value_type = type(value)
    mro = type.__getattribute__(value_type, "__mro__")
    if any(base is time for base in mro) and time.tzinfo.__get__(value, time) is not None:
        raise UnsupportedPhysicalValueError(
            "timezone-bearing datetime.time values are not losslessly representable"
        )
    if any(base is datetime for base in mro):
        timezone_value = datetime.tzinfo.__get__(value, datetime)
        _trusted_timezone_descriptor(timezone_value)
    if any(value_type is supported for supported in _EXACT_SUPPORTED_PHYSICAL_TYPES):
        return
    if any(base is supported for base in mro for supported in _CANONICAL_PHYSICAL_BASES):
        return
    raise UnsupportedPhysicalValueError("unsupported physical scalar type")


def physical_text_body(value: str) -> str:
    """Return the lexical payload used by Arrow's vectorized casts."""
    tagged = _tagged_parts(value)
    return value if tagged is None else tagged[1]


def physical_normalization_value(value: object) -> object:
    """Project transport provenance to a value suitable for normalization.

    Transport payloads favor lossless round trips (for example ``float.hex``
    and ``bytes.hex``). They are deliberately not the lexical representation
    consumed by normalization.
    """
    decoded = decode_physical_value(value)
    temporal = pandas_temporal_payload(decoded)
    if temporal is not None:
        decoded = pandas_temporal_from_payload(temporal)
    if isinstance(decoded, str):
        return decoded
    value_type = type(decoded)
    if value_type in {bool, int, float, Decimal}:
        return str(decoded)
    if value_type in {datetime, date, time}:
        return cast(datetime | date | time, decoded).isoformat()
    if value_type is pd.Timestamp:
        return cast("pd.Timestamp", decoded).isoformat()
    if value_type is timedelta:
        return str(cast(timedelta, decoded))
    if value_type is pd.Timedelta:
        return pd.Timedelta.__str__(cast(pd.Timedelta, decoded))
    raise UnsupportedPhysicalValueError("physical scalar has no lossless normalization lexeme")


def physical_value_is_original_text(value: str) -> bool:
    """Return whether a raw string originated as an Excel text scalar."""
    tagged = _tagged_parts(value)
    return tagged is None or tagged[0] == _TEXT


def decode_physical_value(value: object) -> object:
    """Recover a bounded sample scalar from its raw physical tag."""
    if not isinstance(value, str):
        return value
    tagged = _tagged_parts(value)
    if tagged is None:
        return value
    code, payload = tagged
    if code == _TEXT:
        return payload
    decoder = _DECODERS.get(code)
    if decoder is None:
        raise ValueError("invalid internal physical provenance")
    return decoder(payload)


def physical_value_description(value: object) -> str:
    """Return the safe structural description for one raw physical scalar."""
    if not isinstance(value, str):
        try:
            temporal = pandas_temporal_payload(value)
        except UnsupportedPhysicalValueError:
            temporal = None
        if temporal is not None:
            return "datetime" if temporal.family == "timestamp" else "timedelta"
        value_type = type(value)
        if value_type is pd.Timestamp:
            return "datetime"
        if value_type is pd.Timedelta:
            return "timedelta"
        if value_type in {bool, int, float, date, datetime, time}:
            return value_type.__name__
        return "unsupported value"
    tagged = _tagged_parts(value)
    if tagged is None:
        return f"str(length={len(value)})"
    code, payload = tagged
    if code == _TEXT:
        return f"str(length={len(payload)})"
    return {
        _BOOL: "bool",
        _INT: "int",
        _FLOAT: "float",
        _DATETIME: "datetime",
        _DATE: "date",
        _TIME: "time",
        _TIMEDELTA: "timedelta",
        _PANDAS_TIMESTAMP: "datetime",
        _PANDAS_TIMEDELTA: "timedelta",
        _BYTES: f"bytes(length={len(payload) // 2})",
        _DECIMAL: "unsupported value",
    }[code]


def physical_label_description(value: object) -> str:
    """Describe a display label without invoking user-controlled hooks."""
    value_type = type(value)
    if value_type is str:
        return f"str label(length={len(cast(str, value))})"
    if value_type is int or value_type is float or value_type is bool:
        return f"{value_type.__name__} label"
    return "non-string label"


def physical_value_family(value: object) -> tuple[str, str | None] | None:
    """Classify supported scalars without allocating a singleton Arrow array."""
    ensure_supported_physical_value(value)
    temporal = pandas_temporal_payload(value)
    if temporal is not None:
        if temporal.family == "timestamp":
            return ("timestamp", pandas_temporal_arrow_timezone(temporal))
        return ("duration", None)
    value_type = type(value)
    if isinstance(value, str):
        return ("string", None)
    if value_type is datetime:
        return ("timestamp", _timezone_key(cast(datetime, value)))
    if value_type is pd.Timestamp:
        return ("timestamp", _pandas_timezone_key(cast(pd.Timestamp, value)))
    if value_type is time:
        return ("time", _timezone_key(cast(time, value)))
    return _SIMPLE_PHYSICAL_FAMILIES.get(value_type)


def physical_value_matches_arrow_type(
    value: object,
    target: pa.DataType,
) -> bool:
    """Check the exact physical family accepted by a fixed Arrow field."""
    family = physical_value_family(value)
    if family is None:
        return False
    kind, timezone_key = family
    if pa.types.is_timestamp(target):
        return kind == "timestamp" and timezone_key == cast(pa.TimestampType, target).tz
    if pa.types.is_time(target):
        return kind == "time" and timezone_key is None
    for predicate, expected_kind in _SIMPLE_ARROW_FAMILIES:
        if predicate(target):
            return kind == expected_kind
    return False


def _timezone_key(value: datetime | time) -> str | None:
    timezone = value.tzinfo
    offset = value.utcoffset()
    if timezone is None or offset is None:
        return None
    zone_key = getattr(timezone, "key", None)
    if type(zone_key) is str:
        return zone_key
    name = value.tzname()
    if offset == timedelta(0) and name == "UTC":
        return "UTC"
    total_seconds = int(offset.total_seconds())
    sign = "+" if total_seconds >= 0 else "-"
    total_seconds = abs(total_seconds)
    hours, remainder = divmod(total_seconds, 3_600)
    minutes, seconds = divmod(remainder, 60)
    suffix = f":{seconds:02d}" if seconds else ""
    return f"{sign}{hours:02d}:{minutes:02d}{suffix}"


def _pandas_timezone_key(value: pd.Timestamp) -> str | None:
    payload = _trusted_timezone_payload(value.tzinfo)
    if payload is None:
        return None
    if payload[0] == "zoneinfo":
        return cast(str, payload[1])
    total_microseconds = cast(int, payload[1])
    if total_microseconds == 0:
        return "UTC"
    total_minutes, remainder = divmod(abs(total_microseconds), 60_000_000)
    if remainder:
        raise UnsupportedPhysicalValueError("sub-minute timestamp timezone offset")
    sign = "+" if total_microseconds >= 0 else "-"
    hours, minutes = divmod(total_minutes, 60)
    return f"{sign}{hours:02d}:{minutes:02d}"


def _tagged_parts(value: str) -> tuple[str, str] | None:
    if not value.startswith(_PREFIX):
        return None
    suffix = value[len(_PREFIX) :]
    if not suffix or suffix[0] not in _KNOWN_CODES:
        raise ValueError("invalid internal physical provenance")
    return suffix[0], suffix[1:]
