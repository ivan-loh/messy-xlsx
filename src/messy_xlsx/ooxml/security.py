"""ZIP and XML security boundaries for OOXML metadata inspection."""

from __future__ import annotations

import lzma
import re
import xml.etree.ElementTree as ElementTree
import zipfile
import zlib
from collections.abc import Iterator
from pathlib import PurePosixPath
from typing import IO, NoReturn
from urllib.parse import unquote
from zipfile import ZipFile, ZipInfo

from defusedxml import ElementTree as SafeElementTree
from defusedxml.common import DefusedXmlException

from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import OoxmlLimits

_XML_PART_SUFFIXES = (".xml", ".rels")
_INVALID_PERCENT_ESCAPE = re.compile(r"%(?![0-9A-Fa-f]{2})")
_SUPPORTED_COMPRESSIONS = {
    zipfile.ZIP_STORED,
    zipfile.ZIP_DEFLATED,
    zipfile.ZIP_BZIP2,
    zipfile.ZIP_LZMA,
}
if hasattr(zipfile, "ZIP_ZSTANDARD"):
    _SUPPORTED_COMPRESSIONS.add(zipfile.ZIP_ZSTANDARD)


def _member_identity(name: str) -> tuple[str, bool] | None:
    """Return a decoded comparison key and whether the raw path is canonical."""
    path = PurePosixPath(name)
    raw_drive_like = bool(path.parts and ":" in path.parts[0])
    if (
        not name
        or "\x00" in name
        or "\\" in name
        or raw_drive_like
        or path.is_absolute()
        or _INVALID_PERCENT_ESCAPE.search(name)
    ):
        return None

    canonical = True
    decoded_parts: list[str] = []
    for raw_part in name.split("/"):
        if raw_part in {"", "."}:
            canonical = False
            continue
        try:
            decoded = unquote(raw_part, encoding="utf-8", errors="strict")
        except UnicodeError:
            return None
        if decoded == ".." or any(character in decoded for character in ("\x00", "/", "\\")):
            return None
        if decoded == ".":
            canonical = False
            continue
        decoded_parts.append(decoded)

    if not decoded_parts or ":" in decoded_parts[0]:
        return None
    return "/".join(decoded_parts), canonical


def canonical_archive_name(name: str) -> str:
    """Return a safe canonical comparison key for a validated archive member."""
    identity = _member_identity(name)
    if identity is None or not identity[1]:
        raise ValueError(f"unsafe OOXML archive member: {name!r}")
    return identity[0]


def _raise_unsafe_member(member: ZipInfo) -> NoReturn:
    raise FormatError(
        "OOXML archive contains unsafe archive path",
        member=member.filename,
    )


def _validate_member_metadata(member: ZipInfo, limits: OoxmlLimits) -> None:
    if member.flag_bits & 0x1:
        raise FormatError(
            "OOXML archive contains encrypted member",
            member=member.filename,
        )
    if member.compress_type not in _SUPPORTED_COMPRESSIONS:
        raise FormatError(
            "OOXML archive member uses unsupported compression",
            member=member.filename,
            compression=member.compress_type,
        )
    if (
        member.filename.lower().endswith(_XML_PART_SUFFIXES)
        and member.file_size > limits.max_xml_uncompressed
    ):
        raise FormatError(
            "OOXML XML member exceeds size limit",
            member=member.filename,
            uncompressed=member.file_size,
            limit=limits.max_xml_uncompressed,
        )

    compression_ratio = member.file_size / max(member.compress_size, 1)
    if (
        member.file_size > limits.suspicious_ratio_size
        and compression_ratio > limits.max_compression_ratio
    ):
        raise FormatError(
            "OOXML member has suspicious compression ratio",
            member=member.filename,
            uncompressed=member.file_size,
            compressed=member.compress_size,
            compression_ratio=compression_ratio,
            limit=limits.max_compression_ratio,
        )


def validate_archive(package: ZipFile, limits: OoxmlLimits) -> None:
    """Validate declared ZIP metadata before opening any XML member."""
    members = package.infolist()
    if len(members) > limits.max_members:
        raise FormatError(
            "OOXML archive exceeds member limit",
            member_count=len(members),
            limit=limits.max_members,
        )

    seen: set[str] = set()
    for member in members:
        identity = _member_identity(member.filename)
        if identity is None:
            _raise_unsafe_member(member)
        canonical_name, is_canonical = identity
        if canonical_name in seen:
            raise FormatError(
                "OOXML archive contains duplicate member names",
                member=member.filename,
            )
        if not is_canonical:
            _raise_unsafe_member(member)
        seen.add(canonical_name)

    total = sum(member.file_size for member in members)
    if total > limits.max_total_uncompressed:
        raise FormatError(
            "OOXML archive exceeds total uncompressed limit",
            uncompressed=total,
            limit=limits.max_total_uncompressed,
        )

    for member in members:
        _validate_member_metadata(member, limits)


def reject_unsafe_xml_prefix(prefix: bytes, member: str) -> None:
    """Reject declarations visible in a bounded XML prefix."""
    lowered = prefix.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise FormatError(
            "OOXML XML declarations are not allowed",
            member=member,
        )


def safe_iterparse(
    source: IO[bytes],
    member: str,
    limits: OoxmlLimits,
) -> Iterator[tuple[str, ElementTree.Element]]:
    """Incrementally parse one XML part with entity and resource defenses."""
    depth = 0
    try:
        events = SafeElementTree.iterparse(
            source,
            events=("start", "end"),
            forbid_dtd=True,
            forbid_entities=True,
            forbid_external=True,
        )
        for event, element in events:
            if event == "start":
                depth += 1
                if depth > limits.max_xml_depth:
                    raise FormatError(
                        "OOXML XML exceeds depth limit",
                        member=member,
                        depth=depth,
                        limit=limits.max_xml_depth,
                    )
                if len(element.attrib) > limits.max_element_attributes:
                    raise FormatError(
                        "OOXML element exceeds attribute limit",
                        member=member,
                        attributes=len(element.attrib),
                        limit=limits.max_element_attributes,
                    )
            else:
                text_size = max(len(element.text or ""), len(element.tail or ""))
                if text_size > limits.max_element_text:
                    raise FormatError(
                        "OOXML element exceeds text limit",
                        member=member,
                        text_size=text_size,
                        limit=limits.max_element_text,
                    )
                depth -= 1
            yield event, element
    except FormatError:
        raise
    except DefusedXmlException as error:
        raise FormatError(
            "OOXML XML declarations are not allowed",
            member=member,
        ) from error
    except (
        zipfile.BadZipFile,
        zlib.error,
        lzma.LZMAError,
        EOFError,
        NotImplementedError,
        RuntimeError,
        OSError,
    ) as error:
        raise FormatError(
            "OOXML archive member cannot be read",
            member=member,
        ) from error
    except (ElementTree.ParseError, UnicodeError, ValueError) as error:
        raise FormatError(
            "OOXML package contains malformed XML",
            member=member,
        ) from error
