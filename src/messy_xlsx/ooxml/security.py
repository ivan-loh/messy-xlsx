"""ZIP and XML security boundaries for OOXML metadata inspection."""

from __future__ import annotations

import xml.etree.ElementTree as ElementTree
from collections.abc import Iterator
from pathlib import PurePosixPath
from typing import IO
from zipfile import ZipFile, ZipInfo

from defusedxml import ElementTree as SafeElementTree
from defusedxml.common import DefusedXmlException

from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import OoxmlLimits

_XML_PART_SUFFIXES = (".xml", ".rels")


def _safe_member(name: str) -> bool:
    path = PurePosixPath(name)
    drive_like = bool(path.parts and ":" in path.parts[0])
    return bool(
        name
        and "\x00" not in name
        and "\\" not in name
        and not drive_like
        and not path.is_absolute()
        and ".." not in path.parts
    )


def _raise_unsafe_member(member: ZipInfo) -> None:
    raise FormatError(
        "OOXML archive contains unsafe archive path",
        member=member.filename,
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
        if not _safe_member(member.filename):
            _raise_unsafe_member(member)
        if member.filename in seen:
            raise FormatError(
                "OOXML archive contains duplicate member names",
                member=member.filename,
            )
        seen.add(member.filename)

    total = sum(member.file_size for member in members)
    if total > limits.max_total_uncompressed:
        raise FormatError(
            "OOXML archive exceeds total uncompressed limit",
            uncompressed=total,
            limit=limits.max_total_uncompressed,
        )

    for member in members:
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
    except (ElementTree.ParseError, UnicodeError, ValueError) as error:
        raise FormatError(
            "OOXML package contains malformed XML",
            member=member,
        ) from error
