"""Security contracts for bounded OOXML metadata inspection."""

from __future__ import annotations

import io
import warnings
import zipfile
from typing import cast

import pytest

from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import OoxmlLimits
from messy_xlsx.ooxml.security import safe_iterparse, validate_archive


class _MetadataArchive:
    """Minimal archive index used to test hostile declared ZIP metadata."""

    def __init__(self, members: list[zipfile.ZipInfo]) -> None:
        self._members = members

    def infolist(self) -> list[zipfile.ZipInfo]:
        return self._members


def _member(name: str, *, size: int = 1, compressed_size: int = 1) -> zipfile.ZipInfo:
    member = zipfile.ZipInfo("placeholder")
    member.filename = name
    member.file_size = size
    member.compress_size = compressed_size
    return member


def _metadata_archive(*members: zipfile.ZipInfo) -> zipfile.ZipFile:
    return cast(zipfile.ZipFile, _MetadataArchive(list(members)))


def _archive(entries: list[tuple[str, bytes]]) -> zipfile.ZipFile:
    raw = io.BytesIO()
    with (
        zipfile.ZipFile(raw, "w", zipfile.ZIP_DEFLATED) as package,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore", UserWarning)
        for name, value in entries:
            package.writestr(name, value)
    raw.seek(0)
    return zipfile.ZipFile(raw)


def test_duplicate_members_are_rejected() -> None:
    with (
        _archive([("xl/workbook.xml", b"a"), ("xl/workbook.xml", b"b")]) as package,
        pytest.raises(FormatError, match="duplicate"),
    ):
        validate_archive(package, OoxmlLimits())


@pytest.mark.parametrize(
    "name",
    [
        "../escape.xml",
        "/absolute.xml",
        "xl/../../escape.xml",
        "C:/drive.xml",
        "C:\\drive.xml",
        "\\\\server\\share.xml",
        "xl\\workbook.xml",
        "xl/bad\x00name.xml",
        "",
    ],
)
def test_unsafe_member_paths_are_rejected(name: str) -> None:
    package = _metadata_archive(_member(name))

    with pytest.raises(FormatError, match="unsafe archive path") as raised:
        validate_archive(package, OoxmlLimits())

    assert raised.value.context["member"] == name


def test_member_count_limit_uses_declared_archive_index() -> None:
    package = _metadata_archive(_member("one.xml"), _member("two.xml"))

    with pytest.raises(FormatError, match="member limit") as raised:
        validate_archive(package, OoxmlLimits(max_members=1))

    assert raised.value.context["member_count"] == 2
    assert raised.value.context["limit"] == 1


def test_total_uncompressed_limit_uses_declared_sizes() -> None:
    package = _metadata_archive(
        _member("one.bin", size=4),
        _member("two.bin", size=5),
    )

    with pytest.raises(FormatError, match="total uncompressed limit") as raised:
        validate_archive(package, OoxmlLimits(max_total_uncompressed=8))

    assert raised.value.context["uncompressed"] == 9
    assert raised.value.context["limit"] == 8


@pytest.mark.parametrize("name", ["xl/workbook.xml", "xl/_rels/workbook.xml.rels"])
def test_every_xml_part_type_uses_the_per_member_limit(name: str) -> None:
    package = _metadata_archive(_member(name, size=9))

    with pytest.raises(FormatError, match="XML member exceeds size limit") as raised:
        validate_archive(package, OoxmlLimits(max_xml_uncompressed=8))

    assert raised.value.context["member"] == name
    assert raised.value.context["uncompressed"] == 9


def test_suspicious_compression_ratio_is_rejected_from_metadata() -> None:
    package = _metadata_archive(_member("xl/workbook.xml", size=101, compressed_size=1))
    limits = OoxmlLimits(suspicious_ratio_size=100, max_compression_ratio=100.0)

    with pytest.raises(FormatError, match="suspicious compression ratio") as raised:
        validate_archive(package, limits)

    assert raised.value.context["member"] == "xl/workbook.xml"
    assert raised.value.context["compression_ratio"] == 101.0


@pytest.mark.parametrize(
    "xml",
    [
        b"<!DOCTYPE root><root/>",
        b'<!DOCTYPE root [<!ENTITY x "payload">]><root>&x;</root>',
        b"<!DOCTYPE root [<!ENTITY x SYSTEM 'file:///etc/passwd'>]><root>&x;</root>",
    ],
)
def test_dtd_entities_and_external_entities_become_format_errors(xml: bytes) -> None:
    with pytest.raises(FormatError, match="XML declarations") as raised:
        list(safe_iterparse(io.BytesIO(xml), "xl/workbook.xml", OoxmlLimits()))

    assert raised.value.context["member"] == "xl/workbook.xml"
    assert raised.value.__cause__ is not None


def test_malformed_xml_becomes_format_error_with_preserved_cause() -> None:
    with pytest.raises(FormatError, match="malformed XML") as raised:
        list(safe_iterparse(io.BytesIO(b"<root><child></root>"), "bad.xml", OoxmlLimits()))

    assert raised.value.context["member"] == "bad.xml"
    assert raised.value.__cause__ is not None


@pytest.mark.parametrize(
    ("xml", "limits", "message"),
    [
        (
            b"<root><child><leaf/></child></root>",
            OoxmlLimits(max_xml_depth=2),
            "depth limit",
        ),
        (
            b'<root first="1" second="2"/>',
            OoxmlLimits(max_element_attributes=1),
            "attribute limit",
        ),
        (b"<root>12345</root>", OoxmlLimits(max_element_text=4), "text limit"),
        (b"<root><child/>12345</root>", OoxmlLimits(max_element_text=4), "text limit"),
    ],
)
def test_xml_security_budgets_use_caller_supplied_limits(
    xml: bytes,
    limits: OoxmlLimits,
    message: str,
) -> None:
    with pytest.raises(FormatError, match=message) as raised:
        list(safe_iterparse(io.BytesIO(xml), "bounded.xml", limits))

    assert raised.value.context["member"] == "bounded.xml"
