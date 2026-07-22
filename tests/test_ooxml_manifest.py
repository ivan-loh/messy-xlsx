"""Eager OOXML workbook-manifest contracts."""

from __future__ import annotations

import io
import lzma
import struct
import zipfile
import zlib
from dataclasses import FrozenInstanceError
from typing import Final

import pytest

import messy_xlsx.ooxml.manifest as manifest_module
from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.manifest import build_manifest
from messy_xlsx.ooxml.models import OoxmlLimits, StyleManifest

WORKBOOK_CONTENT_TYPES: Final[dict[str, str]] = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
CONTENT_TYPES_NS: Final = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_RELATIONSHIPS_NS: Final = "http://schemas.openxmlformats.org/package/2006/relationships"
SPREADSHEET_NS: Final = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_RELATIONSHIPS_NS: Final = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
STRICT_CONTENT_TYPES_NS: Final = "http://purl.oclc.org/ooxml/package/content-types"
STRICT_PACKAGE_RELATIONSHIPS_NS: Final = "http://purl.oclc.org/ooxml/package/relationships"
STRICT_SPREADSHEET_NS: Final = "http://purl.oclc.org/ooxml/spreadsheetml/main"
STRICT_OFFICE_RELATIONSHIPS_NS: Final = "http://purl.oclc.org/ooxml/officeDocument/relationships"
WORKSHEET_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
STYLES_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
)
SHARED_STRINGS_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)


def _content_types(workbook_type: str = "xlsx") -> bytes:
    content_type = WORKBOOK_CONTENT_TYPES[workbook_type]
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="{content_type}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="{WORKSHEET_CONTENT_TYPE}"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="{WORKSHEET_CONTENT_TYPE}"/>
  <Override PartName="/xl/styles.xml" ContentType="{STYLES_CONTENT_TYPE}"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="{SHARED_STRINGS_CONTENT_TYPE}"/>
</Types>
""".encode()


def _workbook(*, date1904: str | None = "1") -> bytes:
    date_attribute = "" if date1904 is None else f' date1904="{date1904}"'
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr{date_attribute}/>
  <sheets>
    <sheet name="Second" sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="First" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""".encode()


def _relationships(
    *,
    first_target: str = "worksheets/sheet1.xml",
    second_target: str = "worksheets/sheet2.xml",
    external_target: str = "externalLinks/externalLink1.xml",
) -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="{second_target}"/>
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="{first_target}"/>
  <Relationship Id="rIdStyles"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rIdShared"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    Target="sharedStrings.xml"/>
  <Relationship Id="rIdExternal"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
    Target="{external_target}" TargetMode="External"/>
</Relationships>
""".encode()


def _styles(*, extra: str = "") -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="166" formatCode="0.00"/>
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <cellXfs count="3">
    <xf numFmtId="0"/>
    <xf numFmtId="165"/>
    <xf numFmtId="14"/>
  </cellXfs>
  {extra}
</styleSheet>
""".encode()


def _entries(
    *,
    workbook_type: str = "xlsx",
    workbook: bytes | None = None,
    relationships: bytes | None = None,
    styles: bytes | None = None,
) -> dict[str, bytes]:
    return {
        "[Content_Types].xml": _content_types(workbook_type),
        "xl/workbook.xml": _workbook() if workbook is None else workbook,
        "xl/_rels/workbook.xml.rels": (
            _relationships() if relationships is None else relationships
        ),
        "xl/styles.xml": _styles() if styles is None else styles,
        "xl/sharedStrings.xml": b"<sst>secret shared-string value</sst>",
        "xl/worksheets/sheet1.xml": b"<worksheet>secret first value</worksheet>",
        "xl/worksheets/sheet2.xml": b"<worksheet>secret second value</worksheet>",
        "xl/externalLinks/externalLink1.xml": b"<externalLink>do not read</externalLink>",
    }


def _package(
    entries: dict[str, bytes],
    compression: int = zipfile.ZIP_DEFLATED,
) -> bytes:
    raw = io.BytesIO()
    with zipfile.ZipFile(raw, "w", compression) as package:
        for name, value in entries.items():
            package.writestr(name, value)
    return raw.getvalue()


def _corrupt_compressed_member(content: bytes, member: str) -> bytes:
    corrupted = bytearray(content)
    with zipfile.ZipFile(io.BytesIO(content)) as package:
        member_info = package.getinfo(member)
    filename_size, extra_size = struct.unpack_from(
        "<HH",
        corrupted,
        member_info.header_offset + 26,
    )
    payload_offset = member_info.header_offset + 30 + filename_size + extra_size
    corrupted[payload_offset : payload_offset + member_info.compress_size] = (
        b"\x00" * member_info.compress_size
    )
    return bytes(corrupted)


def _build(entries: dict[str, bytes], limits: OoxmlLimits | None = None):
    with SourceHandle(io.BytesIO(_package(entries)), filename="book.xlsx") as source:
        if limits is None:
            return build_manifest(source)
        return build_manifest(source, limits)


def test_manifest_preserves_order_state_and_metadata_without_cell_values() -> None:
    manifest = _build(_entries())

    assert manifest.workbook_type == "xlsx"
    assert manifest.date_system == "1904"
    assert [(sheet.name, sheet.state) for sheet in manifest.sheets] == [
        ("Second", "hidden"),
        ("First", "visible"),
    ]
    assert [sheet.relationship_id for sheet in manifest.sheets] == ["rId2", "rId1"]
    assert [sheet.target for sheet in manifest.sheets] == [
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet1.xml",
    ]
    assert manifest.has_shared_strings is True
    assert manifest.shared_strings_uncompressed_size == len(
        b"<sst>secret shared-string value</sst>"
    )
    assert manifest.styles == StyleManifest(
        custom_number_formats=((165, "yyyy-mm-dd"), (166, "0.00")),
        date_style_ids=(1, 2),
        number_format_codes=("General", "yyyy-mm-dd", "mm-dd-yy"),
    )
    assert manifest.external_relationships == ("externalLinks/externalLink1.xml",)
    assert not hasattr(manifest, "dataframe")
    assert not hasattr(manifest, "shared_strings")


def test_manifest_rejects_duplicate_sheet_names_before_name_lookup() -> None:
    workbook = _workbook().replace(b'name="First"', b'name="Second"')

    with pytest.raises(FormatError, match="duplicate sheet name") as raised:
        _build(_entries(workbook=workbook))

    assert raised.value.context["sheet"] == "Second"


def test_manifest_models_are_frozen_after_source_borrow_closes() -> None:
    raw = io.BytesIO(_package(_entries()))
    raw.seek(7)
    with SourceHandle(raw, filename="book.xlsx") as source:
        manifest = build_manifest(source)
        assert raw.tell() == 7

    assert raw.closed is False
    with pytest.raises(FrozenInstanceError):
        manifest.date_system = "1900"  # type: ignore[misc]


def test_manifest_never_opens_worksheets_shared_strings_or_external_payloads(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    content = _package(_entries())
    opened: list[str] = []
    original_open = zipfile.ZipFile.open

    def recording_open(
        package: zipfile.ZipFile,
        name: str | zipfile.ZipInfo,
        mode: str = "r",
        pwd: bytes | None = None,
        *,
        force_zip64: bool = False,
    ):
        member = name.filename if isinstance(name, zipfile.ZipInfo) else name
        opened.append(member)
        forbidden = (
            member.startswith("xl/worksheets/"),
            member == "xl/sharedStrings.xml",
            member == "xl/externalLinks/externalLink1.xml",
        )
        if any(forbidden):
            raise AssertionError(f"manifest opened value-bearing part {member}")
        return original_open(
            package,
            name,
            mode=mode,
            pwd=pwd,
            force_zip64=force_zip64,
        )

    monkeypatch.setattr(zipfile.ZipFile, "open", recording_open)

    with SourceHandle(io.BytesIO(content), filename="book.xlsx") as source:
        manifest = build_manifest(source)

    assert manifest.external_relationships == ("externalLinks/externalLink1.xml",)
    assert set(opened) == {
        "[Content_Types].xml",
        "xl/_rels/workbook.xml.rels",
        "xl/workbook.xml",
        "xl/styles.xml",
    }


@pytest.mark.parametrize("workbook_type", ["xlsx", "xlsm", "xltx", "xltm"])
def test_workbook_type_comes_from_content_types(workbook_type: str) -> None:
    assert _build(_entries(workbook_type=workbook_type)).workbook_type == workbook_type


@pytest.mark.parametrize(
    ("member", "namespace"),
    [
        ("[Content_Types].xml", CONTENT_TYPES_NS),
        ("xl/_rels/workbook.xml.rels", PACKAGE_RELATIONSHIPS_NS),
        ("xl/workbook.xml", SPREADSHEET_NS),
        ("xl/styles.xml", SPREADSHEET_NS),
    ],
)
def test_spoofed_metadata_element_namespaces_are_rejected(
    member: str,
    namespace: str,
) -> None:
    entries = _entries()
    entries[member] = entries[member].replace(
        namespace.encode(),
        b"https://attacker.invalid/ooxml",
    )

    with pytest.raises(FormatError, match="namespace") as raised:
        _build(entries)

    assert raised.value.context["member"] == member


def test_spoofed_worksheet_relationship_type_is_rejected() -> None:
    relationships = _relationships().replace(
        f"{OFFICE_RELATIONSHIPS_NS}/worksheet".encode(),
        b"https://attacker.invalid/types/worksheet",
    )

    with pytest.raises(FormatError, match="relationship type") as raised:
        _build(_entries(relationships=relationships))

    assert raised.value.context["relationship_type"] == ("https://attacker.invalid/types/worksheet")


def test_https_purl_spreadsheet_namespace_is_rejected() -> None:
    entries = _entries()
    entries["xl/workbook.xml"] = entries["xl/workbook.xml"].replace(
        SPREADSHEET_NS.encode(),
        b"https://purl.oclc.org/ooxml/spreadsheetml/main",
    )

    with pytest.raises(FormatError, match="namespace") as raised:
        _build(entries)

    assert raised.value.context["member"] == "xl/workbook.xml"


def test_https_purl_critical_relationship_type_is_rejected() -> None:
    relationships = _relationships().replace(
        f"{OFFICE_RELATIONSHIPS_NS}/worksheet".encode(),
        b"https://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
    )

    with pytest.raises(FormatError, match="relationship type") as raised:
        _build(_entries(relationships=relationships))

    assert raised.value.context["relationship_type"] == (
        "https://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
    )


def test_strict_ooxml_namespaces_and_relationship_types_are_supported() -> None:
    entries = _entries()
    replacements = {
        "[Content_Types].xml": [(CONTENT_TYPES_NS, STRICT_CONTENT_TYPES_NS)],
        "xl/_rels/workbook.xml.rels": [
            (PACKAGE_RELATIONSHIPS_NS, STRICT_PACKAGE_RELATIONSHIPS_NS),
            (OFFICE_RELATIONSHIPS_NS, STRICT_OFFICE_RELATIONSHIPS_NS),
        ],
        "xl/workbook.xml": [
            (SPREADSHEET_NS, STRICT_SPREADSHEET_NS),
            (OFFICE_RELATIONSHIPS_NS, STRICT_OFFICE_RELATIONSHIPS_NS),
        ],
        "xl/styles.xml": [(SPREADSHEET_NS, STRICT_SPREADSHEET_NS)],
    }
    for member, member_replacements in replacements.items():
        for old, new in member_replacements:
            entries[member] = entries[member].replace(old.encode(), new.encode())

    manifest = _build(entries)

    assert [sheet.name for sheet in manifest.sheets] == ["Second", "First"]
    assert manifest.styles.date_style_ids == (1, 2)


def test_percent_encoded_legitimate_part_name_resolves_to_archive_member() -> None:
    entries = _entries(relationships=_relationships(first_target="worksheets/sheet%201.xml"))
    entries["[Content_Types].xml"] = entries["[Content_Types].xml"].replace(
        b"/xl/worksheets/sheet1.xml",
        b"/xl/worksheets/sheet%201.xml",
    )
    entries["xl/worksheets/sheet%201.xml"] = entries.pop("xl/worksheets/sheet1.xml")

    manifest = _build(entries)

    assert manifest.sheets[1].target == "xl/worksheets/sheet%201.xml"


@pytest.mark.parametrize(("date1904", "expected"), [(None, "1900"), ("0", "1900"), ("1", "1904")])
def test_date_system_comes_from_workbook_properties(
    date1904: str | None,
    expected: str,
) -> None:
    manifest = _build(_entries(workbook=_workbook(date1904=date1904)))

    assert manifest.date_system == expected


def test_optional_styles_and_shared_strings_have_empty_metadata() -> None:
    entries = _entries()
    del entries["xl/styles.xml"]
    del entries["xl/sharedStrings.xml"]
    entries["xl/_rels/workbook.xml.rels"] = b"""<Relationships
      xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet2.xml"/>
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.xml"/>
    </Relationships>"""

    manifest = _build(entries)

    assert manifest.has_shared_strings is False
    assert manifest.shared_strings_uncompressed_size == 0
    assert manifest.styles == StyleManifest((), ())


@pytest.mark.parametrize(
    ("member", "replacement"),
    [
        (
            "[Content_Types].xml",
            f'<Types xmlns="{CONTENT_TYPES_NS}">'.encode(),
        ),
        ("xl/workbook.xml", f'<workbook xmlns="{SPREADSHEET_NS}">'.encode()),
        (
            "xl/_rels/workbook.xml.rels",
            f'<Relationships xmlns="{PACKAGE_RELATIONSHIPS_NS}">'.encode(),
        ),
        ("xl/styles.xml", f'<styleSheet xmlns="{SPREADSHEET_NS}">'.encode()),
    ],
)
def test_malformed_metadata_xml_becomes_contextual_format_error(
    member: str,
    replacement: bytes,
) -> None:
    entries = _entries()
    entries[member] = replacement

    with pytest.raises(FormatError, match="malformed XML") as raised:
        _build(entries)

    assert raised.value.context["member"] == member
    assert raised.value.__cause__ is not None


@pytest.mark.parametrize(
    "member",
    ["[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"],
)
def test_missing_required_manifest_members_are_format_errors(member: str) -> None:
    entries = _entries()
    del entries[member]

    with pytest.raises(FormatError, match="missing required member") as raised:
        _build(entries)

    assert raised.value.context["member"] == member


def test_duplicate_relationship_ids_are_rejected() -> None:
    relationships = b"""<Relationships
      xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet2.xml"/>
    </Relationships>"""

    with pytest.raises(FormatError, match="duplicate relationship") as raised:
        _build(_entries(relationships=relationships))

    assert raised.value.context["relationship_id"] == "rId1"


def test_missing_sheet_relationship_is_rejected() -> None:
    relationships = _relationships().replace(b'Id="rId1"', b'Id="rIdMissing"')

    with pytest.raises(FormatError, match="missing sheet relationship") as raised:
        _build(_entries(relationships=relationships))

    assert raised.value.context["relationship_id"] == "rId1"
    assert raised.value.context["sheet"] == "First"


@pytest.mark.parametrize(
    "target",
    [
        "../../escape.xml",
        "//server/share/sheet1.xml",
        "/../escape.xml",
        "C:/escape.xml",
        "/C:/escape.xml",
        "https://example.invalid/sheet1.xml",
        "worksheets\\sheet1.xml",
        "worksheets/%2e%2e/%2e%2e/escape.xml",
    ],
)
def test_unsafe_internal_relationship_targets_are_rejected(target: str) -> None:
    relationships = _relationships(first_target=target, second_target="worksheets/sheet2.xml")

    with pytest.raises(FormatError, match="unsafe relationship target") as raised:
        _build(_entries(relationships=relationships))

    assert raised.value.context["target"] == target


def test_package_absolute_internal_relationship_target_is_resolved() -> None:
    relationships = _relationships(
        first_target="/xl/worksheets/sheet1.xml",
        second_target="worksheets/sheet2.xml",
    )

    manifest = _build(_entries(relationships=relationships))

    first = next(sheet for sheet in manifest.sheets if sheet.name == "First")
    assert first.target == "xl/worksheets/sheet1.xml"


def test_relationship_target_must_name_an_archive_member() -> None:
    relationships = _relationships(
        first_target="worksheets/missing.xml",
        second_target="worksheets/sheet2.xml",
    )

    with pytest.raises(FormatError, match="relationship target is missing") as raised:
        _build(_entries(relationships=relationships))

    assert raised.value.context["member"] == "xl/worksheets/missing.xml"


@pytest.mark.parametrize(
    ("styles_target", "forbidden_member"),
    [
        ("worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),
        ("sharedStrings.xml", "xl/sharedStrings.xml"),
    ],
)
def test_styles_relationship_cannot_alias_value_bearing_part(
    monkeypatch: pytest.MonkeyPatch,
    styles_target: str,
    forbidden_member: str,
) -> None:
    relationships = _relationships().replace(
        b'Target="styles.xml"',
        f'Target="{styles_target}"'.encode(),
    )
    content = _package(_entries(relationships=relationships))
    opened: list[str] = []
    original_open = zipfile.ZipFile.open

    def recording_open(
        package: zipfile.ZipFile,
        name: str | zipfile.ZipInfo,
        mode: str = "r",
        pwd: bytes | None = None,
        *,
        force_zip64: bool = False,
    ):
        member = name.filename if isinstance(name, zipfile.ZipInfo) else name
        opened.append(member)
        return original_open(
            package,
            name,
            mode=mode,
            pwd=pwd,
            force_zip64=force_zip64,
        )

    monkeypatch.setattr(zipfile.ZipFile, "open", recording_open)

    with (
        SourceHandle(io.BytesIO(content), filename="book.xlsx") as source,
        pytest.raises(FormatError, match="content type") as raised,
    ):
        build_manifest(source)

    assert raised.value.context["member"] == forbidden_member
    assert forbidden_member not in opened


def test_external_sheet_relationship_is_never_accepted_as_internal() -> None:
    relationships = _relationships().replace(
        b'Target="worksheets/sheet1.xml"/',
        b'Target="https://example.invalid/sheet.xml" TargetMode="External"/',
    )

    with pytest.raises(FormatError, match="external sheet relationship"):
        _build(_entries(relationships=relationships))


def test_caller_limits_are_propagated_to_styles_parser() -> None:
    entries = _entries(styles=_styles(extra=f"<ignored>{'x' * 41}</ignored>"))

    with pytest.raises(FormatError, match="text limit") as raised:
        _build(entries, OoxmlLimits(max_element_text=40))

    assert raised.value.context["member"] == "xl/styles.xml"


def test_manifest_failure_still_closes_archive_borrow_and_restores_cursor() -> None:
    raw = io.BytesIO(_package(_entries(workbook=b"<workbook>")))
    raw.seek(11)
    with SourceHandle(raw, filename="broken.xlsx") as source:
        with pytest.raises(FormatError):
            build_manifest(source)
        assert raw.tell() == 11

    assert raw.closed is False


def test_build_manifest_opens_source_once(monkeypatch: pytest.MonkeyPatch) -> None:
    raw = io.BytesIO(_package(_entries()))
    source = SourceHandle(raw, filename="book.xlsx")
    calls = 0
    original = source.open_binary

    def counted_open_binary():
        nonlocal calls
        calls += 1
        return original()

    monkeypatch.setattr(source, "open_binary", counted_open_binary)

    try:
        build_manifest(source)
    finally:
        source.close()

    assert calls == 1
    assert OoxmlLimits() == manifest_module.DEFAULT_LIMITS


@pytest.mark.parametrize(
    "read_error",
    [
        NotImplementedError("unsupported compression"),
        RuntimeError("encrypted member"),
        zipfile.BadZipFile("bad CRC"),
        zlib.error("corrupt DEFLATE stream"),
        lzma.LZMAError("corrupt LZMA stream"),
        EOFError("truncated compressed stream"),
        OSError("corrupt BZIP2 stream"),
    ],
)
def test_archive_member_read_errors_become_contextual_format_errors(
    monkeypatch: pytest.MonkeyPatch,
    read_error: Exception,
) -> None:
    content = _package(_entries())
    original_open = zipfile.ZipFile.open

    def failing_open(
        package: zipfile.ZipFile,
        name: str | zipfile.ZipInfo,
        mode: str = "r",
        pwd: bytes | None = None,
        *,
        force_zip64: bool = False,
    ):
        member = name.filename if isinstance(name, zipfile.ZipInfo) else name
        if member == "[Content_Types].xml":
            raise read_error
        return original_open(
            package,
            name,
            mode=mode,
            pwd=pwd,
            force_zip64=force_zip64,
        )

    monkeypatch.setattr(zipfile.ZipFile, "open", failing_open)

    with (
        SourceHandle(io.BytesIO(content), filename="broken.xlsx") as source,
        pytest.raises(FormatError, match="archive member") as raised,
    ):
        build_manifest(source)

    assert raised.value.context["member"] == "[Content_Types].xml"
    assert raised.value.__cause__ is read_error


def test_corrupt_deflate_payload_becomes_contextual_format_error() -> None:
    content = _corrupt_compressed_member(
        _package(_entries()),
        "[Content_Types].xml",
    )

    with (
        SourceHandle(io.BytesIO(content), filename="broken.xlsx") as source,
        pytest.raises(FormatError, match="archive member") as raised,
    ):
        build_manifest(source)

    assert raised.value.context["member"] == "[Content_Types].xml"
    assert isinstance(raised.value.__cause__, zlib.error)


@pytest.mark.parametrize(
    ("compression", "error_type"),
    [
        (zipfile.ZIP_BZIP2, OSError),
        (zipfile.ZIP_LZMA, lzma.LZMAError),
    ],
)
def test_corrupt_supported_compression_becomes_contextual_format_error(
    compression: int,
    error_type: type[Exception],
) -> None:
    content = _corrupt_compressed_member(
        _package(_entries(), compression),
        "[Content_Types].xml",
    )

    with (
        SourceHandle(io.BytesIO(content), filename="broken.xlsx") as source,
        pytest.raises(FormatError, match="archive member") as raised,
    ):
        build_manifest(source)

    assert raised.value.context["member"] == "[Content_Types].xml"
    assert isinstance(raised.value.__cause__, error_type)
