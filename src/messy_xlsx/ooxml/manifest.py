"""Build an eager, immutable, metadata-only OOXML workbook manifest."""

from __future__ import annotations

import posixpath
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import IO
from urllib.parse import unquote
from zipfile import BadZipFile, ZipFile

from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format

from messy_xlsx._source import SourceHandle
from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import (
    DEFAULT_LIMITS,
    OoxmlLimits,
    SheetDescriptor,
    StyleManifest,
    WorkbookManifest,
)
from messy_xlsx.ooxml.security import (
    canonical_archive_name,
    safe_iterparse,
    validate_archive,
)

_CONTENT_TYPES_MEMBER = "[Content_Types].xml"
_WORKBOOK_MEMBER = "xl/workbook.xml"
_WORKBOOK_RELATIONSHIPS_MEMBER = "xl/_rels/workbook.xml.rels"
_CONTENT_TYPES_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/package/2006/content-types",
        "http://purl.oclc.org/ooxml/package/content-types",
        "https://purl.oclc.org/ooxml/package/content-types",
    }
)
_PACKAGE_RELATIONSHIPS_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/package/2006/relationships",
        "http://purl.oclc.org/ooxml/package/relationships",
        "https://purl.oclc.org/ooxml/package/relationships",
    }
)
_SPREADSHEET_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://purl.oclc.org/ooxml/spreadsheetml/main",
        "https://purl.oclc.org/ooxml/spreadsheetml/main",
    }
)
_OFFICE_RELATIONSHIPS_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
        "https://purl.oclc.org/ooxml/officeDocument/relationships",
    }
)
_CRITICAL_RELATIONSHIP_TYPES = {
    relationship_type: frozenset(
        f"{namespace}/{relationship_type}" for namespace in _OFFICE_RELATIONSHIPS_NAMESPACES
    )
    for relationship_type in ("worksheet", "styles", "sharedStrings")
}
_WORKBOOK_TYPES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "xlsx",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "xlsm",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "xltx",
    "application/vnd.ms-excel.template.macroEnabled.main+xml": "xltm",
}


@dataclass(frozen=True)
class _Relationship:
    relationship_id: str
    relationship_type: str
    target: str
    resolved_target: str | None
    external: bool


@dataclass(frozen=True)
class _WorkbookSheet:
    name: str
    relationship_id: str
    state: str


@dataclass(frozen=True)
class _WorkbookMetadata:
    date_system: str
    sheets: tuple[_WorkbookSheet, ...]


def _qualified_name(tag: str) -> tuple[str, str]:
    if tag.startswith("{") and "}" in tag:
        namespace, local_name = tag[1:].split("}", 1)
        return namespace, local_name
    return "", tag


def _raise_namespace_error(member: str, tag: str) -> None:
    namespace, local_name = _qualified_name(tag)
    raise FormatError(
        "OOXML package contains malformed XML namespace",
        member=member,
        element=local_name,
        namespace=namespace,
    )


def _root_namespace(
    tag: str,
    expected_local_name: str,
    allowed_namespaces: frozenset[str],
    member: str,
) -> str:
    namespace, local_name = _qualified_name(tag)
    if local_name != expected_local_name or namespace not in allowed_namespaces:
        _raise_namespace_error(member, tag)
    return namespace


def _validated_local_name(
    tag: str,
    namespace: str | None,
    namespaced_elements: frozenset[str],
    member: str,
) -> str:
    element_namespace, local_name = _qualified_name(tag)
    if (
        namespace is not None
        and local_name in namespaced_elements
        and element_namespace != namespace
    ):
        _raise_namespace_error(member, tag)
    return local_name


def _open_archive_member(package: ZipFile, member: str) -> IO[bytes]:
    try:
        return package.open(member)
    except (BadZipFile, NotImplementedError, RuntimeError, OSError) as error:
        raise FormatError(
            "OOXML archive member cannot be read",
            member=member,
        ) from error


def _required_member(package: ZipFile, member: str) -> None:
    try:
        package.getinfo(member)
    except KeyError as error:
        raise FormatError(
            "OOXML package is missing required member",
            member=member,
        ) from error


def _read_workbook_type(package: ZipFile, limits: OoxmlLimits) -> str:
    member = _CONTENT_TYPES_MEMBER
    _required_member(package, member)
    workbook_types: list[str] = []
    namespace: str | None = None
    with _open_archive_member(package, member) as source:
        for event, element in safe_iterparse(source, member, limits):
            element_namespace, local_name = _qualified_name(element.tag)
            if event == "start" and namespace is None:
                namespace = _root_namespace(
                    element.tag,
                    "Types",
                    _CONTENT_TYPES_NAMESPACES,
                    member,
                )
            elif local_name == "Override" and element_namespace != namespace:
                _raise_namespace_error(member, element.tag)
            if (
                event == "end"
                and local_name == "Override"
                and (element.attrib.get("PartName") == "/xl/workbook.xml")
            ):
                content_type = element.attrib.get("ContentType")
                if content_type is None:
                    raise FormatError(
                        "OOXML workbook content type is malformed",
                        member=member,
                    )
                workbook_types.append(content_type)
            if event == "end":
                element.clear()

    if len(workbook_types) != 1:
        raise FormatError(
            "OOXML package must declare one workbook content type",
            member=member,
            declarations=len(workbook_types),
        )
    try:
        return _WORKBOOK_TYPES[workbook_types[0]]
    except KeyError as error:
        raise FormatError(
            "OOXML workbook content type is unsupported",
            member=member,
            content_type=workbook_types[0],
        ) from error


def _decode_relationship_target(target: str) -> str:
    try:
        return unquote(target, encoding="utf-8", errors="strict")
    except UnicodeError as error:
        raise FormatError(
            "OOXML package contains unsafe relationship target",
            target=target,
        ) from error


def _internal_target(target: str, *, source_part: str = _WORKBOOK_MEMBER) -> str:
    decoded = _decode_relationship_target(target)
    target_path = PurePosixPath(decoded)
    drive_like = bool(target_path.parts and ":" in target_path.parts[0])
    unsafe = (
        not decoded
        or "\x00" in decoded
        or "\\" in decoded
        or "?" in decoded
        or "#" in decoded
        or drive_like
        or target_path.is_absolute()
        or ".." in target_path.parts
    )
    if unsafe:
        raise FormatError(
            "OOXML package contains unsafe relationship target",
            target=target,
        )

    source_directory = PurePosixPath(source_part).parent.as_posix()
    resolved = posixpath.normpath(posixpath.join(source_directory, decoded))
    resolved_path = PurePosixPath(resolved)
    if resolved_path.is_absolute() or ".." in resolved_path.parts:
        raise FormatError(
            "OOXML package contains unsafe relationship target",
            target=target,
        )
    return resolved_path.as_posix()


def _relationship_type_name(relationship_type: str) -> str:
    return relationship_type.rstrip("/").rsplit("/", 1)[-1]


def _relationship_has_type(relationship_type: str, expected_type: str) -> bool:
    return relationship_type in _CRITICAL_RELATIONSHIP_TYPES[expected_type]


def _validate_critical_relationship_type(
    relationship_type: str,
    relationship_id: str,
    member: str,
) -> None:
    type_name = _relationship_type_name(relationship_type)
    allowed = _CRITICAL_RELATIONSHIP_TYPES.get(type_name)
    if allowed is not None and relationship_type not in allowed:
        raise FormatError(
            "OOXML package contains unsupported relationship type",
            member=member,
            relationship_id=relationship_id,
            relationship_type=relationship_type,
        )


def _read_relationships(
    package: ZipFile,
    member: str,
    limits: OoxmlLimits,
) -> tuple[dict[str, _Relationship], tuple[str, ...]]:
    _required_member(package, member)
    relationships: dict[str, _Relationship] = {}
    external_targets: list[str] = []
    archive_names = {
        canonical_archive_name(archive_name): archive_name for archive_name in package.namelist()
    }
    namespace: str | None = None
    namespaced_elements = frozenset({"Relationships", "Relationship"})

    with _open_archive_member(package, member) as source:
        for event, element in safe_iterparse(source, member, limits):
            if event == "start" and namespace is None:
                namespace = _root_namespace(
                    element.tag,
                    "Relationships",
                    _PACKAGE_RELATIONSHIPS_NAMESPACES,
                    member,
                )
            local_name = _validated_local_name(
                element.tag,
                namespace,
                namespaced_elements,
                member,
            )
            if event == "end" and local_name == "Relationship":
                relationship_id = element.attrib.get("Id")
                relationship_type = element.attrib.get("Type")
                target = element.attrib.get("Target")
                if not relationship_id or not relationship_type or target is None:
                    raise FormatError(
                        "OOXML package contains malformed relationship",
                        member=member,
                    )
                if relationship_id in relationships:
                    raise FormatError(
                        "OOXML package contains duplicate relationship IDs",
                        member=member,
                        relationship_id=relationship_id,
                    )
                _validate_critical_relationship_type(
                    relationship_type,
                    relationship_id,
                    member,
                )

                external = element.attrib.get("TargetMode", "Internal").lower() == "external"
                resolved_target = None if external else _internal_target(target)
                if resolved_target is not None and resolved_target not in archive_names:
                    raise FormatError(
                        "OOXML relationship target is missing from archive",
                        member=resolved_target,
                        relationship_id=relationship_id,
                    )
                if resolved_target is not None:
                    resolved_target = archive_names[resolved_target]
                if external:
                    external_targets.append(target)

                relationships[relationship_id] = _Relationship(
                    relationship_id=relationship_id,
                    relationship_type=relationship_type,
                    target=target,
                    resolved_target=resolved_target,
                    external=external,
                )
            if event == "end":
                element.clear()
    return relationships, tuple(external_targets)


def _relationship_for_type(
    relationships: dict[str, _Relationship],
    relationship_type: str,
) -> _Relationship | None:
    matches = [
        relationship
        for relationship in relationships.values()
        if _relationship_has_type(relationship.relationship_type, relationship_type)
        and not relationship.external
    ]
    if len(matches) > 1:
        raise FormatError(
            "OOXML package contains duplicate workbook relationships",
            relationship_type=relationship_type,
        )
    return matches[0] if matches else None


def _sheet_relationship_id(attributes: Mapping[str, str], member: str) -> str | None:
    relationship_ids = [
        value
        for namespace in _OFFICE_RELATIONSHIPS_NAMESPACES
        if (value := attributes.get(f"{{{namespace}}}id")) is not None
    ]
    if len(relationship_ids) > 1:
        raise FormatError(
            "OOXML workbook contains ambiguous sheet relationship",
            member=member,
        )
    return relationship_ids[0] if relationship_ids else None


def _read_workbook_xml(
    package: ZipFile,
    member: str,
    limits: OoxmlLimits,
) -> _WorkbookMetadata:
    _required_member(package, member)
    date_system = "1900"
    sheets: list[_WorkbookSheet] = []
    namespace: str | None = None
    namespaced_elements = frozenset({"workbook", "workbookPr", "sheets", "sheet"})

    with _open_archive_member(package, member) as source:
        for event, element in safe_iterparse(source, member, limits):
            if event == "start" and namespace is None:
                namespace = _root_namespace(
                    element.tag,
                    "workbook",
                    _SPREADSHEET_NAMESPACES,
                    member,
                )
            local_name = _validated_local_name(
                element.tag,
                namespace,
                namespaced_elements,
                member,
            )
            if event == "start" and local_name == "workbookPr":
                date1904 = element.attrib.get("date1904", "0").lower()
                if date1904 in {"1", "true", "on"}:
                    date_system = "1904"
                elif date1904 not in {"0", "false", "off"}:
                    raise FormatError(
                        "OOXML workbook has invalid date system",
                        member=member,
                        date1904=date1904,
                    )
            elif event == "end" and local_name == "sheet":
                name = element.attrib.get("name")
                relationship_id = _sheet_relationship_id(element.attrib, member)
                state = element.attrib.get("state", "visible")
                if not name or not relationship_id:
                    raise FormatError(
                        "OOXML workbook contains malformed sheet metadata",
                        member=member,
                    )
                if state not in {"visible", "hidden", "veryHidden"}:
                    raise FormatError(
                        "OOXML workbook contains invalid sheet state",
                        member=member,
                        sheet=name,
                        state=state,
                    )
                sheets.append(_WorkbookSheet(name, relationship_id, state))
            if event == "end":
                element.clear()
    return _WorkbookMetadata(date_system, tuple(sheets))


def _build_sheet_descriptors(
    workbook: _WorkbookMetadata,
    relationships: dict[str, _Relationship],
) -> tuple[SheetDescriptor, ...]:
    descriptors: list[SheetDescriptor] = []
    for sheet in workbook.sheets:
        try:
            relationship = relationships[sheet.relationship_id]
        except KeyError as error:
            raise FormatError(
                "OOXML workbook is missing sheet relationship",
                sheet=sheet.name,
                relationship_id=sheet.relationship_id,
            ) from error
        if relationship.external:
            raise FormatError(
                "OOXML workbook uses an external sheet relationship",
                sheet=sheet.name,
                relationship_id=sheet.relationship_id,
            )
        if not _relationship_has_type(relationship.relationship_type, "worksheet"):
            raise FormatError(
                "OOXML sheet relationship has unexpected type",
                sheet=sheet.name,
                relationship_id=sheet.relationship_id,
                relationship_type=relationship.relationship_type,
            )
        if relationship.resolved_target is None:
            raise FormatError(
                "OOXML sheet relationship has no internal target",
                sheet=sheet.name,
                relationship_id=sheet.relationship_id,
            )
        descriptors.append(
            SheetDescriptor(
                name=sheet.name,
                relationship_id=sheet.relationship_id,
                target=relationship.resolved_target,
                state=sheet.state,
            )
        )
    return tuple(descriptors)


def _custom_number_format(
    attributes: Mapping[str, str],
    member: str,
) -> tuple[int, str]:
    try:
        return int(attributes["numFmtId"]), attributes["formatCode"]
    except (KeyError, ValueError) as error:
        raise FormatError(
            "OOXML styles contain malformed number format",
            member=member,
        ) from error


def _style_number_format_id(attributes: Mapping[str, str], member: str) -> int:
    try:
        return int(attributes.get("numFmtId", "0"))
    except ValueError as error:
        raise FormatError(
            "OOXML styles contain malformed style index",
            member=member,
        ) from error


def _read_styles(
    package: ZipFile,
    member: str | None,
    limits: OoxmlLimits,
) -> StyleManifest:
    if member is None:
        return StyleManifest((), ())

    custom: dict[int, str] = {}
    date_styles: list[int] = []
    in_cell_xfs = False
    style_index = 0
    namespace: str | None = None
    namespaced_elements = frozenset({"styleSheet", "numFmt", "cellXfs", "xf"})
    with _open_archive_member(package, member) as source:
        for event, element in safe_iterparse(source, member, limits):
            if event == "start" and namespace is None:
                namespace = _root_namespace(
                    element.tag,
                    "styleSheet",
                    _SPREADSHEET_NAMESPACES,
                    member,
                )
            local_name = _validated_local_name(
                element.tag,
                namespace,
                namespaced_elements,
                member,
            )
            if event == "start" and local_name == "cellXfs":
                in_cell_xfs = True
            elif event == "end" and local_name == "numFmt":
                number_format_id, format_code = _custom_number_format(element.attrib, member)
                custom[number_format_id] = format_code
            elif event == "end" and local_name == "xf" and in_cell_xfs:
                number_format_id = _style_number_format_id(element.attrib, member)
                format_code = custom.get(
                    number_format_id,
                    BUILTIN_FORMATS.get(number_format_id, ""),
                )
                if format_code and is_date_format(format_code):
                    date_styles.append(style_index)
                style_index += 1
            elif event == "end" and local_name == "cellXfs":
                in_cell_xfs = False
            if event == "end":
                element.clear()
    return StyleManifest(tuple(sorted(custom.items())), tuple(date_styles))


def _manifest_from_package(package: ZipFile, limits: OoxmlLimits) -> WorkbookManifest:
    validate_archive(package, limits)
    workbook_type = _read_workbook_type(package, limits)
    relationships, external_targets = _read_relationships(
        package,
        _WORKBOOK_RELATIONSHIPS_MEMBER,
        limits,
    )
    workbook = _read_workbook_xml(package, _WORKBOOK_MEMBER, limits)
    sheets = _build_sheet_descriptors(workbook, relationships)

    shared_relationship = _relationship_for_type(relationships, "sharedStrings")
    shared_target = None if shared_relationship is None else shared_relationship.resolved_target
    shared_info = None if shared_target is None else package.getinfo(shared_target)

    styles_relationship = _relationship_for_type(relationships, "styles")
    styles_target = None if styles_relationship is None else styles_relationship.resolved_target
    styles = _read_styles(package, styles_target, limits)

    return WorkbookManifest(
        workbook_type=workbook_type,
        date_system=workbook.date_system,
        sheets=sheets,
        has_shared_strings=shared_info is not None,
        shared_strings_uncompressed_size=0 if shared_info is None else shared_info.file_size,
        styles=styles,
        external_relationships=external_targets,
    )


def build_manifest(
    source: SourceHandle,
    limits: OoxmlLimits = DEFAULT_LIMITS,
) -> WorkbookManifest:
    """Build and return a closed-resource, metadata-only workbook manifest."""
    try:
        with source.open_binary() as binary, ZipFile(binary) as package:
            manifest = _manifest_from_package(package, limits)
    except BadZipFile as error:
        raise FormatError(
            "Source is not a valid OOXML archive",
            file_path=source.description,
        ) from error
    return manifest
