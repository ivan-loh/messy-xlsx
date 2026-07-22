"""Build an eager, immutable, metadata-only OOXML workbook manifest."""

from __future__ import annotations

import lzma
import posixpath
import zlib
from array import array
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import IO, cast
from urllib.parse import unquote
from zipfile import BadZipFile, ZipFile

from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from messy_xlsx._source import SourceHandle
from messy_xlsx.cache import PathIdentity
from messy_xlsx.detection.locale_detector import is_european_number_format
from messy_xlsx.exceptions import FormatError
from messy_xlsx.ooxml.models import (
    CELL_EUROPEAN_FORMAT,
    CELL_HAS_FORMULA,
    CELL_HAS_VALUE,
    CELL_KIND_BOOLEAN,
    CELL_KIND_DATE,
    CELL_KIND_NUMBER,
    CELL_KIND_OTHER,
    CELL_KIND_TEXT,
    DEFAULT_LIMITS,
    CellEvidenceIndex,
    Interval,
    IntervalIndex,
    MergeRange,
    OoxmlLimits,
    RowBitSet,
    SheetDescriptor,
    SheetManifest,
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
    }
)
_PACKAGE_RELATIONSHIPS_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/package/2006/relationships",
        "http://purl.oclc.org/ooxml/package/relationships",
    }
)
_SPREADSHEET_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://purl.oclc.org/ooxml/spreadsheetml/main",
    }
)
_OFFICE_RELATIONSHIPS_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
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
_CRITICAL_PART_CONTENT_TYPES = {
    "worksheet": frozenset(
        {"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}
    ),
    "styles": frozenset({"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"}),
    "sharedStrings": frozenset(
        {"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"}
    ),
}
_MAX_EXCEL_ROW = 1_048_576
_MAX_EXCEL_COLUMN = 16_384
_ROW_BITSET_BYTES = (_MAX_EXCEL_ROW + 7) // 8
_HEADER_EVIDENCE_ROWS = 19
_LOCALE_EVIDENCE_ROWS = 51
_LOCALE_EVIDENCE_COLUMNS = 21
_SPARSE_EVIDENCE_ROWS = 1_000


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


@dataclass(frozen=True)
class _ContentTypes:
    workbook_type: str
    overrides: tuple[tuple[str, str], ...]
    defaults: tuple[tuple[str, str], ...]


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
    except (
        BadZipFile,
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


def _required_member(package: ZipFile, member: str) -> None:
    try:
        package.getinfo(member)
    except KeyError as error:
        raise FormatError(
            "OOXML package is missing required member",
            member=member,
        ) from error


def _content_type_part_name(part_name: str, member: str) -> str:
    if not part_name.startswith("/") or part_name.startswith("//"):
        raise FormatError(
            "OOXML content type declaration has unsafe part name",
            member=member,
            part_name=part_name,
        )
    try:
        return cast(str, canonical_archive_name(part_name[1:]))
    except ValueError as error:
        raise FormatError(
            "OOXML content type declaration has unsafe part name",
            member=member,
            part_name=part_name,
        ) from error


def _add_default_content_type(
    attributes: Mapping[str, str],
    defaults: dict[str, str],
    member: str,
) -> None:
    extension = attributes.get("Extension")
    content_type = attributes.get("ContentType")
    if not extension or not content_type:
        raise FormatError(
            "OOXML default content type is malformed",
            member=member,
        )
    extension_key = extension.casefold()
    if extension_key in defaults:
        raise FormatError(
            "OOXML package contains duplicate content type declarations",
            member=member,
            extension=extension,
        )
    defaults[extension_key] = content_type


def _add_override_content_type(
    attributes: Mapping[str, str],
    overrides: dict[str, str],
    member: str,
) -> None:
    part_name = attributes.get("PartName")
    content_type = attributes.get("ContentType")
    if not part_name or not content_type:
        raise FormatError(
            "OOXML override content type is malformed",
            member=member,
        )
    part_key = _content_type_part_name(part_name, member)
    if part_key in overrides:
        raise FormatError(
            "OOXML package contains duplicate content type declarations",
            member=member,
            part_name=part_name,
        )
    overrides[part_key] = content_type


def _read_content_types(package: ZipFile, limits: OoxmlLimits) -> _ContentTypes:
    member = _CONTENT_TYPES_MEMBER
    _required_member(package, member)
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    namespace: str | None = None
    namespaced_elements = frozenset({"Types", "Default", "Override"})
    with _open_archive_member(package, member) as source:
        for event, element in safe_iterparse(source, member, limits):
            if event == "start" and namespace is None:
                namespace = _root_namespace(
                    element.tag,
                    "Types",
                    _CONTENT_TYPES_NAMESPACES,
                    member,
                )
            local_name = _validated_local_name(
                element.tag,
                namespace,
                namespaced_elements,
                member,
            )
            if event == "end" and local_name == "Default":
                _add_default_content_type(element.attrib, defaults, member)
            elif event == "end" and local_name == "Override":
                _add_override_content_type(element.attrib, overrides, member)
            if event == "end":
                element.clear()

    workbook_content_type = overrides.get(_WORKBOOK_MEMBER)
    if workbook_content_type is None:
        raise FormatError(
            "OOXML package must declare one workbook content type",
            member=member,
            declarations=0,
        )
    try:
        workbook_type = _WORKBOOK_TYPES[workbook_content_type]
    except KeyError as error:
        raise FormatError(
            "OOXML workbook content type is unsupported",
            member=member,
            content_type=workbook_content_type,
        ) from error
    return _ContentTypes(
        workbook_type=workbook_type,
        overrides=tuple(overrides.items()),
        defaults=tuple(defaults.items()),
    )


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
    package_absolute = decoded.startswith("/") and not decoded.startswith("//")
    first_part = decoded.lstrip("/").split("/", 1)[0]
    drive_like = ":" in first_part
    unsafe = (
        not decoded
        or "\x00" in decoded
        or "\\" in decoded
        or "?" in decoded
        or "#" in decoded
        or drive_like
        or decoded.startswith("//")
        or ".." in target_path.parts
    )
    if unsafe:
        raise FormatError(
            "OOXML package contains unsafe relationship target",
            target=target,
        )

    if package_absolute:
        resolved = posixpath.normpath(decoded[1:])
    else:
        source_directory = PurePosixPath(source_part).parent.as_posix()
        resolved = posixpath.normpath(posixpath.join(source_directory, decoded))
    resolved_path = PurePosixPath(resolved)
    if (
        not resolved
        or resolved == "."
        or resolved_path.is_absolute()
        or ".." in resolved_path.parts
    ):
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


def _content_type_for_member(content_types: _ContentTypes, member: str) -> str | None:
    member_key = canonical_archive_name(member)
    overrides = dict(content_types.overrides)
    if member_key in overrides:
        return overrides[member_key]
    _, separator, extension = member_key.rpartition(".")
    if not separator:
        return None
    return dict(content_types.defaults).get(extension.casefold())


def _validate_critical_relationship_targets(
    relationships: dict[str, _Relationship],
    content_types: _ContentTypes,
) -> None:
    target_kinds: dict[str, str] = {}
    for relationship in relationships.values():
        critical_kinds = [
            kind
            for kind in _CRITICAL_PART_CONTENT_TYPES
            if _relationship_has_type(relationship.relationship_type, kind)
        ]
        if not critical_kinds or relationship.external:
            continue
        target = relationship.resolved_target
        if target is None:
            raise FormatError(
                "OOXML critical relationship has no internal target",
                relationship_id=relationship.relationship_id,
                relationship_type=relationship.relationship_type,
            )
        kind = critical_kinds[0]
        content_type = _content_type_for_member(content_types, target)
        if content_type not in _CRITICAL_PART_CONTENT_TYPES[kind]:
            raise FormatError(
                "OOXML critical relationship target has unexpected content type",
                member=target,
                relationship_id=relationship.relationship_id,
                relationship_type=relationship.relationship_type,
                content_type=content_type,
                expected_part_kind=kind,
            )
        previous_kind = target_kinds.get(target)
        if previous_kind is not None and previous_kind != kind:
            raise FormatError(
                "OOXML critical relationships assign conflicting content types",
                member=target,
                relationship_id=relationship.relationship_id,
                relationship_type=relationship.relationship_type,
                content_type=content_type,
                expected_part_kind=kind,
                conflicting_part_kind=previous_kind,
            )
        target_kinds[target] = kind


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
    seen_names: set[str] = set()
    for sheet in workbook.sheets:
        if sheet.name in seen_names:
            raise FormatError(
                "OOXML workbook contains duplicate sheet name",
                member=_WORKBOOK_MEMBER,
                sheet=sheet.name,
            )
        seen_names.add(sheet.name)
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
    number_format_codes: list[str] = []
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
                number_format_codes.append(format_code)
                if format_code and is_date_format(format_code):
                    date_styles.append(style_index)
                style_index += 1
            elif event == "end" and local_name == "cellXfs":
                in_cell_xfs = False
            if event == "end":
                element.clear()
    return StyleManifest(
        tuple(sorted(custom.items())),
        tuple(date_styles),
        tuple(number_format_codes),
    )


def _manifest_from_package(package: ZipFile, limits: OoxmlLimits) -> WorkbookManifest:
    validate_archive(package, limits)
    content_types = _read_content_types(package, limits)
    relationships, external_targets = _read_relationships(
        package,
        _WORKBOOK_RELATIONSHIPS_MEMBER,
        limits,
    )
    _validate_critical_relationship_targets(relationships, content_types)
    workbook = _read_workbook_xml(package, _WORKBOOK_MEMBER, limits)
    sheets = _build_sheet_descriptors(workbook, relationships)

    shared_relationship = _relationship_for_type(relationships, "sharedStrings")
    shared_target = None if shared_relationship is None else shared_relationship.resolved_target
    shared_info = None if shared_target is None else package.getinfo(shared_target)

    styles_relationship = _relationship_for_type(relationships, "styles")
    styles_target = None if styles_relationship is None else styles_relationship.resolved_target
    styles = _read_styles(package, styles_target, limits)

    return WorkbookManifest(
        workbook_type=content_types.workbook_type,
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


def _one_based_int(
    value: str | None,
    member: str,
    attribute: str,
    maximum: int,
) -> int:
    try:
        parsed = int(value or "")
    except ValueError as error:
        raise FormatError(
            "OOXML worksheet contains malformed coordinate metadata",
            member=member,
            attribute=attribute,
            value=value,
        ) from error
    if parsed < 1 or parsed > maximum:
        raise FormatError(
            "OOXML worksheet contains out-of-bounds coordinate metadata",
            member=member,
            attribute=attribute,
            value=value,
        )
    return parsed


def _style_index(value: str | None, member: str) -> int:
    try:
        parsed = int(value or "0")
    except ValueError as error:
        raise FormatError(
            "OOXML worksheet contains malformed style metadata",
            member=member,
            style=value,
        ) from error
    if parsed < 0:
        raise FormatError(
            "OOXML worksheet contains malformed style metadata",
            member=member,
            style=value,
        )
    return parsed


def _coordinate(value: str | None, member: str) -> tuple[int, int]:
    try:
        row, column = coordinate_to_tuple(value or "")
    except (TypeError, ValueError) as error:
        raise FormatError(
            "OOXML worksheet contains malformed cell coordinate",
            member=member,
            coordinate=value,
        ) from error
    if row < 1 or column < 1 or row > _MAX_EXCEL_ROW or column > _MAX_EXCEL_COLUMN:
        raise FormatError(
            "OOXML worksheet contains out-of-bounds cell coordinate",
            member=member,
            coordinate=value,
        )
    return row, column


def _range(value: str | None, member: str) -> tuple[int, int, int, int]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(value or "")
    except (TypeError, ValueError) as error:
        raise FormatError(
            "OOXML worksheet contains malformed range coordinate",
            member=member,
            coordinate=value,
        ) from error
    if None in {min_col, min_row, max_col, max_row}:
        raise FormatError(
            "OOXML worksheet contains out-of-bounds range coordinate",
            member=member,
            coordinate=value,
        )
    parsed = int(min_row), int(min_col), int(max_row), int(max_col)
    if (
        min(parsed) < 1
        or parsed[0] > _MAX_EXCEL_ROW
        or parsed[2] > _MAX_EXCEL_ROW
        or parsed[1] > _MAX_EXCEL_COLUMN
        or parsed[3] > _MAX_EXCEL_COLUMN
    ):
        raise FormatError(
            "OOXML worksheet contains out-of-bounds range coordinate",
            member=member,
            coordinate=value,
        )
    if parsed[2] < parsed[0] or parsed[3] < parsed[1]:
        raise FormatError(
            "OOXML worksheet contains malformed range coordinate",
            member=member,
            coordinate=value,
        )
    return parsed


def _xml_boolean(value: str | None) -> bool:
    return value is not None and value.casefold() in {"1", "true"}


@dataclass
class _SemanticRegionState:
    """Constant-space streaming state for the legacy data-region scan."""

    first_row: int | None = None
    last_row: int | None = None
    start_col: int | None = None
    end_col: int | None = None
    scan_end: int | None = None
    stopped: bool = False

    def observe(self, row: int, column: int) -> bool:
        """Record one value-bearing coordinate and return whether it is in-region."""
        if self.first_row is None:
            self.first_row = row
            self.scan_end = (1 if row <= 10_000 else row) + 9_999
        assert self.scan_end is not None
        if self.stopped or row > self.scan_end:
            return False
        if self.last_row is not None and row - self.last_row >= 102:
            self.stopped = True
            return False
        self.last_row = row
        self.start_col = column if self.start_col is None else min(self.start_col, column)
        self.end_col = column if self.end_col is None else max(self.end_col, column)
        return True

    def finish(self, observed_max_row: int) -> tuple[int, int, int, int]:
        """Return the final region after applying the legacy max-row hint rule."""
        if (
            self.first_row is None
            or self.last_row is None
            or self.start_col is None
            or self.end_col is None
            or self.scan_end is None
        ):
            return 1, 1, 1, 1
        scan_limit = min(observed_max_row, self.scan_end)
        end_row = self.last_row
        if scan_limit < observed_max_row and scan_limit - end_row < 10:
            end_row = observed_max_row
        return self.first_row, end_row, self.start_col, self.end_col


def _set_row_bit(bits: bytearray, row: int, member: str) -> None:
    if row < 1 or row > _MAX_EXCEL_ROW or row > len(bits) * 8:
        raise FormatError(
            "OOXML worksheet contains out-of-bounds row bit coordinate",
            member=member,
            coordinate=row,
        )
    bit = row - 1
    bits[bit >> 3] |= 1 << (bit & 7)


def _set_packed_code(
    codes: bytearray,
    offset: int,
    code: int,
    member: str,
    coordinate: str,
) -> None:
    if offset < 0 or offset >= len(codes):
        raise FormatError(
            "OOXML worksheet contains out-of-bounds packed coordinate",
            member=member,
            coordinate=coordinate,
        )
    codes[offset] = code


def _increment_sparse_counts(
    counts: array,
    columns: set[int],
    member: str,
) -> None:
    ordered_columns = sorted(columns)
    for column in ordered_columns:
        if column < 1 or column > len(counts) or counts[column - 1] >= 65_535:
            raise FormatError(
                "OOXML worksheet sparse count exceeds bounded storage",
                member=member,
                column=column,
            )
    for column in ordered_columns:
        counts[column - 1] += 1


def _packed_cell_code(
    data_type: str,
    has_value: bool,
    has_formula: bool,
    number_format: str,
) -> int:
    if data_type in {"inlineStr", "s", "str", "e"}:
        kind = CELL_KIND_TEXT
    elif data_type == "b":
        kind = CELL_KIND_BOOLEAN
    elif data_type == "d" or is_date_format(number_format):
        kind = CELL_KIND_DATE
    elif data_type in {"n", ""}:
        kind = CELL_KIND_NUMBER
    else:
        kind = CELL_KIND_OTHER
    return (
        kind
        | (CELL_HAS_VALUE if has_value else 0)
        | (CELL_HAS_FORMULA if has_formula else 0)
        | (
            CELL_EUROPEAN_FORMAT
            if number_format != "General" and is_european_number_format(number_format)
            else 0
        )
    )


def _cell_evidence_index(
    candidates: bytearray | None,
    region: tuple[int, int, int, int],
) -> tuple[CellEvidenceIndex, bool]:
    if candidates is None:
        return CellEvidenceIndex(), False
    start_row, end_row, start_col, end_col = region
    total_rows = end_row - start_row + 1
    width = end_col - start_col + 1
    header_rows = min(_HEADER_EVIDENCE_ROWS, total_rows)
    locale_rows = min(_LOCALE_EVIDENCE_ROWS, total_rows)
    locale_width = min(_LOCALE_EVIDENCE_COLUMNS, width)
    header = bytearray(header_rows * width)
    locale_extra_rows = max(0, locale_rows - header_rows)
    locale = bytearray(locale_extra_rows * locale_width)
    locale_has_european_format = False

    for row_offset in range(locale_rows):
        source_offset = row_offset * _MAX_EXCEL_COLUMN + start_col - 1
        locale_slice = candidates[source_offset : source_offset + locale_width]
        if any(code & CELL_EUROPEAN_FORMAT for code in locale_slice):
            locale_has_european_format = True
        if row_offset < header_rows:
            target_offset = row_offset * width
            header[target_offset : target_offset + width] = candidates[
                source_offset : source_offset + width
            ]
        else:
            target_offset = (row_offset - header_rows) * locale_width
            locale[target_offset : target_offset + locale_width] = locale_slice

    return (
        CellEvidenceIndex(
            start_row=start_row,
            start_col=start_col,
            end_col=end_col,
            header_row_count=header_rows,
            header_codes=bytes(header),
            locale_row_offset=header_rows,
            locale_row_count=locale_extra_rows,
            locale_column_count=locale_width,
            locale_codes=bytes(locale),
        ),
        locale_has_european_format,
    )


class ManifestReader:
    """Build workbook metadata eagerly and sheet metadata on first request.

    Path-backed readers reject identity changes across their lifetime. Caller-owned
    seekable streams remain live views and must not be mutated while this reader is
    in use; non-seekable streams are snapshotted by :class:`SourceHandle`.
    """

    def __init__(
        self,
        source: SourceHandle,
        limits: OoxmlLimits = DEFAULT_LIMITS,
        on_member_open: Callable[[str], None] | None = None,
    ) -> None:
        self._source = source
        self._limits = limits
        self._on_member_open = on_member_open
        self.workbook = build_manifest(source, limits)
        self._path_identity = PathIdentity.before(source.path) if source.path is not None else None
        self._sheets: dict[str, SheetManifest] = {}

    def sheet(self, name: str) -> SheetManifest:
        """Return one cached sheet index, parsing its XML at most once."""
        self._assert_source_unchanged()
        cached = self._sheets.get(name)
        if cached is not None:
            return cached
        try:
            descriptor = next(sheet for sheet in self.workbook.sheets if sheet.name == name)
        except StopIteration as error:
            raise KeyError(name) from error
        parsed = self._parse_sheet(descriptor)
        self._assert_source_unchanged()
        self._sheets[name] = parsed
        return parsed

    def _assert_source_unchanged(self) -> None:
        identity = self._path_identity
        path = self._source.path
        if identity is not None and path is not None and not identity.unchanged(path):
            raise FormatError(
                "OOXML source changed during lazy metadata access",
                file_path=self._source.description,
                operation="read worksheet metadata",
            )

    def _parse_sheet(self, descriptor: SheetDescriptor) -> SheetManifest:
        member = descriptor.target
        try:
            with self._source.open_binary() as binary, ZipFile(binary) as package:
                validate_archive(package, self._limits)
                _required_member(package, member)
                if self._on_member_open is not None:
                    self._on_member_open(member)
                with _open_archive_member(package, member) as xml_source:
                    return self._parse_sheet_xml(descriptor, xml_source)
        except BadZipFile as error:
            raise FormatError(
                "Source is not a valid OOXML archive",
                file_path=self._source.description,
            ) from error

    def _parse_sheet_xml(  # noqa: C901 - one bounded dispatch over worksheet XML tags
        self,
        descriptor: SheetDescriptor,
        source: IO[bytes],
    ) -> SheetManifest:
        member = descriptor.target
        declared_dimension: tuple[int, int, int, int] | None = None
        observed_max_row = 0
        observed_max_col = 0
        observed_min_col = 0
        hidden_rows: list[Interval] = []
        hidden_columns: list[Interval] = []
        merged_ranges: list[MergeRange] = []
        formula_samples: list[str] = []
        region_state = _SemanticRegionState()
        semantic_row_bits = bytearray(_ROW_BITSET_BYTES)
        scoring_candidates: bytearray | None = None
        pre_semantic_row = 0
        pre_semantic_codes = bytearray(_MAX_EXCEL_COLUMN)
        sparse_counts = array("H", [0]) * _MAX_EXCEL_COLUMN
        sparse_row = 0
        sparse_row_columns: set[int] = set()
        has_formulas = False
        current_cell: str | None = None
        current_row = 0
        current_column = 0
        current_data_type = "n"
        current_number_format = "General"
        current_has_value = False
        current_has_formula = False
        enclosing_row = 0
        last_worksheet_row = 0
        last_cell_column = 0
        namespace: str | None = None
        namespaced_elements = frozenset(
            {"worksheet", "dimension", "row", "col", "c", "f", "v", "t", "mergeCell"}
        )

        for event, element in safe_iterparse(source, member, self._limits):
            if event == "start" and namespace is None:
                namespace = _root_namespace(
                    element.tag,
                    "worksheet",
                    _SPREADSHEET_NAMESPACES,
                    member,
                )
            local_name = _validated_local_name(
                element.tag,
                namespace,
                namespaced_elements,
                member,
            )
            if event == "start" and local_name == "row":
                row = _one_based_int(element.attrib.get("r"), member, "r", _MAX_EXCEL_ROW)
                if row <= last_worksheet_row:
                    raise FormatError(
                        "OOXML worksheet requires strictly increasing worksheet row coordinates",
                        member=member,
                        coordinate=row,
                        previous_coordinate=last_worksheet_row,
                    )
                enclosing_row = row
                last_worksheet_row = row
                last_cell_column = 0
            elif event == "start" and local_name == "c":
                current_cell = element.attrib.get("r")
                current_row, current_column = _coordinate(current_cell, member)
                if enclosing_row == 0:
                    raise FormatError(
                        "OOXML worksheet cell has no enclosing row",
                        member=member,
                        coordinate=current_cell,
                    )
                if current_row != enclosing_row:
                    raise FormatError(
                        "OOXML worksheet cell coordinate disagrees with enclosing row",
                        member=member,
                        coordinate=current_cell,
                        enclosing_row=enclosing_row,
                    )
                if current_column <= last_cell_column:
                    raise FormatError(
                        "OOXML worksheet requires strictly increasing worksheet cell coordinates",
                        member=member,
                        coordinate=current_cell,
                        previous_column=last_cell_column,
                    )
                last_cell_column = current_column
                if region_state.first_row is None and current_row != pre_semantic_row:
                    pre_semantic_codes[:] = b"\x00" * _MAX_EXCEL_COLUMN
                    pre_semantic_row = current_row
                if sparse_row and current_row != sparse_row:
                    _increment_sparse_counts(sparse_counts, sparse_row_columns, member)
                    sparse_row_columns.clear()
                sparse_row = current_row
                observed_max_row = max(observed_max_row, current_row)
                observed_max_col = max(observed_max_col, current_column)
                observed_min_col = (
                    current_column
                    if observed_min_col == 0
                    else min(observed_min_col, current_column)
                )
                current_data_type = element.attrib.get("t", "n")
                current_has_value = False
                current_has_formula = False
                style_index = _style_index(element.attrib.get("s"), member)
                style_codes = self.workbook.styles.number_format_codes
                if style_index < len(style_codes):
                    current_number_format = style_codes[style_index]
                elif style_index:
                    raise FormatError(
                        "OOXML worksheet references an unknown style",
                        member=member,
                        style=style_index,
                    )
                else:
                    current_number_format = "General"
            elif event == "end" and local_name == "dimension":
                declared_dimension = _range(element.attrib.get("ref"), member)
            elif event == "end" and local_name == "row":
                if _xml_boolean(element.attrib.get("hidden")):
                    row = _one_based_int(element.attrib.get("r"), member, "r", _MAX_EXCEL_ROW)
                    hidden_rows.append(Interval(row, row))
                enclosing_row = 0
                last_cell_column = 0
            elif (
                event == "end"
                and local_name == "col"
                and _xml_boolean(element.attrib.get("hidden"))
            ):
                hidden_columns.append(
                    Interval(
                        _one_based_int(element.attrib.get("min"), member, "min", _MAX_EXCEL_COLUMN),
                        _one_based_int(element.attrib.get("max"), member, "max", _MAX_EXCEL_COLUMN),
                    )
                )
            elif event == "end" and local_name == "mergeCell":
                min_row, min_col, max_row, max_col = _range(element.attrib.get("ref"), member)
                merged_ranges.append(MergeRange(min_row, min_col, max_row, max_col))
            elif event == "end" and local_name in {"v", "t"} and current_cell is not None:
                if element.text not in {None, ""}:
                    current_has_value = True
            elif event == "end" and local_name == "f":
                has_formulas = True
                current_has_formula = True
                if (
                    current_cell is not None
                    and len(formula_samples) < self._limits.max_formula_samples
                ):
                    formula_samples.append(current_cell)
            elif event == "end" and local_name == "c":
                packed_code = _packed_cell_code(
                    current_data_type,
                    current_has_value,
                    current_has_formula,
                    current_number_format,
                )
                in_region = False
                if current_has_value:
                    _set_row_bit(semantic_row_bits, current_row, member)
                    in_region = region_state.observe(current_row, current_column)
                if (
                    in_region
                    and region_state.first_row is not None
                    and current_row < region_state.first_row + _SPARSE_EVIDENCE_ROWS
                ):
                    sparse_row_columns.add(current_column)
                if region_state.first_row is not None and scoring_candidates is None:
                    scoring_candidates = bytearray(_LOCALE_EVIDENCE_ROWS * _MAX_EXCEL_COLUMN)
                    if pre_semantic_row == region_state.first_row:
                        scoring_candidates[:_MAX_EXCEL_COLUMN] = pre_semantic_codes
                if region_state.first_row is None:
                    _set_packed_code(
                        pre_semantic_codes,
                        current_column - 1,
                        packed_code,
                        member,
                        current_cell or "",
                    )
                if (
                    scoring_candidates is not None
                    and region_state.first_row is not None
                    and region_state.first_row
                    <= current_row
                    < region_state.first_row + _LOCALE_EVIDENCE_ROWS
                ):
                    offset = (
                        (current_row - region_state.first_row) * _MAX_EXCEL_COLUMN
                        + current_column
                        - 1
                    )
                    _set_packed_code(
                        scoring_candidates,
                        offset,
                        packed_code,
                        member,
                        current_cell or "",
                    )
                current_cell = None
            if event == "end":
                element.clear()

        _increment_sparse_counts(sparse_counts, sparse_row_columns, member)
        semantic_data_region = region_state.finish(observed_max_row)
        cell_evidence, locale_has_european_format = _cell_evidence_index(
            scoring_candidates,
            semantic_data_region,
        )
        start_row, _end_row, start_col, end_col = semantic_data_region
        legacy_has_formulas = any(
            cell_evidence.code(row, column) & CELL_HAS_FORMULA
            for row in range(start_row, min(_end_row, start_row + 49) + 1)
            for column in range(start_col, min(end_col, start_col + 9) + 1)
        )
        sparse_filled_counts = bytearray(_MAX_EXCEL_COLUMN * 2)
        for offset, count in enumerate(sparse_counts):
            byte_offset = offset * 2
            sparse_filled_counts[byte_offset : byte_offset + 2] = count.to_bytes(2, "little")

        return SheetManifest(
            name=descriptor.name,
            target=member,
            declared_dimension=declared_dimension,
            observed_max_row=observed_max_row,
            observed_max_col=observed_max_col,
            hidden_rows=IntervalIndex(tuple(hidden_rows)),
            hidden_columns=IntervalIndex(tuple(hidden_columns)),
            merged_ranges=tuple(merged_ranges),
            has_formulas=has_formulas,
            formula_samples=tuple(formula_samples),
            observed_min_col=observed_min_col,
            semantic_data_region=semantic_data_region,
            semantic_nonempty_rows=RowBitSet(bytes(semantic_row_bits)),
            cell_evidence=cell_evidence,
            sparse_filled_counts=bytes(sparse_filled_counts),
            locale_has_european_format=locale_has_european_format,
            legacy_has_formulas=legacy_has_formulas,
        )
