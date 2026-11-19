"""Atomic, source-bound assembly of approved Word body blocks.

The module deliberately supports a compact OOXML intersection: direct body paragraphs
and tables, block, run, row and cell content controls including nested w15 repeating
sections, drawing canvases with their VML fallbacks, styles, numbering, bookmarks,
hyperlinks, images, notes and linked header or footer parts. Imported conflict
revisions, control bindings and shape identities retain their content with
collision-free identities; conflict markup requires matching w14 conflict modes.
Field instructions and caches survive; supported source-theme formatting becomes
explicit properties before publication. Package-level web add-in graphs survive
without activation: equal identities and their cached dependencies are reused,
while distinct opaque pane collections retain independent package-root owners.
Divergent overlapping payloads or unestablished overlapping pane merges refuse.
Web-linked content controls remain unsupported: preserving their package graph
does not establish the fidelity of their body-control links.
Package assembly still refuses ordinary tracked changes and active content.
Session node import additionally preserves whole revisions and locked controls,
with independent cleanup and paste policies exposed directly from this module.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import posixpath
import re
import tempfile
import unicodedata
import uuid
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from enum import Enum
from fractions import Fraction
from io import BytesIO
from itertools import pairwise
from pathlib import Path
from typing import TYPE_CHECKING, Any
from urllib.parse import quote, unquote, urlsplit
from xml.etree import ElementTree as ET

from ._comment_threads import (
    collect_existing_durable_ids,
    collect_existing_para_ids,
    generate_collision_safe_durable_id,
    generate_collision_safe_para_id,
    last_comment_paragraph,
)
from ._optional import feature_module
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    PackageLimits,
    Q,
    _escape_text,
    _parse_bounded_xml,
    _QNameElement,
    _QNameTreeBuilder,
    _quote_attribute,
    _XmlSafetyError,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .ooxml_order import _PPR_ORDER, _RPR_ORDER, _sort_children
from .package import PackageError, PackageGraph, PackageSafetyError, validate_relationships_part

if TYPE_CHECKING:
    from .document import Document
    from .load_options import ProcessorContext
    from .model import Node

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
# Transitional and Strict relationship-attribute namespaces: both name .rels entries.
_RELATIONSHIP_PREFIXES = (f"{{{_R}}}", f"{{{_STRICT_R}}}")
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_DOCUMENT = "word/document.xml"
_STYLES = "word/styles.xml"
_STYLES_RELATIONSHIP = f"{_R}/styles"
_STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_WML_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml."
_RELS = "_rels/.rels"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_ALTERNATE_CONTENT = f"{{{_MC}}}AlternateContent"
_PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
_VML = "urn:schemas-microsoft-com:vml"
_VML_OFFICE = "urn:schemas-microsoft-com:office:office"
_VML_SPID = f"{{{_VML_OFFICE}}}spid"
_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
# Word anchor identities per part: anchor ids are unique across VML (w14) and DrawingML
# (wp14) anchors; DrawingML edit ids form a second domain.
_ANCHOR_ID_DOMAINS = {
    f"{{{_W14}}}anchorId": "anchorId",
    f"{{{_WP14}}}anchorId": "anchorId",
    f"{{{_WP14}}}editId": "editId",
}
# Body-control links require more than preservation of the package add-in graph.
_SDT_FOREIGN_PROPERTIES = frozenset(
    f"{{{_W15}}}{name}" for name in ("webExtensionLinked", "webExtensionCreated")
)
_WEB_EXTENSION_NS = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
_WEB_TASK_PANES_NS = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
_WEB_EXTENSION_REL = "http://schemas.microsoft.com/office/2011/relationships/webextension"
_WEB_TASK_PANES_REL = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
_WEB_EXTENSION_TYPE = "application/vnd.ms-office.webextension+xml"
_WEB_TASK_PANES_TYPE = "application/vnd.ms-office.webextensiontaskpanes+xml"
_OPC_RELS_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
_DIAGRAM_REL_IDS = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds"
_DIAGRAM_DRAWING_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
_RELATIONSHIP_ATTRIBUTES = tuple(
    f"{prefix}{name}"
    for prefix in _RELATIONSHIP_PREFIXES
    for name in ("id", "embed", "link", "dm", "lo", "qs", "cs")
)

# Drawing IDs are unsigned integers, distinct from relationship IDs, paragraph
# IDs, SmartArt modelId values and VML shape names. The outer Word drawing and
# the inner DrawingML objects have separate nonvisual identity domains.
_DRAWING_ID_DOMAINS = {
    **{
        f"{{{namespace}}}docPr": "docPr"
        for namespace in (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
        )
    },
    **{
        f"{{{namespace}}}cNvPr": "cNvPr"
        for namespace in (
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "http://purl.oclc.org/ooxml/drawingml/picture",
            "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
            "http://purl.oclc.org/ooxml/drawingml/chartDrawing",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "http://schemas.microsoft.com/office/drawing/2008/diagram",
        )
    },
}
_DRAWING_CONNECTIONS = {
    f"{{{namespace}}}{name}"
    for namespace in (
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
    )
    for name in ("stCxn", "endCxn")
}
_WORD_DRAWINGS = {
    f"{{{namespace}}}drawing"
    for namespace in (_W, "http://purl.oclc.org/ooxml/wordprocessingml/main")
}
_MAX_DRAWING_ID = 0xFFFFFFFF


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _q(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _word_ns(root: ET.Element) -> str | None:
    return (
        root.tag[1:].split("}", 1)[0]
        if root.tag.startswith("{") and root.tag[1:].split("}", 1)[0] in SUPPORTED_WORD_NAMESPACES
        else None
    )


def _xml_hash(element: ET.Element) -> str:
    """SHA-256 of ``element`` (with its tail) written in expanded-name form.

    ``ET.tostring`` picks prefixes from the process-global namespace registry,
    which feature modules extend when first imported, so its bytes depend on
    import order. Discovery and assembly must fingerprint the same XML
    identically in any process, so names stay in Clark notation and no prefix
    is ever chosen.
    """
    chunks: list[str] = []
    write = chunks.append

    def name(value: str | ET.QName) -> str:
        return value.text if isinstance(value, ET.QName) else value

    def serialize(node: ET.Element) -> None:
        tag = node.tag
        if tag is ET.Comment:
            write(f"<!--{node.text or ''}-->")
        elif tag is ET.ProcessingInstruction:
            write(f"<?{node.text or ''}?>")
        else:
            write("<" + name(tag))
            for key, value in node.items():
                write(f" {name(key)}={_quote_attribute(name(value))}")
            if node.text or len(node):
                write(">")
                if node.text:
                    write(_escape_text(node.text))
                for child in node:
                    serialize(child)
                write(f"</{name(tag)}>")
            else:
                write(" />")
        if node.tail:
            write(_escape_text(node.tail))

    serialize(element)
    return hashlib.sha256("".join(chunks).encode("utf-8")).hexdigest()


def assembly_range_fingerprint(fingerprints: Sequence[str]) -> str:
    """Bind one selection to the discovered fingerprints of its targets, in order.

    A one-target selection binds that target's ``structure_fingerprint_sha256``;
    a multi-target selection binds the SHA-256 of the targets' fingerprints joined
    by newlines. Both are computable from ``discover_assembly_targets`` output.
    """
    values = [_sha(value, "assembly_fingerprint") for value in fingerprints]
    if not values:
        raise ValueError("assembly_selection_range_invalid")
    if len(values) == 1:
        return values[0]
    return hashlib.sha256("\n".join(values).encode("ascii")).hexdigest()


def _body_items(root: ET.Element, ns: str) -> list[ET.Element]:
    body = root.find(Q(ns, "body"))
    if body is None:
        return []
    return [node for node in body if node.tag in {Q(ns, "p"), Q(ns, "tbl")}]


def _assembly_items(root: ET.Element, ns: str) -> list[ET.Element]:
    """Return stable ordinary blocks, alternate content, then atomic SDT targets.

    Append new target families without renumbering existing public identifiers.
    """
    body = root.find(Q(ns, "body"))
    if body is None:
        return []
    return [
        *_body_items(root, ns),
        *(node for node in body if node.tag == _MC_ALTERNATE_CONTENT),
        *(node for node in body if node.tag == Q(ns, "sdt")),
    ]


def _rels_name(part: str) -> str:
    parent, name = posixpath.split(part)
    return posixpath.join(parent, "_rels", name + ".rels")


def _resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _relative(part: str, target: str) -> str:
    return posixpath.relpath(target, posixpath.dirname(part))


@dataclass(frozen=True)
class AssemblySelection:
    start: int
    end: int
    structure_fingerprint_sha256: str


@dataclass(frozen=True)
class AssemblySource:
    path: Path
    sha256: str
    selections: tuple[AssemblySelection, ...]


@dataclass(frozen=True)
class SectionPolicy:
    header_footer_inheritance: str
    page_numbering: str


@dataclass(frozen=True)
class FieldPolicy:
    mark_dirty: bool
    request_on_open: bool


class ImportFormatMode(str, Enum):
    """Style collision policy; existing requests retain source formatting."""

    USE_DESTINATION_STYLES = "use_destination_styles"
    KEEP_SOURCE_FORMATTING = "keep_source_formatting"
    KEEP_DIFFERENT_STYLES = "keep_different_styles"


@dataclass(frozen=True)
class ImportFormatOptions:
    """Paste policies shared by package assembly and session node import.

    Theme colors resolve by default, retaining the historical source-formatting
    behavior. Spacing adjusts only inline insertion boundaries, never split runs
    within the source. ``force_copy_styles`` overrides style reuse.
    """

    smart_style_behavior: bool = False
    keep_source_numbering: bool = True
    ignore_header_footer: bool = False
    merge_pasted_lists: bool = False
    adjust_sentence_and_word_spacing: bool = False
    append_with_new_page: bool = False
    force_copy_styles: bool = False
    ignore_text_boxes: bool = False
    resolve_theme_colors: bool = True

    def __post_init__(self) -> None:
        if any(not isinstance(getattr(self, name), bool) for name in self.__dataclass_fields__):
            raise ValueError("assembly_import_format_options_invalid")


_DEFAULT_IMPORT_OPTIONS = ImportFormatOptions()


@dataclass(frozen=True)
class AssemblyRequest:
    destination: Path
    destination_sha256: str
    output: Path
    insert_after: int
    anchor_fingerprint_sha256: str | None
    sources: tuple[AssemblySource, ...]
    section_policy: SectionPolicy
    field_policy: FieldPolicy
    import_format_mode: ImportFormatMode = ImportFormatMode.KEEP_SOURCE_FORMATTING
    import_format_options: ImportFormatOptions = _DEFAULT_IMPORT_OPTIONS


@dataclass(frozen=True)
class AssemblyReport:
    status: str
    code: str | None
    destination_sha256: str
    output_sha256: str | None
    inserted_blocks: int
    sources: list[dict[str, Any]]
    unselected_parts_preserved: bool | None = None
    invalidated_signatures: tuple[str, ...] = ()

    @property
    def passed(self) -> bool:
        return self.status == "published"

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "code": self.code,
            "destination_sha256": self.destination_sha256,
            "output_sha256": self.output_sha256,
            "inserted_blocks": self.inserted_blocks,
            "sources": self.sources,
            "unselected_parts_preserved": self.unselected_parts_preserved,
            "invalidated_signatures": list(self.invalidated_signatures),
        }


def _sha(value: Any, name: str) -> str:
    if (
        not isinstance(value, str)
        or len(value) != 64
        or any(c not in "0123456789abcdefABCDEF" for c in value)
    ):
        raise ValueError(f"{name}_invalid")
    return value.lower()


def _positive(value: Any, name: str, *, zero: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < (0 if zero else 1):
        raise ValueError(f"{name}_invalid")
    return value


def _parse_section_policy(value: Any) -> SectionPolicy:
    if value is None:
        return SectionPolicy("source", "source")
    if not isinstance(value, Mapping) or set(value) != {
        "header_footer_inheritance",
        "page_numbering",
    }:
        raise ValueError("assembly_section_policy_invalid")
    header_footer = value["header_footer_inheritance"]
    page_numbering = value["page_numbering"]
    if header_footer not in {"source", "link_previous"}:
        raise ValueError("assembly_header_footer_policy_invalid")
    if page_numbering not in {"source", "continue"}:
        raise ValueError("assembly_page_numbering_policy_invalid")
    return SectionPolicy(header_footer, page_numbering)


def _parse_field_policy(value: Any) -> FieldPolicy:
    if value is None:
        return FieldPolicy(False, False)
    if (
        not isinstance(value, Mapping)
        or set(value) != {"mark_dirty", "request_on_open"}
        or not isinstance(value["mark_dirty"], bool)
        or not isinstance(value["request_on_open"], bool)
    ):
        raise ValueError("assembly_field_policy_invalid")
    return FieldPolicy(value["mark_dirty"], value["request_on_open"])


def _parse_import_mode(value: Any) -> ImportFormatMode:
    try:
        return ImportFormatMode(value)
    except (TypeError, ValueError) as error:
        raise ValueError("assembly_import_format_mode_invalid") from error


def _parse_import_options(value: Any) -> ImportFormatOptions:
    if value is None:
        return ImportFormatOptions()
    names = set(ImportFormatOptions.__dataclass_fields__)
    if (
        not isinstance(value, Mapping)
        or not set(value) <= names
        or any(not isinstance(option, bool) for option in value.values())
    ):
        raise ValueError("assembly_import_format_options_invalid")
    return ImportFormatOptions(**value)


def parse_assembly_request(value: Any) -> AssemblyRequest:
    if not isinstance(value, Mapping):
        raise TypeError("assembly_request_invalid")
    required = {
        "destination",
        "destination_sha256",
        "output",
        "insert_after",
        "anchor_fingerprint_sha256",
        "sources",
    }
    fields = set(value)
    if not required <= fields or not fields <= required | {
        "section_policy",
        "field_policy",
        "import_format_mode",
        "import_format_options",
    }:
        raise ValueError("assembly_request_fields_invalid")
    raw_sources = value["sources"]
    if not isinstance(raw_sources, list) or not raw_sources:
        raise ValueError("assembly_sources_invalid")
    sources: list[AssemblySource] = []
    for raw in raw_sources:
        if (
            not isinstance(raw, Mapping)
            or set(raw) != {"path", "sha256", "selections"}
            or not isinstance(raw["path"], str)
            or not isinstance(raw["selections"], list)
            or not raw["selections"]
        ):
            raise ValueError("assembly_source_invalid")
        selections: list[AssemblySelection] = []
        for selection in raw["selections"]:
            if not isinstance(selection, Mapping) or set(selection) != {
                "start",
                "end",
                "structure_fingerprint_sha256",
            }:
                raise ValueError("assembly_selection_invalid")
            start, end = (
                _positive(selection["start"], "assembly_start", zero=True),
                _positive(selection["end"], "assembly_end", zero=True),
            )
            if (start == 0) != (end == 0) or (start and end < start):
                raise ValueError("assembly_selection_range_invalid")
            selections.append(
                AssemblySelection(
                    start,
                    end,
                    _sha(
                        selection["structure_fingerprint_sha256"],
                        "assembly_fingerprint",
                    ),
                )
            )
        sources.append(
            AssemblySource(
                Path(raw["path"]),
                _sha(raw["sha256"], "assembly_source_sha256"),
                tuple(selections),
            )
        )
    anchor = value["anchor_fingerprint_sha256"]
    if anchor is not None:
        anchor = _sha(anchor, "assembly_anchor_fingerprint")
    insert_after = _positive(value["insert_after"], "assembly_insert_after", zero=True)
    if insert_after == 0 and anchor is not None:
        raise ValueError("assembly_anchor_without_target")
    if insert_after and anchor is None:
        raise ValueError("assembly_anchor_required")
    if not all(isinstance(value[key], str) for key in ("destination", "output")):
        raise ValueError("assembly_path_invalid")
    return AssemblyRequest(
        Path(value["destination"]),
        _sha(value["destination_sha256"], "assembly_destination_sha256"),
        Path(value["output"]),
        insert_after,
        anchor,
        tuple(sources),
        _parse_section_policy(value.get("section_policy")),
        _parse_field_policy(value.get("field_policy")),
        _parse_import_mode(value.get("import_format_mode", "keep_source_formatting")),
        _parse_import_options(value.get("import_format_options")),
    )


def load_assembly_request(path: str | Path) -> AssemblyRequest:
    with Path(path).open(encoding="utf-8") as stream:
        return parse_assembly_request(json.load(stream))


def _refuse(request: AssemblyRequest, code: str) -> AssemblyReport:
    return AssemblyReport("refused", code, request.destination_sha256, None, 0, [])


def _package_limit_refusal(path: Path) -> str | None:
    """The package-limit code ``Document`` and ``inspect_docx`` would report, else ``None``.

    Assembly reads every member into memory, so a package outside the default
    ``PackageLimits`` is refused before any member is decompressed.
    """
    try:
        PackageGraph.from_file(path, limits=PackageLimits())
    except PackageSafetyError as error:
        return error.code
    except PackageError:
        return None
    return None


def _section_metadata(section: ET.Element, ns: str) -> dict[str, Any]:
    page_size = section.find(Q(ns, "pgSz"))
    page_numbers = section.find(Q(ns, "pgNumType"))
    return {
        "orientation": (page_size.get(Q(ns, "orient")) if page_size is not None else None),
        "page_width": (page_size.get(Q(ns, "w")) if page_size is not None else None),
        "page_height": (page_size.get(Q(ns, "h")) if page_size is not None else None),
        "header_footer_references": [
            {
                "kind": child.tag.rsplit("}", 1)[-1],
                "type": child.get(Q(ns, "type")),
                "relationship_id": child.get(Q(_R, "id")),
            }
            for child in section
            if child.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
        ],
        "page_number_start": (
            page_numbers.get(Q(ns, "start")) if page_numbers is not None else None
        ),
    }


def _section_nodes(root: ET.Element, ns: str) -> list[ET.Element]:
    body = root.find(Q(ns, "body"))
    if body is None:
        return []
    sections = [
        section
        for item in _body_items(root, ns)
        if (properties := item.find(Q(ns, "pPr"))) is not None
        and (section := properties.find(Q(ns, "sectPr"))) is not None
    ]
    final = body.find(Q(ns, "sectPr"))
    return [*sections, *([final] if final is not None else [])]


def discover_assembly_targets(source: str | Path) -> dict[str, Any]:
    """Return direct source-body blocks and exact hash bindings without mutation."""
    path = Path(source)
    refusal = _package_limit_refusal(path)
    if refusal is not None:
        return {"passed": False, "code": refusal, "targets": []}
    try:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read(_DOCUMENT))
    except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError):
        return {"passed": False, "code": "source_package_invalid", "targets": []}
    ns = _word_ns(root)
    if ns is None:
        return {
            "passed": False,
            "code": "source_word_namespace_unsupported",
            "targets": [],
        }
    targets = []
    for index, item in enumerate(_assembly_items(root, ns), start=1):
        section = item.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}") if item.tag == Q(ns, "p") else None
        targets.append(
            {
                "start": index,
                "end": index,
                "kind": "section-boundary"
                if section is not None
                else "alternate-content"
                if item.tag == _MC_ALTERNATE_CONTENT
                else item.tag.rsplit("}", 1)[-1],
                "structure_fingerprint_sha256": _xml_hash(item),
                "text": "".join(node.text or "" for node in item.iter(Q(ns, "t"))),
                **({"section": _section_metadata(section, ns)} if section is not None else {}),
            }
        )
    body = root.find(Q(ns, "body"))
    final_section = body.find(Q(ns, "sectPr")) if body is not None else None
    if final_section is not None:
        targets.append(
            {
                "start": 0,
                "end": 0,
                "kind": "section-schedule",
                "structure_fingerprint_sha256": _xml_hash(final_section),
                "text": "",
                "section": _section_metadata(final_section, ns),
            }
        )
    return {
        "passed": True,
        "code": None,
        "source_sha256": _sha256(path),
        "targets": targets,
    }


def _supported_sdt(control: ET.Element, parents: Mapping[ET.Element, ET.Element], ns: str) -> bool:
    """Admit one unlocked control whose repeating structure is whole.

    Block, run, row and cell controls nest freely. A repeating section holds only
    direct items and an item sits directly in a repeating section, so no partial
    collection is imported.
    """
    templates = feature_module("templates")

    properties = control.findall(Q(ns, "sdtPr"))
    if len(properties) != 1 or len(control.findall(Q(ns, "sdtContent"))) != 1:
        return False
    if any(lock.get(Q(ns, "val")) != "unlocked" for lock in properties[0].findall(Q(ns, "lock"))):
        return False
    if any(child.tag in _SDT_FOREIGN_PROPERTIES for child in properties[0]):
        return False
    if any(node.tag == Q(ns, "sectPr") for node in control.iter()):
        return False
    section = templates._sdt_has_property(control, "repeatingSection", ns)
    item = templates._sdt_has_property(control, "repeatingSectionItem", ns)
    if section and (item or templates._repeating_section_direct_items(control, ns) is None):
        return False
    if not item:
        return True
    content = parents.get(control)
    owner = parents.get(content) if content is not None else None
    return (
        content is not None
        and content.tag == Q(ns, "sdtContent")
        and owner is not None
        and templates._sdt_has_property(owner, "repeatingSection", ns)
    )


def _unsupported(elements: list[ET.Element], ns: str, *, allow_smart_tags: bool = False) -> bool:
    from .review import _CONFLICT_REVISION_TAGS, _part_revision_index

    # Comment markers are imported with their comments (_collect_comment_imports).
    forbidden = {
        "altChunk",
        "customXml",
        "del",
        "ins",
        "moveFrom",
        "moveTo",
        "smartTag",
    }
    if allow_smart_tags:
        forbidden.remove("smartTag")
    for element in elements:
        if any(node.tag.rsplit("}", 1)[-1] in forbidden for node in element.iter()):
            return True
        parents = {child: parent for parent in element.iter() for child in parent}
        if not all(_supported_sdt(control, parents, ns) for control in element.iter(Q(ns, "sdt"))):
            return True
        revisions, _, _ = _part_revision_index(element, ns)
        if any(revision.node.tag not in _CONFLICT_REVISION_TAGS for revision in revisions):
            return True
    return False


def _read_rels(parts: Mapping[str, bytes], owner: str) -> ET.Element:
    name = _rels_name(owner)
    return (
        parse_xml_with_preserved_qnames(parts[name])
        if name in parts
        else ET.Element(_q(_PR, "Relationships"))
    )


def _ids(root: ET.Element) -> set[str]:
    return {node.get("Id") for node in root.findall(_q(_PR, "Relationship")) if node.get("Id")}


def _new_rel_id(root: ET.Element) -> str:
    used = _ids(root)
    serial = 1
    while f"rId{serial}" in used:
        serial += 1
    return f"rId{serial}"


def _content_type(part: str, types: ET.Element) -> tuple[str, str] | None:
    suffix = "rels" if part == _RELS else posixpath.splitext(part)[1].lstrip(".")
    for node in types.findall(_q(_CT, "Override")):
        if node.get("PartName") == "/" + part:
            return "Override", node.get("ContentType", "")
    for node in types.findall(_q(_CT, "Default")):
        if node.get("Extension") == suffix:
            return "Default", node.get("ContentType", "")
    return None


def _transitional_assembly_parts(source: Mapping[str, bytes]) -> dict[str, bytes]:
    """Translate a private Strict view; source bytes and selection hashes stay original.

    Reuse save-time vocabulary conversion, including table border names and
    namespace declarations used by lexical MCE QNames. URI-valued DrawingML
    graphic-data selectors need the same vocabulary; visible text and external
    relationship targets are user data and must remain untouched. Unknown Strict
    namespaces refuse the transaction rather than leaking into the destination.
    """
    from .session_ops.rich_objects import _OOXML_NAMESPACES, _convert_ooxml

    mapping = {strict: transitional for transitional, strict in _OOXML_NAMESPACES.items()}
    types = ET.fromstring(source["[Content_Types].xml"])
    result = dict(source)
    for part, data in source.items():
        content_type = _content_type(part, types)
        if not (
            part.endswith((".xml", ".rels", ".vml"))
            or content_type is not None
            and (
                content_type[1].endswith("+xml")
                or content_type[1] in {"application/xml", "text/xml"}
            )
        ):
            continue
        converted = _convert_ooxml(
            data, strict=False, main=part == _DOCUMENT, limits=PackageLimits()
        )
        root = parse_xml_with_preserved_qnames(converted)
        changed = False
        for node in root.iter():
            # Include namespace bindings used only in mc:* or other QName values.
            namespaces = [uri for _, uri in getattr(node, "_namespace_scope", ())]
            namespaces.extend(
                name[1:].split("}", 1)[0]
                for name in (node.tag, *node.attrib)
                if isinstance(name, str) and name.startswith("{")
            )
            if any(uri.startswith("http://purl.oclc.org/ooxml/") for uri in namespaces):
                raise ValueError("assembly_strict_namespace_unsupported")
            if node.tag == "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData":
                uri = node.get("uri", "")
                if uri in mapping:
                    node.set("uri", mapping[uri])
                    changed = True
                elif uri.startswith("http://purl.oclc.org/ooxml/"):
                    raise ValueError("assembly_strict_namespace_unsupported")
        result[part] = serialize_xml_with_preserved_qnames(root) if changed else converted
    return result


class _Importer:
    def __init__(
        self,
        destination: dict[str, bytes],
        source: Mapping[str, bytes],
        label: str,
        source_path: Path,
    ) -> None:
        self.destination, self.source, self.label = destination, source, label
        self.types = parse_xml_with_preserved_qnames(destination["[Content_Types].xml"])
        self.source_types = ET.fromstring(source["[Content_Types].xml"])
        self.rel_map: dict[tuple[str, str, str], str] = {}
        self.part_map: dict[str, str] = {}
        self.source_path = source_path
        self.smartart_graph: PackageGraph | None = None
        self.diagram_caches: dict[str, str] = {}
        self.binding_stores: dict[str, str] = {}
        self.binding_graph: PackageGraph | None = None

    def _copy_type(self, source_part: str, dest_part: str) -> None:
        found = _content_type(source_part, self.source_types)
        if found is None:
            raise ValueError("assembly_content_type_missing")
        kind, content_type = found
        if kind == "Default":
            suffix = "rels" if dest_part == _RELS else posixpath.splitext(dest_part)[1].lstrip(".")
            destination_default = next(
                (
                    node
                    for node in self.types.findall(_q(_CT, "Default"))
                    if node.get("Extension") == suffix
                ),
                None,
            )
            if destination_default is None:
                self.types.append(
                    ET.Element(
                        _q(_CT, "Default"),
                        {"Extension": suffix, "ContentType": content_type},
                    )
                )
                return
            if destination_default.get("ContentType") == content_type:
                return
        if not any(
            node.get("PartName") == "/" + dest_part and node.get("ContentType") == content_type
            for node in self.types.findall(_q(_CT, "Override"))
        ):
            self.types.append(
                ET.Element(
                    _q(_CT, "Override"),
                    {"PartName": "/" + dest_part, "ContentType": content_type},
                )
            )

    def copy_relation(
        self, source_owner: str, source_id: str, destination_owner: str | None = None
    ) -> str:
        destination_owner = destination_owner or source_owner
        key = (source_owner, source_id, destination_owner)
        if key in self.rel_map:
            return self.rel_map[key]
        source_rels = _read_rels(self.source, source_owner)
        relation = next(
            (
                node
                for node in source_rels.findall(_q(_PR, "Relationship"))
                if node.get("Id") == source_id
            ),
            None,
        )
        if relation is None:
            raise ValueError("assembly_relationship_missing")
        destination_rels = _read_rels(self.destination, destination_owner)
        destination_id = _new_rel_id(destination_rels)
        imported = copy.deepcopy(relation)
        attributes = imported.attrib
        attributes["Id"] = destination_id
        if attributes.get("TargetMode") != "External":
            target = relation.get("Target")
            if target is None:
                raise ValueError("assembly_relationship_missing")
            attributes["Target"] = _relative(
                destination_owner, self.copy_part(_resolve(source_owner, target))
            )
        destination_rels.append(imported)
        self.destination[_rels_name(destination_owner)] = serialize_xml_with_preserved_qnames(
            destination_rels
        )
        self.rel_map[key] = destination_id
        return destination_id

    def _bind_diagram_cache(self, rel_ids: ET.Element, source_owner: str) -> None:
        smartart = feature_module("smartart")

        if self.smartart_graph is None:
            self.smartart_graph = PackageGraph.from_file(self.source_path)
        graph = self.smartart_graph
        diagram_parts, _, cache = smartart._resolve_smartart_parts(graph, source_owner, rel_ids)
        if cache is None:
            raise ValueError("assembly_smartart_unsupported")
        data_part = diagram_parts["dm"]
        if data_part in self.diagram_caches:
            if self.diagram_caches[data_part] != cache:
                raise ValueError("assembly_smartart_relationship_invalid")
            return
        self.diagram_caches[data_part] = cache
        destination_data = self.copy_part(data_part)
        if any(rel.type == _DIAGRAM_DRAWING_REL for rel in graph.relationships_for(data_part)):
            # copy_part retains scoped IDs and the entire related-part closure.
            return
        # Promote only this diagram's validated legacy owner association. Never
        # append implicit caches to the destination story's relationship set.
        relation = next(
            rel for rel in graph.relationships_for(source_owner) if rel.type == _DIAGRAM_DRAWING_REL
        )
        identifier = self.copy_relation(source_owner, relation.id, destination_data)
        data = parse_xml_with_preserved_qnames(self.destination[destination_data])
        diagram_ns = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
        extensions = data.find(_q(diagram_ns, "extLst"))
        if extensions is None:
            extensions = ET.SubElement(data, _q(diagram_ns, "extLst"))
        extension = ET.SubElement(
            extensions,
            _q("http://schemas.openxmlformats.org/drawingml/2006/main", "ext"),
            {"uri": "{C7B3CF3C-8C70-41FD-926A-BAF1BF945CEB}"},
        )
        ET.SubElement(extension, smartart._DATA_MODEL_EXT, {"relId": identifier})
        self.destination[destination_data] = serialize_xml_with_preserved_qnames(data)

    def copy_references(
        self,
        roots: list[ET.Element],
        source_owner: str,
        destination_owner: str | None = None,
    ) -> None:
        for root in roots:
            for node in root.iter():
                if node.tag == _DIAGRAM_REL_IDS:
                    self._bind_diagram_cache(node, source_owner)
                for attr in _RELATIONSHIP_ATTRIBUTES:
                    if (relationship := node.get(attr)) is not None:
                        node.set(
                            attr,
                            self.copy_relation(source_owner, relationship, destination_owner),
                        )

    def copy_bindings(self, roots: list[ET.Element], ns: str) -> None:
        """Copy bound datastores through their owners and assign private store IDs."""
        content_controls = feature_module("content_controls", required=False)
        from .properties import _store_item_key

        for root in roots:
            for binding in (
                node
                for node in root.iter()
                if node.tag in {Q(ns, "dataBinding"), Q(_W15, "dataBinding")}
            ):
                if content_controls is None:
                    content_controls = feature_module("content_controls")
                if self.binding_graph is None:
                    self.binding_graph = PackageGraph.from_file(self.source_path)
                graph = self.binding_graph
                old_id = binding.get(Q(ns, "storeItemID"))
                if (
                    old_id is None
                    or content_controls.resolve_custom_xml_databinding(graph, binding) is None
                ):
                    raise ValueError("assembly_sdt_binding_unsupported")
                key = _store_item_key(old_id)
                if key not in self.binding_stores:
                    store = content_controls._custom_xml_store_parts(graph, old_id)
                    if store is None:
                        raise ValueError("assembly_sdt_binding_unsupported")
                    properties_part, data_part = store
                    owners = [
                        rel
                        for rel in graph.relationships_for(_DOCUMENT)
                        if rel.type.endswith("/customXml")
                        and not rel.is_external
                        and rel.resolved_target() == data_part
                    ]
                    properties_edges = [
                        rel
                        for rel in graph.relationships_for(data_part)
                        if rel.type.endswith("/customXmlProps")
                        and not rel.is_external
                        and rel.resolved_target() == properties_part
                    ]
                    if len(owners) != 1 or len(properties_edges) != 1:
                        raise ValueError("assembly_sdt_binding_unsupported")
                    self.copy_relation(_DOCUMENT, owners[0].id)
                    destination_properties = self.part_map[properties_part]
                    properties = parse_xml_with_preserved_qnames(
                        self.destination[destination_properties]
                    )
                    occupied = {
                        _store_item_key(value)
                        for name, data in self.destination.items()
                        if name.startswith("customXml/itemProps") and name.endswith(".xml")
                        for node in [ET.fromstring(data)]
                        for attr, value in node.attrib.items()
                        if attr.rsplit("}", 1)[-1] == "itemID"
                    }
                    serial = 0
                    while True:
                        identifier = (
                            "{"
                            + str(
                                uuid.uuid5(
                                    uuid.NAMESPACE_OID,
                                    f"{self.label}:{destination_properties}:{key}:{serial}",
                                )
                            ).upper()
                            + "}"
                        )
                        if _store_item_key(identifier) not in occupied:
                            break
                        serial += 1
                    item_id_attributes = [
                        attr for attr in properties.attrib if attr.rsplit("}", 1)[-1] == "itemID"
                    ]
                    if len(item_id_attributes) != 1:
                        raise ValueError("assembly_sdt_binding_unsupported")
                    properties.set(item_id_attributes[0], identifier)
                    self.destination[destination_properties] = serialize_xml_with_preserved_qnames(
                        properties
                    )
                    self.binding_stores[key] = identifier
                binding.set(Q(ns, "storeItemID"), self.binding_stores[key])

    def copy_part(self, source_part: str) -> str:
        if source_part in self.part_map:
            return self.part_map[source_part]
        if source_part not in self.source:
            raise ValueError("assembly_related_part_missing")
        source_rels_name = _rels_name(source_part)
        dest_rels = (
            parse_xml_with_preserved_qnames(self.source[source_rels_name])
            if source_rels_name in self.source
            else None
        )
        self._require_resolved_references(source_part, dest_rels)
        directory, filename = posixpath.split(source_part)
        # Store discovery deliberately accepts conventional itemProps names only.
        prefix = (
            "itemProps-" if directory == "customXml" and filename.startswith("itemProps") else ""
        )
        destination_part = posixpath.join(directory, f"{prefix}assembly-{self.label}-{filename}")
        serial = 2
        while destination_part in self.destination:
            destination_part = posixpath.join(
                directory, f"{prefix}assembly-{self.label}-{serial}-{filename}"
            )
            serial += 1
        self.part_map[source_part] = destination_part
        data = self.source[source_part]
        self.destination[destination_part] = data
        self._copy_type(source_part, destination_part)
        if dest_rels is not None:
            for relation in dest_rels.findall(_q(_PR, "Relationship")):
                attributes = relation.attrib
                if attributes.get("TargetMode") != "External":
                    attributes["Target"] = _relative(
                        destination_part,
                        self.copy_part(_resolve(source_part, attributes["Target"])),
                    )
            self.destination[_rels_name(destination_part)] = serialize_xml_with_preserved_qnames(
                dest_rels
            )
        return destination_part

    def _require_resolved_references(self, source_part: str, relations: ET.Element | None) -> None:
        """Refuse a related XML part (a header, footer, chart, ...) copied with its own
        relationships when one of its r:* references names no relationship of that part."""
        found = _content_type(source_part, self.source_types)
        if found is None or not (
            found[1].endswith("+xml") or found[1] in {"application/xml", "text/xml"}
        ):
            return
        declared = _ids(relations) if relations is not None else set()
        for node in ET.fromstring(self.source[source_part]).iter():
            for name, value in node.attrib.items():
                if name.startswith(_RELATIONSHIP_PREFIXES) and value not in declared:
                    raise ValueError("assembly_relationship_missing")

    def finish(self) -> None:
        current = parse_xml_with_preserved_qnames(self.destination["[Content_Types].xml"])
        if _xml_hash(current) != _xml_hash(self.types):
            self.destination["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(
                self.types
            )


@dataclass(frozen=True)
class _WebAddInGraph:
    add_ins: tuple[tuple[str, str], ...]
    panes: tuple[tuple[str, ET.Element, ET.Element], ...]


class _WebAddInElement(_QNameElement):
    """Carry document-level comments/PIs through graph comparisons and copies."""

    __slots__ = ("_document_misc",)

    def __init__(self, tag: str, attrib: dict[str, str]) -> None:
        super().__init__(tag, attrib)
        self._document_misc: tuple[tuple[bool, str, str], ...] = ()

    def __deepcopy__(self, memo: dict[int, Any]) -> _WebAddInElement:
        clone = _WebAddInElement(self.tag, {})
        memo[id(self)] = clone
        clone.attrib = copy.deepcopy(self.attrib, memo)
        clone._namespace_scope = self._namespace_scope
        clone._document_misc = self._document_misc
        clone.text, clone.tail = self.text, self.tail
        clone.extend(copy.deepcopy(child, memo) for child in self)
        return clone


class _WebAddInXmlBuilder(_QNameTreeBuilder):
    """Retain comments/PIs with the existing namespace scopes and node limits."""

    def __init__(self) -> None:
        super().__init__(PackageLimits(), [0])
        self.builder = ET.TreeBuilder(element_factory=_WebAddInElement)
        self.outside: list[tuple[bool, str, str]] = []
        self.root_started = False

    def start(self, tag: str, attrs: dict[str, str]) -> ET.Element:
        if len(self.scopes) == 1:
            self.root_started = True
        return super().start(tag, attrs)

    def _misc(self, tag: Any, kind: str, text: str) -> None:
        if len(self.scopes) == 1:
            self.part_elements += 1
            self.total_elements[0] += 1
            limits = PackageLimits()
            if self.part_elements > limits.max_xml_elements:
                raise _XmlSafetyError("xml_element_limit", "XML node limit exceeded.")
            if self.total_elements[0] > limits.max_total_xml_elements:
                raise _XmlSafetyError("xml_total_element_limit", "XML node limit exceeded.")
            self.outside.append((not self.root_started, kind, text))
            return
        self.start(tag, {})
        self.data(text)
        self.end(tag)

    def comment(self, text: str) -> None:
        self._misc(ET.Comment, "comment", text)

    def pi(self, target: str, text: str | None) -> None:
        self._misc(ET.ProcessingInstruction, "pi", target + (" " + text if text else ""))

    def close(self) -> ET.Element:
        root = super().close()
        root._document_misc = tuple(self.outside)
        return root


def _web_add_in_xml(data: bytes) -> ET.Element:
    try:
        # Admission still handles byte limits, encodings and forbidden declarations.
        _parse_bounded_xml(data, limits=PackageLimits(), total_elements=[0], preserve_qnames=True)
        parser = ET.XMLParser(target=_WebAddInXmlBuilder())
        for offset in range(0, len(data), 64 * 1024):
            parser.feed(data[offset : offset + 64 * 1024])
        return parser.close()
    except (_XmlSafetyError, ET.ParseError) as error:
        raise ValueError("assembly_web_add_in_graph_invalid") from error


def _web_add_in_serialize(root: ET.Element) -> bytes:
    data = serialize_xml_with_preserved_qnames(root)
    leading: list[bytes] = []
    trailing: list[bytes] = []
    for before, kind, text in getattr(root, "_document_misc", ()):
        markup = ("<!--" + text + "-->" if kind == "comment" else "<?" + text + "?>").encode()
        (leading if before else trailing).append(markup)
    declaration, newline, body = data.partition(b"\n")
    return declaration + newline + b"".join(leading) + body + b"".join(trailing)


def _web_add_in_relationships(parts: Mapping[str, bytes], owner: str) -> ET.Element:
    name = _rels_name(owner)
    return _web_add_in_xml(parts[name]) if name in parts else ET.Element(_q(_PR, "Relationships"))


def _web_add_in_target(owner: str, target: str) -> str:
    """Resolve only local OPC part targets; never interpret a target as a fetch."""
    try:
        uri = urlsplit(target)
    except ValueError as error:
        raise ValueError("assembly_web_add_in_graph_invalid") from error
    path = unquote(uri.path)
    if (
        uri.scheme
        or uri.netloc
        or uri.query
        or uri.fragment
        or not path
        or "\\" in path
        or any(ord(character) < 32 for character in path)
    ):
        raise ValueError("assembly_web_add_in_link_unsupported")
    resolved = _resolve(owner, path)
    if resolved in {"", ".", "..", "[Content_Types].xml"} or resolved.startswith("../"):
        raise ValueError("assembly_web_add_in_graph_invalid")
    if "_rels" in resolved.split("/") or resolved.endswith(".rels"):
        raise ValueError("assembly_web_add_in_graph_invalid")
    return resolved


def _web_add_in_rels(parts: Mapping[str, bytes], types: ET.Element, owner: str) -> ET.Element:
    name = _rels_name(owner)
    if name in parts:
        found = _content_type(name, types)
        if found is None or found[1] != _OPC_RELS_TYPE:
            raise ValueError("assembly_web_add_in_content_type_invalid")
    root = _web_add_in_relationships(parts, owner)
    try:
        # The shared OPC validator consumes the comment-free admission projection.
        validation_root = copy.deepcopy(root)
        for parent in validation_root.iter():
            for child in list(parent):
                if child.tag in {ET.Comment, ET.ProcessingInstruction}:
                    if child.tail:
                        parent.text = (parent.text or "") + child.tail
                    parent.remove(child)
        validate_relationships_part(validation_root, name)
    except PackageError as error:
        raise ValueError("assembly_web_add_in_graph_invalid") from error
    occupied: set[str] = set()
    for relation in root.findall(_q(_PR, "Relationship")):
        identifier = relation.get("Id")
        target = relation.get("Target")
        if (
            not identifier
            or identifier in occupied
            or not relation.get("Type")
            or not target
            or relation.get("TargetMode", "Internal") not in {"Internal", "External"}
        ):
            raise ValueError("assembly_web_add_in_graph_invalid")
        occupied.add(identifier)
        if relation.get("TargetMode") != "External":
            part = _web_add_in_target(owner, target)
            if part not in parts:
                raise ValueError("assembly_web_add_in_graph_invalid")
    return root


def _web_add_in_identity(value: str) -> str:
    try:
        return str(uuid.UUID(value.strip("{}")))
    except ValueError:
        return value


def _web_add_in_graph(parts: Mapping[str, bytes], *, relocating: bool) -> _WebAddInGraph:
    """Discover package graphs; opaque pane markup need not be editable to survive."""
    types = _web_add_in_xml(parts["[Content_Types].xml"])
    declared = {
        name: found[1]
        for name in parts
        if (found := _content_type(name, types)) is not None
        and found[1] in {_WEB_EXTENSION_TYPE, _WEB_TASK_PANES_TYPE}
    }
    for override in types.findall(_q(_CT, "Override")):
        content_type = override.get("ContentType")
        if content_type in {_WEB_EXTENSION_TYPE, _WEB_TASK_PANES_TYPE}:
            part = override.get("PartName", "").lstrip("/")
            if part not in parts:
                raise ValueError("assembly_web_add_in_graph_invalid")
            if declared.get(part) != content_type:
                raise ValueError("assembly_web_add_in_content_type_invalid")
    root_rels = _web_add_in_relationships(parts, "")
    pane_edges = [
        relation
        for relation in root_rels.findall(_q(_PR, "Relationship"))
        if relation.get("Type") == _WEB_TASK_PANES_REL
    ]
    if not declared and not pane_edges:
        return _WebAddInGraph((), ())
    _web_add_in_rels(parts, types, "")
    if relocating:
        # Relocating a non-package owner could require story or body-link changes.
        pane_owners = {
            _rels_name(part)
            for part, content_type in declared.items()
            if content_type == _WEB_TASK_PANES_TYPE
        }
        for name, data in parts.items():
            if not name.endswith(".rels"):
                continue
            relationships = root_rels if name == _RELS else _web_add_in_xml(data)
            for relation in relationships.findall(_q(_PR, "Relationship")):
                kind = relation.get("Type")
                if (
                    kind == _WEB_TASK_PANES_REL
                    and name != _RELS
                    or kind == _WEB_EXTENSION_REL
                    and name not in pane_owners
                ):
                    raise ValueError("assembly_web_add_in_link_unsupported")
    add_ins: list[tuple[str, str]] = []
    identities: set[str] = set()
    for part in sorted(declared):
        if declared[part] != _WEB_EXTENSION_TYPE:
            continue
        root = _web_add_in_xml(parts[part])
        identifier = root.get("id")
        if root.tag != _q(_WEB_EXTENSION_NS, "webextension") or not identifier:
            if relocating:
                raise ValueError("assembly_web_add_in_graph_invalid")
            # Unchanged opaque destination XML is not an authoring request.
            continue
        identity = _web_add_in_identity(identifier)
        if relocating and identity in identities:
            raise ValueError("assembly_web_add_in_identity_collision")
        identities.add(identity)
        add_ins.append((part, identity))
        if not relocating:
            continue
        rels = _web_add_in_rels(parts, types, part)
        declared_ids = _ids(rels)
        if any(
            name.startswith(_RELATIONSHIP_PREFIXES) and value not in declared_ids
            for node in root.iter()
            for name, value in node.attrib.items()
        ):
            raise ValueError("assembly_web_add_in_graph_invalid")
        for relation in rels.findall(_q(_PR, "Relationship")):
            # Cached images are opaque leaves. Other links can bind to stories.
            if relation.get("Type") not in {f"{_R}/image", f"{_STRICT_R}/image"}:
                raise ValueError("assembly_web_add_in_link_unsupported")
            if relation.get("TargetMode") == "External":
                continue
            target = _web_add_in_target(part, relation.attrib["Target"])
            found = _content_type(target, types)
            if found is None or not found[1].startswith("image/"):
                raise ValueError("assembly_web_add_in_content_type_invalid")
            if _web_add_in_rels(parts, types, target).findall(_q(_PR, "Relationship")):
                raise ValueError("assembly_web_add_in_link_unsupported")
    panes: list[tuple[str, ET.Element, ET.Element]] = []
    linked: set[str] = set()
    for relation in pane_edges:
        if relation.get("TargetMode") == "External":
            if relocating:
                raise ValueError("assembly_web_add_in_link_unsupported")
            continue
        part = _web_add_in_target("", relation.attrib["Target"])
        if declared.get(part) != _WEB_TASK_PANES_TYPE:
            raise ValueError("assembly_web_add_in_content_type_invalid")
        if relocating and part in linked:
            raise ValueError("assembly_web_add_in_graph_unsupported")
        linked.add(part)
        root = _web_add_in_xml(parts[part])
        if relocating:
            if root.tag != _q(_WEB_TASK_PANES_NS, "taskpanes"):
                raise ValueError("assembly_web_add_in_graph_invalid")
            rels = _web_add_in_rels(parts, types, part)
            by_id = {edge.attrib["Id"]: edge for edge in rels.findall(_q(_PR, "Relationship"))}
            add_in_parts = {name for name, _ in add_ins}
            for edge in by_id.values():
                if edge.get("Type") != _WEB_EXTENSION_REL or edge.get("TargetMode") == "External":
                    raise ValueError("assembly_web_add_in_link_unsupported")
                if _web_add_in_target(part, edge.attrib["Target"]) not in add_in_parts:
                    raise ValueError("assembly_web_add_in_content_type_invalid")
            if any(
                name.startswith(_RELATIONSHIP_PREFIXES) and value not in by_id
                for node in root.iter()
                for name, value in node.attrib.items()
            ):
                raise ValueError("assembly_web_add_in_graph_invalid")
        panes.append((part, root, relation))
    if relocating and linked != {
        part for part, content_type in declared.items() if content_type == _WEB_TASK_PANES_TYPE
    }:
        raise ValueError("assembly_web_add_in_graph_unsupported")
    return _WebAddInGraph(tuple(add_ins), tuple(panes))


def _copy_web_add_in_part(importer: _Importer, source_part: str) -> str:
    """Relocate a validated opaque graph without touching destination members."""
    if source_part in importer.part_map:
        return importer.part_map[source_part]
    reserved = {name.casefold() for name in importer.destination}
    reserved.update(
        node.get("PartName", "").lstrip("/").casefold()
        for node in importer.types.findall(_q(_CT, "Override"))
    )
    directory, filename = posixpath.split(source_part)
    part = source_part
    serial = 1
    while part.casefold() in reserved or _rels_name(part).casefold() in reserved:
        suffix = "" if serial == 1 else f"-{serial}"
        part = posixpath.join(directory, f"assembly-{importer.label}{suffix}-{filename}")
        serial += 1
    importer.part_map[source_part] = part
    importer.destination[part] = importer.source[source_part]
    importer._copy_type(source_part, part)
    source_rels = _rels_name(source_part)
    if source_rels in importer.source:
        rels = _web_add_in_relationships(importer.source, source_part)
        changed = False
        for relation in rels.findall(_q(_PR, "Relationship")):
            if relation.get("TargetMode") == "External":
                continue
            old_target = relation.attrib["Target"]
            target = _copy_web_add_in_part(importer, _web_add_in_target(source_part, old_target))
            if _web_add_in_target(part, old_target) != target:
                relation.set("Target", quote(_relative(part, target), safe="/"))
                changed = True
        destination_rels = _rels_name(part)
        importer.destination[destination_rels] = (
            _web_add_in_serialize(rels) if changed else importer.source[source_rels]
        )
        importer._copy_type(source_rels, destination_rels)
    return part


def _web_add_in_fingerprint(root: ET.Element) -> str:
    """Conservative XML equality, including bindings used by opaque QName values."""
    root = copy.deepcopy(root)
    scopes = []
    for node in root.iter():
        attributes = sorted(node.attrib.items())
        node.attrib.clear()
        node.attrib.update(attributes)
        scopes.append(sorted(getattr(node, "_namespace_scope", ())))
    return hashlib.sha256(
        json.dumps(
            [_xml_hash(root), scopes, getattr(root, "_document_misc", ())],
            ensure_ascii=True,
            separators=(",", ":"),
        ).encode()
    ).hexdigest()


def _web_add_in_relation_fingerprint(relation: ET.Element, target: str) -> str:
    relation = copy.deepcopy(relation)
    relation.attrib.pop("Id", None)
    relation.set("Target", target)
    if relation.get("TargetMode") == "Internal":
        relation.attrib.pop("TargetMode")
    return _web_add_in_fingerprint(relation)


def _web_add_in_normalize_references(root: ET.Element, identifiers: Mapping[str, str]) -> None:
    for node in root.iter():
        for name, value in list(node.attrib.items()):
            if name.startswith(_RELATIONSHIP_PREFIXES):
                if value not in identifiers:
                    raise ValueError("assembly_web_add_in_graph_invalid")
                node.set(name, identifiers[value])


def _web_add_in_payload_fingerprint(
    parts: Mapping[str, bytes], types: ET.Element, part: str
) -> tuple[str, str]:
    """Include cached-image bytes/types and external links, not just the add-in id."""
    root = _web_add_in_xml(parts[part])
    root.set("id", _web_add_in_identity(root.attrib["id"]))
    rels = _web_add_in_rels(parts, types, part)
    identifiers: dict[str, str] = {}
    for relation in rels.findall(_q(_PR, "Relationship")):
        target = relation.attrib["Target"]
        if relation.get("TargetMode") != "External":
            target_part = _web_add_in_target(part, target)
            found = _content_type(target_part, types)
            image_rels = _web_add_in_rels(parts, types, target_part)
            if (
                relation.get("Type") not in {f"{_R}/image", f"{_STRICT_R}/image"}
                or found is None
                or not found[1].startswith("image/")
                or image_rels.findall(_q(_PR, "Relationship"))
            ):
                raise ValueError("assembly_web_add_in_identity_collision")
            target = (
                found[1]
                + ":"
                + hashlib.sha256(parts[target_part]).hexdigest()
                + ":"
                + _web_add_in_fingerprint(image_rels)
            )
        identifier = relation.attrib["Id"]
        fingerprint = _web_add_in_relation_fingerprint(relation, target)
        identifiers[identifier] = fingerprint
        relation.set("Id", fingerprint)
        relation.set("Target", target)
        if relation.get("TargetMode") == "Internal":
            relation.attrib.pop("TargetMode")
    _web_add_in_normalize_references(root, identifiers)
    return _web_add_in_fingerprint(root), _web_add_in_fingerprint(rels)


def _web_add_in_pane_edges(
    parts: Mapping[str, bytes], graph: _WebAddInGraph, owner: str
) -> tuple[ET.Element, dict[str, str], dict[str, str]]:
    rels = _web_add_in_relationships(parts, owner)
    identities = dict(graph.add_ins)
    targets: dict[str, str] = {}
    fingerprints: dict[str, str] = {}
    for relation in rels.findall(_q(_PR, "Relationship")):
        if relation.get("Type") != _WEB_EXTENSION_REL or relation.get("TargetMode") == "External":
            raise ValueError("assembly_web_add_in_link_unsupported")
        part = _web_add_in_target(owner, relation.attrib["Target"])
        if part not in identities:
            raise ValueError("assembly_web_add_in_content_type_invalid")
        identifier = relation.attrib["Id"]
        targets[identifier] = identities[part]
        fingerprints[identifier] = _web_add_in_relation_fingerprint(relation, identities[part])
    return rels, targets, fingerprints


def _web_add_in_collection_fingerprint(
    root: ET.Element, rels: ET.Element, fingerprints: Mapping[str, str], relation: ET.Element
) -> tuple[str, str, str]:
    root = copy.deepcopy(root)
    _web_add_in_normalize_references(root, fingerprints)
    rels = copy.deepcopy(rels)
    for edge in rels.findall(_q(_PR, "Relationship")):
        fingerprint = fingerprints[edge.attrib["Id"]]
        edge.set("Id", fingerprint)
        edge.set("Target", fingerprint)
        if edge.get("TargetMode") == "Internal":
            edge.attrib.pop("TargetMode")
    return (
        _web_add_in_fingerprint(root),
        _web_add_in_fingerprint(rels),
        _web_add_in_relation_fingerprint(relation, "taskpanes"),
    )


def _web_add_in_pane_container_fingerprint(root: ET.Element) -> str:
    root = copy.deepcopy(root)
    for pane in list(root.findall(_q(_WEB_TASK_PANES_NS, "taskpane"))):
        root.remove(pane)
    return _web_add_in_fingerprint(root)


def _web_add_in_pane_reference(reference: ET.Element) -> str:
    return reference.get(_q(_R, "id")) or reference.attrib[_q(_STRICT_R, "id")]


def _web_add_in_panes_mergeable(root: ET.Element, rels: ET.Element) -> bool:
    """Unknown pane markup is preserved independently, never guessed or discarded."""
    pane_tag = _q(_WEB_TASK_PANES_NS, "taskpane")
    ref_tag = _q(_WEB_TASK_PANES_NS, "webextensionref")
    if (
        root.tag != _q(_WEB_TASK_PANES_NS, "taskpanes")
        or getattr(root, "_document_misc", ())
        or getattr(rels, "_document_misc", ())
        or any(child.tag not in {pane_tag, ET.Comment, ET.ProcessingInstruction} for child in root)
        or (root.text or "").strip()
        or rels.attrib
        or any(child.tag != _q(_PR, "Relationship") for child in rels)
        or (rels.text or "").strip()
    ):
        return False
    references = list(root.iter(ref_tag))
    panes = root.findall(pane_tag)
    if len(references) != len(panes) or any(len(pane.findall(ref_tag)) != 1 for pane in panes):
        return False
    declared = _ids(rels)
    for node in root.iter():
        attributes = [
            (name, value)
            for name, value in node.attrib.items()
            if name.startswith(_RELATIONSHIP_PREFIXES)
        ]
        if node.tag == ref_tag:
            if (
                len(attributes) != 1
                or attributes[0][0] not in {_q(_R, "id"), _q(_STRICT_R, "id")}
                or attributes[0][1] not in declared
            ):
                return False
        elif attributes:
            return False
    return True


def _import_web_add_ins(importer: _Importer) -> None:
    """Deduplicate equal identities, retaining independent opaque pane collections.

    Add-in payloads, cached-image bytes/types and pane relationships participate
    in equality. Divergent overlapping identities refuse atomically. Unknown
    destination pane XML is never rewritten just to preserve it; distinct source
    collections can be relocated beside it. Body-linked controls still pass
    through the independent ``_supported_sdt`` refusal gate.
    """
    source = _web_add_in_graph(importer.source, relocating=True)
    if not source.add_ins and not source.panes:
        return
    destination = _web_add_in_graph(importer.destination, relocating=False)
    destination_parts: dict[str, list[str]] = {}
    for part, identity in destination.add_ins:
        destination_parts.setdefault(identity, []).append(part)
    overlap: set[str] = set()
    for part, identity in source.add_ins:
        matches = destination_parts.get(identity, [])
        if not matches:
            continue
        if len(matches) != 1 or _web_add_in_payload_fingerprint(
            importer.source, importer.source_types, part
        ) != _web_add_in_payload_fingerprint(importer.destination, importer.types, matches[0]):
            raise ValueError("assembly_web_add_in_identity_collision")
        overlap.add(identity)
        importer.part_map[part] = matches[0]
    # Prepare the unchanged destination collections before copying any new parts.
    destination_collections = []
    destination_relations = {part: relation for part, _, relation in destination.panes}
    for part, root, relation in destination.panes:
        try:
            rels, targets, fingerprints = _web_add_in_pane_edges(
                importer.destination, destination, part
            )
            signature = _web_add_in_collection_fingerprint(root, rels, fingerprints, relation)
            mergeable = _web_add_in_panes_mergeable(root, rels)
        except ValueError:
            # This collection remains opaque and untouched; it cannot be deduped.
            rels = _web_add_in_relationships(importer.destination, part)
            targets, fingerprints, signature, mergeable = {}, {}, None, False
        destination_collections.append(
            (part, root, rels, targets, fingerprints, signature, mergeable)
        )
    for part, _ in source.add_ins:
        _copy_web_add_in_part(importer, part)
    imported_graph = _WebAddInGraph(
        tuple(
            {
                **dict(destination.add_ins),
                **{importer.part_map[part]: identity for part, identity in source.add_ins},
            }.items()
        ),
        (),
    )
    pane_tag = _q(_WEB_TASK_PANES_NS, "taskpane")
    ref_tag = _q(_WEB_TASK_PANES_NS, "webextensionref")
    for source_part, source_root, root_relation in source.panes:
        source_rels, source_targets, source_fingerprints = _web_add_in_pane_edges(
            importer.source, source, source_part
        )
        signature = _web_add_in_collection_fingerprint(
            source_root, source_rels, source_fingerprints, root_relation
        )
        matching = next((entry for entry in destination_collections if entry[5] == signature), None)
        if matching is not None:
            importer.part_map[source_part] = matching[0]
            continue
        root = copy.deepcopy(source_root)
        rels = copy.deepcopy(source_rels)
        mergeable = _web_add_in_panes_mergeable(root, rels)
        collection_overlap = set(source_targets.values()) & overlap
        if collection_overlap:
            if not mergeable or any(
                not entry[6] and (not entry[3] or collection_overlap & set(entry[3].values()))
                for entry in destination_collections
            ):
                raise ValueError("assembly_web_add_in_pane_merge_unsupported")
            existing: dict[tuple[str, str], int] = {}
            associated: set[str] = set()
            existing_edges: set[str] = set()
            for (
                _,
                other_root,
                other_rels,
                targets,
                fingerprints,
                _,
                editable,
            ) in destination_collections:
                if not editable:
                    continue
                existing_edges.update(fingerprints.values())
                for pane in other_root.findall(pane_tag):
                    reference = pane.find(ref_tag)
                    assert reference is not None
                    identifier = _web_add_in_pane_reference(reference)
                    identity = targets[identifier]
                    if identity not in collection_overlap:
                        continue
                    if _web_add_in_pane_container_fingerprint(
                        other_root
                    ) != _web_add_in_pane_container_fingerprint(root):
                        raise ValueError("assembly_web_add_in_pane_merge_unsupported")
                    source_sidecar = copy.deepcopy(rels)
                    destination_sidecar = copy.deepcopy(other_rels)
                    for sidecar in (source_sidecar, destination_sidecar):
                        for edge in list(sidecar.findall(_q(_PR, "Relationship"))):
                            sidecar.remove(edge)
                    if _web_add_in_fingerprint(source_sidecar) != _web_add_in_fingerprint(
                        destination_sidecar
                    ):
                        raise ValueError("assembly_web_add_in_pane_merge_unsupported")
                    associated.add(identity)
                    normalized = copy.deepcopy(pane)
                    _web_add_in_normalize_references(normalized, fingerprints)
                    key = identity, _web_add_in_fingerprint(normalized)
                    existing[key] = existing.get(key, 0) + 1
            for pane in list(root.findall(pane_tag)):
                reference = pane.find(ref_tag)
                assert reference is not None
                identifier = _web_add_in_pane_reference(reference)
                identity = source_targets[identifier]
                if identity not in collection_overlap:
                    continue
                if (
                    _web_add_in_relation_fingerprint(root_relation, "taskpanes")
                    not in {
                        _web_add_in_relation_fingerprint(destination_relations[part], "taskpanes")
                        for part, _, _, targets, _, _, _ in destination_collections
                        if identity in targets.values()
                    }
                    and identity in associated
                ):
                    raise ValueError("assembly_web_add_in_pane_merge_unsupported")
                normalized = copy.deepcopy(pane)
                _web_add_in_normalize_references(normalized, source_fingerprints)
                key = identity, _web_add_in_fingerprint(normalized)
                if existing.get(key, 0):
                    existing[key] -= 1
                    root.remove(pane)
                elif identity in associated:
                    raise ValueError("assembly_web_add_in_identity_collision")
            used = {
                value
                for node in root.iter()
                for name, value in node.attrib.items()
                if name.startswith(_RELATIONSHIP_PREFIXES)
            }
            for edge in list(rels.findall(_q(_PR, "Relationship"))):
                identifier = edge.attrib["Id"]
                if identifier not in used and source_targets[identifier] in collection_overlap:
                    if source_fingerprints[identifier] not in existing_edges:
                        raise ValueError("assembly_web_add_in_identity_collision")
                    rels.remove(edge)
            if not root.findall(pane_tag) and not rels.findall(_q(_PR, "Relationship")):
                continue
        target = (
            destination_collections[0]
            if len(destination_collections) == 1
            and destination_collections[0][6]
            and mergeable
            and destination_collections[0][1].attrib == root.attrib
            and getattr(destination_collections[0][1], "_namespace_scope", ())
            == getattr(root, "_namespace_scope", ())
            and getattr(destination_collections[0][2], "_namespace_scope", ())
            == getattr(rels, "_namespace_scope", ())
            and _web_add_in_relation_fingerprint(
                destination_relations[destination_collections[0][0]], "taskpanes"
            )
            == _web_add_in_relation_fingerprint(root_relation, "taskpanes")
            else None
        )
        if target is None:
            # Independent ownership is valid OPC and preserves opaque XML bytes.
            part = _copy_web_add_in_part(importer, source_part)
            if collection_overlap:
                importer.destination[part] = _web_add_in_serialize(root)
                relocated_rels = copy.deepcopy(rels)
                for edge in relocated_rels.findall(_q(_PR, "Relationship")):
                    imported_part = importer.part_map[
                        _web_add_in_target(source_part, edge.attrib["Target"])
                    ]
                    edge.set("Target", quote(_relative(part, imported_part), safe="/"))
                importer.destination[_rels_name(part)] = _web_add_in_serialize(relocated_rels)
            package_rels = _web_add_in_relationships(importer.destination, "")
            relation = copy.deepcopy(root_relation)
            relation.set("Id", _new_rel_id(package_rels))
            relation.set("Target", quote(part, safe="/"))
            package_rels.append(relation)
            importer.destination[_RELS] = _web_add_in_serialize(package_rels)
            importer._copy_type(_RELS, _RELS)
            destination_relations[part] = relation
            imported_root = _web_add_in_xml(importer.destination[part])
            imported_rels, imported_targets, imported_fingerprints = _web_add_in_pane_edges(
                importer.destination, imported_graph, part
            )
            destination_collections.append(
                (
                    part,
                    imported_root,
                    imported_rels,
                    imported_targets,
                    imported_fingerprints,
                    _web_add_in_collection_fingerprint(
                        imported_root, imported_rels, imported_fingerprints, relation
                    ),
                    _web_add_in_panes_mergeable(imported_root, imported_rels),
                )
            )
            continue
        target_part, panes, target_rels, _, _, _, _ = target
        remap: dict[str, str] = {}
        for edge in rels.findall(_q(_PR, "Relationship")):
            imported = copy.deepcopy(edge)
            identifier = _new_rel_id(target_rels)
            remap[edge.attrib["Id"]] = identifier
            imported.set("Id", identifier)
            imported_part = importer.part_map[
                _web_add_in_target(source_part, edge.attrib["Target"])
            ]
            imported.set("Target", quote(_relative(target_part, imported_part), safe="/"))
            target_rels.append(imported)
        for pane in root:
            imported = copy.deepcopy(pane)
            _web_add_in_normalize_references(imported, remap)
            panes.append(imported)
        if len(root):
            importer.destination[target_part] = _web_add_in_serialize(panes)
        if remap:
            importer.destination[_rels_name(target_part)] = _web_add_in_serialize(target_rels)
            importer._copy_type(_rels_name(source_part), _rels_name(target_part))
        imported_rels, imported_targets, imported_fingerprints = _web_add_in_pane_edges(
            importer.destination, imported_graph, target_part
        )
        destination_collections[0] = (
            target_part,
            panes,
            imported_rels,
            imported_targets,
            imported_fingerprints,
            _web_add_in_collection_fingerprint(
                panes, imported_rels, imported_fingerprints, destination_relations[target_part]
            ),
            True,
        )


_NOTE_KINDS = {
    "footnote": ("footnotes", "footnote", "footnoteReference"),
    "endnote": ("endnotes", "endnote", "endnoteReference"),
}


@dataclass
class _NoteImports:
    roots: list[ET.Element]
    parts: dict[str, tuple[str, str, ET.Element]]


def _relationship_part(
    parts: Mapping[str, bytes],
    owner: str,
    relationship_type: str,
    *,
    code: str,
) -> str:
    matches = [
        relation
        for relation in _read_rels(parts, owner).findall(_q(_PR, "Relationship"))
        if relation.get("Type") == relationship_type
    ]
    if len(matches) != 1 or matches[0].get("TargetMode") is not None:
        raise ValueError(code)
    target = matches[0].get("Target")
    if target is None:
        raise ValueError(code)
    part = _resolve(owner, target)
    if not part.startswith("word/") or part not in parts:
        raise ValueError(code)
    return part


def _note_source_part(source: Mapping[str, bytes], kind: str, ns: str) -> tuple[str, ET.Element]:
    relationship_kind, element_name, _ = _NOTE_KINDS[kind]
    part = _relationship_part(
        source,
        _DOCUMENT,
        f"{_R}/{relationship_kind}",
        code="assembly_note_dependency_missing",
    )
    expected = (
        f"application/vnd.openxmlformats-officedocument.wordprocessingml.{relationship_kind}+xml"
    )
    if _content_type(part, ET.fromstring(source["[Content_Types].xml"])) != (
        "Override",
        expected,
    ):
        raise ValueError("assembly_note_dependency_missing")
    root = parse_xml_with_preserved_qnames(source[part])
    if _word_ns(root) != ns or root.tag != Q(ns, relationship_kind):
        raise ValueError("assembly_note_dependency_missing")
    if any(node.tag == Q(ns, element_name) and node.get(Q(ns, "id")) is None for node in root):
        raise ValueError("assembly_note_dependency_missing")
    return part, root


def _ensure_override(destination: dict[str, bytes], part: str, content_type: str) -> None:
    types = parse_xml_with_preserved_qnames(destination["[Content_Types].xml"])
    if not any(
        node.get("PartName") == "/" + part and node.get("ContentType") == content_type
        for node in types.findall(_q(_CT, "Override"))
    ):
        types.append(
            ET.Element(
                _q(_CT, "Override"),
                {"PartName": "/" + part, "ContentType": content_type},
            )
        )
    destination["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types)


def _add_document_relationship(
    destination: dict[str, bytes], relationship_type: str, target: str
) -> None:
    rels_name = _rels_name(_DOCUMENT)
    rels = _read_rels(destination, _DOCUMENT)
    rels.append(
        ET.Element(
            _q(_PR, "Relationship"),
            {
                "Id": _new_rel_id(rels),
                "Type": relationship_type,
                "Target": _relative(_DOCUMENT, target),
            },
        )
    )
    destination[rels_name] = serialize_xml_with_preserved_qnames(rels)


def _ensure_related_destination(
    destination: dict[str, bytes],
    relationship_type: str,
    content_type: str,
    root_tag: str,
    stem: str,
    code: str,
    *,
    create: bool = True,
    template: ET.Element | None = None,
) -> tuple[str | None, ET.Element | None, bool]:
    """Resolve the destination part one document relationship type names.

    Returns ``(part, root, created)``. A missing part is named but not written when
    ``create`` is true; the caller publishes it with :func:`_publish_related_destination`.
    A created root copies ``template``'s namespace declarations and attributes, so it
    keeps the source part's prefixes (and any mc:Ignorable they name).
    """
    matches = [
        relation
        for relation in _read_rels(destination, _DOCUMENT).findall(_q(_PR, "Relationship"))
        if relation.get("Type") == relationship_type
    ]
    if len(matches) > 1:
        raise ValueError(code)
    if matches:
        relation = matches[0]
        target = relation.get("Target")
        if relation.get("TargetMode") is not None or target is None:
            raise ValueError(code)
        part = _resolve(_DOCUMENT, target)
        if part not in destination or _content_type(
            part, ET.fromstring(destination["[Content_Types].xml"])
        ) != ("Override", content_type):
            raise ValueError(code)
        root = parse_xml_with_preserved_qnames(destination[part])
        if root.tag != root_tag:
            raise ValueError(code)
        return part, root, False
    if not create:
        return None, None, False
    part = f"word/{stem}.xml"
    serial = 2
    while part in destination:
        part = f"word/assembly-{stem}-{serial}.xml"
        serial += 1
    if template is None:
        return part, ET.Element(root_tag), True
    root = copy.copy(template)
    for child in list(root):
        root.remove(child)
    root.text, root.attrib = None, dict(template.attrib)
    return part, root, True


def _publish_related_destination(
    destination: dict[str, bytes],
    part: str,
    root: ET.Element,
    relationship_type: str,
    content_type: str,
) -> None:
    destination[part] = serialize_xml_with_preserved_qnames(root)
    _ensure_override(destination, part, content_type)
    _add_document_relationship(destination, relationship_type, part)


def _ensure_note_destination(
    destination: dict[str, bytes],
    source_part: str,
    source_root: ET.Element,
    kind: str,
    ns: str,
) -> tuple[str, ET.Element]:
    relationship_kind, element_name, _ = _NOTE_KINDS[kind]
    relationship_type = f"{_R}/{relationship_kind}"
    content_type = f"{_WML_CONTENT_TYPE}{relationship_kind}+xml"
    part, root, created = _ensure_related_destination(
        destination,
        relationship_type,
        content_type,
        Q(ns, relationship_kind),
        relationship_kind,
        "assembly_destination_note_invalid",
    )
    assert part is not None and root is not None
    if created:
        for note in source_root.findall(Q(ns, element_name)):
            note_id = note.get(Q(ns, "id"))
            if note_id is not None and note_id.lstrip("-").isdigit() and int(note_id) <= 0:
                root.append(copy.deepcopy(note))
        _publish_related_destination(destination, part, root, relationship_type, content_type)
    return part, root


def _note_references(roots: list[ET.Element], ns: str) -> list[tuple[str, str]]:
    references: list[tuple[str, str]] = []
    names = {reference_name: kind for kind, (_, _, reference_name) in _NOTE_KINDS.items()}
    for root in roots:
        for node in root.iter():
            kind = names.get(node.tag.rsplit("}", 1)[-1])
            if kind is None:
                continue
            note_id = node.get(Q(ns, "id"))
            if note_id is None or not note_id.isdigit() or int(note_id) < 1:
                raise ValueError("assembly_note_reference_invalid")
            references.append((kind, note_id))
    return references


def _collect_note_imports(
    destination: dict[str, bytes],
    source: Mapping[str, bytes],
    imported_roots: list[ET.Element],
    ns: str,
    *,
    standalone: ET.Element | None = None,
) -> _NoteImports:
    pending = _note_references(imported_roots, ns)
    if standalone is not None:
        kind = standalone.tag.rsplit("}", 1)[-1]
        identifier = standalone.get(Q(ns, "id"), "")
        if kind not in _NOTE_KINDS or not identifier.isdigit() or int(identifier) < 1:
            raise ValueError("node_import_note_entry_invalid")
        pending.append((kind, identifier))
    source_parts: dict[str, tuple[str, ET.Element]] = {}
    destination_parts: dict[str, tuple[str, ET.Element]] = {}
    source_notes: dict[str, dict[str, ET.Element]] = {}
    remaps: dict[str, dict[str, str]] = {}
    selected: list[ET.Element] = []
    while pending:
        kind, old_id = pending.pop()
        if kind not in source_parts:
            source_part, source_root = _note_source_part(source, kind, ns)
            source_parts[kind] = (source_part, source_root)
            _, element_name, _ = _NOTE_KINDS[kind]
            by_id: dict[str, ET.Element] = {}
            for note in source_root.findall(Q(ns, element_name)):
                note_id = note.get(Q(ns, "id"))
                if note_id is None:
                    continue
                if note_id in by_id:
                    raise ValueError("assembly_note_dependency_missing")
                by_id[note_id] = note
            source_notes[kind] = by_id
            destination_parts[kind] = _ensure_note_destination(
                destination, source_part, source_root, kind, ns
            )
            remaps[kind] = {}
        if old_id in remaps[kind]:
            continue
        note = source_notes[kind].get(old_id)
        if note is None:
            raise ValueError("assembly_note_dependency_missing")
        destination_root = destination_parts[kind][1]
        _, element_name, _ = _NOTE_KINDS[kind]
        used = {
            int(value)
            for node in destination_root.findall(Q(ns, element_name))
            if (value := node.get(Q(ns, "id"))) is not None and value.isdigit()
        }
        candidate = int(old_id)
        is_entry = (
            standalone is not None
            and standalone.tag == Q(ns, kind)
            and standalone.get(Q(ns, "id")) == old_id
        )
        if candidate in used or is_entry:
            candidate = max(used, default=0) + 1
            while candidate in used:
                candidate += 1
        remaps[kind][old_id] = str(candidate)
        imported = standalone if is_entry else copy.deepcopy(note)
        imported.set(Q(ns, "id"), str(candidate))
        destination_root.append(imported)
        selected.append(imported)
        pending.extend(_note_references([imported], ns))
    for root in dict.fromkeys([*imported_roots, *selected]):
        for node in root.iter():
            local = node.tag.rsplit("}", 1)[-1]
            for kind, (_, _, reference_name) in _NOTE_KINDS.items():
                if local == reference_name and (old_id := node.get(Q(ns, "id"))) in remaps[kind]:
                    node.set(Q(ns, "id"), remaps[kind][old_id])
    return _NoteImports(
        selected,
        {
            kind: (
                source_parts[kind][0],
                destination_parts[kind][0],
                destination_parts[kind][1],
            )
            for kind in source_parts
        },
    )


def _write_note_imports(destination: dict[str, bytes], imports: _NoteImports) -> None:
    for _, destination_part, root in imports.parts.values():
        destination[destination_part] = serialize_xml_with_preserved_qnames(root)


# Word comments: the markers that name a w:comment by its w:id, and the modern metadata
# parts keyed by the comment's last paragraph w14:paraId (commentsExtended, commentsIds)
# or by the commentsIds durable id (commentsExtensible); people rows are keyed by author.
_COMMENT_MARKERS = ("commentRangeStart", "commentRangeEnd", "commentReference")
_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
_COMMENT_METADATA = {
    "commentsExtended": (
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        f"{{{_W15}}}commentsEx",
        f"{{{_W15}}}commentEx",
    ),
    "commentsIds": (
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        f"{{{_W16CID}}}commentsIds",
        f"{{{_W16CID}}}commentId",
    ),
    "commentsExtensible": (
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
        f"{{{_W16CEX}}}commentsExtensible",
        f"{{{_W16CEX}}}commentExtensible",
    ),
    "people": (
        "http://schemas.microsoft.com/office/2011/relationships/people",
        f"{{{_W15}}}people",
        f"{{{_W15}}}person",
    ),
}


@dataclass
class _CommentImports:
    roots: list[ET.Element]
    source_part: str | None = None
    part: str | None = None
    root: ET.Element | None = None
    # The destination comments part before import, for per-part identity domains.
    baseline: ET.Element | None = None
    # Paragraph ids already used by destination comments and their metadata rows.
    para_ids: set[str] | None = None
    # (last paragraph of an imported comment, its source w14:paraId).
    links: list[tuple[ET.Element, str]] | None = None
    metadata: dict[str, ET.Element] | None = None


def _comment_id(value: str | None, code: str) -> int:
    if value is None or not re.fullmatch(r"[0-9]+", value):
        raise ValueError(code)
    return int(value)


def _comment_markers(roots: list[ET.Element], ns: str) -> dict[int, list[str]]:
    """Every comment id the markers below ``roots`` name, with the marker kinds."""
    tags = {Q(ns, name): name for name in _COMMENT_MARKERS}
    markers: dict[int, list[str]] = {}
    for root in roots:
        for node in root.iter():
            if (name := tags.get(node.tag)) is not None:
                identifier = _comment_id(
                    node.get(Q(ns, "id")), "assembly_comment_reference_invalid"
                )
                markers.setdefault(identifier, []).append(name)
    return markers


def _destination_comment_ids(
    destination: Mapping[str, bytes], destination_root: ET.Element, ns: str
) -> set[int]:
    """Comment ids destination markers already name, so imports never adopt them."""
    roots = [destination_root]
    for part, data in destination.items():
        if part == _DOCUMENT or not part.startswith("word/") or not part.endswith(".xml"):
            continue
        if b"commentR" not in data:
            continue
        try:
            roots.append(ET.fromstring(data))
        except ET.ParseError:
            continue
    tags = {Q(ns, name) for name in _COMMENT_MARKERS}
    used: set[int] = set()
    for root in roots:
        for node in root.iter():
            value = node.get(Q(ns, "id")) if node.tag in tags else None
            if value is not None and re.fullmatch(r"[0-9]+", value):
                used.add(int(value))
    return used


def _source_comment_metadata(source: Mapping[str, bytes]) -> dict[str, ET.Element]:
    types = ET.fromstring(source["[Content_Types].xml"])
    metadata: dict[str, ET.Element] = {}
    for kind, (relationship_type, root_tag, _) in _COMMENT_METADATA.items():
        if not any(
            relation.get("Type") == relationship_type
            for relation in _read_rels(source, _DOCUMENT).findall(_q(_PR, "Relationship"))
        ):
            continue
        part = _relationship_part(
            source, _DOCUMENT, relationship_type, code="assembly_comment_metadata_invalid"
        )
        if _content_type(part, types) != ("Override", f"{_WML_CONTENT_TYPE}{kind}+xml"):
            raise ValueError("assembly_comment_metadata_invalid")
        root = parse_xml_with_preserved_qnames(source[part])
        if root.tag != root_tag:
            raise ValueError("assembly_comment_metadata_invalid")
        metadata[kind] = root
    return metadata


def _collect_comment_imports(
    destination: dict[str, bytes],
    source: Mapping[str, bytes],
    imported_roots: list[ET.Element],
    destination_root: ET.Element,
    ns: str,
    *,
    standalone: ET.Element | None = None,
) -> _CommentImports:
    """Copy every comment imported markers name, renumbering only colliding ids.

    Each imported comment must arrive whole: one reference and, when ranged, one
    start and one end. Non-colliding ids are reserved before any collision is
    renumbered, so a fresh id never takes an incoming one.
    """
    markers = _comment_markers(imported_roots, ns)
    if standalone is not None:
        identifier = _comment_id(standalone.get(Q(ns, "id")), "node_import_comment_entry_invalid")
        markers[identifier] = ["commentReference"]
    if not markers:
        return _CommentImports([])
    for names in markers.values():
        starts = names.count("commentRangeStart")
        if (
            names.count("commentReference") != 1
            or starts > 1
            or starts != names.count("commentRangeEnd")
        ):
            raise ValueError("assembly_comment_range_incomplete")
    relationship_type = f"{_R}/comments"
    content_type = f"{_WML_CONTENT_TYPE}comments+xml"
    source_part = _relationship_part(
        source, _DOCUMENT, relationship_type, code="assembly_comment_dependency_missing"
    )
    if _content_type(source_part, ET.fromstring(source["[Content_Types].xml"])) != (
        "Override",
        content_type,
    ):
        raise ValueError("assembly_comment_dependency_missing")
    source_root = parse_xml_with_preserved_qnames(source[source_part])
    if source_root.tag != Q(ns, "comments"):
        raise ValueError("assembly_comment_dependency_missing")
    by_id: dict[int, ET.Element] = {}
    for comment in source_root.findall(Q(ns, "comment")):
        identifier = _comment_id(comment.get(Q(ns, "id")), "assembly_comment_dependency_missing")
        if identifier in by_id:
            raise ValueError("assembly_comment_dependency_missing")
        by_id[identifier] = comment
    if any(identifier not in by_id for identifier in markers):
        raise ValueError("assembly_comment_dependency_missing")
    if _comment_markers([by_id[identifier] for identifier in markers], ns):
        raise ValueError("assembly_comment_reference_invalid")
    metadata = _source_comment_metadata(source)
    part, root, created = _ensure_related_destination(
        destination,
        relationship_type,
        content_type,
        Q(ns, "comments"),
        "comments",
        "assembly_destination_comment_invalid",
        template=source_root,
    )
    assert part is not None and root is not None
    used = _destination_comment_ids(destination, destination_root, ns)
    used.update(
        _comment_id(comment.get(Q(ns, "id")), "assembly_destination_comment_invalid")
        for comment in root.findall(Q(ns, "comment"))
    )
    para_ids = collect_existing_para_ids({part: root})
    for kind in ("commentsExtended", "commentsIds"):
        relationship, root_tag, _ = _COMMENT_METADATA[kind]
        _, existing, _ = _ensure_related_destination(
            destination,
            relationship,
            f"{_WML_CONTENT_TYPE}{kind}+xml",
            root_tag,
            kind,
            "assembly_destination_comment_invalid",
            create=False,
        )
        if existing is not None:
            para_ids |= collect_existing_para_ids({kind: existing})
    baseline = copy.deepcopy(root)
    remap = {
        identifier: identifier
        for identifier in markers
        if identifier not in used
        and (standalone is None or identifier != int(standalone.get(Q(ns, "id"))))
    }
    used.update(remap)
    for identifier in sorted(markers):
        if identifier not in remap:
            remap[identifier] = max(used, default=-1) + 1
            used.add(remap[identifier])
    imported: list[ET.Element] = []
    links: list[tuple[ET.Element, str]] = []
    for identifier in sorted(markers):
        comment = (
            standalone
            if standalone is not None and identifier == int(standalone.get(Q(ns, "id")))
            else copy.deepcopy(by_id[identifier])
        )
        comment.set(Q(ns, "id"), str(remap[identifier]))
        root.append(comment)
        imported.append(comment)
        last = last_comment_paragraph(comment, ns)
        if last is not None and (para_id := last.get(_PARA_ID)) is not None:
            links.append((last, para_id.upper()))
    tags = {Q(ns, name) for name in _COMMENT_MARKERS}
    for element in imported_roots:
        for node in element.iter():
            if node.tag in tags:
                node.set(Q(ns, "id"), str(remap[int(node.get(Q(ns, "id")))]))
    if created:
        _publish_related_destination(destination, part, root, relationship_type, content_type)
    return _CommentImports(imported, source_part, part, root, baseline, para_ids, links, metadata)


def _comment_metadata_rows(
    kind: str,
    source_root: ET.Element,
    para_map: Mapping[str, str],
    durable_map: dict[str, str],
    durable_ids: set[str],
    authors: set[str],
) -> list[ET.Element]:
    """Copy the source rows bound to imported comments with their remapped links."""
    _, _, row_tag = _COMMENT_METADATA[kind]
    rows: list[ET.Element] = []
    for source_row in source_root.findall(row_tag):
        row = copy.deepcopy(source_row)
        if kind == "commentsExtended":
            para_id = (row.get(f"{{{_W15}}}paraId") or "").upper()
            if para_id not in para_map:
                continue
            row.set(f"{{{_W15}}}paraId", para_map[para_id])
            if (parent := row.get(f"{{{_W15}}}paraIdParent")) is not None:
                if parent.upper() not in para_map:
                    raise ValueError("assembly_comment_thread_incomplete")
                row.set(f"{{{_W15}}}paraIdParent", para_map[parent.upper()])
        elif kind == "commentsIds":
            para_id = (row.get(f"{{{_W16CID}}}paraId") or "").upper()
            if para_id not in para_map:
                continue
            row.set(f"{{{_W16CID}}}paraId", para_map[para_id])
            if (durable := row.get(f"{{{_W16CID}}}durableId")) is not None:
                value = durable.upper()
                if value in durable_ids:
                    fresh = generate_collision_safe_durable_id(durable_ids, value)
                else:
                    durable_ids.add(value)
                    fresh = durable
                durable_map[value] = fresh
                row.set(f"{{{_W16CID}}}durableId", fresh)
        elif kind == "commentsExtensible":
            durable = (row.get(f"{{{_W16CEX}}}durableId") or "").upper()
            if durable not in durable_map:
                continue
            row.set(f"{{{_W16CEX}}}durableId", durable_map[durable])
        elif row.get(f"{{{_W15}}}author") not in authors:
            continue
        rows.append(row)
    return rows


def _remap_comment_identities(
    imports: _CommentImports,
    ns: str,
    paragraph_ids: set[str],
    revision_ids: set,
    *,
    allow_revisions: bool = False,
) -> None:
    """Keep imported comment paragraphs, drawings and revisions unique in the destination."""
    assert imports.baseline is not None and imports.para_ids is not None
    paragraph_ids |= imports.para_ids
    _remap_paragraph_ids(imports.roots, ns, paragraph_ids)
    _remap_drawing_ids(imports.roots, _drawing_ids(imports.baseline))
    _remap_vml_ids(imports.roots, ns, _vml_ids(imports.baseline))
    _remap_anchor_ids(imports.roots, _anchor_ids(imports.baseline))
    _remap_conflict_ids(imports.roots, ns, revision_ids, allow_revisions=allow_revisions)


def _write_comment_imports(
    destination: dict[str, bytes], imports: _CommentImports, ns: str
) -> None:
    """Write the comments part and the metadata rows linked to imported comments."""
    if imports.part is None or imports.root is None:
        return
    destination[imports.part] = serialize_xml_with_preserved_qnames(imports.root)
    para_map = {old: paragraph.get(_PARA_ID, old).upper() for paragraph, old in imports.links or []}
    authors = {
        author for comment in imports.roots if (author := comment.get(Q(ns, "author"))) is not None
    }
    targets: dict[str, tuple[str | None, ET.Element | None, bool]] = {}
    durable_ids: set[str] = set()
    for kind, (relationship_type, root_tag, _) in _COMMENT_METADATA.items():
        targets[kind] = _ensure_related_destination(
            destination,
            relationship_type,
            f"{_WML_CONTENT_TYPE}{kind}+xml",
            root_tag,
            kind,
            "assembly_destination_comment_invalid",
            template=(imports.metadata or {}).get(kind),
        )
        if kind in {"commentsIds", "commentsExtensible"} and not targets[kind][2]:
            durable_ids |= collect_existing_durable_ids({kind: targets[kind][1]})
    durable_map: dict[str, str] = {}
    for kind, source_root in (imports.metadata or {}).items():
        part, root, created = targets[kind]
        assert part is not None and root is not None
        if kind == "people":
            authors -= {person.get(f"{{{_W15}}}author") for person in root}
        rows = _comment_metadata_rows(
            kind, source_root, para_map, durable_map, durable_ids, authors
        )
        if not rows:
            continue
        root.extend(rows)
        relationship_type = _COMMENT_METADATA[kind][0]
        if created:
            _publish_related_destination(
                destination, part, root, relationship_type, f"{_WML_CONTENT_TYPE}{kind}+xml"
            )
        else:
            destination[part] = serialize_xml_with_preserved_qnames(root)


def _theme_color(value: str, colors: Mapping[str, str]) -> str:
    color = colors.get(value.casefold())
    if color is None or not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
        raise ValueError("assembly_theme_dependency_missing")
    return color.upper()


def _tint_or_shade(value: str, tint: str | None, shade: str | None, *, hsl: bool = True) -> str:
    """Transform color with one final nearest-byte, half-up quantization.

    CT_Shd/CT_Border use HSL luminance (ECMA-376-1 17.3.4-5);
    run CT_Color uses RGB channels (17.3.2.6). Callers select precedence.
    """
    for transform in (tint, shade):
        if transform is not None and not re.fullmatch(r"[0-9A-Fa-f]{2}", transform):
            raise ValueError("assembly_theme_dependency_missing")
    if tint is None and shade is None:
        return value.upper()
    channels = [int(value[index : index + 2], 16) for index in range(0, 6, 2)]
    if not hsl:
        if tint is not None:
            factor = int(tint, 16)
            channels = [255 - ((255 - channel) * factor + 127) // 255 for channel in channels]
        if shade is not None:
            factor = int(shade, 16)
            channels = [(channel * factor + 127) // 255 for channel in channels]
        return "".join(f"{channel:02X}" for channel in channels)

    # Work in doubled byte luminance: total = 510*L. Keeping H and S
    # constant scales each channel's offset from L by the chroma ratio.
    # Integer ratios avoid floating drift and intermediate RGB rounding.
    total = min(channels) + max(channels)
    numerator, denominator = total, 1
    if tint is not None:
        factor = int(tint, 16)
        numerator = numerator * factor + 510 * denominator * (255 - factor)
        denominator *= 255
    if shade is not None:
        numerator *= int(shade, 16)
        denominator *= 255
    span = min(total, 510 - total)
    if span == 0:
        channel = (numerator + denominator) // (2 * denominator)
        return f"{channel:02X}" * 3
    transformed_span = min(numerator, 510 * denominator - numerator)
    divisor = 2 * denominator * span
    return "".join(
        f"{(numerator * span + (2 * channel - total) * transformed_span + divisor // 2) // divisor:02X}"
        for channel in channels
    )


def _theme_maps(
    source: Mapping[str, bytes],
    ns: str,
    *,
    owner: str = _DOCUMENT,
) -> tuple[dict[str, str], dict[str, str]]:
    # Theme authoring is an optional feature that itself imports assembly.
    _color_value = feature_module("theme_authoring")._color_value

    themed_owner = (
        owner
        if any(item.get("Type", "").endswith("/theme") for item in _read_rels(source, owner))
        else _DOCUMENT
    )
    theme_part = _relationship_part(
        source, themed_owner, f"{_R}/theme", code="assembly_theme_dependency_missing"
    )
    theme = ET.fromstring(source[theme_part])
    overrides = [
        item
        for item in _read_rels(source, owner)
        if item.get("Type", "").endswith("/themeOverride")
    ]
    if len(overrides) > 1:
        raise ValueError("node_import_theme_override_invalid")
    for relation in overrides:
        if relation.get("TargetMode") == "External":
            raise ValueError("node_import_theme_override_invalid")
        part = _resolve(owner, relation.get("Target", ""))
        if part not in source:
            raise ValueError("node_import_theme_override_invalid")
        override = ET.fromstring(source[part])
        for replacement in override:
            local = replacement.tag.rsplit("}", 1)[-1]
            parents = {child: parent for parent in theme.iter() for child in parent}
            existing = next(
                (item for item in theme.iter() if item.tag.rsplit("}", 1)[-1] == local), None
            )
            if existing is not None and existing in parents:
                parent = parents[existing]
                position = list(parent).index(existing)
                parent.remove(existing)
                parent.insert(position, copy.deepcopy(replacement))
    scheme = next(
        (node for node in theme.iter() if node.tag.rsplit("}", 1)[-1] == "clrScheme"),
        None,
    )
    font_scheme = next(
        (node for node in theme.iter() if node.tag.rsplit("}", 1)[-1] == "fontScheme"),
        None,
    )
    if scheme is None or font_scheme is None:
        raise ValueError("assembly_theme_dependency_missing")
    colors: dict[str, str] = {}
    for entry in scheme:
        color_value = _color_value(entry)
        if color_value is not None:
            colors[entry.tag.rsplit("}", 1)[-1].casefold()] = color_value
    mapping = {
        "bg1": "lt1",
        "tx1": "dk1",
        "bg2": "lt2",
        "tx2": "dk2",
        "dark1": "dk1",
        "light1": "lt1",
        "dark2": "dk2",
        "light2": "lt2",
        "hyperlink": "hlink",
        "followedhyperlink": "folhlink",
    }
    settings_relation = next(
        (
            item
            for item in _read_rels(source, _DOCUMENT)
            if item.get("Type", "").endswith("/settings")
        ),
        None,
    )
    settings_part = (
        _resolve(_DOCUMENT, settings_relation.get("Target", ""))
        if settings_relation is not None and settings_relation.get("TargetMode") != "External"
        else None
    )
    if settings_part is not None:
        settings = ET.fromstring(source[settings_part])
        scheme_mapping = settings.find(Q(ns, "clrSchemeMapping"))
        if scheme_mapping is not None:
            scheme_names = {
                "t1": "tx1",
                "t2": "tx2",
            }
            color_names = {
                "light1": "lt1",
                "dark1": "dk1",
                "light2": "lt2",
                "dark2": "dk2",
                "hyperlink": "hlink",
                "followedhyperlink": "folhlink",
            }
            mapping.update(
                {
                    scheme_names.get(
                        name.rsplit("}", 1)[-1].casefold(),
                        name.rsplit("}", 1)[-1].casefold(),
                    ): color_names.get(target.casefold(), target.casefold())
                    for name, target in scheme_mapping.attrib.items()
                }
            )
    if owner != _DOCUMENT and owner in source:
        related_root = ET.fromstring(source[owner])
        color_override = next(
            (item for item in related_root.iter() if item.tag.rsplit("}", 1)[-1] == "clrMapOvr"),
            None,
        )
        if color_override is not None:
            override = next(
                (
                    item
                    for item in color_override
                    if item.tag.rsplit("}", 1)[-1] == "overrideClrMapping"
                ),
                None,
            )
            if override is not None:
                mapping.update(
                    {name.casefold(): target.casefold() for name, target in override.attrib.items()}
                )
    palette = dict(colors)
    for name, target in mapping.items():
        if target in palette:
            colors[name] = palette[target]
    for name, target in {
        "background1": "bg1",
        "text1": "tx1",
        "background2": "bg2",
        "text2": "tx2",
    }.items():
        if target in colors:
            colors[name] = colors[target]
    fonts: dict[str, str] = {}
    for family in font_scheme:
        family_name = family.tag.rsplit("}", 1)[-1]
        if family_name not in {"majorFont", "minorFont"}:
            continue
        prefix = "major" if family_name == "majorFont" else "minor"
        for child in family:
            local = child.tag.rsplit("}", 1)[-1]
            typeface = child.get("typeface")
            if local == "latin" and typeface is not None:
                fonts[f"{prefix}ascii"] = typeface
                fonts[f"{prefix}hansi"] = typeface
            elif local == "ea" and typeface is not None:
                fonts[f"{prefix}eastasia"] = typeface
            elif local == "cs" and typeface is not None:
                fonts[f"{prefix}bidi"] = typeface
            elif local == "font" and typeface:
                script = child.get("script")
                if script:
                    script = script.casefold()
                    fonts[f"{prefix}bidi:{script}"] = typeface
                    fonts[f"{prefix}eastasia:{script}"] = typeface
    return colors, fonts


def _text_scripts(roots: list[ET.Element], ns: str) -> set[str]:
    scripts: set[str] = set()
    for root in roots:
        for text in root.iter(Q(ns, "t")):
            for character in text.text or "":
                if not unicodedata.category(character).startswith("L"):
                    continue
                name = unicodedata.name(character, "")
                if name.startswith(("LATIN ", "GREEK ", "CYRILLIC ")):
                    continue
                if "CJK UNIFIED IDEOGRAPH" in name or "CJK COMPATIBILITY IDEOGRAPH" in name:
                    scripts.add("han")
                elif "HIRAGANA" in name or "KATAKANA" in name:
                    scripts.add("jpan")
                elif name.startswith("HANGUL "):
                    scripts.add("hang")
                elif name.startswith("ARABIC "):
                    scripts.add("arab")
                elif name.startswith("HEBREW "):
                    scripts.add("hebr")
                elif name.startswith("THAI "):
                    scripts.add("thai")
                elif name.startswith("DEVANAGARI "):
                    scripts.add("deva")
                else:
                    scripts.add("unsupported")
    return scripts


def _resolve_theme_dependencies(
    source: Mapping[str, bytes],
    roots: list[ET.Element],
    ns: str,
    *,
    resolve_colors: bool = True,
    owner: str = _DOCUMENT,
) -> None:
    parent = {child: element for root in roots for element in root.iter() for child in element}
    themed = [
        node
        for root in roots
        for node in root.iter()
        if node.tag.startswith(f"{{{ns}}}")
        and any(
            name.startswith(f"{{{ns}}}")
            and (
                name.rsplit("}", 1)[-1].casefold().startswith("theme")
                or name.rsplit("}", 1)[-1].casefold().endswith("theme")
            )
            for name in node.attrib
        )
    ]
    drawing_namespaces = {
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
        _W14,
    }
    scheme_colors = [
        item
        for root in roots
        for item in root.iter()
        if resolve_colors and item.tag in {Q(uri, "schemeClr") for uri in drawing_namespaces}
    ]
    if not themed and not scheme_colors:
        return
    if not resolve_colors and not any(
        name.rsplit("}", 1)[-1].casefold().endswith("theme")
        for node in themed
        for name in node.attrib
    ):
        return
    colors, fonts = _theme_maps(source, ns, owner=owner)
    for item in scheme_colors:
        uri = item.tag[1:].split("}", 1)[0]
        attribute = Q(uri, "val") if uri == _W14 else "val"
        value = item.get(attribute)
        if value is None:
            raise ValueError("assembly_theme_dependency_missing")
        item.tag = Q(uri, "srgbClr")
        item.set(attribute, _theme_color(value, colors))
    scripts = {
        "ar": "arab",
        "fa": "arab",
        "ur": "arab",
        "he": "hebr",
        "yi": "hebr",
        "th": "thai",
        "hi": "deva",
        "mr": "deva",
        "ne": "deva",
        "ja": "jpan",
        "ko": "hang",
    }
    observed_scripts = _text_scripts(roots, ns)
    east_asian_scripts = {"han", "hang", "jpan", "unsupported"}
    complex_scripts = {"arab", "deva", "hebr", "thai", "unsupported"}
    for node in themed:
        attributes = {
            name.rsplit("}", 1)[-1]: (name, value)
            for name, value in node.attrib.items()
            if name.startswith(f"{{{ns}}}")
        }
        for local, (_, value) in tuple(attributes.items()):
            normalized = local.casefold()
            if not resolve_colors and normalized in {
                "themecolor",
                "themetint",
                "themeshade",
                "themefill",
                "themefilltint",
                "themefillshade",
            }:
                continue
            if local == "themeColor":
                node.set(
                    Q(ns, "val" if node.tag == Q(ns, "color") else "color"),
                    _tint_or_shade(
                        _theme_color(value, colors),
                        attributes.get("themeTint", (None, None))[1],
                        None
                        if "themeTint" in attributes
                        else attributes.get("themeShade", (None, None))[1],
                        hsl=node.tag != Q(ns, "color"),
                    ),
                )
            elif local == "themeFill":
                # Word 16.113.2 resolves themeFillTint over themeFillShade,
                # matching the foreground rule; shade applies only alone.
                node.set(
                    Q(ns, "fill"),
                    _tint_or_shade(
                        _theme_color(value, colors),
                        attributes.get("themeFillTint", (None, None))[1],
                        None
                        if "themeFillTint" in attributes
                        else attributes.get("themeFillShade", (None, None))[1],
                    ),
                )
            elif normalized.endswith("theme"):
                target = {
                    "asciitheme": "ascii",
                    "hansitheme": "hAnsi",
                    "eastasiatheme": "eastAsia",
                    "cstheme": "cs",
                }.get(normalized)
                font = fonts.get(value.casefold())
                if target is None or font is None:
                    raise ValueError("assembly_theme_dependency_missing")
                if not font and target in {"eastAsia", "cs"}:
                    properties = parent.get(node)
                    language = properties.find(Q(ns, "lang")) if properties is not None else None
                    locale = (
                        language.get(Q(ns, "eastAsia" if target == "eastAsia" else "bidi"), "")
                        if language is not None
                        else ""
                    )
                    language_code, _, region = locale.casefold().partition("-")
                    relevant = (
                        observed_scripts & east_asian_scripts
                        if target == "eastAsia"
                        else observed_scripts & complex_scripts
                    )
                    script = None
                    if relevant:
                        script = scripts.get(language_code)
                        if target == "eastAsia" and language_code == "zh":
                            script = "hant" if region in {"hk", "mo", "tw"} else "hans"
                        candidates = relevant - {"han"}
                        if len(candidates) == 1:
                            script = next(iter(candidates))
                        elif len(candidates) > 1 or target == "eastAsia" and script is None:
                            raise ValueError("assembly_theme_dependency_missing")
                        if script:
                            font = fonts.get(f"{value.casefold()}:{script}")
                        if not font:
                            raise ValueError("assembly_theme_dependency_missing")
                if font:
                    node.set(Q(ns, target), font)
            elif normalized.startswith("theme") and normalized not in {
                "themetint",
                "themeshade",
                "themefilltint",
                "themefillshade",
            }:
                raise ValueError("assembly_theme_dependency_missing")
        for local, (name, _) in attributes.items():
            normalized = local.casefold()
            if not resolve_colors and normalized in {
                "themecolor",
                "themetint",
                "themeshade",
                "themefill",
                "themefilltint",
                "themefillshade",
            }:
                continue
            if normalized.startswith("theme") or normalized.endswith("theme"):
                node.attrib.pop(name, None)


def _field_instruction_nodes(root: ET.Element, ns: str) -> list[tuple[str, list[ET.Element]]]:
    fields: list[tuple[str, list[ET.Element]]] = []
    stack: list[list[ET.Element]] = []
    for node in root.iter():
        if node.tag == Q(ns, "fldSimple"):
            instruction = node.get(Q(ns, "instr"))
            if not instruction:
                raise ValueError("assembly_field_structure_unsupported")
            fields.append((instruction, [node]))
        elif node.tag == Q(ns, "fldChar"):
            kind = node.get(Q(ns, "fldCharType"))
            if kind == "begin":
                stack.append([node])
            elif kind == "separate":
                if not stack:
                    raise ValueError("assembly_field_structure_unsupported")
                stack[-1].append(node)
            elif kind == "end":
                if not stack:
                    raise ValueError("assembly_field_structure_unsupported")
                stack[-1].append(node)
                nodes = stack.pop()
                instruction = "".join(
                    child.text or "" for child in nodes if child.tag == Q(ns, "instrText")
                )
                fields.append((instruction, nodes))
        elif node.tag == Q(ns, "instrText") and stack:
            stack[-1].append(node)
    if stack:
        raise ValueError("assembly_field_structure_unsupported")
    return fields


def _field_target(instruction: str, code: str) -> str | None:
    match = re.match(
        rf"^\s*{code}\s+(?:\"([^\"]+)\"|([A-Za-z_][A-Za-z0-9_]*))",
        instruction,
        re.IGNORECASE,
    )
    return (match.group(1) or match.group(2)) if match is not None else None


def _validate_and_rewrite_fields(
    roots: list[ET.Element],
    ns: str,
    policy: FieldPolicy,
    bookmark_names: Mapping[str, str],
    source_bookmarks: set[str],
    style_names: Mapping[str, str],
) -> None:
    for root in roots:
        for instruction, nodes in _field_instruction_nodes(root, ns):
            match = re.match(r"^\s*([A-Za-z]+)", instruction)
            if match is None:
                raise ValueError("assembly_field_structure_unsupported")
            code = match.group(1).upper()
            replacement: tuple[str, str] | None = None
            if code in {"PAGE", "NUMPAGES", "SECTION", "SECTIONPAGES"}:
                pass
            elif code in {"REF", "PAGEREF"}:
                target = _field_target(instruction, code)
                if target is None or target not in source_bookmarks or target not in bookmark_names:
                    raise ValueError("assembly_field_dependency_missing")
                replacement = (target, bookmark_names[target])
            elif code == "STYLEREF":
                target = _field_target(instruction, code)
                if target is None or target not in style_names:
                    raise ValueError("assembly_field_dependency_missing")
                replacement = (target, style_names[target])
            else:
                raise ValueError("assembly_field_unsupported")
            if replacement is not None:
                old, new = replacement
                simple = next((node for node in nodes if node.tag == Q(ns, "fldSimple")), None)
                if simple is not None:
                    simple.set(
                        Q(ns, "instr"),
                        simple.get(Q(ns, "instr"), "").replace(old, new, 1),
                    )
                else:
                    instructions = [node for node in nodes if node.tag == Q(ns, "instrText")]
                    if not instructions:
                        raise ValueError("assembly_field_structure_unsupported")
                    instructions[0].text = instruction.replace(old, new, 1)
                    for node in instructions[1:]:
                        node.text = ""
            if policy.mark_dirty:
                for node in nodes:
                    if node.tag in {Q(ns, "fldChar"), Q(ns, "fldSimple")}:
                        node.set(Q(ns, "dirty"), "true")


def _conflict_mode(parts: Mapping[str, bytes]) -> bool:
    """Return the package's w14 conflict-mode switch from its related settings part.

    Conflict markup reads differently inside and outside conflict mode, so callers
    compare both packages instead of rewriting the destination's global switch.
    """
    settings_type = f"{_R}/settings"
    if not any(
        relation.get("Type") == settings_type
        for relation in _read_rels(parts, _DOCUMENT).findall(_q(_PR, "Relationship"))
    ):
        return False
    part = _relationship_part(
        parts, _DOCUMENT, settings_type, code="assembly_conflict_mode_unsupported"
    )
    root = ET.fromstring(parts[part])
    switches = list(root.iter(_q(_W14, "conflictMode")))
    if not switches:
        return False
    if len(switches) != 1 or switches[0] not in list(root):
        raise ValueError("assembly_conflict_mode_unsupported")
    value = switches[0].get(_q(_W14, "val"), "true")
    if value in {"1", "true", "on"}:
        return True
    if value in {"0", "false", "off"}:
        return False
    raise ValueError("assembly_conflict_mode_unsupported")


def _request_fields_on_open(destination: dict[str, bytes], ns: str) -> None:
    settings_part = "word/settings.xml"
    if settings_part in destination:
        root = parse_xml_with_preserved_qnames(destination[settings_part])
        if _word_ns(root) != ns or root.tag != Q(ns, "settings"):
            raise ValueError("assembly_destination_settings_invalid")
    else:
        root = ET.Element(Q(ns, "settings"))
        _ensure_override(
            destination,
            settings_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        )
        _add_document_relationship(destination, f"{_R}/settings", settings_part)
    update = root.find(Q(ns, "updateFields"))
    if update is None:
        update = ET.Element(Q(ns, "updateFields"))
        root.insert(0, update)
    update.set(Q(ns, "val"), "true")
    destination[settings_part] = serialize_xml_with_preserved_qnames(root)


def _styles(root: ET.Element, ns: str) -> set[str]:
    names = {"pStyle", "rStyle", "tblStyle", "style", "styleLink", "numStyleLink"}
    return {
        node.get(Q(ns, "val"))
        for node in root.iter()
        if node.tag.rsplit("}", 1)[-1] in names and node.get(Q(ns, "val"))
    }


def _required_styles(
    source_root: ET.Element, selected: list[ET.Element], ns: str
) -> tuple[set[str], dict[str, ET.Element]]:
    requested = set().union(*(_styles(node, ns) for node in selected))
    by_id = {
        node.get(Q(ns, "styleId")): node
        for node in source_root.findall(Q(ns, "style"))
        if node.get(Q(ns, "styleId"))
    }
    required: set[str] = set()
    pending = list(requested)
    while pending:
        style_id = pending.pop()
        if style_id in required:
            continue
        style = by_id.get(style_id)
        if style is None:
            raise ValueError("assembly_style_missing")
        required.add(style_id)
        for dependency in (
            style.findall(f".//{Q(ns, 'basedOn')}")
            + style.findall(f".//{Q(ns, 'next')}")
            + style.findall(f".//{Q(ns, 'link')}")
        ):
            if (value := dependency.get(Q(ns, "val"))) is not None:
                pending.append(value)
    return required, by_id


def _field_style_tokens(roots: list[ET.Element], ns: str) -> set[str]:
    tokens: set[str] = set()
    for root in roots:
        for instruction, _ in _field_instruction_nodes(root, ns):
            match = re.match(r"^\s*STYLEREF\b", instruction, re.IGNORECASE)
            if match is not None:
                target = _field_target(instruction, "STYLEREF")
                if target is None:
                    raise ValueError("assembly_field_dependency_missing")
                tokens.add(target)
    return tokens


def _extend_required_styles(
    required: set[str],
    by_id: Mapping[str, ET.Element],
    tokens: set[str],
    ns: str,
) -> None:
    names = {
        name.get(Q(ns, "val")): style_id
        for style_id, style in by_id.items()
        if (name := style.find(Q(ns, "name"))) is not None and name.get(Q(ns, "val"))
    }
    pending = []
    for token in tokens:
        style_id = token if token in by_id else names.get(token)
        if style_id is None:
            raise ValueError("assembly_field_dependency_missing")
        pending.append(style_id)
    while pending:
        style_id = pending.pop()
        if style_id in required:
            continue
        style = by_id.get(style_id)
        if style is None:
            raise ValueError("assembly_style_missing")
        required.add(style_id)
        pending.extend(
            dependency.get(Q(ns, "val"))
            for dependency in (
                style.findall(f".//{Q(ns, 'basedOn')}")
                + style.findall(f".//{Q(ns, 'next')}")
                + style.findall(f".//{Q(ns, 'link')}")
            )
            if dependency.get(Q(ns, "val")) is not None
        )


def _extend_numbering_style_dependencies(
    source: Mapping[str, bytes],
    roots: list[ET.Element],
    required: set[str],
    by_id: Mapping[str, ET.Element],
    ns: str,
) -> None:
    """Close style/list cross-links before allocating either identity domain."""
    if "word/numbering.xml" not in source:
        return
    numbering = parse_xml_with_preserved_qnames(source["word/numbering.xml"])
    numbers = {item.get(Q(ns, "numId")): item for item in numbering.findall(Q(ns, "num"))}
    abstracts = {
        item.get(Q(ns, "abstractNumId")): item for item in numbering.findall(Q(ns, "abstractNum"))
    }
    seen: set[str] = set()
    while True:
        pending = {
            item.get(Q(ns, "val"))
            for root in [*roots, *(by_id[identifier] for identifier in required)]
            for item in root.iter(Q(ns, "numId"))
            if item.get(Q(ns, "val")) not in {None, "0"}
        } - seen
        if not pending:
            return
        seen.update(pending)
        tokens: set[str] = set()
        for identifier in pending:
            number = numbers.get(identifier)
            if number is None:
                raise ValueError("assembly_num_missing")
            reference = number.find(Q(ns, "abstractNumId"))
            abstract = abstracts.get(reference.get(Q(ns, "val"))) if reference is not None else None
            if abstract is None:
                raise ValueError("assembly_abstract_num_missing")
            tokens.update(_styles(number, ns))
            tokens.update(_styles(abstract, ns))
        _extend_required_styles(required, by_id, tokens, ns)


def _defaults_properties(defaults: ET.Element | None, ns: str) -> list[ET.Element]:
    if defaults is None:
        return []
    return [
        copy.deepcopy(properties)
        for parent, child in (("pPrDefault", "pPr"), ("rPrDefault", "rPr"))
        if (properties := defaults.find(f"{Q(ns, parent)}/{Q(ns, child)}")) is not None
    ]


def _inject_source_defaults(style: ET.Element, defaults: list[ET.Element], ns: str) -> None:
    for properties in defaults:
        if style.find(properties.tag) is None:
            style.append(copy.deepcopy(properties))


def _styles_root(parts: Mapping[str, bytes]) -> ET.Element | None:
    """Return declared style definitions, or ``None`` for a package without any.

    A package with no styles part renders with application defaults. A styles
    relationship without the part is unaddressable and refuses.
    """
    if _STYLES in parts:
        return parse_xml_with_preserved_qnames(parts[_STYLES])
    if any(
        relation.get("Type") == _STYLES_RELATIONSHIP
        for relation in _read_rels(parts, _DOCUMENT).findall(_q(_PR, "Relationship"))
    ):
        raise ValueError("assembly_styles_missing")
    return None


def _definition_key(node: ET.Element, ns: str, kind: str) -> bytes:
    """Compare definition content, not identifiers, display names or whitespace."""
    clone = copy.deepcopy(node)
    for attribute in ("styleId", "numId", "abstractNumId"):
        clone.attrib.pop(Q(ns, attribute), None)
    ignored = {"name", "aliases"} if kind == "style" else {"nsid", "tmpl", "abstractNumId"}
    for child in list(clone):
        if child.tag in {Q(ns, name) for name in ignored}:
            clone.remove(child)
    for child in clone.iter():
        child.tail = None
        if child.text is not None and not child.text.strip():
            child.text = None
        child.attrib = dict(sorted(child.attrib.items()))
    return ET.tostring(clone)


def _theme_resolved(
    parts: Mapping[str, bytes], nodes: list[ET.Element], ns: str, *, resolve_colors: bool = True
) -> dict[int, ET.Element]:
    """Map each node's ``id`` to a copy with the package theme resolved.

    An unresolvable theme yields no copies, so callers compare raw definitions.
    """
    copies = [copy.deepcopy(node) for node in nodes]
    try:
        _resolve_theme_dependencies(parts, copies, ns, resolve_colors=resolve_colors)
    except ValueError:
        return {}
    return {id(node): clone for node, clone in zip(nodes, copies)}


def _equivalent_styles(
    left_id: str,
    left: Mapping[str, ET.Element],
    right_id: str,
    right: Mapping[str, ET.Element],
    ns: str,
    seen: set[tuple[str, str]] | None = None,
) -> bool:
    """Compare the dependency graph, not merely a child style's local XML."""
    if left_id not in left or right_id not in right:
        return False
    seen = set() if seen is None else seen
    pair = (left_id, right_id)
    if pair in seen:
        return True
    seen.add(pair)
    first, second = copy.deepcopy(left[left_id]), copy.deepcopy(right[right_id])
    dependencies = []
    for tag in ("basedOn", "next", "link"):
        first_dependency, second_dependency = (
            first.find(Q(ns, tag)),
            second.find(Q(ns, tag)),
        )
        if (first_dependency is None) != (second_dependency is None):
            return False
        if first_dependency is not None and second_dependency is not None:
            dependencies.append(
                (
                    first_dependency.get(Q(ns, "val")),
                    second_dependency.get(Q(ns, "val")),
                )
            )
            first_dependency.set(Q(ns, "val"), "")
            second_dependency.set(Q(ns, "val"), "")
    return _definition_key(first, ns, "style") == _definition_key(second, ns, "style") and all(
        _equivalent_styles(a, left, b, right, ns, seen) for a, b in dependencies
    )


def _merge_styles(
    destination: dict[str, bytes],
    source: Mapping[str, bytes],
    selected: list[ET.Element],
    ns: str,
    label: str,
    num_remap: Mapping[str, str],
    field_style_tokens: set[str],
    mode: ImportFormatMode = ImportFormatMode.KEEP_SOURCE_FORMATTING,
    options: ImportFormatOptions = _DEFAULT_IMPORT_OPTIONS,
    *,
    allow_style_resources: bool = False,
) -> dict[str, str]:
    source_root = _styles_root(source)
    destination_root = _styles_root(destination)
    destination_style_bytes = (
        ET.tostring(destination_root) if destination_root is not None else None
    )
    required, by_id = _required_styles(
        source_root if source_root is not None else ET.Element(Q(ns, "styles")),
        selected,
        ns,
    )
    _extend_required_styles(required, by_id, field_style_tokens, ns)
    if not allow_style_resources and any(
        attribute in item.attrib
        for identifier in required
        for item in by_id[identifier].iter()
        for attribute in _RELATIONSHIP_ATTRIBUTES
    ):
        raise ValueError("assembly_style_resource_unsupported")
    by_id = {identifier: copy.deepcopy(node) for identifier, node in by_id.items()}
    for identifier in required:
        for number in by_id[identifier].iter(Q(ns, "numId")):
            if (value := number.get(Q(ns, "val"))) in num_remap:
                number.set(Q(ns, "val"), num_remap[value])
    source_defaults = source_root.find(Q(ns, "docDefaults")) if source_root is not None else None
    destination_defaults = (
        destination_root.find(Q(ns, "docDefaults")) if destination_root is not None else None
    )
    # Undeclared source defaults carry no formatting to preserve; declared ones are
    # pinned whenever the destination declares anything else.
    defaults_differ = source_defaults is not None and (
        destination_defaults is None
        or ET.tostring(source_defaults) != ET.tostring(destination_defaults)
    )
    if mode == ImportFormatMode.USE_DESTINATION_STYLES and not options.force_copy_styles:
        defaults_differ = False
    reuse_equivalent = not options.force_copy_styles and (
        mode == ImportFormatMode.KEEP_DIFFERENT_STYLES or options.smart_style_behavior
    )
    # Source definitions arrive with theme references already resolved, so
    # equivalence compares them against theme-resolved destination copies.
    destination_resolved = (
        _theme_resolved(
            destination,
            [
                *(destination_root.findall(Q(ns, "style")) if destination_root is not None else []),
                *([destination_defaults] if destination_defaults is not None else []),
            ],
            ns,
            resolve_colors=options.resolve_theme_colors,
        )
        if reuse_equivalent
        else {}
    )
    defaults_equivalent = not defaults_differ or (
        destination_defaults is not None
        and _definition_key(source_defaults, ns, "defaults")
        == _definition_key(
            destination_resolved.get(id(destination_defaults), destination_defaults),
            ns,
            "defaults",
        )
    )
    source_default_properties = _defaults_properties(source_defaults, ns)
    if not required and not defaults_differ:
        return {}
    if defaults_differ:
        if "Normal" not in by_id:
            raise ValueError("assembly_source_normal_missing")
        required.add("Normal")
    if destination_root is None:
        destination_root = ET.Element(Q(ns, "styles"))
        _ensure_override(destination, _STYLES, _STYLES_CONTENT_TYPE)
        _add_document_relationship(destination, _STYLES_RELATIONSHIP, _STYLES)
    existing = {
        node.get(Q(ns, "styleId")): ET.tostring(node)
        for node in destination_root.findall(Q(ns, "style"))
    }
    occupied_names = {
        name.get(Q(ns, "val"), "").casefold()
        for node in destination_root.findall(Q(ns, "style"))
        if (name := node.find(Q(ns, "name"))) is not None
    }
    remap: dict[str, str] = {}
    occupied = set(existing)
    destination_styles = {
        node.get(Q(ns, "styleId")): node for node in destination_root.findall(Q(ns, "style"))
    }
    resolved_styles = {
        identifier: destination_resolved.get(id(node), node)
        for identifier, node in destination_styles.items()
    }
    by_name = {
        (node.get(Q(ns, "type")), name.get(Q(ns, "val"), "").casefold()): identifier
        for identifier, node in destination_styles.items()
        if (name := node.find(Q(ns, "name"))) is not None
    }
    for style_id in sorted(required):
        source_style = by_id[style_id]
        has_resources = any(
            attribute in item.attrib
            for item in source_style.iter()
            for attribute in _RELATIONSHIP_ATTRIBUTES
        )
        name = source_style.find(Q(ns, "name"))
        matched = by_name.get(
            (
                source_style.get(Q(ns, "type")),
                name.get(Q(ns, "val"), "").casefold() if name is not None else "",
            )
        )
        if (
            matched is None
            and style_id in destination_styles
            and source_style.get(Q(ns, "type")) == destination_styles[style_id].get(Q(ns, "type"))
        ):
            matched = style_id
        if (
            mode == ImportFormatMode.USE_DESTINATION_STYLES
            and matched is not None
            and not options.force_copy_styles
        ):
            remap[style_id] = matched
            continue
        if reuse_equivalent and defaults_equivalent and not has_resources:
            equivalent = next(
                (
                    identifier
                    for identifier in destination_styles
                    if _equivalent_styles(style_id, by_id, identifier, resolved_styles, ns)
                    and (options.smart_style_behavior or identifier == matched)
                ),
                None,
            )
            if equivalent is not None:
                remap[style_id] = equivalent
                continue
        if (
            not options.force_copy_styles
            and not defaults_differ
            and (
                style_id not in existing
                or (
                    not has_resources
                    and ET.tostring(by_id[style_id]) == existing[style_id]
                    and _equivalent_styles(style_id, by_id, style_id, destination_styles, ns)
                )
            )
        ):
            remap[style_id] = style_id
            continue
        if not options.force_copy_styles and defaults_differ and style_id not in existing:
            remap[style_id] = style_id
            occupied.add(style_id)
            continue
        base = f"assembly-{label}-{style_id}"
        candidate = base
        serial = 2
        while candidate in occupied:
            candidate = f"{base}-{serial}"
            serial += 1
        remap[style_id] = candidate
        occupied.add(candidate)
    style_importer = (
        _NodeResourceImporter(destination, source, label, Path("."))
        if allow_style_resources
        else None
    )
    for style_id in sorted(required):
        destination_id = remap[style_id]
        if destination_id in existing:
            continue
        clone = copy.deepcopy(by_id[style_id])
        clone.set(Q(ns, "styleId"), destination_id)
        clone.attrib.pop(Q(ns, "default"), None)
        name = clone.find(Q(ns, "name"))
        if name is not None:
            display_name = name.get(Q(ns, "val"), "")
            if display_name.casefold() in occupied_names:
                display_name = destination_id
                serial = 2
                while display_name.casefold() in occupied_names:
                    display_name = f"{destination_id}-{serial}"
                    serial += 1
                name.set(Q(ns, "val"), display_name)
            occupied_names.add(display_name.casefold())
        for dependency in (
            clone.findall(f".//{Q(ns, 'basedOn')}")
            + clone.findall(f".//{Q(ns, 'next')}")
            + clone.findall(f".//{Q(ns, 'link')}")
        ):
            if (value := dependency.get(Q(ns, "val"))) in remap:
                dependency.set(Q(ns, "val"), remap[value])
        if defaults_differ and clone.find(Q(ns, "basedOn")) is None:
            _inject_source_defaults(clone, source_default_properties, ns)
        if style_importer is not None:
            style_importer.copy_references([clone], _STYLES, _STYLES)
        destination_root.append(clone)
        existing[destination_id] = ET.tostring(clone)
    if defaults_differ:
        for block in selected:
            for paragraph in block.iter(Q(ns, "p")):
                if paragraph.find(f"{Q(ns, 'pPr')}/{Q(ns, 'pStyle')}") is not None:
                    continue
                properties = paragraph.find(Q(ns, "pPr"))
                if properties is None:
                    properties = ET.Element(Q(ns, "pPr"))
                    paragraph.insert(0, properties)
                ET.SubElement(properties, Q(ns, "pStyle"), {Q(ns, "val"): remap["Normal"]})
    for block in selected:
        for node in block.iter():
            if (
                node.tag.rsplit("}", 1)[-1]
                in {"pStyle", "rStyle", "tblStyle", "styleLink", "numStyleLink"}
                and (value := node.get(Q(ns, "val"))) in remap
            ):
                node.set(Q(ns, "val"), remap[value])
    if destination_style_bytes != ET.tostring(destination_root):
        destination[_STYLES] = serialize_xml_with_preserved_qnames(destination_root)
    if style_importer is not None and style_importer.rel_map:
        style_importer.finish()
    return remap


def _field_style_names(
    source: Mapping[str, bytes],
    destination: Mapping[str, bytes],
    ns: str,
    remap: Mapping[str, str],
) -> dict[str, str]:
    if not remap:
        return {}
    source_root = ET.fromstring(source[_STYLES])
    destination_root = ET.fromstring(destination[_STYLES])
    source_styles = {
        style.get(Q(ns, "styleId")): style
        for style in source_root.findall(Q(ns, "style"))
        if style.get(Q(ns, "styleId"))
    }
    destination_styles = {
        style.get(Q(ns, "styleId")): style
        for style in destination_root.findall(Q(ns, "style"))
        if style.get(Q(ns, "styleId"))
    }
    names: dict[str, str] = {}
    for source_id, destination_id in remap.items():
        destination_style = destination_styles.get(destination_id)
        source_style = source_styles.get(source_id)
        if destination_style is None or source_style is None:
            continue
        destination_name = destination_style.find(Q(ns, "name"))
        source_name = source_style.find(Q(ns, "name"))
        if destination_name is None:
            continue
        target = destination_name.get(Q(ns, "val"))
        if target is None:
            continue
        names[source_id] = target
        if source_name is not None and (source_display := source_name.get(Q(ns, "val"))):
            names[source_display] = target
    return names


def _ensure_numbering(destination: dict[str, bytes], source: Mapping[str, bytes], ns: str) -> None:
    if "word/numbering.xml" in destination:
        return
    if "word/numbering.xml" not in source:
        raise ValueError("assembly_numbering_missing")
    destination["word/numbering.xml"] = serialize_xml_with_preserved_qnames(
        ET.Element(Q(ns, "numbering"))
    )
    types = parse_xml_with_preserved_qnames(destination["[Content_Types].xml"])
    if not any(
        node.get("PartName") == "/word/numbering.xml" for node in types.findall(_q(_CT, "Override"))
    ):
        types.append(
            ET.Element(
                _q(_CT, "Override"),
                {
                    "PartName": "/word/numbering.xml",
                    "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                },
            )
        )
    destination["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types)
    rels = _read_rels(destination, _DOCUMENT)
    if not any(
        relation.get("Type")
        == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
        for relation in rels.findall(_q(_PR, "Relationship"))
    ):
        rels.append(
            ET.Element(
                _q(_PR, "Relationship"),
                {
                    "Id": _new_rel_id(rels),
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                    "Target": "numbering.xml",
                },
            )
        )
    destination[_rels_name(_DOCUMENT)] = serialize_xml_with_preserved_qnames(rels)


def _merge_numbering(
    destination: dict[str, bytes],
    source: Mapping[str, bytes],
    selected: list[ET.Element],
    ns: str,
    field_style_tokens: set[str],
    options: ImportFormatOptions = _DEFAULT_IMPORT_OPTIONS,
    *,
    reuse_numbers: Mapping[str, str] | None = None,
) -> dict[str, str]:
    source_styles = _styles_root(source)
    if source_styles is None:
        source_styles = ET.Element(Q(ns, "styles"))
    required, source_styles_by_id = _required_styles(source_styles, selected, ns)
    _extend_required_styles(required, source_styles_by_id, field_style_tokens, ns)
    _extend_numbering_style_dependencies(source, selected, required, source_styles_by_id, ns)
    used = {
        node.get(Q(ns, "val"))
        for block in selected
        for node in block.iter(Q(ns, "numId"))
        if node.get(Q(ns, "val"))
    } | {
        node.get(Q(ns, "val"))
        for style_id in required
        for node in source_styles_by_id[style_id].iter(Q(ns, "numId"))
        if node.get(Q(ns, "val"))
    }
    used.discard("0")  # ECMA-376: zero suppresses numbering, it is not a list definition.
    if not used:
        return {}
    _ensure_numbering(destination, source, ns)
    source_root = parse_xml_with_preserved_qnames(source["word/numbering.xml"])
    destination_root = parse_xml_with_preserved_qnames(destination["word/numbering.xml"])
    source_nums = {node.get(Q(ns, "numId")): node for node in source_root.findall(Q(ns, "num"))}
    source_abstract = {
        node.get(Q(ns, "abstractNumId")): node for node in source_root.findall(Q(ns, "abstractNum"))
    }
    next_num = (
        max(
            [
                int(value)
                for value in (
                    node.get(Q(ns, "numId")) for node in destination_root.findall(Q(ns, "num"))
                )
                if value and value.isdigit()
            ]
            + [0]
        )
        + 1
    )
    next_abstract = (
        max(
            [
                int(value)
                for value in (
                    node.get(Q(ns, "abstractNumId"))
                    for node in destination_root.findall(Q(ns, "abstractNum"))
                )
                if value and value.isdigit()
            ]
            + [0]
        )
        + 1
    )
    remap: dict[str, str] = {}
    destination_abstract = {
        node.get(Q(ns, "abstractNumId")): node
        for node in destination_root.findall(Q(ns, "abstractNum"))
    }
    bullet_map: dict[str, str] = {}
    source_bullets = {
        node.get(Q(ns, "numPicBulletId")): node
        for node in source_root.findall(Q(ns, "numPicBullet"))
    }
    bullet_importer: _Importer | None = None
    for old in sorted(used, key=lambda value: int(value or "-1")):
        num = source_nums.get(old)
        if num is None:
            raise ValueError("assembly_num_missing")
        if reuse_numbers is not None and old in reuse_numbers:
            remap[old] = reuse_numbers[old]
            continue
        abstract_node = num.find(Q(ns, "abstractNumId"))
        abstract = source_abstract.get(
            abstract_node.get(Q(ns, "val")) if abstract_node is not None else None
        )
        if abstract is None:
            raise ValueError("assembly_abstract_num_missing")
        if (not options.keep_source_numbering or options.merge_pasted_lists) and not any(
            next(item.iter(Q(ns, "lvlPicBulletId")), None) is not None for item in (abstract, num)
        ):
            for candidate in destination_root.findall(Q(ns, "num")):
                reference = candidate.find(Q(ns, "abstractNumId"))
                candidate_abstract = destination_abstract.get(
                    reference.get(Q(ns, "val")) if reference is not None else None
                )
                if (
                    candidate_abstract is not None
                    and _definition_key(candidate_abstract, ns, "numbering")
                    == _definition_key(abstract, ns, "numbering")
                    and _definition_key(candidate, ns, "numbering")
                    == _definition_key(num, ns, "numbering")
                ):
                    remap[old] = candidate.get(Q(ns, "numId"))
                    break
            if old in remap:
                continue
        new_abstract, new_num = str(next_abstract), str(next_num)
        next_abstract += 1
        next_num += 1
        abstract_copy, num_copy = copy.deepcopy(abstract), copy.deepcopy(num)
        abstract_copy.set(Q(ns, "abstractNumId"), new_abstract)
        num_copy.set(Q(ns, "numId"), new_num)
        num_copy.find(Q(ns, "abstractNumId")).set(Q(ns, "val"), new_abstract)  # type: ignore[union-attr]
        for reference in [
            *abstract_copy.iter(Q(ns, "lvlPicBulletId")),
            *num_copy.iter(Q(ns, "lvlPicBulletId")),
        ]:
            identifier = reference.get(Q(ns, "val"))
            if identifier not in bullet_map:
                bullet = source_bullets.get(identifier)
                if bullet is None:
                    raise ValueError("assembly_picture_bullet_missing")
                imported_bullet = copy.deepcopy(bullet)
                if any(item.tag == _DIAGRAM_REL_IDS for item in imported_bullet.iter()):
                    raise ValueError("assembly_picture_bullet_resource_unsupported")
                if bullet_importer is None:
                    bullet_importer = _Importer(destination, source, "numbering", Path("."))
                for item in imported_bullet.iter():
                    for attribute in _RELATIONSHIP_ATTRIBUTES:
                        rid = item.get(attribute)
                        if rid is None:
                            continue
                        relation = next(
                            (
                                row
                                for row in _read_rels(source, "word/numbering.xml")
                                if row.get("Id") == rid
                            ),
                            None,
                        )
                        if (
                            relation is None
                            or relation.get("Type") != f"{_R}/image"
                            or relation.get("TargetMode") == "External"
                        ):
                            raise ValueError("assembly_picture_bullet_resource_unsupported")
                        target = _resolve("word/numbering.xml", relation.get("Target", ""))
                        bullet_type = _content_type(target, bullet_importer.source_types)
                        if bullet_type is None or not bullet_type[1].startswith("image/"):
                            raise ValueError("assembly_picture_bullet_resource_unsupported")
                occupied = {
                    int(value)
                    for item in destination_root.findall(Q(ns, "numPicBullet"))
                    if (value := item.get(Q(ns, "numPicBulletId"), "")).isdigit()
                }
                new_id = str(max(occupied, default=-1) + 1)
                imported_bullet.set(Q(ns, "numPicBulletId"), new_id)
                _remap_vml_ids([imported_bullet], ns, _vml_ids(destination_root))
                bullet_importer.copy_references(
                    [imported_bullet], "word/numbering.xml", "word/numbering.xml"
                )
                destination_root.insert(0, imported_bullet)
                bullet_map[identifier] = new_id
            reference.set(Q(ns, "val"), bullet_map[identifier])
        destination_root.extend((abstract_copy, num_copy))
        destination_abstract[new_abstract] = abstract_copy
        remap[old] = new_num
    for block in selected:
        for node in block.iter(Q(ns, "numId")):
            if (value := node.get(Q(ns, "val"))) in remap:
                node.set(Q(ns, "val"), remap[value])
    destination["word/numbering.xml"] = serialize_xml_with_preserved_qnames(destination_root)
    if bullet_importer is not None:
        bullet_importer.finish()
    return remap


def _effective_section_references(
    root: ET.Element, ns: str
) -> dict[int, dict[tuple[str, str], ET.Element]]:
    active: dict[tuple[str, str], ET.Element] = {}
    resolved: dict[int, dict[tuple[str, str], ET.Element]] = {}
    for section in _section_nodes(root, ns):
        direct: dict[tuple[str, str], ET.Element] = {}
        for reference in section:
            local = reference.tag.rsplit("}", 1)[-1]
            if local not in {"headerReference", "footerReference"}:
                continue
            kind = reference.get(Q(ns, "type"))
            relationship = reference.get(Q(_R, "id")) or reference.get(Q(_STRICT_R, "id"))
            if kind is None or relationship is None or (local, kind) in direct:
                raise ValueError("assembly_section_header_conflict")
            direct[(local, kind)] = reference
        active.update(direct)
        resolved[id(section)] = dict(active)
    return resolved


def _apply_section_policy(
    selected_sections: Mapping[ET.Element, ET.Element],
    source_root: ET.Element,
    ns: str,
    policy: SectionPolicy,
) -> None:
    resolved = _effective_section_references(source_root, ns)
    for section, source_section in selected_sections.items():
        if policy.header_footer_inheritance == "link_previous":
            for reference in list(section):
                if reference.tag in {
                    Q(ns, "headerReference"),
                    Q(ns, "footerReference"),
                }:
                    section.remove(reference)
        else:
            direct = {
                (reference.tag.rsplit("}", 1)[-1], reference.get(Q(ns, "type")))
                for reference in section
                if reference.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
            }
            if id(source_section) not in resolved:
                raise ValueError("assembly_section_header_conflict")
            for key, reference in resolved[id(source_section)].items():
                if key not in direct:
                    section.append(copy.deepcopy(reference))
        if policy.page_numbering == "continue":
            number = section.find(Q(ns, "pgNumType"))
            if number is not None:
                section.remove(number)


def _gather_related_roots(
    source: Mapping[str, bytes],
    selected: list[ET.Element],
    ns: str,
    *,
    node_import: bool = False,
) -> dict[str, ET.Element]:
    """Return selected section header/footer roots after validating their contract."""
    relationships = _read_rels(source, _DOCUMENT)
    source_types = ET.fromstring(source["[Content_Types].xml"])
    related: dict[str, ET.Element] = {}
    for block in selected:
        for section in block.iter(Q(ns, "sectPr")):
            for reference in section:
                reference_name = reference.tag.rsplit("}", 1)[-1]
                if reference_name not in {"headerReference", "footerReference"}:
                    continue
                section_type = reference.get(Q(ns, "type"))
                relationship_id = reference.get(Q(_R, "id"))
                part_kind = {
                    "headerReference": "header",
                    "footerReference": "footer",
                }[reference_name]
                relationship = next(
                    (
                        candidate
                        for candidate in relationships.findall(_q(_PR, "Relationship"))
                        if candidate.get("Id") == relationship_id
                    ),
                    None,
                )
                if (
                    section_type not in {"default", "even", "first"}
                    or relationship_id is None
                    or relationship is None
                    or relationship.get("Type") != f"{_R}/{part_kind}"
                    or relationship.get("TargetMode") is not None
                    or (target := relationship.get("Target")) is None
                ):
                    raise ValueError("assembly_related_structure_unsupported")
                part = _resolve(_DOCUMENT, target)
                expected_content_type = (
                    "application/vnd.openxmlformats-officedocument."
                    f"wordprocessingml.{part_kind}+xml"
                )
                if (
                    not part.startswith("word/")
                    or part not in source
                    or _content_type(part, source_types) != ("Override", expected_content_type)
                ):
                    raise ValueError("assembly_related_structure_unsupported")
                root = parse_xml_with_preserved_qnames(source[part])
                if (
                    _word_ns(root) != ns
                    or root.tag != Q(ns, "hdr" if part_kind == "header" else "ftr")
                    or (not node_import and _unsupported([root], ns))
                    or (not node_import and _comment_markers([root], ns))
                ):
                    raise ValueError("assembly_related_structure_unsupported")
                related.setdefault(part, root)
    return related


def _section_owner_index(body: ET.Element, section: ET.Element, ns: str) -> int:
    if section is body.find(Q(ns, "sectPr")):
        return len(body) - 1
    for index, item in enumerate(body):
        if section is item.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}"):
            return index
    raise ValueError("assembly_section_header_conflict")


def _following_destination_section(
    destination_root: ET.Element, ns: str, insertion: int
) -> ET.Element:
    body = destination_root.find(Q(ns, "body"))
    if body is None:
        raise ValueError("assembly_destination_invalid")
    following = next(
        (
            section
            for section in _section_nodes(destination_root, ns)
            if _section_owner_index(body, section, ns) >= insertion
        ),
        None,
    )
    if following is None:
        raise ValueError("assembly_section_header_conflict")
    return following


def _section_reference_keys(section: ET.Element, ns: str) -> set[tuple[str, str]]:
    return {
        (reference.tag.rsplit("}", 1)[-1], section_type)
        for reference in section
        if reference.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
        and (section_type := reference.get(Q(ns, "type"))) is not None
    }


def _empty_destination_section_references(
    destination: dict[str, bytes],
    ns: str,
    references: set[tuple[str, str]],
) -> list[ET.Element]:
    if not references:
        return []
    types = parse_xml_with_preserved_qnames(destination["[Content_Types].xml"])
    rels_name = _rels_name(_DOCUMENT)
    rels = _read_rels(destination, _DOCUMENT)
    result: list[ET.Element] = []
    for reference_name, section_type in sorted(references):
        part_kind = {
            "headerReference": "header",
            "footerReference": "footer",
        }[reference_name]
        part = f"word/assembly-empty-{part_kind}-{section_type}.xml"
        serial = 2
        while part in destination:
            part = f"word/assembly-empty-{part_kind}-{section_type}-{serial}.xml"
            serial += 1
        relationship_id = _new_rel_id(rels)
        rels.append(
            ET.Element(
                _q(_PR, "Relationship"),
                {
                    "Id": relationship_id,
                    "Type": f"{_R}/{part_kind}",
                    "Target": _relative(_DOCUMENT, part),
                },
            )
        )
        if not any(
            node.get("PartName") == "/" + part for node in types.findall(_q(_CT, "Override"))
        ):
            types.append(
                ET.Element(
                    _q(_CT, "Override"),
                    {
                        "PartName": "/" + part,
                        "ContentType": (
                            "application/vnd.openxmlformats-officedocument."
                            f"wordprocessingml.{part_kind}+xml"
                        ),
                    },
                )
            )
        root = ET.Element(Q(ns, "hdr" if part_kind == "header" else "ftr"))
        ET.SubElement(root, Q(ns, "p"))
        destination[part] = serialize_xml_with_preserved_qnames(root)
        result.append(
            ET.Element(
                Q(ns, reference_name),
                {
                    Q(ns, "type"): section_type,
                    Q(_R, "id"): relationship_id,
                },
            )
        )
    destination[rels_name] = serialize_xml_with_preserved_qnames(rels)
    destination["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types)
    return result


def _pin_destination_following_section(
    destination_root: ET.Element,
    destination: dict[str, bytes],
    ns: str,
    insertion: int,
    source_references: set[tuple[str, str]],
) -> None:
    following = _following_destination_section(destination_root, ns, insertion)
    resolved = _effective_section_references(destination_root, ns)
    direct = _section_reference_keys(following, ns)
    effective = resolved[id(following)]
    for key, reference in effective.items():
        if key not in direct:
            following.append(copy.deepcopy(reference))
    following.extend(
        _empty_destination_section_references(
            destination, ns, source_references - direct - set(effective)
        )
    )


def _pin_destination_preceding_section(
    destination_root: ET.Element, ns: str, insertion: int
) -> int:
    if insertion == 0:
        return 0
    body = destination_root.find(Q(ns, "body"))
    if body is None:
        raise ValueError("assembly_destination_invalid")
    preceding = body[insertion - 1]
    properties = preceding.find(Q(ns, "pPr"))
    if properties is not None and properties.find(Q(ns, "sectPr")) is not None:
        return 0
    section = copy.deepcopy(_following_destination_section(destination_root, ns, insertion))
    if preceding.tag == Q(ns, "p"):
        if properties is None:
            properties = ET.Element(Q(ns, "pPr"))
            preceding.insert(0, properties)
        properties.append(section)
        return 0
    boundary = ET.Element(Q(ns, "p"))
    properties = ET.SubElement(boundary, Q(ns, "pPr"))
    properties.append(section)
    body.insert(insertion, boundary)
    return 1


def _unselected_destination_parts_preserved(
    baseline: Mapping[str, bytes], candidate: Path, changed: set[str]
) -> bool:
    with zipfile.ZipFile(candidate) as archive:
        return all(
            archive.read(name) == data for name, data in baseline.items() if name not in changed
        )


def _remap_bookmarks(
    destination: Mapping[str, bytes],
    roots: list[ET.Element],
    ns: str,
    *,
    occupied: Sequence[ET.Element] = (),
) -> dict[str, str]:
    # ``occupied`` carries live trees whose bookmarks ``destination`` bytes lag.
    destination_roots: list[ET.Element] = list(occupied)
    for part, data in destination.items():
        if not part.startswith("word/") or not part.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        if _word_ns(root) == ns:
            destination_roots.append(root)
    used_ids = {
        int(node.get(Q(ns, "id")))
        for root in destination_roots
        for node in root.iter(Q(ns, "bookmarkStart"))
        if (node.get(Q(ns, "id")) or "").isdigit()
    }
    used_names = {
        node.get(Q(ns, "name"))
        for root in destination_roots
        for node in root.iter(Q(ns, "bookmarkStart"))
        if node.get(Q(ns, "name"))
    }
    next_id = max(used_ids, default=0) + 1
    names: dict[str, str] = {}
    ids: dict[str, str] = {}
    for root in roots:
        for node in root.iter(Q(ns, "bookmarkStart")):
            old_id, old_name = node.get(Q(ns, "id")), node.get(Q(ns, "name"))
            if old_id is None or old_name is None or old_id in ids or old_name in names:
                raise ValueError("assembly_bookmark_invalid")
            new_id = str(next_id)
            next_id += 1
            base = old_name[:32] or "Bookmark"
            new_name = base
            serial = 2
            while new_name in used_names:
                new_name = f"{base}_{serial}"
                serial += 1
            used_names.add(new_name)
            ids[old_id] = new_id
            names[old_name] = new_name
            node.set(Q(ns, "id"), new_id)
            node.set(Q(ns, "name"), new_name)
    for root in roots:
        for node in root.iter(Q(ns, "bookmarkEnd")):
            old = node.get(Q(ns, "id"))
            if old not in ids:
                raise ValueError("assembly_bookmark_end_invalid")
            node.set(Q(ns, "id"), ids[old])
        for node in root.iter(Q(ns, "hyperlink")):
            anchor = node.get(Q(ns, "anchor"))
            if anchor in names:
                node.set(Q(ns, "anchor"), names[anchor])
    return names


def _remap_conflict_ids(
    roots: list[ET.Element], ns: str, occupied: set[int], *, allow_revisions: bool = False
) -> None:
    """Rebase imported conflicts in one part's shared revision identity domain."""
    from .review import _CONFLICT_REVISION_TAGS, _CUSTOM_XML_CONFLICT_TAGS, _part_revision_index

    if allow_revisions:
        from .review import _SELECTABLE_REVISION_TAGS

        tags = (
            _CONFLICT_REVISION_TAGS
            | _CUSTOM_XML_CONFLICT_TAGS
            | {Q(ns, name) for name in _SELECTABLE_REVISION_TAGS}
        )
        nodes = [node for root in roots for node in root.iter() if node.tag in tags]
        identifiers = {node.get(Q(ns, "id"), "") for node in nodes}
        if any(not value.isascii() or not value.isdigit() for value in identifiers):
            raise ValueError("node_import_revision_id_invalid")
        serial = max(occupied | {int(value) for value in identifiers}, default=0) + 1
        remap = {}
        for value in sorted(identifiers, key=int):
            if int(value) in occupied:
                remap[value] = str(serial)
                serial += 1
            else:
                remap[value] = value
        for node in nodes:
            node.set(Q(ns, "id"), remap[node.get(Q(ns, "id"))])
        occupied.update(int(value) for value in remap.values())
        return
    next_id = 1
    for root in roots:
        revisions, _, _ = _part_revision_index(root, ns)
        for revision in revisions:
            if revision.node.tag not in _CONFLICT_REVISION_TAGS:
                raise ValueError("assembly_structure_unsupported")
            if not revision.identifier.isdigit():
                raise ValueError("assembly_revision_id_invalid")
            identifier = int(revision.identifier)
            if identifier in occupied:
                while next_id in occupied:
                    next_id += 1
                identifier = next_id
                next_id += 1
                revision.node.set(Q(ns, "id"), str(identifier))
            occupied.add(identifier)
    # A custom XML conflict range is one Start/End pair sharing a kind and a w:id; the
    # pair is rebased together inside the same id domain as the revisions above.
    ranges: dict[tuple[str, str], list[ET.Element]] = {}
    for root in roots:
        for node in root.iter():
            if node.tag in _CUSTOM_XML_CONFLICT_TAGS:
                value = node.get(Q(ns, "id"), "")
                if not value.isascii() or not value.isdigit():
                    raise ValueError("assembly_revision_id_invalid")
                kind = "Ins" if "InsRange" in node.tag else "Del"
                ranges.setdefault((kind, value), []).append(node)
    for (_, value), members in ranges.items():
        identifier = int(value)
        if identifier in occupied:
            while next_id in occupied:
                next_id += 1
            identifier = next_id
            next_id += 1
            for node in members:
                node.set(Q(ns, "id"), str(identifier))
        occupied.add(identifier)


def _remap_sdt_ids(roots: list[ET.Element], ns: str, occupied: set[int]) -> None:
    content_controls = feature_module("content_controls", required=False)

    for root in roots:
        for control in root.iter(Q(ns, "sdt")):
            identifiers = control.findall(f"{Q(ns, 'sdtPr')}/{Q(ns, 'id')}")
            if len(identifiers) != 1:
                raise ValueError("assembly_sdt_identity_unsupported")
            node = identifiers[0]
            value = node.get(Q(ns, "val"), "")
            if not value.isdigit() or int(value) > 2_147_483_647:
                raise ValueError("assembly_sdt_identity_unsupported")
            identifier = int(value)
            if identifier in occupied:
                if content_controls is None:
                    content_controls = feature_module("content_controls")
                identifier = content_controls._allocate_sdt_id(occupied)
                if identifier is None:
                    raise ValueError("assembly_sdt_identity_exhausted")
                node.set(Q(ns, "val"), str(identifier))
            occupied.add(identifier)


def _remap_paragraph_ids(roots: list[ET.Element], ns: str, occupied: set[str]) -> None:
    """Reserve identities within one destination part, changing only collisions."""
    for root in roots:
        for paragraph in root.iter(Q(ns, "p")):
            identifier = paragraph.get(_PARA_ID)
            if identifier is None:
                continue
            if identifier.upper() in occupied:
                paragraph.set(_PARA_ID, generate_collision_safe_para_id(occupied, identifier))
            else:
                occupied.add(identifier.upper())


def _drawing_identifier(node: ET.Element) -> int | None:
    value = node.get("id")
    if value is None:
        return None
    if not re.fullmatch(r"\+?[0-9]+", value.strip()):
        raise ValueError("assembly_drawing_identity_invalid")
    identifier = int(value)
    if identifier > _MAX_DRAWING_ID:
        raise ValueError("assembly_drawing_identity_invalid")
    return identifier


def _drawing_ids(root: ET.Element) -> dict[str, set[int]]:
    occupied: dict[str, set[int]] = {"docPr": set(), "cNvPr": set()}
    for node in root.iter():
        domain = _DRAWING_ID_DOMAINS.get(node.tag)
        if domain is not None and (identifier := _drawing_identifier(node)) is not None:
            occupied[domain].add(identifier)
    return occupied


def _remap_drawing_ids(roots: list[ET.Element], occupied: dict[str, set[int]]) -> None:
    """Allocate only imported collisions within one destination package part.

    Reserve all incoming values before allocation so a collision cannot steal a
    later noncolliding identity. Resolve connector endpoints in their original
    drawing graph before changing anything; repeated imported canvases may have
    identical source IDs without sharing their endpoint bindings.

    Related parts copied by _Importer retain their own identity scope and bytes.
    Only main-document blocks and notes merged into existing parts need rebasing.
    """
    definitions: list[tuple[ET.Element, ET.Element, str, int]] = []
    connections: list[tuple[ET.Element, ET.Element, int]] = []
    local: dict[tuple[ET.Element, str, int], list[ET.Element]] = {}
    imported: dict[tuple[str, int], list[ET.Element]] = {}
    reserved = {domain: set(values) for domain, values in occupied.items()}
    for root in roots:
        pending = [(root, root)]
        while pending:
            node, scope = pending.pop()
            if node.tag in _WORD_DRAWINGS:
                scope = node
            domain = _DRAWING_ID_DOMAINS.get(node.tag)
            if domain is not None:
                identifier = _drawing_identifier(node)
                if identifier is not None:
                    definitions.append((node, scope, domain, identifier))
                    local.setdefault((scope, domain, identifier), []).append(node)
                    imported.setdefault((domain, identifier), []).append(node)
                    reserved[domain].add(identifier)
            elif node.tag in _DRAWING_CONNECTIONS:
                identifier = _drawing_identifier(node)
                if identifier is not None:
                    connections.append((node, scope, identifier))
            pending.extend((child, scope) for child in reversed(node))

    bindings: list[tuple[ET.Element, ET.Element, int]] = []
    for connection, scope, identifier in connections:
        # Inner canvas objects take precedence over their outer Word wrapper.
        targets = (
            local.get((scope, "cNvPr", identifier))
            or local.get((scope, "docPr", identifier))
            or imported.get(("cNvPr", identifier))
            or imported.get(("docPr", identifier))
        )
        if not targets:
            raise ValueError("assembly_drawing_reference_missing")
        if len(targets) != 1:
            raise ValueError("assembly_drawing_reference_ambiguous")
        bindings.append((connection, targets[0], identifier))

    next_ids = {domain: 1 for domain in occupied}
    for node, _, domain, identifier in definitions:
        if identifier in occupied[domain]:
            candidate = next_ids[domain]
            while candidate in reserved[domain]:
                candidate += 1
            if candidate > _MAX_DRAWING_ID:
                raise ValueError("assembly_drawing_identity_exhausted")
            node.set("id", str(candidate))
            reserved[domain].add(candidate)
            next_ids[domain] = candidate + 1
            occupied[domain].add(candidate)
        else:
            occupied[domain].add(identifier)
    for connection, target, identifier in bindings:
        if int(target.attrib["id"]) != identifier:
            connection.set("id", target.attrib["id"])


def _vml_identity_attributes(node: ET.Element) -> tuple[str, ...]:
    if not node.tag.startswith(f"{{{_VML}}}"):
        return ()
    return tuple(name for name in ("id", _VML_SPID) if node.get(name))


def _vml_ids(root: ET.Element) -> set[str]:
    return {node.attrib[name] for node in root.iter() for name in _vml_identity_attributes(node)}


def _fresh_vml_id(value: str, reserved: set[str]) -> str:
    serial_id = re.fullmatch(r"(_x0000_[si])([0-9]+)", value)
    if serial_id is not None:
        # Word derives shape numbers from these spellings; keep them numeric.
        prefix, serial = serial_id.group(1), int(serial_id.group(2))
        while f"{prefix}{serial}" in reserved:
            serial += 1
        return f"{prefix}{serial}"
    serial = 2
    while f"{value}_{serial}" in reserved:
        serial += 1
    return f"{value}_{serial}"


def _remap_vml_ids(roots: list[ET.Element], ns: str, occupied: set[str]) -> None:
    """Rename imported VML shape identities that collide within one destination part.

    VML ``id`` and ``o:spid`` values share one lexical domain. A renamed value
    changes at every imported definition and reference, so an id equal to its spid,
    shapetype uses, connector rules and OLE or ActiveX shape links stay bound.
    """
    incoming = {
        node.attrib[name]
        for root in roots
        for node in root.iter()
        for name in _vml_identity_attributes(node)
    }
    reserved = occupied | incoming
    renamed: dict[str, str] = {}
    for root in roots:
        for node in root.iter():
            for name in _vml_identity_attributes(node):
                value = node.attrib[name]
                if value in occupied and value not in renamed:
                    renamed[value] = _fresh_vml_id(value, reserved)
                    reserved.add(renamed[value])
    occupied.update(incoming - renamed.keys(), renamed.values())
    if not renamed:
        return
    for root in roots:
        for node in root.iter():
            for name in _vml_identity_attributes(node):
                node.set(name, renamed.get(node.attrib[name], node.attrib[name]))
            if node.tag.startswith(f"{{{_VML}}}"):
                references = ("type",)
            elif node.tag == _q(_VML_OFFICE, "OLEObject"):
                references = ("ShapeID",)
            elif node.tag.startswith(f"{{{_VML_OFFICE}}}"):
                references = ("idref",)
            elif node.tag == Q(ns, "control"):
                references = (Q(ns, "shapeid"),)
            elif node.tag.endswith("}compatExt") and "spid" in node.attrib:
                references = ("spid",)
            else:
                continue
            for name in references:
                value = node.get(name)
                if value is None:
                    continue
                fragment = value.startswith("#")
                if fragment != (name in {"type", "idref"}):
                    continue
                target = renamed.get(value.removeprefix("#"))
                if target is not None:
                    node.set(name, f"#{target}" if fragment else target)


def _anchor_ids(root: ET.Element) -> dict[str, set[str]]:
    """Every Word anchor identity in one part, by domain (anchorId, editId)."""
    occupied: dict[str, set[str]] = {"anchorId": set(), "editId": set()}
    for node in root.iter():
        for name, value in node.attrib.items():
            if (domain := _ANCHOR_ID_DOMAINS.get(name)) is not None:
                occupied[domain].add(value.upper())
    return occupied


def _remap_anchor_ids(roots: list[ET.Element], occupied: dict[str, set[str]]) -> None:
    """Give imported anchors whose w14/wp14 anchorId or wp14 editId collides a fresh id.

    Word keeps these ids unique within a part (VML and DrawingML anchors share one
    anchorId domain), and no other markup refers to them, so only the colliding
    attribute changes. Reserve every incoming id first so a reallocation cannot take
    a later imported value.
    """
    reserved = {domain: set(values) for domain, values in occupied.items()}
    for root in roots:
        for domain, values in _anchor_ids(root).items():
            reserved[domain] |= values
    for root in roots:
        for node in root.iter():
            for name, domain in _ANCHOR_ID_DOMAINS.items():
                if (value := node.get(name)) is None:
                    continue
                value = value.upper()
                if value in occupied[domain]:
                    value = generate_collision_safe_para_id(reserved[domain], value)
                    node.set(name, value)
                occupied[domain].add(value)


def _write(
    path: Path,
    parts: Mapping[str, bytes],
    original: Path,
    *,
    removed_parts: set[str] | None = None,
) -> None:
    with (
        zipfile.ZipFile(original) as archive,
        zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as output,
    ):
        seen: set[str] = set()
        for info in archive.infolist():
            if removed_parts is not None and info.filename in removed_parts:
                continue
            if info.filename in seen:
                continue
            seen.add(info.filename)
            output.writestr(info, parts.get(info.filename, archive.read(info.filename)))
        for name, data in parts.items():
            if name not in seen:
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o600 << 16
                output.writestr(info, data)


def _omit_import_text_boxes(roots: list[ET.Element], ns: str) -> None:
    """Omit text-box drawings atomically, including their alternate fallback."""
    containers = {Q(ns, "drawing"), Q(ns, "pict"), Q(ns, "object"), _MC_ALTERNATE_CONTENT}
    box = Q(ns, "txbxContent")
    for root in roots:
        if root.tag == box or root.tag in containers and next(root.iter(box), None) is not None:
            raise ValueError("assembly_text_box_root_ignored")
        for parent in root.iter():
            for child in list(parent):
                if child.tag in containers and next(child.iter(box), None) is not None:
                    parent.remove(child)


def _start_import_on_new_page(selected: list[ET.Element], ns: str) -> None:
    """Use a paragraph boundary, not a section break or altered page geometry."""
    if not selected:
        return
    first = selected[0]
    if first.tag != Q(ns, "p"):
        first = ET.Element(Q(ns, "p"))
        selected.insert(0, first)
    properties = first.find(Q(ns, "pPr"))
    if properties is None:
        properties = ET.Element(Q(ns, "pPr"))
        first.insert(0, properties)
    page = properties.find(Q(ns, "pageBreakBefore"))
    if page is None:
        page = ET.Element(Q(ns, "pageBreakBefore"))
        properties.append(page)
    page.set(Q(ns, "val"), "1")
    _sort_children(properties, _PPR_ORDER)


def _adjust_sentence_word_spacing(
    incoming: ET.Element, left: ET.Element | None, right: ET.Element | None, ns: str
) -> None:
    """Use word spacing, not ASCII spaces to simulate East Asian glyph spacing."""
    texts = [item for item in incoming.iter(Q(ns, "t")) if item.text]
    if not texts:
        return
    before = "".join(item.text or "" for item in left.iter(Q(ns, "t"))) if left is not None else ""
    after = "".join(item.text or "" for item in right.iter(Q(ns, "t"))) if right is not None else ""

    def word(character: str) -> bool:
        return character.isalnum() or character == "_"

    def unspaced(character: str) -> bool:
        codepoint = ord(character)
        return (
            unicodedata.east_asian_width(character) in {"W", "F"}
            or 0x0E00 <= codepoint <= 0x0FFF
            or 0x1780 <= codepoint <= 0x17FF
            or 0x1000 <= codepoint <= 0x109F
        )

    def space_between(previous_text: str, following_text: str) -> bool:
        if unicodedata.category(following_text[0]).startswith("M"):
            return False
        previous = next(
            (
                character
                for character in reversed(previous_text)
                if not unicodedata.category(character).startswith("M")
            ),
            previous_text[-1],
        )
        following = following_text[0]
        if unspaced(previous) or unspaced(following):
            return False
        return word(following) and (word(previous) or previous in ".!?؟؛")

    first, last = texts[0], texts[-1]
    if before and first.text:
        if before.endswith(" ") and first.text.startswith(" "):
            first.text = first.text.lstrip(" ")
        elif space_between(before, first.text):
            first.text = " " + first.text
    if after and last.text:
        if after.startswith(" ") and last.text.endswith(" "):
            last.text = last.text.rstrip(" ")
        elif space_between(last.text, after):
            last.text += " "
    for item in (first, last):
        if item.text and (item.text.startswith(" ") or item.text.endswith(" ")):
            item.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _pin_implicit_import_styles(roots: list[ET.Element], styles: ET.Element, ns: str) -> None:
    defaults = {
        item.get(Q(ns, "type")): item.get(Q(ns, "styleId"))
        for item in styles.findall(Q(ns, "style"))
        if item.get(Q(ns, "default"), "").lower() in {"1", "true", "on"}
    }
    if "paragraph" not in defaults and any(
        item.get(Q(ns, "styleId")) == "Normal" for item in styles.findall(Q(ns, "style"))
    ):
        defaults["paragraph"] = "Normal"
    for root in roots:
        for kind, tag, properties_tag, reference_tag in (
            ("paragraph", "p", "pPr", "pStyle"),
            ("table", "tbl", "tblPr", "tblStyle"),
        ):
            identifier = defaults.get(kind)
            if not identifier:
                continue
            for item in root.iter(Q(ns, tag)):
                properties = item.find(Q(ns, properties_tag))
                if properties is not None and properties.find(Q(ns, reference_tag)) is not None:
                    continue
                if properties is None:
                    properties = ET.Element(Q(ns, properties_tag))
                    item.insert(0, properties)
                properties.insert(0, ET.Element(Q(ns, reference_tag), {Q(ns, "val"): identifier}))


def _prepare_import_theme(
    source: dict[str, bytes],
    roots: list[ET.Element],
    ns: str,
    tokens: set[str],
    options: ImportFormatOptions,
    *,
    pin_implicit: bool = True,
) -> None:
    """Resolve only formatting definitions reachable from the imported content."""
    themed = list(roots)
    styles = _styles_root(source)
    by_id: dict[str, ET.Element] = {}
    required: set[str] = set()
    if styles is not None:
        if pin_implicit:
            _pin_implicit_import_styles(roots, styles, ns)
        required, by_id = _required_styles(styles, roots, ns)
        _extend_required_styles(required, by_id, tokens, ns)
        _extend_numbering_style_dependencies(source, roots, required, by_id, ns)
        themed.extend(by_id[identifier] for identifier in required)
        defaults = styles.find(Q(ns, "docDefaults"))
        if defaults is not None:
            themed.append(defaults)
    numbering = (
        parse_xml_with_preserved_qnames(source["word/numbering.xml"])
        if "word/numbering.xml" in source
        else None
    )
    if numbering is not None:
        used = {
            item.get(Q(ns, "val"))
            for root in [*roots, *(by_id[identifier] for identifier in required)]
            for item in root.iter(Q(ns, "numId"))
        }
        numbers = {item.get(Q(ns, "numId")): item for item in numbering.findall(Q(ns, "num"))}
        abstracts = {
            item.get(Q(ns, "abstractNumId")): item
            for item in numbering.findall(Q(ns, "abstractNum"))
        }
        for identifier in used:
            number = numbers.get(identifier)
            if number is not None:
                themed.append(number)
                reference = number.find(Q(ns, "abstractNumId"))
                abstract = (
                    abstracts.get(reference.get(Q(ns, "val"))) if reference is not None else None
                )
                if abstract is not None:
                    themed.append(abstract)
    _resolve_theme_dependencies(source, themed, ns, resolve_colors=options.resolve_theme_colors)
    if styles is not None:
        source[_STYLES] = serialize_xml_with_preserved_qnames(styles)
    if numbering is not None:
        source["word/numbering.xml"] = serialize_xml_with_preserved_qnames(numbering)


class _NodeResourceImporter(_Importer):
    """Session import permits only resource graphs with established copy semantics."""

    _RELATION_KINDS = frozenset(
        {
            "image",
            "hyperlink",
            "header",
            "footer",
            "chart",
            "chartStyle",
            "chartColorStyle",
            "diagramData",
            "diagramLayout",
            "diagramQuickStyle",
            "diagramColors",
            "diagramDrawing",
            "customXml",
            "customXmlProps",
            "theme",
            "themeOverride",
        }
    )

    def copy_relation(
        self, source_owner: str, source_id: str, destination_owner: str | None = None
    ) -> str:
        relation = next(
            (item for item in _read_rels(self.source, source_owner) if item.get("Id") == source_id),
            None,
        )
        if relation is None:
            raise ValueError("assembly_relationship_missing")
        self._check_relation(relation)
        return super().copy_relation(source_owner, source_id, destination_owner)

    def _check_relation(self, relation: ET.Element) -> None:
        kind = relation.get("Type", "").rsplit("/", 1)[-1]
        if kind in {"oleObject", "control", "vbaProject", "activeXControl", "activeXControlBinary"}:
            raise ValueError("node_import_active_content_unsupported")
        if kind not in self._RELATION_KINDS or (
            relation.get("TargetMode") == "External" and kind not in {"hyperlink", "image"}
        ):
            raise ValueError("node_import_resource_unsupported")

    def copy_part(self, source_part: str) -> str:
        content_type = _content_type(source_part, self.source_types)
        if content_type is None:
            raise ValueError("assembly_content_type_missing")
        lowered = content_type[1].casefold()
        if any(token in lowered for token in ("macro", "vba", "activex", "oleobject")):
            raise ValueError("node_import_active_content_unsupported")
        relations = _read_rels(self.source, source_part)
        for relation in relations:
            self._check_relation(relation)
        if (
            getattr(self, "resolve_colors", False)
            and lowered.endswith("+xml")
            and ("drawingml" in lowered or "wordprocessingml" in lowered)
        ):
            root = parse_xml_with_preserved_qnames(self.source[source_part])
            _resolve_theme_dependencies(self.source, [root], self.word_namespace, owner=source_part)
            self.source[source_part] = serialize_xml_with_preserved_qnames(root)
        return super().copy_part(source_part)


def _import_session_bytes(document: Document, parts: Mapping[str, bytes]) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.comment = document._comment
        for name, data in parts.items():
            info = copy.copy(document._entries.get(name)) if name in document._entries else None
            if info is None:
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o600 << 16
            archive.writestr(info, data)
    return buffer.getvalue()


def _pin_inline_import_context(
    node: Node, incoming: ET.Element, parts: Mapping[str, bytes], ns: str
) -> None:
    """Materialize inherited run formatting when its source paragraph is not imported."""

    if incoming.tag in {Q(ns, name) for name in ("tbl", "sectPr")}:
        return
    ancestor = node.parent
    paragraph = None
    while ancestor is not None:
        if ancestor.xml_name == Q(ns, "p") and paragraph is None:
            paragraph = ancestor._element
        ancestor = ancestor.parent
    if incoming.tag == Q(ns, "p"):
        return
    styles = _styles_root(parts)
    if styles is None:
        return
    definitions = {item.get(Q(ns, "styleId")): item for item in styles.findall(Q(ns, "style"))}
    defaults = styles.find(f"{Q(ns, 'docDefaults')}/{Q(ns, 'rPrDefault')}/{Q(ns, 'rPr')}")
    paragraph_style = None
    if paragraph is not None:
        reference = paragraph.find(f"{Q(ns, 'pPr')}/{Q(ns, 'pStyle')}")
        paragraph_style = reference.get(Q(ns, "val")) if reference is not None else None
    if paragraph_style is None:
        paragraph_style = next(
            (
                identifier
                for identifier, item in definitions.items()
                if item.get(Q(ns, "type")) == "paragraph"
                and item.get(Q(ns, "default"), "").lower() in {"1", "true", "on"}
            ),
            "Normal" if "Normal" in definitions else None,
        )

    def chain(identifier: str | None) -> list[ET.Element]:
        layers = []
        seen = set()
        while identifier:
            if identifier in seen:
                raise ValueError("node_import_style_cycle")
            seen.add(identifier)
            style = definitions.get(identifier)
            if style is None:
                raise ValueError("assembly_style_missing")
            properties = style.find(Q(ns, "rPr"))
            if properties is not None:
                layers.append(properties)
            reference = style.find(Q(ns, "basedOn"))
            identifier = reference.get(Q(ns, "val")) if reference is not None else None
        return list(reversed(layers))

    toggles = {
        Q(ns, name)
        for name in (
            "b",
            "bCs",
            "i",
            "iCs",
            "caps",
            "smallCaps",
            "strike",
            "dstrike",
            "outline",
            "shadow",
            "emboss",
            "imprint",
            "vanish",
            "specVanish",
        )
    }
    paragraph_runs = {
        id(run) for paragraph_node in incoming.iter(Q(ns, "p")) for run in paragraph_node.iter()
    }
    for run in incoming.iter():
        if run.tag != Q(ns, "r") and not (
            run.tag.endswith("}r") and run.tag[1:].split("}", 1)[0].endswith("/math")
        ):
            continue
        if id(run) in paragraph_runs:
            continue
        direct = run.find(Q(ns, "rPr"))
        reference = direct.find(Q(ns, "rStyle")) if direct is not None else None
        character_style = reference.get(Q(ns, "val")) if reference is not None else None
        layers = [
            *([(defaults, False)] if defaults is not None else []),
            *((item, True) for item in chain(paragraph_style)),
            *((item, True) for item in chain(character_style)),
            *([(direct, False)] if direct is not None else []),
        ]
        effective: dict[str, ET.Element] = {}
        for properties, style_layer in layers:
            for item in properties:
                if item.tag in toggles and style_layer:
                    enabled = item.get(Q(ns, "val"), "1").lower() not in {"0", "false", "off"}
                    if not enabled:
                        continue
                    previous = effective.get(item.tag)
                    was_enabled = previous is not None and previous.get(
                        Q(ns, "val"), "1"
                    ).lower() not in {"0", "false", "off"}
                    effective[item.tag] = ET.Element(
                        item.tag, {Q(ns, "val"): "0" if was_enabled else "1"}
                    )
                else:
                    clone = copy.deepcopy(item)
                    if item.tag in {Q(ns, "rFonts"), Q(ns, "lang")} and item.tag in effective:
                        clone.attrib = {**effective[item.tag].attrib, **clone.attrib}
                        if item.tag == Q(ns, "rFonts"):
                            for plain, themed in (
                                ("ascii", "asciiTheme"),
                                ("hAnsi", "hAnsiTheme"),
                                ("eastAsia", "eastAsiaTheme"),
                                ("cs", "cstheme"),
                            ):
                                if Q(ns, plain) in item.attrib and Q(ns, themed) not in item.attrib:
                                    clone.attrib.pop(Q(ns, themed), None)
                                elif (
                                    Q(ns, themed) in item.attrib and Q(ns, plain) not in item.attrib
                                ):
                                    clone.attrib.pop(Q(ns, plain), None)
                    effective[item.tag] = clone
        if effective:
            if direct is None:
                direct = ET.Element(Q(ns, "rPr"))
                math_properties = next(
                    (
                        index
                        for index, item in enumerate(run)
                        if item.tag.endswith("}rPr")
                        and item.tag[1:].split("}", 1)[0].endswith("/math")
                    ),
                    None,
                )
                run.insert(0 if math_properties is None else math_properties + 1, direct)
            direct[:] = list(effective.values())
            _sort_children(direct, _RPR_ORDER)


def _numbering_identity(root: ET.Element, identifier: str, ns: str) -> tuple[str, str] | None:
    number = next(
        (item for item in root.findall(Q(ns, "num")) if item.get(Q(ns, "numId")) == identifier),
        None,
    )
    if number is None:
        return None
    reference = number.find(Q(ns, "abstractNumId"))
    abstract = next(
        (
            item
            for item in root.findall(Q(ns, "abstractNum"))
            if reference is not None
            and item.get(Q(ns, "abstractNumId")) == reference.get(Q(ns, "val"))
        ),
        None,
    )
    return (_xml_hash(number), _xml_hash(abstract)) if abstract is not None else None


def _import_definition_context(parts: Mapping[str, bytes]) -> str:
    signature = hashlib.sha256()
    for name, data in sorted(parts.items()):
        if name in {_STYLES, "word/settings.xml"} or name.startswith("word/theme/"):
            signature.update(name.encode("utf-8"))
            signature.update(hashlib.sha256(data).digest())
    return signature.hexdigest()


def _validate_import_fragment(node: Node, roots: list[ET.Element], ns: str) -> None:
    """Keep whole supported constructs, never detach their semantic envelope."""
    from .review import _SELECTABLE_REVISION_TAGS

    if node.xml_name == Q(ns, "txbxContent"):
        raise ValueError("node_import_text_box_pair_incomplete")
    story = node
    while story.parent is not None and story.parent._element is not None:
        story = story.parent
    order = list(story._element.iter())
    selected = {id(item) for item in node._element.iter()}
    positions = [index for index, item in enumerate(order) if id(item) in selected]
    if positions:
        for start, end, construct in (
            ("bookmarkStart", "bookmarkEnd", "bookmark"),
            ("commentRangeStart", "commentRangeEnd", "comment"),
            ("permStart", "permEnd", "permission"),
        ):
            openings = {}
            for index, item in enumerate(order):
                if item.tag == Q(ns, start):
                    openings[item.get(Q(ns, "id"))] = (index, item)
                elif item.tag == Q(ns, end):
                    opening = openings.pop(item.get(Q(ns, "id")), None)
                    if (
                        opening is not None
                        and opening[0] <= positions[-1]
                        and index >= positions[0]
                        and (id(opening[1]) not in selected or id(item) not in selected)
                    ):
                        raise ValueError(f"node_import_{construct}_range_incomplete")
            if any(index <= positions[-1] for index, _ in openings.values()):
                raise ValueError(f"node_import_{construct}_range_incomplete")

    ancestor = node.parent
    while ancestor is not None and ancestor._element is not None:
        local = ancestor.xml_name.rsplit("}", 1)[-1]
        if local in {"ins", "del", "moveFrom", "moveTo"}:
            raise ValueError(f"node_import_{local}_incomplete")
        if local == "sdt":
            properties = ancestor._element.find(Q(ns, "sdtPr"))
            if properties is not None and any(
                item.get(Q(ns, "val")) != "unlocked" for item in properties.findall(Q(ns, "lock"))
            ):
                raise ValueError("node_import_sdt_locked_fragment")
        if local == "txbxContent":
            raise ValueError("node_import_text_box_pair_incomplete")
        if ancestor.xml_name == _MC_ALTERNATE_CONTENT and any(
            item.tag == Q(ns, "txbxContent") for item in ancestor._element.iter()
        ):
            raise ValueError("node_import_text_box_pair_incomplete")
        if local in {"wgp", "grpSp", "group"}:
            raise ValueError("node_import_drawing_group_incomplete")
        ancestor = ancestor.parent
    for root in roots:
        parents = {child: parent for parent in root.iter() for child in parent}
        shared_revisions = {}
        tags = {Q(ns, name) for name in _SELECTABLE_REVISION_TAGS}
        for revision in root.iter():
            if revision.tag not in tags:
                continue
            identifier = revision.get(Q(ns, "id"))
            if identifier in shared_revisions:
                construct = revision.tag.rsplit("}", 1)[-1]
                raise ValueError(f"node_import_{construct}_shared_identity_unsupported")
            shared_revisions[identifier] = revision
        for item in root.iter():
            local = item.tag.rsplit("}", 1)[-1]
            if local in {"OLEObject", "control", "activeX", "object"}:
                raise ValueError("node_import_active_content_unsupported")
            if item.tag in {Q(ns, name) for name in ("altChunk", "customXml")}:
                raise ValueError(f"node_import_{local}_unsupported")
            if item.tag == Q(ns, "sdt"):
                properties = item.findall(Q(ns, "sdtPr"))
                if len(properties) != 1 or len(item.findall(Q(ns, "sdtContent"))) != 1:
                    raise ValueError("node_import_sdt_incomplete")
                clone = copy.deepcopy(item)
                for lock in clone.find(Q(ns, "sdtPr")).findall(Q(ns, "lock")):
                    lock.set(Q(ns, "val"), "unlocked")
                cloned_parents = {
                    **parents,
                    **{child: parent for parent in clone.iter() for child in parent},
                }
                cloned_parents[clone] = parents.get(item)
                if not _supported_sdt(clone, cloned_parents, ns):
                    construct = (
                        "sdt_web_extension"
                        if any(child.tag in _SDT_FOREIGN_PROPERTIES for child in properties[0])
                        else "sdt_repeating_section"
                        if any(child.tag.endswith("}repeatingSection") for child in properties[0])
                        else "sdt"
                    )
                    raise ValueError(f"node_import_{construct}_unsupported")
            if item.tag.startswith(f"{{{ns}}}") and (
                local in {"moveFrom", "moveTo"}
                or local.startswith(("moveFromRange", "moveToRange", "customXml"))
                or local.endswith("Change")
                and local not in _SELECTABLE_REVISION_TAGS
            ):
                raise ValueError(f"node_import_{local}_unsupported")


def _reconcile_import_math(
    source: Mapping[str, bytes],
    destination: Mapping[str, bytes],
    roots: list[ET.Element],
    ns: str,
) -> None:
    """Materialize local math defaults without changing destination mathPr."""
    math = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    if not any(item.tag.startswith(f"{{{math}}}") for root in roots for item in root.iter()):
        return

    def settings(parts: Mapping[str, bytes]) -> dict[str, ET.Element]:
        relations = [
            item
            for item in _read_rels(parts, _DOCUMENT)
            if item.get("Type", "").endswith("/settings")
        ]
        if not relations:
            return {}
        if len(relations) != 1 or relations[0].get("TargetMode") == "External":
            raise ValueError("node_import_math_settings_invalid")
        part = _resolve(_DOCUMENT, relations[0].get("Target", ""))
        root = parse_xml_with_preserved_qnames(parts[part])
        properties = root.find(Q(math, "mathPr"))
        return {
            item.tag.rsplit("}", 1)[-1]: item
            for item in (list(properties) if properties is not None else [])
        }

    original, target = settings(source), settings(destination)
    defaults = {
        "mathFont": "Cambria Math",
        "defJc": "centerGroup",
        "naryLim": "undOvr",
        "intLim": "subSup",
    }
    differences = set()
    for name in original.keys() | target.keys():
        left, right = original.get(name), target.get(name)
        if name in defaults:
            left_value = (
                left.get(Q(math, "val"), defaults[name]) if left is not None else defaults[name]
            )
            right_value = (
                right.get(Q(math, "val"), defaults[name]) if right is not None else defaults[name]
            )
            if left_value != right_value:
                differences.add(name)
        elif (ET.tostring(left) if left is not None else None) != (
            ET.tostring(right) if right is not None else None
        ):
            differences.add(name)
    for name in sorted(differences - defaults.keys()):
        raise ValueError(f"node_import_math_{name}_unsupported")
    for root in roots:
        if "mathFont" in differences:
            font = original.get("mathFont")
            face = (
                font.get(Q(math, "val"), defaults["mathFont"])
                if font is not None
                else defaults["mathFont"]
            )
            for run in root.iter(Q(math, "r")):
                properties = run.find(Q(ns, "rPr"))
                if properties is None:
                    properties = ET.Element(Q(ns, "rPr"))
                    run.insert(1 if run.find(Q(math, "rPr")) is not None else 0, properties)
                fonts = properties.find(Q(ns, "rFonts"))
                if fonts is None:
                    fonts = ET.Element(Q(ns, "rFonts"))
                    properties.insert(0, fonts)
                for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                    fonts.attrib.setdefault(Q(ns, attribute), face)
                _sort_children(properties, _RPR_ORDER)
        if "defJc" in differences:
            value = original.get("defJc")
            for paragraph in root.iter(Q(math, "oMathPara")):
                properties = paragraph.find(Q(math, "oMathParaPr"))
                if properties is None:
                    properties = ET.Element(Q(math, "oMathParaPr"))
                    paragraph.insert(0, properties)
                if properties.find(Q(math, "jc")) is None:
                    ET.SubElement(
                        properties,
                        Q(math, "jc"),
                        {
                            Q(math, "val"): value.get(Q(math, "val"))
                            if value is not None
                            else defaults["defJc"]
                        },
                    )
        for nary in root.iter(Q(math, "nary")):
            properties = nary.find(Q(math, "naryPr"))
            if properties is None:
                properties = ET.Element(Q(math, "naryPr"))
                nary.insert(0, properties)
            character = properties.find(Q(math, "chr"))
            symbol = character.get(Q(math, "val"), "∫") if character is not None else "∫"
            name = "intLim" if symbol in "∫∬∭∮∯∰" else "naryLim"
            if name in differences and properties.find(Q(math, "limLoc")) is None:
                value = original.get(name)
                limit = ET.Element(
                    Q(math, "limLoc"),
                    {
                        Q(math, "val"): value.get(Q(math, "val"))
                        if value is not None
                        else defaults[name]
                    },
                )
                properties.insert(1 if character is not None else 0, limit)


def _materialize_import_table_context(
    node: Node,
    incoming: ET.Element,
    parts: Mapping[str, bytes],
    ns: str,
    memo: Mapping[int, Any],
) -> None:
    """Pin cell-position-dependent formatting before leaving the source table."""
    from .table_cell_properties import _TCPR_CHILD_ORDER
    from .table_row_properties import _TRPR_CHILD_ORDER

    owner = node if node.xml_name == Q(ns, "tbl") else node.parent
    while owner is not None and owner.xml_name != Q(ns, "tbl"):
        owner = owner.parent
    if owner is None or owner._element is None:
        return
    table = owner._element
    properties = table.find(Q(ns, "tblPr"))
    reference = properties.find(Q(ns, "tblStyle")) if properties is not None else None
    styles = _styles_root(parts)
    if reference is None or styles is None:
        return
    by_id = {item.get(Q(ns, "styleId")): item for item in styles.findall(Q(ns, "style"))}
    layers = []
    identifier = reference.get(Q(ns, "val"))
    seen = set()
    while identifier:
        if identifier in seen:
            raise ValueError("node_import_style_cycle")
        seen.add(identifier)
        style = by_id.get(identifier)
        if style is None:
            raise ValueError("node_import_style_missing")
        layers.insert(0, style)
        base = style.find(Q(ns, "basedOn"))
        identifier = base.get(Q(ns, "val")) if base is not None else None
    parents = {child: parent for parent in table.iter() for child in parent}

    def table_owner(element: ET.Element) -> ET.Element | None:
        while element in parents:
            element = parents[element]
            if element.tag == Q(ns, "tbl"):
                return element
        return None

    rows = [item for item in table.iter(Q(ns, "tr")) if table_owner(item) is table]
    look = ET.Element(Q(ns, "tblLook"))
    for layer in [*(item.find(Q(ns, "tblPr")) for item in layers), properties]:
        declared = layer.find(Q(ns, "tblLook")) if layer is not None else None
        if declared is not None:
            look.attrib.update(declared.attrib)
    try:
        bits = int(look.get(Q(ns, "val"), "0000"), 16)
    except ValueError as error:
        raise ValueError("node_import_table_look_invalid") from error

    def flag(name: str, mask: int) -> bool:
        value = look.get(Q(ns, name)) if look is not None else None
        return value not in {"0", "false", "off"} if value is not None else bool(bits & mask)

    def band(name: str) -> int:
        value = None
        for layer in [*(item.find(Q(ns, "tblPr")) for item in layers), properties]:
            item = layer.find(Q(ns, name)) if layer is not None else None
            if item is not None:
                value = item.get(Q(ns, "val"))
        if value is not None and (not value.isdigit() or int(value) < 1):
            raise ValueError("node_import_table_band_size_invalid")
        return int(value or "1")

    def merge(target: ET.Element, source_properties: ET.Element | None) -> None:
        if source_properties is None:
            return
        for child in source_properties:
            previous = target.find(child.tag)
            clone = copy.deepcopy(child)
            if previous is not None:
                if child.tag in {Q(ns, name) for name in ("rFonts", "lang", "ind", "spacing")}:
                    clone.attrib = {**previous.attrib, **clone.attrib}
                elif child.tag.endswith("Borders"):
                    merged = copy.deepcopy(previous)
                    merge(merged, child)
                    clone = merged
                target.remove(previous)
            target.append(clone)

    for row_index, row in enumerate(rows):
        cells = [item for item in row.iter(Q(ns, "tc")) if table_owner(item) is table]
        grid_before = row.find(f"{Q(ns, 'trPr')}/{Q(ns, 'gridBefore')}")
        column = int(grid_before.get(Q(ns, "val"), "0")) if grid_before is not None else 0
        for cell_index, cell in enumerate(cells):
            conditions = ["wholeTable"]
            if not flag("noHBand", 0x200):
                offset = row_index - int(flag("firstRow", 0x20))
                if offset >= 0:
                    conditions.append(
                        "band1Horz"
                        if offset // band("tblStyleRowBandSize") % 2 == 0
                        else "band2Horz"
                    )
            if not flag("noVBand", 0x400):
                offset = column - int(flag("firstColumn", 0x80))
                if offset >= 0:
                    conditions.append(
                        "band1Vert"
                        if offset // band("tblStyleColBandSize") % 2 == 0
                        else "band2Vert"
                    )
            first_row = row_index == 0 and flag("firstRow", 0x20)
            last_row = row_index == len(rows) - 1 and flag("lastRow", 0x40)
            first_col = cell_index == 0 and flag("firstColumn", 0x80)
            last_col = cell_index == len(cells) - 1 and flag("lastColumn", 0x100)
            conditions.extend(
                name
                for name, enabled in (
                    ("firstCol", first_col),
                    ("lastCol", last_col),
                    ("firstRow", first_row),
                    ("lastRow", last_row),
                    ("nwCell", first_row and first_col),
                    ("neCell", first_row and last_col),
                    ("swCell", last_row and first_col),
                    ("seCell", last_row and last_col),
                )
                if enabled
            )
            conditional_flags = (
                ("firstRow", "firstRow"),
                ("lastRow", "lastRow"),
                ("firstColumn", "firstCol"),
                ("lastColumn", "lastCol"),
                ("oddVBand", "band1Vert"),
                ("evenVBand", "band2Vert"),
                ("oddHBand", "band1Horz"),
                ("evenHBand", "band2Horz"),
                ("firstRowFirstColumn", "nwCell"),
                ("firstRowLastColumn", "neCell"),
                ("lastRowFirstColumn", "swCell"),
                ("lastRowLastColumn", "seCell"),
            )

            def apply_flags(
                active: list[str],
                declared: ET.Element | None,
                flags: tuple[tuple[str, str], ...] = conditional_flags,
            ) -> None:
                if declared is None:
                    return
                value = declared.get(Q(ns, "val"))
                if value is not None and not re.fullmatch("[01]{12}", value):
                    raise ValueError("node_import_table_conditional_flags_invalid")
                for position, (attribute, condition) in enumerate(flags):
                    enabled = declared.get(Q(ns, attribute))
                    if enabled is None and value is not None:
                        enabled = value[position]
                    if enabled is None:
                        continue
                    if condition in active:
                        active.remove(condition)
                    if enabled not in {"0", "false", "off"}:
                        active.append(condition)

            for context in (row, cell):
                properties_tag = "trPr" if context is row else "tcPr"
                declared = context.find(f"{Q(ns, properties_tag)}/{Q(ns, 'cnfStyle')}")
                apply_flags(conditions, declared)
            precedence = {
                name: rank
                for rank, name in enumerate(
                    (
                        "wholeTable",
                        "band1Vert",
                        "band2Vert",
                        "band1Horz",
                        "band2Horz",
                        "firstCol",
                        "lastCol",
                        "firstRow",
                        "lastRow",
                        "nwCell",
                        "neCell",
                        "swCell",
                        "seCell",
                    )
                )
            }

            def resolve(
                active: list[str], ranks: dict[str, int] = precedence
            ) -> dict[str, ET.Element]:
                effective = {
                    name: ET.Element(Q(ns, name)) for name in ("pPr", "rPr", "tcPr", "trPr")
                }
                for layer in layers:
                    for name, target in effective.items():
                        merge(target, layer.find(Q(ns, name)))
                for condition in sorted(active, key=ranks.__getitem__):
                    for layer in layers:
                        for override in layer.findall(Q(ns, "tblStylePr")):
                            if override.get(Q(ns, "type")) == condition:
                                for name, target in effective.items():
                                    merge(target, override.find(Q(ns, name)))
                return effective

            effective = resolve(conditions)
            merge(effective["tcPr"], cell.find(Q(ns, "tcPr")))
            for original in cell.iter():
                if original is not cell and table_owner(original) is not table:
                    continue
                cloned = memo.get(id(original))
                if cloned is None:
                    continue
                name, order = {
                    Q(ns, "p"): ("pPr", _PPR_ORDER),
                    Q(ns, "r"): ("rPr", _RPR_ORDER),
                    Q(ns, "tc"): (
                        "tcPr",
                        {name: rank for rank, name in enumerate(_TCPR_CHILD_ORDER)},
                    ),
                }.get(original.tag, (None, None))
                if name is None:
                    continue
                active = list(conditions)
                context = original
                while context is not cell and context.tag != Q(ns, "p") and context in parents:
                    context = parents[context]
                if context.tag == Q(ns, "p"):
                    apply_flags(active, context.find(f"{Q(ns, 'pPr')}/{Q(ns, 'cnfStyle')}"))
                local_effective = resolve(active) if active != conditions else effective
                pinned = copy.deepcopy(local_effective[name])
                if name == "pPr" and node.xml_name == Q(ns, "p"):
                    shading = effective["tcPr"].find(Q(ns, "shd"))
                    if shading is not None and pinned.find(Q(ns, "shd")) is None:
                        pinned.append(copy.deepcopy(shading))
                direct = cloned.find(Q(ns, name))
                merge(pinned, direct)
                if len(pinned):
                    if direct is not None:
                        cloned.remove(direct)
                    cloned.insert(0, pinned)
                    _sort_children(pinned, order)
            cloned_row = memo.get(id(row))
            if cloned_row is not None and len(effective["trPr"]):
                pinned = copy.deepcopy(effective["trPr"])
                direct = cloned_row.find(Q(ns, "trPr"))
                merge(pinned, direct)
                if direct is not None:
                    cloned_row.remove(direct)
                cloned_row.insert(0, pinned)
                _sort_children(pinned, {name: rank for rank, name in enumerate(_TRPR_CHILD_ORDER)})
            span = cell.find(f"{Q(ns, 'tcPr')}/{Q(ns, 'gridSpan')}")
            column += int(span.get(Q(ns, "val"), "1")) if span is not None else 1


_FRAG_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_FRAG_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
_FRAG_WPC = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
_FRAG_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_FRAG_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_FRAG_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_FRAG_VML_SHAPES = frozenset(
    Q(_VML, name)
    for name in (
        "shape",
        "rect",
        "roundrect",
        "oval",
        "line",
        "polyline",
        "image",
        "arc",
        "curve",
    )
)
_FRAG_GROUPS = frozenset(
    {Q(_FRAG_WPG, "wgp"), Q(_FRAG_WPG, "grpSp"), Q(_FRAG_WPC, "wpc"), Q(_VML, "group")}
)
_FRAG_CHILDREN = frozenset(
    {Q(_FRAG_WPS, "wsp"), Q(_FRAG_PIC, "pic"), Q(_FRAG_WPG, "grpSp"), Q(_FRAG_WPG, "wgp")}
    | _FRAG_VML_SHAPES
    | {Q(_VML, "group")}
)


def _is_drawing_fragment(node: Node, ns: str) -> bool:
    """A text-box story or an individual child of a drawing group."""
    if node.xml_name == Q(ns, "txbxContent"):
        return True
    if node.xml_name not in _FRAG_CHILDREN:
        return False
    ancestor = node.parent
    while ancestor is not None and ancestor._element is not None:
        if ancestor.xml_name in _FRAG_GROUPS:
            return True
        ancestor = ancestor.parent
    return False


def _sync_text_box_pair(target: ET.Element, container: ET.Element, ns: str) -> bool:
    """Make every other AlternateContent branch carry the target's text-box content."""
    parents = {child: parent for parent in container.iter() for child in parent}
    box = Q(ns, "txbxContent")
    alternate = None
    branch = target
    while branch in parents:
        if parents[branch].tag == _MC_ALTERNATE_CONTENT:
            alternate = parents[branch]
            break
        branch = parents[branch]
    if alternate is None:
        return False
    mine = list(branch.iter(box))
    ordinal = next(index for index, item in enumerate(mine) if item is target)
    changed = False
    for other in alternate:
        if other is branch or other.tag not in {
            Q(_MC, "Choice"),
            Q(_MC, "Fallback"),
        }:
            continue
        theirs = list(other.iter(box))
        if len(theirs) != len(mine):
            raise ValueError("node_import_text_box_pair_incomplete")
        counterpart = theirs[ordinal]
        replacement = [copy.deepcopy(child) for child in target]
        for child in replacement:
            for item in child.iter():
                item.attrib.pop(_PARA_ID, None)
                item.attrib.pop(f"{{{_W14}}}textId", None)
        before = [ET.tostring(child) for child in counterpart]
        if before != [ET.tostring(child) for child in replacement] or (counterpart.text or "") != (
            target.text or ""
        ):
            counterpart[:] = replacement
            counterpart.text = target.text
            changed = True
    return changed


def _extract_drawing_fragment(
    node: Node, ns: str
) -> tuple[ET.Element, dict[int, Any], Node] | None:
    """Detach one text box or group child as a complete run with a synchronized pair.

    Returns the detached run, its deep-copy memo and the live source run node, or
    ``None`` when the selection has no enclosing run (the caller then refuses).
    Both AlternateContent branches carry the selection; group-child coordinates are
    rebased to a standalone drawing at the origin.
    """
    shapes = feature_module("shapes")

    run_node = node.parent
    while run_node is not None and run_node._element is not None:
        if run_node.xml_name == Q(ns, "r"):
            break
        run_node = run_node.parent
    if run_node is None or run_node._element is None:
        return None
    memo: dict[int, Any] = {}
    run = copy.deepcopy(run_node._element, memo)
    selected = memo[id(node._element)]
    parents = {child: parent for parent in run.iter() for child in parent}
    envelope = selected
    while parents.get(envelope) is not None and parents[envelope] is not run:
        envelope = parents[envelope]
    if parents.get(envelope) is not run:
        return None
    if any(item.tag == Q(_FRAG_WPS, "linkedTxbx") for item in envelope.iter()):
        raise ValueError("node_import_text_box_link_unsupported")
    target = selected
    if selected.tag == Q(ns, "txbxContent"):
        owner = parents.get(selected)
        while owner is not None and owner.tag not in (_FRAG_CHILDREN - {Q(_VML, "group")}):
            owner = parents.get(owner)
        if (
            owner is not None
            and parents.get(owner) is not None
            and (parents[owner].tag in _FRAG_GROUPS)
        ):
            target = owner
    try:
        if target.tag == Q(ns, "txbxContent"):
            _sync_text_box_pair(selected, run, ns)
        else:
            envelope = _extract_group_child(
                run,
                parents,
                target,
                envelope,
                ns,
                (
                    shapes._GROUP_TAGS,
                    shapes._group_lift,
                    shapes._group_local_transform,
                    shapes._group_vml_decimal,
                    shapes._group_vml_style,
                    shapes._metadata_children,
                    shapes._metadata_container,
                    shapes._metadata_root,
                    shapes._metadata_vml_fallback,
                ),
            )
    except shapes.DrawingShapeError as error:
        reason = error.code.removeprefix("drawing_group_").removeprefix("drawing_")
        raise ValueError(f"node_import_group_child_{reason}") from error
    run[:] = [item for item in run if item.tag == Q(ns, "rPr") or item is envelope]
    run.text = None
    envelope.tail = None
    return run, memo, run_node


def _extract_group_child(
    run: ET.Element,
    parents: dict[ET.Element, ET.Element],
    target: ET.Element,
    envelope: ET.Element,
    ns: str,
    helpers: tuple[Any, ...],
) -> ET.Element:
    (
        group_tags,
        group_lift,
        local_transform,
        vml_decimal,
        vml_style,
        children_of,
        container_of,
        root_of,
        vml_fallback,
    ) = helpers
    alternate = envelope if envelope.tag == _MC_ALTERNATE_CONTENT else None
    choice = fallback = None
    if alternate is not None:
        choices = [item for item in alternate if item.tag == Q(_MC, "Choice")]
        fallbacks = [item for item in alternate if item.tag == Q(_MC, "Fallback")]
        if len(choices) != 1 or len(fallbacks) != 1:
            raise ValueError("node_import_group_child_pair_unsupported")
        choice, fallback = choices[0], fallbacks[0]
        if len(choice) != 1 or len(fallback) != 1 or (choice.text or "").strip():
            raise ValueError("node_import_group_child_pair_unsupported")
    legacy = None
    if target.tag.startswith(f"{{{_VML}}}"):
        if alternate is None:
            raise ValueError("node_import_group_child_vml_unpaired_unsupported")
        lineage = [target]
        while parents[lineage[-1]] is not fallback:
            lineage.append(parents[lineage[-1]])
        legacy_chain = [
            item for item in reversed(lineage) if item.tag in _FRAG_VML_SHAPES | {Q(_VML, "group")}
        ]
        legacy_path = [
            [c for c in outer if c.tag in _FRAG_VML_SHAPES | {Q(_VML, "group")}].index(inner) + 1
            for outer, inner in pairwise(legacy_chain)
        ]
        drawings = list(choice.iter(Q(ns, "drawing")))
        if len(drawings) != 1:
            raise ValueError("node_import_group_child_vml_pairing_unsupported")
        drawing = drawings[0]
        legacy = vml_fallback(run, drawing, legacy_path)
        if legacy is not target:
            raise ValueError("node_import_group_child_vml_pairing_unsupported")
        dml = root_of(drawing)
        for index in legacy_path:
            dml = children_of(dml)[index - 1]
        path = legacy_path
    else:
        drawing = target
        while drawing.tag != Q(ns, "drawing"):
            if drawing not in parents:
                raise ValueError("node_import_group_child_drawing_missing")
            drawing = parents[drawing]
        dml = target
    top = root_of(drawing)
    lineage = [dml]
    while lineage[-1] is not top:
        lineage.append(parents.get(lineage[-1]))
        if lineage[-1] is None:
            raise ValueError("node_import_group_child_drawing_missing")
    lineage.reverse()
    path = [children_of(outer).index(inner) + 1 for outer, inner in pairwise(lineage)]
    if not path:
        raise ValueError("node_import_drawing_group_incomplete")
    if legacy is None and alternate is not None:
        legacy = vml_fallback(run, drawing, path)
    groups = [item for item in lineage[:-1] if item.tag in group_tags]
    identifiers = [
        int(value)
        for item in run.iter()
        for key, value in item.attrib.items()
        if key == "id" and value.isdigit()
    ]
    counter = max(identifiers, default=0)

    def allocate() -> int:
        nonlocal counter
        counter += 1
        return counter

    lifted = group_lift(copy.deepcopy(dml), groups, allocate)
    transform = local_transform(lifted)
    offset, extent = transform.find(Q(_FRAG_A, "off")), transform.find(Q(_FRAG_A, "ext"))
    shift_x, shift_y = int(offset.get("x")), int(offset.get("y"))
    width, height = int(extent.get("cx")), int(extent.get("cy"))
    offset.set("x", "0")
    offset.set("y", "0")
    wrapped = lifted.tag in group_tags and dml.tag not in group_tags
    if lifted.tag in group_tags:
        lifted.tag = Q(_FRAG_WPG, "wgp")
    uri = {
        Q(_FRAG_WPS, "wsp"): _FRAG_WPS,
        Q(_FRAG_PIC, "pic"): _FRAG_PIC,
        Q(_FRAG_WPG, "wgp"): _FRAG_WPG,
    }[lifted.tag]
    new_drawing = copy.deepcopy(drawing)
    container = container_of(new_drawing)
    data = container.find(f"{Q(_FRAG_A, 'graphic')}/{Q(_FRAG_A, 'graphicData')}")
    if data is None:
        raise ValueError("node_import_group_child_drawing_missing")
    data[:] = [lifted]
    data.text = None
    data.set("uri", uri)
    size = container.find(Q(_FRAG_WP, "extent"))
    if size is None:
        raise ValueError("node_import_group_child_drawing_missing")
    size.set("cx", str(width))
    size.set("cy", str(height))
    if container.tag == Q(_FRAG_WP, "anchor"):
        for name, delta in (("positionH", shift_x), ("positionV", shift_y)):
            position = container.find(Q(_FRAG_WP, name))
            offset_node = position.find(Q(_FRAG_WP, "posOffset")) if position is not None else None
            if offset_node is None or not (offset_node.text or "").strip().lstrip("-").isdigit():
                if delta:
                    raise ValueError("node_import_group_child_anchor_unsupported")
                continue
            offset_node.text = str(int(offset_node.text) + delta)
    if alternate is None:
        position = list(run).index(envelope)
        run[position] = new_drawing
        return new_drawing
    if wrapped:
        raise ValueError("node_import_group_child_vml_pair_unsupported")
    style = vml_style(legacy)
    for key in ("left", "top", "margin-left", "margin-top", "position"):
        style.pop(key, None)
    style["width"] = vml_decimal(Fraction(width, 12700)) + "pt"
    style["height"] = vml_decimal(Fraction(height, 12700)) + "pt"
    if container.tag == Q(_FRAG_WP, "anchor"):
        style["position"] = "absolute"
        for name, margin, axis in (
            ("positionH", "margin-left", "horizontal"),
            ("positionV", "margin-top", "vertical"),
        ):
            position = container.find(Q(_FRAG_WP, name))
            offset_node = position.find(Q(_FRAG_WP, "posOffset")) if position is not None else None
            if (
                position is None
                or position.get("relativeFrom") not in {"page", "margin"}
                or offset_node is None
                or not (offset_node.text or "").strip().lstrip("-").isdigit()
            ):
                raise ValueError("node_import_group_child_anchor_unsupported")
            style[margin] = vml_decimal(Fraction(int(offset_node.text), 12700)) + "pt"
            style[f"mso-position-{axis}-relative"] = position.get("relativeFrom")
    shape = copy.deepcopy(legacy)
    shape.set("style", ";".join(f"{key}:{value}" for key, value in style.items()))
    pict = legacy
    while pict.tag != Q(ns, "pict"):
        pict = parents[pict]
    types = {
        item.get("type")[1:] for item in shape.iter() if (item.get("type") or "").startswith("#")
    }
    new_pict = ET.Element(pict.tag, dict(pict.attrib))
    new_pict.extend(
        copy.deepcopy(item) for item in pict.iter(Q(_VML, "shapetype")) if item.get("id") in types
    )
    new_pict.append(shape)
    choice[:] = [new_drawing]
    fallback[:] = [new_pict]
    return alternate


class NodeImporter:
    """Import self-contained XML subtrees into one destination Document session.

    ``import_node`` returns a destination-owned detached node after transferring
    reachable definitions and resources. Insert it into a fresh destination
    snapshot. ``insert_node`` combines transfer and insertion atomically and
    applies context-sensitive spacing. Source nodes and packages never mutate.
    Reuse this instance across calls; every call reads the sessions' current bytes.

    Standalone notes and comments install as new entries and return live roots.
    ``import_story`` installs a header/footer in a one-based section slot; the
    detached ``import_node`` story contract is retained. Whole documents, partial
    semantic envelopes, active content and unestablished resources raise typed
    ``NodeEditError`` without changing either session. No resource is fetched.
    """

    def __init__(
        self,
        source_document: Document,
        destination_document: Document,
        format_mode: ImportFormatMode = ImportFormatMode.KEEP_SOURCE_FORMATTING,
        format_options: ImportFormatOptions = _DEFAULT_IMPORT_OPTIONS,
    ) -> None:
        from .document import Document
        from .model import NodeEditError

        if not isinstance(source_document, Document) or not isinstance(
            destination_document, Document
        ):
            raise NodeEditError("node_import_document_invalid")
        if not isinstance(format_options, ImportFormatOptions):
            raise NodeEditError("node_import_format_options_invalid")
        try:
            self.format_mode = _parse_import_mode(format_mode)
            self.format_options = _parse_import_options(
                {
                    name: getattr(format_options, name)
                    for name in ImportFormatOptions.__dataclass_fields__
                }
            )
        except ValueError as error:
            raise NodeEditError(str(error).replace("assembly_", "node_import_", 1)) from error
        self.source_document = source_document
        self.destination_document = destination_document
        self._numbering_source: str | None = None
        self._numbering_cache: dict[str, tuple[str, tuple[str, str]]] = {}
        self._numbering_destination_context: str | None = None

    def import_node(
        self,
        node: Node,
        deep: bool = True,
        *,
        destination_part: str = _DOCUMENT,
        parent: Node | None = None,
        index: int | None = None,
    ) -> Node:
        """Transfer a root; ``parent`` also inserts its standalone page boundary.

        Without a parent, tables cannot carry a preceding page boundary and
        refuse that option rather than return a different node type.
        """
        if parent is not None:
            return self.insert_node(node, parent, index, deep=deep)
        if index is not None:
            from .model import NodeEditError

            raise NodeEditError("node_import_parent_required")
        return self._import_node(node, deep, destination_part=destination_part, page_boundary=True)

    def import_story(
        self,
        node: Node,
        *,
        section: int = 1,
        variant: str = "default",
        deep: bool = True,
    ) -> Node:
        """Install a header/footer atomically; ``primary`` aliases ``default``."""
        from .model import Node, NodeEditError

        if not isinstance(node, Node) or node.xml_name not in {Q(_W, "hdr"), Q(_W, "ftr")}:
            raise NodeEditError("node_import_header_footer_root_invalid")
        destination = self.destination_document
        snapshot = destination._snapshot()
        cache = (
            self._numbering_source,
            dict(self._numbering_cache),
            self._numbering_destination_context,
        )
        try:
            imported = self.import_node(node, deep)
            return destination.insert_story(
                imported, section=section, variant="default" if variant == "primary" else variant
            )
        except (ValueError, PackageError, OSError):
            destination._restore(snapshot)
            self._numbering_source, self._numbering_cache, self._numbering_destination_context = (
                cache
            )
            raise

    def _import_node(
        self, node: Node, deep: bool, *, destination_part: str, page_boundary: bool
    ) -> Node:
        from .model import Node, NodeEditError, NodeType
        from .review import _CONFLICT_REVISION_TAGS, _all_revision_ids

        smartart = feature_module("smartart", required=False)
        import_errors = (ValueError, PackageError, ET.ParseError, KeyError, OSError, StopIteration)
        if smartart is not None:
            import_errors += (smartart.SmartArtError,)

        if not isinstance(node, Node) or node.document is not self.source_document:
            raise NodeEditError("node_import_source_invalid")
        if not isinstance(deep, bool):
            raise NodeEditError("node_import_depth_invalid")
        if node._element is None or node.node_type == NodeType.DOCUMENT:
            raise NodeEditError("node_import_kind_unsupported")
        entry_kind = node.xml_name.rsplit("}", 1)[-1]
        installed_entry = entry_kind in {"footnote", "endnote", "comment"}
        if entry_kind in {"document", "headerReference", "footerReference"}:
            raise NodeEditError("node_import_story_root_unsupported")
        if installed_entry and not deep:
            raise NodeEditError("node_import_story_depth_unsupported")
        if entry_kind in _NOTE_KINDS and node._element.get(Q(_W, "type"), "normal") != "normal":
            construct = node._element.get(Q(_W, "type"))
            raise NodeEditError(f"node_import_{entry_kind}_{construct}_unsupported")
        if node._edit is not None:
            node._edit("check", node, node, 0)
        memo: dict[int, Any] = {}
        prebuilt: ET.Element | None = None
        if _is_drawing_fragment(node, _W):
            if not deep:
                raise NodeEditError("node_import_drawing_fragment_depth_unsupported")
            try:
                fragment = _extract_drawing_fragment(node, _W)
            except ValueError as error:
                if isinstance(error, NodeEditError):
                    raise
                code = str(error)
                raise NodeEditError(
                    code if code.startswith("node_import_") else "node_import_package_invalid"
                ) from error
            if fragment is not None:
                prebuilt, memo, node = fragment
                entry_kind = "r"
        source = self.source_document
        destination = self.destination_document
        snapshot = destination._snapshot()
        try:
            source_parts = {name: source.part_bytes(name) for name in source.part_names}
            parts = {name: destination.part_bytes(name) for name in destination.part_names}
            baseline = dict(parts)
            signature = hashlib.sha256()
            signature.update(repr((self.format_mode, self.format_options)).encode("utf-8"))
            for name, data in sorted(source_parts.items()):
                signature.update(name.encode("utf-8"))
                signature.update(hashlib.sha256(data).digest())
            numbering_source = signature.hexdigest()
            cached_numbering = (
                self._numbering_cache
                if numbering_source == self._numbering_source
                and _import_definition_context(parts) == self._numbering_destination_context
                else {}
            )
            destination_numbering = (
                parse_xml_with_preserved_qnames(parts["word/numbering.xml"])
                if "word/numbering.xml" in parts
                else ET.Element(Q(_W, "numbering"))
            )
            if _DOCUMENT not in source_parts or _DOCUMENT not in parts:
                raise NodeEditError("node_import_main_part_unsupported")
            from .document import _main_part

            if (
                _main_part(source._graph) != _DOCUMENT
                or _main_part(destination._graph) != _DOCUMENT
            ):
                raise NodeEditError("node_import_main_part_unsupported")
            source_types = parse_xml_with_preserved_qnames(source_parts["[Content_Types].xml"])
            main_type = _content_type(_DOCUMENT, source_types)
            if (main_type is not None and "macroenabled" in main_type[1].casefold()) or any(
                item.get("Type", "").rsplit("/", 1)[-1]
                in {"vbaProject", "vbaData", "activeXControl", "activeXControlBinary"}
                for item in _read_rels(source_parts, _DOCUMENT)
            ):
                raise NodeEditError("node_import_active_content_unsupported")
            if destination_part not in parts or node.part not in source_parts:
                raise NodeEditError("node_import_part_invalid")
            destination_root = parse_xml_with_preserved_qnames(parts[_DOCUMENT])
            ns = _word_ns(destination_root)
            if ns != _W or _word_ns(parse_xml_with_preserved_qnames(source_parts[_DOCUMENT])) != ns:
                raise NodeEditError("node_import_namespace_unsupported")
            target_root = parse_xml_with_preserved_qnames(parts[destination_part])
            if _word_ns(target_root) != ns:
                raise NodeEditError("node_import_destination_story_unsupported")
            incoming = prebuilt if prebuilt is not None else copy.deepcopy(node._element, memo)
            if not deep:
                incoming[:] = []
                incoming.text = None
            options = self.format_options
            if deep:
                sections = {
                    memo[id(section)]: section for section in node._element.iter(Q(ns, "sectPr"))
                }
                _apply_section_policy(
                    sections,
                    parse_xml_with_preserved_qnames(source_parts[_DOCUMENT]),
                    ns,
                    SectionPolicy("source", "source"),
                )
            if options.ignore_text_boxes:
                _omit_import_text_boxes([incoming], ns)
            if options.ignore_header_footer:
                for section in incoming.iter(Q(ns, "sectPr")):
                    for item in list(section):
                        if item.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}:
                            section.remove(item)
            if deep and (
                self.format_mode != ImportFormatMode.USE_DESTINATION_STYLES
                or options.force_copy_styles
            ):
                _materialize_import_table_context(node, incoming, source_parts, ns, memo)
                # Pin mathematical defaults before inherited Word run fonts are
                # made explicit; otherwise a paragraph default masks mathFont.
                _reconcile_import_math(source_parts, parts, [incoming], ns)
                _pin_inline_import_context(node, incoming, source_parts, ns)
            related = _gather_related_roots(source_parts, [incoming], ns, node_import=True)
            roots = [incoming, *related.values()]
            notes = _collect_note_imports(
                parts,
                source_parts,
                roots,
                ns,
                standalone=incoming if entry_kind in _NOTE_KINDS else None,
            )
            roots.extend(item for item in notes.roots if item is not incoming)
            comments = _collect_comment_imports(
                parts,
                source_parts,
                roots,
                destination_root,
                ns,
                standalone=incoming if entry_kind == "comment" else None,
            )
            roots.extend(item for item in comments.roots if item is not incoming)
            if entry_kind in _NOTE_KINDS:
                source_entry_part, destination_part, _ = notes.parts[entry_kind]
                if source_entry_part != node.part:
                    raise NodeEditError("node_import_note_entry_invalid")
            elif entry_kind == "comment":
                if comments.source_part != node.part:
                    raise NodeEditError("node_import_comment_entry_invalid")
                destination_part = comments.part
            if installed_entry:
                target_root = parse_xml_with_preserved_qnames(parts[destination_part])
            _reconcile_import_math(source_parts, parts, roots, ns)
            if options.ignore_text_boxes:
                _omit_import_text_boxes(roots, ns)
            _validate_import_fragment(node, roots, ns)
            if any(
                item.tag in _CONFLICT_REVISION_TAGS for root in roots for item in root.iter()
            ) and _conflict_mode(source_parts) != _conflict_mode(parts):
                raise NodeEditError("node_import_conflict_mode_mismatch")
            if not deep or not any(
                root.tag == Q(ns, "p") or root.find(f".//{Q(ns, 'p')}") is not None
                for root in roots
            ):
                source_styles = _styles_root(source_parts)
                if source_styles is not None:
                    defaults = source_styles.find(Q(ns, "docDefaults"))
                    if defaults is not None:
                        source_styles.remove(defaults)
                        source_parts[_STYLES] = serialize_xml_with_preserved_qnames(source_styles)
            tokens = _field_style_tokens(roots, ns)
            owners = {id(incoming): node.part, **{id(root): part for part, root in related.items()}}
            for kind, (source_part, _, _) in notes.parts.items():
                owners.update(
                    {id(root): source_part for root in notes.roots if root.tag == Q(ns, kind)}
                )
            owners.update({id(root): comments.source_part for root in comments.roots})
            for root in roots:
                _resolve_theme_dependencies(
                    source_parts,
                    [root],
                    ns,
                    resolve_colors=options.resolve_theme_colors,
                    owner=owners.get(id(root), _DOCUMENT),
                )
            _prepare_import_theme(source_parts, roots, ns, tokens, options, pin_implicit=deep)
            previous = {
                item.get(Q(ns, "abstractNumId"))
                for item in parse_xml_with_preserved_qnames(
                    parts.get("word/numbering.xml", f'<w:numbering xmlns:w="{ns}"/>'.encode())
                ).findall(Q(ns, "abstractNum"))
            }
            reusable = {
                old: identifier
                for old, (identifier, identity) in cached_numbering.items()
                if _numbering_identity(destination_numbering, identifier, ns) == identity
            }
            numbers = _merge_numbering(
                parts, source_parts, roots, ns, tokens, options, reuse_numbers=reusable
            )
            numbering = (
                parse_xml_with_preserved_qnames(parts["word/numbering.xml"]) if numbers else None
            )
            numbering_styles = (
                [
                    item
                    for item in numbering.findall(Q(ns, "abstractNum"))
                    if item.get(Q(ns, "abstractNumId")) not in previous
                ]
                + [
                    item
                    for item in numbering.findall(Q(ns, "num"))
                    if item.get(Q(ns, "numId"))
                    not in {
                        row.get(Q(ns, "numId"))
                        for row in destination_numbering.findall(Q(ns, "num"))
                    }
                ]
                if numbering is not None
                else []
            )
            label = _xml_hash(incoming)[:12]
            styles = _merge_styles(
                parts,
                source_parts,
                [*roots, *numbering_styles],
                ns,
                label,
                numbers,
                tokens,
                self.format_mode,
                options,
                allow_style_resources=True,
            )
            if numbering is not None:
                parts["word/numbering.xml"] = serialize_xml_with_preserved_qnames(numbering)
            bookmark_names = {
                item.get(Q(ns, "name"))
                for root in roots
                for item in root.iter(Q(ns, "bookmarkStart"))
                if item.get(Q(ns, "name"))
            }
            starts = [
                item.get(Q(ns, "id"))
                for root in roots
                for item in root.iter(Q(ns, "bookmarkStart"))
            ]
            ends = [
                item.get(Q(ns, "id")) for root in roots for item in root.iter(Q(ns, "bookmarkEnd"))
            ]
            if sorted(starts, key=str) != sorted(ends, key=str):
                raise NodeEditError("node_import_bookmark_range_incomplete")
            bookmarks = _remap_bookmarks(parts, roots, ns)
            types = parse_xml_with_preserved_qnames(parts["[Content_Types].xml"])
            occupied_sdt = {
                int(value)
                for name, data in parts.items()
                if (
                    name.endswith((".xml", ".rels"))
                    or (
                        (kind := _content_type(name, types)) is not None
                        and kind[1].endswith("+xml")
                    )
                )
                for item in parse_xml_with_preserved_qnames(data).iter(Q(ns, "sdtPr"))
                for identity in item.findall(Q(ns, "id"))
                if (value := identity.get(Q(ns, "val"), "")).isdigit()
            }
            _remap_sdt_ids(roots, ns, occupied_sdt)
            permission_roots = {id(incoming): (destination_part, target_root)}
            permission_roots.update(
                {
                    id(root): (f"related:{part}", ET.Element(root.tag))
                    for part, root in related.items()
                }
            )
            for kind, (_, part, _) in notes.parts.items():
                existing = parse_xml_with_preserved_qnames(parts[part])
                permission_roots.update(
                    {id(root): (part, existing) for root in notes.roots if root.tag == Q(ns, kind)}
                )
            if comments.baseline is not None:
                permission_roots.update(
                    {id(root): (comments.part, comments.baseline) for root in comments.roots}
                )
            permission_ids = {}
            for root in roots:
                starts = list(root.iter(Q(ns, "permStart")))
                ends = list(root.iter(Q(ns, "permEnd")))
                identifiers = [item.get(Q(ns, "id")) for item in starts]
                if (
                    None in identifiers
                    or len(identifiers) != len(set(identifiers))
                    or sorted(identifiers, key=str)
                    != sorted((item.get(Q(ns, "id")) for item in ends), key=str)
                ):
                    raise NodeEditError("node_import_permission_range_incomplete")
                owner, existing = permission_roots[id(root)]
                occupied = permission_ids.setdefault(
                    owner, {item.get(Q(ns, "id")) for item in existing.iter(Q(ns, "permStart"))}
                )
                remapped = {}
                serial = (
                    max((int(value) for value in occupied if value and value.isdigit()), default=0)
                    + 1
                )
                for item in starts:
                    old = item.get(Q(ns, "id"))
                    while str(serial) in occupied:
                        serial += 1
                    remapped[old] = str(serial)
                    item.set(Q(ns, "id"), str(serial))
                    occupied.add(str(serial))
                    serial += 1
                for item in ends:
                    item.set(Q(ns, "id"), remapped[item.get(Q(ns, "id"))])
            if not installed_entry:
                _remap_paragraph_ids(
                    [incoming], ns, collect_existing_para_ids({destination_part: target_root})
                )
                _remap_drawing_ids([incoming], _drawing_ids(target_root))
                _remap_vml_ids([incoming], ns, _vml_ids(target_root))
                _remap_anchor_ids([incoming], _anchor_ids(target_root))
                _remap_conflict_ids(
                    [incoming],
                    ns,
                    _all_revision_ids({destination_part: parts[destination_part]}),
                    allow_revisions=True,
                )
            for root in related.values():
                _remap_conflict_ids([root], ns, set(), allow_revisions=True)
            for kind, (_, part, _) in notes.parts.items():
                existing = parse_xml_with_preserved_qnames(parts[part])
                imported = [item for item in notes.roots if item.tag == Q(ns, kind)]
                _remap_paragraph_ids(imported, ns, collect_existing_para_ids({part: existing}))
                _remap_drawing_ids(imported, _drawing_ids(existing))
                _remap_vml_ids(imported, ns, _vml_ids(existing))
                _remap_anchor_ids(imported, _anchor_ids(existing))
                _remap_conflict_ids(
                    imported, ns, _all_revision_ids({part: parts[part]}), allow_revisions=True
                )
            if comments.baseline is not None:
                _remap_comment_identities(
                    comments,
                    ns,
                    collect_existing_para_ids({destination_part: target_root}),
                    _all_revision_ids(
                        {comments.part: serialize_xml_with_preserved_qnames(comments.baseline)}
                    ),
                    allow_revisions=True,
                )
            _validate_and_rewrite_fields(
                roots,
                ns,
                FieldPolicy(False, False),
                bookmarks,
                bookmark_names,
                _field_style_names(source_parts, parts, ns, styles),
            )
            resource_importer = _NodeResourceImporter(parts, source_parts, label, source.path)
            resource_importer.resolve_colors = options.resolve_theme_colors
            resource_importer.word_namespace = ns
            current_source = PackageGraph.from_bytes(
                source._materialize(), path=source.path, limits=source._graph.limits
            )
            resource_importer.smartart_graph = current_source
            resource_importer.binding_graph = current_source
            resource_importer.copy_bindings(roots, ns)
            for part, root in related.items():
                source_parts[part] = serialize_xml_with_preserved_qnames(root)
            if not installed_entry:
                resource_importer.copy_references([incoming], node.part, destination_part)
            for kind, (source_part, part, _) in notes.parts.items():
                resource_importer.copy_references(
                    [item for item in notes.roots if item.tag == Q(ns, kind)], source_part, part
                )
            if comments.source_part is not None:
                resource_importer.copy_references(
                    comments.roots, comments.source_part, comments.part
                )
            resource_importer.finish()
            _write_note_imports(parts, notes)
            _write_comment_imports(parts, comments, ns)
            if options.append_with_new_page and page_boundary:
                if incoming.tag in {
                    Q(ns, name)
                    for name in (
                        "body",
                        "hdr",
                        "ftr",
                        "footnote",
                        "endnote",
                        "comment",
                        "txbxContent",
                    )
                }:
                    blocks = list(incoming)
                    _start_import_on_new_page(blocks, ns)
                    incoming[:] = blocks
                elif incoming.tag == Q(ns, "p"):
                    _start_import_on_new_page([incoming], ns)
                elif incoming.tag == Q(ns, "sdt"):
                    content = incoming.find(Q(ns, "sdtContent"))
                    if content is None or not any(
                        item.tag in {Q(ns, "p"), Q(ns, "tbl")} for item in content
                    ):
                        raise NodeEditError("node_import_page_boundary_sdt_requires_parent")
                    blocks = list(content)
                    _start_import_on_new_page(blocks, ns)
                    content[:] = blocks
                else:
                    raise NodeEditError(f"node_import_page_boundary_{entry_kind}_requires_parent")
                if installed_entry:
                    _write_note_imports(parts, notes)
                    _write_comment_imports(parts, comments, ns)
            if parts != baseline:
                destination._adopt(
                    "node-import-dependencies",
                    _import_session_bytes(destination, parts),
                    {
                        "source_part": node.part,
                        "destination_part": destination_part,
                        "format_mode": self.format_mode.value,
                    },
                )
            if installed_entry:
                imported = next(
                    item
                    for item in destination.node.children
                    if item.part == destination_part
                    and item.xml_name == incoming.tag
                    and item._element.get(Q(ns, "id")) == incoming.get(Q(ns, "id"))
                )
            else:
                imported = destination._bind_imported_node(
                    incoming, destination_part, node.node_type, node.story_type
                )
            cache = dict(cached_numbering)
            if numbers:
                final_numbering = parse_xml_with_preserved_qnames(parts["word/numbering.xml"])
                for old, identifier in numbers.items():
                    identity = _numbering_identity(final_numbering, identifier, ns)
                    if identity is not None:
                        cache[old] = (identifier, identity)
            self._numbering_source, self._numbering_cache = numbering_source, cache
            self._numbering_destination_context = _import_definition_context(parts)
            return imported
        except import_errors as error:
            destination._restore(snapshot)
            if isinstance(error, NodeEditError):
                raise
            code = getattr(error, "code", str(error))
            if code.startswith("assembly_"):
                code = code.replace("assembly_", "node_import_", 1)
            if not code.startswith("node_import_"):
                code = "node_import_package_invalid"
            raise NodeEditError(code) from error

    def insert_node(
        self, node: Node, parent: Node, index: int | None = None, *, deep: bool = True
    ) -> Node:
        """Import and insert at a visible child index, rolling back any refusal."""
        from .model import Node, NodeEditError, NodeType

        destination = self.destination_document
        if (
            not isinstance(parent, Node)
            or parent.document is not destination
            or not parent._live
            or parent.part is None
        ):
            raise NodeEditError("node_import_parent_invalid")
        if index is None:
            index = len(parent.children)
            if (
                parent.xml_name
                and parent.xml_name.endswith("}body")
                and parent.children
                and parent.children[-1].xml_name is not None
                and parent.children[-1].xml_name.endswith("}sectPr")
            ):
                index -= 1
        if (
            isinstance(index, bool)
            or not isinstance(index, int)
            or not 0 <= index <= len(parent.children)
        ):
            raise NodeEditError("node_index_invalid")
        parent._edit("check", parent, parent, 0)
        route = []
        cursor = parent
        while cursor.parent is not None and cursor.parent._element is not None:
            route.append(cursor._index)
            cursor = cursor.parent
        story_part, story_name, story_type = cursor.part, cursor.xml_name, cursor.story_type
        story_id = next(
            (value for key, value in cursor._element.attrib.items() if key.endswith("}id")), None
        )
        from .model import StoryType

        if (
            isinstance(node, Node)
            and node._element is not None
            and _is_drawing_fragment(node, _W)
            and parent.xml_name != Q(_W, "p")
        ):
            raise NodeEditError("node_import_drawing_fragment_requires_paragraph")
        sync_ordinal = None
        if parent.xml_name == Q(_W, "txbxContent"):
            sync_ordinal = next(
                (
                    position
                    for position, item in enumerate(cursor._element.iter(Q(_W, "txbxContent")))
                    if item is parent._element
                ),
                None,
            )
        snapshot = destination._snapshot()
        numbering_snapshot = (
            self._numbering_source,
            dict(self._numbering_cache),
            self._numbering_destination_context,
        )
        try:
            imported = self._import_node(
                node, deep, destination_part=parent.part, page_boundary=False
            )
            fresh_parent = next(
                item
                for item in destination.node.children
                if item.part == story_part
                and item.xml_name == story_name
                and item.story_type == story_type
                and next(
                    (value for key, value in item._element.attrib.items() if key.endswith("}id")),
                    None,
                )
                == story_id
            )
            for position in reversed(route):
                fresh_parent = fresh_parent.children[position]
            ns = _word_ns(parse_xml_with_preserved_qnames(destination.part_bytes(_DOCUMENT)))
            if self.format_options.adjust_sentence_and_word_spacing and not any(
                item.tag in {Q(ns, local) for local in ("p", "tbl")}
                for item in imported._element.iter()
            ):
                before = fresh_parent.children[index - 1]._element if index else None
                after = (
                    fresh_parent.children[index]._element
                    if index < len(fresh_parent.children)
                    else None
                )
                _adjust_sentence_word_spacing(imported._element, before, after, ns)
            if self.format_options.append_with_new_page:
                if imported.xml_name not in {Q(ns, "p"), Q(ns, "tbl"), Q(ns, "sdt")}:
                    if fresh_parent.xml_name == Q(ns, "p") and route:
                        from .document import _validate_node_structure

                        package_parts = {
                            name: destination.part_bytes(name) for name in destination.part_names
                        }
                        root = parse_xml_with_preserved_qnames(package_parts[story_part])
                        if story_name == Q(ns, "body"):
                            story = root.find(Q(ns, "body"))
                        elif root.tag == story_name:
                            story = root
                        else:
                            story = next(
                                item
                                for item in root
                                if item.tag == story_name and item.get(Q(ns, "id")) == story_id
                            )
                        container = story
                        for position in reversed(route[1:]):
                            container = container[position]
                        paragraph_position = route[0]
                        paragraph = container[paragraph_position]
                        following = (
                            fresh_parent.children[index]._element
                            if index < len(fresh_parent.children)
                            else None
                        )
                        xml_index = (
                            list(fresh_parent._element).index(following)
                            if following is not None
                            else len(paragraph)
                        )
                        suffix = list(paragraph)[xml_index:]
                        properties = paragraph.find(Q(ns, "pPr"))
                        new_paragraph = _shell(paragraph)
                        if properties is not None:
                            new_paragraph.append(copy.deepcopy(properties))
                            section = properties.find(Q(ns, "sectPr"))
                            if section is not None:
                                properties.remove(section)
                        new_paragraph.append(copy.deepcopy(imported._element))
                        new_paragraph.extend(item for item in suffix if item.tag != Q(ns, "pPr"))
                        for item in suffix:
                            if item.tag != Q(ns, "pPr"):
                                paragraph.remove(item)
                        _remap_paragraph_ids(
                            [new_paragraph], ns, collect_existing_para_ids({story_part: root})
                        )
                        _start_import_on_new_page([new_paragraph], ns)
                        container.insert(paragraph_position + 1, new_paragraph)
                        _validate_node_structure(root, ns)
                        package_parts[story_part] = serialize_xml_with_preserved_qnames(root)
                        destination._adopt(
                            "node-import-page-boundary",
                            _import_session_bytes(destination, package_parts),
                            {"part": story_part},
                        )
                        result_parent = next(
                            item
                            for item in destination.node.children
                            if item.part == story_part
                            and item.xml_name == story_name
                            and item._element.get(Q(ns, "id")) == story_id
                        )
                        for position in reversed(route[1:]):
                            result_parent = result_parent.children[position]
                        return result_parent.children[paragraph_position + 1].children[1]
                    local = imported.xml_name.rsplit("}", 1)[-1]
                    raise NodeEditError(f"node_import_page_boundary_{local}_requires_block_parent")
                blocks = [imported._element]
                _start_import_on_new_page(blocks, ns)
                if len(blocks) == 2:
                    boundary = destination._bind_imported_node(
                        blocks[0], parent.part, NodeType.PARAGRAPH, fresh_parent.story_type
                    )
                    fresh_parent.insert_child(index, boundary)
                    index += 1
                else:
                    imported = destination._bind_imported_node(
                        blocks[0], parent.part, imported.node_type, imported.story_type
                    )
            inserted = fresh_parent.insert_child(index, imported)
            if sync_ordinal is None:
                return inserted
            package_parts = {name: destination.part_bytes(name) for name in destination.part_names}
            root = parse_xml_with_preserved_qnames(package_parts[story_part])
            if story_name == Q(ns, "body"):
                story = root.find(Q(ns, "body"))
            elif root.tag == story_name:
                story = root
            else:
                story = next(
                    item
                    for item in root
                    if item.tag == story_name and item.get(Q(ns, "id")) == story_id
                )
            boxes = list(story.iter(Q(ns, "txbxContent")))
            if not _sync_text_box_pair(boxes[sync_ordinal], story, ns):
                return inserted
            package_parts[story_part] = serialize_xml_with_preserved_qnames(root)
            destination._adopt(
                "node-import-text-box-sync",
                _import_session_bytes(destination, package_parts),
                {"part": story_part},
            )
            fresh_story = next(
                item
                for item in destination.node.children
                if item.part == story_part
                and item.xml_name == story_name
                and item.story_type == story_type
                and next(
                    (value for key, value in item._element.attrib.items() if key.endswith("}id")),
                    None,
                )
                == story_id
            )
            box = [
                item
                for item in fresh_story.get_child_nodes(NodeType.STORY, True)
                if item.story_type == StoryType.TEXT_BOX
            ][sync_ordinal]
            return box.children[index]
        except (ValueError, PackageError, IndexError, StopIteration, OSError) as error:
            destination._restore(snapshot)
            self._numbering_source, self._numbering_cache, self._numbering_destination_context = (
                numbering_snapshot
            )
            if isinstance(error, NodeEditError):
                raise
            raise NodeEditError("node_import_insertion_invalid") from error


def assemble_document(request: AssemblyRequest) -> AssemblyReport:
    from .review import _CONFLICT_REVISION_TAGS, _all_revision_ids

    content_controls = feature_module("content_controls")
    smartart = feature_module("smartart")

    try:
        _parse_import_mode(request.import_format_mode)
        _parse_import_options(
            {
                name: getattr(request.import_format_options, name)
                for name in ImportFormatOptions.__dataclass_fields__
            }
        )
    except (AttributeError, ValueError):
        return _refuse(request, "assembly_import_format_options_invalid")
    destination, output = request.destination.resolve(), request.output.resolve()
    if destination == output or not destination.is_file() or not output.parent.is_dir():
        return _refuse(request, "assembly_destination_invalid")
    if _sha256(destination) != request.destination_sha256:
        return _refuse(request, "assembly_destination_hash_mismatch")
    if any(
        not source.path.is_file() or _sha256(source.path) != source.sha256
        for source in request.sources
    ):
        return _refuse(request, "assembly_source_hash_mismatch")
    for admitted in (destination, *(source.path for source in request.sources)):
        refusal = _package_limit_refusal(admitted)
        if refusal is not None:
            return _refuse(request, refusal)
    try:
        with zipfile.ZipFile(destination) as archive:
            parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        baseline_parts = dict(parts)
        destination_root = parse_xml_with_preserved_qnames(parts[_DOCUMENT])
        ns = _word_ns(destination_root)
        if ns is None:
            return _refuse(request, "assembly_destination_namespace_unsupported")
        body = destination_root.find(Q(ns, "body"))
        assert body is not None
        existing = _body_items(destination_root, ns)
        destination_diagrams = list(destination_root.iter(_DIAGRAM_REL_IDS))
        if destination_diagrams:
            destination_graph = PackageGraph.from_file(destination)
            for rel_ids in destination_diagrams:
                smartart._resolve_smartart_parts(destination_graph, _DOCUMENT, rel_ids)
        if request.insert_after > len(existing) or (
            request.insert_after
            and _xml_hash(existing[request.insert_after - 1]) != request.anchor_fingerprint_sha256
        ):
            return _refuse(request, "assembly_anchor_stale")
        insertion = len(body) - 1 if body[-1].tag == Q(ns, "sectPr") else len(body)
        if request.insert_after:
            anchor = existing[request.insert_after - 1]
            insertion = next(index for index, child in enumerate(body) if child is anchor) + 1
        inserted: list[ET.Element] = []
        source_reports: list[dict[str, Any]] = []
        pinned_destination_section = False
        paragraph_ids = collect_existing_para_ids({_DOCUMENT: destination_root})
        drawing_ids = _drawing_ids(destination_root)
        vml_ids = _vml_ids(destination_root)
        anchor_ids = _anchor_ids(destination_root)
        destination_conflict_mode: bool | None = None
        revision_ids = _all_revision_ids({_DOCUMENT: parts[_DOCUMENT]})
        sdt_ids = content_controls._scan_package_sdt_ids(PackageGraph.from_file(destination), ns)
        for source in request.sources:
            with zipfile.ZipFile(source.path) as archive:
                source_parts = {
                    info.filename: archive.read(info.filename) for info in archive.infolist()
                }
            source_root = parse_xml_with_preserved_qnames(source_parts[_DOCUMENT])
            source_ns = _word_ns(source_root)
            strict_import = source_ns == _STRICT_W and ns == _W
            if source_ns != ns and not strict_import:
                return _refuse(request, "assembly_source_namespace_unsupported")
            source_fingerprints: list[str] | None = None
            source_section_fingerprint: str | None = None
            if strict_import:
                # Bind selections before any vocabulary, border or conformance edit.
                source_fingerprints = [
                    _xml_hash(item) for item in _assembly_items(source_root, source_ns)
                ]
                source_body = source_root.find(Q(source_ns, "body"))
                source_section = (
                    source_body.find(Q(source_ns, "sectPr")) if source_body is not None else None
                )
                if source_section is not None:
                    source_section_fingerprint = _xml_hash(source_section)
                source_parts = _transitional_assembly_parts(source_parts)
                source_root = parse_xml_with_preserved_qnames(source_parts[_DOCUMENT])
            items = _assembly_items(source_root, ns)
            selected: list[ET.Element] = []
            selected_sections: dict[ET.Element, ET.Element] = {}
            for selection in source.selections:
                if selection.start == 0:
                    source_body = source_root.find(Q(ns, "body"))
                    source_section = (
                        source_body.find(Q(ns, "sectPr")) if source_body is not None else None
                    )
                    if (
                        source_section is None
                        or (
                            source_section_fingerprint
                            if strict_import
                            else _xml_hash(source_section)
                        )
                        != selection.structure_fingerprint_sha256
                    ):
                        return _refuse(request, "assembly_source_target_stale")
                    schedule = ET.Element(Q(ns, "p"))
                    properties = ET.SubElement(schedule, Q(ns, "pPr"))
                    imported_section = copy.deepcopy(source_section)
                    properties.append(imported_section)
                    selected_sections[imported_section] = source_section
                    selected.append(schedule)
                    continue
                candidate = items[selection.start - 1 : selection.end]
                if len(candidate) != selection.end - selection.start + 1:
                    return _refuse(request, "assembly_source_target_stale")
                fingerprint = assembly_range_fingerprint(
                    source_fingerprints[selection.start - 1 : selection.end]
                    if source_fingerprints is not None
                    else [_xml_hash(item) for item in candidate]
                )
                if fingerprint != selection.structure_fingerprint_sha256:
                    return _refuse(request, "assembly_source_target_stale")
                for item in candidate:
                    clone_memo: dict[int, Any] = {}
                    imported = copy.deepcopy(item, clone_memo)
                    source_section = item.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
                    if source_section is not None:
                        selected_sections[clone_memo[id(source_section)]] = source_section
                    selected.append(imported)
            if request.import_format_options.ignore_text_boxes:
                _omit_import_text_boxes(selected, ns)
            if _unsupported(selected, ns):
                return _refuse(request, "assembly_structure_unsupported")
            _apply_section_policy(selected_sections, source_root, ns, request.section_policy)
            if request.import_format_options.ignore_header_footer:
                for section in selected_sections:
                    for reference in list(section):
                        if reference.tag in {
                            Q(ns, "headerReference"),
                            Q(ns, "footerReference"),
                        }:
                            section.remove(reference)
            related_roots = _gather_related_roots(source_parts, selected, ns)
            imported_roots = [*selected, *related_roots.values()]
            note_imports = _collect_note_imports(parts, source_parts, imported_roots, ns)
            imported_roots.extend(note_imports.roots)
            if request.import_format_options.ignore_text_boxes:
                _omit_import_text_boxes(imported_roots, ns)
            comment_imports = _collect_comment_imports(
                parts, source_parts, imported_roots, destination_root, ns
            )
            imported_roots.extend(comment_imports.roots)
            if request.import_format_options.ignore_text_boxes:
                _omit_import_text_boxes(comment_imports.roots, ns)
            if _unsupported(imported_roots, ns):
                return _refuse(request, "assembly_structure_unsupported")
            if any(
                node.tag in _CONFLICT_REVISION_TAGS
                for root in imported_roots
                for node in root.iter()
            ):
                if destination_conflict_mode is None:
                    destination_conflict_mode = _conflict_mode(baseline_parts)
                if _conflict_mode(source_parts) != destination_conflict_mode:
                    return _refuse(request, "assembly_conflict_mode_mismatch")
            field_style_tokens = _field_style_tokens(imported_roots, ns)
            _prepare_import_theme(
                source_parts,
                imported_roots,
                ns,
                field_style_tokens,
                request.import_format_options,
                # Assembly retains implicit styles, including synthetic section
                # carriers. Only isolated node imports pin their source context.
                pin_implicit=False,
            )
            source_section = next(
                (
                    section
                    for block in reversed(selected)
                    if (section := block.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")) is not None
                ),
                None,
            )
            if not pinned_destination_section and source_section is not None:
                _pin_destination_following_section(
                    destination_root,
                    parts,
                    ns,
                    insertion,
                    _section_reference_keys(source_section, ns),
                )
                insertion += _pin_destination_preceding_section(destination_root, ns, insertion)
                pinned_destination_section = True
            label = hashlib.sha256(source.sha256.encode()).hexdigest()[:12]
            previous_abstracts = {
                node.get(Q(ns, "abstractNumId"))
                for node in parse_xml_with_preserved_qnames(
                    parts.get("word/numbering.xml", f'<w:numbering xmlns:w="{ns}"/>'.encode())
                ).findall(Q(ns, "abstractNum"))
            }
            num_remap = _merge_numbering(
                parts,
                source_parts,
                imported_roots,
                ns,
                field_style_tokens,
                request.import_format_options,
            )
            imported_numbering = (
                parse_xml_with_preserved_qnames(parts["word/numbering.xml"]) if num_remap else None
            )
            numbering_styles = (
                [
                    node
                    for node in imported_numbering.findall(Q(ns, "abstractNum"))
                    if node.get(Q(ns, "abstractNumId")) not in previous_abstracts
                ]
                if imported_numbering is not None
                else []
            )
            style_remap = _merge_styles(
                parts,
                source_parts,
                [*imported_roots, *numbering_styles],
                ns,
                label,
                num_remap,
                field_style_tokens,
                request.import_format_mode,
                request.import_format_options,
            )
            if imported_numbering is not None and numbering_styles:
                parts["word/numbering.xml"] = serialize_xml_with_preserved_qnames(
                    imported_numbering
                )
            source_bookmarks = {
                name
                for root in imported_roots
                for node in root.iter(Q(ns, "bookmarkStart"))
                if (name := node.get(Q(ns, "name"))) is not None
            }
            # Earlier sources are already in the live body, not yet in parts.
            bookmark_names = _remap_bookmarks(
                parts, imported_roots, ns, occupied=[destination_root]
            )
            _remap_paragraph_ids(selected, ns, paragraph_ids)
            _remap_drawing_ids(selected, drawing_ids)
            _remap_vml_ids(selected, ns, vml_ids)
            _remap_anchor_ids(selected, anchor_ids)
            _remap_conflict_ids(selected, ns, revision_ids)
            _remap_sdt_ids(imported_roots, ns, sdt_ids)
            for root in related_roots.values():
                _remap_conflict_ids([root], ns, set())
            for kind, (_, destination_part, _) in note_imports.parts.items():
                note_destination_root = ET.fromstring(parts[destination_part])
                occupied = collect_existing_para_ids({destination_part: note_destination_root})
                _remap_paragraph_ids(
                    [note for note in note_imports.roots if note.tag == Q(ns, kind)],
                    ns,
                    occupied,
                )
                _remap_drawing_ids(
                    [note for note in note_imports.roots if note.tag == Q(ns, kind)],
                    _drawing_ids(note_destination_root),
                )
                _remap_vml_ids(
                    [note for note in note_imports.roots if note.tag == Q(ns, kind)],
                    ns,
                    _vml_ids(note_destination_root),
                )
                _remap_anchor_ids(
                    [note for note in note_imports.roots if note.tag == Q(ns, kind)],
                    _anchor_ids(note_destination_root),
                )
                _remap_conflict_ids(
                    [note for note in note_imports.roots if note.tag == Q(ns, kind)],
                    ns,
                    _all_revision_ids({destination_part: parts[destination_part]}),
                )
            if comment_imports.baseline is not None:
                _remap_comment_identities(
                    comment_imports,
                    ns,
                    paragraph_ids,
                    _all_revision_ids(
                        {
                            comment_imports.part: serialize_xml_with_preserved_qnames(
                                comment_imports.baseline
                            )
                        }
                    ),
                )
            _validate_and_rewrite_fields(
                imported_roots,
                ns,
                request.field_policy,
                bookmark_names,
                source_bookmarks,
                _field_style_names(source_parts, parts, ns, style_remap),
            )
            importer = _Importer(parts, source_parts, label, source.path)
            if strict_import and any(
                node.tag in {_DIAGRAM_REL_IDS, Q(ns, "dataBinding"), Q(_W15, "dataBinding")}
                for root in imported_roots
                for node in root.iter()
            ):
                # Graph-based dependency validation must see the same translated
                # vocabulary as the selected nodes, never the original on-disk XML.
                buffer = BytesIO()
                with zipfile.ZipFile(buffer, "w") as archive:
                    for name, data in source_parts.items():
                        archive.writestr(name, data)
                graph = PackageGraph.from_bytes(buffer.getvalue(), path=source.path)
                importer.smartart_graph = graph
                importer.binding_graph = graph
            _import_web_add_ins(importer)
            # Bound controls in notes and linked headers name the same source stores.
            importer.copy_bindings(imported_roots, ns)
            for part, root in related_roots.items():
                source_parts[part] = serialize_xml_with_preserved_qnames(root)
            importer.copy_references(selected, _DOCUMENT)
            for note in note_imports.roots:
                note_kind = note.tag.rsplit("}", 1)[-1]
                source_part, destination_part, _ = note_imports.parts[note_kind]
                importer.copy_references([note], source_part, destination_part)
            if comment_imports.source_part is not None:
                importer.copy_references(
                    comment_imports.roots, comment_imports.source_part, comment_imports.part
                )
            importer.finish()
            _write_note_imports(parts, note_imports)
            _write_comment_imports(parts, comment_imports, ns)
            if request.field_policy.request_on_open:
                _request_fields_on_open(parts, ns)
            if request.import_format_options.append_with_new_page and insertion:
                _start_import_on_new_page(selected, ns)
            for block in selected:
                body.insert(insertion, block)
                insertion += 1
                inserted.append(block)
            source_reports.append(
                {"source_sha256": source.sha256, "selected_blocks": len(selected)}
            )
        if inserted and insertion == len(body) - 1 and body[-1].tag == Q(ns, "sectPr"):
            final_properties = inserted[-1].find(Q(ns, "pPr"))
            final_section = (
                final_properties.find(Q(ns, "sectPr")) if final_properties is not None else None
            )
            if final_section is not None:
                final_properties.remove(final_section)
                body.remove(body[-1])
                body.append(final_section)
        parts[_DOCUMENT] = serialize_xml_with_preserved_qnames(destination_root)
        changed_parts = {
            name for name, contents in parts.items() if baseline_parts.get(name) != contents
        }
        staged: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                dir=output.parent,
                prefix=f".{output.name}.",
                suffix=".tmp",
                delete=False,
            ) as file:
                staged = Path(file.name)
            _write(staged, parts, destination)
            candidate = inspect_docx(staged)
            source_still_bound = all(
                source.path.is_file() and _sha256(source.path) == source.sha256
                for source in request.sources
            )
            unselected_parts_preserved = _unselected_destination_parts_preserved(
                baseline_parts, staged, changed_parts
            )
            if not (
                candidate.valid
                and _sha256(destination) == request.destination_sha256
                and source_still_bound
                and unselected_parts_preserved
            ):
                return _refuse(
                    request,
                    "assembly_staged_verification_failed"
                    if source_still_bound
                    else "assembly_source_hash_changed",
                )
            invalidated_signatures = feature_module("package_signatures")._invalidated_signatures(
                PackageGraph.from_file(destination), staged
            )
            os.replace(staged, output)
            staged = None
            return AssemblyReport(
                "published",
                None,
                request.destination_sha256,
                candidate.file_sha256,
                len(inserted),
                source_reports,
                True,
                invalidated_signatures,
            )
        finally:
            if staged is not None:
                staged.unlink(missing_ok=True)
    except smartart.SmartArtError as error:
        return _refuse(request, f"assembly_{error.code}")
    except PackageError:
        return _refuse(request, "assembly_package_invalid")
    except (OSError, KeyError, ET.ParseError, ValueError, zipfile.BadZipFile) as error:
        code = str(error) if str(error).startswith("assembly_") else "assembly_package_invalid"
        return _refuse(request, code)


def _processor_context(context: ProcessorContext | None) -> ProcessorContext:
    from .load_options import ProcessorContext

    if context is None:
        return ProcessorContext()
    if not isinstance(context, ProcessorContext):
        raise TypeError("processor_context_invalid")
    return context


def _processor_package_snapshot(source: Path, snapshot: Path, context: ProcessorContext) -> str:
    """Admit a private package snapshot while binding the original source bytes."""
    from .load_options import load_package

    data = source.read_bytes()
    identity = hashlib.sha256(data).hexdigest()
    loaded = load_package(data, path=source, options=context.load_options)
    admitted = loaded.data
    snapshot.write_bytes(admitted.read_bytes() if isinstance(admitted, Path) else admitted)
    return identity


def _relocate_body_bookmark_markers(body: ET.Element, ns: str) -> int:
    """Move body-level bookmark boundaries of a merged source into its paragraphs.

    A boundary between two body blocks marks the same character position as the
    start of the next block's first paragraph, so it moves there in order. A
    boundary after the last block moves to the end of that block's last
    paragraph, which leaves the final paragraph mark outside the range. Other
    body-level range markup, and a boundary next to anything but a paragraph or
    table, still refuses the merge.
    """
    markers = {Q(ns, "bookmarkStart"), Q(ns, "bookmarkEnd")}
    blocks = {Q(ns, "p"), Q(ns, "tbl")}
    paragraph_tag = Q(ns, "p")
    pending: list[ET.Element] = []
    previous: ET.Element | None = None
    moved = 0
    for child in list(body):
        if child.tag in markers:
            pending.append(child)
            continue
        if child.tag == Q(ns, "sectPr"):
            break
        if pending:
            paragraph = next(child.iter(paragraph_tag), None) if child.tag in blocks else None
            if paragraph is None:
                raise ValueError("assembly_merge_body_boundary_unsupported")
            index = 1 if len(paragraph) and paragraph[0].tag == Q(ns, "pPr") else 0
            for marker in pending:
                body.remove(marker)
                paragraph.insert(index, marker)
                index += 1
            moved += len(pending)
            pending = []
        previous = child
    if pending:
        if previous is None or previous.tag not in blocks:
            raise ValueError("assembly_merge_body_boundary_unsupported")
        paragraph = list(previous.iter(paragraph_tag))[-1]
        for marker in pending:
            body.remove(marker)
            paragraph.append(marker)
        moved += len(pending)
    return moved


def _drop_absent_style_references(styles: ET.Element, ns: str) -> int:
    """Remove basedOn, next and link elements naming no style in the same part.

    ECMA-376 Part 1 sections 17.7.4.3, 17.7.4.6 and 17.7.4.10 ignore such a
    reference, so removing it keeps the source's meaning. Kept, it would refuse
    the dependency walk or resolve to an unrelated destination style.
    """
    definitions = styles.findall(Q(ns, "style"))
    identifiers = {style.get(Q(ns, "styleId")) for style in definitions}
    dropped = 0
    for style in definitions:
        for name in ("basedOn", "next", "link"):
            for reference in style.findall(Q(ns, name)):
                if reference.get(Q(ns, "val")) not in identifiers:
                    style.remove(reference)
                    dropped += 1
    return dropped


def merge_docx(
    sources: Sequence[str | Path],
    output: str | Path,
    *,
    format_mode: ImportFormatMode = ImportFormatMode.KEEP_SOURCE_FORMATTING,
    format_options: ImportFormatOptions = _DEFAULT_IMPORT_OPTIONS,
    context: ProcessorContext | None = None,
) -> AssemblyReport:
    """Merge complete packages in order, using the first as the destination.

    All sources are admitted using ``context.load_options`` and bound to their
    original bytes. The existing assembly importer handles formatting, dependency
    relocation and sections in one transaction. Publication never replaces a
    destination, including one created concurrently. Body-level bookmark
    boundaries of an appended source move into the adjacent paragraph at the
    same character position, and style basedOn, next and link references to a
    style the source does not define are dropped as the specification ignores
    them. Each source receipt counts both under
    ``relocated_body_bookmark_markers`` and ``dropped_style_references``. Other
    body-level content without an assembly target refuses rather than
    disappearing from the merge. A single source is published unchanged after
    admission.

    Example::

        merge_docx(["intro.docx", "body.docx"], "book.docx",
                   format_mode=ImportFormatMode.KEEP_DIFFERENT_STYLES)
    """
    from .load_options import LoadError
    from .review import _link_no_clobber

    destination_hash = ""

    def refuse(code: str) -> AssemblyReport:
        return AssemblyReport("refused", code, destination_hash, None, 0, [])

    if isinstance(sources, (str, bytes, Path)) or not isinstance(sources, Sequence) or not sources:
        return refuse("assembly_sources_invalid")
    try:
        paths = tuple(Path(source).absolute() for source in sources)
        target = Path(output).absolute()
        mode = _parse_import_mode(format_mode)
        if not isinstance(format_options, ImportFormatOptions):
            return refuse("assembly_import_format_options_invalid")
        _parse_import_options(
            {
                name: getattr(format_options, name)
                for name in ImportFormatOptions.__dataclass_fields__
            }
        )
        settings = _processor_context(context)
        if target.exists() or target.is_symlink():
            return refuse("assembly_destination_exists")
        if not target.parent.is_dir():
            return refuse("assembly_destination_invalid")
        with tempfile.TemporaryDirectory(prefix=".wolfdocx-merge-", dir=target.parent) as directory:
            staging = Path(directory)
            snapshots: list[Path] = []
            identities: list[str] = []
            for index, path in enumerate(paths):
                snapshot = staging / f"source-{index:03d}.docx"
                identities.append(_processor_package_snapshot(path, snapshot, settings))
                snapshots.append(snapshot)
                if index == 0:
                    destination_hash = identities[0]
            imported: list[AssemblySource] = []
            normalized: list[tuple[int, int]] = []
            for snapshot in snapshots[1:]:
                with zipfile.ZipFile(snapshot) as archive:
                    root = parse_xml_with_preserved_qnames(archive.read(_DOCUMENT))
                ns = _word_ns(root)
                if ns is None:
                    return refuse("assembly_source_namespace_unsupported")
                body = root.find(Q(ns, "body"))
                if body is None:
                    return refuse("assembly_source_body_missing")
                rewritten: dict[str, bytes] = {}
                moved = _relocate_body_bookmark_markers(body, ns)
                if moved:
                    rewritten[_DOCUMENT] = serialize_xml_with_preserved_qnames(root)
                dropped = 0
                with zipfile.ZipFile(snapshot) as archive:
                    style_bytes = archive.read(_STYLES) if _STYLES in archive.namelist() else None
                if style_bytes is not None:
                    styles = parse_xml_with_preserved_qnames(style_bytes)
                    dropped = _drop_absent_style_references(styles, ns)
                    if dropped:
                        rewritten[_STYLES] = serialize_xml_with_preserved_qnames(styles)
                normalized.append((moved, dropped))
                if rewritten:
                    bound = snapshot
                    snapshot = bound.with_name(f"{bound.stem}-normalized.docx")
                    _write(snapshot, rewritten, bound)
                    # Bind selections to the bytes the importer will parse.
                    with zipfile.ZipFile(snapshot) as archive:
                        root = parse_xml_with_preserved_qnames(archive.read(_DOCUMENT))
                    body = root.find(Q(ns, "body"))
                    assert body is not None
                items = _assembly_items(root, ns)
                positions = {id(item): index for index, item in enumerate(items, start=1)}
                selections: list[AssemblySelection] = []
                for child in body:
                    if child.tag == Q(ns, "sectPr"):
                        selections.append(AssemblySelection(0, 0, _xml_hash(child)))
                    elif id(child) in positions:
                        position = positions[id(child)]
                        selections.append(AssemblySelection(position, position, _xml_hash(child)))
                    else:
                        return refuse("assembly_merge_body_boundary_unsupported")
                imported.append(AssemblySource(snapshot, _sha256(snapshot), tuple(selections)))
            staged_output = staging / "merged.docx"
            if imported:
                with zipfile.ZipFile(snapshots[0]) as archive:
                    root = parse_xml_with_preserved_qnames(archive.read(_DOCUMENT))
                ns = _word_ns(root)
                if ns is None:
                    return refuse("assembly_destination_namespace_unsupported")
                request = AssemblyRequest(
                    snapshots[0],
                    _sha256(snapshots[0]),
                    staged_output,
                    0,
                    None,
                    tuple(imported),
                    SectionPolicy("source", "source"),
                    FieldPolicy(False, False),
                    mode,
                    format_options,
                )
                report = assemble_document(request)
                if not report.passed:
                    return replace(report, destination_sha256=destination_hash)
            else:
                staged_output.write_bytes(snapshots[0].read_bytes())
                report = AssemblyReport(
                    "published", None, destination_hash, _sha256(staged_output), 0, [], True
                )
                if not inspect_docx(staged_output).valid:
                    return refuse("assembly_staged_verification_failed")
            if any(
                not path.is_file() or _sha256(path) != identity
                for path, identity in zip(paths, identities, strict=True)
            ):
                return refuse("assembly_source_hash_changed")
            code = _link_no_clobber(
                staged_output,
                target,
                exists_code="assembly_destination_exists",
                unavailable_code="assembly_publication_unavailable",
            )
            if code is not None:
                return refuse(code)
            return replace(
                report,
                destination_sha256=destination_hash,
                sources=[
                    {
                        **receipt,
                        "source_sha256": identity,
                        "relocated_body_bookmark_markers": moved,
                        "dropped_style_references": dropped,
                    }
                    for receipt, identity, (moved, dropped) in zip(
                        report.sources, identities[1:], normalized, strict=True
                    )
                ],
            )
    except LoadError as error:
        return refuse(error.code)
    except (OSError, TypeError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as error:
        code = str(error)
        return refuse(
            code if code.startswith(("assembly_", "processor_")) else "assembly_package_invalid"
        )


@dataclass(frozen=True)
class CleanupOptions:
    """Clean definitions and containers without changing retained content.

    ``remove_duplicate_referenced_styles`` opts into coalescing custom duplicates
    and rewriting their package-wide references, including STYLEREF display-name
    targets. Default, latent and relationship-bearing duplicates coalesce when every
    dependency is equivalent: the default flag survives on the retained style,
    latent exceptions carry identical settings, and relationship ids resolve to the
    same target. Anything else stays protected.
    ``remove_unused_picture_bullets`` also prunes orphaned image relationships
    and media, preserving images referenced from any retained part.
    ``remove_empty_containers`` applies across stories and cells while retaining
    section breaks, formatting, whitespace and required terminal paragraphs.
    """

    remove_unused_styles: bool = True
    remove_duplicate_styles: bool = True
    remove_unused_lists: bool = True
    remove_empty_containers: bool = False
    remove_unused_builtin_styles: bool = False
    remove_duplicate_referenced_styles: bool = False
    remove_unused_picture_bullets: bool = False

    def __post_init__(self) -> None:
        if any(not isinstance(getattr(self, name), bool) for name in self.__dataclass_fields__):
            raise ValueError("assembly_cleanup_option_invalid")


@dataclass(frozen=True)
class CleanupReport:
    status: str
    code: str | None
    removed_styles: tuple[str, ...] = ()
    removed_lists: tuple[str, ...] = ()
    duplicate_styles: tuple[tuple[str, str], ...] = ()
    output_sha256: str | None = None
    removed_containers: int = 0

    @property
    def passed(self) -> bool:
        return self.status == "published"

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "code": self.code,
            "removed_styles": list(self.removed_styles),
            "removed_lists": list(self.removed_lists),
            "duplicate_styles": dict(self.duplicate_styles),
            "output_sha256": self.output_sha256,
            "removed_containers": self.removed_containers,
        }


def cleanup_document(
    source: str | Path,
    output: str | Path,
    *,
    source_sha256: str,
    remove_empty_containers: bool | None = None,
    options: CleanupOptions | None = None,
    remove_unused_styles: bool | None = None,
    remove_duplicate_styles: bool | None = None,
    remove_unused_lists: bool | None = None,
    remove_unused_builtin_styles: bool | None = None,
    remove_duplicate_referenced_styles: bool | None = None,
    remove_unused_picture_bullets: bool | None = None,
) -> CleanupReport:
    """Atomically prune custom definitions, protecting every package-wide reference.

    Default, latent, field-target and relationship-bearing styles stay protected
    unless an equivalent duplicate coalesces under
    ``remove_duplicate_referenced_styles``, which rewrites every package-wide
    style, dependency and STYLEREF reference. Picture bullets and
    their otherwise unreferenced image resources have an independent opt-in.
    Empty paragraphs and wrappers are pruned across stories and cells, preserving
    section breaks, formatting, whitespace, anchors, drawings,
    and required terminal cell/story paragraphs remain intact.
    """
    overrides = {
        "remove_empty_containers": remove_empty_containers,
        "remove_unused_styles": remove_unused_styles,
        "remove_duplicate_styles": remove_duplicate_styles,
        "remove_unused_lists": remove_unused_lists,
        "remove_unused_builtin_styles": remove_unused_builtin_styles,
        "remove_duplicate_referenced_styles": remove_duplicate_referenced_styles,
        "remove_unused_picture_bullets": remove_unused_picture_bullets,
    }
    if options is not None and not isinstance(options, CleanupOptions):
        return CleanupReport("refused", "assembly_cleanup_option_invalid")
    if any(value is not None and not isinstance(value, bool) for value in overrides.values()):
        return CleanupReport("refused", "assembly_cleanup_option_invalid")
    policies = replace(
        options or CleanupOptions(),
        **{name: value for name, value in overrides.items() if value is not None},
    )
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or not source.is_file() or not output.parent.is_dir():
        return CleanupReport("refused", "assembly_destination_invalid")
    if _sha256(source) != source_sha256:
        return CleanupReport("refused", "assembly_source_hash_mismatch")
    staged = None
    try:
        with zipfile.ZipFile(source) as archive:
            parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        baseline = dict(parts)
        document = parse_xml_with_preserved_qnames(parts[_DOCUMENT])
        ns = _word_ns(document)
        if ns is None:
            return CleanupReport("refused", "assembly_source_namespace_unsupported")
        content_types = parse_xml_with_preserved_qnames(parts["[Content_Types].xml"])
        roots = {
            name: parse_xml_with_preserved_qnames(data)
            for name, data in parts.items()
            if name.endswith(".xml")
            or (
                (content_type := _content_type(name, content_types)) is not None
                and (
                    content_type[1].endswith("+xml")
                    or content_type[1] in {"application/xml", "text/xml"}
                )
            )
        }
        removed_containers = 0
        if policies.remove_empty_containers:

            def empty_paragraph(node: ET.Element) -> bool:
                allowed = {Q(ns, local) for local in ("p", "pPr", "r", "rPr", "t")}
                return all(
                    item.tag in allowed and not item.attrib and not item.text
                    for item in node.iter()
                )

            def terminal_paragraph(parent: ET.Element) -> ET.Element | None:
                for child in reversed(list(parent)):
                    if child.tag == Q(ns, "p"):
                        return child
                    if child.tag in {Q(ns, "sdt"), Q(ns, "customXml")}:
                        content = (
                            child.find(Q(ns, "sdtContent")) if child.tag == Q(ns, "sdt") else child
                        )
                        if content is not None:
                            paragraph = terminal_paragraph(content)
                            if paragraph is not None:
                                return paragraph
                return None

            def prune(parent: ET.Element, keep: ET.Element | None = None) -> None:
                nonlocal removed_containers
                required_terminal = parent.tag in {
                    Q(ns, name)
                    for name in (
                        "tc",
                        "hdr",
                        "ftr",
                        "footnote",
                        "endnote",
                        "comment",
                        "txbxContent",
                    )
                }
                terminal = terminal_paragraph(parent) if required_terminal else keep
                for child in list(parent):
                    if child.tag == Q(ns, "p") and empty_paragraph(child) and child is not terminal:
                        parent.remove(child)
                        removed_containers += 1
                    elif child.tag in {Q(ns, "sdt"), Q(ns, "customXml")}:
                        content = (
                            child.find(Q(ns, "sdtContent")) if child.tag == Q(ns, "sdt") else child
                        )
                        if content is not None:
                            prune(content, terminal)
                            if not len(content) and not content.text:
                                parent.remove(child)
                                removed_containers += 1
                    elif child.tag in {
                        Q(ns, name)
                        for name in ("tbl", "tr", "tc", "footnote", "endnote", "comment")
                    }:
                        prune(child)
                    else:
                        for box in child.iter(Q(ns, "txbxContent")):
                            prune(box)

            for name, root in roots.items():
                story_roots = (
                    [root.find(Q(ns, "body"))]
                    if root.tag == Q(ns, "document")
                    else [root]
                    if root.tag
                    in {Q(ns, name) for name in ("hdr", "ftr", "footnotes", "endnotes", "comments")}
                    else []
                )
                previous = removed_containers
                for story in story_roots:
                    if story is not None:
                        prune(story)
                if removed_containers != previous:
                    parts[name] = serialize_xml_with_preserved_qnames(root)
        styles = roots.get(_STYLES)
        removed_styles: list[str] = []
        duplicates: list[tuple[str, str]] = []
        if styles is not None:
            definitions = {
                node.get(Q(ns, "styleId")): node
                for node in styles.findall(Q(ns, "style"))
                if node.get(Q(ns, "styleId"))
            }
            referenced = set().union(*(_styles(root, ns) for root in roots.values()))
            for root in roots.values():
                referenced.update(
                    node.get(Q(ns, "val"))
                    for node in root.iter()
                    if node.tag
                    in {
                        Q(ns, name)
                        for name in (
                            "basedOn",
                            "next",
                            "link",
                            "styleLink",
                            "numStyleLink",
                        )
                    }
                    and node.get(Q(ns, "val"))
                )
            tokens = {token.casefold() for token in _field_style_tokens(list(roots.values()), ns)}
            latent = {
                node.get(Q(ns, "name"), "").casefold()
                for node in styles.iter(Q(ns, "lsdException"))
            }
            protected = set(referenced)
            intrinsic = set()
            for identifier, node in definitions.items():
                name = node.find(Q(ns, "name"))
                display = name.get(Q(ns, "val"), "") if name is not None else ""
                if (
                    node.get(Q(ns, "default"), "false") in {"1", "true", "on"}
                    or identifier in {"Normal", "DefaultParagraphFont", "TableNormal"}
                    or display.casefold() in latent
                    or identifier.casefold() in latent
                    or display.casefold() in tokens
                    or identifier.casefold() in tokens
                    or any(
                        attribute in item.attrib
                        for item in node.iter()
                        for attribute in _RELATIONSHIP_ATTRIBUTES
                    )
                ):
                    protected.add(identifier)
                    intrinsic.add(identifier)
            relationship_identity: dict[str, tuple[str, ...]] = {}
            for relation in _read_rels(parts, _STYLES):
                if relation.get("Id") is None:
                    continue
                relationship_identity[relation.get("Id")] = (
                    relation.get("Type", ""),
                    "external",
                    relation.get("Target", ""),
                )
                if relation.get("TargetMode") != "External":
                    relationship_identity[relation.get("Id")] = (
                        relation.get("Type", ""),
                        "internal",
                        _resolve(_STYLES, relation.get("Target", "")),
                    )
            latent_parents = {child: parent for parent in styles.iter() for child in parent}
            display_names = {
                identifier: (
                    node.find(Q(ns, "name")).get(Q(ns, "val"), "")
                    if node.find(Q(ns, "name")) is not None
                    else ""
                )
                for identifier, node in definitions.items()
            }
            targeted = {
                identifier
                for identifier, display in display_names.items()
                if display.casefold() in tokens or identifier.casefold() in tokens
            }

            def is_default(node: ET.Element) -> bool:
                return node.get(Q(ns, "default"), "false") in {"1", "true", "on"}

            def relationship_attributes(node: ET.Element) -> list[tuple[ET.Element, str]]:
                return [
                    (item, attribute)
                    for item in node.iter()
                    for attribute in _RELATIONSHIP_ATTRIBUTES
                    if attribute in item.attrib
                ]

            def style_key(node: ET.Element) -> bytes:
                # Definition content, with the default flag and relationship ids
                # replaced by what they mean (target identity), never their spelling.
                clone = copy.deepcopy(node)
                clone.attrib.pop(Q(ns, "default"), None)
                for item, attribute in relationship_attributes(clone):
                    value = item.get(attribute)
                    item.set(
                        attribute,
                        f"relationship:{relationship_identity[value]!r}"
                        if value in relationship_identity
                        else f"unresolved:{value}",
                    )
                return _definition_key(clone, ns, "style")

            def latent_exception(identifier: str) -> ET.Element | None:
                names = {identifier.casefold(), display_names[identifier].casefold()}
                return next(
                    (
                        item
                        for item in styles.iter(Q(ns, "lsdException"))
                        if item.get(Q(ns, "name"), "").casefold() in names
                    ),
                    None,
                )

            def coalescible(identifier: str, previous: str) -> bool:
                """An intrinsic duplicate merges only when every dependency is equivalent."""
                node = definitions[identifier]
                if is_default(node):
                    return False
                if any(
                    item.get(attribute) not in relationship_identity
                    for owner in (node, definitions[previous])
                    for item, attribute in relationship_attributes(owner)
                ):
                    return False
                removed, kept = latent_exception(identifier), latent_exception(previous)
                if (removed is None) != (kept is None):
                    return False
                if removed is None or removed is kept:
                    return True
                return {k: v for k, v in removed.attrib.items() if k != Q(ns, "name")} == {
                    k: v for k, v in kept.attrib.items() if k != Q(ns, "name")
                }

            removed_definitions: dict[str, ET.Element] = {}
            while True:
                canonical: dict[bytes, str] = {}
                round_duplicates: list[tuple[str, str]] = []
                for identifier in sorted(
                    definitions,
                    key=lambda key: (
                        not is_default(definitions[key]),
                        key not in targeted,
                        key not in protected,
                        key,
                    ),
                ):
                    node = definitions[identifier]
                    key = style_key(node)
                    previous = canonical.get(key)
                    builtin = node.get(Q(ns, "customStyle"), "false") not in {"1", "true", "on"}
                    removable = identifier not in protected and (
                        not builtin or policies.remove_unused_builtin_styles
                    )
                    duplicate = previous is not None and (
                        policies.remove_duplicate_styles
                        or policies.remove_duplicate_referenced_styles
                    )
                    referenced_duplicate = (
                        duplicate
                        and policies.remove_duplicate_referenced_styles
                        and not builtin
                        and (identifier not in intrinsic or coalescible(identifier, previous))
                    )
                    unused = (
                        policies.remove_unused_builtin_styles
                        if builtin
                        else policies.remove_unused_styles
                    )
                    if removable and (unused or duplicate) or referenced_duplicate:
                        if referenced_duplicate and identifier in intrinsic:
                            removed_latent = latent_exception(identifier)
                            if (
                                removed_latent is not None
                                and removed_latent is not latent_exception(previous)
                            ):
                                latent_parents[removed_latent].remove(removed_latent)
                        styles.remove(node)
                        del definitions[identifier]
                        removed_definitions[identifier] = node
                        removed_styles.append(identifier)
                        if duplicate:
                            duplicates.append((identifier, previous))
                            round_duplicates.append((identifier, previous))
                    else:
                        canonical.setdefault(key, identifier)
                if not (policies.remove_duplicate_referenced_styles and round_duplicates):
                    break
                # Dependents of a coalesced style now compare equal on its survivor.
                round_remap = dict(round_duplicates)
                for item in styles.iter():
                    if (
                        item.tag in {Q(ns, name) for name in ("basedOn", "next", "link")}
                        and item.get(Q(ns, "val")) in round_remap
                    ):
                        item.set(Q(ns, "val"), round_remap[item.get(Q(ns, "val"))])
            if duplicates and policies.remove_duplicate_referenced_styles:
                remap = dict(duplicates)

                def survivor(identifier: str) -> str:
                    while identifier in remap:
                        identifier = remap[identifier]
                    return identifier

                remap = {identifier: survivor(identifier) for identifier in remap}
                tags = {
                    Q(ns, name)
                    for name in (
                        "pStyle",
                        "rStyle",
                        "tblStyle",
                        "style",
                        "basedOn",
                        "next",
                        "link",
                        "styleLink",
                        "numStyleLink",
                    )
                }
                field_targets: dict[str, str] = {}
                for identifier, kept in remap.items():
                    removed_display = display_names[identifier]
                    field_targets.setdefault(identifier.casefold(), kept)
                    if removed_display:
                        field_targets[removed_display.casefold()] = display_names[kept] or kept
                for name, root in roots.items():
                    changed_references = False
                    for node in root.iter():
                        if node.tag in tags and node.get(Q(ns, "val")) in remap:
                            node.set(Q(ns, "val"), remap[node.get(Q(ns, "val"))])
                            changed_references = True
                    for instruction, nodes in _field_instruction_nodes(root, ns):
                        match = re.match(
                            r'^(\s*STYLEREF\s+)(?:"([^"]+)"|([A-Za-z_][A-Za-z0-9_]*))',
                            instruction,
                            re.IGNORECASE,
                        )
                        if match is None:
                            continue
                        target = match.group(2) or match.group(3)
                        replacement = field_targets.get(target.casefold())
                        if replacement is None or replacement == target:
                            continue
                        rewritten = f'{match.group(1)}"{replacement}"{instruction[match.end() :]}'
                        simple = next((n for n in nodes if n.tag == Q(ns, "fldSimple")), None)
                        if simple is not None:
                            simple.set(Q(ns, "instr"), rewritten)
                        else:
                            texts = [n for n in nodes if n.tag == Q(ns, "instrText")]
                            texts[0].text = rewritten
                            for extra in texts[1:]:
                                extra.text = ""
                        changed_references = True
                    if changed_references:
                        parts[name] = serialize_xml_with_preserved_qnames(root)
            if removed_styles:
                parts[_STYLES] = serialize_xml_with_preserved_qnames(styles)
        removed_lists: list[str] = []
        numbering = roots.get("word/numbering.xml")
        if numbering is not None and policies.remove_unused_lists:
            used = {
                node.get(Q(ns, "val"))
                for name, root in roots.items()
                if name != "word/numbering.xml"
                for node in root.iter(Q(ns, "numId"))
            }
            for node in list(numbering.findall(Q(ns, "num"))):
                identifier = node.get(Q(ns, "numId"))
                if identifier not in used:
                    numbering.remove(node)
                    removed_lists.append(f"num:{identifier}")
            abstracts = {node.get(Q(ns, "val")) for node in numbering.iter(Q(ns, "abstractNumId"))}
            for node in list(numbering.findall(Q(ns, "abstractNum"))):
                identifier = node.get(Q(ns, "abstractNumId"))
                if identifier not in abstracts:
                    numbering.remove(node)
                    removed_lists.append(f"abstractNum:{identifier}")
            if removed_lists:
                parts["word/numbering.xml"] = serialize_xml_with_preserved_qnames(numbering)
        removed_parts: set[str] = set()
        if numbering is not None and policies.remove_unused_picture_bullets:
            used_bullets = {
                item.get(Q(ns, "val")) for item in numbering.iter(Q(ns, "lvlPicBulletId"))
            }
            removed_relations = set()
            for bullet in list(numbering.findall(Q(ns, "numPicBullet"))):
                identifier = bullet.get(Q(ns, "numPicBulletId"))
                if identifier not in used_bullets:
                    removed_relations.update(
                        value
                        for item in bullet.iter()
                        for name, value in item.attrib.items()
                        if name.startswith(_RELATIONSHIP_PREFIXES)
                    )
                    numbering.remove(bullet)
                    removed_lists.append(f"numPicBullet:{identifier}")
            retained_relations = {
                value
                for item in numbering.iter()
                for name, value in item.attrib.items()
                if name.startswith(_RELATIONSHIP_PREFIXES)
            }
            rels = _read_rels(parts, "word/numbering.xml")
            media_candidates = set()
            relations_changed = False
            for relation in list(rels):
                if relation.get("Id") in removed_relations - retained_relations:
                    if relation.get("TargetMode") != "External":
                        media_candidates.add(
                            _resolve("word/numbering.xml", relation.get("Target", ""))
                        )
                    rels.remove(relation)
                    relations_changed = True
            if relations_changed:
                parts[_rels_name("word/numbering.xml")] = serialize_xml_with_preserved_qnames(rels)
            retained_targets = set()
            for name, data in parts.items():
                if not name.endswith(".rels"):
                    continue
                if name == _RELS:
                    owner = ""
                else:
                    directory, filename = posixpath.split(name)
                    owner = posixpath.join(
                        posixpath.dirname(directory), filename.removesuffix(".rels")
                    )
                for relation in parse_xml_with_preserved_qnames(data):
                    if relation.get("TargetMode") != "External":
                        retained_targets.add(_resolve(owner, relation.get("Target", "")))
            for media in sorted(media_candidates - retained_targets):
                content_type = _content_type(media, content_types)
                if (
                    media in parts
                    and content_type is not None
                    and content_type[1].startswith("image/")
                    and _rels_name(media) not in parts
                ):
                    removed_parts.add(media)
                    del parts[media]
                    for override in list(content_types.findall(_q(_CT, "Override"))):
                        if override.get("PartName") == "/" + media:
                            content_types.remove(override)
            if removed_parts:
                parts["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(content_types)
            if removed_relations or any(
                value.startswith("numPicBullet:") for value in removed_lists
            ):
                parts["word/numbering.xml"] = serialize_xml_with_preserved_qnames(numbering)
        with tempfile.NamedTemporaryFile(
            dir=output.parent, prefix=f".{output.name}.", suffix=".tmp", delete=False
        ) as stream:
            staged = Path(stream.name)
        _write(staged, parts, source, removed_parts=removed_parts)
        changed = {
            name for name, data in parts.items() if baseline.get(name) != data
        } | removed_parts
        candidate = inspect_docx(staged)
        if (
            not candidate.valid
            or _sha256(source) != source_sha256
            or not _unselected_destination_parts_preserved(baseline, staged, changed)
        ):
            return CleanupReport("refused", "assembly_staged_verification_failed")
        os.replace(staged, output)
        staged = None
        return CleanupReport(
            "published",
            None,
            tuple(removed_styles),
            tuple(removed_lists),
            tuple(duplicates),
            candidate.file_sha256,
            removed_containers,
        )
    except (OSError, KeyError, ET.ParseError, ValueError, zipfile.BadZipFile):
        return CleanupReport("refused", "assembly_package_invalid")
    finally:
        if staged is not None:
            staged.unlink(missing_ok=True)


# -- document splitting -------------------------------------------------------

SPLIT_MODES = ("section", "page_break", "heading1")
# Inline containers a body paragraph may wrap an effective page break in. A break
# inside a deleted run, a text box, a nested table or any other story is not a
# boundary of the body flow.
_SPLIT_BREAK_CONTAINERS = frozenset(
    {
        "r",
        "hyperlink",
        "ins",
        "moveTo",
        "smartTag",
        "customXml",
        "sdt",
        "sdtContent",
        "fldSimple",
    }
)
_SPLIT_PAIRED_MARKERS = (
    ("bookmarkStart", "bookmarkEnd"),
    ("commentRangeStart", "commentRangeEnd"),
)
_FALSE_VALUES = frozenset({"0", "false", "off"})


@dataclass(frozen=True)
class SplitPart:
    """One published split output and the body content it carries."""

    name: str
    sha256: str
    blocks: int
    sections: int
    text_sha256: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "sha256": self.sha256,
            "blocks": self.blocks,
            "sections": self.sections,
            "text_sha256": self.text_sha256,
        }


@dataclass(frozen=True)
class SplitReport:
    """Content-safe receipt for one atomically published document split."""

    status: str
    code: str | None
    source_sha256: str
    by: str
    parts: tuple[SplitPart, ...] = ()
    text_preserved: bool | None = None
    target: str | None = None
    layout_boundaries: tuple[str, ...] = ()

    @property
    def passed(self) -> bool:
        return self.status == "published"

    def to_dict(self) -> dict[str, Any]:
        result = {
            "schema_version": 1,
            "operation": "split",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "by": self.by,
            "parts": [part.to_dict() for part in self.parts],
            "text_preserved": self.text_preserved,
        }
        if self.target is not None:
            result["target"] = self.target
        if self.layout_boundaries:
            result["layout_boundaries"] = list(self.layout_boundaries)
        return result


def _local(tag: Any) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _body_text(elements: Sequence[ET.Element], ns: str) -> str:
    return "".join(node.text or "" for element in elements for node in element.iter(Q(ns, "t")))


def _enabled(element: ET.Element | None, ns: str) -> bool:
    return element is not None and (element.get(Q(ns, "val")) or "true") not in _FALSE_VALUES


def _section_property(paragraph: ET.Element, ns: str) -> ET.Element | None:
    properties = paragraph.find(Q(ns, "pPr"))
    return properties.find(Q(ns, "sectPr")) if properties is not None else None


def _heading_style_ids(
    parts: Mapping[str, bytes], ns: str, heading_level: int = 1
) -> tuple[set[str], str | None]:
    """Styles with the requested effective outline level, and the default style."""

    data = parts.get(_STYLES)
    if data is None:
        return set(), None
    root = ET.fromstring(data)
    based_on: dict[str, str | None] = {}
    levels: dict[str, str | None] = {}
    names: dict[str, str] = {}
    default: str | None = None
    for style in root.findall(Q(ns, "style")):
        if style.get(Q(ns, "type")) != "paragraph":
            continue
        identifier = style.get(Q(ns, "styleId"))
        if identifier is None:
            continue
        if _enabled(style, ns) and style.get(Q(ns, "default")) in {"1", "true", "on"}:
            default = identifier
        parent = style.find(Q(ns, "basedOn"))
        based_on[identifier] = parent.get(Q(ns, "val")) if parent is not None else None
        properties = style.find(Q(ns, "pPr"))
        outline = properties.find(Q(ns, "outlineLvl")) if properties is not None else None
        levels[identifier] = outline.get(Q(ns, "val")) if outline is not None else None
        name = style.find(Q(ns, "name"))
        names[identifier] = (name.get(Q(ns, "val")) or "").casefold() if name is not None else ""

    def level(identifier: str) -> str | None:
        seen: set[str] = set()
        current: str | None = identifier
        while current is not None and current not in seen and current in levels:
            seen.add(current)
            if levels[current] is not None:
                return levels[current]
            if names[current] == f"heading {heading_level}":
                return str(heading_level - 1)
            current = based_on[current]
        return None

    return {
        identifier for identifier in levels if level(identifier) == str(heading_level - 1)
    }, default


def _is_heading1(
    paragraph: ET.Element, ns: str, headings: set[str], default_style: str | None
) -> bool:
    properties = paragraph.find(Q(ns, "pPr"))
    style = default_style
    if properties is not None:
        outline = properties.find(Q(ns, "outlineLvl"))
        if outline is not None:
            return outline.get(Q(ns, "val")) == "0"
        reference = properties.find(Q(ns, "pStyle"))
        if reference is not None:
            style = reference.get(Q(ns, "val"))
    return style in headings


def _shell(element: ET.Element) -> ET.Element:
    clone = copy.copy(element)
    del clone[:]
    clone.tail = None
    return clone


def _page_break_paths(paragraph: ET.Element, ns: str) -> list[list[int]]:
    """Child-index paths to every effective page break in one body paragraph."""

    paths: list[list[int]] = []

    def walk(element: ET.Element, path: list[int]) -> None:
        for index, child in enumerate(element):
            local = _local(child.tag)
            if (
                local == "br"
                and child.tag == Q(ns, "br")
                and child.get(Q(ns, "type")) == "page"
                and element is not paragraph
            ):
                paths.append([*path, index])
            elif child.tag.startswith(f"{{{ns}}}") and local in _SPLIT_BREAK_CONTAINERS:
                walk(child, [*path, index])

    walk(paragraph, [])
    return paths


def _split_element(element: ET.Element, path: Sequence[int]) -> tuple[ET.Element, ET.Element]:
    """Split one element at a descendant break, dropping the break itself.

    Property children (``pPr``, ``rPr``, ``hyperlink`` attributes and so on) stay on
    both halves, so each fragment keeps the formatting the source applied to it.
    """

    children = list(element)
    index = path[0]
    before, after = _shell(element), _shell(element)
    properties = [
        copy.deepcopy(child) for child in children[:index] if _local(child.tag).endswith("Pr")
    ]
    if len(path) == 1:
        before.extend(children[:index])
        after.extend([*properties, *children[index + 1 :]])
    else:
        head, tail = _split_element(children[index], path[1:])
        before.extend([*children[:index], head])
        after.extend([*properties, tail, *children[index + 1 :]])
    return before, after


def _has_content(paragraph: ET.Element) -> bool:
    return any(not _local(child.tag).endswith("Pr") for child in paragraph)


def _split_paragraph(paragraph: ET.Element, ns: str) -> list[ET.Element]:
    """Cut one paragraph at its page breaks; only the last fragment keeps ``sectPr``."""

    fragments: list[ET.Element] = []
    current = paragraph
    while paths := _page_break_paths(current, ns):
        before, current = _split_element(current, paths[0])
        fragments.append(before)
    fragments.append(current)
    for fragment in fragments[:-1]:
        properties = fragment.find(Q(ns, "pPr"))
        section = properties.find(Q(ns, "sectPr")) if properties is not None else None
        if section is not None and properties is not None:
            properties.remove(section)
    return fragments


def _materialize_references(
    section: ET.Element, references: Mapping[tuple[str, str], ET.Element], ns: str
) -> ET.Element:
    """One section copy that names every inherited header and footer explicitly."""

    clone = copy.deepcopy(section)
    for child in list(clone):
        if child.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}:
            clone.remove(child)
    ordered = sorted(
        references.items(),
        key=lambda item: (item[0][0] != "headerReference", item[0][1]),
    )
    for position, (_key, reference) in enumerate(ordered):
        reference_copy = copy.deepcopy(reference)
        reference_copy.tail = None
        clone.insert(position, reference_copy)
    return clone


def _body_range_marker_edge(node: ET.Element) -> str | None:
    """Return ``start`` or ``end`` for a zero-width range marker, else ``None``.

    Bookmarks, permission ranges and every ``*RangeStart``/``*RangeEnd`` family
    may sit directly in the body between blocks, where they have no geometry.
    """

    local = node.tag.rsplit("}", 1)[-1]
    if local.endswith("RangeStart") or local in {"bookmarkStart", "permStart"}:
        return "start"
    if local.endswith("RangeEnd") or local in {"bookmarkEnd", "permEnd"}:
        return "end"
    return None


def _drop_unpaired_markers(body: ET.Element, ns: str) -> None:
    """Remove range starts and ends whose partner landed in another output."""

    for start_name, end_name in _SPLIT_PAIRED_MARKERS:
        starts: dict[str, list[tuple[ET.Element, ET.Element]]] = {}
        ends: dict[str, list[tuple[ET.Element, ET.Element]]] = {}
        for parent in body.iter():
            for child in parent:
                if child.tag == Q(ns, start_name):
                    starts.setdefault(child.get(Q(ns, "id"), ""), []).append((parent, child))
                elif child.tag == Q(ns, end_name):
                    ends.setdefault(child.get(Q(ns, "id"), ""), []).append((parent, child))
        for identifier, entries in starts.items():
            if identifier not in ends:
                for parent, child in entries:
                    parent.remove(child)
        for identifier, entries in ends.items():
            if identifier not in starts:
                for parent, child in entries:
                    parent.remove(child)


def _split_chunks(
    root: ET.Element, ns: str, by: str, parts: Mapping[str, bytes]
) -> list[list[tuple[ET.Element, int]]]:
    """Group body blocks into outputs; each block carries its source section index."""

    body = root.find(Q(ns, "body"))
    if body is None:
        raise ValueError("split_body_missing")
    headings, default_style = _heading_style_ids(parts, ns) if by == "heading1" else (set(), None)
    chunks: list[list[tuple[ET.Element, int]]] = [[]]
    section = 0
    for child in body:
        if child.tag == Q(ns, "sectPr"):
            continue
        if child.tag != Q(ns, "p"):
            chunks[-1].append((child, section))
            continue
        carries_section = _section_property(child, ns) is not None
        if by == "heading1" and _is_heading1(child, ns, headings, default_style) and chunks[-1]:
            chunks.append([])
        if by == "page_break":
            properties = child.find(Q(ns, "pPr"))
            before = properties.find(Q(ns, "pageBreakBefore")) if properties is not None else None
            if _enabled(before, ns) and chunks[-1]:
                chunks.append([])
            fragments = _split_paragraph(child, ns)
            for position, fragment in enumerate(fragments):
                if position:
                    chunks.append([])
                if position < len(fragments) - 1 and not _has_content(fragment):
                    continue
                chunks[-1].append((fragment, section))
        else:
            chunks[-1].append((child, section))
        if carries_section:
            section += 1
            if by == "section":
                chunks.append([])
    return [chunk for chunk in chunks if chunk]


def _split_document_xml(
    root: ET.Element,
    ns: str,
    chunk: Sequence[tuple[ET.Element, int]],
    sections: Sequence[ET.Element],
    references: Mapping[int, Mapping[tuple[str, str], ET.Element]],
    *,
    drop_unpaired_markers: bool = True,
) -> tuple[bytes, list[ET.Element], int]:
    """Serialize one output: its blocks, then the section in effect at its end."""

    body = root.find(Q(ns, "body"))
    assert body is not None
    units = [copy.deepcopy(element) for element, _section in chunk]
    section_count = 0
    for unit, (_element, index) in zip(units, chunk, strict=True):
        properties = unit.find(Q(ns, "pPr")) if unit.tag == Q(ns, "p") else None
        carried = properties.find(Q(ns, "sectPr")) if properties is not None else None
        if carried is None or properties is None or index >= len(sections):
            continue
        position = list(properties).index(carried)
        properties[position] = _materialize_references(
            carried, references.get(id(sections[index]), {}), ns
        )
        section_count += 1
    final: ET.Element | None = None
    # Range ends that close after the output's last block follow it: that block
    # still decides the section, so the body section lands after those markers.
    tail = len(units) - 1
    while tail and _body_range_marker_edge(units[tail]) is not None:
        tail -= 1
    last = units[tail]
    properties = last.find(Q(ns, "pPr")) if last.tag == Q(ns, "p") else None
    carried = properties.find(Q(ns, "sectPr")) if properties is not None else None
    if carried is not None and properties is not None:
        # The output's last paragraph ends its section: that becomes the body section.
        properties.remove(carried)
        if not len(properties) and not properties.attrib:
            last.remove(properties)
        final = carried
    elif chunk[tail][1] < len(sections):
        final = _materialize_references(
            sections[chunk[tail][1]], references.get(id(sections[chunk[tail][1]]), {}), ns
        )
        section_count += 1
    document = copy.copy(root)
    new_body = _shell(body)
    new_body.extend(units)
    if final is not None:
        final.tail = None
        new_body.append(final)
    if drop_unpaired_markers:
        _drop_unpaired_markers(new_body, ns)
    document[list(root).index(body)] = new_body
    return serialize_xml_with_preserved_qnames(document), units, section_count


def _split_note_parts(parts: Mapping[str, bytes]) -> dict[str, str]:
    """Map ``footnote`` and ``endnote`` to the notes part the main document relates."""

    data = parts.get("word/_rels/document.xml.rels")
    if data is None:
        return {}
    found: dict[str, str] = {}
    for relationship in ET.fromstring(data):
        kind = (relationship.get("Type") or "").rsplit("/", 1)[-1]
        target = relationship.get("Target") or ""
        if kind not in {"footnotes", "endnotes"} or relationship.get("TargetMode") == "External":
            continue
        name = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("word", target))
        )
        if name in parts:
            found[kind[:-1]] = name
    return found


def _split_notes(
    parts: Mapping[str, bytes],
    note_parts: Mapping[str, str],
    units: Sequence[ET.Element],
    ns: str,
) -> dict[str, bytes]:
    """Notes parts that keep only the notes one output references.

    Separator notes stay. A reader that pairs references with notes by position
    rather than by id (LibreOffice does) would otherwise show another part's note.
    """

    roots = {
        kind: parse_xml_with_preserved_qnames(parts[name]) for kind, name in note_parts.items()
    }
    notes = {
        (kind, note.get(Q(ns, "id"))): note
        for kind, root in roots.items()
        for note in root.findall(Q(ns, kind))
    }
    referenced: set[tuple[str, str]] = set()
    pending = _note_references(list(units), ns)
    while pending:
        key = pending.pop()
        if key in referenced:
            continue
        note = notes.get(key)
        if note is None:
            raise ValueError("assembly_note_reference_invalid")
        referenced.add(key)
        pending.extend(_note_references([note], ns))
    replaced: dict[str, bytes] = {}
    for kind, name in note_parts.items():
        root = roots[kind]
        changed = False
        for note in list(root):
            if (
                note.tag == Q(ns, kind)
                and note.get(Q(ns, "type"), "normal") == "normal"
                and (kind, note.get(Q(ns, "id"))) not in referenced
            ):
                root.remove(note)
                changed = True
        replaced[name] = serialize_xml_with_preserved_qnames(root) if changed else parts[name]
    return replaced


def split_document(
    source: str | Path,
    output_dir: str | Path,
    *,
    by: str,
    source_sha256: str | None = None,
) -> SplitReport:
    """Split one DOCX into consecutive DOCX files and publish them atomically.

    ``by`` selects the boundaries: ``"section"`` ends an output after every section
    break, ``"page_break"`` starts one at every explicit page break (``w:br
    w:type="page"`` in a body paragraph, or ``w:pageBreakBefore``), and
    ``"heading1"`` starts one before every body paragraph whose effective outline
    level is 0. Every output keeps the source's styles, numbering, theme, settings,
    and header or footer parts, and the notes it references; the section properties
    in effect at the end of an output become its final ``w:sectPr``, with inherited header and footer
    references written explicitly. The outputs' body text, concatenated in order,
    equals the source body text, and nothing is published unless every output
    inspects as a valid package. Outputs are ``part-001.docx`` onward inside the new
    directory ``output_dir``, which must not exist.
    """

    source_path = Path(source).resolve()
    destination = Path(output_dir).resolve()
    actual = _sha256(source_path) if source_path.is_file() else ""

    def refuse(code: str) -> SplitReport:
        return SplitReport("refused", code, actual or (source_sha256 or ""), by)

    if by not in SPLIT_MODES:
        return refuse("split_mode_invalid")
    if not source_path.is_file():
        return refuse("source_missing")
    if source_sha256 is not None and source_sha256 != actual:
        return refuse("source_hash_mismatch")
    if destination.exists():
        return refuse("destination_exists")
    if not destination.parent.is_dir():
        return refuse("destination_parent_missing")
    if not inspect_docx(source_path).valid:
        return refuse("source_inspection_invalid")
    staging: Path | None = None
    try:
        with zipfile.ZipFile(source_path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        root = parse_xml_with_preserved_qnames(parts[_DOCUMENT])
        ns = _word_ns(root)
        if ns is None:
            return refuse("source_word_namespace_unsupported")
        body = root.find(Q(ns, "body"))
        if body is None:
            return refuse("split_body_missing")
        source_text = _body_text([body], ns)
        sections = _section_nodes(root, ns)
        references = _effective_section_references(root, ns)
        note_parts = _split_note_parts(parts)
        chunks = _split_chunks(root, ns, by, parts)
        if not chunks:
            return refuse("split_body_empty")
        staging = Path(tempfile.mkdtemp(prefix=f".{destination.name}.", dir=destination.parent))
        published: list[SplitPart] = []
        texts: list[str] = []
        for index, chunk in enumerate(chunks, start=1):
            document_xml, units, section_count = _split_document_xml(
                root, ns, chunk, sections, references
            )
            name = f"part-{index:03d}.docx"
            path = staging / name
            notes = _split_notes(parts, note_parts, units, ns)
            _write(path, {_DOCUMENT: document_xml, **notes}, source_path)
            if not inspect_docx(path).valid:
                return refuse("split_staged_verification_failed")
            reread = parse_xml_with_preserved_qnames(
                PackageGraph.from_file(path).read_part_bytes(_DOCUMENT)
            )
            reread_body = reread.find(Q(ns, "body"))
            text = _body_text([reread_body] if reread_body is not None else [], ns)
            texts.append(text)
            published.append(
                SplitPart(
                    name,
                    _sha256(path),
                    sum(1 for unit in units if unit.tag in {Q(ns, "p"), Q(ns, "tbl")}),
                    section_count,
                    hashlib.sha256(text.encode("utf-8")).hexdigest(),
                )
            )
        if "".join(texts) != source_text:
            return refuse("split_text_not_preserved")
        if _sha256(source_path) != actual:
            return refuse("source_hash_changed")
        os.rename(staging, destination)
        staging = None
        return SplitReport("published", None, actual, by, tuple(published), True)
    except PackageError:
        return refuse("source_package_invalid")
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile):
        return refuse("source_package_invalid")
    except ValueError as error:
        message = str(error)
        return refuse(
            message if message.startswith(("split_", "assembly_")) else "source_package_invalid"
        )
    finally:
        if staging is not None:
            for child in staging.iterdir():
                child.unlink()
            staging.rmdir()
