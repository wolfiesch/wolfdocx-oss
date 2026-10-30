"""Byte-preserving package edits, node bindings, and verified DOCX publication."""

from __future__ import annotations

import copy
import hashlib
import mmap
import re
import struct
import tempfile
import zipfile
from collections.abc import Callable, Mapping
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO, cast
from xml.etree import ElementTree as ET
from xml.parsers import expat

from ..document import Document
from ..inspect import Q, _parse_bounded_xml, _verify_staged_changes, _XmlSafetyError
from ..model import (
    DocumentVisitor,
    FieldEnd,
    FieldSeparator,
    FieldStart,
    Node,
    NodeEditError,
    NodeType,
    StoryType,
    VisitorAction,
)
from ..output import SaveOutputParameters
from ..package import PackageError, PackageGraph, PackageLimits
from . import (
    ListEdits,
    RunEdits,
    SectionEdits,
    SessionEditReport,
    StylingEdits,
    TableEdits,
)
from ._bridge import _APPLIED as _EDIT_APPLIED
from ._bridge import operation_detail
from ._bridge import refused as _edit_refused

if TYPE_CHECKING:
    from ..authoring import CellTextAuthoringReport
    from ..find_replace import FindReplaceOptions, FindReplaceReport
    from ..paragraph_authoring import ParagraphReport, ParagraphRequest
    from ..render_warnings import WarningCollection


def _plan_paragraph_edit(*args: Any, **kwargs: Any) -> Any:
    """:func:`wolfdocx.paragraph_authoring._plan_paragraph_edit`, imported on first use."""
    from ..paragraph_authoring import _plan_paragraph_edit

    return _plan_paragraph_edit(*args, **kwargs)



__all__ = ["DocumentOperations", "SessionSaveReport"]


@dataclass(frozen=True)
class SessionSaveReport(SaveOutputParameters):
    """Content-safe receipt for one :meth:`Document.save`."""

    status: str
    code: str | None
    source_sha256: str
    destination: str
    destination_sha256: str | None = None
    changed_parts: tuple[str, ...] = ()
    added_parts: tuple[str, ...] = ()
    removed_parts: tuple[str, ...] = ()
    raw_copied_parts: int = 0
    edits: tuple[Mapping[str, Any], ...] = ()

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "document-save",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination": self.destination,
            "destination_sha256": self.destination_sha256,
            "changed_parts": list(self.changed_parts),
            "added_parts": list(self.added_parts),
            "removed_parts": list(self.removed_parts),
            "raw_copied_parts": self.raw_copied_parts,
            "edits": [dict(edit) for edit in self.edits],
        }


SessionSaveReport.__module__ = Document.__module__


_PUBLISHED = "published"
_DOCUMENT = "word/document.xml"
_REFUSED = "refused"
_UNBOUND_FINGERPRINT = "0" * 64
# [MS-OFFMACRO2] 2.1: these main types belong to .docm/.dotm names, never .docx.
_MACRO_MAIN_CONTENT_TYPES = frozenset(
    {
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    }
)
_LOCAL_FILE_HEADER = b"PK\x03\x04"
_LOCAL_FILE_HEADER_SIZE = 30
_ZIP64_EXTRA_ID = 0x0001
_NOTE_RELATIONSHIPS = ("/footnotes", "/endnotes")
_PART_ERRORS = (PackageError, KeyError, OSError, ET.ParseError, _XmlSafetyError)
# Entries, changed parts, package graph, and edit count: what a refusal restores.
_Snapshot = tuple[dict[str, zipfile.ZipInfo], dict[str, bytes], PackageGraph, int]


def _patch_node_texts(data: bytes, ns: str, replacements: Mapping[int, str]) -> bytes:
    """Patch w:t contents by XML-order ordinal, never serializing their markup.

    Expat supplies byte offsets (including the BOM); ECMA-376 Part 1 17.3.3.31
    defines w:t. Empty-element text cannot be inserted without changing markup
    and is therefore refused. The lexical scanner mirrors _paragraph_xml_span.
    """
    from ..authoring import _start_tag_end

    encoding = (
        "utf-16-le"
        if data.startswith((b"\xff\xfe", b"<\x00"))
        else "utf-16-be"
        if data.startswith((b"\xfe\xff", b"\x00<"))
        else "utf-8"
    )
    wide = encoding.startswith("utf-16")
    parser = expat.ParserCreate(namespace_separator="}")
    count = 0
    active: tuple[int, int] | None = None
    spans: list[tuple[int, int, str]] = []

    def declaration(version: str, declared: str | None, standalone: int) -> None:
        nonlocal encoding
        if declared and not wide:
            encoding = declared

    def start(name: str, attributes: dict[str, str]) -> None:
        nonlocal count, active
        if name != f"{ns}}}t":
            return
        count += 1
        if count not in replacements:
            return
        begin = parser.CurrentByteIndex
        if wide:
            quote = None
            end = None
            for index in range(begin, len(data), 2):
                unit = data[index : index + 2]
                if quote is not None:
                    if unit == quote:
                        quote = None
                elif unit in ('"'.encode(encoding), "'".encode(encoding)):
                    quote = unit
                elif unit == ">".encode(encoding):
                    end = index + 2
                    break
        else:
            end = _start_tag_end(data, begin)
        if end is None or data[begin:end].decode(encoding).rstrip().endswith("/>"):
            raise NodeEditError("run_text_empty_element")
        active = (end, count)

    def end(name: str) -> None:
        nonlocal active
        if name == f"{ns}}}t" and active is not None:
            begin, ordinal = active
            spans.append((begin, parser.CurrentByteIndex, replacements[ordinal]))
            active = None

    parser.XmlDeclHandler = declaration
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    if len(spans) != len(replacements):
        raise NodeEditError("run_text_target_missing")
    chunks: list[bytes] = []
    offset = 0
    for begin, end_offset, text in spans:
        # Comments, PIs and CDATA delimiters are not run text. Do not erase
        # their lexical bytes when replacing the selected character data.
        if "<" in data[begin:end_offset].decode(encoding):
            raise NodeEditError("run_text_markup_not_preserved")
        escaped = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\r", "&#13;")
        )
        chunks.extend((data[offset:begin], escaped.encode(encoding, errors="xmlcharrefreplace")))
        offset = end_offset
    chunks.append(data[offset:])
    return b"".join(chunks)


def _validate_node_structure(
    root: ET.Element, ns: str, *, opaque_elements: set[int] | None = None
) -> None:
    """Validate core hierarchy, leaving unchanged opaque glossary entries alone."""
    opaque_elements = opaque_elements or set()
    nodes = [item for item in root.iter() if id(item) not in opaque_elements]
    blocks = {"p", "tbl", "sdt", "customXml", "altChunk"}
    markers = {
        "bookmarkStart",
        "bookmarkEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "proofErr",
        "permStart",
        "permEnd",
        "ins",
        "del",
        "moveFrom",
        "moveTo",
    }
    allowed = {
        "document": {"body", "background"},
        "body": blocks | markers | {"sectPr"},
        "glossaryDocument": {"docParts"},
        "docParts": {"docPart"},
        "docPart": {"docPartPr", "docPartBody"},
        "docPartBody": blocks | markers | {"sectPr"},
        "docPartPr": {"name", "style", "category", "types", "behaviors", "description", "guid"},
        "category": {"name", "gallery"},
        "types": {"type"},
        "behaviors": {"behavior"},
        "hdr": blocks | markers,
        "ftr": blocks | markers,
        "footnote": blocks | markers,
        "endnote": blocks | markers,
        "comment": blocks | markers,
        "txbxContent": blocks | markers,
        "tbl": {"tblPr", "tblGrid", "tr", "sdt", "customXml"} | markers,
        "tr": {"trPr", "tc", "sdt", "customXml"} | markers,
        "tc": blocks | markers | {"tcPr"},
        "p": {
            "pPr",
            "r",
            "hyperlink",
            "fldSimple",
            "sdt",
            "customXml",
            "subDoc",
            "smartTag",
            "dir",
            "bdo",
        }
        | markers,
        "r": {
            "rPr",
            "t",
            "instrText",
            "delText",
            "delInstrText",
            "tab",
            "br",
            "cr",
            "ptab",
            "drawing",
            "pict",
            "object",
            "fldChar",
            "footnoteReference",
            "endnoteReference",
            "commentReference",
            "noBreakHyphen",
            "softHyphen",
            "sym",
            "lastRenderedPageBreak",
            "separator",
            "continuationSeparator",
            "footnoteRef",
            "endnoteRef",
            "annotationRef",
            "pgNum",
            "ruby",
            "contentPart",
        },
        "hyperlink": {"r", "sdt", "customXml"} | markers,
        "fldSimple": {"r", "hyperlink", "sdt", "customXml"} | markers,
        "footnotes": {"footnote"},
        "smartTag": {"r", "hyperlink", "sdt", "customXml", "smartTag"} | markers,
        "endnotes": {"endnote"},
        "comments": {"comment"},
    }
    properties = {"p": "pPr", "r": "rPr", "tr": "trPr", "tc": "tcPr", "docPart": "docPartPr"}
    placements = {
        "p": {
            "body",
            "hdr",
            "docPartBody",
            "ftr",
            "footnote",
            "endnote",
            "comment",
            "txbxContent",
            "tc",
            "sdtContent",
            "customXml",
            "ins",
            "del",
            "moveFrom",
            "moveTo",
        },
        "r": {
            "p",
            "hyperlink",
            "fldSimple",
            "sdtContent",
            "customXml",
            "ins",
            "del",
            "moveFrom",
            "moveTo",
            "rubyBase",
            "rt",
            "smartTag",
            "dir",
            "bdo",
        },
        "tr": {"tbl", "sdtContent", "customXml", "ins", "del"},
        "tc": {"tr", "sdtContent", "customXml", "ins", "del"},
        "pPr": {"p", "pPrChange"},
        "rPr": {"r", "pPr", "sdtPr", "rPrChange"},
        "trPr": {"tr", "trPrChange"},
        "tcPr": {"tc", "tcPrChange"},
        "tblPr": {"tbl", "tblPrChange"},
        "tblGrid": {"tbl", "tblGridChange"},
        "docParts": {"glossaryDocument"},
        "docPart": {"docParts"},
        "docPartPr": {"docPart"},
        "docPartBody": {"docPart"},
        "sectPr": {"body", "docPartBody", "pPr", "sectPrChange"},
        "t": {"r"},
        "tab": {"r"},
        "br": {"r"},
        "cr": {"r"},
        "ptab": {"r"},
    }
    placements["tbl"] = placements["p"]
    for parent in nodes:
        if not isinstance(parent.tag, str) or not parent.tag.startswith("{" + ns + "}"):
            continue
        local = parent.tag.rsplit("}", 1)[-1]
        children = [
            child.tag.rsplit("}", 1)[-1]
            for child in parent
            if isinstance(child.tag, str) and child.tag.startswith("{" + ns + "}")
        ]
        if local in allowed and any(child not in allowed[local] for child in children):
            raise NodeEditError("node_structure_invalid")
        if any(child in placements and local not in placements[child] for child in children):
            raise NodeEditError("node_structure_invalid")
        prop = properties.get(local)
        if prop is not None and (
            children.count(prop) > 1 or (prop in children and children[0] != prop)
        ):
            raise NodeEditError("node_property_order_invalid")
        if local == "document" and children.count("body") != 1:
            raise NodeEditError("node_body_required")
        if local == "glossaryDocument" and children.count("docParts") > 1:
            raise NodeEditError("node_glossary_structure_invalid")
        if local == "docPart" and (
            children.count("docPartPr") != 1 or children.count("docPartBody") != 1
        ):
            raise NodeEditError("node_glossary_entry_invalid")
        if local in {"body", "docPartBody"} and (
            children.count("sectPr") > 1 or ("sectPr" in children and children[-1] != "sectPr")
        ):
            raise NodeEditError("node_section_order_invalid")
        if local in {
            "tc",
            "txbxContent",
            "hdr",
            "ftr",
            "footnote",
            "endnote",
            "comment",
            "docPartBody",
        }:
            if not any(child in blocks for child in children):
                raise NodeEditError("node_block_required")
            if local == "tc" and children[-1] != "p":
                raise NodeEditError("node_cell_final_paragraph_required")
        if local == "tbl":
            if not any(child in {"tr", "sdt", "customXml"} for child in children):
                raise NodeEditError("node_row_required")
            if children.count("tblPr") > 1 or children.count("tblGrid") > 1:
                raise NodeEditError("node_table_properties_invalid")
            if "tblPr" in children and children[0] != "tblPr":
                raise NodeEditError("node_property_order_invalid")
            if "tblGrid" in children and children.index("tblGrid") != int("tblPr" in children):
                raise NodeEditError("node_property_order_invalid")
        if local == "tr" and not any(child in {"tc", "sdt", "customXml"} for child in children):
            raise NodeEditError("node_cell_required")
    # Glossary entries are independent stories: neither paired ranges nor fields
    # may be completed by another entry's content.
    range_scopes = (
        [item for item in nodes if item.tag == Q(ns, "docPartBody")]
        if root.tag == Q(ns, "glossaryDocument")
        else [root]
    )
    for start, end in (
        ("bookmarkStart", "bookmarkEnd"),
        ("commentRangeStart", "commentRangeEnd"),
        ("moveFromRangeStart", "moveFromRangeEnd"),
        ("moveToRangeStart", "moveToRangeEnd"),
        ("permStart", "permEnd"),
    ):
        for scope in range_scopes:
            starts = [
                item.get(Q(ns, "id"))
                for item in scope.iter(Q(ns, start))
                if id(item) not in opaque_elements
            ]
            ends = [
                item.get(Q(ns, "id"))
                for item in scope.iter(Q(ns, end))
                if id(item) not in opaque_elements
            ]
            if (
                None in starts
                or None in ends
                or len(starts) != len(set(starts))
                or len(ends) != len(set(ends))
                or set(starts) != set(ends)
            ):
                raise NodeEditError("node_range_incomplete")
    for scope in range_scopes:
        names = [
            item.get(Q(ns, "name"))
            for item in scope.iter(Q(ns, "bookmarkStart"))
            if id(item) not in opaque_elements
        ]
        if len(names) != len(set(names)):
            raise NodeEditError("node_bookmark_name_duplicate")
    boundaries = {
        Q(ns, name) for name in ("txbxContent", "footnote", "endnote", "comment", "docPartBody")
    }
    scopes = [root, *(item for item in nodes if item is not root and item.tag in boundaries)]
    for scope in scopes:
        fields: list[bool] = []
        pending = [scope]
        while pending:
            item = pending.pop()
            if id(item) in opaque_elements or (item is not scope and item.tag in boundaries):
                continue
            if item.tag == Q(ns, "fldChar"):
                kind = item.get(Q(ns, "fldCharType"))
                if kind == "begin":
                    fields.append(False)
                elif kind == "separate":
                    if not fields or fields[-1]:
                        raise NodeEditError("node_field_incomplete")
                    fields[-1] = True
                elif kind == "end":
                    if not fields:
                        raise NodeEditError("node_field_incomplete")
                    fields.pop()
                else:
                    raise NodeEditError("node_field_incomplete")
            pending.extend(reversed(item))
        if fields:
            raise NodeEditError("node_field_incomplete")


def _structural(name: str) -> bool:
    """True for the content-types and relationship parts only operation edits change."""
    return name == "[Content_Types].xml" or name.endswith(".rels")


def _added_entry(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    """A clock-free deflated entry for a part an edit added."""
    entry = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
    entry.compress_type = zipfile.ZIP_DEFLATED
    return entry


def _parse_part(data: bytes, limits: PackageLimits, *, editable: bool) -> ET.Element:
    """Parse one part under the package limits; the only parse a session performs.

    ``editable`` trees keep qualified-name provenance so they re-serialize exactly;
    read-only trees take the fast path.
    """
    return _parse_bounded_xml(data, limits=limits, total_elements=[0], preserve_qnames=editable)


def _raw_member(data: bytes | mmap.mmap, info: zipfile.ZipInfo) -> memoryview:
    """The exact local entry, including its header and optional data descriptor."""
    offset = info.header_offset
    if data[offset : offset + 4] != _LOCAL_FILE_HEADER:
        raise zipfile.BadZipFile("local file header missing")
    name_length, extra_length = struct.unpack_from("<HH", data, offset + 26)
    start = offset + _LOCAL_FILE_HEADER_SIZE + name_length + extra_length
    # The stdlib reader records the next local-header/central-directory boundary.
    # Preserve descriptors and padding verbatim rather than interpreting them.
    end = info._end_offset
    if end is None or start + info.compress_size > end or end > len(data):
        raise zipfile.BadZipFile("member data truncated")
    return memoryview(data)[offset:end]


def _without_zip64_extra(extra: bytes) -> bytes:
    """Drop ZIP64 extra records; ``ZipInfo.FileHeader`` adds one when sizes need it."""
    kept = bytearray()
    offset = 0
    while offset + 4 <= len(extra):
        record_id, size = struct.unpack_from("<HH", extra, offset)
        end = offset + 4 + size
        if record_id != _ZIP64_EXTRA_ID:
            kept += extra[offset:end]
        offset = end
    return bytes(kept)


class _PreservedZipInfo(zipfile.ZipInfo):
    """Keep the source filename encoding and flags in the stdlib central writer."""

    __slots__ = ()

    def _encodeFilenameFlags(self) -> tuple[bytes, int]:
        encoding = "utf-8" if self.flag_bits & 0x800 else "cp437"
        return self.filename.encode(encoding), self.flag_bits


def _copy_member_raw(out: zipfile.ZipFile, info: zipfile.ZipInfo, raw: memoryview) -> None:
    """Append an original local entry and retain its central-directory metadata.

    Only the central-directory offset changes. Reconstructing a local header
    would normalize UTF-8 flags, local extras and data descriptors. The stdlib
    central writer also normalizes ASCII UTF-8 names on newer Python runtimes,
    so the copied ZipInfo uses the original encoding rather than that policy.
    """
    entry = copy.copy(info)
    entry.__class__ = _PreservedZipInfo
    assert out.fp is not None
    out.fp.seek(out.start_dir)
    entry.header_offset = out.fp.tell()
    out.fp.write(raw)
    out.start_dir = out.fp.tell()
    out.filelist.append(entry)
    out.NameToInfo[entry.filename] = entry


_TRANSITIONAL_BASE = "http://schemas.openxmlformats.org/"
_STRICT_BASE = "http://purl.oclc.org/ooxml/"
def _main_part(graph: PackageGraph) -> str:
    for relationship in graph.relationships:
        if relationship.type.endswith("/officeDocument") and not relationship.is_external:
            return relationship.resolved_target()
    raise NodeEditError("document_main_part_missing")











class DocumentOperations(
    RunEdits,
    TableEdits,
    SectionEdits,
    ListEdits,
    StylingEdits,
):
    """Apply in-memory edits and save a verified DOCX without clobbering files."""

    @property
    def warnings(self) -> WarningCollection:
        """Save refusal diagnostics in operation order."""
        if self._warnings is None:
            from ..render_warnings import WarningCollection

            self._warnings = WarningCollection()
        return self._warnings

    

    

    @property
    def part_names(self) -> tuple[str, ...]:
        """The session's current parts in package order.

        Tolerated ZIP metadata is never edited, so an entry is a part when the
        current graph admits it or a pending edit (such as a builder part) writes it.
        """
        return tuple(
            name for name in self._entries if name in self._graph.parts or name in self._changed
        )

    

    

    

    def insert_story(self, story: Node, *, section: int = 1, variant: str = "default") -> Node:
        """Install a detached header/footer clone into an unoccupied section slot.

        Relationships remain local package references. The source physical part
        is not overwritten; a deterministic new part is allocated. Remove the
        existing story or its section reference before reusing an occupied slot.
        """
        import posixpath

        from ..inspect import serialize_xml_with_preserved_qnames
        from ..package import PKG_REL

        if (
            not isinstance(story, Node)
            or story.node_type != NodeType.STORY
            or story.story_type is None
            or story.story_type.header_footer_kind is None
            or story.document is not self
            or story.parent is not None
            or story._live
            or story._element is None
            or story.part not in self._graph.parts
        ):
            raise NodeEditError("node_story_insert_invalid")
        if not isinstance(section, int) or isinstance(section, bool) or section < 1:
            raise NodeEditError("node_section_invalid")
        if variant not in {"default", "first", "even"}:
            raise NodeEditError("node_story_variant_invalid")
        main = _main_part(self._graph)
        _, cached = self._editable(main)
        root = copy.deepcopy(cached)
        ns = root.tag[1:].split("}", 1)[0]
        sections = list(root.iter(Q(ns, "sectPr")))
        if section > len(sections):
            raise NodeEditError("node_section_invalid")
        target = sections[section - 1]
        kind = story.story_type.header_footer_kind
        reference_tag = Q(ns, kind + "Reference")
        if any(item.tag == reference_tag and item.get(Q(ns, "type")) == variant for item in target):
            raise NodeEditError("node_story_slot_occupied")
        suffix = 1
        directory = posixpath.dirname(main)
        while (part := f"{directory}/{kind}{suffix}.xml") in self._entries:
            suffix += 1
        relation_namespace = (
            _STRICT_BASE + "officeDocument/relationships"
            if ns.startswith(_STRICT_BASE)
            else _TRANSITIONAL_BASE + "officeDocument/2006/relationships"
        )
        rel_part = posixpath.join(directory, "_rels", posixpath.basename(main) + ".rels")
        rel_root = (
            copy.deepcopy(self._editable(rel_part)[1])
            if rel_part in self._entries
            else ET.Element(Q(PKG_REL, "Relationships"))
        )
        used = {item.get("Id") for item in rel_root}
        serial = 1
        while (rid := f"rId{serial}") in used:
            serial += 1
        ET.SubElement(
            rel_root,
            Q(PKG_REL, "Relationship"),
            {
                "Id": rid,
                "Type": relation_namespace + "/" + kind,
                "Target": posixpath.basename(part),
            },
        )
        reference = ET.Element(
            reference_tag, {Q(ns, "type"): variant, Q(relation_namespace, "id"): rid}
        )
        # Header/footer references precede all other section properties.
        position = 0
        while position < len(target) and target[position].tag in {
            Q(ns, "headerReference"),
            Q(ns, "footerReference"),
        }:
            position += 1
        target.insert(position, reference)
        if variant == "first":
            from ..section_break import _insert_sect_pr_child

            first_page = target.find(Q(ns, "titlePg"))
            if first_page is None:
                _insert_sect_pr_child(target, ET.Element(Q(ns, "titlePg")))
            else:
                first_page.set(Q(ns, "val"), "true")
        content_root = copy.deepcopy(self._editable("[Content_Types].xml")[1])
        content_ns = content_root.tag[1:].split("}", 1)[0]
        ET.SubElement(
            content_root,
            Q(content_ns, "Override"),
            {
                "PartName": "/" + part,
                "ContentType": (
                    "application/vnd.openxmlformats-officedocument.wordprocessingml."
                    + kind
                    + "+xml"
                ),
            },
        )
        story_settings: dict[str, bytes] = {}
        if variant == "even":
            from ..protection import _insert_settings_child

            settings_part = self._definition_part("settings")
            if settings_part is None:
                settings_part = directory + "/settings.xml"
                if settings_part in self._entries:
                    raise NodeEditError("node_story_settings_unbound")
                settings = ET.Element(Q(ns, "settings"))
                ET.SubElement(
                    content_root,
                    Q(content_ns, "Override"),
                    {
                        "PartName": "/" + settings_part,
                        "ContentType": (
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
                        ),
                    },
                )
                while (settings_id := f"rId{serial}") in {item.get("Id") for item in rel_root}:
                    serial += 1
                ET.SubElement(
                    rel_root,
                    Q(PKG_REL, "Relationship"),
                    {
                        "Id": settings_id,
                        "Type": relation_namespace + "/settings",
                        "Target": "settings.xml",
                    },
                )
            else:
                settings = copy.deepcopy(self._editable(settings_part)[1])
            flag = settings.find(Q(ns, "evenAndOddHeaders"))
            if flag is None or flag.get(Q(ns, "val"), "true") in {"0", "false", "off"}:
                if flag is None:
                    _insert_settings_child(settings, ET.Element(Q(ns, "evenAndOddHeaders")), ns)
                else:
                    flag.set(Q(ns, "val"), "true")
                story_settings[settings_part] = serialize_xml_with_preserved_qnames(settings)
        incoming = copy.deepcopy(story._element)
        if incoming.tag != Q(ns, "hdr" if kind == "header" else "ftr"):
            raise NodeEditError("node_story_namespace_invalid")
        source_ids = {relation.id for relation in self._graph.parts[story.part].relationships}
        used_ids = {
            value
            for item in incoming.iter()
            for key, value in item.attrib.items()
            if key.startswith("{") and key[1:].partition("}")[0].endswith("/relationships")
        }
        if any(
            key.startswith("{")
            and key[1:].split("}", 1)[0].endswith("/relationships")
            and value not in source_ids
            for item in incoming.iter()
            for key, value in item.attrib.items()
        ):
            raise NodeEditError("node_relationship_missing")
        from ..assembly import _remap_bookmarks

        try:
            _remap_bookmarks(
                {name: self.part_bytes(name) for name in self.part_names if name.endswith(".xml")},
                [incoming],
                ns,
            )
        except ValueError as exc:
            raise NodeEditError("node_identity_invalid") from exc
        _validate_node_structure(incoming, ns)
        changed = {
            **self._changed,
            main: serialize_xml_with_preserved_qnames(root),
            rel_part: serialize_xml_with_preserved_qnames(rel_root),
            "[Content_Types].xml": serialize_xml_with_preserved_qnames(content_root),
            part: serialize_xml_with_preserved_qnames(incoming),
        }
        changed.update(story_settings)
        source_rels = posixpath.join(
            posixpath.dirname(story.part), "_rels", posixpath.basename(story.part) + ".rels"
        )
        if source_rels in self._entries:
            relationship_root = copy.deepcopy(self._editable(source_rels)[1])
            by_id = {item.id: item for item in self._graph.parts[story.part].relationships}
            for item in list(relationship_root):
                if item.get("Id") not in used_ids:
                    relationship_root.remove(item)
                    continue
                relation = by_id.get(item.get("Id"))
                if relation is not None and not relation.is_external:
                    item.set("Target", posixpath.relpath(relation.resolved_target(), directory))
            new_rels = posixpath.join(directory, "_rels", posixpath.basename(part) + ".rels")
            changed[new_rels] = serialize_xml_with_preserved_qnames(relationship_root)
        buffer = BytesIO(self._write(self._source(), changed))
        with zipfile.ZipFile(buffer, "a") as package:
            for name, data in changed.items():
                if name not in self._entries:
                    info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_DEFLATED
                    package.writestr(info, data)
        staged = buffer.getvalue()
        verified = _verify_staged_changes(
            staged,
            {name: data for name, data in changed.items() if not _structural(name)},
            limits=self._graph.limits,
            unchanged_identical=lambda graph: self._unchanged_raw_bytes(graph, staged, changed),
            unchanged_elements=0,
            baseline_tree=self._current_tree,
            check_notes=True,
        )
        if verified is None:
            raise NodeEditError("node_staged_verification_failed")
        self._adopt("node-insert-story", staged, {"section": section, "variant": variant})
        return next(node for node in self.node.children if node.part == part)

    @property
    def stories(self) -> tuple[Node, ...]:
        """Physical stories in visit order, including glossary entry bodies.

        Glossary metadata remains in the navigation tree; text boxes stay inline.
        """
        return tuple(
            story
            for child in self.node.children
            for story in (
                (child,)
                if child.node_type == NodeType.STORY
                else tuple(
                    node
                    for node in child.get_child_nodes(NodeType.STORY, True)
                    if node.story_type == StoryType.GLOSSARY
                )
            )
        )

    def get_story(self, story_type: StoryType, *, index: int = 0) -> Node:
        """Select a physical story of a type, with a zero-based occurrence index."""
        if (
            not isinstance(story_type, StoryType)
            or not isinstance(index, int)
            or isinstance(index, bool)
            or index < 0
        ):
            raise NodeEditError("node_story_target_invalid")
        matches = tuple(story for story in self.stories if story.story_type == story_type)
        if index >= len(matches):
            raise NodeEditError("node_story_target_missing")
        return matches[index]

    def _bind_imported_node(
        self, element: ET.Element, part: str, node_type: NodeType, story_type: StoryType | None
    ) -> Node:
        snapshot = self.node
        assert snapshot._bind is not None
        imported = snapshot._bind(element, part, story_type)
        if imported.node_type != node_type:
            raise NodeEditError("node_import_type_mismatch")
        return imported

    def _definition_part(self, kind: str) -> str | None:
        return next(
            (
                relation.resolved_target()
                for relation in self._graph.parts[_main_part(self._graph)].relationships
                if not relation.is_external and relation.type.endswith("/" + kind)
            ),
            None,
        )

    @property
    def styles(self) -> ET.Element | None:
        """Lossless copy of all style definitions, defaults and latent styles.

        Assign the edited tree back to persist it; partial StyleInfo records are
        deliberately avoided so no paragraph/table/conditional properties vanish.
        """
        part = self._definition_part("styles")
        return None if part is None else copy.deepcopy(self._editable(part)[1])

    @styles.setter
    def styles(self, value: ET.Element) -> None:
        self._replace_definitions("styles", value)

    @property
    def lists(self) -> ET.Element | None:
        """Lossless copy of concrete/abstract lists, levels, overrides and bullets."""
        part = self._definition_part("numbering")
        return None if part is None else copy.deepcopy(self._editable(part)[1])

    @lists.setter
    def lists(self, value: ET.Element) -> None:
        self._replace_definitions("numbering", value)

    def _replace_definitions(self, kind: str, value: ET.Element) -> None:
        from ..assembly import _add_document_relationship, _ensure_override
        from ..inspect import serialize_xml_with_preserved_qnames

        main = _main_part(self._graph)
        ns = self._editable(main)[1].tag[1:].partition("}")[0]
        if not isinstance(value, ET.Element) or value.tag != Q(ns, kind):
            raise NodeEditError("document_definition_root_invalid")
        payload = serialize_xml_with_preserved_qnames(value)
        self._parse(payload, editable=False)
        parts = {name: self.part_bytes(name) for name in self.part_names}
        part = self._definition_part(kind)
        added = part is None
        if added:
            if main != "word/document.xml":
                raise NodeEditError("document_definition_layout_unsupported")
            part = "word/" + kind + ".xml"
            if part in parts:
                raise NodeEditError("document_definition_part_unbound")
            relation_ns = (
                _STRICT_BASE + "officeDocument/relationships"
                if ns.startswith(_STRICT_BASE)
                else _TRANSITIONAL_BASE + "officeDocument/2006/relationships"
            )
            _ensure_override(
                parts,
                part,
                "application/vnd.openxmlformats-officedocument.wordprocessingml." + kind + "+xml",
            )
            _add_document_relationship(parts, relation_ns + "/" + kind, part)
        assert part is not None
        # Protect references even when the removed definition is used only by a
        # dependency in another style or an override in numbering.
        old = ET.Element(Q(ns, kind)) if added else self._editable(part)[1]
        id_tag = "style" if kind == "styles" else "num"
        id_attribute = "styleId" if kind == "styles" else "numId"
        old_ids = {item.get(Q(ns, id_attribute)) for item in old.findall(Q(ns, id_tag))}
        new_ids = [item.get(Q(ns, id_attribute)) for item in value.findall(Q(ns, id_tag))]
        if None in new_ids or len(new_ids) != len(set(new_ids)):
            raise NodeEditError("document_definition_identity_invalid")
        removed = old_ids - set(new_ids)
        references = (
            {"pStyle", "rStyle", "tblStyle", "basedOn", "next", "link", "styleLink", "numStyleLink"}
            if kind == "styles"
            else {"numId"}
        )
        parts[part] = payload
        for name, data in parts.items():
            if not name.endswith(".xml"):
                continue
            root = value if name == part else self._parse(data, editable=False)
            if any(
                item.tag in {Q(ns, tag) for tag in references} and item.get(Q(ns, "val")) in removed
                for item in root.iter()
            ):
                raise NodeEditError("document_definition_referenced")
        if kind == "numbering":
            abstract_ids = [
                item.get(Q(ns, "abstractNumId")) for item in value.findall(Q(ns, "abstractNum"))
            ]
            if (
                None in abstract_ids
                or len(abstract_ids) != len(set(abstract_ids))
                or any(
                    item.get(Q(ns, "val")) not in abstract_ids
                    for item in value.iter(Q(ns, "abstractNumId"))
                )
            ):
                raise NodeEditError("document_list_definition_invalid")
        changed = {
            **self._changed,
            **{
                name: data
                for name, data in parts.items()
                if name in self._entries and data != self.part_bytes(name)
            },
        }
        staged = self._write(self._source(), changed)
        if added:
            buffer = BytesIO(staged)
            with zipfile.ZipFile(buffer, "a") as archive:
                archive.writestr(_added_entry(zipfile.ZipInfo(part)), payload)
            staged = buffer.getvalue()
        if (
            _verify_staged_changes(
                staged,
                {name: data for name, data in changed.items() if not _structural(name)}
                | {part: payload},
                limits=self._graph.limits,
                unchanged_identical=lambda graph: self._unchanged_raw_bytes(graph, staged, changed),
                unchanged_elements=0,
                baseline_tree=self._current_tree,
                check_notes=True,
            )
            is None
        ):
            raise NodeEditError("document_definition_verification_failed")
        self._adopt("document-" + kind, staged, {})

    @property
    def background(self) -> str | None:
        """Page background RGB (or ``auto``); background_xml retains full markup."""
        element = self.background_xml
        if element is None:
            return None
        ns = element.tag[1:].partition("}")[0]
        return element.get(Q(ns, "color"))

    @background.setter
    def background(self, color: str | None) -> None:
        if color is not None and (
            not isinstance(color, str) or not re.fullmatch(r"[0-9A-Fa-f]{6}|auto", color)
        ):
            raise NodeEditError("document_background_color_invalid")
        main = _main_part(self._graph)
        ns = self._editable(main)[1].tag[1:].partition("}")[0]
        element = self.background_xml
        if color is None:
            self.background_xml = None
        else:
            if element is None:
                element = ET.Element(Q(ns, "background"))
            element.set(Q(ns, "color"), color.upper() if color != "auto" else color)
            self.background_xml = element

    @property
    def background_xml(self) -> ET.Element | None:
        """Lossless background markup, including theme and VML extensions."""
        main = _main_part(self._graph)
        root = self._editable(main)[1]
        ns = root.tag[1:].partition("}")[0]
        element = root.find(Q(ns, "background"))
        return None if element is None else copy.deepcopy(element)

    @background_xml.setter
    def background_xml(self, value: ET.Element | None) -> None:
        from ..inspect import serialize_xml_with_preserved_qnames

        main = _main_part(self._graph)
        root = copy.deepcopy(self._editable(main)[1])
        ns = root.tag[1:].partition("}")[0]
        if value is not None and (
            not isinstance(value, ET.Element) or value.tag != Q(ns, "background")
        ):
            raise NodeEditError("document_background_invalid")
        for element in root.findall(Q(ns, "background")):
            root.remove(element)
        if value is not None:
            root.insert(0, copy.deepcopy(value))
        payload = serialize_xml_with_preserved_qnames(root)
        _validate_node_structure(self._parse(payload, editable=False), ns)
        changed = {**self._changed, main: payload}
        staged = self._write(self._source(), changed)
        if (
            _verify_staged_changes(
                staged,
                {main: payload},
                limits=self._graph.limits,
                unchanged_identical=lambda graph: self._unchanged_raw_bytes(graph, staged, changed),
                unchanged_elements=0,
                baseline_tree=self._current_tree,
                check_notes=True,
            )
            is None
        ):
            raise NodeEditError("document_background_verification_failed")
        self._set_part(main, payload)
        self._record("document-background", (main,))

    @property
    def node(self) -> Node:
        """The session's navigation snapshot; story order matches :meth:`accept`.

        Metadata/unknown XML is exposed as XML nodes with its expanded name.
        w:t is a run's text payload rather than a separate child. No formatting,
        field, revision or unsupported drawing markup is dropped from the view.
        A relationship-resolved glossaryDocument follows the other parts as an
        XML subtree. Ordered docPart entries and metadata remain XML nodes;
        each docPartBody is a GLOSSARY story with semantic content descendants.
        Complex characters are FieldStart/FieldSeparator/FieldEnd subclasses;
        cloning and generic atomic edits preserve their XML and marker flags.
        """
        from ..authoring import (
            _MAX_TEXT_CHARACTERS,
            _text_replacements_preserve_whitespace,
        )
        from ..inspect import serialize_xml_with_preserved_qnames
        from ..paragraph_authoring import _is_valid_xml_text

        self._source()
        main = _main_part(self._graph)
        contexts: dict[str, tuple[bytes, ET.Element, Callable[[ET.Element, str], None]]] = {}
        current_bytes: dict[str, bytes] = {}
        detached_elements: set[int] = set()

        def validate(
            staged_root: ET.Element,
            source_root: ET.Element,
            target: ET.Element,
            memo: dict[int, Any],
            ns: str,
        ) -> set[int]:
            # Existing opaque entries are exposed, not newly admitted by this
            # edit. Only the affected entry and newly inserted content are gated.
            opaque: set[int] = set()
            if source_root.tag == Q(ns, "glossaryDocument"):
                for entry in source_root.iter(Q(ns, "docPart")):
                    if any(item is target for item in entry.iter()):
                        continue
                    staged = memo.get(id(entry))
                    if staged is not None:
                        opaque.update(id(item) for item in staged.iter())
            _validate_node_structure(staged_root, ns, opaque_elements=opaque)
            return opaque

        def parse_edit(candidate: bytes) -> ET.Element:
            try:
                return self._parse(candidate, editable=False)
            except _PART_ERRORS as exc:
                raise NodeEditError("node_xml_invalid") from exc

        def commit(part: str, candidate: bytes) -> None:
            parse_edit(candidate)
            changed = {**self._changed, part: candidate}
            staged = self._write(self._source(), changed)
            verified = _verify_staged_changes(
                staged,
                {name: data for name, data in changed.items() if not _structural(name)},
                limits=self._graph.limits,
                unchanged_identical=lambda graph: self._unchanged_raw_bytes(graph, staged, changed),
                unchanged_elements=0,
                baseline_tree=self._current_tree,
                check_notes=True,
            )
            if verified is None:
                raise NodeEditError("node_staged_verification_failed")
            self._set_part(part, candidate)
            current_bytes[part] = self.part_bytes(part)

        def edit(operation: str, parent: Node, child: Node, index: int) -> None:
            if parent.document is not self or child.document is not self:
                raise NodeEditError("node_document_mismatch")
            if operation == "check":
                if parent._live and (
                    parent.part not in self._entries
                    or self.part_bytes(parent.part) != current_bytes.get(parent.part)
                ):
                    raise NodeEditError("node_snapshot_stale")
                return
            if operation in {"clear", "replace-content"}:
                if (
                    parent.node_type == NodeType.DOCUMENT
                    or parent._element is None
                    or parent.part is None
                ):
                    raise NodeEditError("node_document_structure_required")
                part = parent.part
                root, _ = context(part)
                attached = parent._live
                if attached and (
                    part not in self._entries
                    or self.part_bytes(part) != current_bytes[part]
                    or not any(item is parent._element for item in root.iter())
                ):
                    raise NodeEditError("node_snapshot_stale")
                memo: dict[int, Any] = {}
                staged_root = copy.deepcopy(root if attached else parent._element, memo)
                staged_parent = memo[id(parent._element)]
                staged_parent[:] = (
                    copy.deepcopy(list(child._element)) if operation == "replace-content" else []
                )
                staged_parent.text = None
                candidate = serialize_xml_with_preserved_qnames(staged_root)
                ns = root.tag[1:].partition("}")[0]
                parse_edit(candidate)
                validate(
                    staged_root, root if attached else parent._element, parent._element, memo, ns
                )
                if attached:
                    commit(part, candidate)
                old_children = tuple(parent.children)
                parent._element[:] = list(child._element) if operation == "replace-content" else []
                parent._element.text = None
                parent._attach(tuple(child.children) if operation == "replace-content" else ())
                for old in old_children:
                    pending = [old]
                    while pending:
                        item = pending.pop()
                        item._live = False
                        if item._element is not None:
                            detached_elements.add(id(item._element))
                        pending.extend(item.children)
                pending = list(parent.children)
                while pending:
                    item = pending.pop()
                    item._live = attached
                    if item._element is not None:
                        detached_elements.discard(id(item._element))
                    pending.extend(item.children)
                if attached:
                    self._record("node-" + operation, (part,))
                return
            if (
                parent.node_type == NodeType.DOCUMENT
                and child.story_type is not None
                and child.story_type.header_footer_kind is not None
                and operation == "remove"
            ):
                main_root, _ = context(main)
                if main not in self._entries or self.part_bytes(main) != current_bytes[main]:
                    raise NodeEditError("node_snapshot_stale")
                ids = {
                    relation.id
                    for relation in self._graph.parts[main].relationships
                    if not relation.is_external and relation.resolved_target() == child.part
                }
                memo: dict[int, Any] = {}
                staged_root = copy.deepcopy(main_root, memo)
                ns = main_root.tag[1:].split("}", 1)[0]
                removed: list[ET.Element] = []
                for section in main_root.iter(Q(ns, "sectPr")):
                    for reference in section:
                        if reference.tag in {
                            Q(ns, "headerReference"),
                            Q(ns, "footerReference"),
                        } and any(
                            key.endswith("}id") and value in ids
                            for key, value in reference.attrib.items()
                        ):
                            memo[id(section)].remove(memo[id(reference)])
                            removed.append(reference)
                commit(main, serialize_xml_with_preserved_qnames(staged_root))
                for section in main_root.iter(Q(ns, "sectPr")):
                    for reference in removed:
                        if reference in list(section):
                            section.remove(reference)
                for node in parent.get_child_nodes(NodeType.HEADER_FOOTER, True):
                    if node._element in removed and node.parent is not None:
                        node.parent._attach(
                            tuple(item for item in node.parent.children if item is not node)
                        )
                        node._live = False
                parent._attach(parent.children[:index] + parent.children[index + 1 :])
                pending = [child]
                while pending:
                    node = pending.pop()
                    node._live = False
                    if node._element is not None:
                        detached_elements.add(id(node._element))
                    pending.extend(node.children)
                self._record("node-remove-story", (main,))
                return
            if parent.node_type == NodeType.DOCUMENT:
                if child.story_type not in {
                    StoryType.FOOTNOTE,
                    StoryType.ENDNOTE,
                    StoryType.COMMENT,
                }:
                    raise NodeEditError("node_story_root_required")
                part = child.part
                expected = {
                    StoryType.FOOTNOTE: "footnote",
                    StoryType.ENDNOTE: "endnote",
                    StoryType.COMMENT: "comment",
                }[child.story_type]
            else:
                part = parent.part
            if (
                part is None
                or part not in contexts
                or parent._element is None
                and parent.node_type != NodeType.DOCUMENT
            ):
                raise NodeEditError("node_target_invalid")
            root, _ = context(part)
            ns = root.tag[1:].split("}", 1)[0]
            if parent._live and (
                part not in self._entries or self.part_bytes(part) != current_bytes[part]
            ):
                raise NodeEditError("node_snapshot_stale")
            if child._element is None:
                raise NodeEditError("node_child_invalid")
            if parent.node_type == NodeType.DOCUMENT:
                if child.xml_name != Q(ns, expected):
                    raise NodeEditError("node_story_insert_invalid")
                if operation == "remove" and child.story_type == StoryType.COMMENT:
                    identifier = child._element.get(Q(ns, "id"))
                    if any(
                        item.get(Q(ns, "id")) == identifier
                        for _, story_root, _ in contexts.values()
                        for tag in ("commentReference", "commentRangeStart", "commentRangeEnd")
                        for item in story_root.iter(Q(ns, tag))
                    ):
                        raise NodeEditError("node_story_referenced")
            container = root if parent.node_type == NodeType.DOCUMENT else parent._element
            assert container is not None
            attached = parent._live
            if attached and not any(item is container for item in root.iter()):
                raise NodeEditError("node_snapshot_stale")
            memo: dict[int, Any] = {}
            staged_root = copy.deepcopy(root if attached else container, memo)
            staged_container = memo[id(container)]
            incoming = copy.deepcopy(child._element)
            if operation == "insert":
                if child.part != part:
                    source_relations = {
                        item.id: item for item in self._graph.parts[child.part].relationships
                    }
                    destination_relations = self._graph.parts[part].relationships
                    for item in incoming.iter():
                        for key, value in tuple(item.attrib.items()):
                            if not key.startswith("{") or not key[1:].split("}", 1)[0].endswith(
                                "/relationships"
                            ):
                                continue
                            relation = source_relations.get(value)
                            equivalent = next(
                                (
                                    other
                                    for other in destination_relations
                                    if relation is not None
                                    and other.type == relation.type
                                    and other.is_external == relation.is_external
                                    and other.resolved_target() == relation.resolved_target()
                                ),
                                None,
                            )
                            if equivalent is None:
                                raise NodeEditError("node_relationship_scope_invalid")
                            item.set(key, equivalent.id)
                from ..assembly import (
                    _PARA_ID,
                    _anchor_ids,
                    _drawing_ids,
                    _remap_anchor_ids,
                    _remap_bookmarks,
                    _remap_drawing_ids,
                    _remap_paragraph_ids,
                    _remap_sdt_ids,
                    _remap_vml_ids,
                    _vml_ids,
                )

                try:
                    _remap_drawing_ids([incoming], _drawing_ids(root))
                    _remap_anchor_ids([incoming], _anchor_ids(root))
                    _remap_vml_ids([incoming], ns, _vml_ids(root))
                    _remap_paragraph_ids(
                        [incoming],
                        ns,
                        {value.upper() for item in root.iter() if (value := item.get(_PARA_ID))},
                    )
                    _remap_sdt_ids(
                        [incoming],
                        ns,
                        {
                            int(item.get(Q(ns, "val"), "0"))
                            for item in root.findall(f".//{Q(ns, 'sdtPr')}/{Q(ns, 'id')}")
                            if item.get(Q(ns, "val"), "").isdigit()
                        },
                    )
                    if any(item.tag == Q(ns, "bookmarkStart") for item in incoming.iter()):
                        _remap_bookmarks(
                            {name: self.part_bytes(name) for name in contexts}, [incoming], ns
                        )
                    elif any(item.tag == Q(ns, "bookmarkEnd") for item in incoming.iter()):
                        raise NodeEditError("node_bookmark_range_incomplete")
                    from ..review import (
                        _CONFLICT_REVISION_TAGS,
                        _SELECTABLE_REVISION_TAGS,
                    )

                    revision_tags = {
                        Q(ns, tag) for tag in _SELECTABLE_REVISION_TAGS
                    } | _CONFLICT_REVISION_TAGS
                    occupied = {
                        int(value)
                        for item in root.iter()
                        if item.tag in revision_tags
                        and (value := item.get(Q(ns, "id"), "")).isdigit()
                    }
                    for item in incoming.iter():
                        if item.tag not in revision_tags:
                            continue
                        value = item.get(Q(ns, "id"), "")
                        if not value.isdigit():
                            raise NodeEditError("node_identity_invalid")
                        identifier = int(value)
                        if identifier in occupied:
                            identifier = max(occupied, default=0) + 1
                            item.set(Q(ns, "id"), str(identifier))
                        occupied.add(identifier)
                    permissions = list(incoming.iter(Q(ns, "permStart")))
                    permission_ends = list(incoming.iter(Q(ns, "permEnd")))
                    if permissions or permission_ends:
                        identifiers = [item.get(Q(ns, "id")) for item in permissions]
                        if (
                            None in identifiers
                            or len(identifiers) != len(set(identifiers))
                            or set(identifiers)
                            != {item.get(Q(ns, "id")) for item in permission_ends}
                        ):
                            raise NodeEditError("node_range_incomplete")
                        used = {item.get(Q(ns, "id")) for item in root.iter(Q(ns, "permStart"))}
                        next_id = (
                            max(
                                (int(value) for value in used if value and value.isdigit()),
                                default=0,
                            )
                            + 1
                        )
                        remapped = {}
                        for item in permissions:
                            old = item.get(Q(ns, "id"))
                            remapped[old] = str(next_id)
                            item.set(Q(ns, "id"), str(next_id))
                            next_id += 1
                        for item in permission_ends:
                            item.set(Q(ns, "id"), remapped[item.get(Q(ns, "id"))])
                except ValueError as exc:
                    if isinstance(exc, NodeEditError):
                        raise
                    raise NodeEditError("node_identity_invalid") from exc
                if parent.node_type == NodeType.DOCUMENT:
                    identifier = incoming.get(Q(ns, "id"))
                    if identifier is None or not identifier.isdigit():
                        raise NodeEditError("node_story_identity_invalid")
                    used = {item.get(Q(ns, "id")) for item in root}
                    incoming.set(
                        Q(ns, "id"),
                        str(
                            max(
                                (int(value) for value in used if value and value.isdigit()),
                                default=0,
                            )
                            + 1
                        ),
                    )
                visible = parent.children
                if parent.node_type == NodeType.DOCUMENT:
                    following = next((item for item in visible[index:] if item.part == part), None)
                else:
                    following = visible[index] if index < len(visible) else None
                xml_index = (
                    list(container).index(following._element)
                    if following is not None
                    else len(container)
                )
                staged_container.insert(xml_index, incoming)
            else:
                staged_container.remove(memo[id(child._element)])
            candidate = serialize_xml_with_preserved_qnames(staged_root)
            parse_edit(candidate)
            opaque = validate(staged_root, root if attached else container, container, memo, ns)
            if attached:
                relationships = {item.id for item in self._graph.parts[part].relationships}
                for item in staged_root.iter():
                    if id(item) in opaque:
                        continue
                    for key, value in item.attrib.items():
                        if (
                            key.startswith("{")
                            and key[1:].split("}", 1)[0].endswith("/relationships")
                            and value not in relationships
                        ):
                            raise NodeEditError("node_relationship_missing")
            if attached:
                commit(part, candidate)
            if operation == "insert":
                for original, remapped in zip(child._element.iter(), incoming.iter(), strict=True):
                    original.attrib.clear()
                    original.attrib.update(remapped.attrib)
                container.insert(xml_index, child._element)
                parent._attach(parent.children[:index] + (child,) + parent.children[index:])
            else:
                container.remove(child._element)
                parent._attach(parent.children[:index] + parent.children[index + 1 :])
            pending = [(child, parent.story_type or child.story_type)]
            while pending:
                node, inherited = pending.pop()
                node._live = attached and operation == "insert"
                node._part = part
                node._document = self
                node._edit = edit
                node._rewrite = context(part)[1]
                if operation == "insert" and node.node_type != NodeType.STORY:
                    node._story_type = inherited
                pending.extend((item, node.story_type) for item in node.children)
            if attached:
                self._record("node-" + operation, (part,))

        def clone_document(deep: bool) -> Node:
            for part, current in current_bytes.items():
                if part not in self._entries or self.part_bytes(part) != current:
                    raise NodeEditError("node_snapshot_stale")
            clone = Document.from_bytes(
                self._materialize(), path=self.path, limits=self._graph.limits
            )
            if not deep:
                from ..inspect import serialize_xml_with_preserved_qnames

                clone_main = _main_part(clone._graph)
                root = copy.deepcopy(clone._editable(clone_main)[1])
                ns = root.tag[1:].partition("}")[0]
                body = root.find(Q(ns, "body"))
                assert body is not None
                body[:] = [item for item in body if item.tag == Q(ns, "sectPr")]
                for section in body.iter(Q(ns, "sectPr")):
                    section[:] = [
                        item
                        for item in section
                        if item.tag not in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                    ]
                rel_part = (
                    clone_main.rpartition("/")[0]
                    + "/_rels/"
                    + clone_main.rpartition("/")[2]
                    + ".rels"
                )
                if rel_part in clone._entries:
                    rel_root = copy.deepcopy(clone._editable(rel_part)[1])
                    for relation in list(rel_root):
                        if relation.get("Type", "").rsplit("/", 1)[-1] in {
                            "header",
                            "footer",
                            "footnotes",
                            "endnotes",
                            "comments",
                            "glossaryDocument",
                        }:
                            rel_root.remove(relation)
                    clone._set_part(rel_part, serialize_xml_with_preserved_qnames(rel_root))
                clone._set_part(clone_main, serialize_xml_with_preserved_qnames(root))
                clone._adopt("node-shallow-document-clone", clone._materialize(), {})
            return clone.node

        def context(part: str) -> tuple[ET.Element, Callable[[ET.Element, str], None]]:
            if part not in contexts:
                original, cached = self._editable(part)
                root = copy.deepcopy(cached)
                current_bytes[part] = original
                ns = root.tag[1:].split("}", 1)[0]

                def rewrite(run: ET.Element, text: str) -> None:
                    if not isinstance(text, str):
                        raise TypeError("run_text_invalid")
                    if not _is_valid_xml_text(text):
                        raise NodeEditError("run_text_invalid")
                    if len(text) > _MAX_TEXT_CHARACTERS:
                        raise NodeEditError("run_text_limit_exceeded")
                    attached = id(run) not in detached_elements and any(
                        item is run for item in root.iter()
                    )
                    if attached and (
                        part not in self._entries or self.part_bytes(part) != current_bytes[part]
                    ):
                        raise NodeEditError("node_snapshot_stale")
                    texts = list(run.findall(Q(ns, "t")))
                    if "".join(item.text or "" for item in texts) == text:
                        return
                    if not attached:
                        if texts:
                            texts[0].text = text
                            texts[0].set(
                                Q("http://www.w3.org/XML/1998/namespace", "space"), "preserve"
                            )
                            for item in texts[1:]:
                                item.text = ""
                        else:
                            payload = ET.SubElement(run, Q(ns, "t"))
                            payload.text = text
                            payload.set(
                                Q("http://www.w3.org/XML/1998/namespace", "space"), "preserve"
                            )
                        return
                    if not texts or any(len(item) for item in texts):
                        raise NodeEditError("run_text_target_missing")
                    # Keep all w:t elements and special characters. Distribute over
                    # the old segment lengths, putting the remainder in the last.
                    values: dict[int, str] = {}
                    offset = 0
                    for index, item in enumerate(texts):
                        stop = (
                            len(text) if index + 1 == len(texts) else offset + len(item.text or "")
                        )
                        values[id(item)] = text[offset:stop]
                        offset = stop
                    if not _text_replacements_preserve_whitespace(root, values):
                        raise NodeEditError("run_text_space_not_preserved")
                    ordinals = {
                        id(item): ordinal for ordinal, item in enumerate(root.iter(Q(ns, "t")), 1)
                    }
                    replacements = {
                        ordinals[id(item)]: values[id(item)]
                        for item in texts
                        if values[id(item)] != (item.text or "")
                    }
                    candidate = _patch_node_texts(current_bytes[part], ns, replacements)
                    parsed = self._parse(candidate, editable=False)
                    observed = list(parsed.iter(Q(ns, "t")))
                    if any(
                        (observed[index - 1].text or "") != value
                        for index, value in replacements.items()
                    ):
                        raise NodeEditError("run_text_staged_verification_failed")
                    for item in texts:
                        item.text = values[id(item)]
                    self._set_part(part, candidate)
                    if (
                        part in self._source_graph.parts
                        and candidate == self._source_graph.read_part_bytes(part)
                    ):
                        self._changed.pop(part, None)
                    current_bytes[part] = self.part_bytes(part)
                    self._record("visitor-run-text", (part,))

                contexts[part] = (original, root, rewrite)
            _, root, rewrite = contexts[part]
            return root, rewrite

        types = {
            "p": NodeType.PARAGRAPH,
            "r": NodeType.RUN,
            "tbl": NodeType.TABLE,
            "tr": NodeType.ROW,
            "tc": NodeType.CELL,
            "fldSimple": NodeType.FIELD,
            "fldChar": NodeType.FIELD,
            "bookmarkStart": NodeType.BOOKMARK,
            "bookmarkEnd": NodeType.BOOKMARK,
            "commentRangeStart": NodeType.COMMENT,
            "commentRangeEnd": NodeType.COMMENT,
            "commentReference": NodeType.COMMENT,
            "drawing": NodeType.SHAPE,
            "pict": NodeType.SHAPE,
            "object": NodeType.SHAPE,
            "sdt": NodeType.STRUCTURED_TAG,
            "sectPr": NodeType.SECTION,
            "headerReference": NodeType.HEADER_FOOTER,
            "footerReference": NodeType.HEADER_FOOTER,
            "smartTag": NodeType.SMART_TAG,
            "tab": NodeType.SPECIAL_CHARACTER,
            "br": NodeType.SPECIAL_CHARACTER,
            "cr": NodeType.SPECIAL_CHARACTER,
            "ptab": NodeType.SPECIAL_CHARACTER,
            "noBreakHyphen": NodeType.SPECIAL_CHARACTER,
            "softHyphen": NodeType.SPECIAL_CHARACTER,
        }
        words = {
            _TRANSITIONAL_BASE + "wordprocessingml/2006/main",
            _STRICT_BASE + "wordprocessingml/main",
        }

        def build(
            element: ET.Element, part: str, story: StoryType | None, *, story_root: bool = True
        ) -> Node:
            _, rewrite = context(part)

            def make(item: ET.Element, inherited: StoryType | None) -> Node:
                namespace, _, local = item.tag[1:].partition("}")
                box = namespace in words and local == "txbxContent"
                glossary = namespace in words and local == "docPartBody"
                math = namespace.endswith("/math") and local in {"oMath", "oMathPara"}
                node_class = (
                    {
                        "begin": FieldStart,
                        "separate": FieldSeparator,
                        "end": FieldEnd,
                    }.get(item.get(Q(namespace, "fldCharType")), Node)
                    if namespace in words and local == "fldChar"
                    else Node
                )
                result = node_class(
                    NodeType.STORY
                    if box or glossary
                    else NodeType.OFFICE_MATH
                    if math
                    else types.get(local, NodeType.XML)
                    if namespace in words
                    else NodeType.XML,
                    part=part,
                    element=item,
                    story_type=(
                        StoryType.TEXT_BOX if box else StoryType.GLOSSARY if glossary else inherited
                    ),
                    rewrite=rewrite,
                    edit=edit,
                    document=self,
                )
                result._bind = bind
                return result

            top = (
                Node(
                    NodeType.STORY,
                    part=part,
                    element=element,
                    story_type=story,
                    rewrite=rewrite,
                    edit=edit,
                    document=self,
                )
                if story_root
                else make(element, story)
            )
            top._bind = bind
            pending = [(top, element)]
            while pending:
                parent, item = pending.pop()
                children = tuple(
                    make(child, parent.story_type or story)
                    for child in item
                    if not (
                        parent.node_type == NodeType.RUN
                        and child.tag in {Q(ns, "t") for ns in words}
                    )
                )
                parent._attach(children)
                pending.extend(
                    zip(
                        children,
                        (
                            child
                            for child in item
                            if not (
                                parent.node_type == NodeType.RUN
                                and child.tag in {Q(ns, "t") for ns in words}
                            )
                        ),
                        strict=True,
                    )
                )
            return top

        def bind(element: ET.Element, part: str, story: StoryType | None) -> Node:
            if part not in self._entries:
                raise NodeEditError("node_import_part_missing")
            local = element.tag.rpartition("}")[2]
            if local == "docPartBody":
                story = StoryType.GLOSSARY
            top = build(
                element,
                part,
                story,
                story_root=local
                in {
                    "body",
                    "hdr",
                    "ftr",
                    "footnote",
                    "endnote",
                    "comment",
                    "txbxContent",
                    "docPartBody",
                },
            )
            pending = [top]
            while pending:
                item = pending.pop()
                item._live = False
                if item._element is not None:
                    detached_elements.add(id(item._element))
                pending.extend(item.children)
            return top

        root, _ = context(main)
        ns = root.tag[1:].split("}", 1)[0]
        body = root.find(Q(ns, "body"))
        if body is None:
            raise PackageError("document_body_missing", "Main part has no body", main)
        stories = [build(body, main, StoryType.BODY)]
        relationships = self._graph.parts[main].relationships
        by_id = {relationship.id: relationship for relationship in relationships}
        seen: set[str] = set()
        for section in body.iter(Q(ns, "sectPr")):
            for reference in section:
                local = reference.tag.rsplit("}", 1)[-1]
                if reference.tag not in {
                    Q(ns, "headerReference"),
                    Q(ns, "footerReference"),
                }:
                    continue
                rid = next(
                    (value for key, value in reference.attrib.items() if key.endswith("}id")),
                    None,
                )
                relation = by_id.get(rid)
                kind = "header" if local == "headerReference" else "footer"
                if (
                    relation is None
                    or relation.is_external
                    or not relation.type.endswith("/" + kind)
                ):
                    raise PackageError("story_reference_invalid", "Invalid story reference", main)
                part = relation.resolved_target()
                if part not in seen:
                    seen.add(part)
                    story_root, _ = context(part)
                    variant = reference.get(Q(ns, "type"), "default")
                    story_type = {
                        "default": StoryType(kind),
                        "first": StoryType("first_" + kind),
                        "even": StoryType("even_" + kind),
                    }.get(variant)
                    if story_type is None:
                        raise NodeEditError("node_story_variant_invalid")
                    stories.append(build(story_root, part, story_type))
        for kind, tag, story in (
            ("footnotes", "footnote", StoryType.FOOTNOTE),
            ("endnotes", "endnote", StoryType.ENDNOTE),
            ("comments", "comment", StoryType.COMMENT),
        ):
            for relationship in relationships:
                if relationship.is_external or not relationship.type.endswith("/" + kind):
                    continue
                part = relationship.resolved_target()
                if part in seen:
                    continue
                seen.add(part)
                story_root, _ = context(part)
                story_ns = story_root.tag[1:].split("}", 1)[0]
                for item in story_root.findall(Q(story_ns, tag)):
                    note_type = item.get(Q(story_ns, "type"))
                    story_type = (
                        StoryType(
                            story.value
                            + "_"
                            + {
                                "separator": "separator",
                                "continuationSeparator": "continuation_separator",
                                "continuationNotice": "continuation_notice",
                            }[note_type]
                        )
                        if kind != "comments"
                        and note_type
                        in {"separator", "continuationSeparator", "continuationNotice"}
                        else story
                    )
                    stories.append(build(item, part, story_type))
        glossary_relations = [
            relationship
            for relationship in relationships
            if relationship.type
            in {
                _TRANSITIONAL_BASE + "officeDocument/2006/relationships/glossaryDocument",
                _STRICT_BASE + "officeDocument/relationships/glossaryDocument",
            }
        ]
        if len(glossary_relations) > 1:
            raise NodeEditError("node_glossary_ambiguous")
        if glossary_relations:
            relationship = glossary_relations[0]
            part = relationship.resolved_target()
            if relationship.is_external or part not in self._graph.parts:
                raise NodeEditError("node_glossary_reference_invalid")
            glossary_root, _ = context(part)
            if glossary_root.tag != Q(ns, "glossaryDocument"):
                raise NodeEditError("node_glossary_root_invalid")
            stories.append(build(glossary_root, part, None, story_root=False))
        document = Node(NodeType.DOCUMENT, edit=edit, clone_document=clone_document, document=self)
        document._bind = bind
        document._attach(tuple(stories))
        return document

    def accept(self, visitor: DocumentVisitor) -> VisitorAction:
        """Visit typed start/end nodes in deterministic document order.

        Body first; then explicitly referenced headers/footers in section
        (sectPr) order, preserving reference order within each section. Shared
        physical parts are visited once, at their first reference; inherited
        references are not duplicated. Footnotes, endnotes, then comments follow
        in relationship order, each part's entries in XML order (including note
        separators). The glossary XML subtree follows, with entry bodies in XML
        order. Text boxes are inline at their anchor's XML position.
        XML descendants use preorder/start and postorder/end. STOP returns
        immediately; SKIP_CHILDREN at start still emits the matching end.
        Generic clone/insert/remove edits keep navigation coherent and validate
        the staged package before committing. Other session edits invalidate the
        snapshot. Header/footer installation selects a slot via insert_story.
        """
        return self.node.accept(visitor)

    # -- edits ------------------------------------------------------------------

    def append_paragraph(
        self, text: str = "", *, style: str | None = None, part: str = _DOCUMENT
    ) -> ParagraphReport:
        """Insert a paragraph after the last top-level paragraph of ``part``.

        ``text`` and ``style`` follow :class:`~wolfdocx.ParagraphRequest`; the anchor
        and its structure fingerprint are taken from the session's current tree.
        """
        from ..authoring import _cell_structure_fingerprint
        from ..paragraph_authoring import _paragraph_refusal, _resolve_target_story

        self._source()
        probe = self._paragraph_request(1, _UNBOUND_FINGERPRINT, text, style, part)
        root, container, ns, _, err = _resolve_target_story(self._graph, part, self._editable)
        if err is not None:
            return _paragraph_refusal(probe, err)
        assert root is not None and ns is not None
        anchors = (
            [child for child in container if child.tag == Q(ns, "p")]
            if container is not None
            else []
        )
        if not anchors:
            return _paragraph_refusal(probe, "paragraph_target_missing")
        paragraph = next(
            index for index, node in enumerate(root.iter(Q(ns, "p")), 1) if node is anchors[-1]
        )
        request = self._paragraph_request(
            paragraph, _cell_structure_fingerprint(anchors[-1]), text, style, part
        )
        return self._apply_paragraph_request(request)

    def _apply_paragraph_request(self, request: ParagraphRequest) -> ParagraphReport:
        """Plan and verify a paragraph edit before committing its current part bytes."""
        from ..authoring import _APPLIED
        from ..paragraph_authoring import ParagraphReport, _paragraph_refusal

        # Planning edits the cached tree in place, so the cache gives it up first.
        plan = _plan_paragraph_edit(request, self._graph, self._take_editable)
        if isinstance(plan, str):
            return _paragraph_refusal(request, plan)
        try:
            checked = self._parse_checked(plan.serialized)
            landed = plan.verify_story(checked[1])
        except _PART_ERRORS:
            landed = False
        if not landed:
            return _paragraph_refusal(request, "paragraph_staged_verification_failed")
        part = request.part
        touched = (part,) if plan.serialized != self.part_bytes(part) else ()
        if (
            part in self._source_graph.parts
            and plan.serialized == self._source_graph.read_part_bytes(part)
        ):
            # Inverse edits restore the original raw ZIP member, as _adopt does.
            self._changed.pop(part, None)
            self._trees.pop(part, None)
        else:
            self._set_part(part, plan.serialized)
        self._checked[part] = checked
        self._record("paragraph", touched)
        return ParagraphReport(
            status=_APPLIED,
            code=None,
            source_sha256=self.source_sha256,
            target=request.target_dict(),
            action=request.action,
            paragraph=plan.change_report,
            target_change_report=plan.change_report,
            non_target_parts_preserved=True,
        )

    def set_cell_text(
        self,
        *,
        table: int,
        row: int,
        cell: int,
        text: str,
        expected_old_text: str | None = None,
        structure_fingerprint_sha256: str | None = None,
    ) -> CellTextAuthoringReport:
        """Replace one simple table cell's text, as :func:`author_cell_text` would.

        ``expected_old_text`` and ``structure_fingerprint_sha256`` are optional
        guards; when omitted they are read from the session's current tree.
        """
        from ..authoring import (
            _APPLIED,
            _MAX_TEXT_CHARACTERS,
            CellRequestLimitError,
            CellTextAuthoringReport,
            CellTextRequest,
            _cell_structure_fingerprint,
            _cell_text_landed,
            _plan_cell_text,
            _require_positive_coordinate,
            _selected_cell,
            _simple_text_node,
        )
        from ..authoring import _refusal as _cell_refusal
        from ..authoring import _word_namespace as _cell_word_namespace
        from ..content import _evaluate_target_changes

        self._source()
        coordinates = {
            name: _require_positive_coordinate(value, f"cell_target_{name}")
            for name, value in (("table", table), ("row", row), ("cell", cell))
        }
        if not isinstance(text, str):
            raise TypeError("cell_text_invalid")
        if len(text) > _MAX_TEXT_CHARACTERS:
            raise CellRequestLimitError("text", len(text))
        request = CellTextRequest(
            source=self.path,
            destination=self.path,
            source_sha256=self.source_sha256,
            expected_old_text=expected_old_text or "",
            structure_fingerprint_sha256=structure_fingerprint_sha256 or _UNBOUND_FINGERPRINT,
            text=text,
            **coordinates,
        )
        try:
            document_xml = self.part_bytes(_DOCUMENT)
            # Planning only reads the tree, so a cached editable tree or a read-only
            # one parsed from these bytes serves.
            root = self._trees.get(_DOCUMENT)
            if root is None:
                root = self._checked_tree(_DOCUMENT, document_xml)
        except _PART_ERRORS:
            return _cell_refusal(request, "source_document_invalid")
        namespace = _cell_word_namespace(root)
        selected = _selected_cell(root, request, namespace) if namespace else None
        if selected is not None:
            simple = _simple_text_node(selected, namespace) if namespace else None
            request = replace(
                request,
                expected_old_text=(
                    expected_old_text
                    if expected_old_text is not None or simple is None
                    else simple[1]
                ),
                structure_fingerprint_sha256=(
                    structure_fingerprint_sha256 or _cell_structure_fingerprint(selected)
                ),
            )
        plan = _plan_cell_text(request, root, document_xml)
        if isinstance(plan, str):
            return _cell_refusal(request, plan)
        try:
            checked = self._parse_checked(plan.document_xml)
        except _PART_ERRORS:
            return _cell_refusal(request, "staged_verification_failed")
        candidate = checked[1]
        if not _cell_text_landed(candidate, request, plan.word_namespace):
            return _cell_refusal(request, "staged_verification_failed")
        # The change check masks both trees in place and restores them before it returns.
        changes = _evaluate_target_changes(
            [plan.target_change_scope],
            {_DOCUMENT: hashlib.sha256(document_xml).hexdigest()},
            {_DOCUMENT: hashlib.sha256(plan.document_xml).hexdigest()},
            load_baseline_root=lambda _name: root,
            load_candidate_root=lambda _name: candidate,
        )
        if not changes.passed:
            return _cell_refusal(request, "staged_verification_failed")
        self._set_part(_DOCUMENT, plan.document_xml)
        self._checked[_DOCUMENT] = checked
        self._record("cell-text", (_DOCUMENT,))
        return CellTextAuthoringReport(
            status=_APPLIED,
            code=None,
            source_sha256=self.source_sha256,
            target=request.target_scope.to_dict(),
            structure_fingerprint_sha256=request.structure_fingerprint_sha256,
            target_change_report=changes.to_dict(),
            non_target_parts_preserved=True,
        )

    def set_run_text(
        self,
        paragraph: int,
        run: int,
        text: str,
        *,
        expected_old_text: str | None = None,
    ) -> None:
        """Set one body run's text atomically; publish only with :meth:`save`.

        ``paragraph`` and ``run`` are zero-based indices of direct ``w:p``
        children of ``w:body`` and direct ``w:r`` children of that paragraph.
        Table, hyperlink, revision and content-control descendants do not count.
        Both Transitional and Strict documents are supported.

        Replace only ``w:t``, ``w:tab``, ``w:br``, ``w:cr`` and
        ``w:noBreakHyphen``. Insert the new content at the first replaced child,
        or append it when none exists. Keep run properties, attributes, and
        non-text children in their original relative order; never merge runs.
        Emit tabs, LF, CR and U+2011 as tab, break, carriage-return and
        nonbreaking-hyphen elements. For the guard, CR and wrapping breaks read
        as LF, nonbreaking hyphens as "-", and page/column breaks as empty text.
        Preserve leading/trailing spaces with ``xml:space="preserve"``.

        Raise :class:`~wolfdocx.RunTextEditError` for invalid indices or text,
        out-of-range addresses, or an ``expected_old_text`` mismatch, leaving
        the session and filesystem unchanged. Text and guard values are limited
        to 32,768 XML 1.0 characters. An empty guard checks for an empty run.
        Saving preserves every part other than ``word/document.xml`` byte-for-byte.

        Example::

            with Document("source.docx") as document:
                document.set_run_text(2, 0, "Updated", expected_old_text="Old")
                report = document.save("updated.docx")
        """
        from ..inspect import serialize_xml_with_preserved_qnames
        from ..run_text import _replace_run_text, _validate_request

        self._source()
        _validate_request(paragraph, run, text, expected_old_text)
        # Parse a private editable tree: refusals never mutate cached session trees.
        root = self._parse(self.part_bytes(_DOCUMENT), editable=True)
        _replace_run_text(root, paragraph, run, text, expected_old_text)
        data = serialize_xml_with_preserved_qnames(root)
        checked = self._parse_checked(data)
        self._set_part(_DOCUMENT, data)
        self._checked[_DOCUMENT] = checked
        self._record("run-text", (_DOCUMENT,))

    def find_replace(
        self,
        pattern: str,
        replacement: str | Callable[[Any], str] = "",
        *,
        part: str = _DOCUMENT,
        options: FindReplaceOptions | Mapping[str, Any] | None = None,
    ) -> FindReplaceReport:
        """Regex find and replace over one part, as :func:`~wolfdocx.find_replace` would.

        A pattern that replaces nothing refuses with ``find_replace_no_match`` and
        leaves the session unchanged.
        """
        from ..find_replace import (
            FindReplaceReport,
            FindReplaceRequest,
            _compile_pattern,
            _replace_in_story,
        )
        from ..find_replace import _refusal as _find_replace_refusal

        self._source()
        request = FindReplaceRequest(
            self.path,
            self.path,
            pattern=pattern,
            replacement=replacement,
            source_sha256=self.source_sha256,
            part=part,
            options=options,
        )
        compiled = _compile_pattern(request.pattern, request.options)
        if compiled is None:
            return _find_replace_refusal(self.source_sha256, "find_replace_pattern_invalid")
        if not self._graph.has_part(part):
            return _find_replace_refusal(self.source_sha256, "find_replace_part_unsupported")
        try:
            _, root = self._take_editable(part)
        except _PART_ERRORS:
            return _find_replace_refusal(self.source_sha256, "source_package_invalid")
        matches, count, data = _replace_in_story(root, compiled, request)
        if not count:
            return _find_replace_refusal(self.source_sha256, "find_replace_no_match")
        self._set_part(part, data)
        self._record("find-replace", (part,))
        return FindReplaceReport(
            status="applied",
            code=None,
            source_sha256=self.source_sha256,
            destination=None,
            match_count=len(matches),
            replacement_count=count,
            matches=matches,
            unselected_parts_preserved=True,
        )

    # -- publication ------------------------------------------------------------

    def save(self, path: str | Path | BinaryIO) -> SessionSaveReport:
        """Verify edits and save without replacing an existing destination."""
        source = self._source()
        sink = path if hasattr(path, "write") else None
        destination = None if sink is not None else Path(path).resolve()
        display = "<stream>" if destination is None else str(destination)
        changed = dict(self._changed)
    
        def refusal(code: str) -> SessionSaveReport:
            return self._record_save_report(
                SessionSaveReport(_REFUSED, code, self.source_sha256, display)
            )
    
        if sink is not None:
            try:
                if not sink.seekable() or sink.tell() != 0 or sink.seek(0, 2) != 0:
                    return refusal("document_destination_invalid")
                sink.seek(0)
            except (AttributeError, OSError, ValueError):
                return refusal("document_destination_invalid")
        else:
            assert destination is not None
            if destination.exists() or destination.is_symlink():
                return refusal("document_destination_exists")
            if not destination.parent.is_dir():
                return refusal("document_destination_invalid")
            if (destination.suffix.lower() == ".docx"
                and self._graph.parts[_main_part(self._graph)].content_type
                in _MACRO_MAIN_CONTENT_TYPES):
                return refusal("document_destination_extension_mismatch")
        removed = tuple(sorted(self._infos.keys() - self._entries.keys()))
        structural = {name for name in changed if _structural(name)}
        staged: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                dir=None if destination is None else destination.parent,
                prefix="wolfdocx-save-" if destination is None else f".{destination.name}.",
                suffix=".tmp", delete=False,
            ) as stream:
                staged = Path(stream.name)
                self._write_to(stream, source, changed)
            verified = _verify_staged_changes(
                staged,
                {name: data for name, data in changed.items() if name not in structural},
                limits=self._graph.limits,
                unchanged_identical=lambda graph: self._unchanged_raw(graph, staged, changed),
                unchanged_elements=0, baseline_tree=self._current_tree, parsed=self._checked,
                check_notes=bool(changed.keys() & self._note_parts()) or bool(removed),
            )
            if verified is None:
                return refusal("document_staged_verification_failed")
            output_sha256 = verified[0]
            if sink is not None:
                try:
                    with staged.open("rb") as stream:
                        while chunk := stream.read(64 * 1024):
                            if sink.write(chunk) != len(chunk):
                                raise OSError("short stream write")
                except (OSError, ValueError):
                    sink.seek(0)
                    sink.truncate(0)
                    return refusal("document_publication_unavailable")
                linked = None
            else:
                from ..review import _link_no_clobber
                assert destination is not None
                linked = _link_no_clobber(
                    staged, destination, exists_code="document_destination_exists",
                    unavailable_code="document_publication_unavailable",
                )
        except OSError:
            return refusal("document_publication_unavailable")
        except (ValueError, zipfile.BadZipFile, struct.error) + _PART_ERRORS:
            return refusal("document_staged_verification_failed")
        finally:
            if staged is not None:
                staged.unlink(missing_ok=True)
        if linked is not None:
            return refusal(linked)
        return SessionSaveReport(
            _PUBLISHED, None, self.source_sha256, display,
            destination_sha256=output_sha256,
            changed_parts=tuple(sorted(name for name in changed if name in self._infos)),
            added_parts=tuple(sorted(name for name in changed if name not in self._infos)),
            removed_parts=removed, raw_copied_parts=len(self.part_names) - len(changed),
            edits=tuple(self._edits),
        )

    

    def _record_save_report(self, report: SessionSaveReport) -> SessionSaveReport:
        if report.code is not None:
            from ..render_warnings import RenderWarning, WarningCategory, WarningSource

            self.warnings._extend(
                (
                    RenderWarning(
                        WarningCategory.PUBLICATION,
                        WarningSource.SAVE,
                        report.code,
                        kind=report.code,
                        locator=report.destination,
                        reason=report.code,
                        severity="error" if not report.passed else "warning",
                    ),
                )
            )
        return report

    

    # -- internals --------------------------------------------------------------

    def _parse(self, data: bytes, *, editable: bool) -> ET.Element:
        return _parse_part(data, self._graph.limits, editable=editable)

    def _parse_checked(self, data: bytes) -> tuple[bytes, ET.Element, int]:
        """A read-only tree of ``data`` with the bytes and element count it came from."""
        root = self._parse(data, editable=False)
        return data, root, sum(1 for _ in root.iter())

    def _checked_tree(self, name: str, data: bytes) -> ET.Element:
        """A read-only tree of part ``name`` holding ``data``, parsed at most once."""
        checked = self._checked.get(name)
        if checked is None or checked[0] != data:
            checked = self._checked[name] = self._parse_checked(data)
        return checked[1]

    def _editable(self, name: str) -> tuple[bytes, ET.Element]:
        """A part's current bytes and its cached editable tree, parsed on first use."""
        data = self.part_bytes(name)
        root = self._trees.get(name)
        if root is None:
            root = self._trees[name] = self._parse(data, editable=True)
        return data, root

    def _take_editable(self, name: str) -> tuple[bytes, ET.Element]:
        """Like :meth:`_editable`, for a caller that mutates the tree it receives."""
        loaded = self._editable(name)
        del self._trees[name]
        return loaded

    

    

    

    

    

    

    def _set_part(self, name: str, data: bytes) -> None:
        self._changed[name] = data
        self._trees.pop(name, None)

    def _record(
        self, operation: str, changed: tuple[str, ...], removed: tuple[str, ...] = ()
    ) -> None:
        self._edits.append(
            {
                "operation": operation,
                "changed_parts": list(changed),
                "removed_parts": list(removed),
            }
        )

    def _current_tree(self, name: str) -> ET.Element | None:
        """A read-only tree of one unchanged current part, for note-reference checks."""
        if name not in self._entries:
            return None
        return self._checked_tree(name, self.part_bytes(name))

    # -- operation edits ----------------------------------------------------------

    def _edit_paragraph(self, paragraph: int, part: str, **fields: object) -> SessionEditReport:
        from ..authoring import _cell_structure_fingerprint
        from ..paragraph_authoring import (
            ParagraphRequest,
            _paragraph_refusal,
            _resolve_target_story,
        )

        self._source()
        request = ParagraphRequest(
            source=self.path,
            destination=self.path,
            source_sha256=self.source_sha256,
            structure_fingerprint_sha256=_UNBOUND_FINGERPRINT,
            paragraph=paragraph,
            part=part,
            **fields,  # type: ignore[arg-type]
        )
        root, _, ns, before, error = _resolve_target_story(self._graph, part, self._editable)
        if error is not None:
            report = _paragraph_refusal(request, error)
        else:
            assert root is not None and ns is not None
            target = next(
                (node for index, node in enumerate(root.iter(Q(ns, "p")), 1) if index == paragraph),
                None,
            )
            if target is not None:
                request = replace(
                    request,
                    structure_fingerprint_sha256=_cell_structure_fingerprint(target),
                )
            report = self._apply_paragraph_request(request)
        return SessionEditReport(
            report.status,
            report.code,
            "paragraph",
            changed_parts=(part,) if report.passed and self.part_bytes(part) != before else (),
            detail=operation_detail(report),
        )

    def _materialize(self) -> bytes:
        """The session's current package, staged exactly as :meth:`save` stages it."""
        return self._write(self._source(), self._changed)

    def _adopt(self, operation: str, edited: bytes, detail: Mapping[str, Any]) -> SessionEditReport:
        """Take an operation's output package as the session's current package.

        A source part whose bytes equal the source's is unchanged again and is copied
        raw on save. An added part gets a fixed entry (the ZIP epoch, deflated) rather
        than the clock time the operation wrote, so the same edits save the same bytes.
        """
        graph = PackageGraph.from_bytes(edited, path=self.path, limits=self._graph.limits)
        entries: dict[str, zipfile.ZipInfo] = {}
        changed: dict[str, bytes] = {}
        touched: list[str] = []
        with zipfile.ZipFile(BytesIO(edited)) as archive:
            for info in archive.infolist():
                name = info.filename
                if name not in graph.parts and name in self._entries:
                    # Tolerated ZIP metadata (directories, legacy trash) is not an
                    # OPC part: keep its source entry, copied raw on save.
                    entries[name] = self._entries[name]
                    continue
                data = archive.read(info)
                is_touched = name not in self._entries or data != self.part_bytes(name)
                if is_touched:
                    touched.append(name)
                source_info = self._infos.get(name)
                if source_info is None:
                    entries[name] = self._entries.get(name) or _added_entry(info)
                    changed[name] = data
                    continue
                entries[name] = source_info
                if is_touched:
                    if data != self._source_graph.read_part_bytes(name):
                        changed[name] = data
                elif name in self._changed:
                    changed[name] = self._changed[name]
        removed = tuple(sorted(self._entries.keys() - entries.keys()))
        for name in (*touched, *removed):
            self._trees.pop(name, None)
        self._entries, self._changed, self._graph = entries, changed, graph
        self._record(operation, tuple(touched), removed)
        return SessionEditReport(
            _EDIT_APPLIED,
            None,
            operation,
            changed_parts=tuple(touched),
            removed_parts=removed,
            detail=detail,
        )

    def _edit_part(
        self,
        operation: str,
        part: str,
        plan: Callable[[Any, Callable[[str], Any]], tuple[bytes, Any] | str],
        *,
        failure_code: str,
        detail: Mapping[str, Any],
    ) -> SessionEditReport:
        from ..content import _xml_trees_equal

        self._source()
        # Planning edits the cached tree in place, so the cache gives it up first.
        planned = plan(self._graph, self._take_editable)
        if isinstance(planned, str):
            return _edit_refused(operation, planned)
        data, expected = planned
        try:
            landed = _xml_trees_equal(self._parse(data, editable=True), expected)
        except _PART_ERRORS:
            landed = False
        if not landed:
            return _edit_refused(operation, failure_code)
        self._set_part(part, data)
        self._record(operation, (part,))
        return SessionEditReport(
            _EDIT_APPLIED, None, operation, changed_parts=(part,), detail=detail
        )

    def _snapshot(self) -> object:
        snapshot: _Snapshot = (
            dict(self._entries),
            dict(self._changed),
            self._graph,
            len(self._edits),
        )
        return snapshot

    def _restore(self, snapshot: object) -> None:
        entries, changed, graph, edits = cast(_Snapshot, snapshot)
        self._entries, self._changed, self._graph = entries, changed, graph
        del self._edits[edits:]
        self._trees.clear()
        self._checked.clear()

    def _note_parts(self) -> set[str]:
        """The parts whose change can affect footnote or endnote references."""
        parts = {_DOCUMENT}
        story = self._graph.parts.get(_DOCUMENT)
        if story is not None:
            parts.update(
                relationship.resolved_target()
                for relationship in story.relationships
                if relationship.type.endswith(_NOTE_RELATIONSHIPS) and not relationship.is_external
            )
        return parts

    def _paragraph_request(
        self, paragraph: int, fingerprint: str, text: str, style: str | None, part: str
    ) -> ParagraphRequest:
        from ..paragraph_authoring import ParagraphRequest

        return ParagraphRequest(
            source=self.path,
            destination=self.path,
            source_sha256=self.source_sha256,
            structure_fingerprint_sha256=fingerprint,
            paragraph=paragraph,
            part=part,
            action="insert",
            position="after",
            text=text,
            style=style,
        )

    def _write(self, source: bytes | mmap.mmap, changed: Mapping[str, bytes]) -> bytes:
        """Build the staged package: raw copies of unchanged members, changed parts written."""
        buffer = BytesIO()
        self._write_to(buffer, source, changed)
        return buffer.getvalue()

    def _write_to(
        self, buffer: BinaryIO, source: bytes | mmap.mmap, changed: Mapping[str, bytes]
    ) -> None:
        if not changed and self._entries.keys() == self._infos.keys():
            buffer.write(source)
            return
        with zipfile.ZipFile(buffer, "w") as out:
            out.comment = self._comment
            for name, info in self._entries.items():
                if name in changed:
                    # A source entry keeps its compression method and metadata.
                    out.writestr(copy.copy(info), changed[name])
                else:
                    _copy_member_raw(out, info, _raw_member(source, info))

    def _unchanged_raw(
        self, graph: PackageGraph, staged: bytes | Path, changed: Mapping[str, bytes]
    ) -> bool:
        """True when the staged package names the session's parts, every changed
        relationship or content-type part holds its intended bytes, and every unchanged
        member holds the source member's exact compressed bytes and entry facts."""
        if set(graph.parts) != set(self.part_names):
            return False
        if any(
            graph.read_part_bytes(name) != data
            for name, data in changed.items()
            if _structural(name)
        ):
            return False
        if isinstance(staged, Path):
            with (
                staged.open("rb") as stream,
                mmap.mmap(stream.fileno(), 0, access=mmap.ACCESS_READ) as mapped,
            ):
                return self._unchanged_raw_bytes(graph, mapped, changed)
        return self._unchanged_raw_bytes(graph, staged, changed)

    def _unchanged_raw_bytes(
        self, graph: PackageGraph, staged: bytes | mmap.mmap, changed: Mapping[str, bytes]
    ) -> bool:
        source = self._source()
        with graph._open_archive() as archive:
            staged_infos = {info.filename: info for info in archive.infolist()}
        # Every session entry must survive; anything new must be an OPC part the
        # operation added, never a stray metadata entry.
        if not self._entries.keys() <= staged_infos.keys() or any(
            name not in graph.parts for name in staged_infos.keys() - self._entries.keys()
        ):
            return False
        for name, info in self._entries.items():
            if name in changed:
                continue
            copied = staged_infos.get(name)
            if copied is None or (
                copied.CRC,
                copied.file_size,
                copied.compress_size,
                copied.compress_type,
            ) != (info.CRC, info.file_size, info.compress_size, info.compress_type):
                return False
            if _raw_member(staged, copied) != _raw_member(source, info):
                return False
        return True
