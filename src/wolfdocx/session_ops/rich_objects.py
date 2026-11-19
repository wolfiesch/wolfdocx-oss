"""One DOCX package opened once, edited in memory, and published without clobbering.

:class:`Document` validates the package graph, member CRCs and bounded main-story
XML before opening. Explicit load options require whole-package admission or
verified recovery unless the caller requests deferred part validation for
diagnosis or opaque preservation. Other parts are parsed when an edit needs them.
Every edit reuses its stateless counterpart; a refusal leaves the session
unchanged. :meth:`Document.append_paragraph`, :meth:`Document.set_cell_text`,
:meth:`Document.find_replace`, and the paragraph and run edits plan in memory with
the stateless planning step. Table, picture, section, page-setup, header and footer,
list numbering, and styling edits (:mod:`wolfdocx.session_ops`), and the comment,
footnote and endnote, bookmark, hyperlink, field and table-of-contents, content
control, watermark, and text box edits (:class:`RichObjectEdits`) run the stateless
operation itself against the session's current package in a private scratch directory
and adopt its output, including parts it adds or removes.

:meth:`Document.save` copies every unchanged ZIP member's compressed bytes and
entry metadata from the source and writes only the changed and added parts.
Without changes, the staged digest must equal the admitted immutable source
digest, preserving every byte without another package parse. With changes,
the staged package must validate as a package graph, every
unchanged member must hold the source's compressed bytes, every changed
relationship or content-type part must hold its intended bytes, each changed part
must re-parse under the package limits, and note references are re-checked when
the story or a notes part changed or a part was removed. Unchanged parts are never
parsed, so the aggregate element limit covers the changed parts only. The staged
file is published with a hard link, which never replaces an existing destination.
"""

from __future__ import annotations

import copy
import hashlib
import mmap
import re
import struct
import tempfile
import zipfile
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO, cast
from xml.etree import ElementTree as ET
from xml.parsers import expat

from .._optional import feature_module
from ..canonical import sha256_bytes, sha256_file
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
    _FieldTextState,
)
from ..output import (
    OoxmlCompliance,
    OoxmlSaveOptions,
    SaveOptionsError,
    SaveOutputParameters,
    TxtExportHeadersFootersMode,
    TxtOfficeMathExportMode,
    TxtSaveOptions,
    Zip64Mode,
    _apply_save_properties,
)
from ..package import PackageError, PackageGraph, PackageLimits
from . import (
    ListEdits,
    PictureEdits,
    RunEdits,
    SectionEdits,
    SessionEditReport,
    StylingEdits,
    TableEdits,
)
from ._bridge import _APPLIED as _EDIT_APPLIED
from ._bridge import EditHost, Staged, bound_fingerprint, operation_detail
from ._bridge import refused as _edit_refused

if TYPE_CHECKING:
    from ..authoring import CellTextAuthoringReport
    from ..find_replace import FindReplaceOptions, FindReplaceReport
    from ..load_options import LoadOptions, LoadReport
    from ..paragraph_authoring import ParagraphReport, ParagraphRequest
    from ..render_warnings import WarningCollection
    from ..web_extensions import WebAddIns


def _plan_paragraph_edit(*args: Any, **kwargs: Any) -> Any:
    """:func:`wolfdocx.paragraph_authoring._plan_paragraph_edit`, imported on first use."""
    from ..paragraph_authoring import _plan_paragraph_edit

    return _plan_paragraph_edit(*args, **kwargs)


from ..document import Document, _load_refusal

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
    load_recovery: tuple[str, ...] = ()
    edits: tuple[Mapping[str, Any], ...] = ()
    invalidated_signatures: tuple[str, ...] = ()

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
            "load_recovery": list(self.load_recovery),
            "edits": [dict(edit) for edit in self.edits],
            "invalidated_signatures": list(self.invalidated_signatures),
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
# Same value as wolfdocx.package_signatures._SIGNATURE_CONTENT_TYPE: a package
# without signature parts has no signature a save can invalidate.
_SIGNATURE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
)
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
# ECMA-376 Part 1 namespace table. OPC namespaces remain unchanged.
_OOXML_NAMESPACE_PATHS = (
    "wordprocessingml/main",
    "drawingml/main",
    "drawingml/chart",
    "drawingml/chartDrawing",
    "drawingml/diagram",
    "drawingml/picture",
    "drawingml/spreadsheetDrawing",
    "drawingml/wordprocessingDrawing",
    "spreadsheetml/main",
    "presentationml/main",
    "officeDocument/math",
    "officeDocument/relationships",
    "officeDocument/extended-properties",
    "officeDocument/custom-properties",
    "officeDocument/docPropsVTypes",
    "officeDocument/bibliography",
    "officeDocument/customXml",
    "officeDocument/characteristics",
    "officeDocument/sharedTypes",
)
_OOXML_NAMESPACES = {
    _TRANSITIONAL_BASE + path.replace("/", "/2006/", 1): _STRICT_BASE + path
    for path in _OOXML_NAMESPACE_PATHS
}
# Property namespaces and relationship names are hyphenated in Transitional.
# Strict uses camelCase for both; neither is a mechanical base-URI substitution.
_OOXML_NAMESPACES.update(
    {
        _TRANSITIONAL_BASE + "officeDocument/2006/extended-properties": _STRICT_BASE
        + "officeDocument/extendedProperties",
        _TRANSITIONAL_BASE + "officeDocument/2006/custom-properties": _STRICT_BASE
        + "officeDocument/customProperties",
        _TRANSITIONAL_BASE + "officeDocument/2006/relationships/extended-properties": _STRICT_BASE
        + "officeDocument/relationships/extendedProperties",
        _TRANSITIONAL_BASE + "officeDocument/2006/relationships/custom-properties": _STRICT_BASE
        + "officeDocument/relationships/customProperties",
    }
)
_XML_ATTRIBUTE = re.compile(rb"""([\w:.-]+)(\s*=\s*)(["'])(.*?)\3""", re.DOTALL)
_XML_TAG = re.compile(
    rb"""<!--.*?-->|<!\[CDATA\[.*?\]\]>|<([A-Za-z_][\w:.-]*)(?:[^"'<>]|"[^"]*"|'[^']*')*>""",
    re.DOTALL,
)


def _xml_utf8(data: bytes) -> bytes:
    """Normalize non-UTF-8 XML encodings only on explicit OOXML conversion."""
    if data.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
        text = data.decode("utf-32")
    elif data.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = data.decode("utf-16")
    elif data.startswith(b"\x00<"):
        text = data.decode("utf-16-be")
    elif data.startswith(b"<\x00"):
        text = data.decode("utf-16-le")
    else:
        declaration = re.match(
            rb"""(?:\xef\xbb\xbf)?<\?xml[^>]*encoding\s*=\s*["']([^"']+)["']""", data
        )
        encoding = declaration.group(1).decode("ascii") if declaration else "utf-8"
        if encoding.lower().replace("_", "-") in {"utf-8", "utf8"}:
            return data
        text = data.decode(encoding)
    text = re.sub(r"""(<\?xml[^>]*encoding\s*=\s*["'])[^"']+""", r"\1UTF-8", text, count=1)
    return text.encode("utf-8")


def _convert_ooxml(data: bytes, *, strict: bool, main: bool, limits: PackageLimits) -> bytes:
    root = _parse_part(data, limits, editable=False)
    if main and root.tag not in {
        Q(_TRANSITIONAL_BASE + "wordprocessingml/2006/main", "document"),
        Q(_STRICT_BASE + "wordprocessingml/main", "document"),
    }:
        raise SaveOptionsError("ooxml_document_namespace_unsupported")
    if strict and any(
        isinstance(node.tag, str) and node.tag.startswith("{urn:schemas-microsoft-com:vml}")
        for node in root.iter()
    ):
        raise SaveOptionsError("ooxml_strict_vml_unsupported")
    mapping = _OOXML_NAMESPACES if strict else {v: k for k, v in _OOXML_NAMESPACES.items()}
    ordered_mapping = sorted(mapping.items(), key=lambda item: len(item[0]), reverse=True)
    data = _xml_utf8(data)
    # Rename only direct Word table-border children. Byte offsets retain prefixes,
    # whitespace, attributes, comments and unrelated left/right elements.
    word_namespaces = {
        _TRANSITIONAL_BASE + "wordprocessingml/2006/main",
        _STRICT_BASE + "wordprocessingml/main",
    }
    sides = {"left": "start", "right": "end"} if strict else {"start": "left", "end": "right"}
    parser = expat.ParserCreate(namespace_separator="}")
    stack: list[tuple[str, str | None, bool]] = []
    patches: list[tuple[int, int, bytes]] = []

    def border_start(name: str, attributes: dict[str, str]) -> None:
        namespace, _, local = name.rpartition("}")
        replacement = (
            sides.get(local)
            if namespace in word_namespaces and stack and stack[-1][0] == namespace + "}tblBorders"
            else None
        )
        offset = parser.CurrentByteIndex
        opening_tag = _XML_TAG.match(data, offset)
        assert opening_tag is not None
        empty = opening_tag.group().rstrip().endswith(b"/>")
        if replacement is not None:
            opening = re.match(rb"<([\w:.-]+)", data[offset:])
            assert opening is not None
            begin, end = opening.span(1)
            prefix = opening.group(1).rpartition(b":")[0]
            qname = (prefix + b":" if prefix else b"") + replacement.encode()
            patches.append((offset + begin, offset + end, qname))
        stack.append((name, replacement, empty))

    def border_end(name: str) -> None:
        _, replacement, empty = stack.pop()
        if replacement is not None and not empty:
            offset = parser.CurrentByteIndex
            closing = re.match(rb"</([\w:.-]+)", data[offset:])
            assert closing is not None
            begin, end = closing.span(1)
            prefix = closing.group(1).rpartition(b":")[0]
            qname = (prefix + b":" if prefix else b"") + replacement.encode()
            patches.append((offset + begin, offset + end, qname))

    parser.StartElementHandler = border_start
    parser.EndElementHandler = border_end
    parser.Parse(data, True)
    for begin, end, replacement in reversed(patches):
        data = data[:begin] + replacement + data[end:]
    relationships = (
        root.tag == "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
    )

    def attribute(match: re.Match[bytes]) -> bytes:
        name, equal, quote, raw = match.groups()
        value = raw.decode("utf-8")
        if name.startswith(b"xmlns") or (relationships and name == b"Type"):
            # Match specific relationship exceptions before their common base.
            for old, new in ordered_mapping:
                if value == old or (name == b"Type" and value.startswith(old + "/")):
                    value = new + value[len(old) :]
                    break
        return name + equal + quote + value.encode("utf-8") + quote

    result = _XML_TAG.sub(
        lambda match: (
            _XML_ATTRIBUTE.sub(attribute, match.group()) if match.group(1) else match.group()
        ),
        data,
    )
    if main:
        word = (
            _STRICT_BASE + "wordprocessingml/main"
            if strict
            else (_TRANSITIONAL_BASE + "wordprocessingml/2006/main")
        )
        # Only the root start tag is changed, preserving lexical QNames and text.
        opening = next(
            (
                match
                for match in _XML_TAG.finditer(result)
                if match.group(1) and match.group(1).split(b":")[-1] == b"document"
            ),
            None,
        )
        if opening is None:
            raise SaveOptionsError("ooxml_document_namespace_unsupported")
        tag = opening.group()
        declaration = re.search(
            rb"""xmlns:([\w.-]+)\s*=\s*["']""" + re.escape(word.encode()) + rb"""["']""",
            tag,
        )
        prefix = declaration.group(1) if declaration else b"wdSave"
        tag = re.sub(rb"""\s+[\w.-]+:conformance\s*=\s*(["']).*?\1""", b"", tag)
        addition = b""
        if declaration is None:
            addition += b" xmlns:" + prefix + b'="' + word.encode() + b'"'
        addition += (
            b" " + prefix + b':conformance="' + (b"strict" if strict else b"transitional") + b'"'
        )
        result = result[: opening.start()] + tag[:-1] + addition + b">" + result[opening.end() :]
    return result


def _main_part(graph: PackageGraph) -> str:
    for relationship in graph.relationships:
        if relationship.type.endswith("/officeDocument") and not relationship.is_external:
            return relationship.resolved_target()
    raise SaveOptionsError("document_main_part_missing")


def _save_ooxml(data: bytes, options: OoxmlSaveOptions, limits: PackageLimits) -> bytes:
    strict = options.compliance == OoxmlCompliance.STRICT
    graph = PackageGraph.from_bytes(data, path=Path("save.docx"), limits=limits)
    main = _main_part(graph)
    # Explicit defaults retain the raw-copy default for Transitional sources.
    if (
        options.compliance == OoxmlCompliance.TRANSITIONAL
        and options.compression_level is None
        and options.zip64_mode == Zip64Mode.AS_NEEDED
        and not options.export_legacy_control_characters
    ) and _parse_part(graph.read_part_bytes(main), limits, editable=False).tag == Q(
        _TRANSITIONAL_BASE + "wordprocessingml/2006/main", "document"
    ):
        return data
    from ..web_extensions import _preserved_parts

    opaque = _preserved_parts(graph)
    opaque.discard(main)
    opaque.add("[Content_Types].xml")
    buffer = BytesIO()
    expected: dict[str, bytes] = {}
    try:
        with (
            zipfile.ZipFile(BytesIO(data)) as source,
            zipfile.ZipFile(
                buffer, "w", allowZip64=options.zip64_mode != Zip64Mode.NEVER
            ) as output,
        ):
            output.comment = source.comment
            for original in source.infolist():
                original_payload = payload = source.read(original)
                content_type = graph.parts[original.filename].content_type
                if original.filename not in opaque and (
                    original.filename.endswith((".xml", ".rels", ".vml"))
                    or content_type.endswith("+xml")
                    or content_type in {"application/xml", "text/xml"}
                ):
                    if options.compliance == OoxmlCompliance.ECMA_376_2006:
                        payload = _first_edition_ooxml(payload, limits)
                    payload = _convert_ooxml(
                        payload,
                        strict=strict,
                        main=original.filename == main,
                        limits=limits,
                    )
                    if options.compliance == OoxmlCompliance.ECMA_376_2006:
                        payload = _first_edition_ooxml(payload, limits)
                    if options.export_legacy_control_characters:
                        payload = _legacy_control_ooxml(payload, limits)
                expected[original.filename] = payload
                if (
                    options.compression_level is None
                    and options.zip64_mode == Zip64Mode.AS_NEEDED
                    and payload == original_payload
                ):
                    _copy_member_raw(output, original, _raw_member(data, original))
                    continue
                entry = copy.copy(original)
                entry.extra = _without_zip64_extra(entry.extra)
                entry.file_size = len(payload)
                if (
                    options.zip64_mode == Zip64Mode.NEVER
                    and len(payload) * 1.05 > zipfile.ZIP64_LIMIT
                ):
                    raise SaveOptionsError("ooxml_zip64_required")
                if options.compression_level is not None:
                    entry.compress_type = zipfile.ZIP_DEFLATED
                    entry._compresslevel = options.compression_level
                with output.open(
                    entry, "w", force_zip64=options.zip64_mode == Zip64Mode.ALWAYS
                ) as stream:
                    stream.write(payload)
    except zipfile.LargeZipFile as exc:
        raise SaveOptionsError("ooxml_zip64_required") from exc
    result = buffer.getvalue()
    admitted = PackageGraph.from_bytes(result, path=Path("save.docx"), limits=limits)
    if any(admitted.read_part_bytes(name) != payload for name, payload in expected.items()):
        raise SaveOptionsError("ooxml_staged_verification_failed")
    return result


def _first_edition_ooxml(data: bytes, limits: PackageLimits) -> bytes:
    """Downgrade admitted 2008 direction properties to the first-edition names.

    Later extension semantics without a native downgrade refuse. Escaped
    control-looking literal text is not interpreted as an OOXML Xstring.
    """
    root = _parse_part(data, limits, editable=False)
    word = _TRANSITIONAL_BASE + "wordprocessingml/2006/main"
    changed = False
    replacements = {
        "start": "left",
        "end": "right",
        "startChars": "leftChars",
        "endChars": "rightChars",
    }
    directional = any(
        node.tag in {Q(word, "start"), Q(word, "end")}
        or any(Q(word, name) in node.attrib for name in replacements)
        for node in root.iter()
    )
    if directional and any(
        node.tag == Q(word, "bidi")
        and node.get(Q(word, "val"), "true").lower() not in {"0", "false", "off"}
        for node in root.iter()
    ):
        raise SaveOptionsError("ooxml_2006_bidi_direction_unsupported")
    for node in root.iter():
        namespace, _, name = str(node.tag).removeprefix("{").partition("}")
        if namespace.startswith("http://schemas.microsoft.com/office/word/"):
            raise SaveOptionsError("ooxml_2006_extension_unsupported")
        if namespace == "http://schemas.openxmlformats.org/markup-compatibility/2006":
            raise SaveOptionsError("ooxml_2006_markup_compatibility_unsupported")
        if any(key.startswith("{http://schemas.microsoft.com/office/word/") for key in node.attrib):
            raise SaveOptionsError("ooxml_2006_extension_unsupported")
        if any(
            key.startswith("{http://schemas.openxmlformats.org/markup-compatibility/2006}")
            for key in node.attrib
        ):
            raise SaveOptionsError("ooxml_2006_markup_compatibility_unsupported")
        if namespace != word:
            continue
        if node.tag == Q(word, "document") and Q(word, "conformance") in node.attrib:
            del node.attrib[Q(word, "conformance")]
            changed = True
        if name in {"ind", "tblCellMar", "tcMar"}:
            for old, new in replacements.items():
                key = Q(word, old)
                if key in node.attrib:
                    if Q(word, new) in node.attrib:
                        raise SaveOptionsError("ooxml_2006_direction_property_conflict")
                    node.set(Q(word, new), node.attrib.pop(key))
                    changed = True
        if name in {"jc", "tblJc"} and node.get(Q(word, "val")) in {"start", "end"}:
            raise SaveOptionsError("ooxml_2006_contextual_alignment_unsupported")
        if name in {"tblBorders", "tcBorders", "tblCellMar", "tcMar"}:
            for child in node:
                local = child.tag.removeprefix("{" + word + "}")
                if local in {"start", "end"}:
                    target = Q(word, replacements[local])
                    if node.find(target) is not None:
                        raise SaveOptionsError("ooxml_2006_direction_property_conflict")
                    child.tag = target
                    changed = True
    if not changed:
        return data
    from ..inspect import serialize_xml_with_preserved_qnames

    return serialize_xml_with_preserved_qnames(root)


def _legacy_control_ooxml(data: bytes, limits: PackageLimits) -> bytes:
    """Represent XML-safe embedded text controls as WordprocessingML run nodes."""
    root = _parse_part(data, limits, editable=False)
    changed = False
    words = {
        _TRANSITIONAL_BASE + "wordprocessingml/2006/main",
        _STRICT_BASE + "wordprocessingml/main",
    }
    for run in root.iter():
        namespace, _, name = str(run.tag).removeprefix("{").partition("}")
        if namespace not in words or name != "r":
            continue
        children = []
        for node in run:
            if node.tag != Q(namespace, "t") or not any(c in (node.text or "") for c in "\t\n\r"):
                children.append(node)
                continue
            changed = True
            for segment in re.split(r"([\t\n\r])", node.text or ""):
                if segment in {"\t", "\n", "\r"}:
                    children.append(
                        ET.Element(Q(namespace, {"\t": "tab", "\n": "br", "\r": "cr"}[segment]))
                    )
                elif segment:
                    text = ET.Element(node.tag, dict(node.attrib))
                    text.text = segment
                    if segment.strip() != segment:
                        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    children.append(text)
            if children:
                children[-1].tail = node.tail
        run[:] = children
    if not changed:
        return data
    from ..inspect import serialize_xml_with_preserved_qnames

    return serialize_xml_with_preserved_qnames(root)


_SHAPE_TREE_NODES = frozenset({"wsp", "wgp", "grpSp", "wpc"})
_DRAWING_IMAGE_NODES = frozenset({"pic", "blip", "chart", "relIds", "contentPart", "graphicFrame"})
_VML_IMAGE_NODES = frozenset({"imagedata", "OLEObject", "ink", "signatureline", "binData"})
_RELATIONSHIP_ATTRIBUTES = frozenset({"id", "embed", "link", "pict"})
_TRUE_VML = frozenset({"t", "true", "1", "on"})


def _shape_text(node: ET.Element) -> list[ET.Element | str]:
    """Text of one selected ``w:drawing`` or ``w:pict`` as stories in reading order.

    A story is a ``w:txbxContent`` element (DrawingML text box or WordArt, VML text
    box) or a string of newline-separated lines (an active VML ``v:textpath``
    WordArt). The caller has already chosen one Markup Compatibility branch, so no
    alternative representation is read twice. A DrawingML ``wps:linkedTxbx``
    continuation owns no story: the chain's text lives once in its head. Anything
    outside a story that carries an image, chart, diagram, ink or embedded object,
    and any drawing with no text box or WordArt at all, is refused with
    ``text_content_unsupported``. Nested drawings inside a story belong to that
    story's own paragraphs and are not visited here.
    """

    def local(item: ET.Element) -> str:
        return item.tag.rpartition("}")[2] if isinstance(item.tag, str) else ""

    drawing = local(node) == "drawing"
    stories: list[ET.Element | str] = []
    bearing = False
    stack = [node]
    while stack:
        item = stack.pop()
        name = local(item)
        if drawing:
            unsupported = name in _DRAWING_IMAGE_NODES or (
                name == "graphicData"
                and any(local(child) not in _SHAPE_TREE_NODES for child in item)
            )
        else:
            unsupported = name in _VML_IMAGE_NODES or any(
                key.rpartition("}")[2] == "relid"
                or (
                    key.partition("}")[0].endswith("/relationships")
                    and key.rpartition("}")[2] in _RELATIONSHIP_ATTRIBUTES
                )
                for key in item.attrib
            )
        if unsupported:
            raise SaveOptionsError("text_content_unsupported")
        if name == ("txbx" if drawing else "textbox"):
            bearing = True
            stories.extend(child for child in item if local(child) == "txbxContent")
        elif drawing and name == "linkedTxbx":
            bearing = True
        elif not drawing and name == "textpath":
            if item.get("on", "f").lower() in _TRUE_VML:
                bearing = True
                stories.append(item.get("string", "").replace("\r\n", "\n").replace("\r", "\n"))
        if name != "txbxContent":
            stack.extend(reversed(item))
    if not bearing:
        raise SaveOptionsError("text_content_unsupported")
    return stories


def _save_text(graph: PackageGraph, options: TxtSaveOptions) -> str:
    _feature_html_export = feature_module("html_export")
    _MarkupCompatibilitySelection = _feature_html_export._MarkupCompatibilitySelection
    _MathML = _feature_html_export._MathML
    ExportStyles = _feature_html_export._Styles
    _mathml_text = feature_module("markdown")._mathml_text
    _feature_native_layout = feature_module("native_layout")
    _Numbering = _feature_native_layout._Numbering
    _Styles = _feature_native_layout._Styles

    styles = _Styles(graph)
    numbering = _Numbering(graph, styles)
    math_styles = ExportStyles(graph)
    counters: dict[int, dict[int, int]] = {}
    field_state = _FieldTextState("text_field_structure_invalid", SaveOptionsError)
    shapes: list[str] = []

    def local(node: ET.Element) -> str:
        return node.tag.rpartition("}")[2] if isinstance(node.tag, str) else ""

    def inline(node: ET.Element, *, hidden: bool = False) -> str:
        name = local(node)
        if name in {"pPr", "rPr", "del", "moveFrom", "instrText", "delInstrText"}:
            return ""
        if name == "fldChar":
            field_state.marker(node)
            return ""
        if name == "r":
            namespace = node.tag[1:].partition("}")[0]
            properties = node.find(Q(namespace, "rPr"))
            hidden = hidden or (
                properties is not None
                and any(
                    local(flag) in {"vanish", "webHidden"}
                    and flag.get(Q(namespace, "val"), "true").lower() not in {"0", "false", "off"}
                    for flag in properties
                )
            )
        if name in {"oMath", "oMathPara"}:
            if hidden or not field_state.visible:
                return ""
            namespace = node.tag[1:].partition("}")[0]
            converter = _MathML(namespace, math_styles, {})
            math = ET.fromstring(converter.argument(node))
            if options.office_math_mode == TxtOfficeMathExportMode.LINEAR:
                return _mathml_text(math)
            return "".join(math.itertext())
        if name in {"t", "tab", "br", "cr", "drawing", "pict", "object", "altChunk"}:
            if hidden or not field_state.visible:
                return ""
            if name == "t":
                return node.text or ""
            if name == "tab":
                return "\t"
            if name in {"br", "cr"}:
                kind = next((v for k, v in node.attrib.items() if k.endswith("}type")), "")
                return ("\f" if options.export_page_breaks else "") if kind == "page" else "\n"
            if name in {"drawing", "pict"}:
                for item in _shape_text(node):
                    chunk = item.split("\n") if isinstance(item, str) else story(item)
                    if any(chunk):
                        shapes.extend(chunk)
                return ""
            raise SaveOptionsError("text_content_unsupported")
        return "".join(inline(child, hidden=hidden) for child in node)

    def paragraph(node: ET.Element) -> list[str]:
        nonlocal shapes
        outer, shapes = shapes, []
        properties = styles.paragraph_properties(
            next((child for child in node if local(child) == "pPr"), None)
        )
        inherited = dict(styles.default_paragraph)
        if style_id := properties.get("pstyle"):
            inherited.update(styles.paragraph_style(style_id)[0])
        inherited.update(properties)
        prefix = ""
        if inherited.get("num_pr") and not inherited.get("num_off"):
            identifier = inherited.get("num_id")
            if not isinstance(identifier, int):
                raise SaveOptionsError("text_numbering_unsupported")
            index = inherited.get("num_ilvl", 0)
            level = numbering.level(identifier, index)
            values = counters.setdefault(identifier, {})
            values[index] = values.get(index, level.start - 1) + 1
            for deeper in tuple(values):
                if deeper > index:
                    restart = numbering.level(identifier, deeper).restart
                    if restart != 0 and (restart is None or restart - 1 <= index):
                        del values[deeper]
            prefix = options.list_indentation.character * (options.list_indentation.count * index)
            if options.export_list_labels:
                prefix += numbering.marker_text(identifier, level, values)
                prefix += {"tab": "\t", "space": " ", "nothing": ""}[level.suffix]
        page = "\f" if options.export_page_breaks and inherited.get("page_break_before") else ""
        text = page + prefix + inline(node)
        anchored, shapes = shapes, outer
        return [text, *anchored]

    def blocks(node: ET.Element) -> list[str]:
        result: list[str] = []
        for child in node:
            name = local(child)
            if name == "p":
                result.extend(paragraph(child))
            elif name == "tbl" and options.export_tables:
                for row in child:
                    if local(row) == "tr":
                        result.append(
                            options.table_cell_separator.join(
                                options.paragraph_break.join(blocks(cell))
                                for cell in row
                                if local(cell) == "tc"
                            )
                        )
            elif name not in {"tbl", "sectPr", "pPr", "rPr", "del", "moveFrom"}:
                result.extend(blocks(child))
        return result

    def story_blocks(node: ET.Element) -> list[str]:
        nonlocal field_state
        field_state = _FieldTextState("text_field_structure_invalid", SaveOptionsError)
        result = blocks(node)
        field_state.finish()
        return result

    def story(content: ET.Element) -> list[str]:
        """One shape story with field state of its own; the host's is untouched."""
        nonlocal field_state
        outer = field_state
        field_state = _FieldTextState("text_field_structure_invalid", SaveOptionsError)
        try:
            result = blocks(content)
            field_state.finish()
        finally:
            field_state = outer
        return result

    def part_root(part: str) -> ET.Element:
        data = graph.read_part_bytes(part)
        root = _parse_part(data, graph.limits, editable=False)
        _MarkupCompatibilitySelection(data, root).select(root)
        return root

    main = part_root(_main_part(graph))
    body = next((child for child in main if local(child) == "body"), None)
    if body is None:
        raise SaveOptionsError("text_document_body_missing")
    headers: list[str] = []
    footers: list[str] = []
    if options.header_footer_mode != TxtExportHeadersFootersMode.NONE:
        relationships = {rel.id: rel for rel in graph.relationships_for(_main_part(graph))}
        seen: set[str] = set()
        for reference in main.iter():
            kind = local(reference)
            if kind not in {"headerReference", "footerReference"}:
                continue
            variant = next((v for k, v in reference.attrib.items() if k.endswith("}type")), "")
            if (
                options.header_footer_mode == TxtExportHeadersFootersMode.PRIMARY_ONLY
                and variant != "default"
            ):
                continue
            identifier = next((v for k, v in reference.attrib.items() if k.endswith("}id")), "")
            rel = relationships.get(identifier)
            if rel is None or rel.is_external:
                raise SaveOptionsError("text_story_relationship_invalid")
            target = rel.resolved_target()
            if target in seen:
                continue
            seen.add(target)
            root = part_root(target)
            (headers if kind == "headerReference" else footers).extend(story_blocks(root))
    values = headers + story_blocks(body) + footers
    return options.paragraph_break.join(values) + (options.paragraph_break if values else "")


_TOC_INSTRUCTION = 'TOC \\o "1-3" \\h \\z \\u'


def _paragraph_target(
    records: Sequence[Mapping[str, Any]], paragraph: int, part: str = _DOCUMENT
) -> Mapping[str, Any]:
    """The discovered record of one paragraph target, or an empty one that binds nothing."""
    return next(
        (
            record
            for record in records
            if record.get("paragraph") == paragraph and record.get("part", part) == part
        ),
        {},
    )


def _span(
    target: Mapping[str, Any], start: int | None, end: int | None, expected: str | None
) -> tuple[int, int, str]:
    """A visible-text range of a discovered paragraph and the text it must still hold.

    ``start`` defaults to the paragraph start and ``end`` to its end; ``expected``
    defaults to the session's current text of the range.
    """
    text: str = target.get("text", "")
    begin = 0 if start is None else start
    stop = len(text) if end is None else end
    return begin, stop, text[begin:stop] if expected is None else expected


class RichObjectEdits(EditHost):
    """Comments, notes, bookmarks, hyperlinks, fields, content controls, watermarks, and
    text boxes, each through its file-to-file operation module.

    ``paragraph`` is the 1-based paragraph index that module's own discovery reports
    for ``part``. Every guard the module binds (the structure fingerprint, the text a
    range must still hold, an existing comment's fingerprint) is read from the
    session's current package; ``expected_old_text`` may be passed as an extra guard.
    """

    # -- comments -------------------------------------------------------------

    def insert_comment(
        self,
        paragraph: int,
        start: int,
        end: int,
        text: str,
        *,
        author: str,
        date: str,
        initials: str | None = None,
        expected_old_text: str | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Anchor a new comment to visible characters ``[start, end)`` of ``paragraph``.

        ``date`` is an ISO 8601 timestamp with a time zone. The report's ``detail``
        lists every comment with its ``comment_id`` and ``comment_sha256``.
        """
        _feature_comments = feature_module("comments")
        CommentInsert = _feature_comments.CommentInsert
        CommentRequest = _feature_comments.CommentRequest
        author_comment = _feature_comments.author_comment
        discover_comment_targets = _feature_comments.discover_comment_targets

        def build(staged: Staged) -> CommentRequest:
            discovered = discover_comment_targets(staged.source, part=part)
            target = _paragraph_target(discovered.get("targets", []), paragraph, part)
            begin, stop, old = _span(target, start, end, expected_old_text)
            return CommentRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                edits=(
                    CommentInsert(
                        part=part,
                        paragraph=paragraph,
                        start=begin,
                        end=stop,
                        expected_old_text=old,
                        structure_fingerprint_sha256=bound_fingerprint(
                            target.get("structure_fingerprint_sha256")
                        ),
                        author=author,
                        initials=initials,
                        date=date,
                        text=text,
                    ),
                ),
            )

        return self._apply_operation("comment", build, author_comment)

    def reply_to_comment(
        self,
        comment_id: int,
        text: str,
        *,
        author: str,
        date: str,
        initials: str | None = None,
    ) -> SessionEditReport:
        """Append a reply to the thread of the existing comment ``comment_id``."""
        from .._comment_threads import CommentReply

        _feature_comments = feature_module("comments")
        CommentRequest = _feature_comments.CommentRequest
        author_comment = _feature_comments.author_comment
        discover_comment_targets = _feature_comments.discover_comment_targets

        def build(staged: Staged) -> CommentRequest:
            comments = discover_comment_targets(staged.source).get("comments", [])
            parent: Mapping[str, Any] = next(
                (c for c in comments if c.get("comment_id") == comment_id), {}
            )
            return CommentRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                edits=(
                    CommentReply(
                        comment_id=comment_id,
                        comment_sha256=bound_fingerprint(parent.get("comment_sha256")),
                        author=author,
                        initials=initials,
                        date=date,
                        text=text,
                    ),
                ),
            )

        return self._apply_operation("comment", build, author_comment)

    # -- footnotes and endnotes -------------------------------------------------

    def insert_footnote(
        self,
        paragraph: int,
        text: str,
        *,
        start: int | None = None,
        end: int | None = None,
        anchor_boundary: str = "after",
        expected_old_text: str | None = None,
        custom_mark: str | None = None,
    ) -> SessionEditReport:
        """Insert a footnote whose reference sits at ``anchor_boundary`` of the range
        ``[start, end)`` of body ``paragraph`` (by default the end of the paragraph)."""
        return self._note(
            "footnote",
            paragraph,
            text,
            start,
            end,
            anchor_boundary,
            expected_old_text,
            custom_mark,
        )

    def insert_endnote(
        self,
        paragraph: int,
        text: str,
        *,
        start: int | None = None,
        end: int | None = None,
        anchor_boundary: str = "after",
        expected_old_text: str | None = None,
        custom_mark: str | None = None,
    ) -> SessionEditReport:
        """Insert an endnote, addressed as :meth:`insert_footnote` addresses a footnote."""
        return self._note(
            "endnote",
            paragraph,
            text,
            start,
            end,
            anchor_boundary,
            expected_old_text,
            custom_mark,
        )

    def _note(
        self,
        kind: str,
        paragraph: int,
        text: str,
        start: int | None,
        end: int | None,
        anchor_boundary: str,
        expected_old_text: str | None,
        custom_mark: str | None,
    ) -> SessionEditReport:
        _feature_notes = feature_module("notes")
        NoteRequest = _feature_notes.NoteRequest
        author_note = _feature_notes.author_note
        discover_note_targets = _feature_notes.discover_note_targets

        def build(staged: Staged) -> NoteRequest:
            discovered = discover_note_targets(staged.source, kind)
            target = _paragraph_target(discovered.get("targets", []), paragraph)
            begin, stop, old = _span(target, start, end, expected_old_text)
            return NoteRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                kind=kind,
                paragraph=paragraph,
                start=begin,
                end=stop,
                expected_old_text=old,
                structure_fingerprint_sha256=bound_fingerprint(
                    target.get("structure_fingerprint_sha256")
                ),
                text=text,
                anchor_boundary=anchor_boundary,
                custom_mark=custom_mark,
            )

        return self._apply_operation("note", build, author_note)

    # -- bookmarks and hyperlinks -----------------------------------------------

    def insert_bookmark(
        self,
        name: str,
        paragraph: int,
        *,
        start: int | None = None,
        end: int | None = None,
        expected_old_text: str | None = None,
    ) -> SessionEditReport:
        """Bookmark visible characters ``[start, end)`` of body ``paragraph`` (by default
        the whole paragraph) as ``name``."""
        _feature_bookmarks = feature_module("bookmarks")
        BookmarkInsert = _feature_bookmarks.BookmarkInsert
        BookmarkRequest = _feature_bookmarks.BookmarkRequest
        author_bookmark = _feature_bookmarks.author_bookmark
        discover_bookmark_targets = _feature_bookmarks.discover_bookmark_targets

        def build(staged: Staged) -> BookmarkRequest:
            discovered = discover_bookmark_targets(staged.source)
            target = _paragraph_target(discovered.get("targets", []), paragraph)
            begin, stop, old = _span(target, start, end, expected_old_text)
            return BookmarkRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                edits=(
                    BookmarkInsert(
                        name=name,
                        part=_DOCUMENT,
                        paragraph=paragraph,
                        start=begin,
                        end=stop,
                        expected_old_text=old,
                        structure_fingerprint_sha256=bound_fingerprint(
                            target.get("structure_fingerprint_sha256")
                        ),
                    ),
                ),
            )

        return self._apply_operation("bookmark", build, author_bookmark)

    def insert_hyperlink(
        self,
        paragraph: int,
        start: int,
        end: int,
        target: str,
        *,
        kind: str | None = None,
        tooltip: str | None = None,
        expected_old_text: str | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Link visible characters ``[start, end)`` of ``paragraph`` to ``target``.

        ``target`` is an allowlisted URL (an external link) or the name of one existing
        bookmark (an internal link); ``kind`` (``"external"`` or ``"internal"``)
        overrides the inference.
        """
        _feature_links = feature_module("links")
        HyperlinkInsert = _feature_links.HyperlinkInsert
        HyperlinkRequest = _feature_links.HyperlinkRequest
        author_hyperlink = _feature_links.author_hyperlink
        discover_hyperlink_targets = _feature_links.discover_hyperlink_targets

        def build(staged: Staged) -> HyperlinkRequest:
            discovered = discover_hyperlink_targets(staged.source, part=part)
            record = _paragraph_target(discovered.get("targets", []), paragraph, part)
            begin, stop, old = _span(record, start, end, expected_old_text)
            return HyperlinkRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                edits=(
                    HyperlinkInsert(
                        paragraph=paragraph,
                        start=begin,
                        end=stop,
                        expected_old_text=old,
                        structure_fingerprint_sha256=bound_fingerprint(
                            record.get("structure_fingerprint_sha256")
                        ),
                        target=target,
                        part=part,
                        kind=kind,
                        tooltip=tooltip,
                    ),
                ),
            )

        return self._apply_operation("hyperlink", build, author_hyperlink)

    # -- fields -------------------------------------------------------------------

    def insert_field(
        self,
        paragraph: int,
        instruction: str,
        *,
        start: int | None = None,
        end: int | None = None,
        result_text: str | None = None,
        expected_old_text: str | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert one complex field with ``instruction`` into ``paragraph``.

        The field replaces visible characters ``[start, end)``; by default it is
        inserted, replacing nothing, at the end of the paragraph.
        """
        _feature_field_insert = feature_module("field_insert")
        FieldInsertRequest = _feature_field_insert.FieldInsertRequest
        discover_field_insert_targets = _feature_field_insert.discover_field_insert_targets
        insert_field = _feature_field_insert.insert_field

        def build(staged: Staged) -> FieldInsertRequest:
            discovered = discover_field_insert_targets(staged.source, part=part)
            target = _paragraph_target(discovered.get("targets", []), paragraph, part)
            length = len(target.get("text", ""))
            begin = length if start is None else start
            stop = begin if end is None else end
            begin, stop, old = _span(target, begin, stop, expected_old_text)
            return FieldInsertRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                paragraph=paragraph,
                start=begin,
                end=stop,
                expected_old_text=old,
                structure_fingerprint_sha256=bound_fingerprint(
                    target.get("structure_fingerprint_sha256")
                ),
                instruction=instruction,
                part=part,
                result_text=result_text,
            )

        return self._apply_operation("field-insert", build, insert_field)

    def refresh_fields(
        self,
        *,
        fonts: Sequence[Any] = (),
        paragraph: int | None = None,
        part: str = _DOCUMENT,
        options: Any | None = None,
        evaluation_time: str | None = None,
        limits: Any | None = None,
    ) -> SessionEditReport:
        """Refresh the cached results of the fields of ``part``, or only those of
        ``paragraph``, as :func:`~wolfdocx.fields.refresh_fields` would.

        Page-dependent values (``PAGE``, ``NUMPAGES``, ``PAGEREF``, table-of-contents
        page numbers) come from the engine's native pagination, laid out with ``fonts``.
        """
        _feature_fields = feature_module("fields")
        FieldOptions = _feature_fields.FieldOptions
        FieldRefreshRequest = _feature_fields.FieldRefreshRequest
        FieldTarget = _feature_fields.FieldTarget
        discover_field_targets = _feature_fields.discover_field_targets
        refresh_fields = _feature_fields.refresh_fields
        RenderLimits = feature_module("native_layout").RenderLimits

        def build(staged: Staged) -> FieldRefreshRequest:
            selected: tuple[FieldTarget, ...] = ()
            if paragraph is not None:
                inventory = discover_field_targets(staged.source, part=part).get("fields", [])
                selected = tuple(
                    FieldTarget(
                        paragraph=entry["paragraph"],
                        field=entry["field"],
                        kind=entry["kind"] or "",
                        instruction=entry["instruction"],
                        structure_sha256=entry["structure_sha256"],
                        part=part,
                    )
                    for entry in inventory
                    if entry["paragraph"] == paragraph
                ) or (
                    # A paragraph without fields binds nothing, so the refresh refuses.
                    FieldTarget(paragraph, 1, "", "", bound_fingerprint(None), part),
                )
            return FieldRefreshRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                output=staged.destination,
                part=part,
                fields=selected,
                fonts=tuple(fonts),
                limits=limits or RenderLimits(),
                options=options or FieldOptions(),
                evaluation_time=evaluation_time,
            )

        return self._apply_operation("field-refresh", build, refresh_fields)

    def insert_table_of_contents(
        self,
        paragraph: int,
        *,
        fonts: Sequence[Any],
        instruction: str = _TOC_INSTRUCTION,
        position: str = "after",
        limits: Any | None = None,
    ) -> SessionEditReport:
        """Insert a table of contents in a new body paragraph before or after
        ``paragraph`` and refresh it against native pagination laid out with ``fonts``.

        The three steps (an empty paragraph, the ``TOC`` field, its refresh) land
        together or not at all.
        """
        toc = paragraph + 1 if position == "after" else paragraph
        return self._apply_steps(
            "table-of-contents",
            (
                lambda: self._edit_paragraph(
                    paragraph, _DOCUMENT, action="insert", position=position, text=""
                ),
                lambda: self.insert_field(toc, instruction),
                lambda: self.refresh_fields(fonts=fonts, paragraph=toc, limits=limits),
            ),
        )

    # -- content controls, watermarks, text boxes --------------------------------

    def insert_content_control(
        self,
        paragraph: int,
        control: Any | None = None,
    ) -> SessionEditReport:
        """Wrap direct body ``paragraph`` in a block-level content control.

        ``control`` defaults to a plain-text control with no tag or alias.
        """
        _feature_content_controls = feature_module("content_controls")
        ContentControlRequest = _feature_content_controls.ContentControlRequest
        ContentControlSpec = _feature_content_controls.ContentControlSpec
        author_content_control = _feature_content_controls.author_content_control
        content_control_structure_fingerprint = (
            _feature_content_controls.content_control_structure_fingerprint
        )

        def build(staged: Staged) -> ContentControlRequest:
            return ContentControlRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                target={"paragraph": paragraph, "part": _DOCUMENT},
                structure_fingerprint_sha256=bound_fingerprint(
                    content_control_structure_fingerprint(staged.source, paragraph=paragraph)
                ),
                control=ContentControlSpec() if control is None else control,
            )

        return self._apply_operation("content-control", build, author_content_control)

    def add_text_watermark(
        self, text: str, *, section: int = 1, variant: str = "default", **style: Any
    ) -> SessionEditReport:
        """Add a text watermark to the ``variant`` header of ``section``.

        ``style`` takes the :class:`~wolfdocx.watermarks.TextWatermark` fields
        (``color``, ``opacity``, ``rotation``, ``font_family``, ``font_size``, and
        the rest). Without an ``identifier`` the module seeds one from the digest of
        the session's current package, as it would from the file it reads.
        """
        _feature_watermarks = feature_module("watermarks")
        TextWatermark = _feature_watermarks.TextWatermark
        WatermarkRequest = _feature_watermarks.WatermarkRequest
        add_text_watermark = _feature_watermarks.add_text_watermark

        watermark = TextWatermark(text=text, **style)

        def build(staged: Staged) -> WatermarkRequest:
            return WatermarkRequest(
                staged.source,
                staged.source_sha256,
                staged.destination,
                section,
                variant=variant,
            )

        return self._apply_operation(
            "watermark", build, lambda request: add_text_watermark(request, watermark)
        )

    def insert_text_box(
        self,
        paragraph: int,
        text: str,
        *,
        extent_cx: int,
        extent_cy: int,
        anchor: Any,
        part: str = _DOCUMENT,
        **style: Any,
    ) -> SessionEditReport:
        """Anchor a page-positioned text box holding ``text`` to ``paragraph``, sized
        in EMU.

        ``style`` takes the optional :class:`~wolfdocx.text_boxes.TextBoxInsertRequest`
        fields (``font_face``, ``font_size``, ``bold``, ``fill``, ``border_color``,
        ``name``, and the rest).
        """
        _feature_text_boxes = feature_module("text_boxes")
        TextBoxInsertRequest = _feature_text_boxes.TextBoxInsertRequest
        author_text_box_insert = _feature_text_boxes.author_text_box_insert
        text_box_structure_fingerprint = _feature_text_boxes.text_box_structure_fingerprint

        def build(staged: Staged) -> TextBoxInsertRequest:
            return TextBoxInsertRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                paragraph=paragraph,
                structure_fingerprint_sha256=bound_fingerprint(
                    text_box_structure_fingerprint(staged.source, paragraph=paragraph, part=part)
                ),
                text=text,
                extent_cx=extent_cx,
                extent_cy=extent_cy,
                anchor=anchor,
                part=part,
                **style,
            )

        return self._apply_operation("text-box", build, author_text_box_insert)


class DocumentOperations(
    RunEdits,
    TableEdits,
    PictureEdits,
    SectionEdits,
    ListEdits,
    StylingEdits,
    RichObjectEdits,
):
    """One DOCX package opened once and edited in memory until :meth:`save`.

    Open paths or open, seekable binary streams; streams are read from the start,
    remain caller-owned and are copied to a private disk snapshot. Closed,
    non-seekable and read-failure inputs raise a named
    :class:`~wolfdocx.package.PackageError`. Default admission validates the main
    story and internal targets. For diagnostic callers, explicitly defer part
    validation through :class:`~wolfdocx.load_options.LoadOptions`.
    :meth:`append_paragraph`, :meth:`set_cell_text`, and :meth:`find_replace` return
    the report types of the stateless APIs with status ``"applied"``; every other
    edit returns a :class:`~wolfdocx.session_ops.SessionEditReport`. Nothing is
    written until :meth:`save`, which may be called more than once.

    :attr:`warnings` is cumulative diagnostic history until ``clear()``. Clearing
    it changes neither package bytes nor load, save, or render receipts.
    :meth:`render` renders current edits without publishing an intermediate DOCX.
    """

    @property
    def warnings(self) -> WarningCollection:
        """Load repairs, save refusals and render diagnostics in operation order."""
        if self._warnings is None:
            from ..render_warnings import WarningCollection

            self._warnings = WarningCollection()
        return self._warnings

    @property
    def editing_language(self) -> str | None:
        """The package's declared default language, else the load preference."""
        return None if self.load_report is None else self.load_report.editing_language

    @property
    def load_recovery(self) -> tuple[str, ...]:
        """Every repair a best-effort load made; empty for a whole package."""
        if self.load_report is None:
            return ()
        return tuple(str(action) for action in self.load_report.recovery)

    @property
    def part_names(self) -> tuple[str, ...]:
        """The session's current parts in package order.

        Tolerated ZIP metadata is never edited, so an entry is a part when the
        current graph admits it or a pending edit (such as a builder part) writes it.
        """
        return tuple(
            name for name in self._entries if name in self._graph.parts or name in self._changed
        )

    def _notify_load_report(self, report: LoadReport, previous: LoadReport | None) -> None:
        from ..render_warnings import RenderWarning, WarningCategory, WarningSource

        old_actions = () if previous is None else previous.recovery
        self.warnings._extend(
            RenderWarning(
                WarningCategory.RECOVERY,
                WarningSource.LOAD,
                str(action),
                kind=action.code,
                locator=action.target,
                reason=action.code,
            )
            for action in report.recovery
            if action not in old_actions
        )

    def _apply_initial_load_options(self, load_options: LoadOptions) -> None:
        if load_options is not None and (
            self.load_report.editing_language_source == "preference"
            or load_options.text_direction.value != "ltr"
            or load_options.language_preferences is not None
        ):
            for name in self._entries:
                if self._preference_story(name):
                    before = self.part_bytes(name)
                    after = self._with_load_preferences(name, before)
                    if after != before:
                        self._set_part(name, after)
            if self._changed:
                self._record("load-preferences", tuple(sorted(self._changed)))
        if load_options is not None:
            self._apply_shared_load_options(load_options)

    @property
    def web_add_ins(self) -> WebAddIns:
        """Inspect and atomically edit local add-in metadata and task panes.

        Import value records from :mod:`wolfdocx.web_extensions`; save through
        this session. References are metadata only and are never fetched.
        """
        from ..web_extensions import WebAddIns

        return WebAddIns(self)

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
                    from ..review import _CONFLICT_REVISION_TAGS, _SELECTABLE_REVISION_TAGS

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

    def save(
        self, path: str | Path | BinaryIO, *, options: OoxmlSaveOptions | None = None
    ) -> SessionSaveReport:
        """Stage and verify privately before publishing to a path or binary stream.

        An existing path is refused and never replaced. A stream must be empty and
        seekable so a failed publication can be rolled back without losing caller data.
        Caller progress exceptions propagate unchanged before publication.
        """
        source = self._source()
        sink = path if hasattr(path, "write") else None
        destination = None if sink is not None else Path(path).resolve()
        display = "<stream>" if destination is None else str(destination)
        changed = dict(self._changed)

        def refusal(code: str) -> SessionSaveReport:
            return self._record_save_report(
                SessionSaveReport(_REFUSED, code, self.source_sha256, display)
            )

        callback_error: Exception | None = None

        def progress(stage: str) -> None:
            nonlocal callback_error
            try:
                assert options is not None
                options._progress(stage)
            except Exception as exc:
                callback_error = exc
                raise

        if options is not None and not isinstance(options, OoxmlSaveOptions):
            return refusal("ooxml_save_options_invalid")
        if options is not None:
            try:
                options._validate_lifecycle()
                progress("read")
                if options.property_updates:
                    updated = _apply_save_properties(self._materialize(), options.property_updates)
                    selected = replace(
                        options,
                        property_updates=None,
                        progress_callback=lambda stage: (
                            progress(stage) if stage != "read" else None
                        ),
                    )
                    with Document.from_bytes(
                        updated, path=self.path, limits=self._graph.limits
                    ) as snapshot:
                        snapshot._signature_source_graph = self._signature_source_graph
                        report = snapshot.save(path, options=selected)
                    self._record_save_report(report)
                    return replace(
                        report,
                        source_sha256=self.source_sha256,
                        edits=tuple(self._edits),
                        load_recovery=self.load_recovery,
                    )
            except SaveOptionsError as exc:
                if exc is callback_error:
                    raise
                return refusal(exc.code)

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
            if (
                destination.suffix.lower() == ".docx"
                and self._graph.parts[_main_part(self._graph)].content_type
                in _MACRO_MAIN_CONTENT_TYPES
            ):
                return refusal("document_destination_extension_mismatch")
        removed = tuple(sorted(self._infos.keys() - self._entries.keys()))
        structural = {name for name in changed if _structural(name)}
        staged: Path | None = None
        try:
            if options is not None:
                progress("render")
            with tempfile.NamedTemporaryFile(
                dir=None if destination is None else destination.parent,
                prefix="wolfdocx-save-" if destination is None else f".{destination.name}.",
                suffix=".tmp",
                delete=False,
            ) as stream:
                staged = Path(stream.name)
                self._write_to(stream, source, changed)
            if not changed and not removed:
                # Exact-byte identity reuses admission, including opaque parts and
                # ZIP metadata. Do not parse an already admitted package again.
                digest = sha256_file(staged)
                verified = (digest, {}) if digest == self.source_sha256 else None
            else:
                verified = _verify_staged_changes(
                    staged,
                    {name: data for name, data in changed.items() if name not in structural},
                    limits=self._graph.limits,
                    unchanged_identical=lambda graph: self._unchanged_raw(graph, staged, changed),
                    unchanged_elements=0,
                    baseline_tree=self._current_tree,
                    parsed=self._checked,
                    check_notes=bool(changed.keys() & self._note_parts()) or bool(removed),
                )
            if verified is None:
                return refusal("document_staged_verification_failed")
            output_sha256 = verified[0]
            raw_copied_parts = len(self.part_names) - len(changed)
            if options is not None:
                original = staged.read_bytes()
                converted = _save_ooxml(original, options, self._graph.limits)
                if converted is not original:
                    raw_copied_parts = 0
                    output_sha256 = sha256_bytes(converted)
                    staged.write_bytes(converted)
                if options.signing_details is not None:
                    _feature_package_signatures = feature_module("package_signatures")
                    DocxSignatureRequest = _feature_package_signatures.DocxSignatureRequest
                    sign_docx = _feature_package_signatures.sign_docx

                    details = dict(options.signing_details)
                    if details.get("signing_time") is None:
                        raise SaveOptionsError("ooxml_signing_time_required")
                    signed = staged.with_suffix(".signed.docx")
                    try:
                        report = sign_docx(
                            DocxSignatureRequest(
                                staged,
                                signed,
                                output_sha256,
                                certificate=Path(details.pop("certificate")),
                                private_key=Path(details.pop("private_key")),
                                private_key_password_env=details.pop(
                                    "private_key_password_env", None
                                ),
                                **details,
                            )
                        )
                        if report.status != "published":
                            raise SaveOptionsError(report.code or "ooxml_signing_failed")
                        staged.write_bytes(signed.read_bytes())
                        output_sha256 = report.destination_sha256
                    finally:
                        signed.unlink(missing_ok=True)
                progress("publish")
            signature_source = self._signature_source_graph
            invalidated_signatures: tuple[str, ...] = ()
            if any(
                part.content_type == _SIGNATURE_CONTENT_TYPE
                for part in signature_source.parts.values()
            ):
                invalidated_signatures = feature_module(
                    "package_signatures"
                )._invalidated_signatures(signature_source, staged)
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
                    staged,
                    destination,
                    exists_code="document_destination_exists",
                    unavailable_code="document_publication_unavailable",
                )
        except SaveOptionsError as exc:
            if exc is callback_error:
                raise
            return refusal(str(exc))
        except OSError as exc:
            if exc is callback_error:
                raise
            return refusal("document_publication_unavailable")
        except (ValueError, zipfile.BadZipFile, struct.error) + _PART_ERRORS as exc:
            if exc is callback_error:
                raise
            return refusal("document_staged_verification_failed")
        finally:
            if staged is not None:
                staged.unlink(missing_ok=True)
        if linked is not None:
            return refusal(linked)
        return SessionSaveReport(
            _PUBLISHED,
            None,
            self.source_sha256,
            display,
            destination_sha256=output_sha256,
            changed_parts=tuple(sorted(name for name in changed if name in self._infos)),
            added_parts=tuple(sorted(name for name in changed if name not in self._infos)),
            removed_parts=removed,
            raw_copied_parts=raw_copied_parts,
            edits=tuple(self._edits),
            load_recovery=self.load_recovery,
            invalidated_signatures=invalidated_signatures,
        )

    def save_text(
        self, path: str | Path | BinaryIO, *, options: TxtSaveOptions | None = None
    ) -> SessionSaveReport:
        """Export with the same full TXT policy as ``TxtSaveOptions.save``.

        Refusals retain the session's original source identity; existing paths
        and caller-owned streams are never clobbered.
        Caller progress exceptions propagate unchanged before publication.
        """
        self._source()
        selected = TxtSaveOptions() if options is None else options
        display = "<stream>" if hasattr(path, "write") else str(Path(path).resolve())
        if not isinstance(selected, TxtSaveOptions):
            return self._record_save_report(
                SessionSaveReport(
                    _REFUSED,
                    "text_save_options_invalid",
                    self.source_sha256,
                    display,
                    content_type="text/plain",
                    written_format="txt",
                )
            )
        callback_error: Exception | None = None
        callback = selected.progress_callback
        if callback is not None:

            def progress(stage: str) -> None:
                nonlocal callback_error
                try:
                    callback(stage)
                except Exception as exc:
                    callback_error = exc
                    raise

            selected = replace(selected, progress_callback=progress)
        try:
            return self._record_save_report(selected.save(self, path))
        except SaveOptionsError as exc:
            if exc is callback_error:
                raise
            return self._record_save_report(
                SessionSaveReport(
                    _REFUSED,
                    exc.code,
                    self.source_sha256,
                    display,
                    content_type="text/plain",
                    written_format="txt",
                )
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

    def render(
        self,
        output_dir: str | Path,
        *,
        formats: Sequence[str] = ("pdf",),
        fonts: Sequence[Any] = (),
        scale: float = 1.0,
        limits: Any | None = None,
        **options: Any,
    ) -> Any:
        """Render current edits with native request options and retain diagnostics.

        ``options`` names the optional fields of
        :class:`~wolfdocx.native_render.NativeRenderRequest`, including font
        sources, substitution rules and save policy. Source bytes and their
        digest are bound to this session's current package. Existing render
        destinations and unsupported capabilities keep the native typed refusal.
        Font discovery callbacks still receive their original diagnostic records;
        callback exceptions propagate unchanged.
        """
        from ..render_warnings import font_discovery_warning, warnings_from_native_report

        FontSettings = feature_module("font_sources").FontSettings
        RenderLimits = feature_module("native_layout").RenderLimits
        _feature_native_render = feature_module("native_render")
        NativeRenderRequest = _feature_native_render.NativeRenderRequest
        render_native = _feature_native_render.render_native

        callback = options.pop("font_warning_callback", None)
        preferences = self._load_preferences
        if preferences is not None and preferences.font_settings is not None:
            options.setdefault("font_settings", preferences.font_settings)
        settings = options.get("font_settings")
        if isinstance(settings, FontSettings):
            callback = settings.warning_callback

        def font_warning(record: dict[str, str]) -> None:
            warning = font_discovery_warning(record)
            self.warnings._extend((warning,))
            if preferences is not None and preferences.warning_callback is not None:
                preferences.warning_callback(warning)
            if callback is not None:
                callback(record)

        # Native validation still sees non-callable callbacks and refuses them.
        # Settings take precedence over direct options; copy only the callback
        # field so reusable caller policy is never mutated.
        selected_callback = font_warning if callback is None or callable(callback) else callback
        if isinstance(settings, FontSettings):
            options["font_settings"] = replace(settings, warning_callback=selected_callback)

        with tempfile.TemporaryDirectory(
            prefix="wolfdocx-session-render-",
            dir=None if preferences is None else preferences.temp_folder,
        ) as staging:
            source = Path(staging) / "document.docx"
            with source.open("wb") as stream:
                self._write_to(stream, self._source(), self._changed)
            report = render_native(
                NativeRenderRequest(
                    source=source,
                    source_sha256=sha256_file(source),
                    output_dir=Path(output_dir),
                    fonts=tuple(fonts),
                    formats=tuple(formats),
                    scale=scale,
                    limits=RenderLimits() if limits is None else limits,
                    font_warning_callback=selected_callback,
                    **options,
                )
            )
        warnings = warnings_from_native_report(report)
        self.warnings._extend(warnings)
        if preferences is not None and preferences.warning_callback is not None:
            for warning in warnings:
                preferences.warning_callback(warning)
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

    def _preference_story(self, name: str, graph: PackageGraph | None = None) -> bool:
        from ..load_options import _MAIN_CONTENT_TYPES, _WML

        part = (graph or self._graph).parts.get(name)
        return part is not None and (
            part.content_type in _MAIN_CONTENT_TYPES
            or part.content_type
            in {
                f"{_WML}.{kind}+xml"
                for kind in ("header", "footer", "footnotes", "endnotes", "comments")
            }
        )

    def _with_load_preferences(
        self, name: str, data: bytes, graph: PackageGraph | None = None
    ) -> bytes:
        """Fill absent story formatting in the overlay, never in the source graph."""
        options = self._load_preferences
        if options is None or not self._preference_story(name, graph):
            return data
        language = (
            self.editing_language
            if self.load_report.editing_language_source == "preference"
            else None
        )
        script_preferences = options.language_preferences
        if (
            language is None
            and options.text_direction.value == "ltr"
            and script_preferences is None
        ):
            return data
        from ..inspect import serialize_xml_with_preserved_qnames
        from ..load_options import contains_right_to_left, first_strong_direction
        from ..ooxml_order import _PPR_ORDER, _RPR_ORDER, _sort_children
        from ..review import _word_namespace

        try:
            root = self._parse(data, editable=True)
            ns = _word_namespace(root)
        except _PART_ERRORS as exc:
            raise _load_refusal(
                "load_preference_story_invalid", "Load preference story is invalid.", name
            ) from exc
        if ns is None:
            raise _load_refusal(
                "load_preference_story_unsupported", "Load preference story is unsupported.", name
            )
        changed = False
        if language is not None or script_preferences is not None:
            defaults = {}
            if language is not None:
                defaults[Q(ns, "val")] = language
            if script_preferences is not None:
                for attribute, value in (
                    ("bidi", script_preferences.bidi),
                    ("eastAsia", script_preferences.east_asia),
                ):
                    if value is not None:
                        defaults[Q(ns, attribute)] = value
            for run in root.iter(Q(ns, "r")):
                properties = run.find(Q(ns, "rPr"))
                if properties is None:
                    properties = ET.Element(Q(ns, "rPr"))
                    run.insert(0, properties)
                node = properties.find(Q(ns, "lang"))
                missing = (
                    defaults
                    if node is None
                    else {key: value for key, value in defaults.items() if key not in node.attrib}
                )
                if missing:
                    if node is None:
                        node = ET.SubElement(properties, Q(ns, "lang"))
                    node.attrib.update(missing)
                    _sort_children(properties, _RPR_ORDER)
                    changed = True
        if options.text_direction.value != "ltr":
            for paragraph in root.iter(Q(ns, "p")):
                properties = paragraph.find(Q(ns, "pPr"))
                if properties is None:
                    properties = ET.Element(Q(ns, "pPr"))
                    paragraph.insert(0, properties)
                if properties.find(Q(ns, "bidi")) is None:
                    rtl = (
                        options.text_direction.value == "rtl"
                        or first_strong_direction(
                            "".join(node.text or "" for node in paragraph.iter(Q(ns, "t")))
                        )
                        == "rtl"
                    )
                    ET.SubElement(properties, Q(ns, "bidi"), {Q(ns, "val"): "1" if rtl else "0"})
                    _sort_children(properties, _PPR_ORDER)
                    changed = True
            for run in root.iter(Q(ns, "r")):
                text = "".join(node.text or "" for node in run.iter(Q(ns, "t")))
                if not contains_right_to_left(text):
                    continue
                properties = run.find(Q(ns, "rPr"))
                if properties is None:
                    properties = ET.Element(Q(ns, "rPr"))
                    run.insert(0, properties)
                if properties.find(Q(ns, "rtl")) is None:
                    ET.SubElement(properties, Q(ns, "rtl"), {Q(ns, "val"): "1"})
                    _sort_children(properties, _RPR_ORDER)
                    changed = True
        return serialize_xml_with_preserved_qnames(root) if changed else data

    def _apply_shared_load_options(self, options: LoadOptions) -> None:
        """Apply actual session transformations only for explicitly selected policies."""
        if options.word_version is not None:
            from ..load_options import _WORD_VERSIONS
            from ..review import _word_namespace

            ns = _word_namespace(self._editable(_main_part(self._graph))[1])
            assert ns is not None
            part = self._definition_part("settings")
            settings = (
                ET.Element(Q(ns, "settings"))
                if part is None
                else copy.deepcopy(self._editable(part)[1])
            )
            compat = settings.find(Q(ns, "compat"))
            if compat is None:
                compat = ET.SubElement(settings, Q(ns, "compat"))
            mode = next(
                (
                    node
                    for node in compat.findall(Q(ns, "compatSetting"))
                    if node.get(Q(ns, "name")) == "compatibilityMode"
                ),
                None,
            )
            if mode is None:
                mode = ET.SubElement(compat, Q(ns, "compatSetting"))
            mode.attrib.update(
                {
                    Q(ns, "name"): "compatibilityMode",
                    Q(ns, "uri"): "http://schemas.microsoft.com/office/word",
                    Q(ns, "val"): str(_WORD_VERSIONS[options.word_version]),
                }
            )
            self._replace_definitions("settings", settings)
        if options.shape_to_math:
            self._load_shape_math()
        if not options.preserve_include_picture:
            self._load_unlink_pictures()
        if options.update_dirty_fields:
            self._load_dirty_fields(options)

    def _load_shape_math(self) -> None:
        """Promote native DrawingML equations without guessing from a picture's pixels."""
        from ..inspect import serialize_xml_with_preserved_qnames
        from ..review import _word_namespace

        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        a14 = "http://schemas.microsoft.com/office/drawing/2010/main"
        office = "urn:schemas-microsoft-com:office:office"
        for name in tuple(self._entries):
            if not self._preference_story(name):
                continue
            root = copy.deepcopy(self._editable(name)[1])
            ns = _word_namespace(root)
            assert ns is not None
            parents = {child: parent for parent in root.iter() for child in parent}
            converted = False
            for run in list(root.iter(Q(ns, "r"))):
                for shape in list(run):
                    if shape.tag not in {Q(ns, "drawing"), Q(ns, "pict"), Q(ns, "object")}:
                        continue
                    equations = [
                        equation
                        for container in shape.iter(Q(a14, "m"))
                        for equation in container
                        if equation.tag == Q(math_ns, "oMath")
                    ]
                    if not equations:
                        if any(
                            node.get("ProgID", "").startswith(("Equation.", "MathType"))
                            for node in shape.iter(Q(office, "OLEObject"))
                        ):
                            raise _load_refusal(
                                "load_shape_math_unsupported",
                                "An embedded binary equation cannot be converted.",
                                name,
                            )
                        continue
                    # A shape with other visible content is not safely replaceable.
                    if any(
                        node.tag.rpartition("}")[2] in {"blip", "imagedata"}
                        for node in shape.iter()
                    ):
                        raise _load_refusal(
                            "load_shape_math_mixed_content",
                            "Equation shape also has image content.",
                            name,
                        )
                    parent = parents.get(run)
                    if parent is None:
                        raise _load_refusal(
                            "load_shape_math_context_unsupported",
                            "Equation context is unsupported.",
                            name,
                        )
                    position = list(parent).index(run)
                    if len(run) != 1:
                        raise _load_refusal(
                            "load_shape_math_mixed_content", "Equation run has other content.", name
                        )
                    parent.remove(run)
                    for offset, equation in enumerate(equations):
                        parent.insert(position + offset, copy.deepcopy(equation))
                    converted = True
            if converted:
                self._set_part(name, serialize_xml_with_preserved_qnames(root))
                self._record("load-shape-math", (name,))

    def _load_unlink_pictures(self) -> None:
        """Retain cached INCLUDEPICTURE drawings while removing field instructions."""
        from ..inspect import serialize_xml_with_preserved_qnames
        from ..review import _word_namespace

        field_nodes = feature_module("fields").field_nodes

        for name in tuple(self._entries):
            if not self._preference_story(name):
                continue
            root = copy.deepcopy(self._editable(name)[1])
            ns = _word_namespace(root)
            assert ns is not None
            fields = [
                entry
                for entry in field_nodes(root, ns, part=name)
                if entry.code == "INCLUDEPICTURE"
            ]
            if not fields:
                continue
            parents = {child: parent for parent in root.iter() for child in parent}
            for entry in reversed(fields):
                if entry.locked:
                    continue
                if entry.simple is not None:
                    parent = parents[entry.simple]
                    position = list(parent).index(entry.simple)
                    parent.remove(entry.simple)
                    for offset, child in enumerate(list(entry.simple)):
                        parent.insert(position + offset, child)
                else:
                    for node in entry.nodes:
                        parent = parents.get(node)
                        if parent is not None and node in list(parent):
                            parent.remove(node)
            self._set_part(name, serialize_xml_with_preserved_qnames(root))
            self._record("load-include-picture-unlink", (name,))

    def _load_dirty_fields(self, options: LoadOptions) -> None:
        """Refresh only dirty unlocked fields through the existing source-bound evaluator."""
        _feature_fields = feature_module("fields")
        FieldRefreshRequest = _feature_fields.FieldRefreshRequest
        FieldTarget = _feature_fields.FieldTarget
        discover_field_targets = _feature_fields.discover_field_targets
        refresh_fields = _feature_fields.refresh_fields

        for name in tuple(self._entries):
            if not self._preference_story(name):
                continue
            with tempfile.TemporaryDirectory(
                prefix="wolfdocx-load-fields-", dir=options.temp_folder
            ) as scratch:
                source = Path(scratch) / "source.docx"
                output = Path(scratch) / "updated.docx"
                current = self._materialize()
                source.write_bytes(current)
                inventory = discover_field_targets(source, part=name)
                if inventory.get("status") == "refused":
                    raise _load_refusal(
                        inventory.get("code") or "load_dirty_fields_invalid",
                        "Dirty field discovery failed.",
                        name,
                    )
                selected = tuple(
                    FieldTarget(
                        entry["paragraph"],
                        entry["field"],
                        entry["kind"] or "",
                        entry["instruction"],
                        entry["structure_sha256"],
                        name,
                    )
                    for entry in inventory.get("fields", [])
                    if entry["dirty"] and not entry["fldLock"]
                )
                if not selected:
                    continue
                options.notify("fields.refresh", 0, len(selected))
                completed = iter(range(1, len(selected) + 1))
                total = len(selected)
                report = refresh_fields(
                    FieldRefreshRequest(
                        source=source,
                        source_sha256=sha256_bytes(current),
                        output=output,
                        part=name,
                        fields=selected,
                        progress_callback=lambda field, steps=completed, total=total: (
                            options.notify("fields.refresh", next(steps), total)
                        ),
                    )
                )
                if not report.passed:
                    raise _load_refusal(
                        report.code or "load_dirty_fields_refused",
                        "Dirty field refresh refused.",
                        name,
                    )
                self._adopt("load-dirty-fields", output.read_bytes(), operation_detail(report))
                options.notify("fields.refresh", len(selected), len(selected))

    def _set_part(self, name: str, data: bytes) -> None:
        self._changed[name] = self._with_load_preferences(name, data)
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
                data = self._with_load_preferences(name, archive.read(info), graph)
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
