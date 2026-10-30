"""Source-bound DOCX body or referenced header/footer paragraph insertion, deletion, and target discovery.

Provides strict, schema-valid creation and deletion of paragraphs in
``word/document.xml`` and referenced header/footer parts, bound to source package
hashes and structural fingerprints.

Refuses unsupported or hazardous operations:
- Malformed requests (unknown keys, missing required fields, invalid types).
- Stale source hashes or target structural fingerprints.
- Occupied destination paths.
- Unreferenced or unsupported story parts, wrong part root tags, or invalid content types.
- Unsupported nested paragraph targets (inside text boxes or non-cell unsupported contexts).
- Active revision contexts, comment ranges, complex fields, or content controls.
- Annotation and range hazards (cross-paragraph bookmarks, referenced bookmarks, sibling markers).
- Final paragraph deletion (destroying document body, required final cell paragraph, terminal paragraph after nested tables, or required final header/footer paragraph).
- Section-boundary deletion (destroying paragraph-carried sectPr).
- Non-existent paragraph styles.
- Invalid XML 1.0 characters or character-limit exceedances.

Preserves everything else:
- Non-target package parts are byte-identical.
- All non-target elements retain exact order and XML serialization.
- Section properties and semantics are preserved.
- Package relationships are preserved.
"""

from __future__ import annotations

import copy
import hashlib
import json
import zipfile
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from xml.parsers import expat

from .authoring import (
    _MAX_TEXT_CHARACTERS,
    _XML_SPACE_ATTRIBUTE,
    _cell_structure_fingerprint,
    _ContextState,
    _start_tag_end,
    _unsupported_context_by_element,
)
from .content import _xml_trees_equal
from .formatting import _catalog_package_styles
from .inspect import (
    Q,
    _inspect_docx,
    _Inspection,
    _parse_bounded_xml,
    _remembered_inspection,
    _require_verify_mode,
    _XmlSafetyError,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .reporting import _supported_part
from .review import (
    _DOCUMENT,
    ReviewReport,
    _publish,
    _sha256,
    _StagedPackage,
    _word_namespace,
)

__all__ = [
    "ParagraphReport",
    "ParagraphRequest",
    "ParagraphTarget",
    "author_paragraph",
    "discover_paragraph_targets",
    "load_paragraph_request",
    "paragraph_authoring_structure_fingerprint",
    "parse_paragraph_request",
]

_DISCOVERED = "discovered"
_PUBLISHED = "published"
# A session edit held in memory until ``Document.save`` publishes it.
_APPLIED = "applied"
_REFUSED = "refused"
_SHA256_LENGTH = 64
_STYLES = "word/styles.xml"

_REQUIRED_REQUEST_FIELDS = frozenset(
    {"source", "destination", "source_sha256", "target", "structure_fingerprint_sha256"}
)
_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "action",
        "position",
        "text",
        "style",
        "offset",
    }
)
_ALLOWED_INSERT_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "action",
        "position",
        "text",
        "style",
    }
)
_ALLOWED_DELETE_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "action",
    }
)
_ALLOWED_SPLIT_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "action",
        "offset",
    }
)

_REQUIRED_TARGET_FIELDS = frozenset({"paragraph"})
_ALLOWED_TARGET_FIELDS = frozenset({"paragraph", "part", "offset"})


def _require_sha256(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != _SHA256_LENGTH:
        raise ValueError(f"{label}_invalid")
    if any(c not in "0123456789abcdefABCDEF" for c in value):
        raise ValueError(f"{label}_invalid")
    return value.lower()


def _sha256_file(path: Path) -> str:
    return _sha256(path)


def _visible_text(paragraph: ET.Element, ns: str) -> str:
    return "".join(t.text or "" for t in paragraph.iter(Q(ns, "t")))


def _paragraph_style(paragraph: ET.Element, ns: str) -> str | None:
    ppr = paragraph.find(Q(ns, "pPr"))
    if ppr is not None:
        pstyle = ppr.find(Q(ns, "pStyle"))
        if pstyle is not None:
            return pstyle.get(Q(ns, "val"))
    return None


def _is_valid_xml_text(text: str) -> bool:
    """True when text consists only of valid XML 1.0 characters."""
    for char in text:
        cp = ord(char)
        if cp in (0x9, 0xA, 0xD):
            continue
        if 0x20 <= cp <= 0xD7FF:
            continue
        if 0xE000 <= cp <= 0xFFFD:
            continue
        if 0x10000 <= cp <= 0x10FFFF:
            continue
        return False
    return True


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is byte-identical to source."""
    with zipfile.ZipFile(source) as src_zip, zipfile.ZipFile(candidate) as cand_zip:
        src_names = set(src_zip.namelist())
        cand_names = set(cand_zip.namelist())
        if src_names != cand_names:
            return False
        for name in src_names:
            if name in changed:
                continue
            if src_zip.read(name) != cand_zip.read(name):
                return False
    return True


def _paragraph_hazard_reason(paragraph: ET.Element, ns: str) -> str | None:
    """Inspect paragraph for unsupported child elements or range hazards."""
    for tag_name in ("drawing", "pict", "wgp"):
        if paragraph.findall(f".//{Q(ns, tag_name)}"):
            return "paragraph_drawing_unsupported"

    if (
        paragraph.findall(f".//{Q(ns, 'fldSimple')}")
        or paragraph.findall(f".//{Q(ns, 'instrText')}")
        or paragraph.findall(f".//{Q(ns, 'fldChar')}")
    ):
        return "paragraph_field_unsupported"

    if (
        paragraph.findall(f".//{Q(ns, 'ins')}")
        or paragraph.findall(f".//{Q(ns, 'del')}")
        or paragraph.findall(f".//{Q(ns, 'moveFrom')}")
        or paragraph.findall(f".//{Q(ns, 'moveTo')}")
        or paragraph.findall(f".//{Q(ns, 'rPrChange')}")
        or paragraph.findall(f".//{Q(ns, 'pPrChange')}")
        or paragraph.findall(f".//{Q(ns, 'sectPrChange')}")
    ):
        return "paragraph_revision_unsupported"

    if (
        paragraph.findall(f".//{Q(ns, 'commentRangeStart')}")
        or paragraph.findall(f".//{Q(ns, 'commentRangeEnd')}")
        or paragraph.findall(f".//{Q(ns, 'commentReference')}")
    ):
        return "paragraph_comment_unsupported"

    if paragraph.findall(f".//{Q(ns, 'sdt')}"):
        return "paragraph_control_unsupported"

    starts = [e.get(Q(ns, "id")) for e in paragraph.findall(f".//{Q(ns, 'bookmarkStart')}")]
    ends = [e.get(Q(ns, "id")) for e in paragraph.findall(f".//{Q(ns, 'bookmarkEnd')}")]
    if set(starts) != set(ends) or len(starts) != len(ends):
        return "paragraph_bookmark_unsupported"

    return None


def _resolve_target_story(
    graph: PackageGraph,
    part: str,
    load: Callable[[str], tuple[bytes, ET.Element]] | None = None,
) -> tuple[ET.Element | None, ET.Element | None, str | None, bytes | None, str | None]:
    """Validate part reference, content-type, root tag, and resolve (root, container, ns, part_bytes, error_code).

    ``load`` returns a part's current bytes and a tree parsed from them; by default
    both come from ``graph``.
    """
    if not _supported_part(graph, part):
        return None, None, None, None, "paragraph_part_unsupported"

    try:
        if load is None:
            part_bytes = graph.read_part_bytes(part)
            root = _parse_bounded_xml(part_bytes, limits=graph.limits, total_elements=[0])
        else:
            part_bytes, root = load(part)
    except (PackageError, ET.ParseError, KeyError, OSError, _XmlSafetyError):
        return None, None, None, None, "paragraph_package_invalid"

    ns = _word_namespace(root)
    if ns is None:
        return None, None, None, None, "paragraph_namespace_unsupported"

    if part == _DOCUMENT:
        if root.tag != Q(ns, "document"):
            return None, None, None, None, "paragraph_part_unsupported"
        container = root.find(Q(ns, "body"))
    elif part.startswith("word/header") and part.endswith(".xml"):
        if root.tag != Q(ns, "hdr"):
            return None, None, None, None, "paragraph_part_unsupported"
        container = root
    elif part.startswith("word/footer") and part.endswith(".xml"):
        if root.tag != Q(ns, "ftr"):
            return None, None, None, None, "paragraph_part_unsupported"
        container = root
    else:
        return None, None, None, None, "paragraph_part_unsupported"

    return root, container, ns, part_bytes, None


def _is_admitted_paragraph(
    paragraph: ET.Element,
    container: ET.Element,
    parent_map: dict[ET.Element, ET.Element],
    ns: str,
) -> bool:
    """True when paragraph is a direct child of container or a direct child of a table cell (w:tc)."""
    parent = parent_map.get(paragraph)
    if parent is None:
        return False
    if parent is container:
        return True
    if parent.tag == Q(ns, "tc"):
        curr: ET.Element | None = parent
        while curr is not None and curr is not container:
            if curr.tag not in (Q(ns, "tc"), Q(ns, "tr"), Q(ns, "tbl")):
                return False
            curr = parent_map.get(curr)
        return curr is container
    return False


def _cell_ancestor_revision_hazard(
    paragraph: ET.Element,
    container: ET.Element,
    parent_map: dict[ET.Element, ET.Element],
    ns: str,
) -> str | None:
    """Inspect ancestor table, row, and cell properties for tracked revisions."""
    parent = parent_map.get(paragraph)
    if parent is None or parent is container or parent.tag != Q(ns, "tc"):
        return None

    curr: ET.Element | None = parent
    while curr is not None and curr is not container:
        tag = curr.tag
        if tag == Q(ns, "tc"):
            tc_pr = curr.find(Q(ns, "tcPr"))
            if tc_pr is not None:
                for rev_tag in ("cellIns", "cellDel", "cellMerge", "tcPrChange"):
                    if tc_pr.find(f".//{Q(ns, rev_tag)}") is not None:
                        return "paragraph_revision_unsupported"
        elif tag == Q(ns, "tr"):
            tr_pr = curr.find(Q(ns, "trPr"))
            if tr_pr is not None:
                for rev_tag in ("ins", "del", "trPrChange"):
                    if tr_pr.find(f".//{Q(ns, rev_tag)}") is not None:
                        return "paragraph_revision_unsupported"
        elif tag == Q(ns, "tbl"):
            tbl_pr = curr.find(Q(ns, "tblPr"))
            if tbl_pr is not None and tbl_pr.find(f".//{Q(ns, 'tblPrChange')}") is not None:
                return "paragraph_revision_unsupported"
            tbl_grid = curr.find(Q(ns, "tblGrid"))
            if tbl_grid is not None and tbl_grid.find(f".//{Q(ns, 'tblGridChange')}") is not None:
                return "paragraph_revision_unsupported"
        curr = parent_map.get(curr)

    return None


_HAZARD_TAG_NAMES = (
    "bookmarkStart",
    "bookmarkEnd",
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
    "fldSimple",
    "instrText",
    "fldChar",
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "rPrChange",
    "pPrChange",
    "sectPrChange",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "sdt",
    "drawing",
    "pict",
    "wgp",
)
_REVISION_TAG_NAMES = (
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "rPrChange",
    "pPrChange",
    "sectPrChange",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
)


def _paragraph_hazard_callbacks(
    ns: str, paragraph_hazards: dict[int, str]
) -> tuple[
    frozenset[str],
    Callable[[ET.Element, dict[str, list[ET.Element]]], None],
    Callable[[ET.Element], None],
]:
    """Share grouped paragraph/range semantics between tree and lexical scans."""
    active_bookmarks: set[str] = set()
    active_comments: set[str] = set()
    active_fields = 0
    active_revisions: set[tuple[str, str]] = set()
    tag = {name: Q(ns, name) for name in (*_HAZARD_TAG_NAMES, "p")}
    hazard_tags = frozenset(tag[name] for name in _HAZARD_TAG_NAMES)
    revision_tags = tuple(tag[name] for name in _REVISION_TAG_NAMES)
    drawing_tags = (tag["drawing"], tag["pict"], tag["wgp"])
    id_attribute = Q(ns, "id")
    field_char_type = Q(ns, "fldCharType")
    no_elements: list[ET.Element] = []

    def inspect_paragraph(p: ET.Element, found: dict[str, list[ET.Element]]) -> None:
        nonlocal active_fields
        if not found and not (
            active_bookmarks or active_comments or active_fields or active_revisions
        ):
            return
        p_id = id(p)

        def descendants(name: str) -> list[ET.Element]:
            return found.get(tag[name], no_elements)

        entered_in_bookmark = bool(active_bookmarks)
        entered_in_comment = bool(active_comments)
        entered_in_field = active_fields > 0
        entered_in_revision = bool(active_revisions)

        has_bookmark_hazard = entered_in_bookmark
        starts = [e.get(id_attribute) for e in descendants("bookmarkStart")]
        ends = [e.get(id_attribute) for e in descendants("bookmarkEnd")]
        valid_starts = [s for s in starts if s is not None]
        valid_ends = [e for e in ends if e is not None]

        for end_id in valid_ends:
            if end_id in active_bookmarks:
                has_bookmark_hazard = True
                active_bookmarks.discard(end_id)
            elif end_id not in valid_starts:
                has_bookmark_hazard = True

        for start_id in valid_starts:
            if start_id not in valid_ends:
                has_bookmark_hazard = True
                active_bookmarks.add(start_id)

        if bool(active_bookmarks):
            has_bookmark_hazard = True

        if set(starts) != set(ends) or len(starts) != len(ends):
            has_bookmark_hazard = True

        has_comment_hazard = entered_in_comment
        c_starts = [e.get(id_attribute) for e in descendants("commentRangeStart")]
        c_ends = [e.get(id_attribute) for e in descendants("commentRangeEnd")]
        c_refs = descendants("commentReference")
        if c_starts or c_ends or c_refs:
            has_comment_hazard = True

        for cs in c_starts:
            if cs is not None:
                active_comments.add(cs)
        for ce in c_ends:
            if ce is not None:
                active_comments.discard(ce)

        if bool(active_comments):
            has_comment_hazard = True

        has_field_hazard = entered_in_field
        if descendants("fldSimple") or descendants("instrText") or descendants("fldChar"):
            has_field_hazard = True

        for fc in descendants("fldChar"):
            fc_type = fc.get(field_char_type)
            if fc_type == "begin":
                active_fields += 1
            elif fc_type == "end" and active_fields > 0:
                active_fields -= 1

        if active_fields > 0:
            has_field_hazard = True

        has_revision_hazard = entered_in_revision
        if any(revision in found for revision in revision_tags):
            has_revision_hazard = True

        for mr in descendants("moveFromRangeStart") + descendants("moveToRangeStart"):
            rid = mr.get(id_attribute)
            if rid is not None:
                active_revisions.add((mr.tag, rid))
        for mr in descendants("moveFromRangeEnd") + descendants("moveToRangeEnd"):
            rid = mr.get(id_attribute)
            if rid is not None:
                start_tag = (
                    tag["moveFromRangeStart"]
                    if mr.tag == tag["moveFromRangeEnd"]
                    else tag["moveToRangeStart"]
                )
                active_revisions.discard((start_tag, rid))

        if bool(active_revisions):
            has_revision_hazard = True

        has_control_hazard = tag["sdt"] in found

        has_drawing_hazard = any(drawing in found for drawing in drawing_tags)

        if has_drawing_hazard:
            paragraph_hazards[p_id] = "paragraph_drawing_unsupported"
        elif has_field_hazard:
            paragraph_hazards[p_id] = "paragraph_field_unsupported"
        elif has_revision_hazard:
            paragraph_hazards[p_id] = "paragraph_revision_unsupported"
        elif has_comment_hazard:
            paragraph_hazards[p_id] = "paragraph_comment_unsupported"
        elif has_control_hazard:
            paragraph_hazards[p_id] = "paragraph_control_unsupported"
        elif has_bookmark_hazard:
            paragraph_hazards[p_id] = "paragraph_bookmark_unsupported"

    def handle_non_p_elem(elem: ET.Element) -> None:
        nonlocal active_fields
        if elem.tag not in hazard_tags:
            return
        if elem.tag == Q(ns, "bookmarkStart"):
            bid = elem.get(Q(ns, "id"))
            if bid is not None:
                active_bookmarks.add(bid)
        elif elem.tag == Q(ns, "bookmarkEnd"):
            bid = elem.get(Q(ns, "id"))
            if bid is not None:
                active_bookmarks.discard(bid)
        elif elem.tag == Q(ns, "commentRangeStart"):
            cid = elem.get(Q(ns, "id"))
            if cid is not None:
                active_comments.add(cid)
        elif elem.tag == Q(ns, "commentRangeEnd"):
            cid = elem.get(Q(ns, "id"))
            if cid is not None:
                active_comments.discard(cid)
        elif elem.tag == Q(ns, "fldChar"):
            fc_type = elem.get(Q(ns, "fldCharType"))
            if fc_type == "begin":
                active_fields += 1
            elif fc_type == "end" and active_fields > 0:
                active_fields -= 1
        elif elem.tag in {
            Q(ns, "moveFromRangeStart"),
            Q(ns, "moveToRangeStart"),
        }:
            rid = elem.get(Q(ns, "id"))
            if rid is not None:
                active_revisions.add((elem.tag, rid))
        elif elem.tag in {Q(ns, "moveFromRangeEnd"), Q(ns, "moveToRangeEnd")}:
            rid = elem.get(Q(ns, "id"))
            if rid is not None:
                start_tag = (
                    Q(ns, "moveFromRangeStart")
                    if elem.tag == Q(ns, "moveFromRangeEnd")
                    else Q(ns, "moveToRangeStart")
                )
                active_revisions.discard((start_tag, rid))

    return hazard_tags, inspect_paragraph, handle_non_p_elem


def _index_paragraph_hazards(container: ET.Element, ns: str) -> dict[int, str]:
    """Identify paragraphs with local hazards or active range intersections."""
    paragraph_hazards: dict[int, str] = {}
    hazard_tags, inspect_paragraph, handle_non_p_elem = _paragraph_hazard_callbacks(
        ns, paragraph_hazards
    )
    paragraph_tag = Q(ns, "p")

    def walk(node: ET.Element) -> None:
        for child in node:
            if child.tag == paragraph_tag:
                found: dict[str, list[ET.Element]] = {}
                for element in child.iter():
                    if element.tag in hazard_tags:
                        found.setdefault(element.tag, []).append(element)
                inspect_paragraph(child, found)
            else:
                handle_non_p_elem(child)
                walk(child)

    walk(container)
    return paragraph_hazards


def _build_paragraph_element(
    text: str | None,
    style: str | None,
    ns: str,
) -> ET.Element:
    """Build a schema-ordered w:p element with optional pPr/pStyle and r/t."""
    new_p = ET.Element(Q(ns, "p"))
    if style is not None:
        ppr = ET.SubElement(new_p, Q(ns, "pPr"))
        pstyle = ET.SubElement(ppr, Q(ns, "pStyle"))
        pstyle.set(Q(ns, "val"), style)
    if text:
        r = ET.SubElement(new_p, Q(ns, "r"))
        t = ET.SubElement(r, Q(ns, "t"))
        t.text = text
        if text[:1].isspace() or text[-1:].isspace():
            t.set(_XML_SPACE_ATTRIBUTE, "preserve")
    return new_p


class _ParagraphAdmission:
    """Consume lexical events against a private, bounded, namespace-preserving tree.

    Record only the selected ancestor path/context. Group hazards in the same
    outer-paragraph order as discovery, including nested paragraphs in that
    outer paragraph's descendant buckets, rather than walking the tree again.
    """

    def __init__(
        self, root: ET.Element, container: ET.Element, ns: str, target: ET.Element
    ) -> None:
        self.elements = iter(root.iter())
        self.container = container
        self.target = target
        self.paragraph_tag = Q(ns, "p")
        self.context = _ContextState(ns)
        self.stack: list[tuple[ET.Element, bool, bool]] = []
        self.parent_map: dict[ET.Element, ET.Element] = {}
        self.unsupported_context = False
        self.context_recorded = False
        self.hazards: dict[int, str] = {}
        self.hazard_tags, self.inspect_paragraph, self.handle_non_p = _paragraph_hazard_callbacks(
            ns, self.hazards
        )
        self.outer_paragraph: ET.Element | None = None
        self.found: dict[str, list[ET.Element]] = {}

    def start(self, name: str) -> None:
        element = next(self.elements, None)
        if element is None or element.tag != ("{" + name if "}" in name else name):
            raise ValueError("paragraph_xml_span_invalid")
        ancestor_context = self.stack[-1][1] if self.stack else False
        inside_container = self.stack[-1][2] if self.stack else False
        if not self.context_recorded:
            unsupported, child_context = self.context.enter(element, ancestor_context)
            if element is self.target:
                self.unsupported_context = unsupported
                self.context_recorded = True
                path = [entry[0] for entry in self.stack] + [element]
                self.parent_map = dict(zip(path[1:], path))
        else:
            child_context = False
        if inside_container:
            if self.outer_paragraph is None:
                if element.tag == self.paragraph_tag:
                    self.outer_paragraph = element
                else:
                    self.handle_non_p(element)
            if self.outer_paragraph is not None and element.tag in self.hazard_tags:
                self.found.setdefault(element.tag, []).append(element)
        self.stack.append((element, child_context, inside_container or element is self.container))

    def end(self) -> None:
        element, _, _ = self.stack.pop()
        if element is self.outer_paragraph:
            self.inspect_paragraph(element, self.found)
            self.outer_paragraph = None
            self.found = {}


def _paragraph_xml_span(
    data: bytes,
    ns: str,
    paragraph: int,
    *,
    admission: _ParagraphAdmission | None = None,
) -> tuple[int, int, int, str]:
    """Locate an element and its following lexical gap in the original encoding.

    Expat offsets are byte offsets, including a BOM. Empty-element end callbacks
    point *after* the start tag, unlike ordinary closing-tag callbacks.
    """
    if data.startswith((b"\xff\xfe", b"<\x00")):
        encoding = "utf-16-le"
    elif data.startswith((b"\xfe\xff", b"\x00<")):
        encoding = "utf-16-be"
    else:
        encoding = "utf-8"
    wide = encoding.startswith("utf-16")
    parser = expat.ParserCreate(namespace_separator="}")
    depth = 0
    occurrence = 0
    target_depth = None
    start = end = after = None
    opening_end = None
    empty = False
    target_name = f"{ns}}}p"

    def declaration(version: str, declared: str | None, standalone: int) -> None:
        nonlocal encoding
        if declared and not wide:
            encoding = declared

    def refuse_doctype(*args: Any) -> None:
        raise ValueError("DTD and entity declarations are forbidden")

    def tag_end(offset: int) -> int:
        if not wide:
            result = _start_tag_end(data, offset)
            if result is None:
                raise ValueError("paragraph_xml_span_invalid")
            return result
        quote = None
        quotes = ('"'.encode(encoding), "'".encode(encoding))
        close = ">".encode(encoding)
        for index in range(offset, len(data), 2):
            unit = data[index : index + 2]
            if quote is not None:
                if unit == quote:
                    quote = None
            elif unit in quotes:
                quote = unit
            elif unit == close:
                return index + 2
        raise ValueError("paragraph_xml_span_invalid")

    def start_element(name: str, attributes: dict[str, str]) -> None:
        nonlocal depth, occurrence, target_depth, start, opening_end, empty, after
        if admission is not None:
            admission.start(name)
        if end is not None and after is None:
            after = parser.CurrentByteIndex
        depth += 1
        if name == target_name:
            occurrence += 1
            if occurrence == paragraph:
                start = parser.CurrentByteIndex
                opening_end = tag_end(start)
                empty = data[start:opening_end].endswith("/>".encode(encoding))
                target_depth = depth

    def end_element(name: str) -> None:
        nonlocal depth, end, after, target_depth
        if admission is not None:
            admission.end()
        if end is not None and after is None:
            after = parser.CurrentByteIndex
        if depth == target_depth and name == target_name:
            end = opening_end if empty else tag_end(parser.CurrentByteIndex)
            target_depth = None
        depth -= 1

    parser.XmlDeclHandler = declaration
    parser.StartDoctypeDeclHandler = refuse_doctype
    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    parser.Parse(data, True)
    if start is None or end is None or after is None or not start < end <= after:
        raise ValueError("paragraph_xml_span_invalid")
    return start, end, after, encoding


def _paragraph_xml_fragment(paragraph: ET.Element, encoding: str) -> bytes:
    """Serialize owned paragraph content, excluding its unrelated lexical tail."""
    fragment = copy.copy(paragraph)
    fragment.tail = None
    # The shared serializer emits a UTF-8 declaration followed by one newline.
    xml = serialize_xml_with_preserved_qnames(fragment).partition(b"\n")[2]
    return xml.decode("utf-8").encode(encoding, errors="xmlcharrefreplace")


_SUPPORTED_RUN_CONTROLS = frozenset({"tab", "br", "cr", "noBreakHyphen", "softHyphen"})


@dataclass(frozen=True)
class ParagraphTarget:
    """One candidate body paragraph target for authoring."""

    part: str
    paragraph: int
    text: str
    structure_fingerprint_sha256: str
    style: str | None = None
    empty: bool = False
    supported: bool = True
    unsupported_reason: str | None = None

    @property
    def target(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "paragraph": self.paragraph,
        }

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "part": self.part,
            "paragraph": self.paragraph,
            "text": self.text,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "empty": self.empty,
            "supported": self.supported,
            "target": self.target,
        }
        if self.style is not None:
            result["style"] = self.style
        if self.unsupported_reason is not None:
            result["unsupported_reason"] = self.unsupported_reason
        return result


@dataclass(frozen=True)
class ParagraphRequest:
    """One source-bound request to insert, delete, or split a body paragraph."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    action: str = "insert"
    position: str | None = None
    text: str | None = None
    style: str | None = None
    offset: int | None = None

    def __init__(
        self,
        *,
        source: str | Path,
        destination: str | Path,
        source_sha256: str,
        structure_fingerprint_sha256: str,
        target: Mapping[str, Any] | None = None,
        paragraph: int | None = None,
        part: str = _DOCUMENT,
        action: str = "insert",
        position: str | None = None,
        text: str | None = None,
        style: str | None = None,
        offset: int | None = None,
    ) -> None:
        source_path = Path(source) if isinstance(source, (str, Path)) else None
        if not isinstance(source_path, Path) or not str(source_path):
            raise ValueError("paragraph_source_invalid")

        dest_path = Path(destination) if isinstance(destination, (str, Path)) else None
        if not isinstance(dest_path, Path) or not str(dest_path):
            raise ValueError("paragraph_destination_invalid")

        valid_source_sha = _require_sha256(source_sha256, "paragraph_source_sha256")
        valid_fp = _require_sha256(
            structure_fingerprint_sha256, "paragraph_structure_fingerprint_sha256"
        )

        if target is not None:
            if not isinstance(target, Mapping):
                raise TypeError("paragraph_target_invalid")
            target_keys = set(target)
            if not target_keys.issubset(
                _ALLOWED_TARGET_FIELDS
            ) or not _REQUIRED_TARGET_FIELDS.issubset(target_keys):
                raise ValueError("paragraph_target_fields_invalid")
            target_p = target["paragraph"]
            target_part = target.get("part", _DOCUMENT)
        elif paragraph is not None:
            target_p = paragraph
            target_part = part
        else:
            raise ValueError("paragraph_target_missing")

        if isinstance(target_p, bool) or not isinstance(target_p, int) or target_p < 1:
            raise ValueError("paragraph_paragraph_invalid")
        if not isinstance(target_part, str) or not target_part:
            raise ValueError("paragraph_part_invalid")

        if not isinstance(action, str) or action not in ("insert", "delete", "split"):
            raise ValueError("paragraph_action_invalid")

        if action == "delete":
            if position is not None or text is not None or style is not None or offset is not None:
                raise ValueError("paragraph_request_fields_invalid")
            final_position = None
            final_text = None
            final_style = None
            final_offset = None
        elif action == "split":
            if position is not None or text is not None or style is not None:
                raise ValueError("paragraph_request_fields_invalid")
            if offset is None and target is not None and "offset" in target:
                offset = target["offset"]
            if offset is None:
                raise ValueError("paragraph_offset_missing")
            if isinstance(offset, bool) or not isinstance(offset, int) or offset < 0:
                raise ValueError("paragraph_offset_invalid")
            final_position = None
            final_text = None
            final_style = None
            final_offset = offset
        else:
            if offset is not None:
                raise ValueError("paragraph_request_fields_invalid")
            final_offset = None
            final_position = "after" if position is None else position
            if not isinstance(final_position, str) or final_position not in (
                "before",
                "after",
            ):
                raise ValueError("paragraph_position_invalid")

            if text is None:
                final_text = ""
            elif not isinstance(text, str):
                raise TypeError("paragraph_text_invalid")
            else:
                if len(text) > _MAX_TEXT_CHARACTERS:
                    raise ValueError("paragraph_text_limit_exceeded")
                if not _is_valid_xml_text(text):
                    raise ValueError("paragraph_text_invalid")
                final_text = text

            if style is not None:
                if not isinstance(style, str) or not style.strip():
                    raise ValueError("paragraph_style_invalid")
                final_style = style.strip()
            else:
                final_style = None

        object.__setattr__(self, "source", source_path)
        object.__setattr__(self, "destination", dest_path)
        object.__setattr__(self, "source_sha256", valid_source_sha)
        object.__setattr__(self, "target", {"part": target_part, "paragraph": target_p})
        object.__setattr__(self, "structure_fingerprint_sha256", valid_fp)
        object.__setattr__(self, "action", action)
        object.__setattr__(self, "position", final_position)
        object.__setattr__(self, "text", final_text)
        object.__setattr__(self, "style", final_style)
        object.__setattr__(self, "offset", final_offset)

    @property
    def paragraph(self) -> int:
        return self.target["paragraph"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "part": self.part,
            "paragraph": self.paragraph,
        }
        if self.action == "insert" and self.position is not None:
            data["position"] = self.position
        elif self.action == "split" and self.offset is not None:
            data["offset"] = self.offset
        return data

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "target": {
                "part": self.part,
                "paragraph": self.paragraph,
            },
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "action": self.action,
        }
        if self.action == "insert":
            if self.position is not None:
                data["position"] = self.position
            if self.text is not None:
                data["text"] = self.text
            if self.style is not None:
                data["style"] = self.style
        elif self.action == "split":
            if self.offset is not None:
                data["offset"] = self.offset
        return data


@dataclass(frozen=True)
class ParagraphReport:
    """Content-safe public result for a body paragraph authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    action: str | None = None
    paragraph: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status in (_PUBLISHED, _APPLIED)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "schema_version": 1,
            "operation": "paragraph-authoring",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
        }
        if self.target is not None:
            result["target"] = self.target
        if self.action is not None:
            result["action"] = self.action
        if self.paragraph is not None:
            result["paragraph"] = self.paragraph
        if self.target_change_report is not None:
            result["target_change_report"] = self.target_change_report
        if self.non_target_parts_preserved is not None:
            result["non_target_parts_preserved"] = self.non_target_parts_preserved
        return result


def _paragraph_refusal(
    request: ParagraphRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> ParagraphReport:
    return ParagraphReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        action=request.action,
        paragraph=None,
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_paragraph_request(value: Any) -> ParagraphRequest:
    """Parse and strictly validate a paragraph authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("paragraph_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("paragraph_request_fields_invalid")

    action = value.get("action", "insert")
    if not isinstance(action, str) or action not in ("insert", "delete", "split"):
        raise ValueError("paragraph_action_invalid")

    if action == "delete" and not keys.issubset(_ALLOWED_DELETE_REQUEST_FIELDS):
        raise ValueError("paragraph_request_fields_invalid")
    if action == "insert" and not keys.issubset(_ALLOWED_INSERT_REQUEST_FIELDS):
        raise ValueError("paragraph_request_fields_invalid")
    if action == "split" and not keys.issubset(_ALLOWED_SPLIT_REQUEST_FIELDS):
        raise ValueError("paragraph_request_fields_invalid")

    target = value["target"]
    if not isinstance(target, Mapping):
        raise TypeError("paragraph_target_invalid")
    target_keys = set(target)
    if not target_keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(
        target_keys
    ):
        raise ValueError("paragraph_target_fields_invalid")

    req_target = {
        "paragraph": target["paragraph"],
        "part": target.get("part", _DOCUMENT),
    }
    if "offset" in target:
        req_target["offset"] = target["offset"]
    return ParagraphRequest(
        source=Path(value["source"]),
        destination=Path(value["destination"]),
        source_sha256=value["source_sha256"],
        target=req_target,
        structure_fingerprint_sha256=value["structure_fingerprint_sha256"],
        action=action,
        position=value.get("position"),
        text=value.get("text"),
        style=value.get("style"),
        offset=value.get("offset"),
    )


def load_paragraph_request(path: str | Path) -> ParagraphRequest:
    """Load and strictly validate one JSON paragraph request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_paragraph_request(json.load(stream))


def paragraph_authoring_structure_fingerprint(
    source: str | Path,
    *,
    paragraph: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural anchor fingerprint for one exact body or header/footer paragraph."""
    if paragraph < 1:
        return None
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        root, _, ns, _, err = _resolve_target_story(graph, part)
        if err is not None or root is None or ns is None:
            return None
        paragraphs = list(root.iter(Q(ns, "p")))
        if paragraph < 1 or paragraph > len(paragraphs):
            return None
        target_p = paragraphs[paragraph - 1]
        return _cell_structure_fingerprint(target_p)
    except (
        OSError,
        PackageError,
        ET.ParseError,
        KeyError,
        IndexError,
        ValueError,
        zipfile.BadZipFile,
    ):
        return None


def discover_paragraph_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate body or header/footer paragraphs for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "paragraph-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "paragraph_source_missing",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    try:
        source_sha256 = _sha256_file(path)
        graph = PackageGraph.from_file(path)
    except (PackageError, OSError, zipfile.BadZipFile):
        return {
            "schema_version": 1,
            "operation": "paragraph-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "paragraph_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    root, container, ns, _, err = _resolve_target_story(graph, part)
    if err is not None:
        return {
            "schema_version": 1,
            "operation": "paragraph-targets",
            "passed": False,
            "status": _REFUSED,
            "code": err,
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    assert root is not None and ns is not None

    if container is None:
        return {
            "schema_version": 1,
            "operation": "paragraph-targets",
            "passed": True,
            "status": _DISCOVERED,
            "code": None,
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    parent_map = {c: p for p in root.iter() for c in p}
    unsupported_context = _unsupported_context_by_element(root, ns)
    paragraph_hazards = _index_paragraph_hazards(container, ns)

    targets: list[dict[str, Any]] = []
    for paragraph_index, paragraph in enumerate(root.iter(Q(ns, "p")), start=1):
        if not _is_admitted_paragraph(paragraph, container, parent_map, ns):
            continue

        in_unsupported_ctx = unsupported_context.get(id(paragraph), False)
        hazard = (
            paragraph_hazards.get(id(paragraph))
            or _paragraph_hazard_reason(paragraph, ns)
            or _cell_ancestor_revision_hazard(paragraph, container, parent_map, ns)
        )
        is_supported = not in_unsupported_ctx and hazard is None
        unsupported_reason = "paragraph_context_unsupported" if in_unsupported_ctx else hazard

        text = _visible_text(paragraph, ns)
        fp = _cell_structure_fingerprint(paragraph)
        style = _paragraph_style(paragraph, ns)

        target = ParagraphTarget(
            part=part,
            paragraph=paragraph_index,
            text=text,
            structure_fingerprint_sha256=fp,
            style=style,
            empty=len(text.strip()) == 0,
            supported=is_supported,
            unsupported_reason=unsupported_reason,
        )
        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "paragraph-targets",
        "passed": True,
        "status": _DISCOVERED,
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


@dataclass(frozen=True)
class _ParagraphPlan:
    """One validated paragraph edit: the rewritten story part and its re-check."""

    serialized: bytes
    change_report: dict[str, Any]
    verify_story: Callable[[ET.Element], bool]


def _plan_paragraph_edit(
    request: ParagraphRequest,
    graph: PackageGraph,
    load: Callable[[str], tuple[bytes, ET.Element]] | None = None,
) -> _ParagraphPlan | str:
    """Validate one paragraph request against its story part and plan the rewrite.

    Returns a refusal code, or the rewritten part bytes, the change report, and the
    structural check a re-parsed copy of that part must pass. ``load`` returns the
    current bytes of a part and an editable tree parsed from them (by default both
    come from ``graph``); the tree is mutated, so it must be private to this call.
    """
    root, container, ns, part_bytes, err = _resolve_target_story(graph, request.part, load)
    if err is not None:
        return err

    assert root is not None and ns is not None and part_bytes is not None

    if container is None:
        return "paragraph_target_nested"

    paragraphs = list(root.iter(Q(ns, "p")))
    if request.paragraph < 1 or request.paragraph > len(paragraphs):
        return "paragraph_target_missing"

    target_p = paragraphs[request.paragraph - 1]

    admission = _ParagraphAdmission(root, container, ns, target_p)
    try:
        span = _paragraph_xml_span(part_bytes, ns, request.paragraph, admission=admission)
    except (expat.ExpatError, ValueError, LookupError):
        return "paragraph_staged_verification_failed"
    parent_map = admission.parent_map
    if not _is_admitted_paragraph(target_p, container, parent_map, ns):
        return "paragraph_target_nested"

    if admission.unsupported_context:
        return "paragraph_context_unsupported"

    hazard = (
        admission.hazards.get(id(target_p))
        or _paragraph_hazard_reason(target_p, ns)
        or _cell_ancestor_revision_hazard(target_p, container, parent_map, ns)
    )
    if hazard is not None:
        return hazard

    fp = _cell_structure_fingerprint(target_p)
    if fp != request.structure_fingerprint_sha256:
        return "paragraph_structure_fingerprint_mismatch"

    parent_elem = parent_map[target_p]
    original_parent_children = list(parent_elem)
    anchor_idx = original_parent_children.index(target_p)
    is_cell_target = parent_elem is not container
    all_tcs = list(root.iter(Q(ns, "tc"))) if is_cell_target else []
    tc_idx = all_tcs.index(parent_elem) if is_cell_target else -1

    if request.action == "insert":
        text = request.text or ""
        if len(text) > _MAX_TEXT_CHARACTERS:
            return "paragraph_text_limit_exceeded"
        if not _is_valid_xml_text(text):
            return "paragraph_text_invalid"

        # w:pStyle names the w:styleId; a display name resolves to it as in python-docx.
        style_id: str | None = None
        if request.style is not None:
            if _STYLES not in graph.parts:
                return "paragraph_style_missing"
            paragraph_styles, _, _ = _catalog_package_styles(graph, ns)
            style_id = paragraph_styles.get(request.style)
            if style_id is None:
                return "paragraph_style_missing"

        new_p = _build_paragraph_element(request.text, style_id, ns)

        if request.position == "before":
            parent_elem.insert(anchor_idx, new_p)
            expected_p_idx = anchor_idx
        else:
            parent_elem.insert(anchor_idx + 1, new_p)
            expected_p_idx = anchor_idx + 1

        new_paragraphs = list(root.iter(Q(ns, "p")))
        inserted_p_index = new_paragraphs.index(new_p) + 1

        text_str = request.text or ""
        target_change_report = {
            "action": "insert",
            "position": request.position,
            "anchor_paragraph": request.paragraph,
            "inserted_paragraph_index": inserted_p_index,
            "text_characters": len(text_str),
            "text_sha256": hashlib.sha256(text_str.encode("utf-8")).hexdigest(),
            "style": style_id,
        }

    elif request.action == "split":
        if request.offset is None or request.offset < 0:
            return "paragraph_offset_invalid"

        ppr = target_p.find(Q(ns, "pPr"))
        if ppr is not None and ppr.find(Q(ns, "sectPr")) is not None:
            return "paragraph_section_property_unsupported"

        if target_p.findall(f".//{Q(ns, 'bookmarkStart')}") or target_p.findall(
            f".//{Q(ns, 'bookmarkEnd')}"
        ):
            return "paragraph_bookmark_unsupported"

        if (
            target_p.findall(f".//{Q(ns, 'fldSimple')}")
            or target_p.findall(f".//{Q(ns, 'instrText')}")
            or target_p.findall(f".//{Q(ns, 'fldChar')}")
        ):
            return "paragraph_field_unsupported"

        if (
            target_p.findall(f".//{Q(ns, 'ins')}")
            or target_p.findall(f".//{Q(ns, 'del')}")
            or target_p.findall(f".//{Q(ns, 'moveFrom')}")
            or target_p.findall(f".//{Q(ns, 'moveTo')}")
            or target_p.findall(f".//{Q(ns, 'rPrChange')}")
            or target_p.findall(f".//{Q(ns, 'pPrChange')}")
            or target_p.findall(f".//{Q(ns, 'sectPrChange')}")
        ):
            return "paragraph_revision_unsupported"

        if target_p.findall(f".//{Q(ns, 'hyperlink')}"):
            return "paragraph_hyperlink_unsupported"

        for child in target_p:
            if child.tag not in {Q(ns, "pPr"), Q(ns, "r")}:
                return "paragraph_complex_unsupported"
            if child.tag == Q(ns, "r"):
                for r_child in child:
                    r_local = r_child.tag.split("}")[-1] if "}" in r_child.tag else r_child.tag
                    if r_local in ("rPr", "t") or r_local in _SUPPORTED_RUN_CONTROLS:
                        continue
                    if r_local in ("drawing", "pict", "wgp"):
                        return "paragraph_drawing_unsupported"
                    if r_local in ("fldSimple", "instrText", "fldChar"):
                        return "paragraph_field_unsupported"
                    if r_local in (
                        "ins",
                        "del",
                        "moveFrom",
                        "moveTo",
                        "rPrChange",
                        "pPrChange",
                        "sectPrChange",
                    ):
                        return "paragraph_revision_unsupported"
                    if r_local in (
                        "commentRangeStart",
                        "commentRangeEnd",
                        "commentReference",
                    ):
                        return "paragraph_comment_unsupported"
                    if r_local in ("bookmarkStart", "bookmarkEnd"):
                        return "paragraph_bookmark_unsupported"
                    return "paragraph_complex_unsupported"

        target_text = _visible_text(target_p, ns)
        if request.offset > len(target_text):
            return "paragraph_offset_invalid"

        before_text = target_text[: request.offset]
        after_text = target_text[request.offset :]
        original_style = _paragraph_style(target_p, ns)
        split_offset = request.offset
        total_text_len = len(target_text)

        p1_runs: list[ET.Element] = []
        p2_runs: list[ET.Element] = []
        cur_offset = 0

        for r_child in target_p:
            if r_child.tag != Q(ns, "r"):
                continue

            rpr = r_child.find(Q(ns, "rPr"))
            left_children: list[ET.Element] = []
            right_children: list[ET.Element] = []

            for elem in r_child:
                e_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if e_local == "rPr":
                    continue

                if e_local == "t":
                    t_text = elem.text or ""
                    c_start = cur_offset
                    c_end = cur_offset + len(t_text)
                    cur_offset = c_end

                    if c_end <= split_offset:
                        left_children.append(copy.deepcopy(elem))
                    elif c_start >= split_offset:
                        right_children.append(copy.deepcopy(elem))
                    else:
                        split_idx = split_offset - c_start
                        left_text = t_text[:split_idx]
                        right_text = t_text[split_idx:]

                        left_t = ET.Element(Q(ns, "t"))
                        left_t.text = left_text
                        if (
                            left_text[:1].isspace()
                            or left_text[-1:].isspace()
                            or elem.get(_XML_SPACE_ATTRIBUTE) == "preserve"
                        ):
                            left_t.set(_XML_SPACE_ATTRIBUTE, "preserve")
                        left_children.append(left_t)

                        right_t = ET.Element(Q(ns, "t"))
                        right_t.text = right_text
                        if (
                            right_text[:1].isspace()
                            or right_text[-1:].isspace()
                            or elem.get(_XML_SPACE_ATTRIBUTE) == "preserve"
                        ):
                            right_t.set(_XML_SPACE_ATTRIBUTE, "preserve")
                        right_children.append(right_t)
                else:
                    # Supported control element: tab, br, cr, noBreakHyphen, softHyphen
                    pos = cur_offset
                    if pos < split_offset or (
                        pos == split_offset and split_offset == total_text_len and split_offset > 0
                    ):
                        left_children.append(copy.deepcopy(elem))
                    else:
                        right_children.append(copy.deepcopy(elem))

            if left_children:
                left_r = ET.Element(Q(ns, "r"))
                if rpr is not None:
                    left_r.append(copy.deepcopy(rpr))
                for child_elem in left_children:
                    left_r.append(child_elem)
                p1_runs.append(left_r)

            if right_children:
                right_r = ET.Element(Q(ns, "r"))
                if rpr is not None:
                    right_r.append(copy.deepcopy(rpr))
                for child_elem in right_children:
                    right_r.append(child_elem)
                p2_runs.append(right_r)

            if not left_children and not right_children and rpr is not None and len(r_child) == 1:
                empty_r = copy.deepcopy(r_child)
                if cur_offset < split_offset:
                    p1_runs.append(empty_r)
                else:
                    p2_runs.append(empty_r)

        for child in list(target_p):
            if child.tag != Q(ns, "pPr"):
                target_p.remove(child)
        for r in p1_runs:
            target_p.append(r)

        new_p = ET.Element(Q(ns, "p"))
        if ppr is not None:
            new_p.append(copy.deepcopy(ppr))
        for r in p2_runs:
            new_p.append(r)

        parent_elem.insert(anchor_idx + 1, new_p)
        new_p.tail = target_p.tail
        target_p.tail = None

        target_change_report = {
            "action": "split",
            "anchor_paragraph": request.paragraph,
            "split_paragraph_index": request.paragraph,
            "next_paragraph_index": request.paragraph + 1,
            "offset": request.offset,
            "split_offset": request.offset,
            "before_text_characters": len(before_text),
            "before_text_sha256": hashlib.sha256(before_text.encode("utf-8")).hexdigest(),
            "after_text_characters": len(after_text),
            "after_text_sha256": hashlib.sha256(after_text.encode("utf-8")).hexdigest(),
            "style": original_style,
        }

    else:  # delete
        ppr = target_p.find(Q(ns, "pPr"))
        if ppr is not None and ppr.find(Q(ns, "sectPr")) is not None:
            return "paragraph_section_boundary_deletion"

        if not is_cell_target:
            container_paragraphs = [c for c in container if c.tag == Q(ns, "p")]
            if len(container_paragraphs) <= 1:
                return "paragraph_final_paragraph_deletion"
            if request.part == _DOCUMENT and target_p is container_paragraphs[-1]:
                return "paragraph_final_paragraph_deletion"
        else:
            cell_paragraphs = [c for c in parent_elem if c.tag == Q(ns, "p")]
            if len(cell_paragraphs) <= 1:
                return "paragraph_final_paragraph_deletion"
            remaining_children = [c for c in parent_elem if c is not target_p]
            if remaining_children and remaining_children[-1].tag != Q(ns, "p"):
                return "paragraph_final_paragraph_deletion"

        if target_p.findall(f".//{Q(ns, 'bookmarkStart')}") or target_p.findall(
            f".//{Q(ns, 'bookmarkEnd')}"
        ):
            return "paragraph_bookmark_unsupported"

        deleted_text = _visible_text(target_p, ns)
        deleted_style = _paragraph_style(target_p, ns)

        # Removing only the element leaves its lexical tail in the parent.
        if target_p.tail:
            if anchor_idx:
                previous = original_parent_children[anchor_idx - 1]
                previous.tail = (previous.tail or "") + target_p.tail
            else:
                parent_elem.text = (parent_elem.text or "") + target_p.tail
        parent_elem.remove(target_p)
        expected_p_idx = -1

        del_text = deleted_text or ""
        target_change_report = {
            "action": "delete",
            "deleted_paragraph_index": request.paragraph,
            "deleted_text_characters": len(del_text),
            "deleted_text_sha256": hashlib.sha256(del_text.encode("utf-8")).hexdigest(),
            "deleted_style": deleted_style,
        }

    def verify_story(s_root: ET.Element) -> bool:
        """True when a re-parsed copy of the rewritten part holds exactly this edit."""
        try:
            s_ns = _word_namespace(s_root)
            if s_ns != ns:
                return False
            if request.part == _DOCUMENT:
                s_container = s_root.find(Q(s_ns, "body"))
            else:
                s_container = s_root
            if s_container is None:
                return False

            if not is_cell_target:
                s_parent = s_container
            else:
                s_all_tcs = list(s_root.iter(Q(s_ns, "tc")))
                if len(s_all_tcs) != len(all_tcs):
                    return False
                s_parent = s_all_tcs[tc_idx]

            s_children = list(s_parent)

            if request.action == "insert":
                if len(s_children) != len(original_parent_children) + 1:
                    return False

                if s_children[expected_p_idx].tag != Q(s_ns, "p"):
                    return False

                if request.position == "after":
                    anchor_node = s_children[expected_p_idx - 1]
                    if anchor_node.tag != Q(s_ns, "p"):
                        return False
                    if (
                        _cell_structure_fingerprint(anchor_node)
                        != request.structure_fingerprint_sha256
                    ):
                        return False
                else:
                    anchor_node = s_children[expected_p_idx + 1]
                    if anchor_node.tag != Q(s_ns, "p"):
                        return False
                    if (
                        _cell_structure_fingerprint(anchor_node)
                        != request.structure_fingerprint_sha256
                    ):
                        return False

                for orig_idx, orig_elem in enumerate(original_parent_children):
                    s_idx = orig_idx if orig_idx < expected_p_idx else orig_idx + 1
                    s_elem = s_children[s_idx]
                    if s_elem.tag != orig_elem.tag or not _xml_trees_equal(s_elem, orig_elem):
                        return False

                ins_p = s_children[expected_p_idx]
                actual_text = "".join(t.text or "" for t in ins_p.iter(Q(s_ns, "t")))
                if actual_text != (request.text or ""):
                    return False

                actual_style = _paragraph_style(ins_p, s_ns)
                if actual_style != style_id:
                    return False

            elif request.action == "delete":
                if len(s_children) != len(original_parent_children) - 1:
                    return False

                for orig_idx, orig_elem in enumerate(original_parent_children):
                    if orig_idx == anchor_idx:
                        continue
                    s_idx = orig_idx if orig_idx < anchor_idx else orig_idx - 1
                    s_elem = s_children[s_idx]
                    if s_elem.tag != orig_elem.tag or not _xml_trees_equal(s_elem, orig_elem):
                        return False

                s_paragraphs = [c for c in s_children if c.tag == Q(s_ns, "p")]
                if len(s_paragraphs) < 1:
                    return False

                if is_cell_target and s_children[-1].tag != Q(s_ns, "p"):
                    return False

            elif request.action == "split":
                if len(s_children) != len(original_parent_children) + 1:
                    return False

                p1_node = s_children[anchor_idx]
                p2_node = s_children[anchor_idx + 1]
                if p1_node.tag != Q(s_ns, "p") or p2_node.tag != Q(s_ns, "p"):
                    return False

                p1_actual_text = "".join(t.text or "" for t in p1_node.iter(Q(s_ns, "t")))
                p2_actual_text = "".join(t.text or "" for t in p2_node.iter(Q(s_ns, "t")))
                if p1_actual_text != before_text or p2_actual_text != after_text:
                    return False
                if p1_actual_text + p2_actual_text != target_text:
                    return False

                if _paragraph_style(p1_node, s_ns) != original_style:
                    return False
                if _paragraph_style(p2_node, s_ns) != original_style:
                    return False

                for orig_idx, orig_elem in enumerate(original_parent_children):
                    if orig_idx == anchor_idx:
                        continue
                    s_idx = orig_idx if orig_idx < anchor_idx else orig_idx + 1
                    s_elem = s_children[s_idx]
                    if s_elem.tag != orig_elem.tag or not _xml_trees_equal(s_elem, orig_elem):
                        return False

                if is_cell_target and s_children[-1].tag != Q(s_ns, "p"):
                    return False
            if is_cell_target:
                for idx, orig_tc in enumerate(all_tcs):
                    if idx == tc_idx:
                        continue
                    if not _xml_trees_equal(s_all_tcs[idx], orig_tc):
                        return False
        except (
            OSError,
            PackageError,
            ET.ParseError,
            KeyError,
            IndexError,
            ValueError,
            TypeError,
        ):
            return False
        return True

    try:
        start, end, after, encoding = span
        if request.action == "delete":
            payload = b""
        elif request.action == "insert":
            start = end = start if request.position == "before" else after
            payload = _paragraph_xml_fragment(new_p, encoding)
        else:
            payload = _paragraph_xml_fragment(target_p, encoding) + _paragraph_xml_fragment(
                new_p, encoding
            )
        serialized = part_bytes[:start] + payload + part_bytes[end:]
    except (expat.ExpatError, ValueError, LookupError):
        return "paragraph_staged_verification_failed"
    return _ParagraphPlan(serialized, target_change_report, verify_story)


def author_paragraph(request: ParagraphRequest, *, verify: str = "full") -> ParagraphReport:
    """Insert or delete a body or referenced header/footer paragraph atomically and verification-checked.

    ``verify="full"`` (the default) inspects the whole staged package and proves every
    other part byte-identical to the source before publishing. ``verify="changed-parts"``
    keeps every request check and the source-hash re-checks, validates the source
    package once, proves every other staged part (relationships included) hashes
    identically to it, re-parses only the edited story part under the package limits
    and re-runs the same structural checks on it, skipping re-inspection of unchanged
    parts.
    """
    if not isinstance(request, ParagraphRequest):
        raise TypeError("paragraph_request_invalid")
    _require_verify_mode(verify)

    source = request.source
    destination = request.destination

    if not source.is_file():
        return _paragraph_refusal(request, "paragraph_source_missing")

    if destination.exists() or destination.is_symlink():
        return _paragraph_refusal(request, "paragraph_destination_occupied")

    try:
        source_sha256 = _sha256_file(source)
    except OSError:
        return _paragraph_refusal(request, "paragraph_source_missing")

    if source_sha256 != request.source_sha256:
        return _paragraph_refusal(request, "paragraph_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _paragraph_refusal(request, "paragraph_package_invalid")

    plan = _plan_paragraph_edit(request, graph)
    if isinstance(plan, str):
        return _paragraph_refusal(request, plan)
    serialized = plan.serialized
    target_change_report = plan.change_report

    def verify_publication(staged: _StagedPackage) -> str | None:
        # The shared gate already validated the staged package (whole, or its changed
        # parts against the validated source) and parsed the story part from it.
        if verify == "full" and not _unrelated_parts_preserved(
            source, staged.path, frozenset({request.part})
        ):
            return "paragraph_staged_verification_failed"
        try:
            if staged.part_sha256.get(request.part) != hashlib.sha256(serialized).hexdigest():
                return "paragraph_staged_verification_failed"
            s_root = staged.parsed_xml.get(request.part)
            if s_root is None or not plan.verify_story(s_root):
                return "paragraph_staged_verification_failed"
            if "word/_rels/document.xml.rels" in graph.parts and (
                staged.part_sha256.get("word/_rels/document.xml.rels")
                != graph.parts["word/_rels/document.xml.rels"].sha256
            ):
                return "paragraph_staged_verification_failed"

        except (
            OSError,
            PackageError,
            ET.ParseError,
            KeyError,
            IndexError,
            ValueError,
            TypeError,
        ):
            return "paragraph_staged_verification_failed"

        return None

    replacements = {request.part: serialized}

    baseline: _Inspection | None
    if verify == "changed-parts":
        baseline = _inspect_docx(source)
        if not baseline.valid or baseline.file_sha256 != request.source_sha256:
            return _paragraph_refusal(request, "paragraph_staged_verification_failed")
    else:
        # A remembered inspection of these exact source bytes only lends parsed trees
        # to byte-identical staged parts; the staged package is still inspected whole.
        baseline = _remembered_inspection(source, source_sha256)

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=[target_change_report],
        verify_staged=verify_publication,
        baseline=baseline,
        verify_mode=verify,
        failure_code="paragraph_staged_verification_failed",
        exists_code="paragraph_destination_occupied",
        unavailable_code="paragraph_publication_unavailable",
    )

    if not published.passed:
        return _paragraph_refusal(request, published.code or "paragraph_staged_verification_failed")

    return ParagraphReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        action=request.action,
        paragraph=target_change_report,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
