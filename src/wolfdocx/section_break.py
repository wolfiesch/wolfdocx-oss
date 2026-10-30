"""Source-bound DOCX body section break authoring and discovery.

Provides verified, source-bound insertion and deletion of section breaks in
WordprocessingML document bodies, strictly preserving existing section formatting,
headers, footers, and unrelated document parts.

OOXML section break mechanics (ECMA-376 §17.6.22 w:type, §17.6.17 w:sectPr):
    In WordprocessingML, section properties (w:sectPr) define the layout of the section
    preceding them. For all sections except the final one, w:sectPr is stored inside
    w:pPr of the section's terminating paragraph (<w:p><w:pPr><w:sectPr>...</w:sectPr></w:pPr>...).
    For the final section, w:sectPr is stored as a direct child of <w:body> at the end.

    Insertion lifecycle (action="insert"):
    1. The containing section is resolved as the first terminating w:sectPr at or after
       the anchor item in document order.
    2. The inserted paragraph (<w:p><w:pPr>{cloned_sectPr}</w:pPr></w:p>) receives a
       verbatim deep copy of the containing section's terminating w:sectPr, preserving
       its original w:type, page geometry, headers, footers, and column layout, with
       root rsid attributes cleared.
    3. The containing section's original terminating w:sectPr has its w:type set to the
       requested break_type.

    Deletion lifecycle (action="delete"):
    1. The target paragraph containing the intermediate section break is located by
       paragraph/insert_after index and verified by its structural SHA-256 fingerprint.
    2. Only the terminating w:sectPr child of the paragraph's w:pPr is removed.
       The paragraph itself, its runs, text, and paragraph properties remain intact.
    3. The preceding section merges into the succeeding section, which preserves its
       section geometry (w:pgSz, w:pgMar), break type, and header/footer references.
    4. The document-final body-level w:sectPr is protected and cannot be deleted.
    5. Unsupported inherited-header, conflicting header/footer, and complex contexts
       are explicitly refused to prevent silent content loss.

Break type vocabulary (ECMA-376 §17.18.77 ST_SectionMark):
    Requests accept nextPage, nextColumn, continuous, evenPage, and oddPage, plus
    python-docx ``WD_SECTION`` members and their names in any case (``NEW_COLUMN``,
    ``new_page``). Every accepted spelling is normalized to its ST_SectionMark value
    before authoring, so written markup, requests, and reports carry canonical values.
"""

from __future__ import annotations

import copy
import hashlib
import json
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .assembly import _new_rel_id, _rels_name, _resolve
from .authoring import (
    _PUBLISHED,
    _REFUSED,
    _require_sha256,
    _sha256_file,
    _unsupported_context_by_element,
)
from .inspect import (
    STRICT_WORD_NS,
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import _DOCUMENT, ReviewReport, _publish, _word_namespace

_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
_FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
__all__ = [
    "SectionBreakReport",
    "SectionBreakRequest",
    "SectionBreakTarget",
    "author_section_break",
    "discover_section_break_targets",
    "load_section_break_request",
    "parse_section_break_request",
    "section_break_anchor_fingerprint",
]

_ALLOWED_BREAK_TYPES = frozenset({"nextPage", "nextColumn", "continuous", "evenPage", "oddPage"})
# Lowercased python-docx WD_SECTION member names and snake_case ST_SectionMark spellings.
_BREAK_TYPE_ALIASES: dict[str, str] = {
    "new_page": "nextPage",
    "next_page": "nextPage",
    "new_column": "nextColumn",
    "next_column": "nextColumn",
    "continuous": "continuous",
    "even_page": "evenPage",
    "odd_page": "oddPage",
}
_ALLOWED_ACTIONS = frozenset({"insert", "delete", "update"})


def _canonical_break_type(value: object) -> str | None:
    """Return the ST_SectionMark value for a requested break type, or None if unknown.

    Accepts canonical values, python-docx ``WD_SECTION`` members (matched by member
    name), and those names or snake_case spellings in any case.
    """
    if isinstance(value, Enum):
        value = value.name
    if not isinstance(value, str):
        return None
    if value in _ALLOWED_BREAK_TYPES:
        return value
    return _BREAK_TYPE_ALIASES.get(value.lower())


_REQUIRED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "anchor_fingerprint_sha256",
    }
)

_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "insert_after",
        "paragraph",
        "target",
        "anchor_fingerprint_sha256",
        "break_type",
        "part",
        "action",
        "link_to_previous",
    }
)

# OOXML ECMA-376 17.6.17 CT_SectPr child schema ordering
_SECT_PR_ORDER: dict[str, int] = {
    "headerReference": 0,
    "footerReference": 0,
    "footnotePr": 1,
    "endnotePr": 2,
    "type": 3,
    "pgSz": 4,
    "pgMar": 5,
    "paperSrc": 6,
    "pgBorders": 7,
    "lnNumType": 8,
    "pgNumType": 9,
    "cols": 10,
    "formProt": 11,
    "vAlign": 12,
    "noEndnote": 13,
    "titlePg": 14,
    "textDirection": 15,
    "bidi": 16,
    "rtlGutter": 17,
    "docGrid": 18,
    "printerSettings": 19,
    "sectPrChange": 20,
}


def _unrelated_parts_preserved(
    source: Path,
    candidate: Path,
    changed: frozenset[str],
) -> bool:
    """True when every part the request did not change is byte-identical to source."""
    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        names = set(after.namelist())
        return all(
            name in names and before.read(name) == after.read(name)
            for name in before.namelist()
            if name not in changed
        )


def _xml_hash(element: ET.Element) -> str:
    """Return canonical SHA-256 digest of element XML serialization."""
    return hashlib.sha256(ET.tostring(element, encoding="utf-8")).hexdigest()


def _body_items(root: ET.Element, ns: str) -> list[ET.Element]:
    """Return top-level body items (paragraphs and tables) in document order."""
    body = root.find(Q(ns, "body"))
    if body is None:
        return []
    return [node for node in body if node.tag in {Q(ns, "p"), Q(ns, "tbl")}]


def _sect_pr_rank(elem: ET.Element) -> int:
    tag = elem.tag.rsplit("}", 1)[-1] if "}" in elem.tag else elem.tag
    return _SECT_PR_ORDER.get(tag, 99)


def _insert_sect_pr_child(sect_pr: ET.Element, new_child: ET.Element) -> None:
    """Insert a new child element into sect_pr maintaining OOXML schema order."""
    target_rank = _sect_pr_rank(new_child)
    inserted = False
    for i, child in enumerate(sect_pr):
        child_rank = _sect_pr_rank(child)
        if child_rank > target_rank:
            sect_pr.insert(i, new_child)
            inserted = True
            break
    if not inserted:
        sect_pr.append(new_child)


def _find_sections(root: ET.Element, ns: str) -> list[ET.Element]:
    """Find all section property elements in the document body in order."""
    body = root.find(Q(ns, "body"))
    if body is None:
        return []
    sections: list[ET.Element] = []
    for child in body:
        if child.tag == Q(ns, "p"):
            ppr = child.find(Q(ns, "pPr"))
            if ppr is not None:
                sect_pr = ppr.find(Q(ns, "sectPr"))
                if sect_pr is not None:
                    sections.append(sect_pr)
    final = body.find(Q(ns, "sectPr"))
    if final is not None:
        sections.append(final)
    return sections


def _find_containing_sect_pr(
    body: ET.Element,
    anchor: ET.Element,
    ns: str,
) -> ET.Element | None:
    """Resolve containing section as first terminating sectPr strictly after anchor."""
    found_anchor = False
    for child in body:
        if child is anchor:
            found_anchor = True
            continue
        if not found_anchor:
            continue
        if child.tag == Q(ns, "p"):
            ppr = child.find(Q(ns, "pPr"))
            if ppr is not None:
                sect_pr = ppr.find(Q(ns, "sectPr"))
                if sect_pr is not None:
                    return sect_pr
        elif child.tag == Q(ns, "sectPr"):
            return child
    return None


def _is_unsupported_section(
    sect_pr: ET.Element,
    unsupported_by_id: dict[int, bool],
    ns: str,
) -> bool:
    """True if section or any descendant has unsupported context or revision wrappers."""
    if unsupported_by_id.get(id(sect_pr), False):
        return True
    unsupported_tags = {
        Q(ns, "sectPrChange"),
        Q(ns, "ins"),
        Q(ns, "del"),
        Q(ns, "moveFrom"),
        Q(ns, "moveTo"),
        Q(ns, "sdt"),
        Q(ns, "customXml"),
    }
    for child in sect_pr.iter():
        if child.tag in unsupported_tags or unsupported_by_id.get(id(child), False):
            return True
    return False


def _clear_root_rsid_attrs(elem: ET.Element) -> None:
    """Remove root revision tracking attributes (rsid*) from element."""
    for attr in list(elem.attrib.keys()):
        local = attr.rsplit("}", 1)[-1] if "}" in attr else attr
        if local.lower().startswith("rsid"):
            del elem.attrib[attr]


def _is_title_pg_enabled(sect_pr: ET.Element | None, ns: str) -> bool:
    """Return True if first-page header/footer behavior (w:titlePg) is effectively enabled."""
    if sect_pr is None:
        return False
    elem = sect_pr.find(Q(ns, "titlePg"))
    if elem is None:
        return False
    val = elem.get(Q(ns, "val"))
    if val is None:
        val = elem.get("val")
    if val is None:
        return True
    return str(val).strip().lower() not in {"0", "false", "off"}


def _apply_link_to_previous(
    sect_pr: ET.Element,
    predecessors: list[ET.Element],
    link_to_previous: bool,
    graph: PackageGraph,
    rels_root: ET.Element,
    types_root: ET.Element,
    additions: dict[str, bytes],
    ns: str,
) -> None:
    """Apply link_to_previous behavior to sect_pr relative to predecessors."""
    if link_to_previous is True:
        for child in list(sect_pr):
            tag = child.tag.rsplit("}", 1)[-1] if "}" in child.tag else child.tag
            if tag in {"headerReference", "footerReference"}:
                sect_pr.remove(child)
    elif link_to_previous is False:
        relationships_ns = _STRICT_R if ns == STRICT_WORD_NS else _R
        active_stories: dict[tuple[str, str], str] = {}
        for pred in predecessors:
            for child in pred:
                tag = child.tag.rsplit("}", 1)[-1] if "}" in child.tag else child.tag
                if tag in {"headerReference", "footerReference"}:
                    v = child.get(Q(ns, "type")) or "default"
                    rid = child.get(Q(relationships_ns, "id")) or child.get("id") or ""
                    if rid:
                        rel = None
                        for r in rels_root.findall(Q(_PR, "Relationship")):
                            if r.get("Id") == rid:
                                rel = r
                                break
                        if rel is not None:
                            tgt = rel.get("Target", "").strip()
                            part_name = _resolve(_DOCUMENT, tgt)
                            if part_name in graph.parts or part_name in additions:
                                active_stories[(tag, v)] = part_name

        for (tag, v), src_part in sorted(active_stories.items()):
            kind = "header" if tag == "headerReference" else "footer"
            ct = _HEADER_CONTENT_TYPE if kind == "header" else _FOOTER_CONTENT_TYPE
            src_bytes = additions.get(src_part) or graph.read_part_bytes(src_part)
            prefix = f"word/{kind}"
            serial = 1
            while f"{prefix}{serial}.xml" in graph.parts or f"{prefix}{serial}.xml" in additions:
                serial += 1
            new_part = f"{prefix}{serial}.xml"
            new_rel_target = f"{kind}{serial}.xml"
            new_rel_type = f"{relationships_ns}/{kind}"

            additions[new_part] = src_bytes
            new_rel_id = _new_rel_id(rels_root)
            rels_root.append(
                ET.Element(
                    Q(_PR, "Relationship"),
                    {
                        "Id": new_rel_id,
                        "Type": new_rel_type,
                        "Target": new_rel_target,
                    },
                )
            )
            types_root.append(
                ET.Element(
                    Q(_CT, "Override"),
                    {
                        "PartName": f"/{new_part}",
                        "ContentType": ct,
                    },
                )
            )

            found = None
            for child in sect_pr:
                c_tag = child.tag.rsplit("}", 1)[-1] if "}" in child.tag else child.tag
                if c_tag == tag and (child.get(Q(ns, "type")) or "default") == v:
                    found = child
                    break
            if found is not None:
                found.set(Q(relationships_ns, "id"), new_rel_id)
                if "id" in found.attrib:
                    del found.attrib["id"]
            else:
                new_ref = ET.Element(
                    Q(ns, tag),
                    {Q(ns, "type"): v, Q(relationships_ns, "id"): new_rel_id},
                )
                _insert_sect_pr_child(sect_pr, new_ref)


@dataclass(frozen=True)
class SectionBreakTarget:
    """Discovered candidate target for section break authoring (insert, delete, or update)."""

    insert_after: int
    anchor_fingerprint_sha256: str
    section: int
    part: str = _DOCUMENT
    break_type: str = "nextPage"
    tag: str = "p"
    text: str = ""
    supported: bool = True
    has_break: bool = False
    can_delete: bool = False
    can_update: bool = False
    actions: tuple[str, ...] = ("insert",)
    link_to_previous: bool | None = None

    @property
    def paragraph(self) -> int:
        return self.insert_after

    @property
    def target(self) -> dict[str, Any]:
        return {"insert_after": self.insert_after, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "insert_after": self.insert_after,
            "paragraph": self.insert_after,
            "anchor_fingerprint_sha256": self.anchor_fingerprint_sha256,
            "section": self.section,
            "part": self.part,
            "break_type": self.break_type,
            "tag": self.tag,
            "text": self.text,
            "supported": self.supported,
            "has_break": self.has_break,
            "can_delete": self.can_delete,
            "can_update": self.can_update,
            "actions": list(self.actions),
        }
        if self.link_to_previous is not None:
            data["link_to_previous"] = self.link_to_previous
        return data


@dataclass(frozen=True)
class SectionBreakRequest:
    """Strict request for inserting, deleting, or updating a source-bound section break."""

    source: Path
    destination: Path
    source_sha256: str
    insert_after: int | None = None
    anchor_fingerprint_sha256: str = ""
    break_type: str = "nextPage"
    part: str = _DOCUMENT
    action: str = "insert"
    paragraph: int | None = None
    target: dict[str, Any] | None = None
    link_to_previous: bool | None = None

    def __post_init__(self) -> None:
        if isinstance(self.source, str):
            object.__setattr__(self, "source", Path(self.source))
        elif not isinstance(self.source, Path):
            raise TypeError("section_break_source_invalid")
        if not str(self.source):
            raise ValueError("section_break_source_invalid")

        if isinstance(self.destination, str):
            object.__setattr__(self, "destination", Path(self.destination))
        elif not isinstance(self.destination, Path):
            raise TypeError("section_break_destination_invalid")
        if not str(self.destination):
            raise ValueError("section_break_destination_invalid")

        _require_sha256(self.source_sha256, "section_break_source_sha256")
        _require_sha256(
            self.anchor_fingerprint_sha256,
            "section_break_anchor_fingerprint_sha256",
        )

        if not isinstance(self.action, str) or self.action not in _ALLOWED_ACTIONS:
            raise ValueError("section_break_action_invalid")

        target_candidates: list[tuple[str, Any]] = []

        if self.target is not None:
            if not isinstance(self.target, Mapping):
                raise ValueError("section_break_target_invalid")
            if "paragraph" in self.target:
                target_candidates.append(("target.paragraph", self.target["paragraph"]))
            if "insert_after" in self.target:
                target_candidates.append(("target.insert_after", self.target["insert_after"]))

        if self.paragraph is not None:
            target_candidates.append(("paragraph", self.paragraph))
        if self.insert_after is not None:
            target_candidates.append(("insert_after", self.insert_after))

        if not target_candidates:
            if self.action in {"delete", "update"}:
                raise ValueError("section_break_target_invalid")
            else:
                raise ValueError("section_break_insert_after_invalid")

        for name, val in target_candidates:
            if isinstance(val, bool) or not isinstance(val, int) or val < 1:
                if self.action in {"delete", "update"}:
                    raise ValueError("section_break_target_invalid")
                else:
                    raise ValueError("section_break_insert_after_invalid")

        distinct_targets = {val for _, val in target_candidates}
        if len(distinct_targets) > 1:
            raise ValueError("section_break_target_invalid")

        target_idx = target_candidates[0][1]

        object.__setattr__(self, "insert_after", target_idx)
        object.__setattr__(self, "paragraph", target_idx)

        if self.target is not None and isinstance(self.target, Mapping):
            canonical_target = dict(self.target)
            canonical_target["insert_after"] = target_idx
            canonical_target["paragraph"] = target_idx
            object.__setattr__(self, "target", canonical_target)

        canonical_break_type = _canonical_break_type(self.break_type)
        if canonical_break_type is not None:
            object.__setattr__(self, "break_type", canonical_break_type)
        elif self.action in {"insert", "update"}:
            raise ValueError("section_break_break_type_invalid")

        if not isinstance(self.part, str) or not self.part:
            raise ValueError("section_break_part_invalid")

        if self.link_to_previous is not None and not isinstance(self.link_to_previous, bool):
            raise TypeError("section_break_link_to_previous_invalid")

    @property
    def target_dict_value(self) -> dict[str, Any]:
        return self.target_dict()

    def target_dict(self) -> dict[str, Any]:
        idx = self.paragraph if self.paragraph is not None else self.insert_after
        if self.action in {"delete", "update"}:
            return {"paragraph": idx, "insert_after": idx, "part": self.part}
        return {"insert_after": self.insert_after, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "insert_after": self.insert_after,
            "anchor_fingerprint_sha256": self.anchor_fingerprint_sha256,
            "break_type": self.break_type,
            "part": self.part,
            "action": self.action,
        }
        if self.paragraph is not None:
            data["paragraph"] = self.paragraph
        if self.link_to_previous is not None:
            data["link_to_previous"] = self.link_to_previous
        return data


@dataclass(frozen=True)
class SectionBreakReport:
    """Content-safe public report for one section break authoring operation."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    section_break: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    boundaries: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None
    action: str = "insert"
    link_to_previous: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "schema_version": 1,
            "operation": "section-break",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "action": self.action,
        }
        if self.target is not None:
            result["target"] = self.target
        if self.section_break is not None:
            result["section_break"] = self.section_break
        if self.target_change_report is not None:
            result["target_change_report"] = self.target_change_report
        if self.boundaries is not None:
            result["boundaries"] = self.boundaries
        if self.non_target_parts_preserved is not None:
            result["non_target_parts_preserved"] = self.non_target_parts_preserved
        if self.link_to_previous is not None:
            result["link_to_previous"] = self.link_to_previous
        return result


def _section_break_refusal(
    request: SectionBreakRequest,
    code: str,
) -> SectionBreakReport:
    return SectionBreakReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=request.target_dict(),
        section_break=None,
        target_change_report=None,
        boundaries=None,
        non_target_parts_preserved=None,
        action=request.action,
        link_to_previous=request.link_to_previous,
    )


def parse_section_break_request(value: Any) -> SectionBreakRequest:
    """Parse and strictly validate a section break authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("section_break_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS):
        raise ValueError("section_break_request_fields_invalid")

    base_required = {
        "source",
        "destination",
        "source_sha256",
        "anchor_fingerprint_sha256",
    }
    if not base_required.issubset(keys):
        raise ValueError("section_break_request_fields_invalid")

    action = value.get("action", "insert")
    if not isinstance(action, str) or action not in _ALLOWED_ACTIONS:
        raise ValueError("section_break_action_invalid")

    has_target = ("insert_after" in keys) or ("paragraph" in keys) or ("target" in keys)
    if not has_target:
        raise ValueError("section_break_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, (str, Path)) or not str(source):
        raise ValueError("section_break_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, (str, Path)) or not str(destination):
        raise ValueError("section_break_destination_invalid")

    source_sha256 = _require_sha256(value["source_sha256"], "section_break_source_sha256")
    anchor_fingerprint = _require_sha256(
        value["anchor_fingerprint_sha256"],
        "section_break_anchor_fingerprint_sha256",
    )

    break_type = value.get("break_type", "nextPage")
    if action in {"insert", "update"} and _canonical_break_type(break_type) is None:
        raise ValueError("section_break_break_type_invalid")

    part = value.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("section_break_part_invalid")

    target_val = value.get("target")

    return SectionBreakRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        insert_after=value.get("insert_after"),
        paragraph=value.get("paragraph"),
        anchor_fingerprint_sha256=anchor_fingerprint,
        break_type=break_type,
        part=part,
        action=action,
        target=target_val,
        link_to_previous=value.get("link_to_previous"),
    )


def load_section_break_request(path: str | Path) -> SectionBreakRequest:
    """Load and strictly validate one JSON section-break request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_section_break_request(json.load(stream))


def section_break_anchor_fingerprint(
    source: str | Path,
    *,
    insert_after: int = 1,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural anchor fingerprint for one exact body index."""
    if part != _DOCUMENT:
        return None
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        if part not in graph.parts:
            return None
        root = ET.fromstring(graph.read_part_bytes(part))
        ns = _word_namespace(root)
        if ns is None:
            return None
        items = _body_items(root, ns)
        if insert_after < 1 or insert_after > len(items):
            return None
        return _xml_hash(items[insert_after - 1])
    except (OSError, PackageError, ET.ParseError, KeyError, zipfile.BadZipFile):
        return None


def discover_section_break_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate body targets for section break authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "section-break-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "section_break_source_missing",
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
            "operation": "section-break-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "section_break_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    if part != _DOCUMENT or part not in graph.parts:
        return {
            "schema_version": 1,
            "operation": "section-break-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "section_break_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "section-break-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "section_break_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    ns = _word_namespace(root)
    if ns is None:
        return {
            "schema_version": 1,
            "operation": "section-break-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "section_break_namespace_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    body = root.find(Q(ns, "body"))
    if body is None:
        return {
            "schema_version": 1,
            "operation": "section-break-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "section_break_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    unsupported_context = _unsupported_context_by_element(root, ns)
    sections = _find_sections(root, ns)
    items = _body_items(root, ns)

    unsupported_tags = {
        Q(ns, "ins"),
        Q(ns, "del"),
        Q(ns, "moveFrom"),
        Q(ns, "moveTo"),
        Q(ns, "sdt"),
        Q(ns, "customXml"),
        Q(ns, "pPrChange"),
        Q(ns, "sectPrChange"),
    }

    targets: list[dict[str, Any]] = []
    has_duplicate_body_sect = len(body.findall(Q(ns, "sectPr"))) > 1

    for idx, item in enumerate(items, start=1):
        containing_sect = _find_containing_sect_pr(body, item, ns)
        sec_idx = 1
        if containing_sect is not None:
            for s_i, s in enumerate(sections, start=1):
                if s is containing_sect:
                    sec_idx = s_i
                    break

        b_type = "nextPage"
        has_duplicate_type = False
        if containing_sect is not None:
            types = containing_sect.findall(Q(ns, "type"))
            if len(types) > 1:
                has_duplicate_type = True
            elif types:
                b_type = types[0].get(Q(ns, "val")) or types[0].get("val") or "nextPage"

        tag = item.tag.rsplit("}", 1)[-1] if "}" in item.tag else item.tag
        text = "".join(node.text or "" for node in item.iter(Q(ns, "t")))[:100]

        has_duplicate_ppr = item.tag == Q(ns, "p") and len(item.findall(Q(ns, "pPr"))) > 1
        has_unsupported_tags = any(c.tag in unsupported_tags for c in item.iter())
        in_unsupported_ctx = unsupported_context.get(id(item), False)
        sect_unsupported = containing_sect is None or _is_unsupported_section(
            containing_sect, unsupported_context, ns
        )

        is_supported = (
            not in_unsupported_ctx
            and not has_unsupported_tags
            and not has_duplicate_ppr
            and not sect_unsupported
            and not has_duplicate_body_sect
            and not has_duplicate_type
        )

        # Check if item itself has a paragraph sectPr (deletable or updatable break)
        item_ppr = item.find(Q(ns, "pPr")) if item.tag == Q(ns, "p") else None
        item_sect = item_ppr.find(Q(ns, "sectPr")) if item_ppr is not None else None
        has_break = item_sect is not None

        item_sect_unsupported = False
        if item_sect is not None:
            item_sect_unsupported = (
                _is_unsupported_section(item_sect, unsupported_context, ns)
                or len(item_sect.findall(Q(ns, "type"))) > 1
            )

        can_update = has_break and is_supported and not item_sect_unsupported

        can_delete = False
        if has_break and is_supported and not item_sect_unsupported:
            item_sec_idx = None
            for s_i, s in enumerate(sections, start=1):
                if s is item_sect:
                    item_sec_idx = s_i
                    break
            if item_sec_idx is not None and item_sec_idx < len(sections):
                succ_sect = sections[item_sec_idx]
                if not _is_unsupported_section(succ_sect, unsupported_context, ns):
                    rels_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    _prec_refs = [
                        (
                            c.tag,
                            c.attrib.get(Q(ns, "type"), "default"),
                            c.attrib.get(Q(rels_ns, "id"))
                            or c.attrib.get(Q(ns, "id"))
                            or c.attrib.get("id"),
                        )
                        for c in item_sect
                        if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                    ]
                    _succ_refs = [
                        (
                            c.tag,
                            c.attrib.get(Q(ns, "type"), "default"),
                            c.attrib.get(Q(rels_ns, "id"))
                            or c.attrib.get(Q(ns, "id"))
                            or c.attrib.get("id"),
                        )
                        for c in succ_sect
                        if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                    ]
                    _prec_title = _is_title_pg_enabled(item_sect, ns)
                    _succ_title = _is_title_pg_enabled(succ_sect, ns)
                    if _prec_refs == _succ_refs and _prec_title == _succ_title:
                        can_delete = True

        target_link_to_previous: bool | None = None
        if has_break and is_supported and not item_sect_unsupported:
            item_sec_idx = None
            for s_i, s in enumerate(sections, start=1):
                if s is item_sect:
                    item_sec_idx = s_i
                    break
            if item_sec_idx is not None and item_sec_idx < len(sections):
                succ_sect = sections[item_sec_idx]
                succ_refs = [
                    c
                    for c in succ_sect
                    if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                ]
                target_link_to_previous = len(succ_refs) == 0

        actions_list = ["insert"]
        if can_delete:
            actions_list.append("delete")
        if can_update:
            actions_list.append("update")
        actions = tuple(actions_list)

        target = SectionBreakTarget(
            insert_after=idx,
            anchor_fingerprint_sha256=_xml_hash(item),
            section=sec_idx,
            part=part,
            break_type=b_type,
            tag=tag,
            text=text,
            supported=is_supported,
            has_break=has_break,
            can_delete=can_delete,
            can_update=can_update,
            actions=actions,
            link_to_previous=target_link_to_previous,
        )
        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "section-break-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_section_break(request: SectionBreakRequest) -> SectionBreakReport:
    """Insert a source-bound body section break atomically and verification-checked."""
    source = request.source
    destination = request.destination

    if not source.is_file():
        return _section_break_refusal(request, "section_break_source_missing")

    try:
        if source.resolve() == destination.resolve():
            return _section_break_refusal(request, "section_break_destination_invalid")
    except OSError:
        return _section_break_refusal(request, "section_break_destination_invalid")

    if not destination.parent.is_dir():
        return _section_break_refusal(request, "section_break_destination_invalid")

    if destination.exists() or destination.is_symlink():
        return _section_break_refusal(request, "section_break_destination_occupied")

    try:
        actual_source_sha256 = _sha256_file(source)
    except OSError:
        return _section_break_refusal(request, "section_break_source_missing")

    if actual_source_sha256 != request.source_sha256:
        return _section_break_refusal(request, "section_break_source_hash_mismatch")

    if request.part != _DOCUMENT:
        return _section_break_refusal(request, "section_break_part_unsupported")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _section_break_refusal(request, "section_break_package_invalid")

    if _DOCUMENT not in graph.parts:
        return _section_break_refusal(request, "section_break_package_invalid")

    try:
        doc_bytes = graph.read_part_bytes(_DOCUMENT)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (KeyError, ET.ParseError):
        return _section_break_refusal(request, "section_break_package_invalid")

    ns = _word_namespace(root)
    if ns is None:
        return _section_break_refusal(request, "section_break_namespace_unsupported")

    body = root.find(Q(ns, "body"))
    if body is None:
        return _section_break_refusal(request, "section_break_package_invalid")

    # Check for duplicate body-final sectPr
    body_final_sects = body.findall(Q(ns, "sectPr"))
    if len(body_final_sects) > 1:
        return _section_break_refusal(request, "section_break_context_unsupported")

    if request.paragraph != request.insert_after:
        return _section_break_refusal(request, "section_break_target_invalid")

    target_idx = (
        request.paragraph
        if request.paragraph is not None
        else (request.insert_after if request.insert_after is not None else 1)
    )
    existing_items = _body_items(root, ns)
    if target_idx < 1 or target_idx > len(existing_items):
        return _section_break_refusal(request, "section_break_anchor_missing")

    anchor = existing_items[target_idx - 1]
    if _xml_hash(anchor) != request.anchor_fingerprint_sha256:
        return _section_break_refusal(request, "section_break_anchor_stale")
    # Check unsupported context on anchor
    unsupported_context = _unsupported_context_by_element(root, ns)
    if unsupported_context.get(id(anchor), False):
        return _section_break_refusal(request, "section_break_context_unsupported")

    if anchor.tag == Q(ns, "p") and len(anchor.findall(Q(ns, "pPr"))) > 1:
        return _section_break_refusal(request, "section_break_context_unsupported")

    unsupported_tags = {
        Q(ns, "ins"),
        Q(ns, "del"),
        Q(ns, "moveFrom"),
        Q(ns, "moveTo"),
        Q(ns, "sdt"),
        Q(ns, "customXml"),
        Q(ns, "pPrChange"),
        Q(ns, "sectPrChange"),
    }
    if any(child.tag in unsupported_tags for child in anchor.iter()):
        return _section_break_refusal(request, "section_break_context_unsupported")

    if request.action == "delete":
        anchor_ppr = anchor.find(Q(ns, "pPr")) if anchor.tag == Q(ns, "p") else None
        anchor_sect = anchor_ppr.find(Q(ns, "sectPr")) if anchor_ppr is not None else None
        if anchor_sect is None:
            return _section_break_refusal(request, "section_break_target_missing")

        if _is_unsupported_section(anchor_sect, unsupported_context, ns):
            return _section_break_refusal(request, "section_break_context_unsupported")
        if len(anchor_sect.findall(Q(ns, "type"))) > 1:
            return _section_break_refusal(request, "section_break_context_unsupported")

        sections_before = _find_sections(root, ns)
        preceding_sec_index = None
        for s_i, s in enumerate(sections_before, start=1):
            if s is anchor_sect:
                preceding_sec_index = s_i
                break
        if preceding_sec_index is None:
            return _section_break_refusal(request, "section_break_section_missing")

        if preceding_sec_index >= len(sections_before):
            return _section_break_refusal(request, "section_break_final_section_unsupported")

        succeeding_sect_pr = sections_before[preceding_sec_index]
        if _is_unsupported_section(succeeding_sect_pr, unsupported_context, ns):
            return _section_break_refusal(request, "section_break_context_unsupported")
        if len(succeeding_sect_pr.findall(Q(ns, "type"))) > 1:
            return _section_break_refusal(request, "section_break_context_unsupported")

        rels_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        prec_refs = [
            (
                c.tag,
                c.attrib.get(Q(ns, "type"), "default"),
                c.attrib.get(Q(rels_ns, "id")) or c.attrib.get(Q(ns, "id")) or c.attrib.get("id"),
            )
            for c in anchor_sect
            if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
        ]
        succ_refs = [
            (
                c.tag,
                c.attrib.get(Q(ns, "type"), "default"),
                c.attrib.get(Q(rels_ns, "id")) or c.attrib.get(Q(ns, "id")) or c.attrib.get("id"),
            )
            for c in succeeding_sect_pr
            if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
        ]
        if prec_refs != succ_refs:
            return _section_break_refusal(request, "section_break_context_unsupported")

        prec_title_pg = _is_title_pg_enabled(anchor_sect, ns)
        succ_title_pg = _is_title_pg_enabled(succeeding_sect_pr, ns)
        if prec_title_pg != succ_title_pg:
            return _section_break_refusal(request, "section_break_context_unsupported")

        if (anchor_sect.find(Q(ns, "footnotePr")) is not None) != (
            succeeding_sect_pr.find(Q(ns, "footnotePr")) is not None
        ):
            return _section_break_refusal(request, "section_break_context_unsupported")
        if (anchor_sect.find(Q(ns, "endnotePr")) is not None) != (
            succeeding_sect_pr.find(Q(ns, "endnotePr")) is not None
        ):
            return _section_break_refusal(request, "section_break_context_unsupported")

        succ_type_elem = succeeding_sect_pr.find(Q(ns, "type"))
        succ_break_type = (
            succ_type_elem.get(Q(ns, "val")) or succ_type_elem.get("val") or "nextPage"
            if succ_type_elem is not None
            else "nextPage"
        )

        # Remove ONLY anchor_sect from anchor_ppr; anchor paragraph, text, and other pPr children stay intact
        anchor_ppr.remove(anchor_sect)

        sections_after = _find_sections(root, ns)
        if len(sections_after) != len(sections_before) - 1:
            return _section_break_refusal(request, "section_break_staged_verification_failed")

        boundaries = {
            "deleted_section_break_paragraph_index": target_idx,
            "merged_section_index": preceding_sec_index,
            "succeeding_section_break_type": succ_break_type,
            "sections_before": len(sections_before),
            "total_sections": len(sections_after),
            "action": "delete",
        }

        target_change_report = {
            "action": "delete",
            "paragraph": target_idx,
            "insert_after": target_idx,
            "deleted_section_index": preceding_sec_index,
            "merged_into_section_index": preceding_sec_index,
            "total_sections": len(sections_after),
        }

        def verify_deletion_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, frozenset({_DOCUMENT})):
                return "section_break_staged_verification_failed"
            insp = inspect_docx(staged)
            if not insp.valid:
                return "section_break_staged_verification_failed"
            try:
                s_graph = PackageGraph.from_file(staged)
                s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
                s_ns = _word_namespace(s_root)
                if s_ns != ns:
                    return "section_break_staged_verification_failed"

                s_items = _body_items(s_root, s_ns)
                if len(s_items) != len(existing_items):
                    return "section_break_staged_verification_failed"

                s_anchor = s_items[target_idx - 1]
                if s_anchor.tag != Q(s_ns, "p"):
                    return "section_break_staged_verification_failed"
                s_ppr = s_anchor.find(Q(s_ns, "pPr"))
                if s_ppr is not None and s_ppr.find(Q(s_ns, "sectPr")) is not None:
                    return "section_break_staged_verification_failed"

                orig_text = "".join(node.text or "" for node in anchor.iter(Q(ns, "t")))
                staged_text = "".join(node.text or "" for node in s_anchor.iter(Q(s_ns, "t")))
                if orig_text != staged_text:
                    return "section_break_staged_verification_failed"

                s_sections = _find_sections(s_root, s_ns)
                if len(s_sections) != len(sections_before) - 1:
                    return "section_break_staged_verification_failed"

                staged_succ_sect = s_sections[preceding_sec_index - 1]
                for geom_tag in ("pgSz", "pgMar"):
                    orig_geom = succeeding_sect_pr.find(Q(ns, geom_tag))
                    staged_geom = staged_succ_sect.find(Q(s_ns, geom_tag))
                    if orig_geom is not None and staged_geom is None:
                        return "section_break_staged_verification_failed"
                    if (
                        orig_geom is not None
                        and staged_geom is not None
                        and orig_geom.attrib != staged_geom.attrib
                    ):
                        return "section_break_staged_verification_failed"

                s_body = s_root.find(Q(s_ns, "body"))
                if s_body is None or len(s_body.findall(Q(s_ns, "sectPr"))) != 1:
                    return "section_break_staged_verification_failed"

            except (
                OSError,
                PackageError,
                ET.ParseError,
                KeyError,
                IndexError,
                ValueError,
            ):
                return "section_break_staged_verification_failed"
            return None

        serialized = serialize_xml_with_preserved_qnames(root)
        replacements = {_DOCUMENT: serialized}

        published: ReviewReport = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[target_change_report],
            verify=verify_deletion_publication,
            failure_code="section_break_staged_verification_failed",
            exists_code="section_break_destination_occupied",
            unavailable_code="section_break_publication_unavailable",
        )

        if not published.passed:
            return _section_break_refusal(
                request,
                published.code or "section_break_staged_verification_failed",
            )

        return SectionBreakReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            target=request.target_dict(),
            section_break=target_change_report,
            target_change_report=target_change_report,
            boundaries=boundaries,
            non_target_parts_preserved=True,
            action="delete",
        )
    if request.action == "update":
        anchor_ppr = anchor.find(Q(ns, "pPr")) if anchor.tag == Q(ns, "p") else None
        anchor_sect = anchor_ppr.find(Q(ns, "sectPr")) if anchor_ppr is not None else None
        if anchor_sect is None:
            return _section_break_refusal(request, "section_break_target_missing")

        if _is_unsupported_section(anchor_sect, unsupported_context, ns):
            return _section_break_refusal(request, "section_break_context_unsupported")
        if len(anchor_sect.findall(Q(ns, "type"))) > 1:
            return _section_break_refusal(request, "section_break_context_unsupported")

        if request.break_type not in _ALLOWED_BREAK_TYPES:
            return _section_break_refusal(request, "section_break_break_type_invalid")

        sections_before = _find_sections(root, ns)
        target_sec_index = None
        for s_i, s in enumerate(sections_before, start=1):
            if s is anchor_sect:
                target_sec_index = s_i
                break
        if target_sec_index is None:
            return _section_break_refusal(request, "section_break_section_missing")

        orig_type_node = anchor_sect.find(Q(ns, "type"))
        orig_break_type = (
            orig_type_node.get(Q(ns, "val")) or orig_type_node.get("val") or "nextPage"
            if orig_type_node is not None
            else "nextPage"
        )

        if orig_type_node is not None:
            orig_type_node.set(Q(ns, "val"), request.break_type)
            if "val" in orig_type_node.attrib and Q(ns, "val") in orig_type_node.attrib:
                del orig_type_node.attrib["val"]
        else:
            new_type_elem = ET.Element(Q(ns, "type"), {Q(ns, "val"): request.break_type})
            _insert_sect_pr_child(anchor_sect, new_type_elem)

        sections_after = _find_sections(root, ns)
        if len(sections_after) != len(sections_before):
            return _section_break_refusal(request, "section_break_staged_verification_failed")
        rels_name = _rels_name(_DOCUMENT)
        rels_root = (
            parse_xml_with_preserved_qnames(graph.read_part_bytes(rels_name))
            if rels_name in graph.parts
            else ET.Element(Q(_PR, "Relationships"))
        )
        types_root = parse_xml_with_preserved_qnames(graph.read_part_bytes("[Content_Types].xml"))
        additions: dict[str, bytes] = {}

        if request.link_to_previous is not None and target_sec_index < len(sections_before):
            succ_sect = sections_before[target_sec_index]
            preds = sections_before[:target_sec_index]
            _apply_link_to_previous(
                succ_sect,
                preds,
                request.link_to_previous,
                graph,
                rels_root,
                types_root,
                additions,
                ns,
            )

        boundaries = {
            "paragraph": target_idx,
            "insert_after": target_idx,
            "section_index": target_sec_index,
            "previous_break_type": orig_break_type,
            "break_type": request.break_type,
            "total_sections": len(sections_after),
            "action": "update",
        }
        if request.link_to_previous is not None:
            boundaries["link_to_previous"] = request.link_to_previous

        target_change_report = {
            "action": "update",
            "paragraph": target_idx,
            "insert_after": target_idx,
            "section_index": target_sec_index,
            "previous_break_type": orig_break_type,
            "break_type": request.break_type,
            "total_sections": len(sections_after),
        }
        if request.link_to_previous is not None:
            target_change_report["link_to_previous"] = request.link_to_previous

        all_changed_update = frozenset(
            {_DOCUMENT}
            | set(additions.keys())
            | ({rels_name, "[Content_Types].xml"} if additions else set())
        )

        def verify_update_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, all_changed_update):
                return "section_break_staged_verification_failed"
            insp = inspect_docx(staged)
            if not insp.valid:
                return "section_break_staged_verification_failed"
            try:
                s_graph = PackageGraph.from_file(staged)
                s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
                s_ns = _word_namespace(s_root)
                if s_ns != ns:
                    return "section_break_staged_verification_failed"

                s_items = _body_items(s_root, s_ns)
                if len(s_items) != len(existing_items):
                    return "section_break_staged_verification_failed"

                s_anchor = s_items[target_idx - 1]
                if s_anchor.tag != Q(s_ns, "p"):
                    return "section_break_staged_verification_failed"
                s_ppr = s_anchor.find(Q(s_ns, "pPr"))
                if s_ppr is None:
                    return "section_break_staged_verification_failed"
                s_sect = s_ppr.find(Q(s_ns, "sectPr"))
                if s_sect is None:
                    return "section_break_staged_verification_failed"

                orig_text = "".join(node.text or "" for node in anchor.iter(Q(ns, "t")))
                staged_text = "".join(node.text or "" for node in s_anchor.iter(Q(s_ns, "t")))
                if orig_text != staged_text:
                    return "section_break_staged_verification_failed"

                s_sections = _find_sections(s_root, s_ns)
                if len(s_sections) != len(sections_before):
                    return "section_break_staged_verification_failed"

                s_type_elem = s_sect.find(Q(s_ns, "type"))
                if s_type_elem is None:
                    return "section_break_staged_verification_failed"
                s_type_val = s_type_elem.get(Q(s_ns, "val")) or s_type_elem.get("val")
                if s_type_val != request.break_type:
                    return "section_break_staged_verification_failed"

                ranks = [_sect_pr_rank(c) for c in s_sect]
                if ranks != sorted(ranks):
                    return "section_break_staged_verification_failed"

                for geom_tag in ("pgSz", "pgMar"):
                    orig_geom = anchor_sect.find(Q(ns, geom_tag))
                    st_geom = s_sect.find(Q(s_ns, geom_tag))
                    if orig_geom is not None and st_geom is None:
                        return "section_break_staged_verification_failed"
                    if (
                        orig_geom is not None
                        and st_geom is not None
                        and orig_geom.attrib != st_geom.attrib
                    ):
                        return "section_break_staged_verification_failed"

                orig_refs = [
                    (c.tag, c.attrib)
                    for c in anchor_sect
                    if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                ]
                st_refs = [
                    (c.tag, c.attrib)
                    for c in s_sect
                    if c.tag in {Q(s_ns, "headerReference"), Q(s_ns, "footerReference")}
                ]
                if orig_refs != st_refs:
                    return "section_break_staged_verification_failed"
                if request.link_to_previous is not None:
                    s_successor = s_sections[target_sec_index]
                    s_successor_refs = [
                        c
                        for c in s_successor
                        if c.tag in {Q(s_ns, "headerReference"), Q(s_ns, "footerReference")}
                    ]
                    if request.link_to_previous is True and s_successor_refs:
                        return "section_break_staged_verification_failed"
                    if request.link_to_previous is False:
                        source_relationships_ns = _STRICT_R if ns == STRICT_WORD_NS else _R
                        staged_relationships_ns = _STRICT_R if s_ns == STRICT_WORD_NS else _R
                        predecessor_rel_ids = {
                            c.get(Q(source_relationships_ns, "id")) or c.get("id")
                            for predecessor in sections_before[:target_sec_index]
                            for c in predecessor
                            if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                        }
                        if not s_successor_refs or any(
                            (c.get(Q(staged_relationships_ns, "id")) or c.get("id"))
                            in predecessor_rel_ids
                            for c in s_successor_refs
                        ):
                            return "section_break_staged_verification_failed"

                for sec_i, orig_s in enumerate(sections_before):
                    if orig_s is anchor_sect:
                        continue
                    st_s = s_sections[sec_i]
                    orig_type = orig_s.find(Q(ns, "type"))
                    st_type = st_s.find(Q(s_ns, "type"))
                    orig_tv = (
                        (orig_type.get(Q(ns, "val")) or orig_type.get("val"))
                        if orig_type is not None
                        else None
                    )
                    st_tv = (
                        (st_type.get(Q(s_ns, "val")) or st_type.get("val"))
                        if st_type is not None
                        else None
                    )
                    if orig_tv != st_tv:
                        return "section_break_staged_verification_failed"

                s_body = s_root.find(Q(s_ns, "body"))
                if s_body is None or len(s_body.findall(Q(s_ns, "sectPr"))) != 1:
                    return "section_break_staged_verification_failed"

            except (
                OSError,
                PackageError,
                ET.ParseError,
                KeyError,
                IndexError,
                ValueError,
            ):
                return "section_break_staged_verification_failed"
            return None

        serialized = serialize_xml_with_preserved_qnames(root)
        replacements = {_DOCUMENT: serialized}
        if additions:
            replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)
            replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)

        published: ReviewReport = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[target_change_report],
            additions=additions,
            verify=verify_update_publication,
            failure_code="section_break_staged_verification_failed",
            exists_code="section_break_destination_occupied",
            unavailable_code="section_break_publication_unavailable",
        )

        if not published.passed:
            return _section_break_refusal(
                request,
                published.code or "section_break_staged_verification_failed",
            )

        return SectionBreakReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            target=request.target_dict(),
            section_break=target_change_report,
            target_change_report=target_change_report,
            boundaries=boundaries,
            non_target_parts_preserved=True,
            action="update",
            link_to_previous=request.link_to_previous,
        )
    # If anchor carries a terminating sectPr, validate its context and uniqueness
    anchor_ppr = anchor.find(Q(ns, "pPr")) if anchor.tag == Q(ns, "p") else None
    anchor_sect = anchor_ppr.find(Q(ns, "sectPr")) if anchor_ppr is not None else None
    if anchor_sect is not None:
        if _is_unsupported_section(anchor_sect, unsupported_context, ns):
            return _section_break_refusal(request, "section_break_context_unsupported")
        if len(anchor_sect.findall(Q(ns, "type"))) > 1:
            return _section_break_refusal(request, "section_break_context_unsupported")

    # Resolve containing section as first terminating sectPr at/after anchor
    sections_before = _find_sections(root, ns)
    containing_sect_pr = _find_containing_sect_pr(body, anchor, ns)
    if containing_sect_pr is None:
        return _section_break_refusal(request, "section_break_section_missing")

    containing_sec_index = 1
    for s_i, s in enumerate(sections_before, start=1):
        if s is containing_sect_pr:
            containing_sec_index = s_i
            break

    if _is_unsupported_section(containing_sect_pr, unsupported_context, ns):
        return _section_break_refusal(request, "section_break_context_unsupported")

    if len(containing_sect_pr.findall(Q(ns, "type"))) > 1:
        return _section_break_refusal(request, "section_break_context_unsupported")

    if request.break_type not in _ALLOWED_BREAK_TYPES:
        return _section_break_refusal(request, "section_break_break_type_invalid")

    # Extract containing section's original type (for previous-section preservation)
    orig_type_node = containing_sect_pr.find(Q(ns, "type"))
    orig_break_type = (
        orig_type_node.get(Q(ns, "val")) or orig_type_node.get("val") or "nextPage"
        if orig_type_node is not None
        else "nextPage"
    )

    # Clone containing terminating sectPr with root rsid attributes cleared
    cloned_sect_pr = copy.deepcopy(containing_sect_pr)
    _clear_root_rsid_attrs(cloned_sect_pr)

    # Insert new paragraph carrying cloned sectPr directly after anchor
    new_p = ET.Element(Q(ns, "p"))
    new_ppr = ET.SubElement(new_p, Q(ns, "pPr"))
    new_ppr.append(cloned_sect_pr)

    anchor_body_index = list(body).index(anchor)
    body.insert(anchor_body_index + 1, new_p)

    # Update containing section's original terminating sectPr with requested break_type
    if orig_type_node is not None:
        orig_type_node.set(Q(ns, "val"), request.break_type)
        if "val" in orig_type_node.attrib and Q(ns, "val") in orig_type_node.attrib:
            del orig_type_node.attrib["val"]
    else:
        new_type_elem = ET.Element(Q(ns, "type"), {Q(ns, "val"): request.break_type})
        _insert_sect_pr_child(containing_sect_pr, new_type_elem)
    rels_name = _rels_name(_DOCUMENT)
    rels_root = (
        parse_xml_with_preserved_qnames(graph.read_part_bytes(rels_name))
        if rels_name in graph.parts
        else ET.Element(Q(_PR, "Relationships"))
    )
    types_root = parse_xml_with_preserved_qnames(graph.read_part_bytes("[Content_Types].xml"))
    additions: dict[str, bytes] = {}

    if request.link_to_previous is not None:
        preds = list(sections_before[: containing_sec_index - 1]) + [cloned_sect_pr]
        _apply_link_to_previous(
            containing_sect_pr,
            preds,
            request.link_to_previous,
            graph,
            rels_root,
            types_root,
            additions,
            ns,
        )

    sections_after = _find_sections(root, ns)
    if len(sections_after) != len(sections_before) + 1:
        return _section_break_refusal(request, "section_break_staged_verification_failed")

    boundaries = {
        "previous_section": {
            "section_index": containing_sec_index,
            "break_type": orig_break_type,
        },
        "following_section": {
            "section_index": containing_sec_index + 1,
            "break_type": request.break_type,
        },
        "inserted_paragraph_body_index": request.insert_after + 1,
        "total_sections": len(sections_after),
        "break_type": request.break_type,
    }
    if request.link_to_previous is not None:
        boundaries["link_to_previous"] = request.link_to_previous

    target_change_report = {
        "insert_after": request.insert_after,
        "break_type": request.break_type,
        "previous_section_index": containing_sec_index,
        "following_section_index": containing_sec_index + 1,
        "total_sections": len(sections_after),
    }
    if request.link_to_previous is not None:
        target_change_report["link_to_previous"] = request.link_to_previous

    all_changed_insert = frozenset(
        {_DOCUMENT}
        | set(additions.keys())
        | ({rels_name, "[Content_Types].xml"} if additions else set())
    )

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, all_changed_insert):
            return "section_break_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "section_break_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
            s_ns = _word_namespace(s_root)
            if s_ns != ns:
                return "section_break_staged_verification_failed"

            s_items = _body_items(s_root, s_ns)
            if len(s_items) != len(existing_items) + 1:
                return "section_break_staged_verification_failed"

            # Verify the inserted paragraph position
            ins_elem = s_items[request.insert_after]
            if ins_elem.tag != Q(s_ns, "p"):
                return "section_break_staged_verification_failed"
            ins_ppr = ins_elem.find(Q(s_ns, "pPr"))
            if ins_ppr is None:
                return "section_break_staged_verification_failed"
            ins_sect = ins_ppr.find(Q(s_ns, "sectPr"))
            if ins_sect is None:
                return "section_break_staged_verification_failed"

            # Verify no root rsid attributes on inserted sectPr
            if any(
                (k.rsplit("}", 1)[-1] if "}" in k else k).lower().startswith("rsid")
                for k in ins_sect.attrib
            ):
                return "section_break_staged_verification_failed"

            # If anchor carried a terminating sectPr, verify it was preserved byte-for-byte
            if anchor_sect is not None:
                staged_anchor = s_items[request.insert_after - 1]
                staged_anchor_ppr = staged_anchor.find(Q(s_ns, "pPr"))
                if staged_anchor_ppr is None:
                    return "section_break_staged_verification_failed"
                staged_anchor_sect = staged_anchor_ppr.find(Q(s_ns, "sectPr"))
                if staged_anchor_sect is None or ET.tostring(staged_anchor_sect) != ET.tostring(
                    anchor_sect
                ):
                    return "section_break_staged_verification_failed"

            s_sections = _find_sections(s_root, s_ns)
            if len(s_sections) != len(sections_before) + 1:
                return "section_break_staged_verification_failed"

            # Verify previous section preserved its type
            ins_type_elem = ins_sect.find(Q(s_ns, "type"))
            ins_b_type = (
                ins_type_elem.get(Q(s_ns, "val")) or ins_type_elem.get("val") or "nextPage"
                if ins_type_elem is not None
                else "nextPage"
            )
            if ins_b_type != orig_break_type:
                return "section_break_staged_verification_failed"

            # Verify following section received requested break_type
            fol_sec = s_sections[containing_sec_index]
            fol_type_elem = fol_sec.find(Q(s_ns, "type"))
            if fol_type_elem is None:
                return "section_break_staged_verification_failed"
            fol_b_type = fol_type_elem.get(Q(s_ns, "val")) or fol_type_elem.get("val")
            if fol_b_type != request.break_type:
                return "section_break_staged_verification_failed"

            # Verify monotonic schema order on both section properties
            for sec_elem in (ins_sect, fol_sec):
                ranks = [_sect_pr_rank(c) for c in sec_elem]
                if ranks != sorted(ranks):
                    return "section_break_staged_verification_failed"

            # Verify geometry and header/footer references preserved on inserted sectPr
            for geom_tag in ("pgSz", "pgMar"):
                orig_geom = containing_sect_pr.find(Q(ns, geom_tag))
                ins_geom = ins_sect.find(Q(s_ns, geom_tag))
                if orig_geom is not None and ins_geom is None:
                    return "section_break_staged_verification_failed"
                if (
                    orig_geom is not None
                    and ins_geom is not None
                    and orig_geom.attrib != ins_geom.attrib
                ):
                    return "section_break_staged_verification_failed"

            if request.link_to_previous is True:
                fol_refs = [
                    c
                    for c in fol_sec
                    if c.tag in {Q(s_ns, "headerReference"), Q(s_ns, "footerReference")}
                ]
                if fol_refs:
                    return "section_break_staged_verification_failed"
            elif request.link_to_previous is False:
                relationships_ns = _STRICT_R if s_ns == STRICT_WORD_NS else _R
                fol_refs = [
                    c
                    for c in fol_sec
                    if c.tag in {Q(s_ns, "headerReference"), Q(s_ns, "footerReference")}
                ]
                previous_ref_ids = {
                    c.get(Q(relationships_ns, "id")) or c.get("id")
                    for c in ins_sect
                    if c.tag in {Q(s_ns, "headerReference"), Q(s_ns, "footerReference")}
                }
                if not fol_refs or any(
                    (c.get(Q(relationships_ns, "id")) or c.get("id")) in previous_ref_ids
                    for c in fol_refs
                ):
                    return "section_break_staged_verification_failed"
            else:
                orig_refs = [
                    (c.tag, c.attrib)
                    for c in containing_sect_pr
                    if c.tag in {Q(ns, "headerReference"), Q(ns, "footerReference")}
                ]
                ins_refs = [
                    (c.tag, c.attrib)
                    for c in ins_sect
                    if c.tag in {Q(s_ns, "headerReference"), Q(s_ns, "footerReference")}
                ]
                if orig_refs != ins_refs:
                    return "section_break_staged_verification_failed"
        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return "section_break_staged_verification_failed"
        return None

    serialized = serialize_xml_with_preserved_qnames(root)
    replacements = {_DOCUMENT: serialized}
    if additions:
        replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)
        replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=[target_change_report],
        additions=additions,
        verify=verify_publication,
        failure_code="section_break_staged_verification_failed",
        exists_code="section_break_destination_occupied",
        unavailable_code="section_break_publication_unavailable",
    )

    if not published.passed:
        return _section_break_refusal(
            request,
            published.code or "section_break_staged_verification_failed",
        )

    return SectionBreakReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target={"insert_after": request.insert_after, "part": request.part},
        section_break=target_change_report,
        target_change_report=target_change_report,
        boundaries=boundaries,
        non_target_parts_preserved=True,
        link_to_previous=request.link_to_previous,
    )
