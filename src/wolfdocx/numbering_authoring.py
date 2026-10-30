"""Source-bound DOCX paragraph numbering authoring and discovery.

Applies existing numbering definitions or restarts numbering via new w:num
instances referencing existing abstractNum definitions with validated startOverride.
A restart that names neither num_id nor abstract_num_id continues from the target
paragraph's effective numbering instance (direct w:numPr, else style-inherited).
Preserves package parts, unrelated numbering definitions, and schema element order.
Refuses absent or unsupported abstract definitions, style-linked numbering,
picture bullets, and unsafe wrappers.
"""

from __future__ import annotations

import copy
import hashlib
import json
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .authoring import _unsupported_context_by_element
from .inspect import (
    Q,
    STRICT_WORD_NS,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import (
    _DOCUMENT,
    _mask_text,
    _positive,
    _publish,
    _sha,
    _word_namespace,
)

_PUBLISHED = "published"
_REFUSED = "refused"
_NUMBERING_PART = "word/numbering.xml"
_STYLES_PART = "word/styles.xml"
_NUMBERING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
)
_STRICT_NUMBERING_REL_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/numbering"
)
_NUMBERING_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
)
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_NOTE_FORMATS = frozenset(
    {
        "decimal",
        "decimalZero",
        "lowerLetter",
        "upperLetter",
        "lowerRoman",
        "upperRoman",
        "none",
    }
)
_NUMBER_FORMATS = _NOTE_FORMATS | {"bullet"}

_UNSUPPORTED_PARAGRAPH_TAGS = frozenset(
    {
        "fldChar",
        "fldSimple",
        "sdt",
        "sdtContent",
        "customXml",
        "smartTag",
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
    }
)

_PPR_SUCCEEDING_TAGS = frozenset(
    {
        "pBdr",
        "shd",
        "tabs",
        "suppressAutoHyphens",
        "kinsoku",
        "wordWrap",
        "overflowPunct",
        "topLinePunct",
        "autoSpaceDE",
        "autoSpaceDN",
        "bidi",
        "adjustRightInd",
        "snapToGrid",
        "spacing",
        "ind",
        "contextualSpacing",
        "mirrorIndents",
        "suppressOverlap",
        "jc",
        "textDirection",
        "textAlignment",
        "textboxTightWrap",
        "outlineLvl",
        "divId",
        "cnfStyle",
        "rPr",
        "sectPr",
        "pPrChange",
    }
)


@dataclass(frozen=True)
class NumberingEdit:
    """One source-bound paragraph numbering mutation."""

    paragraph: int
    structure_fingerprint_sha256: str
    action: str  # "apply", "restart", or "remove"
    num_id: int | None = None
    abstract_num_id: int | None = None
    level: int = 0
    start_override: int | None = None
    part: str = _DOCUMENT
    restart_group: str | None = None

    def __post_init__(self) -> None:
        if (
            isinstance(self.paragraph, bool)
            or not isinstance(self.paragraph, int)
            or self.paragraph < 1
        ):
            raise ValueError("numbering_paragraph_invalid")
        if (
            not isinstance(self.structure_fingerprint_sha256, str)
            or len(self.structure_fingerprint_sha256) != 64
            or any(
                c not in "0123456789abcdefABCDEF"
                for c in self.structure_fingerprint_sha256
            )
        ):
            raise ValueError("numbering_fingerprint_invalid")
        if not isinstance(self.part, str) or self.part != _DOCUMENT:
            raise ValueError("numbering_part_invalid")
        if self.action not in {"apply", "restart", "remove"}:
            raise ValueError("numbering_action_invalid")
        if (
            isinstance(self.level, bool)
            or not isinstance(self.level, int)
            or self.level < 0
            or self.level > 8
        ):
            raise ValueError("numbering_level_invalid")
        if self.action == "remove":
            if (
                self.num_id is not None
                or self.abstract_num_id is not None
                or self.start_override is not None
                or self.restart_group is not None
            ):
                raise ValueError("numbering_edit_fields_invalid")
            if self.level != 0:
                raise ValueError("numbering_level_invalid")
        elif self.action == "apply":
            if self.num_id is None:
                raise ValueError("numbering_num_missing")
            if (
                self.abstract_num_id is not None
                or self.start_override is not None
                or self.restart_group is not None
            ):
                raise ValueError("numbering_edit_fields_invalid")
        elif self.action == "restart":
            if self.start_override is not None and (
                isinstance(self.start_override, bool)
                or not isinstance(self.start_override, int)
                or self.start_override < 0
            ):
                raise ValueError("numbering_start_override_invalid")
            if self.restart_group is not None and (
                not isinstance(self.restart_group, str) or not self.restart_group
            ):
                raise ValueError("numbering_restart_group_invalid")

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "paragraph": self.paragraph,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "action": self.action,
            "part": self.part,
        }
        if self.action != "remove":
            result["level"] = self.level
        if self.num_id is not None:
            result["num_id"] = self.num_id
        if self.abstract_num_id is not None:
            result["abstract_num_id"] = self.abstract_num_id
        if self.start_override is not None:
            result["start_override"] = self.start_override
        if self.restart_group is not None:
            result["restart_group"] = self.restart_group
        return result


@dataclass(frozen=True)
class NumberingRequest:
    """Strict request for one atomic, verified numbering authoring operation."""

    source: Path
    source_sha256: str
    destination: Path
    edits: tuple[NumberingEdit, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": str(self.source),
            "source_sha256": self.source_sha256,
            "destination": str(self.destination),
            "edits": [edit.to_dict() for edit in self.edits],
        }


@dataclass(frozen=True)
class NumberingReport:
    """Content-safe public result for a numbering authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    numbering: list[dict[str, Any]]
    changed_targets: list[dict[str, Any]]

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "numbering",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "numbering": self.numbering,
            "changed_targets": self.changed_targets,
        }


def _numbering_refusal(request: NumberingRequest, code: str) -> NumberingReport:
    return NumberingReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        numbering=[],
        changed_targets=[],
    )


def _parse_edit(value: Any) -> NumberingEdit:
    if not isinstance(value, Mapping):
        raise TypeError("numbering_edit_invalid")

    action = value.get("action")
    if action not in {"apply", "restart", "remove"}:
        raise ValueError("numbering_action_invalid")

    if action == "apply":
        allowed = {
            "paragraph",
            "structure_fingerprint_sha256",
            "action",
            "num_id",
            "level",
            "part",
        }
        required = {"paragraph", "structure_fingerprint_sha256", "action", "num_id"}
    elif action == "restart":
        allowed = {
            "paragraph",
            "structure_fingerprint_sha256",
            "action",
            "abstract_num_id",
            "num_id",
            "level",
            "start_override",
            "part",
            "restart_group",
        }
        required = {"paragraph", "structure_fingerprint_sha256", "action"}
    else:  # action == "remove"
        allowed = {
            "paragraph",
            "structure_fingerprint_sha256",
            "action",
            "level",
            "part",
        }
        required = {"paragraph", "structure_fingerprint_sha256", "action"}
    present = set(value)
    if not (required <= present <= allowed):
        raise ValueError("numbering_edit_fields_invalid")

    paragraph = _positive(value["paragraph"], "numbering_paragraph")
    fp = _sha(value["structure_fingerprint_sha256"], "numbering_fingerprint")

    part = value.get("part", _DOCUMENT)
    if not isinstance(part, str) or part != _DOCUMENT:
        raise ValueError("numbering_part_invalid")

    level = value.get("level", 0)
    if isinstance(level, bool) or not isinstance(level, int) or level < 0 or level > 8:
        raise ValueError("numbering_level_invalid")
    if action == "remove" and level != 0:
        raise ValueError("numbering_level_invalid")

    num_id: int | None = None
    if "num_id" in value:
        num_id = _positive(value["num_id"], "numbering_num_id")

    abstract_num_id: int | None = None
    if "abstract_num_id" in value:
        raw_aid = value["abstract_num_id"]
        if isinstance(raw_aid, bool) or not isinstance(raw_aid, int) or raw_aid < 0:
            raise ValueError("numbering_abstract_num_id_invalid")
        abstract_num_id = raw_aid

    start_override: int | None = None
    if "start_override" in value:
        start_override = _positive(
            value["start_override"], "numbering_start_override", zero=True
        )

    restart_group: str | None = None
    if "restart_group" in value:
        rg = value["restart_group"]
        if not isinstance(rg, str) or not rg:
            raise ValueError("numbering_restart_group_invalid")
        restart_group = rg

    return NumberingEdit(
        paragraph=paragraph,
        structure_fingerprint_sha256=fp,
        action=action,
        num_id=num_id,
        abstract_num_id=abstract_num_id,
        level=level,
        start_override=start_override,
        part=part,
        restart_group=restart_group,
    )


def parse_numbering_request(value: Any) -> NumberingRequest:
    """Parse the strict JSON request for one atomic numbering authoring operation."""
    if not isinstance(value, Mapping):
        raise TypeError("numbering_request_invalid")

    if set(value) != {"source", "source_sha256", "destination", "edits"}:
        raise ValueError("numbering_request_fields_invalid")

    source = value["source"]
    destination = value["destination"]
    if not isinstance(source, str) or not source:
        raise ValueError("numbering_path_invalid")
    if not isinstance(destination, str) or not destination:
        raise ValueError("numbering_path_invalid")

    source_sha = _sha(value["source_sha256"], "numbering_source_sha256")

    edits_raw = value["edits"]
    if not isinstance(edits_raw, list) or not edits_raw:
        raise ValueError("numbering_edits_invalid")

    parsed_edits = tuple(_parse_edit(item) for item in edits_raw)

    paragraphs = [edit.paragraph for edit in parsed_edits]
    if len(set(paragraphs)) != len(paragraphs):
        raise ValueError("numbering_duplicate_target")

    return NumberingRequest(
        source=Path(source),
        source_sha256=source_sha,
        destination=Path(destination),
        edits=parsed_edits,
    )


def load_numbering_request(path: str | Path) -> NumberingRequest:
    """Load and strictly validate a numbering authoring request from JSON on disk."""
    with Path(path).open(encoding="utf-8") as stream:
        return parse_numbering_request(json.load(stream))


def _unrelated_parts_preserved(
    source: Path, candidate: Path, changed: frozenset[str]
) -> bool:
    """True when every part the request did not change is byte-identical to source."""
    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        names = set(after.namelist())
        return all(
            info.filename in names
            and before.read(info.filename) == after.read(info.filename)
            for info in before.infolist()
            if info.filename not in changed
        )


def _has_unsupported_context(
    root: ET.Element, paragraph: ET.Element, namespace: str
) -> bool:
    unsupported_map = _unsupported_context_by_element(root, namespace)
    if unsupported_map.get(id(paragraph), False):
        return True
    unsupported_qnames = {Q(namespace, tag) for tag in _UNSUPPORTED_PARAGRAPH_TAGS}
    for node in paragraph.iter():
        if node.tag in unsupported_qnames:
            return True
    return False


def _visible_text(element: ET.Element, namespace: str) -> str:
    return "".join(t.text or "" for t in element.iter(Q(namespace, "t")))


def _create_restart_num_element(
    namespace: str,
    num_id: int,
    abstract_num_id: str | int,
    level: int,
    start_override: int,
) -> ET.Element:
    num_elem = ET.Element(Q(namespace, "num"), {Q(namespace, "numId"): str(num_id)})
    ET.SubElement(
        num_elem,
        Q(namespace, "abstractNumId"),
        {Q(namespace, "val"): str(abstract_num_id)},
    )
    lvl_override = ET.SubElement(
        num_elem,
        Q(namespace, "lvlOverride"),
        {Q(namespace, "ilvl"): str(level)},
    )
    ET.SubElement(
        lvl_override,
        Q(namespace, "startOverride"),
        {Q(namespace, "val"): str(start_override)},
    )
    return num_elem


def _insert_num_in_schema_order(
    numbering_root: ET.Element, namespace: str, num_elem: ET.Element
) -> None:
    """Insert a w:num element into w:numbering maintaining schema order.

    ISO/IEC 29500-1 §17.9.15:
    countryRegion?, picture*, abstractNum*, num*, numIdMacAtCleanup?
    """
    cleanup_elem = numbering_root.find(Q(namespace, "numIdMacAtCleanup"))
    if cleanup_elem is not None:
        idx = list(numbering_root).index(cleanup_elem)
        numbering_root.insert(idx, num_elem)
        val_str = cleanup_elem.get(Q(namespace, "val"))
        num_id_val = int(num_elem.get(Q(namespace, "numId"), "0"))
        if val_str and val_str.isdigit() and num_id_val >= int(val_str):
            cleanup_elem.set(Q(namespace, "val"), str(num_id_val + 1))
    else:
        numbering_root.append(num_elem)


def _insert_num_pr_in_schema_order(
    p_pr: ET.Element, namespace: str, num_pr: ET.Element
) -> None:
    """Insert a w:numPr element into w:pPr maintaining schema order."""
    succeeding_qnames = {Q(namespace, tag) for tag in _PPR_SUCCEEDING_TAGS}
    for idx, child in enumerate(p_pr):
        if child.tag in succeeding_qnames:
            p_pr.insert(idx, num_pr)
            return
    p_pr.append(num_pr)


def _apply_num_pr_to_paragraph(
    paragraph: ET.Element,
    namespace: str,
    num_id: int,
    level: int,
) -> None:
    p_pr = paragraph.find(Q(namespace, "pPr"))
    if p_pr is None:
        p_pr = ET.Element(Q(namespace, "pPr"))
        paragraph.insert(0, p_pr)

    num_pr = p_pr.find(Q(namespace, "numPr"))
    if num_pr is not None:
        p_pr.remove(num_pr)

    new_num_pr = ET.Element(Q(namespace, "numPr"))
    ET.SubElement(new_num_pr, Q(namespace, "ilvl"), {Q(namespace, "val"): str(level)})
    ET.SubElement(new_num_pr, Q(namespace, "numId"), {Q(namespace, "val"): str(num_id)})
    _insert_num_pr_in_schema_order(p_pr, namespace, new_num_pr)


def _resolved_paragraph_style_numbering(
    paragraph: ET.Element,
    parts: Mapping[str, bytes],
    namespace: str,
) -> tuple[int, int] | None:
    """Resolve inherited numbering (num_id, ilvl) from word/styles.xml, if any."""
    if _STYLES_PART not in parts:
        return None
    try:
        styles_root = ET.fromstring(parts[_STYLES_PART])
    except ET.ParseError:
        return None
    if styles_root.tag != Q(namespace, "styles"):
        return None

    p_pr = paragraph.find(Q(namespace, "pPr"))
    p_style_elem = p_pr.find(Q(namespace, "pStyle")) if p_pr is not None else None
    declared_style = (
        (p_style_elem.get(Q(namespace, "val")) or p_style_elem.get("val"))
        if p_style_elem is not None
        else None
    )

    styles_by_id: dict[str, ET.Element] = {}
    default_style_id: str | None = None

    for style_elem in styles_root.findall(Q(namespace, "style")):
        sid = style_elem.get(Q(namespace, "styleId")) or style_elem.get("styleId")
        stype = style_elem.get(Q(namespace, "type")) or style_elem.get("type")
        if stype == "paragraph" and sid:
            styles_by_id[sid] = style_elem
            is_default = style_elem.get(Q(namespace, "default")) or style_elem.get(
                "default"
            )
            if is_default in ("1", "true"):
                default_style_id = sid

    target_style = declared_style or default_style_id
    if not target_style:
        doc_defaults = styles_root.find(Q(namespace, "docDefaults"))
        if doc_defaults is not None:
            dd_p_pr = doc_defaults.find(
                f"{Q(namespace, 'pPrDefault')}/{Q(namespace, 'pPr')}"
            )
            if dd_p_pr is not None:
                dd_num_pr = dd_p_pr.find(Q(namespace, "numPr"))
                if dd_num_pr is not None:
                    nid = dd_num_pr.find(Q(namespace, "numId"))
                    val = (
                        (nid.get(Q(namespace, "val")) or nid.get("val"))
                        if nid is not None
                        else None
                    )
                    if val and val.isdigit():
                        nid_int = int(val)
                        if nid_int > 0:
                            ilvl = dd_num_pr.find(Q(namespace, "ilvl"))
                            ilvl_val = (
                                (ilvl.get(Q(namespace, "val")) or ilvl.get("val"))
                                if ilvl is not None
                                else None
                            )
                            ilvl_int = int(ilvl_val) if ilvl_val and ilvl_val.isdigit() else 0
                            return (nid_int, ilvl_int)
        return None

    seen: set[str] = set()
    curr: str | None = target_style
    while curr and curr in styles_by_id and curr not in seen:
        seen.add(curr)
        style_elem = styles_by_id[curr]
        s_p_pr = style_elem.find(Q(namespace, "pPr"))
        if s_p_pr is not None:
            num_pr = s_p_pr.find(Q(namespace, "numPr"))
            if num_pr is not None:
                nid = num_pr.find(Q(namespace, "numId"))
                val = (
                    (nid.get(Q(namespace, "val")) or nid.get("val"))
                    if nid is not None
                    else None
                )
                if val and val.isdigit():
                    nid_int = int(val)
                    if nid_int == 0:
                        return (0, 0)
                    ilvl = num_pr.find(Q(namespace, "ilvl"))
                    ilvl_val = (
                        (ilvl.get(Q(namespace, "val")) or ilvl.get("val"))
                        if ilvl is not None
                        else None
                    )
                    ilvl_int = int(ilvl_val) if ilvl_val and ilvl_val.isdigit() else 0
                    return (nid_int, ilvl_int)
        based_on = style_elem.find(Q(namespace, "basedOn"))
        curr = (
            (based_on.get(Q(namespace, "val")) or based_on.get("val"))
            if based_on is not None
            else None
        )

    doc_defaults = styles_root.find(Q(namespace, "docDefaults"))
    if doc_defaults is not None:
        dd_p_pr = doc_defaults.find(
            f"{Q(namespace, 'pPrDefault')}/{Q(namespace, 'pPr')}"
        )
        if dd_p_pr is not None:
            dd_num_pr = dd_p_pr.find(Q(namespace, "numPr"))
            if dd_num_pr is not None:
                nid = dd_num_pr.find(Q(namespace, "numId"))
                val = (
                    (nid.get(Q(namespace, "val")) or nid.get("val"))
                    if nid is not None
                    else None
                )
                if val and val.isdigit():
                    nid_int = int(val)
                    if nid_int > 0:
                        ilvl = dd_num_pr.find(Q(namespace, "ilvl"))
                        ilvl_val = (
                            (ilvl.get(Q(namespace, "val")) or ilvl.get("val"))
                            if ilvl is not None
                            else None
                        )
                        ilvl_int = int(ilvl_val) if ilvl_val and ilvl_val.isdigit() else 0
                        return (nid_int, ilvl_int)
    return None


def _paragraph_has_inherited_numbering(
    paragraph: ET.Element,
    parts: Mapping[str, bytes],
    namespace: str,
) -> bool:
    res = _resolved_paragraph_style_numbering(paragraph, parts, namespace)
    return res is not None and res[0] > 0


def _paragraph_effective_num_id(
    paragraph: ET.Element,
    parts: Mapping[str, bytes],
    namespace: str,
) -> int | None:
    """Resolve the active numbering instance id for a paragraph, if numbered.

    A direct w:numPr/w:numId wins; otherwise the style chain and document
    defaults decide. numId 0 (explicit suppression) means not numbered.
    """
    p_pr = paragraph.find(Q(namespace, "pPr"))
    num_pr = p_pr.find(Q(namespace, "numPr")) if p_pr is not None else None
    nid_elem = num_pr.find(Q(namespace, "numId")) if num_pr is not None else None
    if nid_elem is not None:
        val = nid_elem.get(Q(namespace, "val"))
        num_id = int(val) if val and val.isdigit() else 0
    else:
        inherited = _resolved_paragraph_style_numbering(paragraph, parts, namespace)
        num_id = inherited[0] if inherited is not None else 0
    return num_id if num_id > 0 else None


def _remove_num_pr_from_paragraph(
    paragraph: ET.Element,
    namespace: str,
    suppress: bool,
) -> None:
    p_pr = paragraph.find(Q(namespace, "pPr"))
    if p_pr is None:
        if not suppress:
            return
        p_pr = ET.Element(Q(namespace, "pPr"))
        paragraph.insert(0, p_pr)

    num_pr = p_pr.find(Q(namespace, "numPr"))
    if num_pr is not None:
        p_pr.remove(num_pr)

    if suppress:
        new_num_pr = ET.Element(Q(namespace, "numPr"))
        ET.SubElement(new_num_pr, Q(namespace, "numId"), {Q(namespace, "val"): "0"})
        _insert_num_pr_in_schema_order(p_pr, namespace, new_num_pr)
    else:
        if len(p_pr) == 0:
            paragraph.remove(p_pr)


def _ensure_numbering_registration(
    parts: dict[str, bytes],
    replacements: dict[str, bytes],
    namespace: str,
) -> None:
    relationship_type = (
        _STRICT_NUMBERING_REL_TYPE
        if namespace == STRICT_WORD_NS
        else _NUMBERING_REL_TYPE
    )
    ct_data = replacements.get("[Content_Types].xml", parts.get("[Content_Types].xml"))
    if ct_data is not None:
        ct_root = parse_xml_with_preserved_qnames(ct_data)
        has_override = any(
            node.get("PartName") == "/word/numbering.xml"
            for node in ct_root.findall(Q(_CT, "Override"))
        )
        if not has_override:
            ct_root.append(
                ET.Element(
                    Q(_CT, "Override"),
                    {
                        "PartName": "/word/numbering.xml",
                        "ContentType": _NUMBERING_CONTENT_TYPE,
                    },
                )
            )
            replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(ct_root)

    rels_name = "word/_rels/document.xml.rels"
    rels_data = replacements.get(rels_name, parts.get(rels_name))
    if rels_data is not None:
        rels_root = parse_xml_with_preserved_qnames(rels_data)
        has_rel = any(
            rel.get("Type") == relationship_type
            for rel in rels_root.findall(Q(_PR, "Relationship"))
        )
        if not has_rel:
            existing_ids = [
                rel.get("Id", "") for rel in rels_root.findall(Q(_PR, "Relationship"))
            ]
            rid_nums = [
                int(rid[3:])
                for rid in existing_ids
                if rid.startswith("rId") and rid[3:].isdigit()
            ]
            next_rid = f"rId{(max(rid_nums) if rid_nums else 0) + 1}"
            rels_root.append(
                ET.Element(
                    Q(_PR, "Relationship"),
                    {
                        "Id": next_rid,
                        "Type": relationship_type,
                        "Target": "numbering.xml",
                    },
                )
            )
            replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)


def discover_numbering_targets(
    source: str | Path, *, part: str = _DOCUMENT
) -> dict[str, Any]:
    """List source-bound paragraph targets, existing abstract definitions, and instances."""
    path = Path(source)
    try:
        graph = PackageGraph.from_file(path)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return {
            "passed": False,
            "code": "numbering_source_package_invalid",
            "source_sha256": "",
            "part": part,
            "numbering_part_exists": False,
            "abstract_nums": [],
            "nums": [],
            "targets": [],
            "unsupported": [],
        }

    if part not in parts:
        return {
            "passed": False,
            "code": "numbering_source_package_invalid",
            "source_sha256": graph.file_sha256,
            "part": part,
            "numbering_part_exists": False,
            "abstract_nums": [],
            "nums": [],
            "targets": [],
            "unsupported": [],
        }

    try:
        document_root = ET.fromstring(parts[part])
    except ET.ParseError:
        return {
            "passed": False,
            "code": "numbering_source_package_invalid",
            "source_sha256": graph.file_sha256,
            "part": part,
            "numbering_part_exists": False,
            "abstract_nums": [],
            "nums": [],
            "targets": [],
            "unsupported": [],
        }

    namespace = _word_namespace(document_root)
    if namespace is None:
        return {
            "passed": False,
            "code": "numbering_namespace_unsupported",
            "source_sha256": graph.file_sha256,
            "part": part,
            "numbering_part_exists": False,
            "abstract_nums": [],
            "nums": [],
            "targets": [],
            "unsupported": [],
        }

    abstract_nums: list[dict[str, Any]] = []
    nums: list[dict[str, Any]] = []
    numbering_part_exists = _NUMBERING_PART in parts

    if numbering_part_exists:
        try:
            num_root = ET.fromstring(parts[_NUMBERING_PART])
            if num_root.tag != Q(namespace, "numbering"):
                return {
                    "passed": False,
                    "code": "numbering_part_malformed",
                    "source_sha256": graph.file_sha256,
                    "part": part,
                    "numbering_part_exists": True,
                    "abstract_nums": [],
                    "nums": [],
                    "targets": [],
                    "unsupported": [],
                }

            for abstract in num_root.findall(Q(namespace, "abstractNum")):
                aid_str = abstract.get(Q(namespace, "abstractNumId"))
                if not aid_str or not aid_str.isdigit():
                    continue
                aid = int(aid_str)
                has_style = (
                    abstract.find(Q(namespace, "styleLink")) is not None
                    or abstract.find(Q(namespace, "numStyleLink")) is not None
                )
                has_pic = any(
                    lvl.find(Q(namespace, "lvlPicBulletId")) is not None
                    for lvl in abstract.findall(Q(namespace, "lvl"))
                )
                levels: list[dict[str, Any]] = []
                for lvl in abstract.findall(Q(namespace, "lvl")):
                    ilvl_str = lvl.get(Q(namespace, "ilvl"))
                    if not ilvl_str or not ilvl_str.isdigit():
                        continue
                    ilvl = int(ilvl_str)
                    fmt_elem = lvl.find(Q(namespace, "numFmt"))
                    num_fmt = (
                        fmt_elem.get(Q(namespace, "val"))
                        if fmt_elem is not None
                        else "decimal"
                    )
                    txt_elem = lvl.find(Q(namespace, "lvlText"))
                    lvl_text = (
                        txt_elem.get(Q(namespace, "val"))
                        if txt_elem is not None
                        else ""
                    )
                    st_elem = lvl.find(Q(namespace, "start"))
                    st_str = (
                        st_elem.get(Q(namespace, "val")) if st_elem is not None else "1"
                    )
                    start_val = int(st_str) if st_str and st_str.isdigit() else 1
                    lvl_pic = lvl.find(Q(namespace, "lvlPicBulletId")) is not None
                    levels.append(
                        {
                            "ilvl": ilvl,
                            "num_fmt": num_fmt or "decimal",
                            "lvl_text": lvl_text or "",
                            "start": start_val,
                            "has_picture_bullet": lvl_pic,
                        }
                    )
                abstract_nums.append(
                    {
                        "abstract_num_id": aid,
                        "has_style_link": has_style,
                        "has_picture_bullet": has_pic,
                        "levels": levels,
                    }
                )

            for num in num_root.findall(Q(namespace, "num")):
                nid_str = num.get(Q(namespace, "numId"))
                if not nid_str or not nid_str.isdigit():
                    continue
                nid = int(nid_str)
                ab_ref = num.find(Q(namespace, "abstractNumId"))
                ab_val_str = (
                    ab_ref.get(Q(namespace, "val")) if ab_ref is not None else None
                )
                ab_val = (
                    int(ab_val_str) if ab_val_str and ab_val_str.isdigit() else None
                )
                overrides: list[dict[str, Any]] = []
                for ov in num.findall(Q(namespace, "lvlOverride")):
                    ilvl_str = ov.get(Q(namespace, "ilvl"))
                    if not ilvl_str or not ilvl_str.isdigit():
                        continue
                    ov_start = ov.find(Q(namespace, "startOverride"))
                    ov_val_str = (
                        ov_start.get(Q(namespace, "val"))
                        if ov_start is not None
                        else None
                    )
                    overrides.append(
                        {
                            "ilvl": int(ilvl_str),
                            "start_override": (
                                int(ov_val_str)
                                if ov_val_str and ov_val_str.isdigit()
                                else None
                            ),
                        }
                    )
                nums.append(
                    {
                        "num_id": nid,
                        "abstract_num_id": ab_val,
                        "level_overrides": overrides,
                    }
                )
        except ET.ParseError:
            return {
                "passed": False,
                "code": "numbering_part_malformed",
                "source_sha256": graph.file_sha256,
                "part": part,
                "numbering_part_exists": True,
                "abstract_nums": [],
                "nums": [],
                "targets": [],
                "unsupported": [],
            }

    targets: list[dict[str, Any]] = []
    unsupported: list[dict[str, Any]] = []

    for index, paragraph in enumerate(document_root.iter(Q(namespace, "p")), start=1):
        fp = _mask_text(paragraph, namespace)
        if _has_unsupported_context(document_root, paragraph, namespace):
            unsupported.append(
                {
                    "paragraph": index,
                    "code": "numbering_context_unsupported",
                    "structure_fingerprint_sha256": fp,
                }
            )
            continue

        p_pr = paragraph.find(Q(namespace, "pPr"))
        num_pr = p_pr.find(Q(namespace, "numPr")) if p_pr is not None else None
        num_id_val: int | None = None
        ilvl_val: int | None = None
        if num_pr is not None:
            nid_elem = num_pr.find(Q(namespace, "numId"))
            nid_str = (
                nid_elem.get(Q(namespace, "val")) if nid_elem is not None else None
            )
            if nid_str and nid_str.isdigit():
                num_id_val = int(nid_str)
            ilvl_elem = num_pr.find(Q(namespace, "ilvl"))
            ilvl_str = (
                ilvl_elem.get(Q(namespace, "val")) if ilvl_elem is not None else None
            )
            if ilvl_str and ilvl_str.isdigit():
                ilvl_val = int(ilvl_str)
            elif num_id_val is not None:
                ilvl_val = 0
        else:
            inherited = _resolved_paragraph_style_numbering(paragraph, parts, namespace)
            if inherited is not None:
                num_id_val, ilvl_val = inherited

        targets.append(
            {
                "paragraph": index,
                "structure_fingerprint_sha256": fp,
                "text": _visible_text(paragraph, namespace),
                "num_id": num_id_val,
                "level": ilvl_val,
                "is_numbered": num_id_val is not None and num_id_val > 0,
            }
        )

    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        "part": part,
        "numbering_part_exists": numbering_part_exists,
        "abstract_nums": abstract_nums,
        "nums": nums,
        "targets": targets,
        "unsupported": unsupported,
    }


def author_numbering(request: NumberingRequest) -> NumberingReport:
    """Execute source-bound paragraph numbering mutations atomically."""
    source = request.source
    destination = request.destination

    if not source.is_file():
        return _numbering_refusal(request, "numbering_source_package_invalid")

    source_bytes = source.read_bytes()
    if hashlib.sha256(source_bytes).hexdigest() != request.source_sha256:
        return _numbering_refusal(request, "numbering_source_sha256_mismatch")

    if destination.exists() or destination.is_symlink():
        return _numbering_refusal(request, "numbering_destination_invalid")

    if not destination.parent.is_dir():
        return _numbering_refusal(request, "numbering_destination_invalid")

    if not request.edits:
        return _numbering_refusal(request, "numbering_edits_empty")

    try:
        graph = PackageGraph.from_file(source)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return _numbering_refusal(request, "numbering_source_package_invalid")

    if _DOCUMENT not in parts:
        return _numbering_refusal(request, "numbering_source_package_invalid")

    try:
        document_root = parse_xml_with_preserved_qnames(parts[_DOCUMENT])
    except ET.ParseError:
        return _numbering_refusal(request, "numbering_source_package_invalid")

    namespace = _word_namespace(document_root)
    if namespace is None:
        return _numbering_refusal(request, "numbering_namespace_unsupported")

    if _NUMBERING_PART not in parts:
        return _numbering_refusal(request, "numbering_part_missing")

    try:
        numbering_root = parse_xml_with_preserved_qnames(parts[_NUMBERING_PART])
    except ET.ParseError:
        return _numbering_refusal(request, "numbering_part_malformed")

    if numbering_root.tag != Q(namespace, "numbering"):
        return _numbering_refusal(request, "numbering_part_malformed")

    abstract_nums: dict[str, ET.Element] = {}
    for abstract in numbering_root.findall(Q(namespace, "abstractNum")):
        aid = abstract.get(Q(namespace, "abstractNumId"))
        if aid:
            abstract_nums[aid] = abstract

    nums: dict[str, ET.Element] = {}
    for num in numbering_root.findall(Q(namespace, "num")):
        nid = num.get(Q(namespace, "numId"))
        if nid:
            nums[nid] = num

    existing_num_ids: set[int] = set()
    for nid_str in nums:
        if nid_str.isdigit():
            existing_num_ids.add(int(nid_str))
    for p in document_root.iter(Q(namespace, "p")):
        num_pr = p.find(f"{Q(namespace, 'pPr')}/{Q(namespace, 'numPr')}")
        if num_pr is not None:
            nid_elem = num_pr.find(Q(namespace, "numId"))
            val = nid_elem.get(Q(namespace, "val")) if nid_elem is not None else None
            if val and val.isdigit():
                existing_num_ids.add(int(val))

    next_num_id = (max(existing_num_ids) if existing_num_ids else 0) + 1
    paragraphs = list(document_root.iter(Q(namespace, "p")))
    restart_abstract_ids: dict[int, str] = {}

    for index, edit in enumerate(request.edits):
        if edit.part != _DOCUMENT:
            return _numbering_refusal(request, "numbering_part_invalid")

        if edit.paragraph < 1 or edit.paragraph > len(paragraphs):
            return _numbering_refusal(request, "numbering_target_missing")

        target_p = paragraphs[edit.paragraph - 1]
        fp = _mask_text(target_p, namespace)
        if fp != edit.structure_fingerprint_sha256:
            return _numbering_refusal(request, "numbering_target_stale")

        if _has_unsupported_context(document_root, target_p, namespace):
            return _numbering_refusal(request, "numbering_context_unsupported")

        if edit.action == "apply":
            if edit.num_id is None:
                return _numbering_refusal(request, "numbering_num_missing")
            num_elem = nums.get(str(edit.num_id))
            if num_elem is None:
                return _numbering_refusal(request, "numbering_num_missing")
            ab_ref = num_elem.find(Q(namespace, "abstractNumId"))
            if ab_ref is None:
                return _numbering_refusal(request, "numbering_abstract_num_missing")
            ab_id = ab_ref.get(Q(namespace, "val"))
            if not ab_id or ab_id not in abstract_nums:
                return _numbering_refusal(request, "numbering_abstract_num_missing")
            abstract_elem = abstract_nums[ab_id]

            if (
                abstract_elem.find(Q(namespace, "styleLink")) is not None
                or abstract_elem.find(Q(namespace, "numStyleLink")) is not None
            ):
                return _numbering_refusal(request, "numbering_style_link_unsupported")

            if any(
                lvl.find(Q(namespace, "lvlPicBulletId")) is not None
                for lvl in abstract_elem.findall(Q(namespace, "lvl"))
            ):
                return _numbering_refusal(
                    request, "numbering_picture_bullet_unsupported"
                )

            lvl = next(
                (
                    level
                    for level in abstract_elem.findall(Q(namespace, "lvl"))
                    if level.get(Q(namespace, "ilvl")) == str(edit.level)
                ),
                None,
            )
            if lvl is None:
                return _numbering_refusal(request, "numbering_level_invalid")

            fmt_elem = lvl.find(Q(namespace, "numFmt"))
            fmt_val = (
                fmt_elem.get(Q(namespace, "val")) if fmt_elem is not None else "decimal"
            )
            if fmt_val not in _NUMBER_FORMATS:
                return _numbering_refusal(request, "numbering_definition_unsupported")

            txt_elem = lvl.find(Q(namespace, "lvlText"))
            if txt_elem is None or txt_elem.get(Q(namespace, "val")) is None:
                return _numbering_refusal(request, "numbering_definition_unsupported")

        elif edit.action == "restart":
            resolved_ab_id: str | None = None
            source_num_id = edit.num_id
            if edit.abstract_num_id is not None:
                resolved_ab_id = str(edit.abstract_num_id)
                if resolved_ab_id not in abstract_nums:
                    return _numbering_refusal(request, "numbering_abstract_num_missing")
            elif source_num_id is None:
                source_num_id = _paragraph_effective_num_id(target_p, parts, namespace)
                if source_num_id is None:
                    return _numbering_refusal(request, "numbering_num_missing")
            if source_num_id is not None:
                num_elem = nums.get(str(source_num_id))
                if num_elem is None:
                    return _numbering_refusal(request, "numbering_num_missing")
                ab_ref = num_elem.find(Q(namespace, "abstractNumId"))
                if ab_ref is None:
                    return _numbering_refusal(request, "numbering_abstract_num_missing")
                num_ab_id = ab_ref.get(Q(namespace, "val"))
                if resolved_ab_id is not None and num_ab_id != resolved_ab_id:
                    return _numbering_refusal(
                        request, "numbering_abstract_num_mismatch"
                    )
                resolved_ab_id = num_ab_id

            if not resolved_ab_id or resolved_ab_id not in abstract_nums:
                return _numbering_refusal(request, "numbering_abstract_num_missing")

            restart_abstract_ids[index] = resolved_ab_id
            abstract_elem = abstract_nums[resolved_ab_id]

            if (
                abstract_elem.find(Q(namespace, "styleLink")) is not None
                or abstract_elem.find(Q(namespace, "numStyleLink")) is not None
            ):
                return _numbering_refusal(request, "numbering_style_link_unsupported")

            if any(
                lvl.find(Q(namespace, "lvlPicBulletId")) is not None
                for lvl in abstract_elem.findall(Q(namespace, "lvl"))
            ):
                return _numbering_refusal(
                    request, "numbering_picture_bullet_unsupported"
                )

            lvl = next(
                (
                    level
                    for level in abstract_elem.findall(Q(namespace, "lvl"))
                    if level.get(Q(namespace, "ilvl")) == str(edit.level)
                ),
                None,
            )
            if lvl is None:
                return _numbering_refusal(request, "numbering_level_invalid")

            fmt_elem = lvl.find(Q(namespace, "numFmt"))
            fmt_val = (
                fmt_elem.get(Q(namespace, "val")) if fmt_elem is not None else "decimal"
            )
            if fmt_val not in _NUMBER_FORMATS:
                return _numbering_refusal(request, "numbering_definition_unsupported")

            txt_elem = lvl.find(Q(namespace, "lvlText"))
            if txt_elem is None or txt_elem.get(Q(namespace, "val")) is None:
                return _numbering_refusal(request, "numbering_definition_unsupported")

            start_val = edit.start_override if edit.start_override is not None else 1
            if start_val < 0:
                return _numbering_refusal(request, "numbering_start_override_invalid")
        elif edit.action == "remove":
            if (
                edit.num_id is not None
                or edit.abstract_num_id is not None
                or edit.start_override is not None
                or edit.restart_group is not None
                or edit.level != 0
            ):
                return _numbering_refusal(request, "numbering_edit_fields_invalid")
        else:
            return _numbering_refusal(request, "numbering_action_invalid")
    mod_doc_root = copy.deepcopy(document_root)
    mod_num_root = copy.deepcopy(numbering_root)
    mod_paragraphs = list(mod_doc_root.iter(Q(namespace, "p")))

    allocated_num_map: dict[Any, int] = {}
    allocated_restarts: dict[int, tuple[int, int, int]] = {}
    numbering_records: list[dict[str, Any]] = []
    changed_targets: list[dict[str, Any]] = []
    numbering_modified = False

    for idx, edit in enumerate(request.edits):
        target_p = mod_paragraphs[edit.paragraph - 1]
        if edit.action == "apply":
            assert edit.num_id is not None
            effective_num_id = edit.num_id
            num_elem = nums[str(effective_num_id)]
            ab_ref = num_elem.find(Q(namespace, "abstractNumId"))
            ab_id_int = (
                int(ab_ref.get(Q(namespace, "val"), "0")) if ab_ref is not None else 0
            )
            numbering_records.append(
                {
                    "num_id": effective_num_id,
                    "abstract_num_id": ab_id_int,
                    "action": "apply",
                    "level": edit.level,
                    "start_override": None,
                }
            )
            _apply_num_pr_to_paragraph(
                target_p, namespace, effective_num_id, edit.level
            )
            changed_targets.append(
                {
                    "paragraph": edit.paragraph,
                    "action": edit.action,
                    "num_id": effective_num_id,
                    "level": edit.level,
                    "part": edit.part,
                }
            )
        elif edit.action == "restart":
            resolved_ab_id = restart_abstract_ids[idx]
            start_val = edit.start_override if edit.start_override is not None else 1
            group_key = edit.restart_group if edit.restart_group is not None else idx
            if group_key in allocated_num_map:
                effective_num_id = allocated_num_map[group_key]
            else:
                effective_num_id = next_num_id
                next_num_id += 1
                allocated_num_map[group_key] = effective_num_id
                new_num_elem = _create_restart_num_element(
                    namespace,
                    effective_num_id,
                    resolved_ab_id,
                    edit.level,
                    start_val,
                )
                _insert_num_in_schema_order(mod_num_root, namespace, new_num_elem)
                allocated_restarts[effective_num_id] = (
                    int(resolved_ab_id),
                    edit.level,
                    start_val,
                )
                numbering_modified = True

            numbering_records.append(
                {
                    "num_id": effective_num_id,
                    "abstract_num_id": int(resolved_ab_id),
                    "action": "restart",
                    "level": edit.level,
                    "start_override": start_val,
                }
            )
            _apply_num_pr_to_paragraph(
                target_p, namespace, effective_num_id, edit.level
            )
            changed_targets.append(
                {
                    "paragraph": edit.paragraph,
                    "action": edit.action,
                    "num_id": effective_num_id,
                    "level": edit.level,
                    "part": edit.part,
                }
            )
        else:  # remove
            orig_p = paragraphs[edit.paragraph - 1]
            has_inherited = _paragraph_has_inherited_numbering(orig_p, parts, namespace)
            _remove_num_pr_from_paragraph(target_p, namespace, suppress=has_inherited)
            eff_nid: int | None = 0 if has_inherited else None
            eff_lvl: int | None = 0 if has_inherited else None
            numbering_records.append(
                {
                    "num_id": eff_nid,
                    "abstract_num_id": None,
                    "action": "remove",
                    "level": eff_lvl,
                    "start_override": None,
                }
            )
            changed_targets.append(
                {
                    "paragraph": edit.paragraph,
                    "action": "remove",
                    "num_id": eff_nid,
                    "level": eff_lvl,
                    "part": edit.part,
                }
            )

    mod_doc_bytes = serialize_xml_with_preserved_qnames(mod_doc_root)
    replacements: dict[str, bytes] = {_DOCUMENT: mod_doc_bytes}

    if numbering_modified:
        mod_num_bytes = serialize_xml_with_preserved_qnames(mod_num_root)
        replacements[_NUMBERING_PART] = mod_num_bytes

    _ensure_numbering_registration(parts, replacements, namespace)
    all_changed = frozenset(replacements)

    def verifier(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, all_changed):
            return "numbering_staged_verification_failed"
        disc = discover_numbering_targets(staged)
        if not disc["passed"]:
            return "numbering_staged_verification_failed"
        disc_targets = {t["paragraph"]: t for t in disc["targets"]}
        for ct in changed_targets:
            p_info = disc_targets.get(ct["paragraph"])
            if (
                p_info is None
                or p_info["num_id"] != ct["num_id"]
                or p_info["level"] != ct["level"]
            ):
                return "numbering_staged_verification_failed"
        disc_nums = {n["num_id"]: n for n in disc["nums"]}
        for num_id, (ab_id, level, start) in allocated_restarts.items():
            staged_num = disc_nums.get(num_id)
            if (
                staged_num is None
                or staged_num["abstract_num_id"] != ab_id
                or {"ilvl": level, "start_override": start}
                not in staged_num["level_overrides"]
            ):
                return "numbering_staged_verification_failed"
        return None

    published = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=changed_targets,
        additions={},
        verify=verifier,
        failure_code="numbering_staged_verification_failed",
        exists_code="numbering_destination_invalid",
        unavailable_code="numbering_publication_unavailable",
    )

    if published.status != _PUBLISHED:
        return NumberingReport(
            status=published.status,
            code=published.code,
            source_sha256=published.source_sha256,
            destination_sha256=published.output_sha256,
            numbering=[],
            changed_targets=[],
        )

    return NumberingReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=published.source_sha256,
        destination_sha256=published.output_sha256,
        numbering=numbering_records,
        changed_targets=changed_targets,
    )
