"""Source-bound DOCX body table deletion and target discovery.

Safely removes a targeted top-level body table from a WordprocessingML document
while guaranteeing preservation of surrounding body paragraphs, document section
properties, and unrelated package parts. Refuses stale, missing, or nested targets,
section or relationship hazards, and annotation or bookmark ranges crossing or
intersecting the removed table where safety cannot be established.
"""

from __future__ import annotations

import json
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .authoring import (
    _PUBLISHED,
    _REFUSED,
    _cell_structure_fingerprint,
    _require_sha256,
    _sha256_file,
    _unsupported_context_by_element,
    _word_namespace,
)
from .inspect import (
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import _DOCUMENT, ReviewReport, _publish

__all__ = [
    "TableDeleteReport",
    "TableDeleteRequest",
    "TableDeleteTarget",
    "author_table_delete",
    "discover_table_delete_targets",
    "load_table_delete_request",
    "parse_table_delete_request",
    "table_delete_structure_fingerprint",
]

_REQUIRED_REQUEST_FIELDS = frozenset(
    {"source", "destination", "source_sha256", "target", "structure_fingerprint_sha256"}
)
_ALLOWED_REQUEST_FIELDS = frozenset(
    {"source", "destination", "source_sha256", "target", "structure_fingerprint_sha256"}
)
_REQUIRED_TARGET_FIELDS = frozenset({"table"})
_ALLOWED_TARGET_FIELDS = frozenset({"table", "part"})

_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_RELS_DOCUMENT = "word/_rels/document.xml.rels"


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is byte-identical to source."""
    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        names = set(after.namelist())
        return all(
            info.filename in names and before.read(info.filename) == after.read(info.filename)
            for info in before.infolist()
            if info.filename not in changed
        )


def _is_supported_table(table: ET.Element, namespace: str) -> bool:
    """True if table has supported body table structure and at least one row."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("tblPr"), word("tblGrid"), word("tr")} for child in table):
        return False
    rows = table.findall(word("tr"))
    return len(rows) > 0


def _table_structure_fingerprint(table: ET.Element) -> str:
    """Hash the table's XML structure while deliberately masking cell text."""
    return _cell_structure_fingerprint(table)


def _validate_table_delete_target_dict(target: Any) -> dict[str, Any]:
    if not isinstance(target, Mapping):
        raise TypeError("table_delete_target_invalid")
    keys = set(target)
    if not keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(keys):
        raise ValueError("table_delete_target_fields_invalid")
    table = target["table"]
    if isinstance(table, bool) or not isinstance(table, int) or table < 1:
        raise ValueError("table_delete_table_invalid")
    part = target.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("table_delete_part_invalid")
    return {"table": table, "part": part}


@dataclass(frozen=True)
class TableDeleteTarget:
    """One discovered candidate table target for deletion."""

    table: int
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    rows_count: int = 0
    columns_count: int = 0
    style: str | None = None
    preview_text: str = ""

    @property
    def target(self) -> dict[str, Any]:
        return {"table": self.table, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "part": self.part,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "rows_count": self.rows_count,
            "columns_count": self.columns_count,
            "style": self.style,
            "preview_text": self.preview_text,
        }


@dataclass(frozen=True)
class TableDeleteRequest:
    """One source-bound request to delete an exact body table."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str

    def __post_init__(self) -> None:
        if isinstance(self.source, str):
            object.__setattr__(self, "source", Path(self.source))
        elif not isinstance(self.source, Path):
            raise TypeError("table_delete_source_invalid")
        if not str(self.source):
            raise ValueError("table_delete_source_invalid")

        if isinstance(self.destination, str):
            object.__setattr__(self, "destination", Path(self.destination))
        elif not isinstance(self.destination, Path):
            raise TypeError("table_delete_destination_invalid")
        if not str(self.destination):
            raise ValueError("table_delete_destination_invalid")

        _require_sha256(self.source_sha256, "table_delete_source_sha256")
        _require_sha256(
            self.structure_fingerprint_sha256,
            "table_delete_structure_fingerprint_sha256",
        )

        validated_target = _validate_table_delete_target_dict(self.target)
        object.__setattr__(self, "target", validated_target)

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        return {"table": self.table, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "target": self.target_dict(),
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
        }


@dataclass(frozen=True)
class TableDeleteReport:
    """Content-safe public result for a table delete authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    deleted_table: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "table-delete",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "deleted_table": self.deleted_table,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _table_delete_refusal(
    request: TableDeleteRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TableDeleteReport:
    return TableDeleteReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        deleted_table=None,
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_delete_request(value: Any) -> TableDeleteRequest:
    """Parse and strictly validate a table delete authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_delete_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_delete_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("table_delete_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("table_delete_destination_invalid")

    source_sha256 = _require_sha256(value["source_sha256"], "table_delete_source_sha256")
    structure_fingerprint = _require_sha256(
        value["structure_fingerprint_sha256"],
        "table_delete_structure_fingerprint_sha256",
    )

    target = _validate_table_delete_target_dict(value["target"])

    return TableDeleteRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        target=target,
        structure_fingerprint_sha256=structure_fingerprint,
    )


def load_table_delete_request(path: str | Path) -> TableDeleteRequest:
    """Load and strictly validate one JSON table-delete request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_delete_request(json.load(stream))


def table_delete_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one target body table."""
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
        tables = list(root.iter(Q(ns, "tbl")))
        if table < 1 or table > len(tables):
            return None
        target_table = tables[table - 1]
        body = root.find(Q(ns, "body"))
        if body is None:
            return None
        parent_map = {c: p for p in root.iter() for c in p}
        if parent_map.get(target_table) is not body:
            return None
        if target_table.find(f".//{Q(ns, 'tbl')}") is not None:
            return None
        if not _is_supported_table(target_table, ns):
            return None
        return _table_structure_fingerprint(target_table)
    except (PackageError, OSError, ET.ParseError, KeyError):
        return None


def _check_nested_hazard(
    root: ET.Element,
    target_table: ET.Element,
    body: ET.Element,
    ns: str,
) -> str | None:
    """Refuse nested table targets or targets containing nested tables."""
    parent_map = {c: p for p in root.iter() for c in p}
    if parent_map.get(target_table) is not body:
        return "table_delete_target_nested"

    curr = parent_map.get(target_table)
    while curr is not None:
        if curr.tag == Q(ns, "tbl"):
            return "table_delete_target_nested"
        curr = parent_map.get(curr)

    if target_table.find(f".//{Q(ns, 'tbl')}") is not None:
        return "table_delete_target_nested"

    return None


def _check_section_hazard(
    target_table: ET.Element,
    body: ET.Element,
    ns: str,
) -> str | None:
    """Refuse targets whose deletion introduces section layout or validity hazards."""
    if target_table.find(f".//{Q(ns, 'sectPr')}") is not None:
        return "table_delete_section_hazard"
    if target_table.find(f".//{Q(ns, 'sectPrChange')}") is not None:
        return "table_delete_section_hazard"

    remaining_paragraphs = [c for c in body if c is not target_table and c.tag == Q(ns, "p")]
    if not remaining_paragraphs:
        return "table_delete_section_hazard"

    return None


def _check_annotation_and_bookmark_hazards(
    root: ET.Element,
    target_table: ET.Element,
    ns: str,
) -> str | None:
    """Refuse annotation or bookmark ranges crossing or intersecting the removed table."""
    table_elems = set(target_table.iter())
    doc_elems = list(root.iter())
    elem_indices = {id(el): idx for idx, el in enumerate(doc_elems)}

    tbl_start_idx = elem_indices[id(target_table)]
    tbl_end_idx = max(elem_indices[id(el)] for el in table_elems)

    # 1. Bookmarks
    bm_starts: dict[str, ET.Element] = {}
    bm_ends: dict[str, ET.Element] = {}
    for el in doc_elems:
        if el.tag == Q(ns, "bookmarkStart"):
            bm_id = el.get(Q(ns, "id")) or el.get("id")
            if bm_id is not None:
                bm_starts[bm_id] = el
        elif el.tag == Q(ns, "bookmarkEnd"):
            bm_id = el.get(Q(ns, "id")) or el.get("id")
            if bm_id is not None:
                bm_ends[bm_id] = el

    all_bm_ids = set(bm_starts) | set(bm_ends)
    for bm_id in all_bm_ids:
        start_el = bm_starts.get(bm_id)
        end_el = bm_ends.get(bm_id)
        start_in = (start_el in table_elems) if start_el is not None else False
        end_in = (end_el in table_elems) if end_el is not None else False

        # One boundary inside table, one outside -> crossing hazard
        if start_in != end_in:
            return "table_delete_bookmark_intersection_unsupported"

        # Neither inside table -> check if range encloses / crosses the table
        if not start_in and not end_in and start_el is not None and end_el is not None:
            s_idx = elem_indices[id(start_el)]
            e_idx = elem_indices[id(end_el)]
            if s_idx < tbl_start_idx and e_idx > tbl_end_idx:
                return "table_delete_bookmark_intersection_unsupported"
            if (
                s_idx < tbl_end_idx
                and e_idx > tbl_start_idx
                and not (e_idx < tbl_start_idx or s_idx > tbl_end_idx)
            ):
                return "table_delete_bookmark_intersection_unsupported"

        # Both inside table -> check if bookmark is referenced outside the table
        if start_in and end_in and start_el is not None:
            bm_name = start_el.get(Q(ns, "name")) or start_el.get("name")
            if bm_name:
                for el in doc_elems:
                    if el not in table_elems:
                        if el.tag == Q(ns, "hyperlink"):
                            anchor = el.get(Q(ns, "anchor")) or el.get("anchor")
                            if anchor == bm_name:
                                return "table_delete_bookmark_intersection_unsupported"
                        elif el.tag == Q(ns, "instrText") and el.text:
                            if bm_name in el.text:
                                return "table_delete_bookmark_intersection_unsupported"

    # 2. Comments
    comment_tags = {
        Q(ns, "commentRangeStart"),
        Q(ns, "commentRangeEnd"),
        Q(ns, "commentReference"),
    }
    for el in table_elems:
        if el.tag in comment_tags:
            return "table_delete_annotation_intersection_unsupported"

    c_starts: dict[str, ET.Element] = {}
    c_ends: dict[str, ET.Element] = {}
    for el in doc_elems:
        if el.tag == Q(ns, "commentRangeStart"):
            cid = el.get(Q(ns, "id")) or el.get("id")
            if cid is not None:
                c_starts[cid] = el
        elif el.tag == Q(ns, "commentRangeEnd"):
            cid = el.get(Q(ns, "id")) or el.get("id")
            if cid is not None:
                c_ends[cid] = el

    for cid in set(c_starts) | set(c_ends):
        s_el = c_starts.get(cid)
        e_el = c_ends.get(cid)
        s_in = (s_el in table_elems) if s_el is not None else False
        e_in = (e_el in table_elems) if e_el is not None else False
        if s_in or e_in:
            return "table_delete_annotation_intersection_unsupported"
        if s_el is not None and e_el is not None:
            s_idx = elem_indices[id(s_el)]
            e_idx = elem_indices[id(e_el)]
            if s_idx < tbl_start_idx and e_idx > tbl_end_idx:
                return "table_delete_annotation_intersection_unsupported"

    # 3. Tracked Revisions (ins, del, move ranges)
    revision_tags = {
        Q(ns, "ins"),
        Q(ns, "del"),
        Q(ns, "moveFromRangeStart"),
        Q(ns, "moveFromRangeEnd"),
        Q(ns, "moveToRangeStart"),
        Q(ns, "moveToRangeEnd"),
        Q(ns, "moveFrom"),
        Q(ns, "moveTo"),
    }
    for el in table_elems:
        if el.tag in revision_tags:
            return "table_delete_annotation_intersection_unsupported"

    for start_tag, end_tag in [
        ("moveFromRangeStart", "moveFromRangeEnd"),
        ("moveToRangeStart", "moveToRangeEnd"),
    ]:
        m_starts: dict[str, ET.Element] = {}
        m_ends: dict[str, ET.Element] = {}
        for el in doc_elems:
            if el.tag == Q(ns, start_tag):
                mid = el.get(Q(ns, "id")) or el.get("id")
                if mid is not None:
                    m_starts[mid] = el
            elif el.tag == Q(ns, end_tag):
                mid = el.get(Q(ns, "id")) or el.get("id")
                if mid is not None:
                    m_ends[mid] = el

        for mid in set(m_starts) | set(m_ends):
            s_el = m_starts.get(mid)
            e_el = m_ends.get(mid)
            s_in = (s_el in table_elems) if s_el is not None else False
            e_in = (e_el in table_elems) if e_el is not None else False
            if s_in or e_in:
                return "table_delete_annotation_intersection_unsupported"
            if s_el is not None and e_el is not None:
                s_idx = elem_indices[id(s_el)]
                e_idx = elem_indices[id(e_el)]
                if s_idx < tbl_start_idx and e_idx > tbl_end_idx:
                    return "table_delete_annotation_intersection_unsupported"

    return None


def _inspect_and_prune_relationships(
    graph: PackageGraph,
    root: ET.Element,
    target_table: ET.Element,
) -> tuple[str | None, bytes | None]:
    """Inspect relationship references in table and safely prune solely-owned external rels.

    Preserves all relationships used outside the target table. Refuses internal part
    hazards, broken/missing relationship records, and complex embedded dependencies.
    """
    table_rel_ids: set[str] = set()
    for el in target_table.iter():
        for attr_name, attr_val in el.attrib.items():
            if attr_name.startswith((f"{{{_RELATIONSHIPS_NS}}}", "r:")):
                table_rel_ids.add(attr_val)

    if not table_rel_ids:
        return None, None

    if _RELS_DOCUMENT not in graph.parts:
        return "table_delete_relationship_hazard", None

    try:
        rels_bytes = graph.read_part_bytes(_RELS_DOCUMENT)
        rels_root = parse_xml_with_preserved_qnames(rels_bytes)
    except (ET.ParseError, KeyError, PackageError):
        return "table_delete_relationship_hazard", None

    rels_by_id: dict[str, ET.Element] = {}
    for r in rels_root:
        rid = r.get("Id")
        if rid:
            rels_by_id[rid] = r

    for rel_id in table_rel_ids:
        if rel_id not in rels_by_id:
            return "table_delete_relationship_hazard", None
        rel_elem = rels_by_id[rel_id]
        rel_type = rel_elem.get("Type", "").lower()
        if any(
            structural in rel_type
            for structural in (
                "header",
                "footer",
                "comments",
                "numbering",
                "styles",
                "settings",
                "footnotes",
                "endnotes",
                "fonttable",
            )
        ):
            return "table_delete_relationship_hazard", None

    # Determine which relationships are used outside target_table
    table_elems = set(target_table.iter())
    outside_rel_ids: set[str] = set()
    for el in root.iter():
        if el not in table_elems:
            for attr_name, attr_val in el.attrib.items():
                if attr_name.startswith((f"{{{_RELATIONSHIPS_NS}}}", "r:")):
                    outside_rel_ids.add(attr_val)

    sole_rel_ids = table_rel_ids - outside_rel_ids

    for rid in sole_rel_ids:
        rel_elem = rels_by_id[rid]
        target_mode = rel_elem.get("TargetMode", "")
        rel_type = rel_elem.get("Type", "")
        if target_mode != "External" and "hyperlink" not in rel_type:
            return "table_delete_relationship_hazard", None

    if sole_rel_ids:
        for rel_elem in list(rels_root):
            if rel_elem.get("Id") in sole_rel_ids:
                rels_root.remove(rel_elem)
        updated_rels = serialize_xml_with_preserved_qnames(rels_root)
        return None, updated_rels

    return None, None


def discover_table_delete_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table targets for safe deletion."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-delete-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_delete_source_missing",
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
            "operation": "table-delete-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_delete_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    if part != _DOCUMENT or part not in graph.parts:
        return {
            "schema_version": 1,
            "operation": "table-delete-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_delete_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError, PackageError):
        return {
            "schema_version": 1,
            "operation": "table-delete-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_delete_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-delete-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_delete_namespace_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    body = root.find(Q(namespace, "body"))
    if body is None:
        return {
            "schema_version": 1,
            "operation": "table-delete-targets",
            "passed": True,
            "status": "discovered",
            "code": None,
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    unsupported_context = _unsupported_context_by_element(root, namespace)
    tables = list(root.iter(Q(namespace, "tbl")))
    targets: list[dict[str, Any]] = []

    for table_idx, table in enumerate(tables, start=1):
        if unsupported_context.get(id(table), False):
            continue
        if not _is_supported_table(table, namespace):
            continue
        if _check_nested_hazard(root, table, body, namespace) is not None:
            continue
        if _check_section_hazard(table, body, namespace) is not None:
            continue
        if _check_annotation_and_bookmark_hazards(root, table, namespace) is not None:
            continue
        rel_hazard, _ = _inspect_and_prune_relationships(graph, root, table)
        if rel_hazard is not None:
            continue

        fp = _table_structure_fingerprint(table)
        rows = table.findall(Q(namespace, "tr"))
        tbl_grid = table.find(Q(namespace, "tblGrid"))
        grid_cols = tbl_grid.findall(Q(namespace, "gridCol")) if tbl_grid is not None else []
        cols_count = (
            len(grid_cols)
            if grid_cols
            else max((len(r.findall(Q(namespace, "tc"))) for r in rows), default=0)
        )

        tbl_pr = table.find(Q(namespace, "tblPr"))
        style: str | None = None
        if tbl_pr is not None:
            style_elem = tbl_pr.find(Q(namespace, "tblStyle"))
            if style_elem is not None:
                style = style_elem.get(Q(namespace, "val"))

        texts: list[str] = []
        for t in table.iter(Q(namespace, "t")):
            if t.text:
                texts.append(t.text)
                if sum(len(x) for x in texts) >= 50:
                    break
        preview_text = "".join(texts)[:50]

        target = TableDeleteTarget(
            table=table_idx,
            structure_fingerprint_sha256=fp,
            part=part,
            rows_count=len(rows),
            columns_count=cols_count,
            style=style,
            preview_text=preview_text,
        )
        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-delete-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_table_delete(request: TableDeleteRequest) -> TableDeleteReport:
    """Delete an exact source-bound body table atomically, preserving surrounding structure."""
    if not isinstance(request, TableDeleteRequest):
        raise TypeError("table_delete_request_invalid")

    source = request.source
    destination = request.destination

    if destination.exists() or destination.is_symlink():
        return _table_delete_refusal(request, "table_delete_destination_occupied")

    if not destination.parent.is_dir():
        return _table_delete_refusal(request, "table_delete_destination_invalid")

    try:
        if source.resolve() == destination.resolve():
            return _table_delete_refusal(request, "table_delete_destination_invalid")
    except OSError:
        if source == destination:
            return _table_delete_refusal(request, "table_delete_destination_invalid")

    if not source.is_file():
        return _table_delete_refusal(request, "table_delete_source_missing")

    try:
        source_sha256 = _sha256_file(source)
    except OSError:
        return _table_delete_refusal(request, "table_delete_source_missing")

    if source_sha256 != request.source_sha256:
        return _table_delete_refusal(request, "table_delete_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_delete_refusal(request, "table_delete_package_invalid")

    if request.part != _DOCUMENT or request.part not in graph.parts:
        return _table_delete_refusal(request, "table_delete_part_unsupported")

    try:
        doc_bytes = graph.read_part_bytes(_DOCUMENT)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (PackageError, ET.ParseError, KeyError, OSError):
        return _table_delete_refusal(request, "table_delete_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_delete_refusal(request, "table_delete_namespace_unsupported")

    body = root.find(Q(namespace, "body"))
    if body is None:
        return _table_delete_refusal(request, "table_delete_target_missing")

    tables = list(root.iter(Q(namespace, "tbl")))
    if request.table < 1 or request.table > len(tables):
        return _table_delete_refusal(request, "table_delete_target_missing")

    target_tbl = tables[request.table - 1]

    nested_err = _check_nested_hazard(root, target_tbl, body, namespace)
    if nested_err is not None:
        return _table_delete_refusal(request, nested_err)

    if not _is_supported_table(target_tbl, namespace):
        return _table_delete_refusal(request, "table_delete_context_unsupported")

    actual_fp = _table_structure_fingerprint(target_tbl)
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_delete_refusal(request, "table_delete_target_stale")

    unsupported_context = _unsupported_context_by_element(root, namespace)
    if unsupported_context.get(id(target_tbl), False):
        return _table_delete_refusal(request, "table_delete_context_unsupported")

    sect_err = _check_section_hazard(target_tbl, body, namespace)
    if sect_err is not None:
        return _table_delete_refusal(request, sect_err)

    annot_err = _check_annotation_and_bookmark_hazards(root, target_tbl, namespace)
    if annot_err is not None:
        return _table_delete_refusal(request, annot_err)

    rel_err, updated_rels = _inspect_and_prune_relationships(graph, root, target_tbl)
    if rel_err is not None:
        return _table_delete_refusal(request, rel_err)

    # Capture original structure metrics for delta verification
    orig_tables_count = len(tables)
    orig_body_tables = [c for c in body if c.tag == Q(namespace, "tbl")]
    orig_body_tables_count = len(orig_body_tables)
    orig_body_paragraphs = [c for c in body if c.tag == Q(namespace, "p")]
    orig_body_paragraphs_xml = [ET.tostring(p) for p in orig_body_paragraphs]

    rows = target_tbl.findall(Q(namespace, "tr"))
    tbl_grid = target_tbl.find(Q(namespace, "tblGrid"))
    grid_cols = tbl_grid.findall(Q(namespace, "gridCol")) if tbl_grid is not None else []
    cols_count = (
        len(grid_cols)
        if grid_cols
        else max((len(r.findall(Q(namespace, "tc"))) for r in rows), default=0)
    )

    tbl_pr = target_tbl.find(Q(namespace, "tblPr"))
    style: str | None = None
    if tbl_pr is not None:
        style_elem = tbl_pr.find(Q(namespace, "tblStyle"))
        if style_elem is not None:
            style = style_elem.get(Q(namespace, "val"))

    target_change_report = {
        "table_index": request.table,
        "rows_deleted": len(rows),
        "columns_deleted": cols_count,
        "style": style,
        "remaining_tables": orig_tables_count - 1,
    }

    # Perform table deletion from body
    body.remove(target_tbl)

    serialized_doc = serialize_xml_with_preserved_qnames(root)
    replacements: dict[str, bytes] = {_DOCUMENT: serialized_doc}
    changed_parts: set[str] = {_DOCUMENT}

    if updated_rels is not None:
        replacements[_RELS_DOCUMENT] = updated_rels
        changed_parts.add(_RELS_DOCUMENT)

    frozen_changed = frozenset(changed_parts)

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozen_changed):
            return "table_delete_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_delete_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
            s_ns = _word_namespace(s_root)
            if s_ns != namespace:
                return "table_delete_staged_verification_failed"
            s_body = s_root.find(Q(s_ns, "body"))
            if s_body is None:
                return "table_delete_staged_verification_failed"

            # 1. Exact table-count delta proof
            s_tables = list(s_root.iter(Q(s_ns, "tbl")))
            if len(s_tables) != orig_tables_count - 1:
                return "table_delete_staged_verification_failed"

            s_body_tables = [c for c in s_body if c.tag == Q(s_ns, "tbl")]
            if len(s_body_tables) != orig_body_tables_count - 1:
                return "table_delete_staged_verification_failed"

            # 2. Surrounding body paragraphs byte-identical proof
            s_body_paragraphs = [c for c in s_body if c.tag == Q(s_ns, "p")]
            if len(s_body_paragraphs) != len(orig_body_paragraphs):
                return "table_delete_staged_verification_failed"
            for orig_xml, s_p in zip(orig_body_paragraphs_xml, s_body_paragraphs):
                if ET.tostring(s_p) != orig_xml:
                    return "table_delete_staged_verification_failed"

            # 3. If rels was replaced, verify it parses cleanly
            if _RELS_DOCUMENT in frozen_changed:
                s_rels_bytes = s_graph.read_part_bytes(_RELS_DOCUMENT)
                ET.fromstring(s_rels_bytes)

            return None
        except (PackageError, ET.ParseError, KeyError, OSError):
            return "table_delete_staged_verification_failed"

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=[target_change_report],
        verify=verify_publication,
        failure_code="table_delete_staged_verification_failed",
        exists_code="table_delete_destination_occupied",
        unavailable_code="table_delete_publication_unavailable",
    )

    if not published.passed:
        return _table_delete_refusal(
            request, published.code or "table_delete_staged_verification_failed"
        )

    return TableDeleteReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        deleted_table=target_change_report,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
