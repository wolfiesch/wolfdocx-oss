"""Source-bound table row authoring and discovery for WordprocessingML documents.

Admitted Merge Boundary:
- WordprocessingML table rows containing merged cells (<w:vMerge>, <w:hMerge>, <w:gridSpan>)
  cannot serve as authoring prototypes for row insertion, nor can rows in tables containing
  active vertical-merge (<w:vMerge>) chains be inserted into, because vertical merges define
  cross-row cell ownership chains that would become desynchronized or corrupted by inserted rows.
- Similarly, table row deletion strictly enforces merge boundaries:
  (1) Deletion of any row containing vertical merge markup (<w:vMerge w:val="restart"> or continuation
      <w:vMerge/> / <w:vMerge w:val="continue">) is refused with 'table_row_merged_unsupported',
      preserving vertical merge ownership.
  (2) Deletion of any row within a table containing vertical merge chains is refused with
      'table_row_merged_unsupported' to prevent breaking vertical merge ownership chains across columns.
  (3) Deletion of rows with horizontal merge (<w:hMerge> or <w:gridSpan>) or non-uniform cell counts
      is refused ('table_row_merged_unsupported' / 'table_row_grid_mismatch').
- Only approved, unmerged rows in tables with intact grid column structures and without vertical
  merge chains are admitted for deletion. A row that contains nested ``w:tbl`` cell blocks is
  copy-safe for insertion and deletion; its text-replacement surface is intentionally unavailable.
"""

from __future__ import annotations

import copy
import json
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .authoring import (
    _MAX_TEXT_CHARACTERS,
    _PUBLISHED,
    _REFUSED,
    _cell_structure_fingerprint,
    _require_sha256,
    _sha256_file,
    _simple_text_node,
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
from .table_row_properties import _TRPR_CHILD_ORDER, _insert_or_replace_trpr_child

_OFF_VALUES = frozenset({"0", "false", "off"})

__all__ = [
    "TableRowReport",
    "TableRowRequest",
    "TableRowTarget",
    "author_table_row",
    "discover_table_row_targets",
    "load_table_row_request",
    "parse_table_row_request",
    "table_row_structure_fingerprint",
]

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
        "position",
        "cell_texts",
        "repeat_header",
        "cant_split",
        "action",
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
_REQUIRED_TARGET_FIELDS = frozenset({"table", "row"})
_ALLOWED_TARGET_FIELDS = frozenset({"table", "row", "part"})

_FORBIDDEN_ROW_ELEMENTS = frozenset(
    {
        "bookmarkStart",
        "bookmarkEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
        "customXml",
        "del",
        "delText",
        "drawing",
        "endnoteReference",
        "fldChar",
        "fldSimple",
        "footnoteReference",
        "hyperlink",
        "ins",
        "instrText",
        "moveFrom",
        "moveFromRangeEnd",
        "moveFromRangeStart",
        "moveTo",
        "moveToRangeEnd",
        "moveToRangeStart",
        "pict",
        "sdt",
        "smartTag",
        "subDoc",
        "wgp",
    }
)


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is the source part byte for byte."""
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


def _has_forbidden_row_descendant(element: ET.Element) -> bool:
    """True if element or any descendant has forbidden elements or IDs."""
    for child in element.iter():
        local = child.tag.rsplit("}", 1)[-1]
        if local in _FORBIDDEN_ROW_ELEMENTS:
            return True
        for attr in child.attrib:
            attr_local = attr.rsplit("}", 1)[-1]
            if attr_local in ("id", "embed", "href"):
                return True
    return False


def _row_cant_split(row: ET.Element, namespace: str) -> bool:
    """Return the effective direct ``w:cantSplit`` state of a row (CT_OnOff semantics)."""
    tr_pr = row.find(Q(namespace, "trPr"))
    if tr_pr is None:
        return False
    cant_split = tr_pr.find(Q(namespace, "cantSplit"))
    if cant_split is None:
        return False
    value = cant_split.get(Q(namespace, "val"))
    return value is None or value.lower() not in _OFF_VALUES


def _on_off_state(element: ET.Element, namespace: str) -> bool:
    """Return the CT_OnOff state of a present element (absent ``w:val`` means on)."""
    value = element.get(Q(namespace, "val"))
    return value is None or value.lower() not in _OFF_VALUES


def _row_repeat_header(row: ET.Element, namespace: str) -> bool:
    """Return the effective direct ``w:tblHeader`` state of a row (CT_OnOff semantics)."""
    tr_pr = row.find(Q(namespace, "trPr"))
    if tr_pr is None:
        return False
    header = tr_pr.find(Q(namespace, "tblHeader"))
    return header is not None and _on_off_state(header, namespace)


def _is_copy_safe_cell(cell: ET.Element, namespace: str) -> bool:
    """True if a cell can be cloned without treating its nested blocks as text."""

    def word(local):
        return Q(namespace, local)
    properties = cell.find(word("tcPr"))
    if properties is not None and any(
        properties.find(word(name)) is not None for name in ("vMerge", "hMerge", "gridSpan")
    ):
        return False
    if any(child.tag not in {word("tcPr"), word("p"), word("tbl")} for child in cell):
        return False
    for child in cell:
        if child.tag == word("p"):
            if any(grandchild.tag not in {word("pPr"), word("r")} for grandchild in child):
                return False
            for run in child.findall(word("r")):
                if any(grandchild.tag not in {word("rPr"), word("t")} for grandchild in run):
                    return False
        elif child.tag == word("tbl"):
            if not _is_supported_table(child, namespace):
                return False
            if any(
                not _is_approved_row(nested_row, namespace)
                for nested_row in child.findall(word("tr"))
            ):
                return False
    return any(child.tag in {word("p"), word("tbl")} for child in cell)


def _is_approved_row(row: ET.Element, namespace: str) -> bool:
    """True if an unmerged row has a safe copy surface."""

    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("trPr"), word("tc")} for child in row):
        return False
    cells = row.findall(word("tc"))
    return bool(cells) and not any(
        _has_forbidden_row_descendant(cell) or not _is_copy_safe_cell(cell, namespace)
        for cell in cells
    )


def _is_supported_table(table: ET.Element, namespace: str) -> bool:
    """True if table has supported body table structure."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("tblPr"), word("tblGrid"), word("tr")} for child in table):
        return False
    rows = table.findall(word("tr"))
    return bool(rows)


def _row_structure_fingerprint(row: ET.Element) -> str:
    """Hash the row's XML structure while deliberately masking cell text."""
    return _cell_structure_fingerprint(row)


def table_row_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    row: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one exact prototype row."""
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
        if not _is_supported_table(target_table, ns):
            return None
        rows = target_table.findall(Q(ns, "tr"))
        if row < 1 or row > len(rows):
            return None
        target_row = rows[row - 1]
        tbl_grid = target_table.find(Q(ns, "tblGrid"))
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(Q(ns, "gridCol"))
            if grid_cols and len(target_row.findall(Q(ns, "tc"))) != len(grid_cols):
                return None
        if not _is_approved_row(target_row, ns):
            return None
        return _row_structure_fingerprint(target_row)
    except (PackageError, OSError, ET.ParseError, KeyError):
        return None


@dataclass(frozen=True)
class TableRowTarget:
    """One discovered candidate table row that can serve as an authoring prototype."""

    table: int
    row: int
    cells_count: int
    structure_fingerprint_sha256: str
    cell_texts: tuple[str, ...] | None
    is_repeat_header: bool
    part: str = _DOCUMENT
    actions: tuple[str, ...] = ("insert", "delete")
    cant_split: bool = False

    @property
    def target(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "row": self.row,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "row": self.row,
            "cells_count": self.cells_count,
            "cell_texts": list(self.cell_texts) if self.cell_texts is not None else None,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "is_repeat_header": self.is_repeat_header,
            "actions": list(self.actions),
            "cant_split": self.cant_split,
        }


@dataclass(frozen=True)
class TableRowRequest:
    """One source-bound request to insert or delete a styled table row."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    position: str | None = None
    cell_texts: tuple[str, ...] | None = None
    repeat_header: bool | None = None
    cant_split: bool | None = None
    action: str = "insert"

    def __post_init__(self) -> None:
        if not isinstance(self.action, str) or self.action not in ("insert", "delete"):
            raise ValueError("table_row_action_invalid")
        if self.cant_split is not None and not isinstance(self.cant_split, bool):
            raise TypeError("table_row_cant_split_invalid")
        if self.action == "delete":
            if self.position is not None:
                raise ValueError("table_row_request_fields_invalid")
            if self.cell_texts is not None:
                raise ValueError("table_row_request_fields_invalid")
            if self.repeat_header is not None or self.cant_split is not None:
                raise ValueError("table_row_request_fields_invalid")
        else:
            if self.position is None:
                object.__setattr__(self, "position", "after")
            elif not isinstance(self.position, str) or self.position not in (
                "before",
                "after",
            ):
                raise ValueError("table_row_position_invalid")

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def row(self) -> int:
        return self.target["row"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "part": self.part,
            "table": self.table,
            "row": self.row,
        }
        if self.action != "delete" and self.position is not None:
            data["position"] = self.position
        return data

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "target": {
                "part": self.part,
                "table": self.table,
                "row": self.row,
            },
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
        }
        if self.action == "delete":
            data["action"] = "delete"
            return data
        data["position"] = self.position
        if self.cell_texts is not None:
            data["cell_texts"] = list(self.cell_texts)
        if self.repeat_header is not None:
            data["repeat_header"] = self.repeat_header
        if self.cant_split is not None:
            data["cant_split"] = self.cant_split
        return data


@dataclass(frozen=True)
class TableRowReport:
    """Content-safe public result for a table row authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "table-row",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _table_row_refusal(
    request: TableRowRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TableRowReport:
    return TableRowReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_row_request(value: Any) -> TableRowRequest:
    """Parse and strictly validate a table row authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_row_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_row_request_fields_invalid")

    action = value.get("action", "insert")
    if not isinstance(action, str) or action not in ("insert", "delete"):
        raise ValueError("table_row_action_invalid")

    if action == "delete" and not keys.issubset(_ALLOWED_DELETE_REQUEST_FIELDS):
        raise ValueError("table_row_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("table_row_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("table_row_destination_invalid")

    source_sha256 = _require_sha256(value["source_sha256"], "table_row_source_sha256")
    structure_fingerprint = _require_sha256(
        value["structure_fingerprint_sha256"], "table_row_structure_fingerprint_sha256"
    )

    target = value["target"]
    if not isinstance(target, Mapping):
        raise TypeError("table_row_target_invalid")
    target_keys = set(target)
    if not target_keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(
        target_keys
    ):
        raise ValueError("table_row_target_fields_invalid")

    table = target["table"]
    if isinstance(table, bool) or not isinstance(table, int) or table < 1:
        raise ValueError("table_row_table_invalid")

    row = target["row"]
    if isinstance(row, bool) or not isinstance(row, int) or row < 1:
        raise ValueError("table_row_row_invalid")

    part = target.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("table_row_part_invalid")
    if action == "delete":
        return TableRowRequest(
            source=Path(source),
            destination=Path(destination),
            source_sha256=source_sha256,
            target={"table": table, "row": row, "part": part},
            structure_fingerprint_sha256=structure_fingerprint,
            action="delete",
        )

    position = value.get("position", "after")
    if not isinstance(position, str) or position not in ("before", "after"):
        raise ValueError("table_row_position_invalid")

    raw_cell_texts = value.get("cell_texts")
    cell_texts: tuple[str, ...] | None = None
    if raw_cell_texts is not None:
        if isinstance(raw_cell_texts, (str, bytes, Mapping)) or not isinstance(
            raw_cell_texts, Sequence
        ):
            raise TypeError("table_row_cell_texts_invalid")
        parsed_texts: list[str] = []
        for text in raw_cell_texts:
            if not isinstance(text, str):
                raise TypeError("table_row_cell_text_invalid")
            if len(text) > _MAX_TEXT_CHARACTERS:
                raise ValueError("table_row_cell_text_limit_exceeded")
            parsed_texts.append(text)
        cell_texts = tuple(parsed_texts)

    repeat_header = value.get("repeat_header")
    if repeat_header is not None and not isinstance(repeat_header, bool):
        raise TypeError("table_row_repeat_header_invalid")

    cant_split = value.get("cant_split")
    if cant_split is not None and not isinstance(cant_split, bool):
        raise TypeError("table_row_cant_split_invalid")

    return TableRowRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        target={"table": table, "row": row, "part": part},
        structure_fingerprint_sha256=structure_fingerprint,
        position=position,
        cell_texts=cell_texts,
        repeat_header=repeat_header,
        cant_split=cant_split,
        action="insert",
    )


def load_table_row_request(path: str | Path) -> TableRowRequest:
    """Load and strictly validate one JSON table-row request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_row_request(json.load(stream))


def discover_table_row_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table row targets for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-row-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_source_missing",
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
            "operation": "table-row-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    if part != _DOCUMENT or part not in graph.parts:
        return {
            "schema_version": 1,
            "operation": "table-row-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "table-row-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-row-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_namespace_unsupported",
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
        tbl_grid = table.find(Q(namespace, "tblGrid"))
        grid_cols = tbl_grid.findall(Q(namespace, "gridCol")) if tbl_grid is not None else []
        rows = table.findall(Q(namespace, "tr"))

        table_has_vmerge = False
        for r in rows:
            for c in r.findall(Q(namespace, "tc")):
                tcp = c.find(Q(namespace, "tcPr"))
                if tcp is not None and tcp.find(Q(namespace, "vMerge")) is not None:
                    table_has_vmerge = True
                    break
            if table_has_vmerge:
                break

        for row_idx, row in enumerate(rows, start=1):
            if unsupported_context.get(id(row), False):
                continue
            if not _is_approved_row(row, namespace):
                continue
            cells = row.findall(Q(namespace, "tc"))
            fp = _row_structure_fingerprint(row)
            simple_cells = [_simple_text_node(cell, namespace) for cell in cells]
            cell_texts = (
                tuple(node[1] for node in simple_cells)
                if all(node is not None for node in simple_cells)
                else None
            )
            if grid_cols and len(cells) != len(grid_cols):
                continue
            is_repeat = _row_repeat_header(row, namespace)
            can_delete = len(rows) > 1 and not table_has_vmerge
            actions = ("insert", "delete") if can_delete else ("insert",)
            target = TableRowTarget(
                table=table_idx,
                row=row_idx,
                cells_count=len(cells),
                structure_fingerprint_sha256=fp,
                cell_texts=cell_texts,
                is_repeat_header=is_repeat,
                part=part,
                actions=actions,
                cant_split=_row_cant_split(row, namespace),
            )
            targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-row-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_table_row(request: TableRowRequest) -> TableRowReport:
    """Insert a source-bound styled table row before or after a bound prototype atomically."""
    source = request.source
    destination = request.destination

    if not source.is_file():
        return _table_row_refusal(request, "table_row_source_missing")
    if source == destination:
        return _table_row_refusal(request, "table_row_destination_invalid")
    if not destination.parent.is_dir():
        return _table_row_refusal(request, "table_row_destination_invalid")
    if destination.exists() or destination.is_symlink():
        return _table_row_refusal(request, "table_row_destination_occupied")

    try:
        actual_source_sha256 = _sha256_file(source)
    except OSError:
        return _table_row_refusal(request, "table_row_source_missing")

    if actual_source_sha256 != request.source_sha256:
        return _table_row_refusal(request, "table_row_source_hash_mismatch")

    if request.part != _DOCUMENT:
        return _table_row_refusal(request, "table_row_part_unsupported")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_row_refusal(request, "table_row_package_invalid")

    if _DOCUMENT not in graph.parts:
        return _table_row_refusal(request, "table_row_package_invalid")

    try:
        doc_bytes = graph.read_part_bytes(_DOCUMENT)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (KeyError, ET.ParseError):
        return _table_row_refusal(request, "table_row_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_row_refusal(request, "table_row_namespace_unsupported")

    unsupported_context = _unsupported_context_by_element(root, namespace)
    tables = list(root.iter(Q(namespace, "tbl")))
    if request.table < 1 or request.table > len(tables):
        return _table_row_refusal(request, "table_row_target_missing")

    table = tables[request.table - 1]
    if unsupported_context.get(id(table), False):
        return _table_row_refusal(request, "table_row_context_unsupported")

    # Check table structure and merged cells
    if not _is_supported_table(table, namespace):
        return _table_row_refusal(request, "table_row_table_unsupported")

    rows = table.findall(Q(namespace, "tr"))
    if request.row < 1 or request.row > len(rows):
        return _table_row_refusal(request, "table_row_target_missing")

    prototype_row = rows[request.row - 1]
    if unsupported_context.get(id(prototype_row), False):
        return _table_row_refusal(request, "table_row_context_unsupported")

    # Refuse merged rows (vMerge, hMerge, gridSpan on prototype or vMerge anywhere in table)
    for r in rows:
        for c in r.findall(Q(namespace, "tc")):
            tcp = c.find(Q(namespace, "tcPr"))
            if tcp is not None:
                if r is prototype_row and any(
                    tcp.find(Q(namespace, m)) is not None for m in ("vMerge", "hMerge", "gridSpan")
                ):
                    return _table_row_refusal(request, "table_row_merged_unsupported")
                if tcp.find(Q(namespace, "vMerge")) is not None:
                    return _table_row_refusal(request, "table_row_merged_unsupported")

    if not _is_approved_row(prototype_row, namespace):
        return _table_row_refusal(request, "table_row_structure_unsupported")

    prototype_cells = prototype_row.findall(Q(namespace, "tc"))
    tbl_grid = table.find(Q(namespace, "tblGrid"))
    if tbl_grid is not None:
        grid_cols = tbl_grid.findall(Q(namespace, "gridCol"))
        if grid_cols and len(prototype_cells) != len(grid_cols):
            return _table_row_refusal(request, "table_row_grid_mismatch")

    # Fingerprint check
    actual_fp = _row_structure_fingerprint(prototype_row)
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_row_refusal(request, "table_row_structure_fingerprint_mismatch")

    if request.action == "delete":
        if len(rows) <= 1:
            return _table_row_refusal(request, "table_row_final_row_unsupported")

        table.remove(prototype_row)

        target_change_report = {
            "action": "delete",
            "table": request.table,
            "row": request.row,
            "deleted_row": request.row,
            "rows_count_before": len(rows),
            "rows_count_after": len(rows) - 1,
            "cells_count": len(prototype_cells),
        }

        def verify_deletion_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, frozenset({_DOCUMENT})):
                return "table_row_staged_verification_failed"
            insp = inspect_docx(staged)
            if not insp.valid:
                return "table_row_staged_verification_failed"
            try:
                s_graph = PackageGraph.from_file(staged)
                s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
                s_tables = list(s_root.iter(Q(namespace, "tbl")))
                expected_table_count = len(tables) - len(
                    list(prototype_row.iter(Q(namespace, "tbl")))
                )
                if len(s_tables) != expected_table_count:
                    return "table_row_staged_verification_failed"
                s_table = s_tables[request.table - 1]
                s_rows = s_table.findall(Q(namespace, "tr"))
                if len(s_rows) != len(rows) - 1:
                    return "table_row_staged_verification_failed"
                orig_grid = table.find(Q(namespace, "tblGrid"))
                staged_grid = s_table.find(Q(namespace, "tblGrid"))
                if (orig_grid is None) != (staged_grid is None):
                    return "table_row_staged_verification_failed"
                if (
                    orig_grid is not None
                    and staged_grid is not None
                    and ET.tostring(orig_grid) != ET.tostring(staged_grid)
                ):
                    return "table_row_staged_verification_failed"
                orig_remaining_rows = [r for i, r in enumerate(rows, start=1) if i != request.row]
                for orig_r, staged_r in zip(orig_remaining_rows, s_rows):
                    if ET.tostring(orig_r) != ET.tostring(staged_r):
                        return "table_row_staged_verification_failed"
            except (
                OSError,
                PackageError,
                ET.ParseError,
                KeyError,
                IndexError,
                ValueError,
            ):
                return "table_row_staged_verification_failed"
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
            failure_code="table_row_staged_verification_failed",
            exists_code="table_row_destination_occupied",
            unavailable_code="table_row_publication_unavailable",
        )

        if not published.passed:
            return _table_row_refusal(
                request, published.code or "table_row_staged_verification_failed"
            )

        return TableRowReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            target=request.target_dict(),
            target_change_report=target_change_report,
            non_target_parts_preserved=True,
        )
    # Cell texts length check
    if request.cell_texts is not None and len(request.cell_texts) != len(prototype_cells):
        return _table_row_refusal(request, "table_row_cell_texts_mismatch")

    nested_tables_in_prototype = len(list(prototype_row.iter(Q(namespace, "tbl"))))

    # Clone prototype row
    new_row = copy.deepcopy(prototype_row)

    # Handle repeat_header
    tr_pr = new_row.find(Q(namespace, "trPr"))
    if request.repeat_header is True:
        if tr_pr is None:
            tr_pr = ET.Element(Q(namespace, "trPr"))
            new_row.insert(0, tr_pr)
        header = tr_pr.find(Q(namespace, "tblHeader"))
        if header is None:
            ET.SubElement(tr_pr, Q(namespace, "tblHeader"))
        elif not _on_off_state(header, namespace):
            # An explicit-off prototype marker must become a bare (on) marker.
            header.attrib.pop(Q(namespace, "val"), None)
    elif request.repeat_header is False:
        if tr_pr is not None:
            for elem in list(tr_pr.findall(Q(namespace, "tblHeader"))):
                tr_pr.remove(elem)

    # Handle cant_split: True writes a bare w:cantSplit in schema order; False removes it.
    tr_pr = new_row.find(Q(namespace, "trPr"))
    if request.cant_split is True:
        if tr_pr is None:
            tr_pr = ET.Element(Q(namespace, "trPr"))
            new_row.insert(0, tr_pr)
        _insert_or_replace_trpr_child(tr_pr, ET.Element(Q(namespace, "cantSplit")), namespace)
    elif request.cant_split is False and tr_pr is not None:
        for elem in list(tr_pr.findall(Q(namespace, "cantSplit"))):
            tr_pr.remove(elem)
        if len(tr_pr) == 0:
            new_row.remove(tr_pr)

    # Handle cell texts replacement
    if request.cell_texts is not None:
        new_cells = new_row.findall(Q(namespace, "tc"))
        for cell, text in zip(new_cells, request.cell_texts):
            simple = _simple_text_node(cell, namespace)
            if simple is None:
                return _table_row_refusal(request, "table_row_structure_unsupported")
            t_node = simple[0]
            t_node.text = text
            if text.startswith(" ") or text.endswith(" "):
                t_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            elif "{http://www.w3.org/XML/1998/namespace}space" in t_node.attrib:
                del t_node.attrib["{http://www.w3.org/XML/1998/namespace}space"]

    # Insert new row into table
    children = list(table)
    proto_idx = children.index(prototype_row)
    if request.position == "before":
        table.insert(proto_idx, new_row)
        inserted_row_num = request.row
    else:
        table.insert(proto_idx + 1, new_row)
        inserted_row_num = request.row + 1

    is_repeat = _row_repeat_header(new_row, namespace)
    is_cant_split = _row_cant_split(new_row, namespace)

    target_change_report = {
        "table": request.table,
        "prototype_row": request.row,
        "position": request.position,
        "inserted_row": inserted_row_num,
        "rows_count_before": len(rows),
        "rows_count_after": len(rows) + 1,
        "cells_count": len(prototype_cells),
        "repeat_header": is_repeat,
        "cant_split": is_cant_split,
    }

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({_DOCUMENT})):
            return "table_row_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_row_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
            s_tables = list(s_root.iter(Q(namespace, "tbl")))
            expected_table_count = len(tables) + nested_tables_in_prototype
            if len(s_tables) != expected_table_count:
                return "table_row_staged_verification_failed"
            s_table = s_tables[request.table - 1]
            s_rows = s_table.findall(Q(namespace, "tr"))
            if len(s_rows) != len(rows) + 1:
                return "table_row_staged_verification_failed"
            orig_grid = table.find(Q(namespace, "tblGrid"))
            staged_grid = s_table.find(Q(namespace, "tblGrid"))
            if (orig_grid is None) != (staged_grid is None):
                return "table_row_staged_verification_failed"
            if (
                orig_grid is not None
                and staged_grid is not None
                and ET.tostring(orig_grid) != ET.tostring(staged_grid)
            ):
                return "table_row_staged_verification_failed"
            ins_row = s_rows[inserted_row_num - 1]
            ins_cells = ins_row.findall(Q(namespace, "tc"))
            if len(ins_cells) != len(prototype_cells):
                return "table_row_staged_verification_failed"
            ins_row_tr_pr = ins_row.find(Q(namespace, "trPr"))
            staged_is_repeat = _row_repeat_header(ins_row, namespace)
            if staged_is_repeat != is_repeat:
                return "table_row_staged_verification_failed"
            if _row_cant_split(ins_row, namespace) != is_cant_split:
                return "table_row_staged_verification_failed"
            if request.cant_split is True:
                if ins_row_tr_pr is None:
                    return "table_row_staged_verification_failed"
                ins_children = list(ins_row_tr_pr)
                markers = ins_row_tr_pr.findall(Q(namespace, "cantSplit"))
                if len(markers) != 1 or markers[0].attrib:
                    return "table_row_staged_verification_failed"
                rank = _TRPR_CHILD_ORDER.index("cantSplit")
                preceding = ins_children[: ins_children.index(markers[0])]
                if any(
                    child.tag.rsplit("}", 1)[-1] not in _TRPR_CHILD_ORDER
                    or _TRPR_CHILD_ORDER.index(child.tag.rsplit("}", 1)[-1]) >= rank
                    for child in preceding
                ):
                    return "table_row_staged_verification_failed"
            if request.cell_texts is not None:
                for c, exp_txt in zip(ins_cells, request.cell_texts):
                    res = _simple_text_node(c, namespace)
                    if res is None or res[1] != exp_txt:
                        return "table_row_staged_verification_failed"
        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return "table_row_staged_verification_failed"
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
        verify=verify_publication,
        failure_code="table_row_staged_verification_failed",
        exists_code="table_row_destination_occupied",
        unavailable_code="table_row_publication_unavailable",
    )

    if not published.passed:
        return _table_row_refusal(request, published.code or "table_row_staged_verification_failed")

    return TableRowReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
