"""Source-bound table column authoring and discovery for WordprocessingML documents.

Supports inserting and deleting columns in admitted unmerged tables and resizing
grid columns without changing merged regions. Body, reachable header/footer,
nested and text-box tables share source binding and publication checks.
"""

from __future__ import annotations

import copy
import hashlib
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
    _XML_SPACE_ATTRIBUTE,
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
from .table_properties import _check_reachable_part

__all__ = [
    "TableColumnReport",
    "TableColumnRequest",
    "TableColumnTarget",
    "author_table_column",
    "discover_table_column_targets",
    "load_table_column_request",
    "parse_table_column_request",
    "table_column_structure_fingerprint",
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
        "width",
        "column_width",
        "cell_texts",
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
_REQUIRED_TARGET_FIELDS = frozenset({"table", "column"})
_ALLOWED_TARGET_FIELDS = frozenset({"table", "column", "col", "part"})

_ALLOWED_RESIZE_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "width",
        "column_width",
        "action",
    }
)
_ACTIONS = ("insert", "delete", "resize")
# Word's maximum page width (22 inches); no single grid column can exceed it.
_MAX_COLUMN_WIDTH_TWIPS = 31680

_FORBIDDEN_COLUMN_ELEMENTS = frozenset(
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

_TCPR_ORDER: tuple[str, ...] = (
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
)


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is byte-for-byte identical to source."""
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


def _has_forbidden_descendant(element: ET.Element) -> bool:
    """True if element or any descendant has forbidden elements or IDs."""
    for child in element.iter():
        local = child.tag.rsplit("}", 1)[-1]
        if local in _FORBIDDEN_COLUMN_ELEMENTS:
            return True
        for attr in child.attrib:
            attr_local = attr.rsplit("}", 1)[-1]
            if attr_local in ("id", "embed", "href"):
                return True
    return False


def _has_illegal_xml_chars(text: str) -> bool:
    """True if text contains characters prohibited in XML 1.0."""
    for ch in text:
        cp = ord(ch)
        if cp < 0x20:
            if ch not in ("\t", "\n", "\r"):
                return True
        elif 0xD800 <= cp <= 0xDFFF or cp in (0xFFFE, 0xFFFF):
            return True
    return False


def _is_approved_row(row: ET.Element, namespace: str, *, allow_merged: bool = False) -> bool:
    """True if row conforms to approved rectangular row conventions."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("trPr"), word("tc")} for child in row):
        return False
    tr_pr = row.find(word("trPr"))
    if tr_pr is not None and (
        tr_pr.find(word("gridBefore")) is not None or tr_pr.find(word("gridAfter")) is not None
    ):
        return False
    cells = row.findall(word("tc"))
    if not cells:
        return False
    for cell in cells:
        if _has_forbidden_descendant(cell):
            return False
        properties = cell.find(word("tcPr"))
        merged = properties is not None and any(
            properties.find(word(name)) is not None for name in ("vMerge", "hMerge", "gridSpan")
        )
        if _simple_text_node(cell, namespace) is None and not (allow_merged and merged):
            return False
        if (
            not allow_merged
            and properties is not None
            and any(
                properties.find(word(name)) is not None for name in ("vMerge", "hMerge", "gridSpan")
            )
        ):
            return False
        if cell.find(f".//{word('tbl')}") is not None:
            return False
    return True


def _is_supported_table(table: ET.Element, namespace: str) -> bool:
    """True if table has supported body table structure."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("tblPr"), word("tblGrid"), word("tr")} for child in table):
        return False
    tbl_grid = table.find(word("tblGrid"))
    if tbl_grid is None:
        return False
    grid_cols = tbl_grid.findall(word("gridCol"))
    if not grid_cols:
        return False
    rows = table.findall(word("tr"))
    return bool(rows)


def _grid_row_cells(row: ET.Element, namespace: str) -> list[tuple[ET.Element, int, int]]:
    """Physical cells and their inclusive, one-based grid intervals."""
    def word(local):
        return Q(namespace, local)
    result = []
    start = 1
    for cell in row.findall(word("tc")):
        span = cell.find(f"{word('tcPr')}/{word('gridSpan')}")
        width = 1 if span is None else int(span.get(word("val"), "0"))
        result.append((cell, start, start + width - 1))
        start += width
    return result


def _table_hazard_reason(
    table: ET.Element, namespace: str, *, allow_merged: bool = False
) -> str | None:
    """Admit a complete rectangular grid, retaining merge markup during resize."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("tblPr"), word("tblGrid"), word("tr")} for child in table):
        return "table_column_table_unsupported"
    tbl_grid = table.find(word("tblGrid"))
    if tbl_grid is None:
        return "table_column_table_unsupported"
    grid_cols = tbl_grid.findall(word("gridCol"))
    if not grid_cols:
        return "table_column_table_unsupported"
    rows = table.findall(word("tr"))
    if not rows:
        return "table_column_table_unsupported"

    num_cols = len(grid_cols)
    previous_vertical: set[tuple[int, int]] = set()
    for row in rows:
        tr_pr = row.find(word("trPr"))
        if tr_pr is not None and (
            tr_pr.find(word("gridBefore")) is not None or tr_pr.find(word("gridAfter")) is not None
        ):
            return "table_column_grid_before_after_unsupported"
        cells = row.findall(word("tc"))
        if not allow_merged and len(cells) != num_cols:
            return "table_column_grid_mismatch"
        horizontal = False
        vertical: set[tuple[int, int]] = set()
        end = 0
        for cell in cells:
            if cell.find(f".//{word('tbl')}") is not None:
                return "table_column_structure_unsupported"
            if _has_forbidden_descendant(cell):
                return "table_column_context_unsupported"
            properties = cell.find(word("tcPr"))
            span = 1
            modes = {}
            if properties is not None:
                for name in ("gridSpan", "hMerge", "vMerge"):
                    nodes = properties.findall(word(name))
                    if nodes and not allow_merged:
                        return "table_column_merged_unsupported"
                    if len(nodes) > 1:
                        return "table_column_merge_markup_invalid"
                    if not nodes:
                        continue
                    value = nodes[0].get(word("val"))
                    if name == "gridSpan":
                        try:
                            span = int(value or "0")
                        except ValueError:
                            return "table_column_merge_markup_invalid"
                        if span < 1:
                            return "table_column_merge_markup_invalid"
                    else:
                        if value not in (None, "restart", "continue"):
                            return "table_column_merge_markup_invalid"
                        modes[name] = value or "continue"
            interval = (end + 1, end + span)
            end += span
            h_mode = modes.get("hMerge")
            if h_mode == "continue" and not horizontal:
                return "table_column_merge_markup_invalid"
            horizontal = h_mode is not None
            v_mode = modes.get("vMerge")
            if v_mode == "continue" and interval not in previous_vertical:
                return "table_column_merge_markup_invalid"
            if v_mode is not None:
                vertical.add(interval)
            if _simple_text_node(cell, namespace) is None and not (
                allow_merged
                and properties is not None
                and any(
                    properties.find(word(name)) is not None
                    for name in ("gridSpan", "hMerge", "vMerge")
                )
            ):
                return "table_column_structure_unsupported"
        if not cells or end != num_cols:
            return "table_column_grid_mismatch"
        previous_vertical = vertical
    return None


def _insert_tcpr_child_in_order(tc_pr: ET.Element, child: ET.Element, namespace: str) -> None:
    """Insert or replace a child element in tcPr while strictly maintaining schema order."""
    child_name = child.tag.rsplit("}", 1)[-1]
    if child_name not in _TCPR_ORDER:
        tc_pr.append(child)
        return
    target_rank = _TCPR_ORDER.index(child_name)
    existing = tc_pr.find(Q(namespace, child_name))
    if existing is not None:
        idx = list(tc_pr).index(existing)
        tc_pr.remove(existing)
        tc_pr.insert(idx, child)
        return
    insert_at = len(tc_pr)
    for idx, c in enumerate(tc_pr):
        c_name = c.tag.rsplit("}", 1)[-1]
        c_rank = _TCPR_ORDER.index(c_name) if c_name in _TCPR_ORDER else 999
        if c_rank > target_rank:
            insert_at = idx
            break
    tc_pr.insert(insert_at, child)


def _is_valid_column_width(width: object) -> bool:
    """True when width is a positive integer twips value no wider than a Word page."""
    return (
        isinstance(width, int)
        and not isinstance(width, bool)
        and 0 < width <= _MAX_COLUMN_WIDTH_TWIPS
    )


def _cell_content_signature(
    cell: ET.Element, namespace: str
) -> tuple[tuple[bytes, ...], tuple[bytes, ...]]:
    """Serialize a cell's content and its tcPr children except tcW, in document order."""
    tc_pr_tag = Q(namespace, "tcPr")
    tc_w_tag = Q(namespace, "tcW")
    content = tuple(ET.tostring(child) for child in cell if child.tag != tc_pr_tag)
    tc_pr = cell.find(tc_pr_tag)
    properties = () if tc_pr is None else tuple(ET.tostring(c) for c in tc_pr if c.tag != tc_w_tag)
    return content, properties


def _get_element_width(elem: ET.Element, namespace: str) -> int | None:
    """Read w:w integer attribute from gridCol, tcW, or tblW element."""
    w_attr = elem.get(Q(namespace, "w")) or elem.get("w")
    if w_attr is not None and w_attr.isdigit():
        return int(w_attr)
    return None


def _column_structure_fingerprint(
    table: ET.Element, column: int, namespace: str, *, part: str = _DOCUMENT
) -> str:
    """Hash the selected column's XML structure across tblGrid and every row, masking cell text."""
    def word(local):
        return Q(namespace, local)
    col_idx = column - 1
    tbl_grid = table.find(word("tblGrid"))
    assert tbl_grid is not None
    grid_cols = tbl_grid.findall(word("gridCol"))
    grid_col = grid_cols[col_idx]
    rows = table.findall(word("tr"))
    wrapper = ET.Element(word("columnPrototype"))
    if part != _DOCUMENT:
        wrapper.set("part", part)
    wrapper.append(copy.deepcopy(grid_col))
    for row in rows:
        for cell, start, end in _grid_row_cells(row, namespace):
            if start <= column <= end:
                wrapper.append(copy.deepcopy(cell))
                break
    return _cell_structure_fingerprint(wrapper)


def table_column_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    column: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one target table column."""
    # Reachability is checked against the package rather than the part's filename alone.
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        valid, _ = _check_reachable_part(graph, part, "table_column")
        if not valid:
            return None
        root = ET.fromstring(graph.read_part_bytes(part))
        ns = _word_namespace(root)
        if ns is None:
            return None
        def word(local):
            return Q(ns, local)
        tables = list(root.iter(word("tbl")))
        if table < 1 or table > len(tables):
            return None
        target_table = tables[table - 1]
        if _table_hazard_reason(target_table, ns, allow_merged=True) is not None:
            return None
        tbl_grid = target_table.find(word("tblGrid"))
        if tbl_grid is None:
            return None
        grid_cols = tbl_grid.findall(word("gridCol"))
        if column < 1 or column > len(grid_cols):
            return None
        return _column_structure_fingerprint(target_table, column, ns, part=part)
    except (PackageError, OSError, ET.ParseError, KeyError, IndexError):
        return None


@dataclass(frozen=True)
class TableColumnTarget:
    """One discovered candidate table column that can serve as an authoring prototype or deletion target."""

    table: int
    column: int
    rows_count: int
    structure_fingerprint_sha256: str
    cell_texts: tuple[str, ...]
    width: int | None
    part: str = _DOCUMENT
    actions: tuple[str, ...] = ("insert", "delete")

    @property
    def target(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "column": self.column,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "column": self.column,
            "rows_count": self.rows_count,
            "cell_texts": list(self.cell_texts),
            "width": self.width,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "target": self.target,
            "actions": list(self.actions),
        }


@dataclass(frozen=True)
class TableColumnRequest:
    """One source-bound request to insert, delete, or resize a table column.

    ``width`` is in twips (dxa). For ``action="resize"`` it is the new width of
    the target column; its range is checked by ``author_table_column``, which
    returns a content-safe ``table_column_width_invalid`` refusal.
    """

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    position: str | None = None
    width: int | None = None
    cell_texts: tuple[str, ...] | None = None
    action: str = "insert"

    def __post_init__(self) -> None:
        if not isinstance(self.action, str) or self.action not in _ACTIONS:
            raise ValueError("table_column_action_invalid")
        if self.action == "resize":
            if self.position is not None or self.cell_texts is not None:
                raise ValueError("table_column_request_fields_invalid")
            if self.width is None:
                raise ValueError("table_column_request_fields_invalid")
        elif self.action == "delete":
            if self.position is not None:
                raise ValueError("table_column_request_fields_invalid")
            if self.cell_texts is not None:
                raise ValueError("table_column_request_fields_invalid")
            if self.width is not None:
                raise ValueError("table_column_request_fields_invalid")
        else:
            if self.position is None:
                object.__setattr__(self, "position", "after")
            elif not isinstance(self.position, str) or self.position not in (
                "before",
                "after",
            ):
                raise ValueError("table_column_position_invalid")
            if self.width is not None and (
                not isinstance(self.width, int) or isinstance(self.width, bool) or self.width <= 0
            ):
                raise ValueError("table_column_width_invalid")
            if self.cell_texts is not None:
                if not isinstance(self.cell_texts, tuple):
                    try:
                        object.__setattr__(self, "cell_texts", tuple(self.cell_texts))
                    except TypeError:
                        raise TypeError("table_column_cell_texts_invalid")
                for text in self.cell_texts:
                    if not isinstance(text, str):
                        raise TypeError("table_column_cell_text_invalid")
                    if len(text) > _MAX_TEXT_CHARACTERS:
                        raise ValueError("table_column_cell_text_limit_exceeded")
                    if _has_illegal_xml_chars(text):
                        raise ValueError("table_column_cell_text_invalid")

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def column(self) -> int:
        val = self.target.get("column")
        if val is None:
            val = self.target.get("col")
        return val

    @property
    def col(self) -> int:
        return self.column

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "part": self.part,
            "table": self.table,
            "column": self.column,
        }
        if self.action == "insert" and self.position is not None:
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
                "column": self.column,
            },
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
        }
        if self.action == "delete":
            data["action"] = "delete"
            return data
        if self.action == "resize":
            data["action"] = "resize"
            data["width"] = self.width
            return data
        data["position"] = self.position
        if self.width is not None:
            data["width"] = self.width
        if self.cell_texts is not None:
            data["cell_texts"] = list(self.cell_texts)
        return data


@dataclass(frozen=True)
class TableColumnReport:
    """Content-safe public result for a table column authoring request."""

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
            "operation": "table-column",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _table_column_refusal(
    request: TableColumnRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TableColumnReport:
    return TableColumnReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_column_request(value: Any) -> TableColumnRequest:
    """Parse and strictly validate a table column authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_column_request_invalid")
    keys = set(value.keys())
    action = value.get("action", "insert")
    if not isinstance(action, str) or action not in _ACTIONS:
        raise ValueError("table_column_action_invalid")

    if action == "delete":
        if not _REQUIRED_REQUEST_FIELDS.issubset(keys):
            raise ValueError("table_column_request_missing_fields")
        if not keys.issubset(_ALLOWED_DELETE_REQUEST_FIELDS):
            raise ValueError("table_column_request_extra_fields")
    elif action == "resize":
        if not _REQUIRED_REQUEST_FIELDS.issubset(keys) or not keys & {"width", "column_width"}:
            raise ValueError("table_column_request_missing_fields")
        if not keys.issubset(_ALLOWED_RESIZE_REQUEST_FIELDS):
            raise ValueError("table_column_request_extra_fields")
    else:
        if not _REQUIRED_REQUEST_FIELDS.issubset(keys):
            raise ValueError("table_column_request_missing_fields")
        if not keys.issubset(_ALLOWED_REQUEST_FIELDS):
            raise ValueError("table_column_request_extra_fields")

    source = Path(value["source"])
    destination = Path(value["destination"])
    source_sha256 = _require_sha256(value["source_sha256"], "table_column_source_sha256_invalid")
    structure_fingerprint_sha256 = _require_sha256(
        value["structure_fingerprint_sha256"],
        "table_column_structure_fingerprint_invalid",
    )

    target_raw = value["target"]
    if not isinstance(target_raw, Mapping):
        raise TypeError("table_column_target_invalid")
    t_keys = set(target_raw.keys())
    has_col = "column" in t_keys or "col" in t_keys
    if "table" not in t_keys or not has_col:
        raise ValueError("table_column_target_missing_fields")
    if not t_keys.issubset(_ALLOWED_TARGET_FIELDS):
        raise ValueError("table_column_target_extra_fields")

    table_idx = target_raw["table"]
    if not isinstance(table_idx, int) or isinstance(table_idx, bool) or table_idx < 1:
        raise ValueError("table_column_target_invalid")
    col_idx = target_raw.get("column", target_raw.get("col"))
    if not isinstance(col_idx, int) or isinstance(col_idx, bool) or col_idx < 1:
        raise ValueError("table_column_target_invalid")
    part = target_raw.get("part", _DOCUMENT)
    if not isinstance(part, str):
        raise TypeError("table_column_part_invalid")

    target = {"table": table_idx, "column": col_idx, "part": part}

    position = value.get("position")
    width = value.get("width", value.get("column_width"))
    cell_texts_raw = value.get("cell_texts")

    cell_texts: tuple[str, ...] | None = None
    if cell_texts_raw is not None:
        if not isinstance(cell_texts_raw, (list, tuple)):
            raise TypeError("table_column_cell_texts_invalid")
        texts_list = []
        for item in cell_texts_raw:
            if not isinstance(item, str):
                raise TypeError("table_column_cell_text_invalid")
            if len(item) > _MAX_TEXT_CHARACTERS:
                raise ValueError("table_column_cell_text_limit_exceeded")
            if _has_illegal_xml_chars(item):
                raise ValueError("table_column_cell_text_invalid")
            texts_list.append(item)
        cell_texts = tuple(texts_list)

    return TableColumnRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha256,
        target=target,
        structure_fingerprint_sha256=structure_fingerprint_sha256,
        position=position,
        width=width,
        cell_texts=cell_texts,
        action=action,
    )


def load_table_column_request(path: str | Path) -> TableColumnRequest:
    """Load and strictly validate one JSON table-column request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_column_request(json.load(stream))


def discover_table_column_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table column targets for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-column-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_column_source_missing",
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
            "operation": "table-column-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_column_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    valid_part, part_err = _check_reachable_part(graph, part, "table_column")
    if not valid_part:
        return {
            "schema_version": 1,
            "operation": "table-column-targets",
            "passed": False,
            "status": _REFUSED,
            "code": part_err or "table_column_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "table-column-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_column_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-column-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_column_namespace_unsupported",
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
        if _table_hazard_reason(table, namespace, allow_merged=True) is not None:
            continue

        tbl_grid = table.find(Q(namespace, "tblGrid"))
        assert tbl_grid is not None
        grid_cols = tbl_grid.findall(Q(namespace, "gridCol"))
        rows = table.findall(Q(namespace, "tr"))
        num_cols = len(grid_cols)
        mapped_rows = [_grid_row_cells(row, namespace) for row in rows]
        merged = any(
            start != end
            or cell.find(f"{Q(namespace, 'tcPr')}/{Q(namespace, 'hMerge')}") is not None
            or cell.find(f"{Q(namespace, 'tcPr')}/{Q(namespace, 'vMerge')}") is not None
            for mapping in mapped_rows
            for cell, start, end in mapping
        )

        row_context_blocked = any(
            unsupported_context.get(id(row), False)
            or not _is_approved_row(row, namespace, allow_merged=True)
            for row in rows
        )
        if row_context_blocked:
            continue

        for col_idx in range(1, num_cols + 1):
            grid_col = grid_cols[col_idx - 1]
            col_width = _get_element_width(grid_col, namespace)
            if col_width is None and rows:
                first_row_cell = next(
                    cell for cell, start, end in mapped_rows[0] if start <= col_idx <= end
                )
                tc_w = first_row_cell.find(f"{Q(namespace, 'tcPr')}/{Q(namespace, 'tcW')}")
                if tc_w is not None:
                    col_width = _get_element_width(tc_w, namespace)

            cell_texts: list[str] = []
            for mapping in mapped_rows:
                cell = next(cell for cell, start, end in mapping if start <= col_idx <= end)
                cell_texts.append("".join(node.text or "" for node in cell.iter(Q(namespace, "t"))))

            fp = _column_structure_fingerprint(table, col_idx, namespace, part=part)
            can_delete = num_cols > 1
            actions = (
                ("resize",) if merged else (("insert", "delete") if can_delete else ("insert",))
            )

            target = TableColumnTarget(
                table=table_idx,
                column=col_idx,
                rows_count=len(rows),
                structure_fingerprint_sha256=fp,
                cell_texts=tuple(cell_texts),
                width=col_width,
                part=part,
                actions=actions,
            )
            targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-column-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def _resize_table_column(
    request: TableColumnRequest,
    *,
    root: ET.Element,
    table: ET.Element,
    tables_count: int,
    original_cell_xmls_by_row: list[list[bytes]],
    namespace: str,
) -> TableColumnReport:
    """Set one admitted column's grid and cell widths while keeping every cell's content intact."""
    def word(local):
        return Q(namespace, local)
    width = request.width
    assert width is not None
    col_idx = request.column - 1
    tbl_grid = table.find(word("tblGrid"))
    assert tbl_grid is not None
    grid_cols = tbl_grid.findall(word("gridCol"))
    rows = table.findall(word("tr"))
    original_grid_xmls = [ET.tostring(grid_col) for grid_col in grid_cols]
    mapped_rows = [_grid_row_cells(row, namespace) for row in rows]
    affected = [
        (r_idx, c_idx, cell, start, end)
        for r_idx, mapping in enumerate(mapped_rows)
        for c_idx, (cell, start, end) in enumerate(mapping)
        if start <= request.column <= end
    ]
    original_signatures = {
        (r_idx, c_idx): _cell_content_signature(cell, namespace)
        for r_idx, c_idx, cell, _, _ in affected
    }
    tc_w_path = f"{word('tcPr')}/{word('tcW')}"

    target_grid_col = grid_cols[col_idx]
    width_before = _get_element_width(target_grid_col, namespace)
    grid_widths = [_get_element_width(col, namespace) for col in grid_cols]
    if width_before is None:
        if affected[0][3] != affected[0][4]:
            return _table_column_refusal(
                request,
                "table_column_grid_mismatch",
                {**request.target_dict(), "row": affected[0][0] + 1, "cell": affected[0][1] + 1},
            )
        first_tc_w = affected[0][2].find(tc_w_path)
        if first_tc_w is not None and (first_tc_w.get(word("type")) or "dxa") == "dxa":
            width_before = _get_element_width(first_tc_w, namespace)
    grid_widths[col_idx] = width

    # 1. Synchronize tblGrid
    target_grid_col.attrib.pop("w", None)
    target_grid_col.set(word("w"), str(width))

    # 2. Synchronize the target cell of every row; tcPr is always the first tc child
    cell_widths = {}
    for r_idx, c_idx, target_cell, start, end in affected:
        widths = grid_widths[start - 1 : end]
        if any(value is None or value <= 0 for value in widths):
            return _table_column_refusal(
                request,
                "table_column_grid_mismatch",
                {**request.target_dict(), "row": r_idx + 1, "cell": c_idx + 1},
            )
        cell_width = sum(widths)
        cell_widths[r_idx, c_idx] = cell_width
        tc_pr = target_cell.find(word("tcPr"))
        if tc_pr is None:
            tc_pr = ET.Element(word("tcPr"))
            target_cell.insert(0, tc_pr)
        tc_w = tc_pr.find(word("tcW"))
        if tc_w is None:
            tc_w = ET.Element(word("tcW"))
            _insert_tcpr_child_in_order(tc_pr, tc_w, namespace)
        tc_w.attrib.pop("w", None)
        tc_w.attrib.pop("type", None)
        tc_w.set(word("w"), str(cell_width))
        tc_w.set(word("type"), "dxa")

    # 3. Synchronize table width by the column delta if explicit dxa
    table_width_before: int | None = None
    table_width_after: int | None = None
    tbl_w = table.find(f"{word('tblPr')}/{word('tblW')}")
    if tbl_w is not None and width_before is not None:
        w_type = tbl_w.get(word("type")) or tbl_w.get("type")
        w_val = tbl_w.get(word("w")) or tbl_w.get("w")
        if (w_type is None or w_type == "dxa") and w_val and w_val.isdigit():
            table_width_before = int(w_val)
            table_width_after = max(0, table_width_before + width - width_before)
            tbl_w.set(word("w"), str(table_width_after))

    target_change_report = {
        "action": "resize",
        "table": request.table,
        "column": request.column,
        "columns_count": len(grid_cols),
        "rows_count": len(rows),
        "width_before": width_before,
        "width_after": width,
        "table_width_before": table_width_before,
        "table_width_after": table_width_after,
    }
    failed = "table_column_staged_verification_failed"

    def verify_resize_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(request.source, staged, frozenset({request.part})):
            return failed
        if not inspect_docx(staged).valid:
            return failed
        try:
            s_root = ET.fromstring(PackageGraph.from_file(staged).read_part_bytes(request.part))
            s_tables = list(s_root.iter(word("tbl")))
            if len(s_tables) != tables_count:
                return failed
            s_table = s_tables[request.table - 1]
            s_grid = s_table.find(word("tblGrid"))
            s_cols = s_grid.findall(word("gridCol")) if s_grid is not None else []
            if len(s_cols) != len(original_grid_xmls):
                return failed
            for g_i, (orig_xml, s_col) in enumerate(zip(original_grid_xmls, s_cols, strict=True)):
                if g_i == col_idx:
                    if _get_element_width(s_col, namespace) != width:
                        return failed
                elif ET.tostring(s_col) != orig_xml:
                    return failed
            s_rows = s_table.findall(word("tr"))
            if len(s_rows) != len(rows):
                return failed
            for r_i, s_r in enumerate(s_rows):
                s_cells = s_r.findall(word("tc"))
                orig_xmls = original_cell_xmls_by_row[r_i]
                if len(s_cells) != len(orig_xmls):
                    return failed
                for c_i, (orig_xml, s_cell) in enumerate(zip(orig_xmls, s_cells, strict=True)):
                    if (r_i, c_i) not in cell_widths:
                        if ET.tostring(s_cell) != orig_xml:
                            return failed
                        continue
                    if _cell_content_signature(s_cell, namespace) != original_signatures[r_i, c_i]:
                        return failed
                    s_tc_w = s_cell.find(tc_w_path)
                    if (
                        s_tc_w is None
                        or _get_element_width(s_tc_w, namespace) != cell_widths[r_i, c_i]
                        or s_tc_w.get(word("type")) != "dxa"
                    ):
                        return failed
            if table_width_after is not None:
                s_tbl_w = s_table.find(f"{word('tblPr')}/{word('tblW')}")
                if s_tbl_w is None or _get_element_width(s_tbl_w, namespace) != table_width_after:
                    return failed
        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return failed
        return None

    published: ReviewReport = _publish(
        source=request.source,
        output=request.destination,
        source_sha256=request.source_sha256,
        replacements={request.part: serialize_xml_with_preserved_qnames(root)},
        revisions=[],
        targets=[target_change_report],
        verify=verify_resize_publication,
        failure_code=failed,
        exists_code="table_column_destination_occupied",
        unavailable_code="table_column_publication_unavailable",
    )
    if not published.passed:
        return _table_column_refusal(request, published.code or failed)

    return TableColumnReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )


def author_table_column(request: TableColumnRequest) -> TableColumnReport:
    """Insert, delete, or resize a source-bound table column atomically with full geometry synchronization."""
    source = request.source
    destination = request.destination

    if request.action == "resize" and not _is_valid_column_width(request.width):
        return _table_column_refusal(request, "table_column_width_invalid")
    if not source.is_file():
        return _table_column_refusal(request, "table_column_source_missing")
    if source == destination:
        return _table_column_refusal(request, "table_column_destination_invalid")
    if not destination.parent.is_dir():
        return _table_column_refusal(request, "table_column_destination_invalid")
    if destination.exists() or destination.is_symlink():
        return _table_column_refusal(request, "table_column_destination_occupied")

    try:
        actual_source_sha256 = _sha256_file(source)
    except OSError:
        return _table_column_refusal(request, "table_column_source_missing")

    if actual_source_sha256 != request.source_sha256:
        return _table_column_refusal(request, "table_column_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_column_refusal(request, "table_column_package_invalid")

    if _DOCUMENT not in graph.parts:
        return _table_column_refusal(request, "table_column_package_invalid")

    valid_part, part_err = _check_reachable_part(graph, request.part, "table_column")
    if not valid_part:
        return _table_column_refusal(request, part_err or "table_column_part_unsupported")

    try:
        doc_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (KeyError, ET.ParseError):
        return _table_column_refusal(request, "table_column_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_column_refusal(request, "table_column_namespace_unsupported")

    unsupported_context = _unsupported_context_by_element(root, namespace)
    def word(local):
        return Q(namespace, local)
    tables = list(root.iter(word("tbl")))

    if request.table < 1 or request.table > len(tables):
        return _table_column_refusal(request, "table_column_target_missing")

    table = tables[request.table - 1]
    if unsupported_context.get(id(table), False):
        return _table_column_refusal(request, "table_column_context_unsupported")

    hazard = _table_hazard_reason(table, namespace, allow_merged=request.action == "resize")
    if hazard is not None:
        target = request.target_dict()
        for r_idx, row in enumerate(table.findall(word("tr")), start=1):
            for c_idx, cell in enumerate(row.findall(word("tc")), start=1):
                properties = cell.find(word("tcPr"))
                if properties is not None and any(
                    properties.find(word(name)) is not None
                    for name in ("gridSpan", "hMerge", "vMerge")
                ):
                    target.update(row=r_idx, cell=c_idx)
                    return _table_column_refusal(request, hazard, target)
        return _table_column_refusal(request, hazard)

    tbl_grid = table.find(word("tblGrid"))
    assert tbl_grid is not None
    grid_cols = tbl_grid.findall(word("gridCol"))
    rows = table.findall(word("tr"))
    num_cols = len(grid_cols)
    original_cell_xmls_by_row: list[list[bytes]] = [
        [ET.tostring(c) for c in r.findall(word("tc"))] for r in rows
    ]

    if request.column < 1 or request.column > num_cols:
        return _table_column_refusal(request, "table_column_target_missing")

    actual_fp = _column_structure_fingerprint(table, request.column, namespace, part=request.part)
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_column_refusal(request, "table_column_structure_fingerprint_mismatch")

    if request.action == "resize":
        return _resize_table_column(
            request,
            root=root,
            table=table,
            tables_count=len(tables),
            original_cell_xmls_by_row=original_cell_xmls_by_row,
            namespace=namespace,
        )

    target_col_idx = request.column - 1
    target_grid_col = grid_cols[target_col_idx]

    expected_surviving_cells_by_row: list[list[bytes]] = [
        [cell_xml for c_i, cell_xml in enumerate(row_xmls) if c_i != target_col_idx]
        for row_xmls in original_cell_xmls_by_row
    ]
    if request.action == "delete":
        if num_cols <= 1:
            return _table_column_refusal(request, "table_column_final_column_unsupported")

        deleted_width = _get_element_width(target_grid_col, namespace)
        if deleted_width is None:
            first_tc = rows[0].findall(word("tc"))[target_col_idx]
            tc_w = first_tc.find(f".//{word('tcW')}")
            deleted_width = _get_element_width(tc_w, namespace) if tc_w is not None else 0

        # 1. Synchronize tblGrid
        tbl_grid.remove(target_grid_col)

        # 2. Synchronize every row
        for r in rows:
            r_cells = r.findall(word("tc"))
            r.remove(r_cells[target_col_idx])

        # 3. Synchronize table width if explicit dxa
        tbl_pr = table.find(word("tblPr"))
        if tbl_pr is not None:
            tbl_w = tbl_pr.find(word("tblW"))
            if tbl_w is not None and deleted_width:
                w_type = tbl_w.get(word("type")) or tbl_w.get("type")
                w_val = tbl_w.get(word("w")) or tbl_w.get("w")
                if (w_type is None or w_type == "dxa") and w_val and w_val.isdigit():
                    new_tbl_w = max(0, int(w_val) - deleted_width)
                    tbl_w.set(word("w"), str(new_tbl_w))

        target_change_report = {
            "action": "delete",
            "table": request.table,
            "deleted_column": request.column,
            "columns_count_before": num_cols,
            "columns_count_after": num_cols - 1,
            "rows_count": len(rows),
            "width": deleted_width,
        }

        def verify_deletion_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, frozenset({request.part})):
                return "table_column_staged_verification_failed"
            insp = inspect_docx(staged)
            if not insp.valid:
                return "table_column_staged_verification_failed"
            try:
                s_graph = PackageGraph.from_file(staged)
                s_root = ET.fromstring(s_graph.read_part_bytes(request.part))
                s_tables = list(s_root.iter(word("tbl")))
                if len(s_tables) != len(tables):
                    return "table_column_staged_verification_failed"
                s_table = s_tables[request.table - 1]
                s_grid = s_table.find(word("tblGrid"))
                if s_grid is None:
                    return "table_column_staged_verification_failed"
                s_cols = s_grid.findall(word("gridCol"))
                if len(s_cols) != num_cols - 1:
                    return "table_column_staged_verification_failed"
                s_rows = s_table.findall(word("tr"))
                if len(s_rows) != len(rows):
                    return "table_column_staged_verification_failed"
                for r_i, s_r in enumerate(s_rows):
                    s_cells = s_r.findall(word("tc"))
                    expected_survivors = expected_surviving_cells_by_row[r_i]
                    if len(s_cells) != len(expected_survivors):
                        return "table_column_staged_verification_failed"
                    for orig_xml, staged_c in zip(expected_survivors, s_cells, strict=True):
                        if ET.tostring(staged_c) != orig_xml:
                            return "table_column_staged_verification_failed"
            except (
                OSError,
                PackageError,
                ET.ParseError,
                KeyError,
                IndexError,
                ValueError,
            ):
                return "table_column_staged_verification_failed"
            return None

        serialized = serialize_xml_with_preserved_qnames(root)
        replacements = {request.part: serialized}

        published: ReviewReport = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[target_change_report],
            verify=verify_deletion_publication,
            failure_code="table_column_staged_verification_failed",
            exists_code="table_column_destination_occupied",
            unavailable_code="table_column_publication_unavailable",
        )

        if not published.passed:
            return _table_column_refusal(
                request, published.code or "table_column_staged_verification_failed"
            )

        return TableColumnReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            target=request.target_dict(),
            target_change_report=target_change_report,
            non_target_parts_preserved=True,
        )

    # Insertion path
    if request.cell_texts is not None and len(request.cell_texts) != len(rows):
        return _table_column_refusal(request, "table_column_cell_texts_mismatch")

    if request.width is not None:
        inserted_width = request.width
    else:
        inferred_w = _get_element_width(target_grid_col, namespace)
        if inferred_w is None:
            first_tc = rows[0].findall(word("tc"))[target_col_idx]
            tc_w = first_tc.find(f".//{word('tcW')}")
            inferred_w = _get_element_width(tc_w, namespace) if tc_w is not None else None
        inserted_width = inferred_w if inferred_w is not None else 2500

    insertion_texts: Sequence[str] = (
        request.cell_texts if request.cell_texts is not None else ("",) * len(rows)
    )

    if request.position == "before":
        ins_grid_idx = target_col_idx
        inserted_col_num = request.column
    else:
        ins_grid_idx = target_col_idx + 1
        inserted_col_num = request.column + 1

    # 1. Synchronize tblGrid
    new_grid_col = ET.Element(word("gridCol"))
    new_grid_col.set(word("w"), str(inserted_width))
    tbl_grid.insert(ins_grid_idx, new_grid_col)

    # 2. Synchronize every row
    for r_idx, r in enumerate(rows):
        r_cells = r.findall(word("tc"))
        proto_cell = r_cells[target_col_idx]
        proto_tc_pr = proto_cell.find(word("tcPr"))

        new_cell = ET.Element(word("tc"))
        if proto_tc_pr is not None:
            new_tc_pr = copy.deepcopy(proto_tc_pr)
        else:
            new_tc_pr = ET.Element(word("tcPr"))

        # Set explicit tcW maintaining schema ordering in tcPr
        tc_w = new_tc_pr.find(word("tcW"))
        if tc_w is None:
            tc_w = ET.Element(word("tcW"))
            _insert_tcpr_child_in_order(new_tc_pr, tc_w, namespace)
        tc_w.set(word("w"), str(inserted_width))
        tc_w.set(word("type"), "dxa")

        # Purge any merge elements
        for tag_name in ("vMerge", "hMerge", "gridSpan"):
            m_elem = new_tc_pr.find(word(tag_name))
            if m_elem is not None:
                new_tc_pr.remove(m_elem)

        new_cell.append(new_tc_pr)

        # Paragraph & Run with cloned formatting
        proto_p = proto_cell.find(word("p"))
        new_p = ET.Element(word("p"))
        if proto_p is not None:
            proto_p_pr = proto_p.find(word("pPr"))
            if proto_p_pr is not None:
                new_p.append(copy.deepcopy(proto_p_pr))

        new_r = ET.Element(word("r"))
        if proto_p is not None:
            proto_r = proto_p.find(word("r"))
            if proto_r is not None:
                proto_r_pr = proto_r.find(word("rPr"))
                if proto_r_pr is not None:
                    new_r.append(copy.deepcopy(proto_r_pr))

        t_elem = ET.Element(word("t"))
        t_text = insertion_texts[r_idx]
        t_elem.text = t_text
        if t_text.startswith(" ") or t_text.endswith(" ") or "  " in t_text:
            t_elem.set(_XML_SPACE_ATTRIBUTE, "preserve")

        new_r.append(t_elem)
        new_p.append(new_r)
        new_cell.append(new_p)

        # Insert at exact position relative to proto_cell child index
        r_children = list(r)
        proto_child_idx = r_children.index(proto_cell)
        if request.position == "before":
            r.insert(proto_child_idx, new_cell)
        else:
            r.insert(proto_child_idx + 1, new_cell)

    # 3. Synchronize table width if explicit dxa
    tbl_pr = table.find(word("tblPr"))
    if tbl_pr is not None:
        tbl_w = tbl_pr.find(word("tblW"))
        if tbl_w is not None and inserted_width:
            w_type = tbl_w.get(word("type")) or tbl_w.get("type")
            w_val = tbl_w.get(word("w")) or tbl_w.get("w")
            if (w_type is None or w_type == "dxa") and w_val and w_val.isdigit():
                new_tbl_w = int(w_val) + inserted_width
                tbl_w.set(word("w"), str(new_tbl_w))

    target_change_report = {
        "action": "insert",
        "table": request.table,
        "prototype_column": request.column,
        "position": request.position,
        "inserted_column": inserted_col_num,
        "columns_count_before": num_cols,
        "columns_count_after": num_cols + 1,
        "rows_count": len(rows),
        "width": inserted_width,
        "cell_text_lengths": [len(t) for t in insertion_texts],
        "cell_texts_sha256": hashlib.sha256("".join(insertion_texts).encode("utf-8")).hexdigest(),
    }

    def verify_insertion_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({request.part})):
            return "table_column_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_column_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(request.part))
            s_tables = list(s_root.iter(word("tbl")))
            if len(s_tables) != len(tables):
                return "table_column_staged_verification_failed"
            s_table = s_tables[request.table - 1]
            s_grid = s_table.find(word("tblGrid"))
            if s_grid is None:
                return "table_column_staged_verification_failed"
            s_cols = s_grid.findall(word("gridCol"))
            if len(s_cols) != num_cols + 1:
                return "table_column_staged_verification_failed"
            ins_grid_col = s_cols[ins_grid_idx]
            if _get_element_width(ins_grid_col, namespace) != inserted_width:
                return "table_column_staged_verification_failed"
            s_rows = s_table.findall(word("tr"))
            if len(s_rows) != len(rows):
                return "table_column_staged_verification_failed"
            for r_i, (orig_r, s_r) in enumerate(zip(rows, s_rows)):
                s_cells = s_r.findall(word("tc"))
                if len(s_cells) != num_cols + 1:
                    return "table_column_staged_verification_failed"
                inserted_cell = s_cells[ins_grid_idx]
                simple = _simple_text_node(inserted_cell, namespace)
                if simple is None or simple[1] != insertion_texts[r_i]:
                    return "table_column_staged_verification_failed"
                tc_w_elem = inserted_cell.find(f".//{word('tcW')}")
                if tc_w_elem is None or _get_element_width(tc_w_elem, namespace) != inserted_width:
                    return "table_column_staged_verification_failed"
                orig_xmls = original_cell_xmls_by_row[r_i]
                staged_non_target = [c for i, c in enumerate(s_cells) if i != ins_grid_idx]
                if len(staged_non_target) != len(orig_xmls):
                    return "table_column_staged_verification_failed"
                for orig_xml, staged_c in zip(orig_xmls, staged_non_target, strict=True):
                    if ET.tostring(staged_c) != orig_xml:
                        return "table_column_staged_verification_failed"
        except (
            OSError,
            PackageError,
            ET.ParseError,
            KeyError,
            IndexError,
            ValueError,
        ):
            return "table_column_staged_verification_failed"
        return None

    serialized = serialize_xml_with_preserved_qnames(root)
    replacements = {request.part: serialized}

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=[target_change_report],
        verify=verify_insertion_publication,
        failure_code="table_column_staged_verification_failed",
        exists_code="table_column_destination_occupied",
        unavailable_code="table_column_publication_unavailable",
    )

    if not published.passed:
        return _table_column_refusal(
            request, published.code or "table_column_staged_verification_failed"
        )

    return TableColumnReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
