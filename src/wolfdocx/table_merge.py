"""Source-bound table cell merge and split authoring and discovery for WordprocessingML documents.

Supports horizontal gridSpan merges, vertical vMerge restarts and continuations,
and combined rectangular merges on explicit grid rectangles in supported tables of
the main document body or of a genuinely referenced header or footer part.
Guarantees zero data loss by refusing merges over non-empty non-origin cells
(including whitespace, tabs, breaks, and control symbols), preserves tblGrid unchanged,
derives split cell widths from explicit positive grid widths, and strictly preserves
OOXML schema element ordering in tcPr.

Legacy ``w:hMerge`` regions are discoverable as split targets but are never mutated:
any request against a table that carries ``w:hMerge`` is refused with
``table_merge_structure_unsupported`` and the refusal's ``target_change_report`` names
the exact one-based ``row`` and ``cell`` of the offending cell. Existing gridSpan and
vMerge regions a request would cross are likewise refused with row and cell details
and are never canonicalized or deleted.
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

from .authoring import (
    _PUBLISHED,
    _REFUSED,
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
from .table_properties import _check_reachable_part

__all__ = [
    "TableMergeReport",
    "TableMergeRequest",
    "TableMergeTarget",
    "author_table_merge",
    "discover_table_merge_targets",
    "load_table_merge_request",
    "parse_table_merge_request",
    "table_merge_structure_fingerprint",
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
        "action",
    }
)
_REQUIRED_TARGET_FIELDS = frozenset({"table", "start_row", "end_row", "start_col", "end_col"})
_ALLOWED_TARGET_FIELDS = frozenset(
    {"table", "start_row", "end_row", "start_col", "end_col", "part"}
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

_FORBIDDEN_CELL_ELEMENTS = frozenset(
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

_USER_CONTENT_LOCAL_TAGS = frozenset(
    {
        "tab",
        "br",
        "cr",
        "sym",
        "noBreakHyphen",
        "softHyphen",
        "drawing",
        "pict",
        "tbl",
        "hyperlink",
        "fldSimple",
        "fldChar",
        "instrText",
        "footnoteReference",
        "endnoteReference",
        "commentReference",
        "subDoc",
        "ptab",
        "dayShort",
        "monthShort",
        "yearShort",
        "dayLong",
        "monthLong",
        "yearLong",
        "annotationRef",
        "footnoteRef",
        "endnoteRef",
        "separator",
        "continuationSeparator",
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


def _has_forbidden_cell_descendant(element: ET.Element) -> bool:
    """True if element or any descendant has forbidden elements or IDs."""
    for child in element.iter():
        local = child.tag.rsplit("}", 1)[-1]
        if local in _FORBIDDEN_CELL_ELEMENTS:
            return True
        for attr in child.attrib:
            attr_local = attr.rsplit("}", 1)[-1]
            if attr_local == "id":
                return True
    return False


def _cell_has_user_content(cell: ET.Element, namespace: str) -> bool:
    """Return True if the cell contains user text, whitespace, control characters, or non-text payload."""
    def word(local):
        return Q(namespace, local)
    for t_node in cell.iter(word("t")):
        if t_node.text is not None and len(t_node.text) > 0:
            return True
    for child in cell.iter():
        local = child.tag.rsplit("}", 1)[-1]
        if local in _USER_CONTENT_LOCAL_TAGS:
            return True
    return len(cell.findall(word("p"))) > 1


def _get_cell_text(cell: ET.Element, namespace: str) -> str:
    """Extract plain text from all w:t descendants in the cell."""
    def word(local):
        return Q(namespace, local)
    texts: list[str] = []
    for t_node in cell.iter(word("t")):
        if t_node.text:
            texts.append(t_node.text)
    return "".join(texts)


def _cell_structure_projection(element: ET.Element) -> list[Any]:
    """Mask text inside w:t while preserving element hierarchy, attributes, and non-text values."""
    return [
        element.tag,
        sorted(element.attrib.items()),
        None if element.tag.rsplit("}", 1)[-1] == "t" else element.text,
        element.tail,
        [_cell_structure_projection(child) for child in element],
    ]


def _is_supported_table(table: ET.Element, namespace: str) -> bool:
    """True if table has supported body table structure with explicit grid columns."""
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


def _cell_hmerge_mode(cell: ET.Element, namespace: str) -> str | None:
    """Return the legacy hMerge mode ('restart', 'continue') or None when absent.

    Raises ValueError on duplicate or invalid hMerge markup.
    """
    def word(local):
        return Q(namespace, local)
    tc_pr = cell.find(word("tcPr"))
    if tc_pr is None:
        return None
    h_merges = tc_pr.findall(word("hMerge"))
    if not h_merges:
        return None
    if len(h_merges) > 1:
        raise ValueError("duplicate_hMerge")
    val = h_merges[0].get(word("val"))
    if val is None or val == "continue":
        return "continue"
    if val == "restart":
        return "restart"
    raise ValueError("invalid_hMerge_val")


def _cell_malformed_reason(
    cell: ET.Element, namespace: str, *, legacy_hmerge: bool = False
) -> str | None:
    """Name the first malformed or unsupported merge marker of a cell, else None.

    A legacy hMerge is reported as ``hMerge_unsupported`` unless ``legacy_hmerge`` admits it,
    in which case it must be a well-formed, single-column marker.
    """
    def word(local):
        return Q(namespace, local)
    tc_pr = cell.find(word("tcPr"))
    if tc_pr is None:
        return None
    try:
        span, _ = _cell_merge_state(cell, namespace, legacy_hmerge=legacy_hmerge)
        if legacy_hmerge and tc_pr.find(word("hMerge")) is not None:
            _cell_hmerge_mode(cell, namespace)
            if span != 1:
                return "hMerge_gridSpan_unsupported"
    except ValueError as error:
        return str(error)
    return None


def _row_legacy_hmerge_runs(row: ET.Element, namespace: str) -> list[tuple[int, int]] | None:
    """Return zero-based inclusive (first, last) cell indexes of every legacy hMerge run.

    A run is a restart cell followed by one or more continue cells. Returns None when the
    row's hMerge markup is malformed or a continue cell has no restart before it.
    """
    runs: list[tuple[int, int]] = []
    start: int | None = None
    last = 0
    for index, cell in enumerate(row.findall(Q(namespace, "tc"))):
        try:
            mode = _cell_hmerge_mode(cell, namespace)
        except ValueError:
            return None
        if mode == "continue":
            if start is None:
                return None
            last = index
            continue
        if start is not None and last > start:
            runs.append((start, last))
        start = index if mode == "restart" else None
        last = index
    if start is not None and last > start:
        runs.append((start, last))
    return runs


def _row_has_malformed_merge_markup(
    row: ET.Element, namespace: str, *, legacy_hmerge: bool = False
) -> bool:
    """True if any cell in row has duplicate or invalid gridSpan/vMerge or unadmitted hMerge."""
    for cell in row.findall(Q(namespace, "tc")):
        if _cell_malformed_reason(cell, namespace, legacy_hmerge=legacy_hmerge) is not None:
            return True
    return legacy_hmerge and _row_legacy_hmerge_runs(row, namespace) is None


def _row_grid_mapping(
    row: ET.Element,
    num_grid_cols: int,
    namespace: str,
    *,
    legacy_hmerge: bool = False,
) -> list[tuple[ET.Element, int, int]] | None:
    """Map each cell in row to (cell, col_start, col_end) (1-based, inclusive).

    Returns None if the row's cells do not cleanly partition the grid columns.
    """
    def word(local):
        return Q(namespace, local)
    cells = row.findall(word("tc"))
    if not cells:
        return None
    mapping: list[tuple[ET.Element, int, int]] = []
    current_col = 1
    for cell in cells:
        tc_pr = cell.find(word("tcPr"))
        span = 1
        if tc_pr is not None:
            if tc_pr.find(word("hMerge")) is not None and not legacy_hmerge:
                return None
            grid_spans = tc_pr.findall(word("gridSpan"))
            if len(grid_spans) > 1:
                return None
            if grid_spans:
                val_str = grid_spans[0].get(word("val"))
                if val_str is None:
                    return None
                try:
                    span = int(val_str)
                except (ValueError, TypeError):
                    return None
                if span < 1:
                    return None
        col_start = current_col
        col_end = current_col + span - 1
        if col_end > num_grid_cols:
            return None
        mapping.append((cell, col_start, col_end))
        current_col += span
    if current_col - 1 != num_grid_cols:
        return None
    return mapping


def _cell_merge_state(
    cell: ET.Element, namespace: str, *, legacy_hmerge: bool = False
) -> tuple[int, str | None]:
    """Return (grid_span, v_merge_mode).

    v_merge_mode is 'restart', 'continue', or None.
    Raises ValueError on malformed or duplicate merge markup.
    """
    def word(local):
        return Q(namespace, local)
    tc_pr = cell.find(word("tcPr"))
    if tc_pr is None:
        return 1, None

    if tc_pr.find(word("hMerge")) is not None and not legacy_hmerge:
        raise ValueError("hMerge_unsupported")

    grid_spans = tc_pr.findall(word("gridSpan"))
    if len(grid_spans) > 1:
        raise ValueError("duplicate_gridSpan")
    span = 1
    if grid_spans:
        val_str = grid_spans[0].get(word("val"))
        if val_str is None:
            raise ValueError("missing_gridSpan_val")
        try:
            span = int(val_str)
        except (ValueError, TypeError):
            raise ValueError("invalid_gridSpan_val") from None
        if span < 1:
            raise ValueError("invalid_gridSpan_val")

    v_merges = tc_pr.findall(word("vMerge"))
    if len(v_merges) > 1:
        raise ValueError("duplicate_vMerge")
    v_merge = None
    if v_merges:
        val = v_merges[0].get(word("val"))
        if val is None or val == "continue":
            v_merge = "continue"
        elif val == "restart":
            v_merge = "restart"
        else:
            raise ValueError("invalid_vMerge_val")

    return span, v_merge


def _insert_or_replace_tcpr_child(tc_pr: ET.Element, new_child: ET.Element, namespace: str) -> None:
    """Insert or replace a child element in tcPr while maintaining schema order."""
    local_name = new_child.tag.rsplit("}", 1)[-1]
    def word(local):
        return Q(namespace, local)
    for existing in list(tc_pr.findall(word(local_name))):
        tc_pr.remove(existing)
    try:
        new_order = _TCPR_ORDER.index(local_name)
    except ValueError:
        new_order = 999
    insert_idx = len(tc_pr)
    for idx, child in enumerate(tc_pr):
        child_local = child.tag.rsplit("}", 1)[-1]
        try:
            child_order = _TCPR_ORDER.index(child_local)
        except ValueError:
            child_order = 999
        if child_order > new_order:
            insert_idx = idx
            break
    tc_pr.insert(insert_idx, new_child)


def _remove_tcpr_child(tc_pr: ET.Element, local_name: str, namespace: str) -> None:
    """Remove any child matching local_name from tcPr."""
    def word(local):
        return Q(namespace, local)
    for child in list(tc_pr.findall(word(local_name))):
        tc_pr.remove(child)


def _set_tc_width(tc_pr: ET.Element, width_val: int | str, namespace: str) -> None:
    def word(local):
        return Q(namespace, local)
    tc_w = ET.Element(word("tcW"))
    tc_w.set(word("w"), str(width_val))
    tc_w.set(word("type"), "dxa")
    _insert_or_replace_tcpr_child(tc_pr, tc_w, namespace)


def _set_grid_span(tc_pr: ET.Element, span: int, namespace: str) -> None:
    def word(local):
        return Q(namespace, local)
    gs = ET.Element(word("gridSpan"))
    gs.set(word("val"), str(span))
    _insert_or_replace_tcpr_child(tc_pr, gs, namespace)


def _set_v_merge(tc_pr: ET.Element, mode: str, namespace: str) -> None:
    def word(local):
        return Q(namespace, local)
    vm = ET.Element(word("vMerge"))
    if mode == "restart":
        vm.set(word("val"), "restart")
    _insert_or_replace_tcpr_child(tc_pr, vm, namespace)


def _create_split_cell(
    prototype_tc_pr: ET.Element | None,
    width_val: str | None,
    namespace: str,
) -> ET.Element:
    """Create a new per-grid cell inheriting applicable properties without merge attributes."""
    def word(local):
        return Q(namespace, local)
    new_tc = ET.Element(word("tc"))
    new_tc_pr = ET.SubElement(new_tc, word("tcPr"))
    if width_val is not None:
        _set_tc_width(new_tc_pr, width_val, namespace)
    if prototype_tc_pr is not None:
        applicable_tags = ("tcBorders", "shd", "tcMar", "vAlign", "noWrap")
        for tag in applicable_tags:
            child = prototype_tc_pr.find(word(tag))
            if child is not None:
                new_child = copy.deepcopy(child)
                _insert_or_replace_tcpr_child(new_tc_pr, new_child, namespace)
    ET.SubElement(new_tc, word("p"))
    return new_tc


def _validate_table_merge_target_dict(target: Mapping[str, Any]) -> dict[str, Any]:
    """Strictly validate and return an immutable copy of target coordinates."""
    if not isinstance(target, Mapping):
        raise TypeError("table_merge_target_invalid")
    target_keys = set(target)
    if not target_keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(
        target_keys
    ):
        raise ValueError("table_merge_target_fields_invalid")

    table = target["table"]
    if isinstance(table, bool) or not isinstance(table, int) or table < 1:
        raise ValueError("table_merge_table_invalid")

    start_row = target["start_row"]
    if isinstance(start_row, bool) or not isinstance(start_row, int) or start_row < 1:
        raise ValueError("table_merge_start_row_invalid")

    end_row = target["end_row"]
    if isinstance(end_row, bool) or not isinstance(end_row, int) or end_row < 1:
        raise ValueError("table_merge_end_row_invalid")

    start_col = target["start_col"]
    if isinstance(start_col, bool) or not isinstance(start_col, int) or start_col < 1:
        raise ValueError("table_merge_start_col_invalid")

    end_col = target["end_col"]
    if isinstance(end_col, bool) or not isinstance(end_col, int) or end_col < 1:
        raise ValueError("table_merge_end_col_invalid")

    if start_row > end_row or start_col > end_col:
        raise ValueError("table_merge_coordinates_invalid")

    part = target.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("table_merge_part_invalid")

    return {
        "table": table,
        "start_row": start_row,
        "end_row": end_row,
        "start_col": start_col,
        "end_col": end_col,
        "part": part,
    }


def table_merge_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for a target table rectangle."""
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        valid_part, _ = _check_reachable_part(graph, part, "table_merge")
        if not valid_part:
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
        tbl_grid = target_table.find(Q(ns, "tblGrid"))
        if tbl_grid is None:
            return None
        grid_cols = tbl_grid.findall(Q(ns, "gridCol"))
        if not grid_cols:
            return None
        rows = target_table.findall(Q(ns, "tr"))
        if start_row < 1 or end_row > len(rows) or start_row > end_row:
            return None
        if start_col < 1 or end_col > len(grid_cols) or start_col > end_col:
            return None

        binding: dict[str, Any] = {
            "table": table,
            "start_row": start_row,
            "end_row": end_row,
            "start_col": start_col,
            "end_col": end_col,
        }
        if part != _DOCUMENT:
            binding["part"] = part
        projections: list[Any] = [_cell_structure_projection(tbl_grid), binding]
        for r_idx in range(start_row, end_row + 1):
            row = rows[r_idx - 1]
            if _row_has_malformed_merge_markup(row, ns, legacy_hmerge=True):
                return None
            mapping = _row_grid_mapping(row, len(grid_cols), ns, legacy_hmerge=True)
            if mapping is None:
                return None
            for cell, c_start, c_end in mapping:
                if not (c_end < start_col or c_start > end_col):
                    projections.append([r_idx, c_start, c_end, _cell_structure_projection(cell)])

        canonical = json.dumps(projections, ensure_ascii=False, separators=(",", ":"))
        return hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    except (PackageError, OSError, zipfile.BadZipFile, ET.ParseError, KeyError, ValueError):
        return None


@dataclass(frozen=True)
class TableMergeTarget:
    """One discovered candidate table rectangle for merge or split authoring."""

    table: int
    start_row: int
    end_row: int
    start_col: int
    end_col: int
    structure_fingerprint_sha256: str
    action: str
    part: str = _DOCUMENT
    origin_text: str = ""
    is_merged: bool = False

    @property
    def target(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "start_row": self.start_row,
            "end_row": self.end_row,
            "start_col": self.start_col,
            "end_col": self.end_col,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "start_row": self.start_row,
            "end_row": self.end_row,
            "start_col": self.start_col,
            "end_col": self.end_col,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "action": self.action,
            "is_merged": self.is_merged,
            "origin_text": self.origin_text,
        }


@dataclass(frozen=True)
class TableMergeRequest:
    """One source-bound request to merge or split an explicit bounded table grid rectangle."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    action: str = "merge"

    def __post_init__(self) -> None:
        if isinstance(self.source, str):
            object.__setattr__(self, "source", Path(self.source))
        elif not isinstance(self.source, Path):
            raise TypeError("table_merge_source_invalid")
        if not str(self.source):
            raise ValueError("table_merge_source_invalid")

        if isinstance(self.destination, str):
            object.__setattr__(self, "destination", Path(self.destination))
        elif not isinstance(self.destination, Path):
            raise TypeError("table_merge_destination_invalid")
        if not str(self.destination):
            raise ValueError("table_merge_destination_invalid")

        _require_sha256(self.source_sha256, "table_merge_source_sha256")
        _require_sha256(
            self.structure_fingerprint_sha256,
            "table_merge_structure_fingerprint_sha256",
        )
        if not isinstance(self.action, str) or self.action not in ("merge", "split"):
            raise ValueError("table_merge_action_invalid")

        validated_target = _validate_table_merge_target_dict(self.target)
        object.__setattr__(self, "target", validated_target)

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def start_row(self) -> int:
        return self.target["start_row"]

    @property
    def end_row(self) -> int:
        return self.target["end_row"]

    @property
    def start_col(self) -> int:
        return self.target["start_col"]

    @property
    def end_col(self) -> int:
        return self.target["end_col"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "table": self.table,
            "start_row": self.start_row,
            "end_row": self.end_row,
            "start_col": self.start_col,
            "end_col": self.end_col,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "target": self.target_dict(),
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "action": self.action,
        }


@dataclass(frozen=True)
class TableMergeReport:
    """Content-safe public result for a table merge or split authoring request."""

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
            "operation": "table-merge",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _table_merge_refusal(
    request: TableMergeRequest,
    code: str,
    target: dict[str, Any] | None = None,
    change_report: dict[str, Any] | None = None,
) -> TableMergeReport:
    return TableMergeReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        target_change_report=change_report,
        non_target_parts_preserved=None,
    )


def parse_table_merge_request(value: Any) -> TableMergeRequest:
    """Parse and strictly validate a table cell merge or split authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_merge_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_merge_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("table_merge_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("table_merge_destination_invalid")

    action = value.get("action", "merge")
    if not isinstance(action, str) or action not in ("merge", "split"):
        raise ValueError("table_merge_action_invalid")

    return TableMergeRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=value["source_sha256"],
        target=value["target"],
        structure_fingerprint_sha256=value["structure_fingerprint_sha256"],
        action=action,
    )


def load_table_merge_request(path: str | Path) -> TableMergeRequest:
    """Load and strictly validate one JSON table-merge request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_merge_request(json.load(stream))


def discover_table_merge_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table cell merge and split targets for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-merge-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_merge_source_missing",
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
            "operation": "table-merge-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_merge_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    valid_part, part_err = _check_reachable_part(graph, part, "table_merge")
    if not valid_part:
        return {
            "schema_version": 1,
            "operation": "table-merge-targets",
            "passed": False,
            "status": _REFUSED,
            "code": part_err or "table_merge_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "table-merge-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_merge_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-merge-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_merge_namespace_unsupported",
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
        num_grid_cols = len(grid_cols)
        if num_grid_cols == 0:
            continue
        rows = table.findall(Q(namespace, "tr"))
        if not rows:
            continue

        row_mappings: list[list[tuple[ET.Element, int, int]]] = []
        table_grid_valid = True
        for row in rows:
            if unsupported_context.get(id(row), False):
                table_grid_valid = False
                break
            if _row_has_malformed_merge_markup(row, namespace, legacy_hmerge=True):
                table_grid_valid = False
                break
            mapping = _row_grid_mapping(row, num_grid_cols, namespace, legacy_hmerge=True)
            if mapping is None:
                table_grid_valid = False
                break
            row_mappings.append(mapping)
        if not table_grid_valid:
            continue

        # 0. Legacy hMerge tables: mutation is refused table-wide, so only the split regions
        # themselves are inventoried; each names the restart cell and its continue cells.
        if any(
            _cell_hmerge_mode(cell, namespace) is not None
            for mapping in row_mappings
            for cell, _, _ in mapping
        ):
            for r_idx, row in enumerate(rows, start=1):
                mapping = row_mappings[r_idx - 1]
                for first, last in _row_legacy_hmerge_runs(row, namespace) or []:
                    run_cells = mapping[first : last + 1]
                    if any(
                        _cell_merge_state(run_cell, namespace, legacy_hmerge=True)[1] is not None
                        for run_cell, _, _ in run_cells
                    ):
                        continue
                    run_start, run_end = mapping[first][1], mapping[last][2]
                    fp = table_merge_structure_fingerprint(
                        source,
                        table=table_idx,
                        start_row=r_idx,
                        end_row=r_idx,
                        start_col=run_start,
                        end_col=run_end,
                        part=part,
                    )
                    if fp is not None:
                        target = TableMergeTarget(
                            table=table_idx,
                            start_row=r_idx,
                            end_row=r_idx,
                            start_col=run_start,
                            end_col=run_end,
                            structure_fingerprint_sha256=fp,
                            action="split",
                            part=part,
                            origin_text=_get_cell_text(mapping[first][0], namespace),
                            is_merged=True,
                        )
                        targets.append(target.to_dict())
            continue

        # 1. Discover existing merges (splittable targets)
        visited_vertical: set[tuple[int, int]] = set()
        for r_idx, row in enumerate(rows, start=1):
            mapping = row_mappings[r_idx - 1]
            for cell, c_start, c_end in mapping:
                if (r_idx, c_start) in visited_vertical:
                    continue
                try:
                    span, v_merge = _cell_merge_state(cell, namespace)
                except ValueError:
                    continue
                if v_merge == "restart":
                    end_r = r_idx
                    while end_r + 1 <= len(rows):
                        next_mapping = row_mappings[end_r]
                        next_cell = None
                        for c_elem, n_cs, n_ce in next_mapping:
                            if n_cs == c_start and n_ce == c_end:
                                next_cell = c_elem
                                break
                        if next_cell is None:
                            break
                        try:
                            n_span, n_vmerge = _cell_merge_state(next_cell, namespace)
                        except ValueError:
                            break
                        if n_vmerge != "continue" or n_span != span:
                            break
                        visited_vertical.add((end_r + 1, c_start))
                        end_r += 1
                    if end_r > r_idx:
                        fp = table_merge_structure_fingerprint(
                            source,
                            table=table_idx,
                            start_row=r_idx,
                            end_row=end_r,
                            start_col=c_start,
                            end_col=c_end,
                            part=part,
                        )
                        if fp is not None:
                            target = TableMergeTarget(
                                table=table_idx,
                                start_row=r_idx,
                                end_row=end_r,
                                start_col=c_start,
                                end_col=c_end,
                                structure_fingerprint_sha256=fp,
                                action="split",
                                part=part,
                                origin_text=_get_cell_text(cell, namespace),
                                is_merged=True,
                            )
                            targets.append(target.to_dict())
                        continue

                if c_end > c_start and v_merge is None:
                    fp = table_merge_structure_fingerprint(
                        source,
                        table=table_idx,
                        start_row=r_idx,
                        end_row=r_idx,
                        start_col=c_start,
                        end_col=c_end,
                        part=part,
                    )
                    if fp is not None:
                        target = TableMergeTarget(
                            table=table_idx,
                            start_row=r_idx,
                            end_row=r_idx,
                            start_col=c_start,
                            end_col=c_end,
                            structure_fingerprint_sha256=fp,
                            action="split",
                            part=part,
                            origin_text=_get_cell_text(cell, namespace),
                            is_merged=True,
                        )
                        targets.append(target.to_dict())

        # 2. Discover unmerged merge candidates (adjacent cells without content loss)
        for r_idx, row in enumerate(rows, start=1):
            mapping = row_mappings[r_idx - 1]
            for idx in range(len(mapping) - 1):
                c1, c1_s, _c1_e = mapping[idx]
                c2, _c2_s, c2_e = mapping[idx + 1]
                try:
                    s1, vm1 = _cell_merge_state(c1, namespace)
                    s2, vm2 = _cell_merge_state(c2, namespace)
                except ValueError:
                    continue
                if (
                    s1 == 1
                    and s2 == 1
                    and vm1 is None
                    and vm2 is None
                    and not _cell_has_user_content(c2, namespace)
                ):
                    fp = table_merge_structure_fingerprint(
                        source,
                        table=table_idx,
                        start_row=r_idx,
                        end_row=r_idx,
                        start_col=c1_s,
                        end_col=c2_e,
                        part=part,
                    )
                    if fp is not None:
                        target = TableMergeTarget(
                            table=table_idx,
                            start_row=r_idx,
                            end_row=r_idx,
                            start_col=c1_s,
                            end_col=c2_e,
                            structure_fingerprint_sha256=fp,
                            action="merge",
                            part=part,
                            origin_text=_get_cell_text(c1, namespace),
                            is_merged=False,
                        )
                        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-merge-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


_MERGE_OWNED_TCPR = frozenset({"tcW", "gridSpan", "vMerge"})
_FAILED_VERIFICATION = "table_merge_staged_verification_failed"


def _conflict_report(
    request: TableMergeRequest,
    reason: str,
    row: int,
    cell: int | None,
    **details: Any,
) -> dict[str, Any]:
    """Content-safe refusal report naming the exact one-based row and cell in conflict."""
    return {
        "table": request.table,
        "action": request.action,
        "part": request.part,
        "start_row": request.start_row,
        "end_row": request.end_row,
        "start_col": request.start_col,
        "end_col": request.end_col,
        "reason": reason,
        "row": row,
        "cell": cell,
        **details,
    }


def _first_unsupported_merge_cell(
    request: TableMergeRequest, rows: list[ET.Element], namespace: str
) -> dict[str, Any] | None:
    """Locate the first cell with malformed or legacy hMerge markup, requested rows first."""
    requested = list(range(request.start_row, request.end_row + 1))
    others = [index for index in range(1, len(rows) + 1) if index not in requested]
    for r_idx in (*requested, *others):
        for ordinal, cell in enumerate(rows[r_idx - 1].findall(Q(namespace, "tc")), start=1):
            reason = _cell_malformed_reason(cell, namespace)
            if reason is None:
                continue
            details: dict[str, Any] = {}
            if reason == "hMerge_unsupported":
                try:
                    details["h_merge"] = _cell_hmerge_mode(cell, namespace)
                except ValueError:
                    details["h_merge"] = None
            return _conflict_report(request, reason, r_idx, ordinal, **details)
    return None


def _entry_at(
    entries: list[tuple[int, ET.Element, int, int]], column: int
) -> tuple[int, ET.Element, int, int] | None:
    """The (ordinal, cell, col_start, col_end) entry whose columns cover ``column``."""
    for entry in entries:
        if entry[2] <= column <= entry[3]:
            return entry
    return None


def _signature_outside_table(root: ET.Element, table: ET.Element) -> bytes | None:
    """Serialize the whole part with ``table`` masked out; the tree is restored afterwards."""
    for parent in root.iter():
        for index, child in enumerate(parent):
            if child is table:
                marker = ET.Element("table-merge-target")
                marker.tail = table.tail
                parent[index] = marker
                try:
                    return ET.tostring(root)
                finally:
                    parent[index] = table
    return None


def _cell_preserved_projection(
    cell: ET.Element, namespace: str, owned: frozenset[str]
) -> tuple[list[bytes], list[bytes]]:
    """(tcPr children not in ``owned``, non-tcPr children) of a cell, serialized in order."""
    tc_pr = cell.find(Q(namespace, "tcPr"))
    props = (
        []
        if tc_pr is None
        else [ET.tostring(child) for child in tc_pr if child.tag.rsplit("}", 1)[-1] not in owned]
    )
    return props, [ET.tostring(child) for child in cell if child is not tc_pr]


def author_table_merge(request: TableMergeRequest) -> TableMergeReport:
    """Execute a source-bound table cell merge or split authoring request atomically."""
    if not isinstance(request, TableMergeRequest):
        raise TypeError("table_merge_request_invalid")

    if request.action not in ("merge", "split"):
        return _table_merge_refusal(request, "table_merge_action_invalid")

    if request.start_row > request.end_row or request.start_col > request.end_col:
        return _table_merge_refusal(request, "table_merge_coordinates_invalid")

    source = request.source
    destination = request.destination

    if not source.is_file():
        return _table_merge_refusal(request, "table_merge_source_missing")
    if source == destination:
        return _table_merge_refusal(request, "table_merge_destination_invalid")
    if not destination.parent.is_dir():
        return _table_merge_refusal(request, "table_merge_destination_invalid")
    if destination.exists() or destination.is_symlink():
        return _table_merge_refusal(request, "table_merge_destination_occupied")

    try:
        actual_source_sha256 = _sha256_file(source)
    except OSError:
        return _table_merge_refusal(request, "table_merge_source_missing")

    if actual_source_sha256 != request.source_sha256:
        return _table_merge_refusal(request, "table_merge_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_merge_refusal(request, "table_merge_package_invalid")

    if request.part == _DOCUMENT and _DOCUMENT not in graph.parts:
        return _table_merge_refusal(request, "table_merge_package_invalid")

    valid_part, part_err = _check_reachable_part(graph, request.part, "table_merge")
    if not valid_part:
        return _table_merge_refusal(request, part_err or "table_merge_part_unsupported")

    try:
        part_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(part_bytes)
    except (KeyError, ET.ParseError):
        return _table_merge_refusal(request, "table_merge_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_merge_refusal(request, "table_merge_namespace_unsupported")

    def word(local):
        return Q(namespace, local)
    unsupported_context = _unsupported_context_by_element(root, namespace)
    tables = list(root.iter(word("tbl")))
    if request.table < 1 or request.table > len(tables):
        return _table_merge_refusal(request, "table_merge_target_missing")

    table = tables[request.table - 1]
    if unsupported_context.get(id(table), False):
        return _table_merge_refusal(request, "table_merge_context_unsupported")

    if not _is_supported_table(table, namespace):
        return _table_merge_refusal(request, "table_merge_table_unsupported")

    tbl_grid = table.find(word("tblGrid"))
    grid_cols = tbl_grid.findall(word("gridCol")) if tbl_grid is not None else []
    num_grid_cols = len(grid_cols)
    if num_grid_cols == 0:
        return _table_merge_refusal(request, "table_merge_grid_mismatch")

    rows = table.findall(word("tr"))
    if not rows:
        return _table_merge_refusal(request, "table_merge_table_unsupported")

    if (
        request.start_row < 1
        or request.end_row > len(rows)
        or request.start_col < 1
        or request.end_col > num_grid_cols
    ):
        return _table_merge_refusal(request, "table_merge_target_missing")

    # Any malformed markup or legacy hMerge anywhere in the table refuses the whole table,
    # naming the exact row and cell so nothing is silently canonicalized or dropped.
    unsupported_markup = _first_unsupported_merge_cell(request, rows, namespace)
    if unsupported_markup is not None:
        return _table_merge_refusal(
            request, "table_merge_structure_unsupported", change_report=unsupported_markup
        )

    for r_idx in range(request.start_row, request.end_row + 1):
        row = rows[r_idx - 1]
        if unsupported_context.get(id(row), False):
            return _table_merge_refusal(request, "table_merge_context_unsupported")

    row_mappings: list[list[tuple[ET.Element, int, int]]] = []
    for row in rows:
        mapping = _row_grid_mapping(row, num_grid_cols, namespace)
        if mapping is None:
            return _table_merge_refusal(request, "table_merge_grid_mismatch")
        row_mappings.append(mapping)
    entries = [
        [(ordinal, cell, c_start, c_end) for ordinal, (cell, c_start, c_end) in enumerate(m, 1)]
        for m in row_mappings
    ]

    def cell_refusal(
        code: str, reason: str, r_idx: int, entry: tuple[int, ET.Element, int, int]
    ) -> TableMergeReport:
        ordinal, conflict_cell, c_start, c_end = entry
        try:
            c_span, c_vmerge = _cell_merge_state(conflict_cell, namespace)
        except ValueError:
            c_span, c_vmerge = None, None
        return _table_merge_refusal(
            request,
            code,
            change_report=_conflict_report(
                request,
                reason,
                r_idx,
                ordinal,
                cell_start_col=c_start,
                cell_end_col=c_end,
                cell_grid_span=c_span,
                cell_v_merge=c_vmerge,
            ),
        )

    # Require valid positive explicit affected grid widths (no silent skipped sizing)
    col_widths: list[int] = []
    for c in range(request.start_col, request.end_col + 1):
        w_attr = grid_cols[c - 1].get(word("w"))
        if w_attr is None:
            return _table_merge_refusal(request, "table_merge_grid_mismatch")
        try:
            col_w = int(w_attr)
            if col_w <= 0:
                return _table_merge_refusal(request, "table_merge_grid_mismatch")
        except (ValueError, TypeError):
            return _table_merge_refusal(request, "table_merge_grid_mismatch")
        col_widths.append(col_w)
    total_width = sum(col_widths)

    actual_fp = table_merge_structure_fingerprint(
        source,
        table=request.table,
        start_row=request.start_row,
        end_row=request.end_row,
        start_col=request.start_col,
        end_col=request.end_col,
        part=request.part,
    )
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_merge_refusal(request, "table_merge_structure_fingerprint_mismatch")

    col_span = request.end_col - request.start_col + 1
    row_span = request.end_row - request.start_row + 1

    if col_span == 1 and row_span == 1:
        return _table_merge_refusal(request, "table_merge_geometry_unsupported")

    # Pristine references for staged verification; the live tree is mutated below.
    pristine_table = copy.deepcopy(table)
    outside_before = _signature_outside_table(root, table)
    if outside_before is None:
        return _table_merge_refusal(request, "table_merge_table_unsupported")

    origin_cell: ET.Element | None = None
    origin_text = ""
    continuation_texts: dict[int, str] = {}

    if request.action == "merge":
        # MERGE VALIDATION
        for r_idx in range(request.start_row, request.end_row + 1):
            for entry in entries[r_idx - 1]:
                _, cell, c_start, c_end = entry
                if c_end < request.start_col or c_start > request.end_col:
                    continue
                if c_start < request.start_col or c_end > request.end_col:
                    return cell_refusal(
                        "table_merge_geometry_unsupported",
                        "cell_crosses_target_boundary",
                        r_idx,
                        entry,
                    )
                if cell.find(f".//{word('tbl')}") is not None:
                    return cell_refusal(
                        "table_merge_geometry_unsupported",
                        "cell_contains_nested_table",
                        r_idx,
                        entry,
                    )
                if _has_forbidden_cell_descendant(cell):
                    return cell_refusal(
                        "table_merge_geometry_unsupported",
                        "cell_contains_unsupported_content",
                        r_idx,
                        entry,
                    )
                try:
                    span, v_merge = _cell_merge_state(cell, namespace)
                except ValueError:
                    return cell_refusal(
                        "table_merge_structure_unsupported",
                        "malformed_merge_markup",
                        r_idx,
                        entry,
                    )
                if span > 1 or v_merge is not None:
                    return cell_refusal(
                        "table_merge_geometry_unsupported", "cell_already_merged", r_idx, entry
                    )
                if r_idx == request.start_row and c_start == request.start_col:
                    origin_cell = cell
                    origin_text = _get_cell_text(cell, namespace)
                elif _cell_has_user_content(cell, namespace):
                    return cell_refusal(
                        "table_merge_content_loss_unsupported", "cell_has_content", r_idx, entry
                    )

        if origin_cell is None:
            return _table_merge_refusal(request, "table_merge_geometry_unsupported")

        # A continue cell directly below would silently join the new vertical merge.
        if row_span > 1 and request.end_row < len(rows):
            below = _entry_at(entries[request.end_row], request.start_col)
            if below is not None and below[2] == request.start_col:
                try:
                    _, below_vmerge = _cell_merge_state(below[1], namespace)
                except ValueError:
                    below_vmerge = None
                if below_vmerge == "continue":
                    return cell_refusal(
                        "table_merge_geometry_unsupported",
                        "vmerge_continues_below_target",
                        request.end_row + 1,
                        below,
                    )

        # EXECUTE MERGE
        origin_tc_pr = origin_cell.find(word("tcPr"))
        if origin_tc_pr is None:
            origin_tc_pr = ET.Element(word("tcPr"))
            origin_cell.insert(0, origin_tc_pr)

        _set_tc_width(origin_tc_pr, total_width, namespace)
        if col_span > 1:
            _set_grid_span(origin_tc_pr, col_span, namespace)
        if row_span > 1:
            _set_v_merge(origin_tc_pr, "restart", namespace)

        origin_row = rows[request.start_row - 1]
        origin_row_mapping = row_mappings[request.start_row - 1]
        for cell, c_start, c_end in origin_row_mapping:
            if c_start > request.start_col and c_end <= request.end_col:
                origin_row.remove(cell)

        for r_idx in range(request.start_row + 1, request.end_row + 1):
            row = rows[r_idx - 1]
            mapping = row_mappings[r_idx - 1]
            cont_cell: ET.Element | None = None
            for cell, c_start, c_end in mapping:
                if c_start == request.start_col:
                    cont_cell = cell
                elif c_start > request.start_col and c_end <= request.end_col:
                    row.remove(cell)

            if cont_cell is None:
                return _table_merge_refusal(request, "table_merge_geometry_unsupported")

            cont_tc_pr = cont_cell.find(word("tcPr"))
            if cont_tc_pr is None:
                cont_tc_pr = ET.Element(word("tcPr"))
                cont_cell.insert(0, cont_tc_pr)

            _set_tc_width(cont_tc_pr, total_width, namespace)
            if col_span > 1:
                _set_grid_span(cont_tc_pr, col_span, namespace)
            _set_v_merge(cont_tc_pr, "continue", namespace)

            for p in cont_cell.findall(word("p")):
                for r in p.findall(word("r")):
                    p.remove(r)
            if not cont_cell.findall(word("p")):
                ET.SubElement(cont_cell, word("p"))

    elif request.action == "split":
        # SPLIT VALIDATION
        origin_entry = _entry_at(entries[request.start_row - 1], request.start_col)
        if origin_entry is None:
            return _table_merge_refusal(request, "table_merge_geometry_unsupported")
        if origin_entry[2] != request.start_col or origin_entry[3] != request.end_col:
            return cell_refusal(
                "table_merge_geometry_unsupported",
                "cell_extent_mismatch",
                request.start_row,
                origin_entry,
            )
        origin_cell = origin_entry[1]

        origin_text = _get_cell_text(origin_cell, namespace)
        try:
            span, v_merge = _cell_merge_state(origin_cell, namespace)
        except ValueError:
            return cell_refusal(
                "table_merge_structure_unsupported",
                "malformed_merge_markup",
                request.start_row,
                origin_entry,
            )

        if span != col_span:
            return cell_refusal(
                "table_merge_geometry_unsupported",
                "grid_span_mismatch",
                request.start_row,
                origin_entry,
            )

        if row_span > 1:
            if v_merge != "restart":
                return cell_refusal(
                    "table_merge_geometry_unsupported",
                    "vmerge_restart_missing",
                    request.start_row,
                    origin_entry,
                )
            if request.start_row > 1:
                prev_entry = _entry_at(entries[request.start_row - 2], request.start_col)
                if prev_entry is not None and prev_entry[2] == request.start_col:
                    try:
                        _, prev_vm = _cell_merge_state(prev_entry[1], namespace)
                    except ValueError:
                        return cell_refusal(
                            "table_merge_structure_unsupported",
                            "malformed_merge_markup",
                            request.start_row - 1,
                            prev_entry,
                        )
                    if prev_vm == "continue":
                        return cell_refusal(
                            "table_merge_geometry_unsupported",
                            "vmerge_continues_above_target",
                            request.start_row - 1,
                            prev_entry,
                        )
        elif v_merge is not None:
            return cell_refusal(
                "table_merge_geometry_unsupported",
                "vmerge_present",
                request.start_row,
                origin_entry,
            )

        # Validate every affected cell across all covered rows
        for r_idx in range(request.start_row, request.end_row + 1):
            affected = _entry_at(entries[r_idx - 1], request.start_col)
            if affected is None:
                return _table_merge_refusal(request, "table_merge_geometry_unsupported")
            if affected[2] != request.start_col or affected[3] != request.end_col:
                return cell_refusal(
                    "table_merge_geometry_unsupported", "cell_extent_mismatch", r_idx, affected
                )
            affected_cell = affected[1]

            if unsupported_context.get(id(affected_cell), False):
                return cell_refusal(
                    "table_merge_context_unsupported",
                    "cell_in_unsupported_context",
                    r_idx,
                    affected,
                )
            if affected_cell.find(f".//{word('tbl')}") is not None:
                return cell_refusal(
                    "table_merge_geometry_unsupported",
                    "cell_contains_nested_table",
                    r_idx,
                    affected,
                )
            if _has_forbidden_cell_descendant(affected_cell):
                return cell_refusal(
                    "table_merge_geometry_unsupported",
                    "cell_contains_unsupported_content",
                    r_idx,
                    affected,
                )

            if r_idx > request.start_row:
                try:
                    c_span, c_vmerge = _cell_merge_state(affected_cell, namespace)
                except ValueError:
                    return cell_refusal(
                        "table_merge_structure_unsupported",
                        "malformed_merge_markup",
                        r_idx,
                        affected,
                    )
                if c_vmerge != "continue" or c_span != span:
                    return cell_refusal(
                        "table_merge_geometry_unsupported",
                        "vmerge_continuation_missing",
                        r_idx,
                        affected,
                    )
                continuation_texts[r_idx] = _get_cell_text(affected_cell, namespace)

        # Check vertical merge does not continue past end_row
        if row_span > 1 and request.end_row < len(rows):
            below = _entry_at(entries[request.end_row], request.start_col)
            if below is not None and below[2] == request.start_col:
                try:
                    _, after_vmerge = _cell_merge_state(below[1], namespace)
                except ValueError:
                    return cell_refusal(
                        "table_merge_structure_unsupported",
                        "malformed_merge_markup",
                        request.end_row + 1,
                        below,
                    )
                if after_vmerge == "continue":
                    return cell_refusal(
                        "table_merge_geometry_unsupported",
                        "vmerge_continues_below_target",
                        request.end_row + 1,
                        below,
                    )

        # EXECUTE SPLIT
        for r_idx in range(request.start_row, request.end_row + 1):
            row = rows[r_idx - 1]
            mapping = row_mappings[r_idx - 1]
            cell_to_split = None
            for cell, c_start, _ in mapping:
                if c_start == request.start_col:
                    cell_to_split = cell
                    break
            assert cell_to_split is not None

            tc_pr = cell_to_split.find(word("tcPr"))
            if tc_pr is not None:
                _remove_tcpr_child(tc_pr, "gridSpan", namespace)
                _remove_tcpr_child(tc_pr, "vMerge", namespace)
                col_w0 = grid_cols[request.start_col - 1].get(word("w"))
                _set_tc_width(tc_pr, col_w0, namespace)

            if not cell_to_split.findall(word("p")):
                ET.SubElement(cell_to_split, word("p"))

            if col_span > 1:
                current_anchor = cell_to_split
                for col in range(request.start_col + 1, request.end_col + 1):
                    col_w = grid_cols[col - 1].get(word("w"))
                    new_cell = _create_split_cell(tc_pr, col_w, namespace)
                    anchor_idx = list(row).index(current_anchor)
                    row.insert(anchor_idx + 1, new_cell)
                    current_anchor = new_cell
    else:
        return _table_merge_refusal(request, "table_merge_action_invalid")

    target_change_report = {
        "table": request.table,
        "part": request.part,
        "action": request.action,
        "start_row": request.start_row,
        "end_row": request.end_row,
        "start_col": request.start_col,
        "end_col": request.end_col,
        "rows_count": len(rows),
        "cols_count": num_grid_cols,
        "origin_text": origin_text,
    }

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({request.part})):
            return _FAILED_VERIFICATION
        insp = inspect_docx(staged)
        if not insp.valid:
            return _FAILED_VERIFICATION
        try:
            s_graph = PackageGraph.from_file(staged)
            staged_part_valid, _ = _check_reachable_part(s_graph, request.part, "table_merge")
            if not staged_part_valid:
                return _FAILED_VERIFICATION
            s_root = ET.fromstring(s_graph.read_part_bytes(request.part))
            if _word_namespace(s_root) != namespace:
                return _FAILED_VERIFICATION

            s_tables = list(s_root.iter(word("tbl")))
            if len(s_tables) != len(tables):
                return _FAILED_VERIFICATION
            s_table = s_tables[request.table - 1]

            # Everything in the selected part outside the target table is unchanged
            outside_after = _signature_outside_table(s_root, s_table)
            if outside_after is None or outside_after != outside_before:
                return _FAILED_VERIFICATION

            # Verify tblGrid is byte-identical
            orig_grid = pristine_table.find(word("tblGrid"))
            staged_grid = s_table.find(word("tblGrid"))
            if (orig_grid is None) != (staged_grid is None):
                return _FAILED_VERIFICATION
            if (
                orig_grid is not None
                and staged_grid is not None
                and ET.tostring(orig_grid) != ET.tostring(staged_grid)
            ):
                return _FAILED_VERIFICATION

            orig_rows = pristine_table.findall(word("tr"))
            s_rows = s_table.findall(word("tr"))
            if len(s_rows) != len(orig_rows):
                return _FAILED_VERIFICATION

            # Verify preservation of unselected rows
            for r_idx in range(1, len(orig_rows) + 1):
                if (r_idx < request.start_row or r_idx > request.end_row) and ET.tostring(
                    s_rows[r_idx - 1]
                ) != ET.tostring(orig_rows[r_idx - 1]):
                    return _FAILED_VERIFICATION

            # Verify every row in [start_row, end_row]
            for r_idx in range(request.start_row, request.end_row + 1):
                s_mapping = _row_grid_mapping(s_rows[r_idx - 1], num_grid_cols, namespace)
                orig_mapping = _row_grid_mapping(orig_rows[r_idx - 1], num_grid_cols, namespace)
                if s_mapping is None or orig_mapping is None:
                    return _FAILED_VERIFICATION

                # Verify unselected cells in affected row
                for s_cell, sc_s, sc_e in s_mapping:
                    if sc_e < request.start_col or sc_s > request.end_col:
                        matched = False
                        for o_cell, oc_s, oc_e in orig_mapping:
                            if oc_s == sc_s and oc_e == sc_e:
                                if ET.tostring(s_cell) == ET.tostring(o_cell):
                                    matched = True
                                break
                        if not matched:
                            return _FAILED_VERIFICATION

                orig_at_start = next(
                    (o_cell for o_cell, oc_s, _ in orig_mapping if oc_s == request.start_col),
                    None,
                )
                if orig_at_start is None:
                    return _FAILED_VERIFICATION
                orig_props, orig_content = _cell_preserved_projection(
                    orig_at_start, namespace, _MERGE_OWNED_TCPR
                )

                # Verify affected cells in affected row
                if request.action == "merge":
                    covered = [
                        (c, cs, ce)
                        for c, cs, ce in s_mapping
                        if not (ce < request.start_col or cs > request.end_col)
                    ]
                    if len(covered) != 1:
                        return _FAILED_VERIFICATION
                    m_cell, mc_s, mc_e = covered[0]
                    if mc_s != request.start_col or mc_e != request.end_col:
                        return _FAILED_VERIFICATION
                    span, vm = _cell_merge_state(m_cell, namespace)
                    if col_span > 1 and span != col_span:
                        return _FAILED_VERIFICATION
                    if col_span == 1 and span != 1:
                        return _FAILED_VERIFICATION
                    m_props, m_content = _cell_preserved_projection(
                        m_cell, namespace, _MERGE_OWNED_TCPR
                    )
                    if m_props != orig_props:
                        return _FAILED_VERIFICATION
                    if r_idx == request.start_row:
                        if row_span > 1 and vm != "restart":
                            return _FAILED_VERIFICATION
                        if row_span == 1 and vm is not None:
                            return _FAILED_VERIFICATION
                        if _get_cell_text(m_cell, namespace) != origin_text:
                            return _FAILED_VERIFICATION
                        if m_content != orig_content:
                            return _FAILED_VERIFICATION
                    else:
                        if vm != "continue":
                            return _FAILED_VERIFICATION
                        if _get_cell_text(m_cell, namespace) != "":
                            return _FAILED_VERIFICATION
                else:
                    # Split action
                    split_cells = [
                        (c, cs, ce)
                        for c, cs, ce in s_mapping
                        if cs >= request.start_col and ce <= request.end_col
                    ]
                    if len(split_cells) != col_span:
                        return _FAILED_VERIFICATION
                    for idx, (s_c, sc_s, sc_e) in enumerate(split_cells):
                        exp_col = request.start_col + idx
                        if sc_s != exp_col or sc_e != exp_col:
                            return _FAILED_VERIFICATION
                        span, vm = _cell_merge_state(s_c, namespace)
                        if span != 1 or vm is not None:
                            return _FAILED_VERIFICATION
                        s_tc_pr = s_c.find(word("tcPr"))
                        if s_tc_pr is None:
                            return _FAILED_VERIFICATION
                        w_elem = s_tc_pr.find(word("tcW"))
                        if w_elem is None or w_elem.get(word("w")) != grid_cols[exp_col - 1].get(
                            word("w")
                        ):
                            return _FAILED_VERIFICATION
                        if exp_col > request.start_col:
                            if _get_cell_text(s_c, namespace) != "":
                                return _FAILED_VERIFICATION
                            continue
                        s_props, s_content = _cell_preserved_projection(
                            s_c, namespace, _MERGE_OWNED_TCPR
                        )
                        if s_props != orig_props:
                            return _FAILED_VERIFICATION
                        if orig_at_start.findall(word("p")) and s_content != orig_content:
                            return _FAILED_VERIFICATION
                        if r_idx == request.start_row:
                            if _get_cell_text(s_c, namespace) != origin_text:
                                return _FAILED_VERIFICATION
                        elif _get_cell_text(s_c, namespace) != continuation_texts.get(r_idx, ""):
                            return _FAILED_VERIFICATION

        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return _FAILED_VERIFICATION
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
        verify=verify_publication,
        failure_code=_FAILED_VERIFICATION,
        exists_code="table_merge_destination_occupied",
        unavailable_code="table_merge_publication_unavailable",
    )

    if not published.passed:
        return _table_merge_refusal(request, published.code or _FAILED_VERIFICATION)

    return TableMergeReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
