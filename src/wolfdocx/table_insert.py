"""Source-bound DOCX table insertion and target discovery.

Provides verified, source-bound insertion of a new rectangular table before or
after a paragraph that sits directly in a block container of a WordprocessingML
part: the body (or a header/footer root), a table cell (``w:tc``, at any nesting
depth), or a text-box story (``w:txbxContent``). Existing document content,
anchor paragraphs, section formatting, and unrelated document parts are
strictly preserved. Paragraphs inside any other wrapper (content controls,
hyperlinks, revisions, fields, comment ranges, Markup Compatibility choices,
footnotes) are refused.
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
    _require_sha256,
    _sha256_file,
    _unsupported_context_by_element,
)
from .inspect import (
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import (
    _DOCUMENT,
    ReviewReport,
    _mask_text,
    _publish,
    _word_namespace,
)
from .table_properties import (
    _HEX_COLOR_PATTERN,
    _canonical_table_alignment,
    _check_reachable_part,
    _validate_referenced_table_style,
)

__all__ = [
    "TableInsertReport",
    "TableInsertRequest",
    "TableInsertTarget",
    "author_table_insert",
    "discover_table_insert_targets",
    "load_table_insert_request",
    "parse_table_insert_request",
    "table_insert_structure_fingerprint",
]

_DISCOVERED = "discovered"


_REQUIRED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "rows",
        "grid_widths",
    }
)

_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "rows",
        "grid_widths",
        "position",
        "style",
        "repeat_header",
        "alignment",
        "shading",
    }
)

_REQUIRED_TARGET_FIELDS = frozenset({"paragraph"})
_ALLOWED_TARGET_FIELDS = frozenset({"paragraph", "part"})
_TABLE_ALIGNMENTS = frozenset({"left", "center", "right"})


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
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


def _paragraph_unsupported_reason(paragraph: ET.Element, ns: str) -> str | None:
    """Inspect paragraph for unsupported child elements or anchors."""
    for tag_name in ("drawing", "pict", "wgp"):
        if paragraph.findall(f".//{Q(ns, tag_name)}"):
            return "table_insert_drawing_unsupported"

    if (
        paragraph.findall(f".//{Q(ns, 'fldSimple')}")
        or paragraph.findall(f".//{Q(ns, 'instrText')}")
        or paragraph.findall(f".//{Q(ns, 'fldChar')}")
    ):
        return "table_insert_field_unsupported"

    if (
        paragraph.findall(f".//{Q(ns, 'ins')}")
        or paragraph.findall(f".//{Q(ns, 'del')}")
        or paragraph.findall(f".//{Q(ns, 'moveFrom')}")
        or paragraph.findall(f".//{Q(ns, 'moveTo')}")
        or paragraph.findall(f".//{Q(ns, 'rPrChange')}")
        or paragraph.findall(f".//{Q(ns, 'pPrChange')}")
    ):
        return "table_insert_revision_unsupported"

    if (
        paragraph.findall(f".//{Q(ns, 'commentRangeStart')}")
        or paragraph.findall(f".//{Q(ns, 'commentRangeEnd')}")
        or paragraph.findall(f".//{Q(ns, 'commentReference')}")
    ):
        return "table_insert_comment_unsupported"

    if paragraph.findall(f".//{Q(ns, 'sdt')}"):
        return "table_insert_control_unsupported"

    if paragraph.findall(f".//{Q(ns, 'hyperlink')}"):
        return "table_insert_hyperlink_unsupported"

    return None


def _visible_text(paragraph: ET.Element, ns: str) -> str:
    return "".join(t.text or "" for t in paragraph.iter(Q(ns, "t")))


_MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_ALTERNATE_CONTENT = f"{{{_MC_NAMESPACE}}}AlternateContent"


def _story_root(root: ET.Element, ns: str, part: str) -> ET.Element | None:
    """The element whose direct paragraphs are the part's top-level block anchors."""
    return root.find(Q(ns, "body")) if part == _DOCUMENT else root


def _cell_hides_content(cell: ET.Element, ns: str) -> bool:
    """True for a cell whose content is tracked, or hidden by a merge continuation."""
    cell_pr = cell.find(Q(ns, "tcPr"))
    if cell_pr is None:
        return False
    if any(cell_pr.find(Q(ns, name)) is not None for name in ("cellIns", "cellDel", "cellMerge")):
        return True
    for merge in ("vMerge", "hMerge"):
        marker = cell_pr.find(Q(ns, merge))
        if marker is not None and marker.get(Q(ns, "val"), "continue") != "restart":
            return True
    return False


def _anchor_context_unsupported(
    target: ET.Element,
    parents: Mapping[ET.Element, ET.Element],
    ns: str,
    context: Mapping[int, bool],
) -> bool:
    """Ancestor/range wrappers, alternate-content duplicates, or hidden cell content."""
    if context.get(id(target), True):
        return True
    node = parents.get(target)
    if node is not None and node.tag == Q(ns, "tc") and _cell_hides_content(node, ns):
        return True
    while node is not None:
        if node.tag == _MC_ALTERNATE_CONTENT:
            return True
        node = parents.get(node)
    return False


def _anchor_admission(
    target: ET.Element,
    parents: Mapping[ET.Element, ET.Element],
    story_root: ET.Element,
    ns: str,
    context: Mapping[int, bool],
) -> ET.Element | str:
    """The block container an anchor may be inserted beside, else the refusal code.

    The one admission rule shared by discovery and authoring: the paragraph must be
    a direct child of the story root, a table cell, or a text-box story, and must
    carry no unsupported wrapper context or paragraph content.
    """
    container = parents.get(target)
    if container is None or not (
        container is story_root or container.tag in (Q(ns, "tc"), Q(ns, "txbxContent"))
    ):
        return "table_insert_anchor_nested"
    if _anchor_context_unsupported(target, parents, ns, context):
        return "table_insert_context_unsupported"
    if _paragraph_unsupported_reason(target, ns) is not None:
        return "table_insert_anchor_unsupported"
    return container


def _child_path(
    element: ET.Element, parents: Mapping[ET.Element, ET.Element], root: ET.Element
) -> list[int]:
    """Child indices leading from ``root`` to ``element``."""
    path: list[int] = []
    while element is not root:
        parent = parents[element]
        path.append(list(parent).index(element))
        element = parent
    path.reverse()
    return path


def _build_table_element(
    rows_data: Sequence[Sequence[str]],
    grid_widths: Sequence[int],
    style: str | None,
    repeat_header: bool,
    ns: str,
    alignment: str | None = None,
    shading: str | None = None,
) -> ET.Element:
    """Build a schema-ordered w:tbl element with tblPr, tblGrid, tr, tc, and p."""
    tbl = ET.Element(Q(ns, "tbl"))

    # 1. tblPr (CT_TblPr order: tblStyle, ..., tblW, jc, ..., shd, ...)
    tbl_pr = ET.SubElement(tbl, Q(ns, "tblPr"))
    if style is not None:
        tbl_style = ET.SubElement(tbl_pr, Q(ns, "tblStyle"))
        tbl_style.set(Q(ns, "val"), style)

    total_w = sum(grid_widths)
    tbl_w = ET.SubElement(tbl_pr, Q(ns, "tblW"))
    tbl_w.set(Q(ns, "w"), str(total_w))
    tbl_w.set(Q(ns, "type"), "dxa")

    if alignment is not None:
        jc = ET.SubElement(tbl_pr, Q(ns, "jc"))
        jc.set(Q(ns, "val"), alignment)

    if shading is not None:
        shd = ET.SubElement(tbl_pr, Q(ns, "shd"))
        shd.set(Q(ns, "val"), "clear")
        shd.set(Q(ns, "color"), "auto")
        shd.set(Q(ns, "fill"), shading)

    # 2. tblGrid
    tbl_grid = ET.SubElement(tbl, Q(ns, "tblGrid"))
    for w in grid_widths:
        grid_col = ET.SubElement(tbl_grid, Q(ns, "gridCol"))
        grid_col.set(Q(ns, "w"), str(w))

    # 3. rows (tr)
    cols = len(grid_widths)
    for r_idx, row in enumerate(rows_data):
        tr = ET.SubElement(tbl, Q(ns, "tr"))
        if repeat_header and r_idx == 0:
            tr_pr = ET.SubElement(tr, Q(ns, "trPr"))
            ET.SubElement(tr_pr, Q(ns, "tblHeader"))

        for c_idx in range(cols):
            cell_text = row[c_idx] if c_idx < len(row) else ""
            tc = ET.SubElement(tr, Q(ns, "tc"))
            tc_pr = ET.SubElement(tc, Q(ns, "tcPr"))
            tc_w = ET.SubElement(tc_pr, Q(ns, "tcW"))
            tc_w.set(Q(ns, "w"), str(grid_widths[c_idx]))
            tc_w.set(Q(ns, "type"), "dxa")

            p = ET.SubElement(tc, Q(ns, "p"))
            if cell_text:
                r = ET.SubElement(p, Q(ns, "r"))
                t = ET.SubElement(r, Q(ns, "t"))
                t.text = cell_text
                if cell_text.startswith(" ") or cell_text.endswith(" ") or "  " in cell_text:
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return tbl


@dataclass(frozen=True)
class TableInsertTarget:
    """One candidate paragraph anchor (body, table cell, or text-box story) for table insertion."""

    part: str
    paragraph: int
    text: str
    structure_fingerprint_sha256: str
    empty: bool = False
    supported: bool = True

    @property
    def target(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "paragraph": self.paragraph,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "paragraph": self.paragraph,
            "text": self.text,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "empty": self.empty,
            "supported": self.supported,
            "target": self.target,
        }


@dataclass(frozen=True)
class TableInsertRequest:
    """One source-bound request to insert a rectangular table before or after a paragraph anchor."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    rows: tuple[tuple[str, ...], ...]
    grid_widths: tuple[int, ...]
    position: str = "after"
    style: str | None = None
    repeat_header: bool = False
    alignment: str | None = None
    shading: str | None = None

    def __init__(
        self,
        *,
        source: str | Path,
        destination: str | Path,
        source_sha256: str,
        target: Mapping[str, Any],
        structure_fingerprint_sha256: str,
        rows: Sequence[Sequence[str]],
        grid_widths: Sequence[int],
        position: str = "after",
        style: str | None = None,
        repeat_header: bool = False,
        alignment: str | None = None,
        shading: str | None = None,
    ) -> None:
        source_path = Path(source) if isinstance(source, str) else source
        if not isinstance(source_path, Path) or not str(source_path):
            raise ValueError("table_insert_source_invalid")

        dest_path = Path(destination) if isinstance(destination, str) else destination
        if not isinstance(dest_path, Path) or not str(dest_path):
            raise ValueError("table_insert_destination_invalid")

        valid_source_sha = _require_sha256(source_sha256, "table_insert_source_sha256")
        valid_fp = _require_sha256(
            structure_fingerprint_sha256, "table_insert_structure_fingerprint_sha256"
        )

        if not isinstance(target, Mapping):
            raise TypeError("table_insert_target_invalid")
        target_keys = set(target)
        if not target_keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(
            target_keys
        ):
            raise ValueError("table_insert_target_fields_invalid")

        target_p = target["paragraph"]
        target_part = target.get("part", _DOCUMENT)

        if isinstance(target_p, bool) or not isinstance(target_p, int) or target_p < 1:
            raise ValueError("table_insert_paragraph_invalid")
        if not isinstance(target_part, str) or not target_part:
            raise ValueError("table_insert_part_invalid")

        if not isinstance(position, str) or position not in ("before", "after"):
            raise ValueError("table_insert_position_invalid")

        if not isinstance(rows, Sequence) or isinstance(rows, (str, bytes)):
            raise TypeError("table_insert_rows_invalid")
        if len(rows) == 0:
            raise ValueError("table_insert_rows_empty")

        parsed_rows: list[tuple[str, ...]] = []
        cols: int | None = None
        for row in rows:
            if not isinstance(row, Sequence) or isinstance(row, (str, bytes)):
                raise TypeError("table_insert_row_invalid")
            if cols is None:
                cols = len(row)
                if cols == 0:
                    raise ValueError("table_insert_rows_empty")
            elif len(row) != cols:
                raise ValueError("table_insert_rows_non_rectangular")

            row_cells: list[str] = []
            for cell in row:
                if not isinstance(cell, str):
                    raise TypeError("table_insert_cell_text_invalid")
                if len(cell) > _MAX_TEXT_CHARACTERS:
                    raise ValueError("table_insert_cell_text_limit_exceeded")
                row_cells.append(cell)
            parsed_rows.append(tuple(row_cells))

        if not isinstance(grid_widths, Sequence) or isinstance(grid_widths, (str, bytes)):
            raise TypeError("table_insert_grid_widths_invalid")
        if len(grid_widths) == 0:
            raise ValueError("table_insert_grid_widths_invalid")
        if len(grid_widths) != cols:
            raise ValueError("table_insert_grid_widths_count_mismatch")

        gw_list: list[int] = []
        for w in grid_widths:
            if isinstance(w, bool) or not isinstance(w, int) or w <= 0:
                raise ValueError("table_insert_grid_widths_invalid")
            gw_list.append(w)

        if style is not None and (not isinstance(style, str) or not style):
            raise ValueError("table_insert_style_invalid")

        if not isinstance(repeat_header, bool):
            raise TypeError("table_insert_repeat_header_invalid")

        if alignment is not None:
            canonical_alignment = _canonical_table_alignment(alignment, _TABLE_ALIGNMENTS)
            if canonical_alignment is None:
                raise ValueError("table_insert_alignment_invalid")
            alignment = canonical_alignment

        if shading is not None and (
            not isinstance(shading, str) or _HEX_COLOR_PATTERN.match(shading) is None
        ):
            raise ValueError("table_insert_shading_invalid")

        object.__setattr__(self, "source", source_path)
        object.__setattr__(self, "destination", dest_path)
        object.__setattr__(self, "source_sha256", valid_source_sha)
        object.__setattr__(
            self,
            "target",
            {"part": target_part, "paragraph": target_p},
        )
        object.__setattr__(self, "structure_fingerprint_sha256", valid_fp)
        object.__setattr__(self, "rows", tuple(parsed_rows))
        object.__setattr__(self, "grid_widths", tuple(gw_list))
        object.__setattr__(self, "position", position)
        object.__setattr__(self, "style", style)
        object.__setattr__(self, "repeat_header", repeat_header)
        object.__setattr__(self, "alignment", alignment)
        object.__setattr__(self, "shading", shading)

    @property
    def paragraph(self) -> int:
        return self.target["paragraph"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "paragraph": self.paragraph,
            "position": self.position,
        }

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
            "position": self.position,
            "rows": [list(r) for r in self.rows],
            "grid_widths": list(self.grid_widths),
        }
        if self.style is not None:
            data["style"] = self.style
        if self.repeat_header:
            data["repeat_header"] = self.repeat_header
        if self.alignment is not None:
            data["alignment"] = self.alignment
        if self.shading is not None:
            data["shading"] = self.shading
        return data


@dataclass(frozen=True)
class TableInsertReport:
    """Content-safe public result for a table insert authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    inserted_table: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "table-insert",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "inserted_table": self.inserted_table,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _table_insert_refusal(
    request: TableInsertRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TableInsertReport:
    return TableInsertReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        inserted_table=None,
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_insert_request(value: Any) -> TableInsertRequest:
    """Parse and strictly validate a table insert authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_insert_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_insert_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, (str, Path)) or not str(source):
        raise ValueError("table_insert_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, (str, Path)) or not str(destination):
        raise ValueError("table_insert_destination_invalid")

    source_sha256 = value["source_sha256"]
    if not isinstance(source_sha256, str):
        raise TypeError("table_insert_source_sha256_invalid")

    fp = value["structure_fingerprint_sha256"]
    if not isinstance(fp, str):
        raise TypeError("table_insert_structure_fingerprint_sha256_invalid")

    target = value["target"]
    if not isinstance(target, Mapping):
        raise TypeError("table_insert_target_invalid")
    target_keys = set(target)
    if not target_keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(
        target_keys
    ):
        raise ValueError("table_insert_target_fields_invalid")

    paragraph = target["paragraph"]
    part = target.get("part", _DOCUMENT)

    rows = value["rows"]
    grid_widths = value["grid_widths"]
    position = value.get("position", "after")
    style = value.get("style")
    repeat_header = value.get("repeat_header", False)
    alignment = value.get("alignment")
    shading = value.get("shading")

    return TableInsertRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        target={"paragraph": paragraph, "part": part},
        structure_fingerprint_sha256=fp,
        rows=rows,
        grid_widths=grid_widths,
        position=position,
        style=style,
        repeat_header=repeat_header,
        alignment=alignment,
        shading=shading,
    )


def load_table_insert_request(path: str | Path) -> TableInsertRequest:
    """Load and strictly validate one JSON table-insert request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_insert_request(json.load(stream))


def table_insert_structure_fingerprint(
    source: str | Path,
    *,
    paragraph: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one exact paragraph.

    The ordinal counts every ``w:p`` in the part in document order, including cell
    and text-box paragraphs. A paragraph directly inside an unsupported container
    (for example a footnote) keeps a stable source fingerprint even though
    :func:`author_table_insert` refuses it as ``table_insert_anchor_nested``, so
    that refusal has an exact, source-bound request. Wrapper context and paragraph
    content that authoring refuses outright yield no fingerprint.
    """
    path = Path(source)
    if not path.is_file() or paragraph < 1:
        return None
    try:
        graph = PackageGraph.from_file(path)
        valid, _ = _check_reachable_part(graph, part, "table_insert")
        if not valid:
            return None
        root = ET.fromstring(graph.read_part_bytes(part))
        ns = _word_namespace(root)
        if ns is None:
            return None
        paragraphs = list(root.iter(Q(ns, "p")))
        if paragraph < 1 or paragraph > len(paragraphs):
            return None
        target_p = paragraphs[paragraph - 1]
        parents = {c: p for p in root.iter() for c in p}
        context = _unsupported_context_by_element(root, ns, stop_at=target_p)
        if _anchor_context_unsupported(target_p, parents, ns, context):
            return None
        if _paragraph_unsupported_reason(target_p, ns) is not None:
            return None
        return _mask_text(target_p, ns)
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


def discover_table_insert_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate paragraph anchors for table insertion."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-insert-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_insert_source_missing",
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
            "operation": "table-insert-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_insert_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    valid_part, part_err = _check_reachable_part(graph, part, "table_insert")
    if not valid_part:
        return {
            "schema_version": 1,
            "operation": "table-insert-targets",
            "passed": False,
            "status": _REFUSED,
            "code": part_err or "table_insert_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        doc_bytes = graph.read_part_bytes(part)
        root = ET.fromstring(doc_bytes)
    except (PackageError, ET.ParseError, KeyError, OSError):
        return {
            "schema_version": 1,
            "operation": "table-insert-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_insert_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    ns = _word_namespace(root)
    if ns is None:
        return {
            "schema_version": 1,
            "operation": "table-insert-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_insert_namespace_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    container = _story_root(root, ns, part)
    if container is None:
        return {
            "schema_version": 1,
            "operation": "table-insert-targets",
            "passed": True,
            "status": _DISCOVERED,
            "code": None,
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    parent_map = {c: p for p in root.iter() for c in p}
    unsupported_context = _unsupported_context_by_element(root, ns)

    targets: list[dict[str, Any]] = []
    for paragraph_index, paragraph in enumerate(root.iter(Q(ns, "p")), start=1):
        if isinstance(
            _anchor_admission(paragraph, parent_map, container, ns, unsupported_context), str
        ):
            continue

        text = _visible_text(paragraph, ns)
        fp = _mask_text(paragraph, ns)
        target = TableInsertTarget(
            part=part,
            paragraph=paragraph_index,
            text=text,
            structure_fingerprint_sha256=fp,
            empty=len(text.strip()) == 0,
            supported=True,
        )
        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-insert-targets",
        "passed": True,
        "status": _DISCOVERED,
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_table_insert(request: TableInsertRequest) -> TableInsertReport:
    """Insert a new rectangular table before or after a bound paragraph atomically."""
    if not isinstance(request, TableInsertRequest):
        raise TypeError("table_insert_request_invalid")

    source = request.source
    destination = request.destination

    if not source.is_file():
        return _table_insert_refusal(request, "table_insert_source_missing")

    if destination.exists():
        return _table_insert_refusal(request, "table_insert_destination_occupied")

    try:
        source_sha256 = _sha256_file(source)
    except OSError:
        return _table_insert_refusal(request, "table_insert_source_missing")

    if source_sha256 != request.source_sha256:
        return _table_insert_refusal(request, "table_insert_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_insert_refusal(request, "table_insert_package_invalid")

    valid_part, part_err = _check_reachable_part(graph, request.part, "table_insert")
    if not valid_part:
        return _table_insert_refusal(request, part_err or "table_insert_part_unsupported")

    try:
        doc_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (PackageError, ET.ParseError, KeyError, OSError):
        return _table_insert_refusal(request, "table_insert_package_invalid")

    ns = _word_namespace(root)
    if ns is None:
        return _table_insert_refusal(request, "table_insert_namespace_unsupported")

    paragraphs = list(root.iter(Q(ns, "p")))
    if request.paragraph < 1 or request.paragraph > len(paragraphs):
        return _table_insert_refusal(request, "table_insert_target_missing")

    target_p = paragraphs[request.paragraph - 1]

    story_root = _story_root(root, ns, request.part)
    if story_root is None:
        return _table_insert_refusal(request, "table_insert_anchor_unsupported")

    parent_map = {c: p for p in root.iter() for c in p}
    unsupported_context = _unsupported_context_by_element(root, ns, stop_at=target_p)
    admitted = _anchor_admission(target_p, parent_map, story_root, ns, unsupported_context)
    if isinstance(admitted, str):
        return _table_insert_refusal(request, admitted)
    container = admitted

    fp = _mask_text(target_p, ns)
    if fp != request.structure_fingerprint_sha256:
        return _table_insert_refusal(request, "table_insert_structure_fingerprint_mismatch")

    if not request.rows or not request.rows[0]:
        return _table_insert_refusal(request, "table_insert_rows_empty")

    cols = len(request.rows[0])
    if any(len(r) != cols for r in request.rows):
        return _table_insert_refusal(request, "table_insert_rows_non_rectangular")

    for r in request.rows:
        for cell_text in r:
            if len(cell_text) > _MAX_TEXT_CHARACTERS:
                return _table_insert_refusal(request, "table_insert_cell_text_limit_exceeded")

    if not request.grid_widths:
        return _table_insert_refusal(request, "table_insert_grid_widths_missing")
    if len(request.grid_widths) != cols:
        return _table_insert_refusal(request, "table_insert_grid_widths_count_mismatch")
    if any(w <= 0 for w in request.grid_widths):
        return _table_insert_refusal(request, "table_insert_grid_widths_invalid")

    final_widths = request.grid_widths

    if request.style is not None:
        style_err = _validate_referenced_table_style(graph, request.style)
        if style_err is not None:
            return _table_insert_refusal(request, style_err)

    original_container_children = list(container)
    anchor_idx = original_container_children.index(target_p)
    # Structural location of the selected container, independent of paragraph ordinals
    # (the inserted table's own cell paragraphs shift every later ordinal).
    container_path = _child_path(container, parent_map, root)
    # A cell or text box must still end with a paragraph after the new table.
    needs_terminal = (
        request.position == "after"
        and container.tag in (Q(ns, "tc"), Q(ns, "txbxContent"))
        and anchor_idx == len(original_container_children) - 1
    )

    new_tbl = _build_table_element(
        request.rows,
        final_widths,
        request.style,
        request.repeat_header,
        ns,
        alignment=request.alignment,
        shading=request.shading,
    )

    expected_tbl_idx = anchor_idx if request.position == "before" else anchor_idx + 1
    container.insert(expected_tbl_idx, new_tbl)
    if needs_terminal:
        ET.SubElement(container, Q(ns, "p"))

    inserted_table_index = list(root.iter(Q(ns, "tbl"))).index(new_tbl) + 1

    target_change_report = {
        "table_index": inserted_table_index,
        "rows": len(request.rows),
        "cols": cols,
        "grid_widths": list(final_widths),
        "style": request.style,
        "repeat_header": request.repeat_header,
        "alignment": request.alignment,
        "shading": request.shading,
        "position": request.position,
        "anchor_paragraph": request.paragraph,
        "part": request.part,
    }

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({request.part})):
            return "table_insert_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_insert_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(request.part))
            s_ns = _word_namespace(s_root)
            if s_ns != ns:
                return "table_insert_staged_verification_failed"
            if _story_root(s_root, s_ns, request.part) is None:
                return "table_insert_staged_verification_failed"

            # 1. Follow the selected container's structural path, not a paragraph ordinal.
            s_container = s_root
            for index in container_path:
                s_siblings = list(s_container)
                if index >= len(s_siblings):
                    return "table_insert_staged_verification_failed"
                s_container = s_siblings[index]
            if s_container.tag != container.tag:
                return "table_insert_staged_verification_failed"

            s_children = list(s_container)
            # 2. Exactly the table, plus the terminal paragraph when one was required.
            if len(s_children) != len(original_container_children) + 1 + int(needs_terminal):
                return "table_insert_staged_verification_failed"

            # 3. Element at expected_tbl_idx must be the inserted table.
            s_tbl = s_children[expected_tbl_idx]
            if s_tbl.tag != Q(s_ns, "tbl"):
                return "table_insert_staged_verification_failed"

            # 4. Exact anchor adjacency.
            anchor_pos = (
                expected_tbl_idx - 1 if request.position == "after" else expected_tbl_idx + 1
            )
            anchor_node = s_children[anchor_pos]
            if anchor_node.tag != Q(s_ns, "p"):
                return "table_insert_staged_verification_failed"
            if _mask_text(anchor_node, s_ns) != request.structure_fingerprint_sha256:
                return "table_insert_staged_verification_failed"

            # 5. The required terminal paragraph is a bare empty paragraph.
            s_terminal = s_children[-1] if needs_terminal else None
            if s_terminal is not None and (
                s_terminal is s_tbl
                or s_terminal.tag != Q(s_ns, "p")
                or len(s_terminal) > 0
                or s_terminal.attrib
                or (s_terminal.text or "").strip()
                or (s_terminal.tail or "").strip()
            ):
                return "table_insert_staged_verification_failed"

            # 6. Detach exactly the additions: everything else in the part must then be
            # identical to the source part, not just the container's children.
            s_tbl_snapshot = copy.deepcopy(s_tbl)
            s_container.remove(s_tbl)
            if s_terminal is not None:
                s_container.remove(s_terminal)
            if ET.tostring(s_root) != ET.tostring(ET.fromstring(doc_bytes)):
                return "table_insert_staged_verification_failed"
            s_tbl = s_tbl_snapshot

            # 7. Verify inserted table shape, widths, and cell content
            s_rows = s_tbl.findall(Q(s_ns, "tr"))
            if len(s_rows) != len(request.rows):
                return "table_insert_staged_verification_failed"

            for s_row, expected_row in zip(s_rows, request.rows):
                s_cells = s_row.findall(Q(s_ns, "tc"))
                if len(s_cells) != cols:
                    return "table_insert_staged_verification_failed"
                for c_idx, (s_cell, expected_text) in enumerate(zip(s_cells, expected_row)):
                    cell_p = s_cell.find(Q(s_ns, "p"))
                    if cell_p is None:
                        return "table_insert_staged_verification_failed"
                    actual_text = "".join(t.text or "" for t in cell_p.iter(Q(s_ns, "t")))
                    if actual_text != expected_text:
                        return "table_insert_staged_verification_failed"
                    tc_pr = s_cell.find(Q(s_ns, "tcPr"))
                    if tc_pr is not None:
                        tc_w = tc_pr.find(Q(s_ns, "tcW"))
                        if tc_w is not None and tc_w.get(Q(s_ns, "w")) != str(final_widths[c_idx]):
                            return "table_insert_staged_verification_failed"

            s_grid = s_tbl.find(Q(s_ns, "tblGrid"))
            if s_grid is None:
                return "table_insert_staged_verification_failed"
            s_cols = s_grid.findall(Q(s_ns, "gridCol"))
            if len(s_cols) != cols:
                return "table_insert_staged_verification_failed"
            for s_col, expected_w in zip(s_cols, final_widths):
                if s_col.get(Q(s_ns, "w")) != str(expected_w):
                    return "table_insert_staged_verification_failed"

            s_tbl_pr = s_tbl.find(Q(s_ns, "tblPr"))
            if s_tbl_pr is None:
                return "table_insert_staged_verification_failed"
            expected_tbl_pr_tags = [
                Q(s_ns, local)
                for local, present in (
                    ("tblStyle", request.style is not None),
                    ("tblW", True),
                    ("jc", request.alignment is not None),
                    ("shd", request.shading is not None),
                )
                if present
            ]
            if [child.tag for child in s_tbl_pr] != expected_tbl_pr_tags:
                return "table_insert_staged_verification_failed"

            if request.style is not None:
                s_tbl_style = s_tbl_pr.find(Q(s_ns, "tblStyle"))
                if s_tbl_style is None or s_tbl_style.get(Q(s_ns, "val")) != request.style:
                    return "table_insert_staged_verification_failed"

            if request.alignment is not None:
                s_jc = s_tbl_pr.find(Q(s_ns, "jc"))
                if s_jc is None or s_jc.get(Q(s_ns, "val")) != request.alignment:
                    return "table_insert_staged_verification_failed"

            if request.shading is not None:
                s_shd = s_tbl_pr.find(Q(s_ns, "shd"))
                if (
                    s_shd is None
                    or s_shd.get(Q(s_ns, "fill")) != request.shading
                    or s_shd.get(Q(s_ns, "val")) != "clear"
                ):
                    return "table_insert_staged_verification_failed"

            if request.repeat_header:
                s_first_row = s_rows[0]
                s_tr_pr = s_first_row.find(Q(s_ns, "trPr"))
                if s_tr_pr is None or s_tr_pr.find(Q(s_ns, "tblHeader")) is None:
                    return "table_insert_staged_verification_failed"

        except (
            OSError,
            PackageError,
            ET.ParseError,
            KeyError,
            IndexError,
            ValueError,
            TypeError,
        ):
            return "table_insert_staged_verification_failed"
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
        failure_code="table_insert_staged_verification_failed",
        exists_code="table_insert_destination_occupied",
        unavailable_code="table_insert_publication_unavailable",
    )

    if not published.passed:
        return _table_insert_refusal(
            request, published.code or "table_insert_staged_verification_failed"
        )

    return TableInsertReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        inserted_table=target_change_report,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
