"""Tests for source-bound table cell properties authoring and discovery."""

from __future__ import annotations

import hashlib
import json
from functools import partial
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import cell, grid, paragraph, row, table, write_package
from wolfdocx.inspect import Q, inspect_docx
from wolfdocx.package import PackageGraph
from wolfdocx.review import _DOCUMENT, _word_namespace
from wolfdocx.table_cell_properties import (
    BorderSpec,
    TableBordersSpec,
    TableCellMarginsSpec,
    TableCellPropertiesReport,
    TableCellPropertiesRequest,
    TableCellPropertiesSpec,
    TableCellPropertiesTarget,
    TableShadingSpec,
    TableWidthSpec,
    author_table_cell_properties,
    discover_table_cell_properties_targets,
    load_table_cell_properties_request,
    parse_table_cell_properties_request,
    table_cell_properties_structure_fingerprint,
)

_DUMMY_SHA = "a" * 64
_DUMMY_FP = "b" * 64
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_cell_doc(
    path: Path,
    *,
    rows_data: list[list[str]] | None = None,
    grid_widths: list[int] | None = None,
    cell_properties: dict[tuple[int, int], str] | None = None,
    tbl_properties: str = "",
    wrapper: str = "",
    styles_extra: str = "",
    row_properties: str = "",
    cell_xml_override: dict[tuple[int, int], str] | None = None,
) -> Path:
    """Create a minimal valid DOCX with an explicit table and optional cell properties."""
    data = rows_data if rows_data is not None else [["A", "B"], ["C", "D"]]
    cols = len(data[0])
    widths = grid_widths if grid_widths is not None else [2500] * cols
    table_grid = grid(widths)

    c_props = cell_properties or {}
    c_overrides = cell_xml_override or {}

    rows_xml: list[str] = []
    for r_idx, r_data in enumerate(data):
        cells_xml: list[str] = []
        for c_idx, txt in enumerate(r_data):
            if (r_idx, c_idx) in c_overrides:
                cells_xml.append(c_overrides[(r_idx, c_idx)])
            else:
                props = c_props.get((r_idx, c_idx), "")
                cells_xml.append(cell(paragraph(txt), properties=props))
        r_pr = row_properties if r_idx == 0 else ""
        rows_xml.append(row("".join(cells_xml), properties=r_pr))

    tbl_xml = table("".join(rows_xml), properties=tbl_properties, grid=table_grid)
    if wrapper == "sdt":
        body = f"<w:sdt><w:sdtContent>{tbl_xml}</w:sdtContent></w:sdt>"
    elif wrapper == "ins":
        body = f'<w:ins w:id="1" w:author="tester">{tbl_xml}</w:ins>'
    else:
        body = tbl_xml

    parts: dict[str, str | bytes] | None = None
    if styles_extra:
        parts = {
            "word/styles.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:styles xmlns:w="{_W_NS}">{styles_extra}</w:styles>'
            )
        }

    return write_package(path, body, parts=parts)


# --- 1. Combined Property Authoring ---


def test_author_table_cell_properties_combined_all_five(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties=TableCellPropertiesSpec(
            width=TableWidthSpec(value=2400, type="dxa"),
            vertical_alignment="center",
            shading=TableShadingSpec(fill="F0F0F0", val="clear", color="auto"),
            borders=TableBordersSpec(
                top=BorderSpec(val="single", sz=4, space=0, color="auto"),
                bottom=BorderSpec(val="single", sz=4, space=0, color="auto"),
            ),
            cell_margins=TableCellMarginsSpec(top=100, left=150, bottom=100, right=150),
        ),
    )

    report = author_table_cell_properties(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == source_sha
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True
    assert report.target == {"table": 1, "row": 1, "cell": 1, "part": _DOCUMENT}

    # Verify staged package validity
    insp = inspect_docx(dest)
    assert insp.valid is True

    # Verify cell properties in XML
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    w_tc = root.findall(f".//{{{_W_NS}}}tc")[0]
    tc_pr = w_tc.find(f"{{{_W_NS}}}tcPr")
    assert tc_pr is not None

    # Check child ordering: tcW < tcBorders < shd < tcMar < vAlign
    children = [child.tag.rsplit("}", 1)[-1] for child in tc_pr]
    expected_order = ["tcW", "tcBorders", "shd", "tcMar", "vAlign"]
    assert children == expected_order

    # Check values
    tc_w = tc_pr.find(f"{{{_W_NS}}}tcW")
    assert tc_w is not None
    assert tc_w.get(f"{{{_W_NS}}}w") == "2400"
    assert tc_w.get(f"{{{_W_NS}}}type") == "dxa"

    v_align = tc_pr.find(f"{{{_W_NS}}}vAlign")
    assert v_align is not None
    assert v_align.get(f"{{{_W_NS}}}val") == "center"

    shd = tc_pr.find(f"{{{_W_NS}}}shd")
    assert shd is not None
    assert shd.get(f"{{{_W_NS}}}fill") == "F0F0F0"

    tc_borders = tc_pr.find(f"{{{_W_NS}}}tcBorders")
    assert tc_borders is not None
    assert tc_borders.find(f"{{{_W_NS}}}top") is not None
    assert tc_borders.find(f"{{{_W_NS}}}bottom") is not None

    tc_mar = tc_pr.find(f"{{{_W_NS}}}tcMar")
    assert tc_mar is not None
    assert tc_mar.find(f"{{{_W_NS}}}top").get(f"{{{_W_NS}}}w") == "100"
    assert tc_mar.find(f"{{{_W_NS}}}left").get(f"{{{_W_NS}}}w") == "150"

    # Cell text remains unchanged
    t_nodes = w_tc.findall(f".//{{{_W_NS}}}t")
    assert len(t_nodes) == 1
    assert t_nodes[0].text == "A"

    # Sibling cell text unchanged
    all_t = [t.text for t in root.findall(f".//{{{_W_NS}}}t")]
    assert all_t == ["A", "B", "C", "D"]

    # Target change report
    change = report.target_change_report
    assert change is not None
    assert change["table"] == 1
    assert change["row"] == 1
    assert change["cell"] == 1
    assert sorted(change["properties_authored"]) == [
        "borders",
        "cell_margins",
        "shading",
        "vertical_alignment",
        "width",
    ]


# --- 2. Partial-Property Preservation ---


def test_author_table_cell_properties_partial_preserves_unspecified(
    tmp_path: Path,
) -> None:
    # Cell initially has width and vertical alignment
    initial_tcpr = '<w:tcW w:w="3200" w:type="dxa"/><w:vAlign w:val="bottom"/>'
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): initial_tcpr},
    )
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    # Apply only shading
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"shading": {"fill": "00FF00", "val": "clear", "color": "auto"}},
    )

    report = author_table_cell_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    w_tc = root.findall(f".//{{{_W_NS}}}tc")[0]
    tc_pr = w_tc.find(f"{{{_W_NS}}}tcPr")
    assert tc_pr is not None

    # Unspecified properties (width and vAlign) preserved
    assert tc_pr.find(f"{{{_W_NS}}}tcW").get(f"{{{_W_NS}}}w") == "3200"
    assert tc_pr.find(f"{{{_W_NS}}}vAlign").get(f"{{{_W_NS}}}val") == "bottom"
    assert tc_pr.find(f"{{{_W_NS}}}shd").get(f"{{{_W_NS}}}fill") == "00FF00"

    # Schema order: tcW < shd < vAlign
    children = [child.tag.rsplit("}", 1)[-1] for child in tc_pr]
    assert children == ["tcW", "shd", "vAlign"]


def test_author_table_cell_properties_borders_partial_preservation(
    tmp_path: Path,
) -> None:
    # Cell initially has bottom border
    initial_tcpr = '<w:tcBorders><w:bottom w:val="single" w:sz="6" w:color="000000"/></w:tcBorders>'
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): initial_tcpr},
    )
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)

    # Request updates only top border
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"borders": {"top": {"val": "double", "sz": 12, "color": "FF0000"}}},
    )

    report = author_table_cell_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    w_tc = root.findall(f".//{{{_W_NS}}}tc")[0]
    tc_borders = w_tc.find(f".//{{{_W_NS}}}tcBorders")
    assert tc_borders is not None

    # Both top and bottom are present
    top_el = tc_borders.find(f"{{{_W_NS}}}top")
    bottom_el = tc_borders.find(f"{{{_W_NS}}}bottom")
    assert top_el is not None and top_el.get(f"{{{_W_NS}}}val") == "double"
    assert bottom_el is not None and bottom_el.get(f"{{{_W_NS}}}val") == "single"

    # Order inside tcBorders is top then bottom
    side_tags = [c.tag.rsplit("}", 1)[-1] for c in tc_borders]
    assert side_tags == ["top", "bottom"]


def test_author_table_cell_properties_margins_partial_preservation(
    tmp_path: Path,
) -> None:
    # Cell initially has top margin
    initial_tcpr = '<w:tcMar><w:top w:w="120" w:type="dxa"/></w:tcMar>'
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): initial_tcpr},
    )
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)

    # Request updates left margin
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"cell_margins": {"left": 200}},
    )

    report = author_table_cell_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    w_tc = root.findall(f".//{{{_W_NS}}}tc")[0]
    tc_mar = w_tc.find(f".//{{{_W_NS}}}tcMar")
    assert tc_mar is not None

    # Both top and left margins are present
    assert tc_mar.find(f"{{{_W_NS}}}top").get(f"{{{_W_NS}}}w") == "120"
    assert tc_mar.find(f"{{{_W_NS}}}left").get(f"{{{_W_NS}}}w") == "200"

    # Order inside tcMar is top then left
    mar_tags = [c.tag.rsplit("}", 1)[-1] for c in tc_mar]
    assert mar_tags == ["top", "left"]


# --- 3. Set and Clear Semantics ---


def test_author_table_cell_properties_set_clear_semantics(tmp_path: Path) -> None:
    # Start with a cell that has thick borders, colored shading, and custom width
    initial_tcpr = (
        '<w:tcW w:w="4000" w:type="dxa"/>'
        '<w:tcBorders><w:top w:val="thick" w:sz="24"/></w:tcBorders>'
        '<w:shd w:fill="FF0000" w:val="clear"/>'
        '<w:vAlign w:val="center"/>'
    )
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): initial_tcpr},
    )
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)

    # Clear border to 'none', clear shading to 'auto', reset width to auto, reset vAlign to top
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties=TableCellPropertiesSpec(
            width=TableWidthSpec(value=0, type="auto"),
            vertical_alignment="top",
            shading=TableShadingSpec(fill="auto", val="clear", color="auto"),
            borders=TableBordersSpec(top=BorderSpec(val="none")),
        ),
    )

    report = author_table_cell_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    w_tc = root.findall(f".//{{{_W_NS}}}tc")[0]
    tc_pr = w_tc.find(f"{{{_W_NS}}}tcPr")
    assert tc_pr is not None

    assert tc_pr.find(f"{{{_W_NS}}}tcW").get(f"{{{_W_NS}}}type") == "auto"
    assert tc_pr.find(f"{{{_W_NS}}}vAlign").get(f"{{{_W_NS}}}val") == "top"
    assert tc_pr.find(f"{{{_W_NS}}}shd").get(f"{{{_W_NS}}}fill") == "auto"
    assert tc_pr.find(f".//{{{_W_NS}}}top").get(f"{{{_W_NS}}}val") == "none"


# --- 4. Child Ordering and Schema Integrity ---


def test_author_table_cell_properties_child_ordering_independent_of_input_order(
    tmp_path: Path,
) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)

    # Request passed in arbitrary order
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "vertical_alignment": "center",
            "cell_margins": {"right": 100},
            "shading": {"fill": "123456"},
            "borders": {"right": {"val": "single"}},
            "width": {"value": 2000, "type": "dxa"},
        },
    )

    report = author_table_cell_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    w_tc = root.findall(f".//{{{_W_NS}}}tc")[0]
    tc_pr = w_tc.find(f"{{{_W_NS}}}tcPr")
    children = [child.tag.rsplit("}", 1)[-1] for child in tc_pr]
    assert children == ["tcW", "tcBorders", "shd", "tcMar", "vAlign"]


# --- 5. Namespace Stability ---


def test_author_table_cell_properties_namespace_stability(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"shading": {"fill": "ABCDEF"}},
    )

    report = author_table_cell_properties(req)
    assert report.passed is True

    dest_bytes = PackageGraph.from_file(dest).read_part_bytes(_DOCUMENT)
    raw_xml = dest_bytes.decode("utf-8")
    assert 'xmlns=""' not in raw_xml

    root = ET.fromstring(dest_bytes)
    ns = _word_namespace(root)
    assert ns == _W_NS
    tbl = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl")
    assert tbl is not None
    tc = tbl.find(f".//{{{_W_NS}}}tc")
    assert tc is not None
    tc_pr = tc.find(f"{{{_W_NS}}}tcPr")
    assert tc_pr is not None
    shd = tc_pr.find(f"{{{_W_NS}}}shd")
    assert shd is not None
    assert shd.get(f"{{{_W_NS}}}fill") == "ABCDEF"


# --- 6. Discovery Tests ---


def test_discover_table_cell_properties_targets(tmp_path: Path) -> None:
    # Create doc with pre-existing property on (1, 1) and plain (1, 2)
    c_props = {(0, 0): '<w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="center"/>'}
    source = _make_cell_doc(tmp_path / "source.docx", cell_properties=c_props)

    disc = discover_table_cell_properties_targets(source)
    assert disc["passed"] is True
    assert disc["status"] == "discovered"
    assert disc["operation"] == "table-cell-properties-targets"
    assert disc["source_sha256"] == _sha256(source)
    assert disc["part"] == _DOCUMENT
    targets = disc["targets"]
    assert len(targets) == 4  # 2x2 table

    # First cell has properties
    t1 = targets[0]
    assert t1["table"] == 1
    assert t1["row"] == 1
    assert t1["cell"] == 1
    assert t1["part"] == _DOCUMENT
    assert t1["current_properties"]["width"] == {"value": 2000, "type": "dxa"}
    assert t1["current_properties"]["vertical_alignment"] == "center"

    # Second cell has no direct properties
    t2 = targets[1]
    assert t2["table"] == 1
    assert t2["row"] == 1
    assert t2["cell"] == 2
    assert t2["current_properties"] == {}


def test_discover_table_cell_properties_skips_unsupported(tmp_path: Path) -> None:
    # Cell (0, 0) is a vMerge continuation
    c_props = {(0, 0): "<w:vMerge/>"}
    source = _make_cell_doc(tmp_path / "source.docx", cell_properties=c_props)

    disc = discover_table_cell_properties_targets(source)
    assert disc["passed"] is True
    # Merged cell is excluded from targets and recorded in unsupported
    targets = disc["targets"]
    assert len(targets) == 3
    assert not any(t["row"] == 1 and t["cell"] == 1 for t in targets)

    unsupported = disc["unsupported"]
    assert len(unsupported) == 1
    assert unsupported[0]["table"] == 1
    assert unsupported[0]["row"] == 1
    assert unsupported[0]["cell"] == 1
    assert unsupported[0]["code"] == "table_cell_properties_cell_merged"


def test_discover_table_cell_properties_source_missing(tmp_path: Path) -> None:
    disc = discover_table_cell_properties_targets(tmp_path / "nonexistent.docx")
    assert disc["passed"] is False
    assert disc["status"] == "refused"
    assert disc["code"] == "table_cell_properties_source_missing"


def test_discover_table_cell_properties_unsupported_part(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    disc = discover_table_cell_properties_targets(source, part="word/header1.xml")
    assert disc["passed"] is False
    assert disc["code"] == "table_cell_properties_part_unsupported"


# --- 7. Structural Fingerprint Tests ---


def test_table_cell_properties_structure_fingerprint(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None
    assert len(fp) == 64

    # Text changes do not change fingerprint
    source_diff_text = _make_cell_doc(
        tmp_path / "source_diff_text.docx", rows_data=[["Different", "B"], ["C", "D"]]
    )
    fp2 = table_cell_properties_structure_fingerprint(source_diff_text, table=1, row=1, cell=1)
    assert fp2 == fp

    # Structural changes (e.g. adding properties) DO change fingerprint
    source_diff_struct = _make_cell_doc(
        tmp_path / "source_diff_struct.docx",
        cell_properties={(0, 0): '<w:shd w:fill="FF0000"/>'},
    )
    fp3 = table_cell_properties_structure_fingerprint(source_diff_struct, table=1, row=1, cell=1)
    assert fp3 != fp

    # Out of bounds returns None
    assert table_cell_properties_structure_fingerprint(source, table=99, row=1, cell=1) is None
    assert table_cell_properties_structure_fingerprint(source, table=1, row=99, cell=1) is None
    assert table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=99) is None
    assert (
        table_cell_properties_structure_fingerprint(
            tmp_path / "absent.docx", table=1, row=1, cell=1
        )
        is None
    )


# --- 8. Dataclass and JSON Parsing / Loading Tests ---


def test_parse_and_load_table_cell_properties_request(tmp_path: Path) -> None:
    req_dict = {
        "source": str(tmp_path / "doc.docx"),
        "destination": str(tmp_path / "out.docx"),
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1, "row": 2, "cell": 3, "part": _DOCUMENT},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {
            "width": {"value": 1500, "type": "dxa"},
            "vertical_alignment": "center",
        },
    }

    req = parse_table_cell_properties_request(req_dict)
    assert req.table == 1
    assert req.row == 2
    assert req.cell == 3
    assert req.part == _DOCUMENT
    assert req.properties["width"] == {"value": 1500, "type": "dxa"}
    assert req.properties["vertical_alignment"] == "center"

    # Idempotent parse
    assert parse_table_cell_properties_request(req) is req

    # Spec object passed in
    spec_req = parse_table_cell_properties_request(
        {
            "source": str(tmp_path / "doc.docx"),
            "destination": str(tmp_path / "out.docx"),
            "source_sha256": _DUMMY_SHA,
            "target": {"table": 1, "row": 1, "cell": 1},
            "structure_fingerprint_sha256": _DUMMY_FP,
            "properties": TableCellPropertiesSpec(
                width=TableWidthSpec(value=2000),
                vertical_alignment="top",
            ),
        }
    )
    assert spec_req.properties["width"]["value"] == 2000

    # Save to disk and load
    json_path = tmp_path / "req.json"
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(req_dict, f)

    loaded = load_table_cell_properties_request(json_path)
    assert loaded.table == 1
    assert loaded.row == 2
    assert loaded.cell == 3


def test_table_cell_properties_dataclass_targets_and_reports() -> None:
    target_obj = TableCellPropertiesTarget(
        table=1,
        row=1,
        cell=1,
        structure_fingerprint_sha256=_DUMMY_FP,
    )
    assert target_obj.table == 1
    assert target_obj.target == {"table": 1, "row": 1, "cell": 1, "part": _DOCUMENT}
    assert target_obj.to_dict()["table"] == 1

    report_obj = TableCellPropertiesReport(
        status="published",
        code=None,
        source_sha256=_DUMMY_SHA,
    )
    assert report_obj.passed is True
    assert report_obj.to_dict()["operation"] == "table-cell-properties"


def test_parse_table_cell_properties_validation_errors(tmp_path: Path) -> None:
    base = {
        "source": str(tmp_path / "doc.docx"),
        "destination": str(tmp_path / "out.docx"),
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1, "row": 1, "cell": 1},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {"vertical_alignment": "center"},
    }

    # Not a mapping
    with pytest.raises(TypeError, match="table_cell_properties_request_invalid"):
        parse_table_cell_properties_request("string")

    # Missing fields
    for field in (
        "source",
        "source_sha256",
        "destination",
        "target",
        "structure_fingerprint_sha256",
        "properties",
    ):
        bad = dict(base)
        del bad[field]
        with pytest.raises(ValueError, match="table_cell_properties_request_missing_fields"):
            parse_table_cell_properties_request(bad)

    # Extra fields
    with pytest.raises(ValueError, match="table_cell_properties_request_extra_fields"):
        parse_table_cell_properties_request({**base, "extra": "unallowed"})

    # Invalid sha
    with pytest.raises(ValueError, match="table_cell_properties_source_sha256"):
        parse_table_cell_properties_request({**base, "source_sha256": "not_hex"})

    # Invalid target
    with pytest.raises(TypeError, match="table_cell_properties_target_invalid"):
        parse_table_cell_properties_request({**base, "target": "not_dict"})

    with pytest.raises(ValueError, match="table_cell_properties_target_missing_fields"):
        parse_table_cell_properties_request({**base, "target": {"table": 1, "row": 1}})

    with pytest.raises(ValueError, match="table_cell_properties_target_table_invalid"):
        parse_table_cell_properties_request({**base, "target": {"table": 0, "row": 1, "cell": 1}})

    with pytest.raises(ValueError, match="table_cell_properties_target_row_invalid"):
        parse_table_cell_properties_request({**base, "target": {"table": 1, "row": -1, "cell": 1}})

    with pytest.raises(ValueError, match="table_cell_properties_target_cell_invalid"):
        parse_table_cell_properties_request(
            {**base, "target": {"table": 1, "row": 1, "cell": "bad"}}
        )

    # Empty properties
    with pytest.raises(ValueError, match="table_cell_properties_properties_empty"):
        parse_table_cell_properties_request({**base, "properties": {}})

    # Extra property keys
    with pytest.raises(ValueError, match="table_cell_properties_properties_extra_fields"):
        parse_table_cell_properties_request({**base, "properties": {"unknown_prop": 123}})

    # Invalid width
    with pytest.raises(ValueError, match="table_cell_properties_width_missing_value"):
        parse_table_cell_properties_request({**base, "properties": {"width": {"type": "dxa"}}})

    with pytest.raises(ValueError, match="table_cell_properties_width_value_invalid"):
        parse_table_cell_properties_request({**base, "properties": {"width": {"value": -5}}})

    with pytest.raises(ValueError, match="table_cell_properties_width_type_invalid"):
        parse_table_cell_properties_request(
            {**base, "properties": {"width": {"value": 100, "type": "bad"}}}
        )

    # Invalid vertical alignment
    with pytest.raises(ValueError, match="table_cell_properties_vertical_alignment_invalid"):
        parse_table_cell_properties_request(
            {**base, "properties": {"vertical_alignment": "sideways"}}
        )

    # Invalid shading
    with pytest.raises(ValueError, match="table_cell_properties_shading_missing_fill"):
        parse_table_cell_properties_request({**base, "properties": {"shading": {"val": "clear"}}})

    with pytest.raises(ValueError, match="table_cell_properties_shading_fill_invalid"):
        parse_table_cell_properties_request({**base, "properties": {"shading": {"fill": "xyz"}}})

    # Invalid borders
    with pytest.raises(ValueError, match="table_cell_properties_borders_empty"):
        parse_table_cell_properties_request({**base, "properties": {"borders": {}}})

    with pytest.raises(ValueError, match="table_cell_properties_border_missing_val"):
        parse_table_cell_properties_request({**base, "properties": {"borders": {"top": {"sz": 4}}}})

    with pytest.raises(ValueError, match="table_cell_properties_border_val_invalid"):
        parse_table_cell_properties_request(
            {**base, "properties": {"borders": {"top": {"val": "unsupported_style"}}}}
        )

    with pytest.raises(ValueError, match="table_cell_properties_border_sz_invalid"):
        parse_table_cell_properties_request(
            {**base, "properties": {"borders": {"top": {"val": "single", "sz": -1}}}}
        )

    # Invalid margins
    with pytest.raises(ValueError, match="table_cell_properties_cell_margins_empty"):
        parse_table_cell_properties_request({**base, "properties": {"cell_margins": {}}})

    with pytest.raises(ValueError, match="table_cell_properties_cell_margin_top_invalid"):
        parse_table_cell_properties_request({**base, "properties": {"cell_margins": {"top": -10}}})


# --- 9. Refusals and Safety Gates ---


def test_refusals_source_missing(tmp_path: Path) -> None:
    req = TableCellPropertiesRequest(
        source=tmp_path / "nonexistent.docx",
        destination=tmp_path / "out.docx",
        source_sha256=_DUMMY_SHA,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_source_missing"
    assert not (tmp_path / "out.docx").exists()


def test_refusals_destination_occupied(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    dest.write_text("occupied")
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_destination_occupied"


def test_refusals_source_destination_same(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    req = TableCellPropertiesRequest(
        source=source,
        destination=source,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_source_destination_same"


def test_refusals_destination_parent_missing(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "nonexistent_dir" / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_destination_parent_missing"


def test_refusals_source_sha_mismatch(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_DUMMY_SHA,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_source_sha256_mismatch"
    assert not dest.exists()


def test_refusals_structure_fingerprint_mismatch(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_structure_fingerprint_mismatch"
    assert not dest.exists()


def test_refusals_target_out_of_bounds(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    source_sha = _sha256(source)

    for target, expected_code in [
        ({"table": 99, "row": 1, "cell": 1}, "table_cell_properties_table_missing"),
        ({"table": 1, "row": 99, "cell": 1}, "table_cell_properties_row_missing"),
        ({"table": 1, "row": 1, "cell": 99}, "table_cell_properties_cell_missing"),
    ]:
        req = TableCellPropertiesRequest(
            source=source,
            destination=dest,
            source_sha256=source_sha,
            target=target,
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"shading": {"fill": "112233"}},
        )
        report = author_table_cell_properties(req)
        assert report.passed is False
        assert report.code == expected_code
        assert not dest.exists()


def test_refusals_unsupported_part(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1, "part": "word/header1.xml"},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_part_unsupported"
    assert not dest.exists()


# --- 10. Refusals for Unsupported Contexts, Merges, and Ambiguity ---


def test_refusals_merged_cell_vmerge(tmp_path: Path) -> None:
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): "<w:vMerge/>"},
    )
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_cell_merged"
    assert not dest.exists()


def test_refusals_merged_cell_gridspan(tmp_path: Path) -> None:
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): '<w:gridSpan w:val="2"/>'},
    )
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_cell_merged"
    assert not dest.exists()


def test_refusals_nested_table(tmp_path: Path) -> None:
    # Cell with nested table inside
    nested_tbl = "<w:tbl><w:tblGrid><w:gridCol w:w='1000'/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Nest</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    cell_override = f"<w:tc><w:p><w:r><w:t>Parent</w:t></w:r></w:p>{nested_tbl}</w:tc>"
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_xml_override={(0, 0): cell_override},
    )
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_structure_fingerprint_mismatch"
    assert not dest.exists()


def test_refusals_table_in_sdt_wrapper(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx", wrapper="sdt")
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_context_unsupported"
    assert not dest.exists()


def test_refusals_ambiguous_duplicate_tcpr(tmp_path: Path) -> None:
    # Cell with two tcPr elements
    cell_override = "<w:tc><w:tcPr/><w:tcPr/><w:p><w:r><w:t>Dup</w:t></w:r></w:p></w:tc>"
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_xml_override={(0, 0): cell_override},
    )
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_cell_ambiguous"
    assert not dest.exists()


def test_refusals_ambiguous_duplicate_child_in_tcpr(tmp_path: Path) -> None:
    # Cell with duplicate tcW inside tcPr
    initial_tcpr = '<w:tcW w:w="1000" w:type="dxa"/><w:tcW w:w="2000" w:type="dxa"/>'
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): initial_tcpr},
    )
    dest = tmp_path / "out.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"shading": {"fill": "112233"}},
    )
    report = author_table_cell_properties(req)
    assert report.passed is False
    assert report.code == "table_cell_properties_cell_ambiguous"
    assert not dest.exists()


# --- 11. Native Render Support Flag Matching ---


def test_vertical_alignment_values_publish(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)

    # 1. Authoring with vAlign="center"
    dest1 = tmp_path / "out_center.docx"
    req1 = TableCellPropertiesRequest(
        source=source,
        destination=dest1,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"vertical_alignment": "center"},
    )
    report1 = author_table_cell_properties(req1)
    assert report1.passed is True

    # 2. Authoring with vAlign="top" and standard border
    dest2 = tmp_path / "out_top.docx"
    req2 = TableCellPropertiesRequest(
        source=source,
        destination=dest2,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "vertical_alignment": "top",
            "borders": {"top": {"val": "single", "sz": 4}},
        },
    )
    report2 = author_table_cell_properties(req2)
    assert report2.passed is True

    # 3. Authoring with vAlign="both" is valid OOXML authoring
    dest3 = tmp_path / "out_both.docx"
    req3 = TableCellPropertiesRequest(
        source=source,
        destination=dest3,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"vertical_alignment": "both"},
    )
    report3 = author_table_cell_properties(req3)
    assert report3.passed is True


# --- 12. No Wrap Authoring and Discovery ---


def test_author_table_cell_properties_nowrap_true(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out_nowrap_true.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"no_wrap": True},
    )
    report = author_table_cell_properties(req)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["properties_authored"] == ["no_wrap"]
    assert report.target_change_report["properties_after"]["no_wrap"] is True

    # Verify saved XML has normative on (<w:noWrap/> without val)
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    word = partial(Q, _W_NS)
    cell_el = next(iter(root.iter(word("tbl")))).findall(word("tr"))[0].findall(word("tc"))[0]
    tc_pr = cell_el.find(word("tcPr"))
    assert tc_pr is not None
    nowrap_el = tc_pr.find(word("noWrap"))
    assert nowrap_el is not None
    # Normative on: val attribute is omitted
    assert nowrap_el.get(word("val")) is None
    assert nowrap_el.get("val") is None


def test_author_table_cell_properties_nowrap_false(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "out_nowrap_false.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"no_wrap": False},
    )
    report = author_table_cell_properties(req)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["properties_after"]["no_wrap"] is False

    # Verify saved XML has explicit w:val="off"
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    word = partial(Q, _W_NS)
    cell_el = next(iter(root.iter(word("tbl")))).findall(word("tr"))[0].findall(word("tc"))[0]
    tc_pr = cell_el.find(word("tcPr"))
    assert tc_pr is not None
    nowrap_el = tc_pr.find(word("noWrap"))
    assert nowrap_el is not None
    assert nowrap_el.get(word("val")) == "off"


def test_author_table_cell_properties_nowrap_omitted_and_none_preserve(
    tmp_path: Path,
) -> None:
    # Source cell with existing <w:noWrap/>
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={(0, 0): "<w:noWrap/>"},
    )
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    # 1. Authoring only shading (no_wrap omitted) preserves existing noWrap
    dest1 = tmp_path / "out_preserve1.docx"
    req1 = TableCellPropertiesRequest(
        source=source,
        destination=dest1,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"shading": {"fill": "AABBCC"}},
    )
    report1 = author_table_cell_properties(req1)
    assert report1.passed is True
    assert report1.target_change_report["properties_before"]["no_wrap"] is True
    assert report1.target_change_report["properties_after"]["no_wrap"] is True

    graph1 = PackageGraph.from_file(dest1)
    root1 = ET.fromstring(graph1.read_part_bytes(_DOCUMENT))
    word = partial(Q, _W_NS)
    tc_pr1 = (
        next(iter(root1.iter(word("tbl"))))
        .findall(word("tr"))[0]
        .findall(word("tc"))[0]
        .find(word("tcPr"))
    )
    assert tc_pr1.find(word("noWrap")) is not None
    assert tc_pr1.find(word("shd")) is not None

    # 2. Spec with no_wrap=None preserves existing noWrap (does NOT delete it)
    dest2 = tmp_path / "out_preserve2.docx"
    dest1_sha = _sha256(dest1)
    fp1 = table_cell_properties_structure_fingerprint(dest1, table=1, row=1, cell=1)
    req2 = TableCellPropertiesRequest(
        source=dest1,
        destination=dest2,
        source_sha256=dest1_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp1,
        properties=TableCellPropertiesSpec(
            vertical_alignment="center",
            no_wrap=None,
        ),
    )
    report2 = author_table_cell_properties(req2)
    assert report2.passed is True
    assert report2.target_change_report["properties_after"]["no_wrap"] is True

    graph2 = PackageGraph.from_file(dest2)
    root2 = ET.fromstring(graph2.read_part_bytes(_DOCUMENT))
    tc_pr2 = (
        next(iter(root2.iter(word("tbl"))))
        .findall(word("tr"))[0]
        .findall(word("tc"))[0]
        .find(word("tcPr"))
    )
    assert tc_pr2.find(word("noWrap")) is not None


def test_author_table_cell_properties_nowrap_aliases_and_validation(
    tmp_path: Path,
) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    # CamelCase noWrap alias accepted
    dest = tmp_path / "out_alias.docx"
    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties={"noWrap": True},
    )
    assert req.properties["no_wrap"] is True
    report = author_table_cell_properties(req)
    assert report.passed is True

    # Incompatible conflicting alias
    with pytest.raises(ValueError, match="table_cell_properties_no_wrap_invalid"):
        TableCellPropertiesRequest(
            source=source,
            destination=tmp_path / "err.docx",
            source_sha256=source_sha,
            target={"table": 1, "row": 1, "cell": 1},
            structure_fingerprint_sha256=fp,
            properties={"no_wrap": True, "noWrap": False},
        )

    # Non-bool rejected
    for invalid_val in ("true", 1, 0, [True]):
        with pytest.raises(ValueError, match="table_cell_properties_no_wrap_invalid"):
            TableCellPropertiesRequest(
                source=source,
                destination=tmp_path / "err.docx",
                source_sha256=source_sha,
                target={"table": 1, "row": 1, "cell": 1},
                structure_fingerprint_sha256=fp,
                properties={"no_wrap": invalid_val},
            )


def test_discover_table_cell_properties_targets_nowrap(tmp_path: Path) -> None:
    source = _make_cell_doc(
        tmp_path / "source.docx",
        cell_properties={
            (0, 0): "<w:noWrap/>",
            (0, 1): '<w:noWrap w:val="0"/>',
        },
    )
    discovery = discover_table_cell_properties_targets(source)
    assert discovery["passed"] is True
    targets = {(t["row"], t["cell"]): t for t in discovery["targets"]}

    # Cell (1, 1) has no_wrap=True
    assert targets[(1, 1)]["current_properties"]["no_wrap"] is True
    # Cell (1, 2) has no_wrap=False
    assert targets[(1, 2)]["current_properties"]["no_wrap"] is False
    # Cell (2, 1) has no nowrap property
    assert "no_wrap" not in targets[(2, 1)]["current_properties"]


def test_author_table_cell_properties_combined_all_six(tmp_path: Path) -> None:
    source = _make_cell_doc(tmp_path / "source.docx")
    dest = tmp_path / "output_six.docx"
    source_sha = _sha256(source)
    fp = table_cell_properties_structure_fingerprint(source, table=1, row=1, cell=1)
    assert fp is not None

    req = TableCellPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "cell": 1},
        structure_fingerprint_sha256=fp,
        properties=TableCellPropertiesSpec(
            width=TableWidthSpec(value=2400, type="dxa"),
            vertical_alignment="center",
            shading=TableShadingSpec(fill="F0F0F0", val="clear", color="auto"),
            borders=TableBordersSpec(
                top=BorderSpec(val="single", sz=4, space=0, color="auto"),
                bottom=BorderSpec(val="single", sz=4, space=0, color="auto"),
            ),
            cell_margins=TableCellMarginsSpec(top=100, left=150, bottom=100, right=150),
            no_wrap=True,
        ),
    )

    report = author_table_cell_properties(req)
    assert report.passed is True
    assert report.target_change_report["properties_authored"] == [
        "borders",
        "cell_margins",
        "no_wrap",
        "shading",
        "vertical_alignment",
        "width",
    ]

    # Verify schema sequence order in output DOCX tcPr
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    word = partial(Q, _W_NS)
    cell_el = next(iter(root.iter(word("tbl")))).findall(word("tr"))[0].findall(word("tc"))[0]
    tc_pr = cell_el.find(word("tcPr"))
    assert tc_pr is not None
    child_tags = [c.tag.rsplit("}", 1)[-1] for c in tc_pr]
    expected_order = ["tcW", "tcBorders", "shd", "noWrap", "tcMar", "vAlign"]
    assert child_tags == expected_order
