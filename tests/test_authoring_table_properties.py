"""Tests for source-bound table properties authoring and discovery."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import (
    cell,
    grid,
    paragraph,
    row,
    table,
    write_package,
)
from wolfdocx.inspect import inspect_docx
from wolfdocx.package import PackageGraph
from wolfdocx.review import _word_namespace
from wolfdocx.table_properties import (
    BorderSpec,
    TableBordersSpec,
    TableCellMarginsSpec,
    TableCellSpacingSpec,
    TableIndentSpec,
    TableLayoutSpec,
    TableLookSpec,
    TablePropertiesReport,
    TablePropertiesRequest,
    TablePropertiesSpec,
    TablePropertiesTarget,
    TableShadingSpec,
    TableWidthSpec,
    author_table_properties,
    discover_table_properties_targets,
    load_table_properties_request,
    parse_table_properties_request,
    table_properties_structure_fingerprint,
)

_DOCUMENT = "word/document.xml"
_DUMMY_SHA = "a" * 64
_DUMMY_FP = "b" * 64
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_table_doc(
    path: Path,
    *,
    rows_data: list[list[str]] | None = None,
    grid_widths: list[int] | None = None,
    tbl_properties: str = "",
    wrapper: str = "",
    styles_extra: str = "",
    row_properties: str = "",
) -> Path:
    """Create a minimal valid DOCX with an explicit table."""
    data = rows_data if rows_data is not None else [["A", "B"], ["C", "D"]]
    cols = len(data[0])
    widths = grid_widths if grid_widths is not None else [2500] * cols
    table_grid = grid(widths)

    rows_xml: list[str] = []
    for r_idx, r_data in enumerate(data):
        cells_xml = "".join(cell(paragraph(txt)) for txt in r_data)
        r_pr = row_properties if r_idx == 0 else ""
        rows_xml.append(row(cells_xml, properties=r_pr))

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


def test_author_table_properties_combined(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "part": _DOCUMENT},
        structure_fingerprint_sha256=fp,
        properties={
            "width": {"value": 9000, "type": "dxa"},
            "alignment": "center",
            "borders": {
                "top": {"val": "single", "sz": 8, "color": "112233"},
                "left": {"val": "single", "sz": 8, "color": "112233"},
                "bottom": {"val": "double", "sz": 12, "color": "445566"},
                "right": {"val": "single", "sz": 8, "color": "112233"},
                "insideH": {"val": "dashed", "sz": 4, "color": "778899"},
                "insideV": {"val": "dotted", "sz": 4, "color": "AABBCC"},
            },
            "shading": {"fill": "EFEFEF", "val": "clear", "color": "auto"},
            "cell_margins": {"top": 120, "left": 160, "bottom": 120, "right": 160},
        },
    )

    report = author_table_properties(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True

    # Check inspection and XML
    insp = inspect_docx(dest)
    assert insp.valid is True

    disc = discover_table_properties_targets(dest)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 1
    t_rec = disc["targets"][0]["current_properties"]
    assert t_rec["width"]["value"] == 9000
    assert "top" in t_rec["borders"]
    assert t_rec["borders"]["top"]["val"] == "single"
    assert t_rec["borders"]["bottom"]["val"] == "double"
    # Verify XML elements
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl")
    assert tbl is not None
    tbl_pr = tbl.find(f"{{{_W_NS}}}tblPr")
    assert tbl_pr is not None

    jc = tbl_pr.find(f"{{{_W_NS}}}jc")
    assert jc is not None and jc.get(f"{{{_W_NS}}}val") == "center"

    shd = tbl_pr.find(f"{{{_W_NS}}}shd")
    assert shd is not None and shd.get(f"{{{_W_NS}}}fill") == "EFEFEF"

    cell_mar = tbl_pr.find(f"{{{_W_NS}}}tblCellMar")
    assert cell_mar is not None
    top_mar = cell_mar.find(f"{{{_W_NS}}}top")
    assert top_mar is not None and top_mar.get(f"{{{_W_NS}}}w") == "120"

    # Verify report change details
    change = report.target_change_report
    assert change is not None
    assert change["table"] == 1
    assert "width" in change["properties_authored"]
    assert "alignment" in change["properties_authored"]
    assert "borders" in change["properties_authored"]
    assert "shading" in change["properties_authored"]
    assert "cell_margins" in change["properties_authored"]
    assert change["boundaries"] == {"table": 1, "rows": 2, "columns": 2}


# --- 2. Styled Tables (tblStyle Preservation & Precedence) ---


def test_author_table_properties_styled_table_preserves_style(tmp_path: Path) -> None:
    source = _make_table_doc(
        tmp_path / "styled.docx",
        tbl_properties='<w:tblStyle w:val="ReportTable"/>',
    )
    dest = tmp_path / "styled_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "width": {"value": 8500, "type": "dxa"},
            "shading": "DDEEFF",
        },
    )

    report = author_table_properties(req)
    assert report.passed is True

    # Style must remain intact
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    style_elem = tbl_pr.find(f"{{{_W_NS}}}tblStyle")
    assert style_elem is not None
    assert style_elem.get(f"{{{_W_NS}}}val") == "ReportTable"

    # Direct formatting applied
    w_elem = tbl_pr.find(f"{{{_W_NS}}}tblW")
    assert w_elem is not None and w_elem.get(f"{{{_W_NS}}}w") == "8500"
    shd_elem = tbl_pr.find(f"{{{_W_NS}}}shd")
    assert shd_elem is not None and shd_elem.get(f"{{{_W_NS}}}fill") == "DDEEFF"

    assert report.target_change_report is not None
    assert report.target_change_report["style"] == "ReportTable"


# --- 3. Existing Properties Preservation & Selective Updating ---


def test_author_table_properties_preserves_unspecified_properties_and_siblings(
    tmp_path: Path,
) -> None:
    # Table has existing tblStyle, tblInd, and partial borders (top & left)
    initial_props = (
        '<w:tblStyle w:val="TableGrid"/>'
        '<w:tblInd w:w="360" w:type="dxa"/>'
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="6" w:color="000000"/>'
        '<w:left w:val="single" w:sz="6" w:color="000000"/>'
        "</w:tblBorders>"
    )
    source = _make_table_doc(tmp_path / "exist.docx", tbl_properties=initial_props)
    dest = tmp_path / "exist_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # Author only bottom border and alignment
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "alignment": "right",
            "borders": {
                "bottom": {"val": "double", "sz": 12, "color": "0000FF"},
            },
        },
    )

    report = author_table_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None

    # tblStyle and tblInd preserved
    assert tbl_pr.find(f"{{{_W_NS}}}tblStyle") is not None
    ind = tbl_pr.find(f"{{{_W_NS}}}tblInd")
    assert ind is not None and ind.get(f"{{{_W_NS}}}w") == "360"

    # jc added
    jc = tbl_pr.find(f"{{{_W_NS}}}jc")
    assert jc is not None and jc.get(f"{{{_W_NS}}}val") == "right"

    # borders: top and left preserved, bottom added
    tbl_borders = tbl_pr.find(f"{{{_W_NS}}}tblBorders")
    assert tbl_borders is not None
    top_b = tbl_borders.find(f"{{{_W_NS}}}top")
    assert top_b is not None and top_b.get(f"{{{_W_NS}}}sz") == "6"
    left_b = tbl_borders.find(f"{{{_W_NS}}}left")
    assert left_b is not None and left_b.get(f"{{{_W_NS}}}sz") == "6"
    bottom_b = tbl_borders.find(f"{{{_W_NS}}}bottom")
    assert bottom_b is not None and bottom_b.get(f"{{{_W_NS}}}val") == "double"


# --- 4. No Data Loss (Cells, Text, Grid & Unrelated Parts Intact) ---


def test_author_table_properties_no_data_loss(tmp_path: Path) -> None:
    rows_data = [
        ["Header 1", "Header 2", "Header 3"],
        ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
        ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"],
    ]
    source = _make_table_doc(
        tmp_path / "loss.docx",
        rows_data=rows_data,
        grid_widths=[2000, 3000, 4000],
    )
    dest = tmp_path / "loss_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 9000},
    )

    report = author_table_properties(req)
    assert report.passed is True
    assert report.non_target_parts_preserved is True

    # Read output and verify every cell text
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl")
    assert tbl is not None

    # Grid intact
    grid_cols = tbl.findall(f"{{{_W_NS}}}tblGrid/{{{_W_NS}}}gridCol")
    assert [int(col.get(f"{{{_W_NS}}}w") or 0) for col in grid_cols] == [
        2000,
        3000,
        4000,
    ]

    # Row and cell content intact
    tr_elems = tbl.findall(f"{{{_W_NS}}}tr")
    assert len(tr_elems) == 3
    for r_idx, tr in enumerate(tr_elems):
        tc_elems = tr.findall(f"{{{_W_NS}}}tc")
        assert len(tc_elems) == 3
        for c_idx, tc in enumerate(tc_elems):
            t_elem = tc.find(f".//{{{_W_NS}}}t")
            assert t_elem is not None
            assert t_elem.text == rows_data[r_idx][c_idx]


# --- 5. Schema Order Enforcement ---


def test_author_table_properties_schema_order(tmp_path: Path) -> None:
    # Initial properties have elements that should precede and succeed new ones
    initial_props = (
        '<w:tblStyle w:val="NormalTable"/>'
        '<w:tblInd w:w="200" w:type="dxa"/>'
        '<w:tblLook w:val="04A0"/>'
    )
    source = _make_table_doc(tmp_path / "order.docx", tbl_properties=initial_props)
    dest = tmp_path / "order_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "cell_margins": {"top": 100, "bottom": 100},
            "borders": {"top": {"val": "single", "sz": 4}},
            "width": 7200,
            "shading": "CCCCCC",
            "alignment": "left",
        },
    )

    report = author_table_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None

    tags = [c.tag.rsplit("}", 1)[-1] for c in tbl_pr]
    expected_order = [
        "tblStyle",
        "tblW",
        "jc",
        "tblInd",
        "tblBorders",
        "shd",
        "tblCellMar",
        "tblLook",
    ]
    assert tags == expected_order

    # Check borders order inside tblBorders
    borders_elem = tbl_pr.find(f"{{{_W_NS}}}tblBorders")
    assert borders_elem is not None
    b_tags = [c.tag.rsplit("}", 1)[-1] for c in borders_elem]
    assert b_tags == ["top"]

    # Check margins order inside tblCellMar
    mar_elem = tbl_pr.find(f"{{{_W_NS}}}tblCellMar")
    assert mar_elem is not None
    m_tags = [c.tag.rsplit("}", 1)[-1] for c in mar_elem]
    assert m_tags == ["top", "bottom"]


# --- 6. Namespace Stability ---


def test_author_table_properties_namespace_stability(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "ns.docx")
    dest = tmp_path / "ns_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 6000, "alignment": "center"},
    )

    report = author_table_properties(req)
    assert report.passed is True

    dest_bytes = PackageGraph.from_file(dest).read_part_bytes(_DOCUMENT)
    raw_xml = dest_bytes.decode("utf-8")
    assert 'xmlns=""' not in raw_xml

    root = ET.fromstring(dest_bytes)
    ns = _word_namespace(root)
    assert ns == _W_NS
    tbl = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl")
    assert tbl is not None
    tbl_pr = tbl.find(f"{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    assert tbl_pr.find(f"{{{_W_NS}}}tblW") is not None
    assert tbl_pr.find(f"{{{_W_NS}}}jc") is not None


# --- 7. API and JSON Strict Validation ---


def test_parse_table_properties_request_valid() -> None:
    data = {
        "source": "in.docx",
        "destination": "out.docx",
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1, "part": _DOCUMENT},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {
            "width": {"value": 5000, "type": "dxa"},
            "alignment": "center",
            "borders": {"top": {"val": "single", "sz": 4, "space": 0, "color": "auto"}},
            "shading": {"fill": "AABBCC", "val": "clear", "color": "auto"},
            "cell_margins": {"top": 100},
        },
    }
    req = parse_table_properties_request(data)
    assert req.table == 1
    assert req.part == _DOCUMENT
    assert req.source == Path("in.docx")
    assert req.destination == Path("out.docx")
    assert req.properties["width"]["value"] == 5000
    assert req.properties["alignment"] == "center"

    # to_dict round-trip
    dumped = req.to_dict()
    assert dumped["target"]["table"] == 1
    assert dumped["properties"]["width"]["value"] == 5000


def test_parse_table_properties_request_invalid_types_and_fields() -> None:
    # Not mapping
    with pytest.raises(TypeError, match="table_properties_request_invalid"):
        parse_table_properties_request("not-a-dict")

    # Missing field
    with pytest.raises(ValueError, match="table_properties_request_fields_invalid"):
        parse_table_properties_request({"source": "a.docx"})

    # Unknown request field
    with pytest.raises(ValueError, match="table_properties_request_fields_invalid"):
        parse_table_properties_request(
            {
                "source": "a.docx",
                "destination": "b.docx",
                "source_sha256": _DUMMY_SHA,
                "target": {"table": 1},
                "structure_fingerprint_sha256": _DUMMY_FP,
                "properties": {"width": 5000},
                "unknown_extra": 123,
            }
        )

    # Empty source / destination
    base = {
        "source": "a.docx",
        "destination": "b.docx",
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {"width": 5000},
    }
    with pytest.raises(ValueError, match="table_properties_source_invalid"):
        parse_table_properties_request({**base, "source": ""})
    with pytest.raises(ValueError, match="table_properties_destination_invalid"):
        parse_table_properties_request({**base, "destination": ""})

    # Invalid target
    with pytest.raises(TypeError, match="table_properties_target_invalid"):
        parse_table_properties_request({**base, "target": 1})
    with pytest.raises(ValueError, match="table_properties_target_fields_invalid"):
        parse_table_properties_request({**base, "target": {"table": 1, "extra": 2}})
    with pytest.raises(ValueError, match="table_properties_table_invalid"):
        parse_table_properties_request({**base, "target": {"table": 0}})
    with pytest.raises(ValueError, match="table_properties_table_invalid"):
        parse_table_properties_request({**base, "target": {"table": True}})
    with pytest.raises(ValueError, match="table_properties_part_invalid"):
        parse_table_properties_request({**base, "target": {"table": 1, "part": ""}})

    # Empty properties
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request({**base, "properties": {}})

    # Unknown property
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request({**base, "properties": {"unknown_prop": 1}})

    # Invalid width
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request({**base, "properties": {"width": -10}})
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request(
            {**base, "properties": {"width": {"value": 5000, "type": "unknown"}}}
        )

    # Invalid alignment
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request({**base, "properties": {"alignment": "justify"}})

    # Invalid borders
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request(
            {**base, "properties": {"borders": {"diagonal": {"val": "single"}}}}
        )
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request(
            {**base, "properties": {"borders": {"top": {"val": "unsupported_style"}}}}
        )
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request(
            {**base, "properties": {"borders": {"top": {"color": "NOT_HEX"}}}}
        )

    # Invalid shading
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request({**base, "properties": {"shading": "BAD_COLOR"}})

    # Invalid cell margins
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        parse_table_properties_request({**base, "properties": {"cell_margins": {"top": -5}}})


def test_table_properties_request_direct_dataclass_validation() -> None:
    # Direct construction must perform same strict checks as parse
    with pytest.raises(TypeError, match="table_properties_source_invalid"):
        TablePropertiesRequest(
            source=123,  # type: ignore[arg-type]
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"width": 5000},
        )

    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={},
        )


def test_property_spec_dataclasses_round_trip() -> None:
    w = TableWidthSpec(value=8000, type="dxa")
    assert w.to_dict() == {"value": 8000, "type": "dxa"}

    b = BorderSpec(val="double", sz=8, space=0, color="123456")
    assert b.to_dict() == {"val": "double", "sz": 8, "space": 0, "color": "123456"}

    tb = TableBordersSpec(top=b, bottom=BorderSpec(val="single"))
    tb_dict = tb.to_dict()
    assert "top" in tb_dict
    assert tb_dict["top"]["val"] == "double"
    assert tb_dict["bottom"]["val"] == "single"

    shd = TableShadingSpec(fill="AABBCC", val="clear", color="auto")
    assert shd.to_dict() == {"fill": "AABBCC", "val": "clear", "color": "auto"}

    cm = TableCellMarginsSpec(top=100, right=200)
    assert cm.to_dict() == {"top": 100, "right": 200}

    spec = TablePropertiesSpec(width=w, borders=tb, shading=shd, cell_margins=cm)
    spec_dict = spec.to_dict()
    assert spec_dict["width"]["value"] == 8000
    assert spec_dict["shading"]["fill"] == "AABBCC"


# --- 8. Discovery and Report JSON ---


def test_discover_table_properties_targets(tmp_path: Path) -> None:
    # Document with 2 tables: one normal, one with existing style
    doc = write_package(
        tmp_path / "multi.docx",
        table(
            row(cell(paragraph("T1"))),
            properties='<w:tblW w:w="5000" w:type="dxa"/>',
            grid=grid([5000]),
        )
        + paragraph("Between")
        + table(
            row(cell(paragraph("T2"))),
            properties='<w:tblStyle w:val="GridTable"/>',
            grid=grid([5000]),
        ),
    )

    discovery = discover_table_properties_targets(doc)
    assert discovery["schema_version"] == 1
    assert discovery["operation"] == "table-properties-targets"
    assert discovery["passed"] is True
    assert discovery["status"] == "discovered"
    assert discovery["code"] is None
    assert discovery["part"] == _DOCUMENT
    assert len(discovery["targets"]) == 2

    t1 = discovery["targets"][0]
    assert t1["table"] == 1
    assert t1["current_properties"]["width"]["value"] == 5000

    t2 = discovery["targets"][1]
    assert t2["table"] == 2
    assert t2["style"] == "GridTable"


def test_discover_table_properties_targets_refusals(tmp_path: Path) -> None:
    # Missing source
    disc = discover_table_properties_targets(tmp_path / "absent.docx")
    assert disc["passed"] is False
    assert disc["status"] == "refused"
    assert disc["code"] == "table_properties_source_missing"

    # Unsupported part
    source = _make_table_doc(tmp_path / "valid.docx")
    disc_part = discover_table_properties_targets(source, part="word/header1.xml")
    assert disc_part["passed"] is False
    assert disc_part["code"] == "table_properties_part_unsupported"


def test_load_table_properties_request(tmp_path: Path) -> None:
    req_file = tmp_path / "request.json"
    req_data = {
        "source": "in.docx",
        "destination": "out.docx",
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {"width": 4500},
    }
    req_file.write_text(json.dumps(req_data), encoding="utf-8")

    req = load_table_properties_request(req_file)
    assert req.table == 1
    assert req.properties["width"]["value"] == 4500


def test_table_properties_target_to_dict() -> None:
    tgt = TablePropertiesTarget(
        table=1,
        structure_fingerprint_sha256=_DUMMY_FP,
        part=_DOCUMENT,
        current_properties={"width": {"value": 5000, "type": "dxa"}},
        style="Normal",
        rows_count=3,
        columns_count=2,
    )
    assert tgt.table == 1
    assert tgt.target == {"table": 1, "part": _DOCUMENT}
    d = tgt.to_dict()
    assert d["table"] == 1
    assert d["rows_count"] == 3
    assert d["columns_count"] == 2


def test_table_properties_report_to_dict() -> None:
    report = TablePropertiesReport(
        status="published",
        code=None,
        source_sha256=_DUMMY_SHA,
        destination_sha256=_DUMMY_SHA,
        target={"table": 1, "part": _DOCUMENT},
        properties={"width": {"value": 5000, "type": "dxa"}},
        target_change_report={
            "table": 1,
            "boundaries": {"table": 1, "rows": 1, "columns": 1},
        },
        non_target_parts_preserved=True,
    )
    d = report.to_dict()
    assert d["schema_version"] == 1
    assert d["operation"] == "table-properties"
    assert d["passed"] is True
    assert d["status"] == "published"
    assert d["destination_sha256"] == _DUMMY_SHA
    assert d["target_change_report"]["boundaries"]["rows"] == 1


# --- 9. Atomic Refusals & Safety Boundaries ---


def test_author_table_properties_refusal_source_missing(tmp_path: Path) -> None:
    req = TablePropertiesRequest(
        source=tmp_path / "absent.docx",
        destination=tmp_path / "out.docx",
        source_sha256=_DUMMY_SHA,
        target={"table": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_source_missing"
    assert not (tmp_path / "out.docx").exists()


def test_author_table_properties_refusal_destination_occupied(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    dest = tmp_path / "occupied.docx"
    dest.write_text("existing content", encoding="utf-8")
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_destination_occupied"
    # Destination must be untouched
    assert dest.read_text(encoding="utf-8") == "existing content"


def test_author_table_properties_refusal_source_hash_mismatch(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256="f" * 64,  # wrong hash
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_source_hash_mismatch"
    assert not dest.exists()


def test_author_table_properties_refusal_fingerprint_mismatch(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256="f" * 64,  # stale fingerprint
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_fingerprint_mismatch"
    assert not dest.exists()


def test_author_table_properties_refusal_target_missing(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 99},  # only 1 table exists
        structure_fingerprint_sha256=fp,
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_target_missing"
    assert not dest.exists()


def test_author_table_properties_refusal_part_unsupported(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "part": "word/header1.xml"},
        structure_fingerprint_sha256=fp,
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_part_unsupported"
    assert not dest.exists()


def test_author_table_properties_refusal_context_unsupported(tmp_path: Path) -> None:
    # Table inside tracked revision <w:ins>
    source = _make_table_doc(tmp_path / "ins.docx", wrapper="ins")
    dest = tmp_path / "ins_out.docx"
    fp = _sha256(source)  # dummy fp since fingerprint helper refuses unsupported context

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 5000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code in (
        "table_properties_context_unsupported",
        "table_properties_fingerprint_mismatch",
    )
    assert not dest.exists()


def test_author_table_properties_refusal_conflict(tmp_path: Path) -> None:
    # Duplicate tblW inside tblPr causes conflict
    dup_props = '<w:tblW w:w="5000" w:type="dxa"/><w:tblW w:w="6000" w:type="dxa"/>'
    source = _make_table_doc(tmp_path / "conflict.docx", tbl_properties=dup_props)
    dest = tmp_path / "conflict_out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 7000},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.code == "table_properties_conflict"
    assert not dest.exists()


# --- 10. Narrow Renderer Metadata (Conditional Styles & Row Overrides) ---


def test_author_table_properties_with_conditional_table_style(tmp_path: Path) -> None:
    # Table style in styles.xml has tblStylePr (conditional formatting)
    style_xml = (
        '<w:style w:type="table" w:styleId="ConditionalStyle">'
        '<w:tblStylePr w:type="firstRow"/>'
        "</w:style>"
    )
    source = _make_table_doc(
        tmp_path / "cond.docx",
        tbl_properties='<w:tblStyle w:val="ConditionalStyle"/>',
        styles_extra=style_xml,
    )
    dest = tmp_path / "cond_out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    disc = discover_table_properties_targets(source)
    assert disc["passed"] is True

    # Authoring should succeed safely, not be refused!
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": 8000},
    )
    report = author_table_properties(req)
    assert report.passed is True
    assert report.status == "published"


def test_author_table_properties_with_row_tbl_pr_ex(
    tmp_path: Path,
) -> None:
    # Table row has tblPrEx as direct child of w:tr (ECMA-376 §17.4.59)
    row_xml = (
        "<w:tr>"
        '<w:tblPrEx><w:tblW w:w="4000" w:type="dxa"/></w:tblPrEx>'
        f"{cell(paragraph('A'))}{cell(paragraph('B'))}"
        "</w:tr>"
    )
    tbl_xml = table(row_xml, grid=grid([2500, 2500]))
    source = write_package(tmp_path / "row_ex.docx", tbl_xml)
    dest = tmp_path / "row_ex_out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    disc = discover_table_properties_targets(source)
    assert disc["passed"] is True

    # Authoring succeeds safely, not refused!
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"shading": "FFFF00"},
    )
    report = author_table_properties(req)
    assert report.passed is True
    assert report.status == "published"


def test_author_table_properties_authored_zero_width(
    tmp_path: Path,
) -> None:
    source = _make_table_doc(tmp_path / "zero_w.docx")
    dest = tmp_path / "zero_w_out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # Authoring width: 0 with dxa is valid authoring per OOXML
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"width": {"value": 0, "type": "dxa"}},
    )
    report = author_table_properties(req)
    assert report.passed is True
    assert report.status == "published"


def test_author_table_properties_malformed_styles_refuses(
    tmp_path: Path,
) -> None:
    source = _make_table_doc(
        tmp_path / "corrupt_styles.docx",
        styles_extra="<UNCLOSED_TAG",
    )
    dest = tmp_path / "corrupt_styles_out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    disc = discover_table_properties_targets(source)
    assert disc["passed"] is True

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"alignment": "center"},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_staged_verification_failed"


def test_author_table_properties_with_inherited_conditional_style(
    tmp_path: Path,
) -> None:
    # Base style has tblStylePr; Child style inherits it via basedOn
    styles_xml = (
        '<w:style w:type="table" w:styleId="BaseStyle">'
        '<w:tblStylePr w:type="firstRow"/>'
        "</w:style>"
        '<w:style w:type="table" w:styleId="ChildStyle">'
        '<w:basedOn w:val="BaseStyle"/>'
        "</w:style>"
    )
    source = _make_table_doc(
        tmp_path / "inherited_cond.docx",
        tbl_properties='<w:tblStyle w:val="ChildStyle"/>',
        styles_extra=styles_xml,
    )
    dest = tmp_path / "inherited_cond_out.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    disc = discover_table_properties_targets(source)
    assert disc["passed"] is True

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"shading": "DDDDDD"},
    )
    report = author_table_properties(req)
    assert report.passed is True


# --- 11. Table Style Binding & Clearing (properties.style) ---


def test_author_table_properties_set_style(tmp_path: Path) -> None:
    custom_style_xml = (
        '<w:style w:type="table" w:styleId="CustomTable1">'
        '<w:name w:val="Custom Table 1"/>'
        '<w:tblPr><w:tblCellMar><w:top w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        "</w:style>"
    )
    source = _make_table_doc(
        tmp_path / "source_set_style.docx",
        styles_extra=custom_style_xml,
    )
    dest = tmp_path / "out_set_style.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"style": "CustomTable1"},
    )
    report = author_table_properties(req)
    assert report.passed is True
    assert report.code is None
    assert report.target_change_report is not None
    assert report.target_change_report["style"] == "CustomTable1"
    assert "style" in report.target_change_report["properties_authored"]

    # Verify staged output has tblStyle as first child of tblPr
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    assert len(tbl_pr) >= 1
    first_child = tbl_pr[0]
    assert first_child.tag == f"{{{_W_NS}}}tblStyle"
    assert first_child.get(f"{{{_W_NS}}}val") == "CustomTable1"


def test_author_table_properties_clear_style(tmp_path: Path) -> None:
    custom_style_xml = (
        '<w:style w:type="table" w:styleId="CustomTable1">'
        '<w:name w:val="Custom Table 1"/>'
        "</w:style>"
    )
    initial_props = '<w:tblStyle w:val="CustomTable1"/><w:tblW w:w="8000" w:type="dxa"/>'
    source = _make_table_doc(
        tmp_path / "source_clear_style.docx",
        tbl_properties=initial_props,
        styles_extra=custom_style_xml,
    )
    dest = tmp_path / "out_clear_style.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"style": None},
    )
    report = author_table_properties(req)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["style"] is None

    # Verify tblStyle is removed, but tblW is preserved
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    assert tbl_pr.find(f"{{{_W_NS}}}tblStyle") is None
    tbl_w = tbl_pr.find(f"{{{_W_NS}}}tblW")
    assert tbl_w is not None
    assert tbl_w.get(f"{{{_W_NS}}}w") == "8000"


def test_author_table_properties_style_missing_refusal(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source_missing_style.docx")
    dest = tmp_path / "out_missing.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"style": "NonExistentTableStyle"},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_style_missing"


def test_author_table_properties_style_type_incompatible_refusal(
    tmp_path: Path,
) -> None:
    source = _make_table_doc(tmp_path / "source_para_style.docx")
    dest = tmp_path / "out_para.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # "Heading" is type="paragraph" in default styles.xml
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"style": "Heading"},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_style_type_incompatible"


def test_author_table_properties_style_unsupported_conditional_formatting(
    tmp_path: Path,
) -> None:
    styles_xml = (
        '<w:style w:type="table" w:styleId="BandedTable">'
        '<w:tblStylePr w:type="firstRow"/>'
        "</w:style>"
    )
    source = _make_table_doc(
        tmp_path / "source_banded.docx",
        styles_extra=styles_xml,
    )
    dest = tmp_path / "out_banded.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"style": "BandedTable"},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_style_unsupported"


def test_author_table_properties_style_unsupported_cycle(tmp_path: Path) -> None:
    styles_xml = (
        '<w:style w:type="table" w:styleId="CycleTableA">'
        '<w:basedOn w:val="CycleTableB"/>'
        "</w:style>"
        '<w:style w:type="table" w:styleId="CycleTableB">'
        '<w:basedOn w:val="CycleTableA"/>'
        "</w:style>"
    )
    source = _make_table_doc(
        tmp_path / "source_cycle.docx",
        styles_extra=styles_xml,
    )
    dest = tmp_path / "out_cycle.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"style": "CycleTableA"},
    )
    report = author_table_properties(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_style_unsupported"


def test_author_table_properties_style_direct_format_precedence_and_schema_order(
    tmp_path: Path,
) -> None:
    custom_style_xml = (
        '<w:style w:type="table" w:styleId="StyledTable"><w:name w:val="Styled Table"/></w:style>'
    )
    initial_props = '<w:tblW w:w="7200" w:type="dxa"/><w:jc w:val="center"/>'
    source = _make_table_doc(
        tmp_path / "source_precedence.docx",
        tbl_properties=initial_props,
        styles_extra=custom_style_xml,
    )
    dest = tmp_path / "out_precedence.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "style": "StyledTable",
            "shading": "EEEEEE",
        },
    )
    report = author_table_properties(req)
    assert report.passed is True

    # Check schema order in tblPr: tblStyle (rank 0), tblW (rank 6), jc (rank 7), shd (rank 11)
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    tags = [c.tag.rsplit("}", 1)[-1] for c in tbl_pr]
    assert tags == ["tblStyle", "tblW", "jc", "shd"]

    # Direct formatting preserved
    assert tbl_pr.find(f"{{{_W_NS}}}tblW").get(f"{{{_W_NS}}}w") == "7200"
    assert tbl_pr.find(f"{{{_W_NS}}}jc").get(f"{{{_W_NS}}}val") == "center"
    assert tbl_pr.find(f"{{{_W_NS}}}shd").get(f"{{{_W_NS}}}fill") == "EEEEEE"


def test_author_table_properties_source_rebinding_after_publication(
    tmp_path: Path,
) -> None:
    custom_style_xml = (
        '<w:style w:type="table" w:styleId="RebindTable"><w:name w:val="Rebind Table"/></w:style>'
    )
    source1 = _make_table_doc(
        tmp_path / "rebind_source.docx",
        styles_extra=custom_style_xml,
    )
    dest1 = tmp_path / "rebind_step1.docx"
    fp1 = table_properties_structure_fingerprint(source1, table=1)
    assert fp1 is not None

    # Step 1: Set style and width
    req1 = TablePropertiesRequest(
        source=source1,
        destination=dest1,
        source_sha256=_sha256(source1),
        target={"table": 1},
        structure_fingerprint_sha256=fp1,
        properties={"style": "RebindTable", "width": {"value": 8200, "type": "dxa"}},
    )
    report1 = author_table_properties(req1)
    assert report1.passed is True
    assert report1.destination_sha256 is not None

    # Step 2: Source rebinding - use dest1 as source with fresh sha256 and structure fingerprint
    dest2 = tmp_path / "rebind_step2.docx"
    source2_sha = _sha256(dest1)
    assert source2_sha == report1.destination_sha256
    fp2 = table_properties_structure_fingerprint(dest1, table=1)
    assert fp2 is not None

    # Clear style in step 2 while preserving width
    req2 = TablePropertiesRequest(
        source=dest1,
        destination=dest2,
        source_sha256=source2_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp2,
        properties={"style": None},
    )
    report2 = author_table_properties(req2)
    assert report2.passed is True
    assert report2.target_change_report["style"] is None

    # Output of step 2 has style cleared, width preserved
    graph2 = PackageGraph.from_file(dest2)
    root2 = ET.fromstring(graph2.read_part_bytes(_DOCUMENT))
    tbl_pr2 = root2.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr2 is not None
    assert tbl_pr2.find(f"{{{_W_NS}}}tblStyle") is None
    assert tbl_pr2.find(f"{{{_W_NS}}}tblW").get(f"{{{_W_NS}}}w") == "8200"


def test_author_table_properties_unchanged_ordinary_properties(tmp_path: Path) -> None:
    custom_style_xml = (
        '<w:style w:type="table" w:styleId="StableTable"><w:name w:val="Stable Table"/></w:style>'
    )
    initial_props = (
        '<w:tblStyle w:val="StableTable"/>'
        '<w:tblW w:w="6000" w:type="dxa"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="AABBCC"/>'
    )
    source = _make_table_doc(
        tmp_path / "source_unchanged.docx",
        tbl_properties=initial_props,
        styles_extra=custom_style_xml,
    )
    dest = tmp_path / "out_unchanged.docx"
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # Author only alignment - style and other properties must remain unchanged
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"alignment": "right"},
    )
    report = author_table_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    assert tbl_pr.find(f"{{{_W_NS}}}tblStyle").get(f"{{{_W_NS}}}val") == "StableTable"
    assert tbl_pr.find(f"{{{_W_NS}}}tblW").get(f"{{{_W_NS}}}w") == "6000"
    assert tbl_pr.find(f"{{{_W_NS}}}shd").get(f"{{{_W_NS}}}fill") == "AABBCC"
    assert tbl_pr.find(f"{{{_W_NS}}}jc").get(f"{{{_W_NS}}}val") == "right"


def test_author_table_properties_spec_dataclass_style() -> None:
    spec_set = TablePropertiesSpec(style="MyTable")
    assert spec_set.to_dict() == {"style": "MyTable"}

    spec_clear = TablePropertiesSpec(style=None)
    assert spec_clear.to_dict() == {"style": None}

    spec_unset = TablePropertiesSpec(width=TableWidthSpec(5000))
    assert spec_unset.to_dict() == {"width": {"value": 5000, "type": "dxa"}}


def test_author_table_properties_validation_refusals_style() -> None:
    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"style": 123},
        )

    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"style": True},
        )

    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"style": "   "},
        )


def test_author_table_properties_spec_dataclass_indent_spacing() -> None:
    spec_indent = TablePropertiesSpec(indent=TableIndentSpec(720))
    assert spec_indent.to_dict() == {"indent": {"value": 720, "type": "dxa"}}

    spec_spacing = TablePropertiesSpec(cell_spacing=TableCellSpacingSpec(0))
    assert spec_spacing.to_dict() == {"cell_spacing": {"value": 0, "type": "dxa"}}

    spec_clear_indent = TablePropertiesSpec(indent=None)
    assert spec_clear_indent.to_dict() == {"indent": None}

    spec_clear_spacing = TablePropertiesSpec(cell_spacing=None)
    assert spec_clear_spacing.to_dict() == {"cell_spacing": None}

    spec_both = TablePropertiesSpec(
        indent=TableIndentSpec(360),
        cell_spacing=TableCellSpacingSpec(0),
    )
    assert spec_both.to_dict() == {
        "indent": {"value": 360, "type": "dxa"},
        "cell_spacing": {"value": 0, "type": "dxa"},
    }


def test_author_table_properties_validation_refusals_indent_spacing() -> None:
    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"indent": True},
        )

    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"cell_spacing": False},
        )

    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"indent": 500000},
        )

    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"cell_spacing": -20},
        )


def test_author_table_properties_spec_dataclass_layout_look() -> None:
    spec_layout = TablePropertiesSpec(layout=TableLayoutSpec("fixed"))
    assert spec_layout.to_dict() == {"layout": {"type": "fixed"}}

    spec_autofit = TablePropertiesSpec(layout=TableLayoutSpec("autofit"))
    assert spec_autofit.to_dict() == {"layout": {"type": "autofit"}}

    spec_clear_layout = TablePropertiesSpec(layout=None)
    assert spec_clear_layout.to_dict() == {"layout": None}

    spec_look = TablePropertiesSpec(look=TableLookSpec(first_row=True, no_vband=True, val="0420"))
    assert spec_look.to_dict() == {"look": {"first_row": True, "no_vband": True, "val": "0420"}}

    spec_clear_look = TablePropertiesSpec(look=None)
    assert spec_clear_look.to_dict() == {"look": None}

    spec_both = TablePropertiesSpec(
        layout=TableLayoutSpec("fixed"),
        look=TableLookSpec(first_row=False),
    )
    assert spec_both.to_dict() == {
        "layout": {"type": "fixed"},
        "look": {"first_row": False},
    }


def test_author_table_properties_validation_refusals_layout_look() -> None:
    # Invalid layout string
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"layout": "expanded"},
        )

    # Layout using invented w:val instead of normative w:type
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"layout": {"val": "fixed"}},
        )

    # Layout non-string type
    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"layout": 123},
        )

    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"layout": True},
        )

    # Look non-boolean flag (strict types)
    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {"first_row": "true"}},
        )

    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {"first_row": 1}},
        )

    # Look invalid hex val
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {"val": "XYZ"}},
        )

    # Look non-string val
    with pytest.raises(TypeError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {"val": 1184}},
        )

    # Look incoherent flag vs val
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {"val": "04A0", "first_row": False}},
        )

    # Look empty mapping
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {}},
        )

    # Look unknown key
    with pytest.raises(ValueError, match="table_properties_fields_invalid"):
        TablePropertiesRequest(
            source=Path("in.docx"),
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"look": {"unexpected_key": True}},
        )


def test_author_table_properties_layout_look_lifecycle(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "lifecycle.docx")
    dest = tmp_path / "lifecycle_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # Discover before authoring
    targets_before = discover_table_properties_targets(source)
    assert targets_before["passed"] is True
    props_initial = targets_before["targets"][0]["current_properties"]
    assert "layout" not in props_initial
    assert "look" not in props_initial

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "layout": "fixed",
            "look": {"first_row": True, "no_vband": True},
        },
    )
    rep = author_table_properties(req)
    assert rep.passed is True
    assert rep.status == "published"
    assert rep.target_change_report is not None
    assert rep.target_change_report["properties_after"]["layout"] == {"type": "fixed"}
    assert rep.target_change_report["properties_after"]["look"]["first_row"] is True
    assert rep.target_change_report["properties_after"]["look"]["no_vband"] is True
    assert rep.target_change_report["properties_after"]["look"]["val"] == "0420"

    # Verify OOXML structure and child element ordering
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None

    layout_elem = tbl_pr.find(f"{{{_W_NS}}}tblLayout")
    assert layout_elem is not None
    assert layout_elem.get(f"{{{_W_NS}}}type") == "fixed"

    look_elem = tbl_pr.find(f"{{{_W_NS}}}tblLook")
    assert look_elem is not None
    assert look_elem.get(f"{{{_W_NS}}}val") == "0420"
    assert look_elem.get(f"{{{_W_NS}}}firstRow") == "1"
    assert look_elem.get(f"{{{_W_NS}}}noVBand") == "1"

    tags = [c.tag.rsplit("}", 1)[-1] for c in tbl_pr]
    # Schema order: tblLayout precedes tblCellMar precedes tblLook
    assert tags.index("tblLayout") < tags.index("tblLook")

    # Reopen and discover on authored output
    targets_after = discover_table_properties_targets(dest)
    assert targets_after["passed"] is True
    props_discovered = targets_after["targets"][0]["current_properties"]
    assert props_discovered["layout"] == {"type": "fixed"}
    assert props_discovered["look"]["first_row"] is True
    assert props_discovered["look"]["no_vband"] is True
    assert props_discovered["look"]["val"] == "0420"


def test_author_table_properties_look_false_omitted_distinction(tmp_path: Path) -> None:
    # Start with table having w:tblLook w:val="04A0" (first_row=True, first_column=True, no_vband=True)
    initial_props = '<w:tblLook w:val="04A0" w:firstRow="1" w:firstColumn="1" w:noVBand="1"/>'
    source = _make_table_doc(tmp_path / "look_distinction.docx", tbl_properties=initial_props)
    dest = tmp_path / "look_distinction_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # Author setting first_row=False, while leaving other flags omitted!
    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "look": {"first_row": False},
        },
    )
    rep = author_table_properties(req)
    assert rep.passed is True
    assert rep.target_change_report is not None
    # first_row is explicitly False; first_column and no_vband remain True!
    assert rep.target_change_report["properties_after"]["look"]["first_row"] is False
    assert rep.target_change_report["properties_after"]["look"]["first_column"] is True
    assert rep.target_change_report["properties_after"]["look"]["no_vband"] is True
    # 0x04A0 & ~0x0020 = 0x0480
    assert rep.target_change_report["properties_after"]["look"]["val"] == "0480"

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    look_elem = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr/{{{_W_NS}}}tblLook")
    assert look_elem is not None
    assert look_elem.get(f"{{{_W_NS}}}firstRow") == "0"
    assert look_elem.get(f"{{{_W_NS}}}firstColumn") == "1"
    assert look_elem.get(f"{{{_W_NS}}}noVBand") == "1"
    assert look_elem.get(f"{{{_W_NS}}}val") == "0480"


def test_author_table_properties_clear_layout_and_look(tmp_path: Path) -> None:
    initial_props = '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/>'
    source = _make_table_doc(tmp_path / "clear.docx", tbl_properties=initial_props)
    dest = tmp_path / "clear_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    req = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={
            "layout": None,
            "look": None,
        },
    )
    rep = author_table_properties(req)
    assert rep.passed is True
    assert rep.target_change_report is not None
    assert "layout" not in rep.target_change_report["properties_after"]
    assert "look" not in rep.target_change_report["properties_after"]

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tbl_pr = root.find(f"{{{_W_NS}}}body/{{{_W_NS}}}tbl/{{{_W_NS}}}tblPr")
    assert tbl_pr is not None
    assert tbl_pr.find(f"{{{_W_NS}}}tblLayout") is None
    assert tbl_pr.find(f"{{{_W_NS}}}tblLook") is None


def test_author_table_properties_json_lifecycle_layout_look(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "json_in.docx")
    dest = tmp_path / "json_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    raw_json_dict = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": source_sha,
        "target": {"table": 1},
        "structure_fingerprint_sha256": fp,
        "properties": {
            "layout": {"type": "autofit"},
            "look": {
                "first_row": True,
                "last_row": False,
                "first_column": False,
                "last_column": False,
                "no_hband": False,
                "no_vband": True,
                "val": "0420",
            },
        },
    }
    req_path = tmp_path / "request.json"
    req_path.write_text(json.dumps(raw_json_dict), encoding="utf-8")

    loaded_req = load_table_properties_request(req_path)
    assert loaded_req.properties["layout"] == {"type": "autofit"}
    assert loaded_req.properties["look"]["first_row"] is True
    assert loaded_req.properties["look"]["val"] == "0420"

    rep = author_table_properties(loaded_req)
    assert rep.passed is True
    assert rep.status == "published"
    assert dest.is_file()
    assert inspect_docx(dest).valid is True


def test_author_table_properties_layout_look_refusals(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "refusals.docx")
    dest = tmp_path / "refusals_out.docx"
    source_sha = _sha256(source)
    fp = table_properties_structure_fingerprint(source, table=1)
    assert fp is not None

    # Stale source hash refusal
    req_stale_sha = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256="f" * 64,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"layout": "fixed"},
    )
    rep_stale_sha = author_table_properties(req_stale_sha)
    assert rep_stale_sha.passed is False
    assert rep_stale_sha.code == "table_properties_source_hash_mismatch"

    # Stale fingerprint refusal
    req_stale_fp = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256="0" * 64,
        properties={"layout": "fixed"},
    )
    rep_stale_fp = author_table_properties(req_stale_fp)
    assert rep_stale_fp.passed is False
    assert rep_stale_fp.code == "table_properties_fingerprint_mismatch"

    # Occupied destination refusal
    dest.write_text("occupied", encoding="utf-8")
    req_occupied = TablePropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
        properties={"layout": "fixed"},
    )
    rep_occupied = author_table_properties(req_occupied)
    assert rep_occupied.passed is False
    assert rep_occupied.code == "table_properties_destination_occupied"
