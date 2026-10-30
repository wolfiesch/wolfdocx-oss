"""Tests for source-bound table row properties authoring and discovery."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any
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
from wolfdocx.inspect import Q, inspect_docx
from wolfdocx.package import PackageGraph
from wolfdocx.table_row_properties import (
    RowHeightSpec,
    TableRowPropertiesReport,
    TableRowPropertiesRequest,
    TableRowPropertiesSpec,
    TableRowPropertiesTarget,
    author_table_row_properties,
    discover_table_row_properties_targets,
    load_table_row_properties_request,
    parse_table_row_properties_request,
    table_row_properties_structure_fingerprint,
)

_DOCUMENT = "word/document.xml"
_DUMMY_SHA = "a" * 64
_DUMMY_FP = "b" * 64
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_row_doc(
    path: Path,
    *,
    rows_data: list[list[str]] | None = None,
    grid_widths: list[int] | None = None,
    row_properties: list[str] | None = None,
    tbl_properties: str = "",
    wrapper: str = "",
    styles_extra: str = "",
) -> Path:
    """Create a minimal valid DOCX with an explicit table and per-row properties."""
    data = rows_data if rows_data is not None else [["Header A", "Header B"], ["Data 1", "Data 2"]]
    cols = len(data[0])
    widths = grid_widths if grid_widths is not None else [2500] * cols
    table_grid = grid(widths)

    rows_xml: list[str] = []
    for r_idx, r_data in enumerate(data):
        cells_xml = "".join(cell(paragraph(txt)) for txt in r_data)
        r_pr = row_properties[r_idx] if row_properties and r_idx < len(row_properties) else ""
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


# --- 1. Combined Property Authoring & Child Order ---


def test_author_table_row_properties_combined(tmp_path: Path) -> None:
    source = _make_row_doc(tmp_path / "source.docx")
    dest = tmp_path / "output.docx"
    source_sha = _sha256(source)

    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1, "part": _DOCUMENT},
        structure_fingerprint_sha256=fp,
        properties={
            "repeat_header": True,
            "cant_split": True,
            "height": {"value": 720, "rule": "atLeast"},
        },
    )

    report = author_table_row_properties(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == source_sha
    assert report.destination_sha256 == _sha256(dest)
    assert report.non_target_parts_preserved is True

    assert dest.is_file()
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    first_row = next(root.iter(Q(_W_NS, "tr")))

    # Verify w:trPr is the first child of w:tr
    assert first_row[0].tag == Q(_W_NS, "trPr")
    tr_pr = first_row[0]

    # Verify child sequence: cantSplit -> trHeight -> tblHeader
    child_tags = [c.tag.rsplit("}", 1)[-1] for c in tr_pr]
    assert child_tags == ["cantSplit", "trHeight", "tblHeader"]

    # Verify attributes
    tr_height = tr_pr.find(Q(_W_NS, "trHeight"))
    assert tr_height is not None
    assert tr_height.get(Q(_W_NS, "val")) == "720"
    assert tr_height.get(Q(_W_NS, "hRule")) == "atLeast"

    # Verify cell text preserved
    cells = first_row.findall(Q(_W_NS, "tc"))
    texts = ["".join(t.text or "" for t in c.iter(Q(_W_NS, "t"))) for c in cells]
    assert texts == ["Header A", "Header B"]

    # Verify second row untouched
    second_row = list(root.iter(Q(_W_NS, "tr")))[1]
    second_cells = second_row.findall(Q(_W_NS, "tc"))
    second_texts = ["".join(t.text or "" for t in c.iter(Q(_W_NS, "t"))) for c in second_cells]
    assert second_texts == ["Data 1", "Data 2"]


# --- 2. Flag Toggles (repeat_header, cant_split) ---


def test_author_table_row_properties_toggle_flags(tmp_path: Path) -> None:
    # Start with row that already has tblHeader and cantSplit
    initial_pr = "<w:cantSplit/><w:tblHeader/>"
    source = _make_row_doc(tmp_path / "initial.docx", row_properties=[initial_pr, ""])
    dest1 = tmp_path / "toggled_off.docx"
    source_sha = _sha256(source)

    fp1 = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp1 is not None

    # Toggle off both repeat_header and cant_split
    req1 = TableRowPropertiesRequest(
        source=source,
        destination=dest1,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp1,
        properties={"repeat_header": False, "cant_split": False},
    )
    rep1 = author_table_row_properties(req1)
    assert rep1.passed is True

    graph1 = PackageGraph.from_file(dest1)
    root1 = ET.fromstring(graph1.read_part_bytes(_DOCUMENT))
    tr_pr1 = next(root1.iter(Q(_W_NS, "tr"))).find(Q(_W_NS, "trPr"))
    assert tr_pr1 is not None
    tbl_h1 = tr_pr1.find(Q(_W_NS, "tblHeader"))
    assert tbl_h1 is not None and tbl_h1.get(Q(_W_NS, "val")) == "off"
    cant_s1 = tr_pr1.find(Q(_W_NS, "cantSplit"))
    assert cant_s1 is not None and cant_s1.get(Q(_W_NS, "val")) == "off"

    # Discovery readback correctly interprets explicit false w:val="off"
    disc1 = discover_table_row_properties_targets(dest1)
    assert disc1["targets"][0]["current_properties"]["repeat_header"] is False
    assert disc1["targets"][0]["current_properties"]["cant_split"] is False
    # Toggle repeat_header back on
    dest2 = tmp_path / "toggled_on.docx"
    dest1_sha = _sha256(dest1)
    fp2 = table_row_properties_structure_fingerprint(dest1, table=1, row=1)
    assert fp2 is not None

    req2 = TableRowPropertiesRequest(
        source=dest1,
        destination=dest2,
        source_sha256=dest1_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp2,
        properties={"repeat_header": True},
    )
    rep2 = author_table_row_properties(req2)
    assert rep2.passed is True

    graph2 = PackageGraph.from_file(dest2)
    root2 = ET.fromstring(graph2.read_part_bytes(_DOCUMENT))
    tr_pr2 = next(root2.iter(Q(_W_NS, "tr"))).find(Q(_W_NS, "trPr"))
    assert tr_pr2 is not None
    tbl_h2 = tr_pr2.find(Q(_W_NS, "tblHeader"))
    assert tbl_h2 is not None and tbl_h2.get(Q(_W_NS, "val")) is None
    # Omitted cant_split in req2 retains previous explicit false w:val="off"
    cant_s2 = tr_pr2.find(Q(_W_NS, "cantSplit"))
    assert cant_s2 is not None and cant_s2.get(Q(_W_NS, "val")) == "off"

    disc2 = discover_table_row_properties_targets(dest2)
    assert disc2["targets"][0]["current_properties"]["repeat_header"] is True
    assert disc2["targets"][0]["current_properties"]["cant_split"] is False


def test_author_table_row_properties_explicit_false_overrides_inherited(
    tmp_path: Path,
) -> None:
    # Row with inherited style setup; explicit false writes w:val="off" elements
    style_xml = (
        '<w:style w:type="table" w:styleId="HeaderTable">'
        '<w:name w:val="HeaderTable"/>'
        "<w:tblPr>"
        '<w:tblCellMar><w:top w:w="100" w:type="dxa"/></w:tblCellMar>'
        "</w:tblPr>"
        "</w:style>"
    )
    source = _make_row_doc(
        tmp_path / "inherited_src.docx",
        styles_extra=style_xml,
        tbl_properties='<w:tblStyle w:val="HeaderTable"/>',
        row_properties=["<w:tblHeader/><w:cantSplit/>", ""],
    )
    dest = tmp_path / "inherited_out.docx"
    source_sha = _sha256(source)
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    # Explicitly set false to override inherited behavior
    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": False, "cant_split": False},
    )
    report = author_table_row_properties(req)
    assert report.passed is True

    # XML has canonical w:val="off" elements in proper schema order
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tr_pr = next(root.iter(Q(_W_NS, "tr"))).find(Q(_W_NS, "trPr"))
    assert tr_pr is not None

    child_tags = [c.tag.rsplit("}", 1)[-1] for c in tr_pr]
    assert child_tags == ["cantSplit", "tblHeader"]

    cant_split = tr_pr.find(Q(_W_NS, "cantSplit"))
    assert cant_split is not None and cant_split.get(Q(_W_NS, "val")) == "off"
    tbl_header = tr_pr.find(Q(_W_NS, "tblHeader"))
    assert tbl_header is not None and tbl_header.get(Q(_W_NS, "val")) == "off"

    # Round-trip readback
    disc = discover_table_row_properties_targets(dest)
    assert disc["passed"] is True
    target_props = disc["targets"][0]["current_properties"]
    assert target_props["repeat_header"] is False
    assert target_props["cant_split"] is False
    assert report.target_change_report["properties_after"]["repeat_header"] is False
    assert report.target_change_report["properties_after"]["cant_split"] is False


# --- 3. Height Rules & Direct Integer Value ---


@pytest.mark.parametrize(
    ("height_input", "expected_val", "expected_rule"),
    [
        (RowHeightSpec(value=1200, rule="exact"), "1200", "exact"),
        ({"value": 400, "rule": "auto"}, "400", "auto"),
        ({"val": 850, "hRule": "atLeast"}, "850", "atLeast"),
        (600, "600", "atLeast"),
    ],
)
def test_author_table_row_properties_height_rules(
    tmp_path: Path, height_input: Any, expected_val: str, expected_rule: str
) -> None:
    source = _make_row_doc(tmp_path / f"source_{expected_rule}.docx")
    dest = tmp_path / f"output_{expected_rule}.docx"
    source_sha = _sha256(source)
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"height": height_input},
    )
    report = author_table_row_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tr_pr = next(root.iter(Q(_W_NS, "tr"))).find(Q(_W_NS, "trPr"))
    assert tr_pr is not None
    tr_height = tr_pr.find(Q(_W_NS, "trHeight"))
    assert tr_height is not None
    assert tr_height.get(Q(_W_NS, "val")) == expected_val
    assert tr_height.get(Q(_W_NS, "hRule")) == expected_rule


# --- 4. No-Op & Preservation of Unspecified Properties ---


def test_author_table_row_properties_preserves_unspecified_properties(
    tmp_path: Path,
) -> None:
    # Row has existing wBefore, tblHeader, and jc
    initial_pr = '<w:wBefore w:w="120" w:type="dxa"/><w:tblHeader/><w:jc w:val="center"/>'
    source = _make_row_doc(tmp_path / "source.docx", row_properties=[initial_pr, ""])
    dest = tmp_path / "preserved.docx"
    source_sha = _sha256(source)
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    # Only modify cant_split (repeat_header and height are unspecified/None)
    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties=TableRowPropertiesSpec(cant_split=True),
    )
    report = author_table_row_properties(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tr_pr = next(root.iter(Q(_W_NS, "tr"))).find(Q(_W_NS, "trPr"))
    assert tr_pr is not None

    # Verify wBefore, tblHeader, jc are preserved, and cantSplit added in schema order
    child_tags = [c.tag.rsplit("}", 1)[-1] for c in tr_pr]
    assert child_tags == ["wBefore", "cantSplit", "tblHeader", "jc"]

    w_before = tr_pr.find(Q(_W_NS, "wBefore"))
    assert w_before is not None and w_before.get(Q(_W_NS, "w")) == "120"
    assert tr_pr.find(Q(_W_NS, "tblHeader")) is not None
    jc = tr_pr.find(Q(_W_NS, "jc"))
    assert jc is not None and jc.get(Q(_W_NS, "val")) == "center"


def test_author_table_row_properties_no_op_empty_properties(tmp_path: Path) -> None:
    source = _make_row_doc(tmp_path / "noop_src.docx")
    dest = tmp_path / "noop_out.docx"
    source_sha = _sha256(source)
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    # Empty properties is valid no-op
    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={},
    )
    report = author_table_row_properties(req)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["properties_authored"] == []


# --- 5. Namespace Stability & Inspections ---


def test_author_table_row_properties_namespace_stability(tmp_path: Path) -> None:
    source = _make_row_doc(tmp_path / "ns_src.docx")
    dest = tmp_path / "ns_out.docx"
    source_sha = _sha256(source)
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    report = author_table_row_properties(req)
    assert report.passed is True

    insp = inspect_docx(dest)
    assert insp.valid is True


# --- 6. Strict Request Parsing & Validation ---


def test_parse_table_row_properties_request_valid() -> None:
    data = {
        "source": "in.docx",
        "destination": "out.docx",
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1, "row": 2, "part": _DOCUMENT},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {
            "repeat_header": True,
            "cant_split": False,
            "height": {"value": 720, "rule": "exact"},
        },
    }
    req = parse_table_row_properties_request(data)
    assert req.source == Path("in.docx")
    assert req.destination == Path("out.docx")
    assert req.table == 1
    assert req.row == 2
    assert req.part == _DOCUMENT
    assert req.properties["repeat_header"] is True
    assert req.properties["cant_split"] is False
    assert req.properties["height"] == {"value": 720, "rule": "exact"}


def test_parse_table_row_properties_request_invalid_types_and_fields() -> None:
    # Not mapping
    with pytest.raises(TypeError, match="table_row_properties_request_invalid"):
        parse_table_row_properties_request("not-a-dict")

    # Missing field
    with pytest.raises(ValueError, match="table_row_properties_request_fields_invalid"):
        parse_table_row_properties_request({"source": "a.docx"})

    # Unknown request field
    base = {
        "source": "a.docx",
        "destination": "b.docx",
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1, "row": 1},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {"repeat_header": True},
    }
    with pytest.raises(ValueError, match="table_row_properties_request_fields_invalid"):
        parse_table_row_properties_request({**base, "extra_unexpected": 123})

    # Empty source / destination
    with pytest.raises(ValueError, match="table_row_properties_source_invalid"):
        parse_table_row_properties_request({**base, "source": ""})
    with pytest.raises(ValueError, match="table_row_properties_destination_invalid"):
        parse_table_row_properties_request({**base, "destination": ""})

    # Invalid sha256
    with pytest.raises(ValueError, match="table_row_properties_source_sha256_invalid"):
        parse_table_row_properties_request({**base, "source_sha256": "tooshort"})
    with pytest.raises(
        ValueError, match="table_row_properties_structure_fingerprint_sha256_invalid"
    ):
        parse_table_row_properties_request({**base, "structure_fingerprint_sha256": "bad"})

    # Invalid target
    with pytest.raises(TypeError, match="table_row_properties_target_invalid"):
        parse_table_row_properties_request({**base, "target": "not-dict"})
    with pytest.raises(ValueError, match="table_row_properties_target_fields_invalid"):
        parse_table_row_properties_request({**base, "target": {"table": 1}})  # missing row
    with pytest.raises(ValueError, match="table_row_properties_table_invalid"):
        parse_table_row_properties_request({**base, "target": {"table": 0, "row": 1}})
    with pytest.raises(ValueError, match="table_row_properties_table_invalid"):
        parse_table_row_properties_request(
            {**base, "target": {"table": True, "row": 1}}
        )  # bool as number
    with pytest.raises(ValueError, match="table_row_properties_row_invalid"):
        parse_table_row_properties_request({**base, "target": {"table": 1, "row": -1}})
    with pytest.raises(ValueError, match="table_row_properties_row_invalid"):
        parse_table_row_properties_request(
            {**base, "target": {"table": 1, "row": False}}
        )  # bool as number
    with pytest.raises(ValueError, match="table_row_properties_part_invalid"):
        parse_table_row_properties_request({**base, "target": {"table": 1, "row": 1, "part": ""}})

    # Unknown property
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request({**base, "properties": {"unknown_prop": 1}})

    # Invalid bool-as-number in properties
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request({**base, "properties": {"repeat_header": "yes"}})
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request({**base, "properties": {"cant_split": 1}})

    # Invalid height bounds and rules
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request({**base, "properties": {"height": -1}})
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request({**base, "properties": {"height": 31681}})
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request(
            {**base, "properties": {"height": True}}
        )  # bool as height
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request(
            {**base, "properties": {"height": {"value": 500, "rule": "unsupported"}}}
        )
    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        parse_table_row_properties_request(
            {**base, "properties": {"height": {"value": True, "rule": "exact"}}}
        )


def test_table_row_properties_request_direct_dataclass_validation() -> None:
    with pytest.raises(TypeError, match="table_row_properties_source_invalid"):
        TableRowPropertiesRequest(
            source=123,  # type: ignore[arg-type]
            destination=Path("out.docx"),
            source_sha256=_DUMMY_SHA,
            target={"table": 1, "row": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            properties={"repeat_header": True},
        )

    with pytest.raises(TypeError, match="table_row_properties_fields_invalid"):
        RowHeightSpec(value="not-an-int")  # type: ignore[arg-type]

    with pytest.raises(ValueError, match="table_row_properties_fields_invalid"):
        RowHeightSpec(value=99999)  # exceeds 31680


def test_property_spec_dataclasses_round_trip() -> None:
    h = RowHeightSpec(value=720, rule="exact")
    assert h.to_dict() == {"value": 720, "rule": "exact"}

    spec = TableRowPropertiesSpec(repeat_header=True, height=h, cant_split=False)
    assert spec.to_dict() == {
        "repeat_header": True,
        "height": {"value": 720, "rule": "exact"},
        "cant_split": False,
    }

    tgt = TableRowPropertiesTarget(
        table=1,
        row=2,
        structure_fingerprint_sha256=_DUMMY_FP,
        cells_count=3,
        current_properties={"repeat_header": True},
    )
    assert tgt.target == {"table": 1, "row": 2, "part": _DOCUMENT}
    tgt_dict = tgt.to_dict()
    assert tgt_dict["table"] == 1
    assert tgt_dict["row"] == 2
    assert tgt_dict["cells_count"] == 3

    rep = TableRowPropertiesReport(
        status="published",
        code=None,
        source_sha256=_DUMMY_SHA,
        destination_sha256=_DUMMY_SHA,
        target={"table": 1, "row": 1, "part": _DOCUMENT},
        properties={"repeat_header": True},
    )
    assert rep.passed is True
    assert rep.to_dict()["operation"] == "table-row-properties"


# --- 7. Target Discovery ---


def test_discover_table_row_properties_targets(tmp_path: Path) -> None:
    doc = _make_row_doc(
        tmp_path / "multi.docx",
        rows_data=[["H1", "H2"], ["D1", "D2"], ["D3", "D4"]],
        row_properties=["<w:tblHeader/>", "", "<w:cantSplit/>"],
    )
    disc = discover_table_row_properties_targets(doc)
    assert disc["passed"] is True
    assert disc["operation"] == "table-row-properties-targets"
    assert disc["status"] == "discovered"
    assert disc["code"] is None
    assert disc["source_sha256"] == _sha256(doc)
    assert disc["part"] == _DOCUMENT
    assert len(disc["targets"]) == 3

    t1, t2, t3 = disc["targets"]
    assert t1["table"] == 1 and t1["row"] == 1
    assert t1["cells_count"] == 2
    assert t1["current_properties"]["repeat_header"] is True

    assert t2["table"] == 1 and t2["row"] == 2
    assert t2["current_properties"] == {}

    assert t3["table"] == 1 and t3["row"] == 3
    assert t3["current_properties"]["cant_split"] is True


def test_discover_table_row_properties_targets_refusals(tmp_path: Path) -> None:
    # Missing file
    d1 = discover_table_row_properties_targets(tmp_path / "missing.docx")
    assert d1["passed"] is False
    assert d1["code"] == "table_row_properties_source_missing"

    # Unsupported part
    doc = _make_row_doc(tmp_path / "ok.docx")
    d2 = discover_table_row_properties_targets(doc, part="word/header1.xml")
    assert d2["passed"] is False
    assert d2["code"] == "table_row_properties_part_unsupported"


def test_load_table_row_properties_request(tmp_path: Path) -> None:
    req_file = tmp_path / "request.json"
    req_data = {
        "source": "in.docx",
        "destination": "out.docx",
        "source_sha256": _DUMMY_SHA,
        "target": {"table": 1, "row": 1, "part": _DOCUMENT},
        "structure_fingerprint_sha256": _DUMMY_FP,
        "properties": {"repeat_header": True},
    }
    req_file.write_text(json.dumps(req_data), encoding="utf-8")
    loaded = load_table_row_properties_request(req_file)
    assert loaded.source == Path("in.docx")
    assert loaded.properties["repeat_header"] is True


# --- 8. Authoring Refusals ---


def test_author_table_row_properties_refusal_source_missing(tmp_path: Path) -> None:
    req = TableRowPropertiesRequest(
        source=tmp_path / "absent.docx",
        destination=tmp_path / "out.docx",
        source_sha256=_DUMMY_SHA,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=_DUMMY_FP,
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is False
    assert rep.code == "table_row_properties_source_missing"


def test_author_table_row_properties_refusal_destination_occupied(
    tmp_path: Path,
) -> None:
    source = _make_row_doc(tmp_path / "source.docx")
    dest = tmp_path / "occupied.docx"
    dest.write_text("existing content", encoding="utf-8")
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is False
    assert rep.code == "table_row_properties_destination_occupied"
    # Destination content must be untouched
    assert dest.read_text(encoding="utf-8") == "existing content"


def test_author_table_row_properties_refusal_source_hash_mismatch(
    tmp_path: Path,
) -> None:
    source = _make_row_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_DUMMY_SHA,  # Stale hash
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is False
    assert rep.code == "table_row_properties_source_hash_mismatch"


def test_author_table_row_properties_refusal_fingerprint_mismatch(
    tmp_path: Path,
) -> None:
    source = _make_row_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=_DUMMY_FP,  # Stale fingerprint
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is False
    assert rep.code == "table_row_properties_fingerprint_mismatch"


def test_author_table_row_properties_refusal_target_missing(tmp_path: Path) -> None:
    source = _make_row_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    # Target table out of range
    req_tbl = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 99, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep_tbl = author_table_row_properties(req_tbl)
    assert rep_tbl.passed is False
    assert rep_tbl.code == "table_row_properties_target_missing"

    # Target row out of range
    req_row = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 99},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep_row = author_table_row_properties(req_row)
    assert rep_row.passed is False
    assert rep_row.code == "table_row_properties_target_missing"


def test_author_table_row_properties_refusal_part_unsupported(tmp_path: Path) -> None:
    source = _make_row_doc(tmp_path / "source.docx")
    dest = tmp_path / "out.docx"
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1, "part": "word/footer1.xml"},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is False
    assert rep.code == "table_row_properties_part_unsupported"


def test_author_table_row_properties_refusal_context_unsupported(
    tmp_path: Path,
) -> None:
    # 1. Table inside tracked revision <w:ins>
    source_ins = _make_row_doc(tmp_path / "ins.docx", wrapper="ins")
    dest_ins = tmp_path / "ins_out.docx"
    fp_ins = table_row_properties_structure_fingerprint(source_ins, table=1, row=1)
    assert fp_ins is not None

    req_ins = TableRowPropertiesRequest(
        source=source_ins,
        destination=dest_ins,
        source_sha256=_sha256(source_ins),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp_ins,
        properties={"repeat_header": True},
    )
    rep_ins = author_table_row_properties(req_ins)
    assert rep_ins.passed is False
    assert rep_ins.code == "table_row_properties_context_unsupported"

    # 2. Row with gridBefore
    source_gb = _make_row_doc(
        tmp_path / "gb.docx", row_properties=['<w:gridBefore w:val="1"/>', ""]
    )
    dest_gb = tmp_path / "gb_out.docx"
    fp_gb = table_row_properties_structure_fingerprint(source_gb, table=1, row=1)
    assert fp_gb is not None

    req_gb = TableRowPropertiesRequest(
        source=source_gb,
        destination=dest_gb,
        source_sha256=_sha256(source_gb),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp_gb,
        properties={"repeat_header": True},
    )
    rep_gb = author_table_row_properties(req_gb)
    assert rep_gb.passed is False
    assert rep_gb.code == "table_row_properties_context_unsupported"


def test_author_table_row_properties_refusal_conflict(tmp_path: Path) -> None:
    # 1. Row has duplicate w:trHeight elements
    dup_pr = '<w:trHeight w:val="400"/><w:trHeight w:val="600"/>'
    source = _make_row_doc(tmp_path / "conflict.docx", row_properties=[dup_pr, ""])
    dest = tmp_path / "conflict_out.docx"
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is False
    assert rep.code == "table_row_properties_conflict"


# --- 9. Conditional Table Styles ---


def test_author_table_row_properties_with_conditional_table_style(tmp_path: Path) -> None:
    # Table style in styles.xml has tblStylePr (conditional formatting)
    style_xml = (
        '<w:style w:type="table" w:styleId="ConditionalStyle">'
        '<w:tblStylePr w:type="firstRow"/>'
        "</w:style>"
    )
    source = _make_row_doc(
        tmp_path / "cond.docx",
        tbl_properties='<w:tblStyle w:val="ConditionalStyle"/>',
        styles_extra=style_xml,
    )
    dest = tmp_path / "cond_out.docx"
    source_sha = _sha256(source)
    fp = table_row_properties_structure_fingerprint(source, table=1, row=1)
    assert fp is not None

    disc = discover_table_row_properties_targets(source)
    assert disc["passed"] is True

    # Authoring still succeeds cleanly!
    req = TableRowPropertiesRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=fp,
        properties={"repeat_header": True},
    )
    rep = author_table_row_properties(req)
    assert rep.passed is True
    assert rep.status == "published"
