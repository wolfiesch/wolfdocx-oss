"""Tests for source-bound table cell merge and split authoring and discovery."""

from __future__ import annotations

import hashlib
import json
from functools import partial
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
from wolfdocx.inspect import Q, inspect_docx
from wolfdocx.package import PackageGraph
from wolfdocx.table_merge import (
    TableMergeReport,
    TableMergeRequest,
    TableMergeTarget,
    author_table_merge,
    discover_table_merge_targets,
    load_table_merge_request,
    parse_table_merge_request,
    table_merge_structure_fingerprint,
)

_DOCUMENT = "word/document.xml"
_DUMMY_SHA = "a" * 64
_DUMMY_FP = "b" * 64


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_table_doc(
    path: Path,
    *,
    rows_data: list[list[str]],
    grid_widths: list[int] | None = None,
    wrapper: str = "",
) -> Path:
    """Create a minimal valid DOCX with an explicit grid table."""
    cols = len(rows_data[0])
    widths = grid_widths if grid_widths is not None else [2500] * cols
    table_grid = grid(widths)

    rows_xml: list[str] = []
    for r_data in rows_data:
        cells_xml = "".join(cell(paragraph(txt)) for txt in r_data)
        rows_xml.append(row(cells_xml))

    tbl_xml = table("".join(rows_xml), grid=table_grid)
    if wrapper == "sdt":
        body = f"<w:sdt><w:sdtContent>{tbl_xml}</w:sdtContent></w:sdt>"
    elif wrapper == "ins":
        body = f'<w:ins w:id="1" w:author="tester">{tbl_xml}</w:ins>'
    else:
        body = tbl_xml
    return write_package(path, body)


# --- Request Parsing & Validation Tests ---


def test_parse_table_merge_request_valid() -> None:
    req_dict = {
        "source": "input.docx",
        "destination": "output.docx",
        "source_sha256": _DUMMY_SHA,
        "structure_fingerprint_sha256": _DUMMY_FP,
        "action": "merge",
        "target": {
            "table": 1,
            "start_row": 1,
            "end_row": 1,
            "start_col": 1,
            "end_col": 2,
            "part": _DOCUMENT,
        },
    }
    req = parse_table_merge_request(req_dict)
    assert req.source == Path("input.docx")
    assert req.destination == Path("output.docx")
    assert req.source_sha256 == _DUMMY_SHA
    assert req.structure_fingerprint_sha256 == _DUMMY_FP
    assert req.action == "merge"
    assert req.table == 1
    assert req.start_row == 1
    assert req.end_row == 1
    assert req.start_col == 1
    assert req.end_col == 2
    assert req.part == _DOCUMENT

    # Round-trip to_dict
    dumped = req.to_dict()
    assert dumped["action"] == "merge"
    assert dumped["target"]["start_row"] == 1


def test_parse_table_merge_request_default_action() -> None:
    req_dict = {
        "source": "input.docx",
        "destination": "output.docx",
        "source_sha256": _DUMMY_SHA,
        "structure_fingerprint_sha256": _DUMMY_FP,
        "target": {
            "table": 1,
            "start_row": 1,
            "end_row": 2,
            "start_col": 1,
            "end_col": 1,
        },
    }
    req = parse_table_merge_request(req_dict)
    assert req.action == "merge"
    assert req.part == _DOCUMENT


def test_parse_table_merge_request_refusals() -> None:
    base = {
        "source": "input.docx",
        "destination": "output.docx",
        "source_sha256": _DUMMY_SHA,
        "structure_fingerprint_sha256": _DUMMY_FP,
        "action": "merge",
        "target": {
            "table": 1,
            "start_row": 1,
            "end_row": 1,
            "start_col": 1,
            "end_col": 2,
        },
    }

    # Non-mapping
    with pytest.raises(TypeError, match="table_merge_request_invalid"):
        parse_table_merge_request("not a mapping")

    # Extra fields
    with pytest.raises(ValueError, match="table_merge_request_fields_invalid"):
        parse_table_merge_request({**base, "unknown_field": True})

    # Missing fields
    with pytest.raises(ValueError, match="table_merge_request_fields_invalid"):
        parse_table_merge_request({k: v for k, v in base.items() if k != "source"})

    # Invalid source
    with pytest.raises(ValueError, match="table_merge_source_invalid"):
        parse_table_merge_request({**base, "source": ""})

    # Invalid destination
    with pytest.raises(ValueError, match="table_merge_destination_invalid"):
        parse_table_merge_request({**base, "destination": 123})

    # Invalid sha256
    with pytest.raises(ValueError, match="table_merge_source_sha256"):
        parse_table_merge_request({**base, "source_sha256": "bad"})

    # Invalid action
    with pytest.raises(ValueError, match="table_merge_action_invalid"):
        parse_table_merge_request({**base, "action": "rotate"})

    # Invalid target
    with pytest.raises(TypeError, match="table_merge_target_invalid"):
        parse_table_merge_request({**base, "target": "bad"})

    with pytest.raises(ValueError, match="table_merge_target_fields_invalid"):
        parse_table_merge_request({**base, "target": {"table": 1}})

    # Coordinates out of order (start_row > end_row)
    with pytest.raises(ValueError, match="table_merge_coordinates_invalid"):
        parse_table_merge_request(
            {
                **base,
                "target": {
                    "table": 1,
                    "start_row": 3,
                    "end_row": 1,
                    "start_col": 1,
                    "end_col": 2,
                },
            }
        )

    # Coordinates out of order (start_col > end_col)
    with pytest.raises(ValueError, match="table_merge_coordinates_invalid"):
        parse_table_merge_request(
            {
                **base,
                "target": {
                    "table": 1,
                    "start_row": 1,
                    "end_row": 1,
                    "start_col": 3,
                    "end_col": 2,
                },
            }
        )

    # Boolean coordinate rejected
    with pytest.raises(ValueError, match="table_merge_table_invalid"):
        parse_table_merge_request(
            {
                **base,
                "target": {
                    "table": True,
                    "start_row": 1,
                    "end_row": 1,
                    "start_col": 1,
                    "end_col": 2,
                },
            }
        )


def test_load_table_merge_request(tmp_path: Path) -> None:
    req_file = tmp_path / "req.json"
    req_data = {
        "source": "src.docx",
        "destination": "dst.docx",
        "source_sha256": _DUMMY_SHA,
        "structure_fingerprint_sha256": _DUMMY_FP,
        "action": "split",
        "target": {
            "table": 1,
            "start_row": 1,
            "end_row": 2,
            "start_col": 1,
            "end_col": 1,
        },
    }
    req_file.write_text(json.dumps(req_data), encoding="utf-8")
    loaded = load_table_merge_request(req_file)
    assert loaded.action == "split"
    assert loaded.end_row == 2


# --- Horizontal Merge & Split Lifecycle ---


def test_horizontal_merge_and_split_lifecycle(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "table_h.docx",
        rows_data=[
            ["Origin Text", "", "Col C"],
            ["Row 2 A", "Row 2 B", "Row 2 C"],
            ["Row 3 A", "Row 3 B", "Row 3 C"],
        ],
        grid_widths=[2000, 3000, 4000],
    )
    src_sha = _sha256(src_doc)
    merged_doc = tmp_path / "table_h_merged.docx"

    # 1. Structure fingerprint
    fp = table_merge_structure_fingerprint(
        src_doc,
        table=1,
        start_row=1,
        end_row=1,
        start_col=1,
        end_col=2,
    )
    assert fp is not None

    # 2. Author horizontal merge (cols 1..2 in row 1)
    req = TableMergeRequest(
        source=src_doc,
        destination=merged_doc,
        source_sha256=src_sha,
        target={
            "table": 1,
            "start_row": 1,
            "end_row": 1,
            "start_col": 1,
            "end_col": 2,
        },
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    report = author_table_merge(req)
    assert isinstance(report, TableMergeReport)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True

    # 3. Independent OOXML inspection of merged document
    insp = inspect_docx(merged_doc)
    assert insp.valid is True

    graph = PackageGraph.from_file(merged_doc)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = _DOCUMENT_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)

    table_elem = next(root.iter(word("tbl")))
    rows = table_elem.findall(word("tr"))
    assert len(rows) == 3

    # Row 1 has 2 cells: origin cell (spanning cols 1-2) and col 3
    r1_cells = rows[0].findall(word("tc"))
    assert len(r1_cells) == 2

    origin = r1_cells[0]
    origin_tc_pr = origin.find(word("tcPr"))
    assert origin_tc_pr is not None

    # gridSpan is 2
    grid_span_elem = origin_tc_pr.find(word("gridSpan"))
    assert grid_span_elem is not None
    assert grid_span_elem.get(word("val")) == "2"

    # Width is 2000 + 3000 = 5000
    tc_w_elem = origin_tc_pr.find(word("tcW"))
    assert tc_w_elem is not None
    assert tc_w_elem.get(word("w")) == "5000"

    # Origin text is preserved
    t_nodes = origin.findall(f".//{word('t')}")
    assert len(t_nodes) == 1
    assert t_nodes[0].text == "Origin Text"

    # Row 2 and 3 cells unchanged (each still has 3 cells)
    assert len(rows[1].findall(word("tc"))) == 3
    assert len(rows[2].findall(word("tc"))) == 3

    # tblGrid unchanged
    tbl_grid = table_elem.find(word("tblGrid"))
    assert tbl_grid is not None
    assert len(tbl_grid.findall(word("gridCol"))) == 3

    # 4. Discovery on merged document discovers the existing merge
    discovery = discover_table_merge_targets(merged_doc)
    assert discovery["passed"] is True
    assert discovery["status"] == "discovered"
    targets = discovery["targets"]
    split_targets = [t for t in targets if t["action"] == "split"]
    assert len(split_targets) == 1
    target = split_targets[0]
    target_obj = TableMergeTarget(
        table=target["table"],
        start_row=target["start_row"],
        end_row=target["end_row"],
        start_col=target["start_col"],
        end_col=target["end_col"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        action=target["action"],
    )
    assert target_obj.to_dict()["action"] == "split"
    assert target["start_row"] == 1
    assert target["is_merged"] is True
    assert target["origin_text"] == "Origin Text"

    # 5. Author split to reverse the merge
    split_doc = tmp_path / "table_h_split.docx"
    merged_sha = _sha256(merged_doc)
    split_req = TableMergeRequest(
        source=merged_doc,
        destination=split_doc,
        source_sha256=merged_sha,
        target={
            "table": target["table"],
            "start_row": target["start_row"],
            "end_row": target["end_row"],
            "start_col": target["start_col"],
            "end_col": target["end_col"],
        },
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        action="split",
    )
    split_report = author_table_merge(split_req)
    assert split_report.passed is True
    assert split_report.status == "published"

    # 6. Verify split document structure
    split_graph = PackageGraph.from_file(split_doc)
    split_root = ET.fromstring(split_graph.read_part_bytes(_DOCUMENT))
    split_tbl = next(split_root.iter(word("tbl")))
    split_r1_cells = split_tbl.findall(word("tr"))[0].findall(word("tc"))
    assert len(split_r1_cells) == 3

    # Cell 1 has origin text and width 2000, gridSpan removed
    c1_tc_pr = split_r1_cells[0].find(word("tcPr"))
    assert c1_tc_pr.find(word("gridSpan")) is None
    assert c1_tc_pr.find(word("tcW")).get(word("w")) == "2000"
    assert split_r1_cells[0].find(f".//{word('t')}").text == "Origin Text"

    # Cell 2 has width 3000, gridSpan removed, empty paragraph
    c2_tc_pr = split_r1_cells[1].find(word("tcPr"))
    assert c2_tc_pr.find(word("gridSpan")) is None
    assert c2_tc_pr.find(word("tcW")).get(word("w")) == "3000"
    assert len(split_r1_cells[1].findall(word("p"))) >= 1


# --- Vertical Merge & Split Lifecycle ---


def test_vertical_merge_and_split_lifecycle(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "table_v.docx",
        rows_data=[
            ["Top Origin", "B1", "C1"],
            ["", "B2", "C2"],
            ["A3", "B3", "C3"],
        ],
        grid_widths=[2500, 2500, 2500],
    )
    src_sha = _sha256(src_doc)
    merged_doc = tmp_path / "table_v_merged.docx"

    fp = table_merge_structure_fingerprint(
        src_doc,
        table=1,
        start_row=1,
        end_row=2,
        start_col=1,
        end_col=1,
    )
    assert fp is not None

    # 1. Author vertical merge (rows 1..2 in col 1)
    req = TableMergeRequest(
        source=src_doc,
        destination=merged_doc,
        source_sha256=src_sha,
        target={
            "table": 1,
            "start_row": 1,
            "end_row": 2,
            "start_col": 1,
            "end_col": 1,
        },
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    report = author_table_merge(req)
    assert report.passed is True

    # 2. Inspect merged doc
    graph = PackageGraph.from_file(merged_doc)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)

    table_elem = next(root.iter(word("tbl")))
    rows = table_elem.findall(word("tr"))

    # Origin cell (row 1, col 1) has vMerge="restart"
    r1_c1 = rows[0].findall(word("tc"))[0]
    vm1 = r1_c1.find(word("tcPr")).find(word("vMerge"))
    assert vm1 is not None
    assert vm1.get(word("val")) == "restart"
    assert r1_c1.find(f".//{word('t')}").text == "Top Origin"

    # Continuation cell (row 2, col 1) has vMerge
    r2_c1 = rows[1].findall(word("tc"))[0]
    vm2 = r2_c1.find(word("tcPr")).find(word("vMerge"))
    assert vm2 is not None

    # 3. Discovery on merged doc
    discovery = discover_table_merge_targets(merged_doc)
    assert discovery["passed"] is True
    split_targets = [t for t in discovery["targets"] if t["action"] == "split"]
    assert len(split_targets) == 1
    t = split_targets[0]
    assert t["start_row"] == 1
    assert t["end_row"] == 2
    assert t["start_col"] == 1
    assert t["end_col"] == 1
    assert t["origin_text"] == "Top Origin"

    # 4. Author split to reverse vertical merge
    split_doc = tmp_path / "table_v_split.docx"
    split_req = TableMergeRequest(
        source=merged_doc,
        destination=split_doc,
        source_sha256=_sha256(merged_doc),
        target={
            "table": t["table"],
            "start_row": t["start_row"],
            "end_row": t["end_row"],
            "start_col": t["start_col"],
            "end_col": t["end_col"],
        },
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
        action="split",
    )
    split_report = author_table_merge(split_req)
    assert split_report.passed is True

    # 5. Verify split doc
    s_graph = PackageGraph.from_file(split_doc)
    s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
    s_rows = next(s_root.iter(word("tbl"))).findall(word("tr"))

    # vMerge removed from both cells
    assert s_rows[0].findall(word("tc"))[0].find(word("tcPr")).find(word("vMerge")) is None
    assert s_rows[1].findall(word("tc"))[0].find(word("tcPr")).find(word("vMerge")) is None
    assert s_rows[0].findall(word("tc"))[0].find(f".//{word('t')}").text == "Top Origin"


# --- Combined Rectangular Merge & Split Lifecycle ---


def test_combined_rectangular_merge_and_split_lifecycle(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "table_rect.docx",
        rows_data=[
            ["Block Origin", "", "Right 1"],
            ["", "", "Right 2"],
            ["Bottom 1", "Bottom 2", "Right 3"],
        ],
        grid_widths=[2000, 3000, 4000],
    )
    src_sha = _sha256(src_doc)
    merged_doc = tmp_path / "table_rect_merged.docx"

    fp = table_merge_structure_fingerprint(
        src_doc,
        table=1,
        start_row=1,
        end_row=2,
        start_col=1,
        end_col=2,
    )
    assert fp is not None

    # Merge 2x2 block
    req = TableMergeRequest(
        source=src_doc,
        destination=merged_doc,
        source_sha256=src_sha,
        target={
            "table": 1,
            "start_row": 1,
            "end_row": 2,
            "start_col": 1,
            "end_col": 2,
        },
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    report = author_table_merge(req)
    assert report.passed is True

    # Inspect merged OOXML
    graph = PackageGraph.from_file(merged_doc)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)

    table_elem = next(root.iter(word("tbl")))
    rows = table_elem.findall(word("tr"))

    # Row 1 and Row 2 each now have 2 cells
    assert len(rows[0].findall(word("tc"))) == 2
    assert len(rows[1].findall(word("tc"))) == 2

    # Origin cell: gridSpan=2, vMerge=restart, width 5000
    r1_c1 = rows[0].findall(word("tc"))[0]
    r1_pr = r1_c1.find(word("tcPr"))
    assert r1_pr.find(word("gridSpan")).get(word("val")) == "2"
    assert r1_pr.find(word("vMerge")).get(word("val")) == "restart"
    assert r1_pr.find(word("tcW")).get(word("w")) == "5000"

    # Continuation cell: gridSpan=2, vMerge=continue, width 5000
    r2_c1 = rows[1].findall(word("tc"))[0]
    r2_pr = r2_c1.find(word("tcPr"))
    assert r2_pr.find(word("gridSpan")).get(word("val")) == "2"
    assert r2_pr.find(word("vMerge")) is not None
    assert r2_pr.find(word("tcW")).get(word("w")) == "5000"

    # Discover and split
    discovery = discover_table_merge_targets(merged_doc)
    t = next(t for t in discovery["targets"] if t["action"] == "split")
    assert t["start_row"] == 1
    assert t["end_row"] == 2
    assert t["start_col"] == 1
    assert t["end_col"] == 2

    split_doc = tmp_path / "table_rect_split.docx"
    split_req = TableMergeRequest(
        source=merged_doc,
        destination=split_doc,
        source_sha256=_sha256(merged_doc),
        target={
            "table": t["table"],
            "start_row": t["start_row"],
            "end_row": t["end_row"],
            "start_col": t["start_col"],
            "end_col": t["end_col"],
        },
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
        action="split",
    )
    split_rep = author_table_merge(split_req)
    assert split_rep.passed is True

    # Split restored 3 cells per row
    s_graph = PackageGraph.from_file(split_doc)
    s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
    s_rows = next(s_root.iter(word("tbl"))).findall(word("tr"))
    assert len(s_rows[0].findall(word("tc"))) == 3
    assert len(s_rows[1].findall(word("tc"))) == 3


# --- Refusal Scenarios ---


def test_author_table_merge_refusal_content_loss(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "loss.docx",
        rows_data=[
            ["Origin Text", "Non-empty Data", "Col C"],
            ["Row 2 A", "Row 2 B", "Row 2 C"],
        ],
    )
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "loss_out.docx",
        source_sha256=_sha256(src_doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "table_merge_content_loss_unsupported"
    assert not (tmp_path / "loss_out.docx").exists()


def test_author_table_merge_refusal_stale_source_hash(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "stale_src.docx",
        rows_data=[["A", "", "C"], ["D", "E", "F"]],
    )
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "stale_out.docx",
        source_sha256=_DUMMY_SHA,
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_source_hash_mismatch"


def test_author_table_merge_refusal_stale_fingerprint(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "stale_fp.docx",
        rows_data=[["A", "", "C"], ["D", "E", "F"]],
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "stale_fp_out.docx",
        source_sha256=_sha256(src_doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_structure_fingerprint_mismatch"


def test_author_table_merge_refusal_destination_occupied(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "occupied_src.docx",
        rows_data=[["A", "", "C"], ["D", "E", "F"]],
    )
    dest = tmp_path / "occupied_dest.docx"
    dest.write_text("occupied")
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=dest,
        source_sha256=_sha256(src_doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_destination_occupied"
    assert dest.read_text() == "occupied"


def test_author_table_merge_refusal_source_missing(tmp_path: Path) -> None:
    req = TableMergeRequest(
        source=tmp_path / "missing.docx",
        destination=tmp_path / "out.docx",
        source_sha256=_DUMMY_SHA,
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_source_missing"


def test_author_table_merge_refusal_target_missing(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "target_missing.docx",
        rows_data=[["A", "", "C"], ["D", "E", "F"]],
    )
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "out.docx",
        source_sha256=_sha256(src_doc),
        target={"table": 5, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_target_missing"


def test_author_table_merge_refusal_unsupported_context(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "sdt_table.docx",
        rows_data=[["A", "", "C"], ["D", "E", "F"]],
        wrapper="sdt",
    )
    # Fingerprint returns None for unsupported context table
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "out.docx",
        source_sha256=_sha256(src_doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp or _DUMMY_FP,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code in (
        "table_merge_context_unsupported",
        "table_merge_structure_fingerprint_mismatch",
    )


def test_author_table_merge_refusal_single_cell_merge(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "single_cell.docx",
        rows_data=[["A", "B", "C"], ["D", "E", "F"]],
    )
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=1
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "out.docx",
        source_sha256=_sha256(src_doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 1},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_geometry_unsupported"


def test_author_table_merge_refusal_split_unmerged_cell(tmp_path: Path) -> None:
    src_doc = _make_table_doc(
        tmp_path / "split_unmerged.docx",
        rows_data=[["A", "B", "C"], ["D", "E", "F"]],
    )
    fp = table_merge_structure_fingerprint(
        src_doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=src_doc,
        destination=tmp_path / "out.docx",
        source_sha256=_sha256(src_doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="split",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_geometry_unsupported"


def test_discover_table_merge_targets_source_missing() -> None:
    discovery = discover_table_merge_targets("nonexistent.docx")
    assert discovery["passed"] is False
    assert discovery["status"] == "refused"
    assert discovery["code"] == "table_merge_source_missing"
    assert discovery["targets"] == []


def test_table_merge_request_direct_admission_refusals() -> None:
    valid_target = {
        "table": 1,
        "start_row": 1,
        "end_row": 1,
        "start_col": 1,
        "end_col": 2,
    }

    # Direct construction with invalid action
    with pytest.raises(ValueError, match="table_merge_action_invalid"):
        TableMergeRequest(
            source=Path("s.docx"),
            destination=Path("d.docx"),
            source_sha256=_DUMMY_SHA,
            target=valid_target,
            structure_fingerprint_sha256=_DUMMY_FP,
            action="dance",
        )

    # Direct construction with boolean coordinate
    with pytest.raises(ValueError, match="table_merge_table_invalid"):
        TableMergeRequest(
            source=Path("s.docx"),
            destination=Path("d.docx"),
            source_sha256=_DUMMY_SHA,
            target={**valid_target, "table": True},
            structure_fingerprint_sha256=_DUMMY_FP,
            action="merge",
        )

    # Direct construction with descending row range
    with pytest.raises(ValueError, match="table_merge_coordinates_invalid"):
        TableMergeRequest(
            source=Path("s.docx"),
            destination=Path("d.docx"),
            source_sha256=_DUMMY_SHA,
            target={**valid_target, "start_row": 3, "end_row": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            action="merge",
        )

    # Direct construction with descending col range
    with pytest.raises(ValueError, match="table_merge_coordinates_invalid"):
        TableMergeRequest(
            source=Path("s.docx"),
            destination=Path("d.docx"),
            source_sha256=_DUMMY_SHA,
            target={**valid_target, "start_col": 3, "end_col": 1},
            structure_fingerprint_sha256=_DUMMY_FP,
            action="merge",
        )

    # Mutable dict passed into target is defensively copied and cannot be mutated externally
    mutable_dict = dict(valid_target)
    req = TableMergeRequest(
        source=Path("s.docx"),
        destination=Path("d.docx"),
        source_sha256=_DUMMY_SHA,
        target=mutable_dict,
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    mutable_dict["start_row"] = 99
    assert req.start_row == 1


def test_author_table_merge_refusal_whitespace_and_control_payload(
    tmp_path: Path,
) -> None:
    # Non-origin cell containing whitespace-only text
    c1 = cell(paragraph("Origin"))
    c2_ws = '<w:tc><w:p><w:r><w:t xml:space="preserve">   </w:t></w:r></w:p></w:tc>'
    c3 = cell(paragraph("Col C"))
    row_xml = f"<w:tr>{c1}{c2_ws}{c3}</w:tr>"
    doc_ws = tmp_path / "ws.docx"
    write_package(doc_ws, f"<w:tbl>{grid([2500, 2500, 2500])}{row_xml}</w:tbl>")

    fp_ws = table_merge_structure_fingerprint(
        doc_ws, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req_ws = TableMergeRequest(
        source=doc_ws,
        destination=tmp_path / "out_ws.docx",
        source_sha256=_sha256(doc_ws),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp_ws,
        action="merge",
    )
    rep_ws = author_table_merge(req_ws)
    assert rep_ws.passed is False
    assert rep_ws.code == "table_merge_content_loss_unsupported"

    # Non-origin cell containing <w:tab/>
    c2_tab = "<w:tc><w:p><w:r><w:tab/></w:r></w:p></w:tc>"
    doc_tab = tmp_path / "tab.docx"
    write_package(
        doc_tab,
        f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c1}{c2_tab}{c3}</w:tr></w:tbl>",
    )
    fp_tab = table_merge_structure_fingerprint(
        doc_tab, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req_tab = TableMergeRequest(
        source=doc_tab,
        destination=tmp_path / "out_tab.docx",
        source_sha256=_sha256(doc_tab),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp_tab,
        action="merge",
    )
    rep_tab = author_table_merge(req_tab)
    assert rep_tab.passed is False
    assert rep_tab.code == "table_merge_content_loss_unsupported"

    # Non-origin cell containing <w:br/>
    c2_br = "<w:tc><w:p><w:r><w:br/></w:r></w:p></w:tc>"
    doc_br = tmp_path / "br.docx"
    write_package(doc_br, f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c1}{c2_br}{c3}</w:tr></w:tbl>")
    fp_br = table_merge_structure_fingerprint(
        doc_br, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req_br = TableMergeRequest(
        source=doc_br,
        destination=tmp_path / "out_br.docx",
        source_sha256=_sha256(doc_br),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp_br,
        action="merge",
    )
    rep_br = author_table_merge(req_br)
    assert rep_br.passed is False
    assert rep_br.code == "table_merge_content_loss_unsupported"

    # Non-origin cell containing <w:sym .../>
    c2_sym = '<w:tc><w:p><w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r></w:p></w:tc>'
    doc_sym = tmp_path / "sym.docx"
    write_package(
        doc_sym,
        f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c1}{c2_sym}{c3}</w:tr></w:tbl>",
    )
    fp_sym = table_merge_structure_fingerprint(
        doc_sym, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req_sym = TableMergeRequest(
        source=doc_sym,
        destination=tmp_path / "out_sym.docx",
        source_sha256=_sha256(doc_sym),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp_sym,
        action="merge",
    )
    rep_sym = author_table_merge(req_sym)
    assert rep_sym.passed is False
    assert rep_sym.code == "table_merge_content_loss_unsupported"


def test_author_table_merge_refusal_malformed_merge_markup(tmp_path: Path) -> None:
    # Duplicate gridSpan
    c_dup_span = (
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc>'
    )
    c_other = cell(paragraph("Other"))
    doc_dup_span = tmp_path / "dup_span.docx"
    write_package(
        doc_dup_span,
        f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c_dup_span}{c_other}</w:tr></w:tbl>",
    )
    req = TableMergeRequest(
        source=doc_dup_span,
        destination=tmp_path / "out.docx",
        source_sha256=_sha256(doc_dup_span),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_structure_unsupported"

    # Duplicate vMerge
    c_dup_vm = '<w:tc><w:tcPr><w:vMerge w:val="restart"/><w:vMerge/></w:tcPr><w:p/></w:tc>'
    doc_dup_vm = tmp_path / "dup_vm.docx"
    write_package(
        doc_dup_vm,
        f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c_dup_vm}{c_other}{c_other}</w:tr></w:tbl>",
    )
    req_vm = TableMergeRequest(
        source=doc_dup_vm,
        destination=tmp_path / "out2.docx",
        source_sha256=_sha256(doc_dup_vm),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    rep_vm = author_table_merge(req_vm)
    assert rep_vm.passed is False
    assert rep_vm.code == "table_merge_structure_unsupported"

    # Invalid vMerge value
    c_bogus_vm = '<w:tc><w:tcPr><w:vMerge w:val="bogus"/></w:tcPr><w:p/></w:tc>'
    doc_bogus_vm = tmp_path / "bogus_vm.docx"
    write_package(
        doc_bogus_vm,
        f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c_bogus_vm}{c_other}{c_other}</w:tr></w:tbl>",
    )
    req_bvm = TableMergeRequest(
        source=doc_bogus_vm,
        destination=tmp_path / "out3.docx",
        source_sha256=_sha256(doc_bogus_vm),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    rep_bvm = author_table_merge(req_bvm)
    assert rep_bvm.passed is False
    assert rep_bvm.code == "table_merge_structure_unsupported"

    # Legacy hMerge
    c_hmerge = '<w:tc><w:tcPr><w:hMerge w:val="restart"/></w:tcPr><w:p/></w:tc>'
    doc_hmerge = tmp_path / "hmerge.docx"
    write_package(
        doc_hmerge,
        f"<w:tbl>{grid([2500, 2500, 2500])}<w:tr>{c_hmerge}{c_other}{c_other}</w:tr></w:tbl>",
    )
    req_hm = TableMergeRequest(
        source=doc_hmerge,
        destination=tmp_path / "out4.docx",
        source_sha256=_sha256(doc_hmerge),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=_DUMMY_FP,
        action="merge",
    )
    rep_hm = author_table_merge(req_hm)
    assert rep_hm.passed is False
    assert rep_hm.code == "table_merge_structure_unsupported"


def test_author_table_merge_refusal_missing_or_non_positive_grid_width(
    tmp_path: Path,
) -> None:
    # Non-positive gridCol width (w="0")
    zero_grid = '<w:tblGrid><w:gridCol w:w="0"/><w:gridCol w:w="2500"/></w:tblGrid>'
    row_xml = f"<w:tr>{cell(paragraph('A'))}{cell(paragraph(''))}</w:tr>"
    doc_zero = tmp_path / "zero_width.docx"
    write_package(doc_zero, f"<w:tbl>{zero_grid}{row_xml}</w:tbl>")

    fp = table_merge_structure_fingerprint(
        doc_zero, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req = TableMergeRequest(
        source=doc_zero,
        destination=tmp_path / "out_zero.docx",
        source_sha256=_sha256(doc_zero),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is False
    assert rep.code == "table_merge_grid_mismatch"

    # Missing width attribute on gridCol
    missing_grid = "<w:tblGrid><w:gridCol/><w:gridCol w:w='2500'/></w:tblGrid>"
    doc_missing = tmp_path / "missing_width.docx"
    write_package(doc_missing, f"<w:tbl>{missing_grid}{row_xml}</w:tbl>")
    fp2 = table_merge_structure_fingerprint(
        doc_missing, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req2 = TableMergeRequest(
        source=doc_missing,
        destination=tmp_path / "out_missing.docx",
        source_sha256=_sha256(doc_missing),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp2,
        action="merge",
    )
    rep2 = author_table_merge(req2)
    assert rep2.passed is False
    assert rep2.code == "table_merge_grid_mismatch"


def test_author_table_merge_split_branch_validation_and_continuation_preservation(
    tmp_path: Path,
) -> None:
    # Split of vertical merge where continuation cell has text
    vm_restart = '<w:vMerge w:val="restart"/>'
    vm_cont = "<w:vMerge/>"
    r1 = f"<w:tr>{cell(paragraph('Origin Header'), properties=vm_restart)}{cell(paragraph('Col B'))}</w:tr>"
    r2 = f"<w:tr>{cell(paragraph('Continuation Note'), properties=vm_cont)}{cell(paragraph('B2'))}</w:tr>"
    doc_cont = tmp_path / "vert_cont_text.docx"
    write_package(doc_cont, f"<w:tbl>{grid([2500, 2500])}{r1}{r2}</w:tbl>")

    fp = table_merge_structure_fingerprint(
        doc_cont, table=1, start_row=1, end_row=2, start_col=1, end_col=1
    )
    split_doc = tmp_path / "vert_cont_split.docx"
    req = TableMergeRequest(
        source=doc_cont,
        destination=split_doc,
        source_sha256=_sha256(doc_cont),
        target={"table": 1, "start_row": 1, "end_row": 2, "start_col": 1, "end_col": 1},
        structure_fingerprint_sha256=fp,
        action="split",
    )
    rep = author_table_merge(req)
    assert rep.passed is True
    assert rep.status == "published"

    # Verify continuation text is preserved in row 2 cell 1
    graph = PackageGraph.from_file(split_doc)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)
    rows = next(iter(root.iter(word("tbl")))).findall(word("tr"))
    r1_text = rows[0].findall(word("tc"))[0].find(f".//{word('t')}").text
    r2_text = rows[1].findall(word("tc"))[0].find(f".//{word('t')}").text
    assert r1_text == "Origin Header"
    assert r2_text == "Continuation Note"

    # Split on cell containing nested table is refused
    nested_c = f'<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:tbl>{grid([1000])}<w:tr>{cell(paragraph("inner"))}</w:tr></w:tbl></w:tc>'
    doc_nested = tmp_path / "nested_split.docx"
    write_package(doc_nested, f"<w:tbl>{grid([2500, 2500])}<w:tr>{nested_c}</w:tr></w:tbl>")
    fp_nested = table_merge_structure_fingerprint(
        doc_nested, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    req_nested = TableMergeRequest(
        source=doc_nested,
        destination=tmp_path / "out_nested.docx",
        source_sha256=_sha256(doc_nested),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp_nested or _DUMMY_FP,
        action="split",
    )
    rep_nested = author_table_merge(req_nested)
    assert rep_nested.passed is False
    assert rep_nested.code in (
        "table_merge_geometry_unsupported",
        "table_merge_structure_fingerprint_mismatch",
    )


def test_author_table_merge_staged_verification_unselected_preservation(
    tmp_path: Path,
) -> None:
    p_before = paragraph("Introduction paragraph before table")
    p_after = paragraph("Conclusion paragraph after table")
    tbl_xml = f"<w:tbl>{grid([2000, 2000, 2000])}<w:tr>{cell(paragraph('A'))}{cell(paragraph(''))}{cell(paragraph('C'))}</w:tr><w:tr>{cell(paragraph('D'))}{cell(paragraph('E'))}{cell(paragraph('F'))}</w:tr></w:tbl>"
    doc = tmp_path / "surrounding.docx"
    write_package(doc, f"{p_before}{tbl_xml}{p_after}")

    fp = table_merge_structure_fingerprint(
        doc, table=1, start_row=1, end_row=1, start_col=1, end_col=2
    )
    merged_doc = tmp_path / "surrounding_merged.docx"
    req = TableMergeRequest(
        source=doc,
        destination=merged_doc,
        source_sha256=_sha256(doc),
        target={"table": 1, "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 2},
        structure_fingerprint_sha256=fp,
        action="merge",
    )
    rep = author_table_merge(req)
    assert rep.passed is True

    # Verify surrounding paragraphs are exactly preserved
    graph = PackageGraph.from_file(merged_doc)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)
    body = root.find(word("body"))
    body_paragraphs = body.findall(word("p"))
    assert len(body_paragraphs) == 2
    assert body_paragraphs[0].find(f".//{word('t')}").text == "Introduction paragraph before table"
    assert body_paragraphs[1].find(f".//{word('t')}").text == "Conclusion paragraph after table"

    # Verify unselected row 2 is preserved
    tbl = body.find(word("tbl"))
    r2_cells = tbl.findall(word("tr"))[1].findall(word("tc"))
    assert len(r2_cells) == 3
    assert r2_cells[0].find(f".//{word('t')}").text == "D"
    assert r2_cells[1].find(f".//{word('t')}").text == "E"
    assert r2_cells[2].find(f".//{word('t')}").text == "F"
