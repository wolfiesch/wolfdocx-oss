"""Tests for source-bound table row authoring and discovery."""

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
    write_styled_table_package,
)
from wolfdocx.inspect import Q
from wolfdocx.package import PackageGraph
from wolfdocx.table_structure import (
    TableRowRequest,
    author_table_row,
    discover_table_row_targets,
    load_table_row_request,
    parse_table_row_request,
    table_row_structure_fingerprint,
)

_DOCUMENT = "word/document.xml"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_sample_package(
    path: Path,
    *,
    rows_data: list[list[str]] | None = None,
    repeat_header: bool = False,
    grid_widths: list[int] | None = None,
) -> Path:
    """Create a clean, approved table fixture package for testing."""
    if rows_data is None:
        rows_data = [
            ["Header A", "Header B", "Header C"],
            ["Data 1", "Data 2", "Data 3"],
            ["Data 4", "Data 5", "Data 6"],
        ]
    cols = len(rows_data[0])
    widths = grid_widths if grid_widths is not None else [2500] * cols
    table_grid = grid(widths)

    rows_xml: list[str] = []
    for idx, r_data in enumerate(rows_data):
        cells_xml = "".join(cell(paragraph(txt)) for txt in r_data)
        r_pr = "<w:tblHeader/>" if (idx == 0 and repeat_header) else ""
        rows_xml.append(row(cells_xml, properties=r_pr))

    body = table("".join(rows_xml), grid=table_grid)
    return write_package(path, body)


def test_discover_table_row_targets_fixture(tmp_path: Path) -> None:
    fixture = write_styled_table_package(tmp_path / "styled-table.docx")
    discovery = discover_table_row_targets(fixture)
    assert discovery["passed"] is True
    assert discovery["operation"] == "table-row-targets"
    assert discovery["status"] == "discovered"
    assert discovery["code"] is None
    assert discovery["source_sha256"] == _sha256(fixture)
    assert discovery["part"] == _DOCUMENT
    assert len(discovery["targets"]) > 0

    first_target = discovery["targets"][0]
    assert first_target["table"] == 1
    assert first_target["row"] == 2
    assert first_target["cells_count"] == 2
    assert first_target["cell_texts"] == ["foo", "bar"]
    assert len(first_target["structure_fingerprint_sha256"]) == 64
    assert isinstance(first_target["is_repeat_header"], bool)


def test_author_table_row_on_styled_table_fixture(tmp_path: Path) -> None:
    fixture = write_styled_table_package(tmp_path / "styled-table.docx")
    dest = tmp_path / "fixture_row_added.docx"
    discovery = discover_table_row_targets(fixture)
    assert discovery["passed"] is True
    target = discovery["targets"][0]
    req = TableRowRequest(
        source=fixture,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target={"table": target["table"], "row": target["row"]},
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="after",
        cell_texts=("added foo", "added bar"),
    )
    report = author_table_row(req)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["table"] == 1
    assert report.target_change_report["prototype_row"] == 2
    assert report.target_change_report["inserted_row"] == 3
    assert report.target_change_report["rows_count_before"] == 4
    assert report.target_change_report["rows_count_after"] == 5
    assert report.non_target_parts_preserved is True
    assert dest.is_file()


def test_discover_table_row_targets_missing_and_unsupported(tmp_path: Path) -> None:
    non_existent = tmp_path / "missing.docx"
    missing_report = discover_table_row_targets(non_existent)
    assert missing_report["passed"] is False
    assert missing_report["code"] == "table_row_source_missing"
    assert missing_report["targets"] == []

    # Valid package but unsupported part
    source = _make_sample_package(tmp_path / "sample.docx")
    unsupported_part_report = discover_table_row_targets(source, part="word/header1.xml")
    assert unsupported_part_report["passed"] is False
    assert unsupported_part_report["code"] == "table_row_part_unsupported"
    assert unsupported_part_report["targets"] == []


def test_author_table_row_after_success(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"

    discovery = discover_table_row_targets(source)
    assert discovery["passed"] is True
    proto_target = discovery["targets"][1]  # row 2

    req_dict = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "row": 2},
        "structure_fingerprint_sha256": proto_target["structure_fingerprint_sha256"],
        "position": "after",
        "cell_texts": ["New Data 1", "New Data 2", "New Data 3"],
    }
    request = parse_table_row_request(req_dict)
    report = author_table_row(request)

    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == discovery["source_sha256"]
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True

    # Verify target_change_report facts
    change = report.target_change_report
    assert change is not None
    assert change["table"] == 1
    assert change["prototype_row"] == 2
    assert change["position"] == "after"
    assert change["inserted_row"] == 3
    assert change["rows_count_before"] == 3
    assert change["rows_count_after"] == 4
    assert change["cells_count"] == 3

    # Inspect resulting document XML
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tables = list(root.iter(Q(ns, "tbl")))
    assert len(tables) == 1
    rows = tables[0].findall(Q(ns, "tr"))
    assert len(rows) == 4

    # Check that row 1 and 2 are untouched
    row1_texts = ["".join(c.itertext()) for c in rows[0].findall(Q(ns, "tc"))]
    assert row1_texts == ["Header A", "Header B", "Header C"]
    row2_texts = ["".join(c.itertext()) for c in rows[1].findall(Q(ns, "tc"))]
    assert row2_texts == ["Data 1", "Data 2", "Data 3"]

    # Check that row 3 is the inserted row with new texts
    row3_texts = ["".join(c.itertext()) for c in rows[2].findall(Q(ns, "tc"))]
    assert row3_texts == ["New Data 1", "New Data 2", "New Data 3"]

    # Check that row 4 is the old row 3
    row4_texts = ["".join(c.itertext()) for c in rows[3].findall(Q(ns, "tc"))]
    assert row4_texts == ["Data 4", "Data 5", "Data 6"]

    # Check tblGrid preservation
    staged_grid = tables[0].find(Q(ns, "tblGrid"))
    assert staged_grid is not None
    grid_cols = staged_grid.findall(Q(ns, "gridCol"))
    assert len(grid_cols) == 3


def test_author_table_row_before_success(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_before.docx")
    dest = tmp_path / "dest_before.docx"

    discovery = discover_table_row_targets(source)
    assert discovery["passed"] is True
    proto_target = discovery["targets"][0]  # row 1

    req_dict = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "row": 1},
        "structure_fingerprint_sha256": proto_target["structure_fingerprint_sha256"],
        "position": "before",
        "cell_texts": ["Inserted Head 1", "Inserted Head 2", "Inserted Head 3"],
    }
    request = parse_table_row_request(req_dict)
    report = author_table_row(request)

    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["inserted_row"] == 1
    assert report.target_change_report["rows_count_before"] == 3
    assert report.target_change_report["rows_count_after"] == 4

    # Verify first row is the inserted row
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 4
    row1_texts = ["".join(c.itertext()) for c in rows[0].findall(Q(ns, "tc"))]
    assert row1_texts == ["Inserted Head 1", "Inserted Head 2", "Inserted Head 3"]
    row2_texts = ["".join(c.itertext()) for c in rows[1].findall(Q(ns, "tc"))]
    assert row2_texts == ["Header A", "Header B", "Header C"]


def test_author_table_row_without_replacement_texts(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_clone.docx")
    dest = tmp_path / "dest_clone.docx"

    discovery = discover_table_row_targets(source)
    proto_target = discovery["targets"][1]  # row 2

    # Omit cell_texts to test deep-copy of prototype content
    req_dict = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "row": 2},
        "structure_fingerprint_sha256": proto_target["structure_fingerprint_sha256"],
        "position": "after",
    }
    request = parse_table_row_request(req_dict)
    assert request.cell_texts is None

    report = author_table_row(request)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 4
    # Cloned row has duplicate of row 2 text
    cloned_texts = ["".join(c.itertext()) for c in rows[2].findall(Q(ns, "tc"))]
    assert cloned_texts == ["Data 1", "Data 2", "Data 3"]


def test_author_table_row_repeat_header_true_and_false(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_hdr.docx", repeat_header=False)
    discovery = discover_table_row_targets(source)
    proto_target = discovery["targets"][0]

    # Test repeat_header=True
    dest_true = tmp_path / "dest_hdr_true.docx"
    req_true = TableRowRequest(
        source=source,
        destination=dest_true,
        source_sha256=discovery["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=proto_target["structure_fingerprint_sha256"],
        position="after",
        repeat_header=True,
    )
    report_true = author_table_row(req_true)
    assert report_true.passed is True
    assert report_true.target_change_report is not None
    assert report_true.target_change_report["repeat_header"] is True

    # Verify w:tblHeader is present on inserted row
    graph_true = PackageGraph.from_file(dest_true)
    root_true = ET.fromstring(graph_true.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl_true = root_true.find(f".//{Q(ns, 'tbl')}")
    assert tbl_true is not None
    rows_true = tbl_true.findall(Q(ns, "tr"))
    ins_trpr = rows_true[1].find(Q(ns, "trPr"))
    assert ins_trpr is not None
    assert ins_trpr.find(Q(ns, "tblHeader")) is not None

    # Test repeat_header=False on a table that originally had repeat_header=True
    source_repeat = _make_sample_package(tmp_path / "source_hdr_repeat.docx", repeat_header=True)
    disc_repeat = discover_table_row_targets(source_repeat)
    dest_false = tmp_path / "dest_hdr_false.docx"
    req_false = TableRowRequest(
        source=source_repeat,
        destination=dest_false,
        source_sha256=disc_repeat["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=disc_repeat["targets"][0]["structure_fingerprint_sha256"],
        position="after",
        repeat_header=False,
    )
    report_false = author_table_row(req_false)
    assert report_false.passed is True
    assert report_false.target_change_report is not None
    assert report_false.target_change_report["repeat_header"] is False

    graph_false = PackageGraph.from_file(dest_false)
    root_false = ET.fromstring(graph_false.read_part_bytes(_DOCUMENT))
    tbl_false = root_false.find(f".//{Q(ns, 'tbl')}")
    assert tbl_false is not None
    rows_false = tbl_false.findall(Q(ns, "tr"))
    ins_trpr_false = rows_false[1].find(Q(ns, "trPr"))
    if ins_trpr_false is not None:
        assert ins_trpr_false.find(Q(ns, "tblHeader")) is None


def test_author_table_row_source_hash_mismatch_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_hash.docx")
    dest = tmp_path / "dest_hash.docx"
    discovery = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256="0" * 64,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=discovery["targets"][0]["structure_fingerprint_sha256"],
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_source_hash_mismatch"
    assert not dest.exists()


def test_author_table_row_fingerprint_mismatch_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_fp.docx")
    dest = tmp_path / "dest_fp.docx"
    discovery = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256="f" * 64,
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_structure_fingerprint_mismatch"
    assert not dest.exists()


def test_author_table_row_destination_occupied_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_occ.docx")
    dest = tmp_path / "dest_occ.docx"
    dest.write_bytes(b"existing content")
    discovery = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=discovery["targets"][0]["structure_fingerprint_sha256"],
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_destination_occupied"
    # Preserved original content unchanged
    assert dest.read_bytes() == b"existing content"


def test_author_table_row_cell_texts_mismatch_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_mismatch.docx")
    dest = tmp_path / "dest_mismatch.docx"
    discovery = discover_table_row_targets(source)

    # 2 texts provided for a 3-column row
    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=discovery["targets"][0]["structure_fingerprint_sha256"],
        cell_texts=("Col 1", "Col 2"),
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_cell_texts_mismatch"
    assert not dest.exists()


def test_author_table_row_merged_cells_refusal(tmp_path: Path) -> None:
    # Build package with a vertically merged cell in row 2
    row1 = row(cell(paragraph("H1")) + cell(paragraph("H2")))
    row2_merged = row(
        cell(paragraph("D1"), properties='<w:vMerge w:val="restart"/>') + cell(paragraph("D2"))
    )
    body = table(row1 + row2_merged, grid=grid([2000, 2000]))
    source = write_package(tmp_path / "merged.docx", body)
    dest = tmp_path / "dest_merged.docx"

    # Row 2 contains vMerge
    fp = table_row_structure_fingerprint(source, table=1, row=2)
    # Since row 2 has vMerge, fingerprint is None (unapproved)
    assert fp is None

    # Even if passing fake fingerprint, author_table_row refuses with table_row_merged_unsupported
    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 2},
        structure_fingerprint_sha256="0" * 64,
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_merged_unsupported"
    assert not dest.exists()


def test_author_table_row_grid_mismatch_refusal(tmp_path: Path) -> None:
    # Table grid defines 3 columns, but row has only 2 cells
    row1 = row(cell(paragraph("A")) + cell(paragraph("B")))
    body = table(row1, grid=grid([2000, 2000, 2000]))
    source = write_package(tmp_path / "grid_mismatch.docx", body)
    dest = tmp_path / "dest_grid_mismatch.docx"

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256="0" * 64,
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code in ("table_row_grid_mismatch", "table_row_table_unsupported")
    assert not dest.exists()


def test_author_table_row_unsupported_descendants_refusal(tmp_path: Path) -> None:
    # Cell with a bookmark
    bk_cell = (
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:bookmarkStart w:id="1" w:name="b1"/>'
        "<w:r><w:t>Bookmarked</w:t></w:r>"
        '<w:bookmarkEnd w:id="1"/></w:p></w:tc>'
    )
    r1 = f"<w:tr>{bk_cell}</w:tr>"
    body = table(r1, grid=grid([2000]))
    source = write_package(tmp_path / "bookmark_row.docx", body)
    dest = tmp_path / "dest_bk.docx"

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256="0" * 64,
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code in (
        "table_row_structure_unsupported",
        "table_row_table_unsupported",
    )
    assert not dest.exists()


def test_parse_and_load_table_row_request(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    source.write_bytes(b"dummy")

    valid_payload = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": "a" * 64,
        "target": {"table": 1, "row": 2},
        "structure_fingerprint_sha256": "b" * 64,
        "position": "after",
        "cell_texts": ["A", "B"],
        "repeat_header": False,
    }
    req = parse_table_row_request(valid_payload)
    assert req.source == source
    assert req.destination == dest
    assert req.table == 1
    assert req.row == 2
    assert req.position == "after"
    assert req.cell_texts == ("A", "B")
    assert req.repeat_header is False

    # Round-trip through to_dict
    serialized = req.to_dict()
    assert serialized["source"] == str(source)
    assert serialized["target"] == {"part": _DOCUMENT, "table": 1, "row": 2}

    # Test load_table_row_request from disk
    req_file = tmp_path / "request.json"
    req_file.write_text(json.dumps(valid_payload), encoding="utf-8")
    loaded = load_table_row_request(req_file)
    assert loaded.source == source
    assert loaded.cell_texts == ("A", "B")

    # Reject unknown keys
    invalid_payload = {**valid_payload, "extra_key": 123}
    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        parse_table_row_request(invalid_payload)

    # Reject unknown target keys
    invalid_target_payload = {
        **valid_payload,
        "target": {"table": 1, "row": 1, "unknown": "val"},
    }
    with pytest.raises(ValueError, match="table_row_target_fields_invalid"):
        parse_table_row_request(invalid_target_payload)

    # Reject non-mapping target
    with pytest.raises(TypeError, match="table_row_target_invalid"):
        parse_table_row_request({**valid_payload, "target": "table1"})

    # Reject invalid position
    with pytest.raises(ValueError, match="table_row_position_invalid"):
        parse_table_row_request({**valid_payload, "position": "invalid"})


def test_table_row_structure_fingerprint(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "fp_sample.docx")
    fp = table_row_structure_fingerprint(source, table=1, row=1)
    assert fp is not None
    assert len(fp) == 64

    # Non-existent table/row returns None
    assert table_row_structure_fingerprint(source, table=99, row=1) is None
    assert table_row_structure_fingerprint(source, table=1, row=99) is None
    assert table_row_structure_fingerprint(source, table=1, row=1, part="word/invalid.xml") is None


def test_delete_first_row_success(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_first.docx")
    dest = tmp_path / "dest_del_first.docx"

    discovery = discover_table_row_targets(source)
    assert discovery["passed"] is True
    proto_target = discovery["targets"][0]  # row 1
    assert "delete" in proto_target["actions"]

    req_dict = {
        "action": "delete",
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "row": 1},
        "structure_fingerprint_sha256": proto_target["structure_fingerprint_sha256"],
    }
    request = parse_table_row_request(req_dict)
    assert request.action == "delete"
    assert request.position is None
    assert request.cell_texts is None
    assert request.repeat_header is None

    report = author_table_row(request)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == discovery["source_sha256"]
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True

    change = report.target_change_report
    assert change is not None
    assert change["action"] == "delete"
    assert change["table"] == 1
    assert change["row"] == 1
    assert change["deleted_row"] == 1
    assert change["rows_count_before"] == 3
    assert change["rows_count_after"] == 2
    assert change["cells_count"] == 3

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 2
    row1_texts = ["".join(c.itertext()) for c in rows[0].findall(Q(ns, "tc"))]
    assert row1_texts == ["Data 1", "Data 2", "Data 3"]
    row2_texts = ["".join(c.itertext()) for c in rows[1].findall(Q(ns, "tc"))]
    assert row2_texts == ["Data 4", "Data 5", "Data 6"]


def test_delete_middle_row_success(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_mid.docx")
    dest = tmp_path / "dest_del_mid.docx"

    discovery = discover_table_row_targets(source)
    assert discovery["passed"] is True
    proto_target = discovery["targets"][1]  # row 2
    assert "delete" in proto_target["actions"]

    req_dict = {
        "action": "delete",
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "row": 2},
        "structure_fingerprint_sha256": proto_target["structure_fingerprint_sha256"],
    }
    request = parse_table_row_request(req_dict)
    report = author_table_row(request)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["deleted_row"] == 2
    assert report.target_change_report["rows_count_before"] == 3
    assert report.target_change_report["rows_count_after"] == 2

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 2
    row1_texts = ["".join(c.itertext()) for c in rows[0].findall(Q(ns, "tc"))]
    assert row1_texts == ["Header A", "Header B", "Header C"]
    row2_texts = ["".join(c.itertext()) for c in rows[1].findall(Q(ns, "tc"))]
    assert row2_texts == ["Data 4", "Data 5", "Data 6"]


def test_delete_last_row_success(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_last.docx")
    dest = tmp_path / "dest_del_last.docx"

    discovery = discover_table_row_targets(source)
    assert discovery["passed"] is True
    proto_target = discovery["targets"][2]  # row 3
    assert "delete" in proto_target["actions"]

    req_dict = {
        "action": "delete",
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "row": 3},
        "structure_fingerprint_sha256": proto_target["structure_fingerprint_sha256"],
    }
    request = parse_table_row_request(req_dict)
    report = author_table_row(request)
    assert report.passed is True
    assert report.target_change_report is not None
    assert report.target_change_report["deleted_row"] == 3
    assert report.target_change_report["rows_count_before"] == 3
    assert report.target_change_report["rows_count_after"] == 2

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 2
    row1_texts = ["".join(c.itertext()) for c in rows[0].findall(Q(ns, "tc"))]
    assert row1_texts == ["Header A", "Header B", "Header C"]
    row2_texts = ["".join(c.itertext()) for c in rows[1].findall(Q(ns, "tc"))]
    assert row2_texts == ["Data 1", "Data 2", "Data 3"]


def test_delete_row_neighbor_preservation_and_grid_integrity(tmp_path: Path) -> None:
    rows_data = [
        ["R1C1", "R1C2", "R1C3"],
        ["R2C1", "R2C2", "R2C3"],
        ["R3C1", "R3C2", "R3C3"],
        ["R4C1", "R4C2", "R4C3"],
    ]
    widths = [1500, 2500, 3500]
    source = _make_sample_package(
        tmp_path / "source_preserve.docx",
        rows_data=rows_data,
        grid_widths=widths,
    )
    dest = tmp_path / "dest_preserve.docx"

    disc = discover_table_row_targets(source)
    assert disc["passed"] is True
    fp_row2 = disc["targets"][1]["structure_fingerprint_sha256"]

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=disc["source_sha256"],
        target={"table": 1, "row": 2},
        structure_fingerprint_sha256=fp_row2,
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is True
    assert report.non_target_parts_preserved is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    staged_grid = tbl.find(Q(ns, "tblGrid"))
    assert staged_grid is not None
    cols = [int(c.get(Q(ns, "w"), "0")) for c in staged_grid.findall(Q(ns, "gridCol"))]
    assert cols == widths

    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 3
    remaining_texts = [["".join(c.itertext()) for c in r.findall(Q(ns, "tc"))] for r in rows]
    assert remaining_texts == [
        ["R1C1", "R1C2", "R1C3"],
        ["R3C1", "R3C2", "R3C3"],
        ["R4C1", "R4C2", "R4C3"],
    ]


def test_delete_row_repeated_follow_on_edits(tmp_path: Path) -> None:
    rows_data = [
        ["Row 1", "Alpha"],
        ["Row 2", "Beta"],
        ["Row 3", "Gamma"],
        ["Row 4", "Delta"],
    ]
    step0_doc = _make_sample_package(tmp_path / "step0.docx", rows_data=rows_data)

    step1_dest = tmp_path / "step1.docx"
    disc0 = discover_table_row_targets(step0_doc)
    req1 = TableRowRequest(
        source=step0_doc,
        destination=step1_dest,
        source_sha256=disc0["source_sha256"],
        target={"table": 1, "row": 2},
        structure_fingerprint_sha256=disc0["targets"][1]["structure_fingerprint_sha256"],
        action="delete",
    )
    rep1 = author_table_row(req1)
    assert rep1.passed is True

    step2_dest = tmp_path / "step2.docx"
    disc1 = discover_table_row_targets(step1_dest)
    assert len(disc1["targets"]) == 3
    req2 = TableRowRequest(
        source=step1_dest,
        destination=step2_dest,
        source_sha256=disc1["source_sha256"],
        target={"table": 1, "row": 2},
        structure_fingerprint_sha256=disc1["targets"][1]["structure_fingerprint_sha256"],
        action="delete",
    )
    rep2 = author_table_row(req2)
    assert rep2.passed is True

    step3_dest = tmp_path / "step3.docx"
    disc2 = discover_table_row_targets(step2_dest)
    assert len(disc2["targets"]) == 2
    req3 = TableRowRequest(
        source=step2_dest,
        destination=step3_dest,
        source_sha256=disc2["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=disc2["targets"][0]["structure_fingerprint_sha256"],
        position="after",
        cell_texts=("Inserted Row", "Epsilon"),
        action="insert",
    )
    rep3 = author_table_row(req3)
    assert rep3.passed is True

    step4_dest = tmp_path / "step4.docx"
    disc3 = discover_table_row_targets(step3_dest)
    assert len(disc3["targets"]) == 3
    req4 = TableRowRequest(
        source=step3_dest,
        destination=step4_dest,
        source_sha256=disc3["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=disc3["targets"][0]["structure_fingerprint_sha256"],
        action="delete",
    )
    rep4 = author_table_row(req4)
    assert rep4.passed is True

    graph = PackageGraph.from_file(step4_dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    assert len(rows) == 2
    final_texts = [["".join(c.itertext()) for c in r.findall(Q(ns, "tc"))] for r in rows]
    assert final_texts == [
        ["Inserted Row", "Epsilon"],
        ["Row 4", "Delta"],
    ]


def test_delete_final_row_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(
        tmp_path / "single_row.docx",
        rows_data=[["Only Row Col 1", "Only Row Col 2"]],
    )
    dest = tmp_path / "dest_single.docx"
    disc = discover_table_row_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 1
    assert disc["targets"][0]["actions"] == ["insert"]
    assert "delete" not in disc["targets"][0]["actions"]

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=disc["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=disc["targets"][0]["structure_fingerprint_sha256"],
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_final_row_unsupported"
    assert not dest.exists()


def test_delete_row_merged_unsupported_refusal(tmp_path: Path) -> None:
    row1 = row(cell(paragraph("H1")) + cell(paragraph("H2")))
    row2_merged = row(
        cell(paragraph("D1"), properties='<w:vMerge w:val="restart"/>') + cell(paragraph("D2"))
    )
    body = table(row1 + row2_merged, grid=grid([2000, 2000]))
    source = write_package(tmp_path / "merged_del.docx", body)
    dest = tmp_path / "dest_merged_del.docx"

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256="0" * 64,
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_merged_unsupported"
    assert not dest.exists()


def test_delete_row_source_hash_mismatch_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_hash.docx")
    dest = tmp_path / "dest_del_hash.docx"
    disc = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256="0" * 64,
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=disc["targets"][0]["structure_fingerprint_sha256"],
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_source_hash_mismatch"
    assert not dest.exists()


def test_delete_row_fingerprint_mismatch_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_fp.docx")
    dest = tmp_path / "dest_del_fp.docx"
    disc = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=disc["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256="f" * 64,
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_structure_fingerprint_mismatch"
    assert not dest.exists()


def test_delete_row_destination_occupied_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_occ.docx")
    dest = tmp_path / "dest_del_occ.docx"
    dest.write_bytes(b"retained content")
    disc = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=disc["source_sha256"],
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256=disc["targets"][0]["structure_fingerprint_sha256"],
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_destination_occupied"
    assert dest.read_bytes() == b"retained content"


def test_delete_row_target_out_of_range_refusal(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source_del_range.docx")
    dest = tmp_path / "dest_del_range.docx"
    disc = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=disc["source_sha256"],
        target={"table": 1, "row": 99},
        structure_fingerprint_sha256="0" * 64,
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_target_missing"
    assert not dest.exists()


def test_delete_row_relationship_bearing_refusal(tmp_path: Path) -> None:
    drawing_cell = (
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:drawing><w:inline><w:docPr id="1" name="P"/></w:inline></w:drawing></w:r></w:p></w:tc>'
    )
    r1 = f"<w:tr>{drawing_cell}</w:tr>"
    r2 = row(cell(paragraph("D1")) + cell(paragraph("D2")))
    body = table(r1 + r2, grid=grid([2000, 2000]))
    source = write_package(tmp_path / "drawing_row.docx", body)
    dest = tmp_path / "dest_drawing.docx"

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 1},
        structure_fingerprint_sha256="0" * 64,
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code in (
        "table_row_structure_unsupported",
        "table_row_table_unsupported",
    )
    assert not dest.exists()


def test_delete_row_parser_strictness_and_round_trip(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    source.write_bytes(b"dummy")

    valid_delete_dict = {
        "action": "delete",
        "source": str(source),
        "destination": str(dest),
        "source_sha256": "a" * 64,
        "target": {"table": 1, "row": 2},
        "structure_fingerprint_sha256": "b" * 64,
    }
    req = parse_table_row_request(valid_delete_dict)
    assert req.action == "delete"
    assert req.position is None
    assert req.cell_texts is None
    assert req.repeat_header is None

    serialized = req.to_dict()
    assert serialized["action"] == "delete"
    assert "position" not in serialized
    assert "cell_texts" not in serialized
    assert "repeat_header" not in serialized

    re_parsed = parse_table_row_request(serialized)
    assert re_parsed.action == "delete"
    assert re_parsed.source == source
    assert re_parsed.destination == dest
    assert re_parsed.table == 1
    assert re_parsed.row == 2
    assert re_parsed.position is None

    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        parse_table_row_request({**valid_delete_dict, "position": "after"})

    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        parse_table_row_request({**valid_delete_dict, "cell_texts": ["A", "B"]})

    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        parse_table_row_request({**valid_delete_dict, "repeat_header": True})

    with pytest.raises(ValueError, match="table_row_action_invalid"):
        parse_table_row_request({**valid_delete_dict, "action": "update"})

    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        TableRowRequest(
            source=source,
            destination=dest,
            source_sha256="a" * 64,
            target={"table": 1, "row": 2},
            structure_fingerprint_sha256="b" * 64,
            position="before",
            action="delete",
        )
    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        TableRowRequest(
            source=source,
            destination=dest,
            source_sha256="a" * 64,
            target={"table": 1, "row": 2},
            structure_fingerprint_sha256="b" * 64,
            cell_texts=("A",),
            action="delete",
        )
    with pytest.raises(ValueError, match="table_row_request_fields_invalid"):
        TableRowRequest(
            source=source,
            destination=dest,
            source_sha256="a" * 64,
            target={"table": 1, "row": 2},
            structure_fingerprint_sha256="b" * 64,
            repeat_header=True,
            action="delete",
        )


def test_delete_row_publication_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    source = _make_sample_package(tmp_path / "source_pub_fail.docx")
    dest = tmp_path / "dest_pub_fail.docx"
    disc = discover_table_row_targets(source)

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=disc["source_sha256"],
        target={"table": 1, "row": 2},
        structure_fingerprint_sha256=disc["targets"][1]["structure_fingerprint_sha256"],
        action="delete",
    )

    import wolfdocx.table_structure as ts
    from wolfdocx.inspect import inspect_docx as real_inspect_docx

    invalid_pkg = tmp_path / "invalid_mock.docx"
    invalid_pkg.write_bytes(b"not a zip")
    failed_report = real_inspect_docx(invalid_pkg)
    assert failed_report.valid is False

    monkeypatch.setattr(ts, "inspect_docx", lambda path: failed_report)

    report = author_table_row(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_row_staged_verification_failed"
    assert not dest.exists()


def test_discovery_communicates_admissible_actions_and_preserves_targets(
    tmp_path: Path,
) -> None:
    source = _make_sample_package(tmp_path / "disc_actions.docx")
    disc = discover_table_row_targets(source)
    assert disc["passed"] is True
    targets = disc["targets"]
    assert len(targets) == 3

    for t in targets:
        assert t["actions"] == ["insert", "delete"]
        assert t["cells_count"] == 3
        assert len(t["structure_fingerprint_sha256"]) == 64
        assert "is_repeat_header" in t


def test_delete_row_neighbor_with_multiple_paragraphs(tmp_path: Path) -> None:
    # Neighbor row 1 has cell with two paragraphs (valid WordprocessingML)
    r1 = row(cell(paragraph("Para 1") + paragraph("Para 2")) + cell(paragraph("Neighbor C2")))
    r2 = row(cell(paragraph("Row 2 Col 1")) + cell(paragraph("Row 2 Col 2")))
    r3 = row(cell(paragraph("Row 3 Col 1")) + cell(paragraph("Row 3 Col 2")))
    body = table(r1 + r2 + r3, grid=grid([2500, 2500]))
    source = write_package(tmp_path / "neighbor_multi_p.docx", body)
    dest = tmp_path / "dest_neighbor_multi_p.docx"

    disc = discover_table_row_targets(source)
    assert disc["passed"] is True
    fp_row2 = disc["targets"][1]["structure_fingerprint_sha256"]

    req = TableRowRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"table": 1, "row": 2},
        structure_fingerprint_sha256=fp_row2,
        action="delete",
    )
    report = author_table_row(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.target_change_report is not None
    assert report.target_change_report["deleted_row"] == 2
    assert report.target_change_report["rows_count_before"] == 3
    assert report.target_change_report["rows_count_after"] == 2
    assert report.non_target_parts_preserved is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    remaining_rows = tbl.findall(Q(ns, "tr"))
    assert len(remaining_rows) == 2
    # Verify untouched neighbor row 1 preserved its two paragraphs intact
    c1_paragraphs = remaining_rows[0].findall(f".//{Q(ns, 'tc')}")[0].findall(Q(ns, "p"))
    assert len(c1_paragraphs) == 2
    assert ["".join(p.itertext()) for p in c1_paragraphs] == ["Para 1", "Para 2"]
