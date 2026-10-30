"""Tests for source-bound table column authoring and discovery."""

from __future__ import annotations

import hashlib
import json
import zipfile
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
from wolfdocx.inspect import Q
from wolfdocx.package import PackageGraph
from wolfdocx.table_column import (
    TableColumnReport,
    TableColumnRequest,
    author_table_column,
    discover_table_column_targets,
    load_table_column_request,
    parse_table_column_request,
    table_column_structure_fingerprint,
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
    tbl_properties: str = "",
    extra_body: str = "",
) -> Path:
    """Create an admitted rectangular table package for testing."""
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
        cells_xml = "".join(
            cell(
                paragraph(txt),
                properties=f'<w:tcW w:w="{widths[c_idx]}" w:type="dxa"/>',
            )
            for c_idx, txt in enumerate(r_data)
        )
        r_pr = "<w:tblHeader/>" if (idx == 0 and repeat_header) else ""
        rows_xml.append(row(cells_xml, properties=r_pr))

    total_w = sum(widths)
    pr = f'<w:tblW w:w="{total_w}" w:type="dxa"/>{tbl_properties}'
    body = f"{table(''.join(rows_xml), grid=table_grid, properties=pr)}{extra_body}"
    return write_package(path, body)


# --- Discovery Tests ---


def test_discover_table_column_targets_success(tmp_path: Path) -> None:
    pkg = _make_sample_package(tmp_path / "sample.docx", grid_widths=[2000, 3000, 4000])
    discovery = discover_table_column_targets(pkg)

    assert discovery["passed"] is True
    assert discovery["operation"] == "table-column-targets"
    assert discovery["status"] == "discovered"
    assert discovery["code"] is None
    assert discovery["source_sha256"] == _sha256(pkg)
    assert discovery["part"] == _DOCUMENT
    assert len(discovery["targets"]) == 3

    t1 = discovery["targets"][0]
    assert t1["table"] == 1
    assert t1["column"] == 1
    assert t1["rows_count"] == 3
    assert t1["width"] == 2000
    assert t1["cell_texts"] == ["Header A", "Data 1", "Data 4"]
    assert len(t1["structure_fingerprint_sha256"]) == 64
    assert t1["actions"] == ["insert", "delete"]

    t2 = discovery["targets"][1]
    assert t2["column"] == 2
    assert t2["width"] == 3000
    assert t2["cell_texts"] == ["Header B", "Data 2", "Data 5"]

    t3 = discovery["targets"][2]
    assert t3["column"] == 3
    assert t3["width"] == 4000
    assert t3["cell_texts"] == ["Header C", "Data 3", "Data 6"]


def test_discover_table_column_targets_single_column(tmp_path: Path) -> None:
    rows = [["Only A"], ["Only B"]]
    pkg = _make_sample_package(tmp_path / "single.docx", rows_data=rows, grid_widths=[3000])
    discovery = discover_table_column_targets(pkg)

    assert discovery["passed"] is True
    assert len(discovery["targets"]) == 1
    t = discovery["targets"][0]
    assert t["column"] == 1
    assert t["actions"] == ["insert"]  # Deletion not allowed on single column


def test_discover_table_column_targets_errors_and_skipped(tmp_path: Path) -> None:
    # 1. Missing source
    missing_report = discover_table_column_targets(tmp_path / "missing.docx")
    assert missing_report["passed"] is False
    assert missing_report["code"] == "table_column_source_missing"

    # 2. Corrupt zip
    bad_zip = tmp_path / "bad.docx"
    bad_zip.write_bytes(b"not a zip")
    bad_report = discover_table_column_targets(bad_zip)
    assert bad_report["passed"] is False
    assert bad_report["code"] == "table_column_package_invalid"

    # 3. Unsupported part
    sample = _make_sample_package(tmp_path / "sample.docx")
    part_report = discover_table_column_targets(sample, part="word/header1.xml")
    assert part_report["passed"] is False
    assert part_report["code"] == "table_column_part_unsupported"

    # 4. Merged columns admit resize without offering insert/delete through the region.
    table_grid = grid([2500, 2500])
    rows_xml = row(
        cell(paragraph("1"), properties="<w:vMerge w:val='restart'/>") + cell(paragraph("2"))
    ) + row(cell(paragraph(""), properties="<w:vMerge/>") + cell(paragraph("4")))
    merged_pkg = write_package(tmp_path / "merged.docx", table(rows_xml, grid=table_grid))
    merged_discovery = discover_table_column_targets(merged_pkg)
    assert merged_discovery["passed"] is True
    assert [target["actions"] for target in merged_discovery["targets"]] == [["resize"], ["resize"]]

    # 5. Row with gridBefore is skipped in discovery
    gb_row = row(
        cell(paragraph("A")) + cell(paragraph("B")),
        properties="<w:gridBefore w:val='1'/>",
    )
    normal_row = row(cell(paragraph("C")) + cell(paragraph("D")))
    gb_pkg = write_package(tmp_path / "gb.docx", table(gb_row + normal_row, grid=table_grid))
    gb_discovery = discover_table_column_targets(gb_pkg)
    assert gb_discovery["passed"] is True
    assert gb_discovery["targets"] == []

    # 6. Non-rectangular table (mismatched cell count) is skipped in discovery
    nr_row1 = row(cell(paragraph("A")) + cell(paragraph("B")))
    nr_row2 = row(cell(paragraph("C")))  # only 1 cell instead of 2
    nr_pkg = write_package(tmp_path / "nr.docx", table(nr_row1 + nr_row2, grid=table_grid))
    nr_discovery = discover_table_column_targets(nr_pkg)
    assert nr_discovery["passed"] is True
    assert nr_discovery["targets"] == []


# --- Structural Fingerprint Tests ---


def test_table_column_structure_fingerprint(tmp_path: Path) -> None:
    pkg1 = _make_sample_package(tmp_path / "pkg1.docx", grid_widths=[2000, 3000, 4000])
    fp1 = table_column_structure_fingerprint(pkg1, table=1, column=2)
    assert fp1 is not None
    assert len(fp1) == 64

    # Fingerprint matches discovery
    discovery = discover_table_column_targets(pkg1)
    assert discovery["targets"][1]["structure_fingerprint_sha256"] == fp1

    # Text modification does not change structural fingerprint
    rows_modified_text = [
        ["Header A", "DIFFERENT TEXT", "Header C"],
        ["Data 1", "MORE DIFFERENT", "Data 3"],
        ["Data 4", "ALSO DIFFERENT", "Data 6"],
    ]
    pkg2 = _make_sample_package(
        tmp_path / "pkg2.docx",
        rows_data=rows_modified_text,
        grid_widths=[2000, 3000, 4000],
    )
    fp2 = table_column_structure_fingerprint(pkg2, table=1, column=2)
    assert fp2 == fp1

    # Structural modification (width change) DOES change structural fingerprint
    pkg3 = _make_sample_package(tmp_path / "pkg3.docx", grid_widths=[2000, 3500, 4000])
    fp3 = table_column_structure_fingerprint(pkg3, table=1, column=2)
    assert fp3 != fp1

    # Out-of-bounds column returns None
    assert table_column_structure_fingerprint(pkg1, table=1, column=99) is None
    assert table_column_structure_fingerprint(pkg1, table=99, column=1) is None


# --- Author Insertion Tests ---


def test_author_table_column_insert_after(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx", grid_widths=[2000, 3000, 4000])
    dest = tmp_path / "inserted_after.docx"
    discovery = discover_table_column_targets(source)
    target = discovery["targets"][0]  # Column 1

    req = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="after",
        width=1500,
        cell_texts=("New H", "New D1", "New D2"),
    )
    report = author_table_column(req)
    assert isinstance(report, TableColumnReport)
    assert report.to_dict()["operation"] == "table-column"

    assert report.passed is True
    assert report.code is None
    assert report.non_target_parts_preserved is True
    assert dest.is_file()

    change = report.target_change_report
    assert change is not None
    assert change["action"] == "insert"
    assert change["prototype_column"] == 1
    assert change["inserted_column"] == 2
    assert change["columns_count_before"] == 3
    assert change["columns_count_after"] == 4
    assert change["rows_count"] == 3
    assert change["width"] == 1500
    assert change["cell_text_lengths"] == [5, 6, 6]
    assert len(change["cell_texts_sha256"]) == 64
    assert "cell_texts" not in change

    # Verify readback geometry and content
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)

    tbl = root.find(f".//{word('tbl')}")
    assert tbl is not None

    # Verify tblGrid
    grid_elem = tbl.find(word("tblGrid"))
    assert grid_elem is not None
    cols = grid_elem.findall(word("gridCol"))
    assert len(cols) == 4
    assert [c.get(word("w")) for c in cols] == ["2000", "1500", "3000", "4000"]

    # Verify tblW in tblPr incremented
    tbl_w = tbl.find(f"{word('tblPr')}/{word('tblW')}")
    assert tbl_w is not None
    assert tbl_w.get(word("w")) == str(2000 + 1500 + 3000 + 4000)

    # Verify rows & cells
    rows = tbl.findall(word("tr"))
    assert len(rows) == 3
    expected_texts = [
        ["Header A", "New H", "Header B", "Header C"],
        ["Data 1", "New D1", "Data 2", "Data 3"],
        ["Data 4", "New D2", "Data 5", "Data 6"],
    ]
    for r, exp in zip(rows, expected_texts):
        cells = r.findall(word("tc"))
        assert len(cells) == 4
        actual_texts = ["".join(t.text or "" for t in c.iter(word("t"))) for c in cells]
        assert actual_texts == exp

        # Verify inserted cell tcW
        ins_cell_tcw = cells[1].find(f"{word('tcPr')}/{word('tcW')}")
        assert ins_cell_tcw is not None
        assert ins_cell_tcw.get(word("w")) == "1500"


def test_author_table_column_insert_before(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx", grid_widths=[2000, 3000, 4000])
    dest = tmp_path / "inserted_before.docx"
    discovery = discover_table_column_targets(source)
    target = discovery["targets"][0]  # Column 1

    req = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="before",
        width=1200,
        cell_texts=("First H", "First D1", "First D2"),
    )
    report = author_table_column(req)

    assert report.passed is True
    change = report.target_change_report
    assert change is not None
    assert change["inserted_column"] == 1
    assert change["columns_count_after"] == 4

    # Verify inserted column is first
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)
    tbl = root.find(f".//{word('tbl')}")
    assert tbl is not None

    cols = tbl.findall(f"{word('tblGrid')}/{word('gridCol')}")
    assert [c.get(word("w")) for c in cols] == ["1200", "2000", "3000", "4000"]

    first_row_cells = tbl.find(word("tr")).findall(word("tc"))
    first_cell_text = "".join(t.text or "" for t in first_row_cells[0].iter(word("t")))
    assert first_cell_text == "First H"


def test_author_table_column_insert_default_width_and_texts(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx", grid_widths=[2200, 3300, 4400])
    dest = tmp_path / "default_insert.docx"
    discovery = discover_table_column_targets(source)
    target = discovery["targets"][1]  # Column 2 (width 3300)

    # Omit width and cell_texts
    req = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="after",
    )
    report = author_table_column(req)

    assert report.passed is True
    change = report.target_change_report
    assert change is not None
    assert change["width"] == 3300  # Inferred from prototype column
    assert change["cell_text_lengths"] == [0, 0, 0]
    assert len(change["cell_texts_sha256"]) == 64
    assert "cell_texts" not in change

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)
    tbl = root.find(f".//{word('tbl')}")
    assert tbl is not None

    cols = tbl.findall(f"{word('tblGrid')}/{word('gridCol')}")
    assert [c.get(word("w")) for c in cols] == ["2200", "3300", "3300", "4400"]


def test_author_table_column_insert_preserves_whitespace(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "whitespace.docx"
    discovery = discover_table_column_targets(source)
    target = discovery["targets"][0]

    req = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="after",
        cell_texts=("  leading", "trailing  ", "double  space"),
    )
    report = author_table_column(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)
    xml_space = "{http://www.w3.org/XML/1998/namespace}space"

    tbl = root.find(f".//{word('tbl')}")
    assert tbl is not None
    rows = tbl.findall(word("tr"))
    for r in rows:
        inserted_cell = r.findall(word("tc"))[1]
        t_elem = inserted_cell.find(f".//{word('t')}")
        assert t_elem is not None
        assert t_elem.get(xml_space) == "preserve"


# --- Author Deletion Tests ---


def test_author_table_column_delete_middle(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx", grid_widths=[2000, 3000, 4000])
    dest = tmp_path / "deleted_mid.docx"
    discovery = discover_table_column_targets(source)
    target = discovery["targets"][1]  # Column 2 (width 3000)

    req = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        action="delete",
    )
    report = author_table_column(req)

    assert report.passed is True
    assert report.code is None
    assert report.non_target_parts_preserved is True
    assert dest.is_file()

    change = report.target_change_report
    assert change is not None
    assert change["action"] == "delete"
    assert change["deleted_column"] == 2
    assert change["columns_count_before"] == 3
    assert change["columns_count_after"] == 2
    assert change["rows_count"] == 3
    assert change["width"] == 3000

    # Readback verification
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = partial(Q, ns)
    tbl = root.find(f".//{word('tbl')}")
    assert tbl is not None

    cols = tbl.findall(f"{word('tblGrid')}/{word('gridCol')}")
    assert len(cols) == 2
    assert [c.get(word("w")) for c in cols] == ["2000", "4000"]

    tbl_w = tbl.find(f"{word('tblPr')}/{word('tblW')}")
    assert tbl_w is not None
    assert tbl_w.get(word("w")) == str(9000 - 3000)

    rows = tbl.findall(word("tr"))
    expected_texts = [
        ["Header A", "Header C"],
        ["Data 1", "Data 3"],
        ["Data 4", "Data 6"],
    ]
    for r, exp in zip(rows, expected_texts):
        cells = r.findall(word("tc"))
        assert len(cells) == 2
        actual_texts = ["".join(t.text or "" for t in c.iter(word("t"))) for c in cells]
        assert actual_texts == exp


def test_author_table_column_delete_first_and_last(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx", grid_widths=[2000, 3000, 4000])
    discovery = discover_table_column_targets(source)

    # Delete column 1
    dest1 = tmp_path / "del_col1.docx"
    t1 = discovery["targets"][0]
    req1 = TableColumnRequest(
        source=source,
        destination=dest1,
        source_sha256=discovery["source_sha256"],
        target=t1["target"],
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
        action="delete",
    )
    rep1 = author_table_column(req1)
    assert rep1.passed is True
    assert rep1.target_change_report["deleted_column"] == 1
    assert rep1.target_change_report["columns_count_after"] == 2

    # Delete column 3
    dest3 = tmp_path / "del_col3.docx"
    t3 = discovery["targets"][2]
    req3 = TableColumnRequest(
        source=source,
        destination=dest3,
        source_sha256=discovery["source_sha256"],
        target=t3["target"],
        structure_fingerprint_sha256=t3["structure_fingerprint_sha256"],
        action="delete",
    )
    rep3 = author_table_column(req3)
    assert rep3.passed is True
    assert rep3.target_change_report["deleted_column"] == 3
    assert rep3.target_change_report["columns_count_after"] == 2


def test_author_table_column_delete_final_column_refused(tmp_path: Path) -> None:
    single = _make_sample_package(
        tmp_path / "single.docx", rows_data=[["A"], ["B"]], grid_widths=[3000]
    )
    dest = tmp_path / "should_not_exist.docx"
    discovery = discover_table_column_targets(single)
    t = discovery["targets"][0]

    req = TableColumnRequest(
        source=single,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
        action="delete",
    )
    report = author_table_column(req)

    assert report.passed is False
    assert report.code == "table_column_final_column_unsupported"
    assert not dest.exists()


# --- Validation & Refusal Tests ---


def test_author_table_column_refusals_source_dest(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    discovery = discover_table_column_targets(source)
    t = discovery["targets"][0]

    # 1. Missing source
    req_missing = TableColumnRequest(
        source=tmp_path / "nonexistent.docx",
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
    )
    rep = author_table_column(req_missing)
    assert rep.passed is False
    assert rep.code == "table_column_source_missing"

    # 2. Source hash mismatch
    req_bad_hash = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256="0" * 64,
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
    )
    rep = author_table_column(req_bad_hash)
    assert rep.passed is False
    assert rep.code == "table_column_source_hash_mismatch"

    # 3. Destination occupied
    dest.write_bytes(b"occupied")
    req_occ = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
    )
    rep = author_table_column(req_occ)
    assert rep.passed is False
    assert rep.code == "table_column_destination_occupied"
    assert dest.read_bytes() == b"occupied"  # Unmodified

    # 4. Destination equals source
    req_same = TableColumnRequest(
        source=source,
        destination=source,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
    )
    rep = author_table_column(req_same)
    assert rep.passed is False
    assert rep.code == "table_column_destination_invalid"


def test_author_table_column_refusals_target_and_fingerprint(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    discovery = discover_table_column_targets(source)
    t = discovery["targets"][0]

    # Target table out of range
    req_bad_table = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target={"table": 99, "column": 1},
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
    )
    assert author_table_column(req_bad_table).code == "table_column_target_missing"

    # Target column out of range
    req_bad_col = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target={"table": 1, "column": 99},
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
    )
    assert author_table_column(req_bad_col).code == "table_column_target_missing"

    # Stale fingerprint
    req_bad_fp = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256="f" * 64,
    )
    assert author_table_column(req_bad_fp).code == "table_column_structure_fingerprint_mismatch"


def test_author_table_column_refusals_hazards(tmp_path: Path) -> None:
    dest = tmp_path / "out.docx"
    table_grid = grid([2500, 2500])

    # 1. vMerge in table
    vmerge_body = table(
        row(cell(paragraph("A"), properties="<w:vMerge w:val='restart'/>") + cell(paragraph("B")))
        + row(cell(paragraph(""), properties="<w:vMerge/>") + cell(paragraph("D"))),
        grid=table_grid,
    )
    vmerge_pkg = write_package(tmp_path / "vm.docx", vmerge_body)
    req = TableColumnRequest(
        source=vmerge_pkg,
        destination=dest,
        source_sha256=_sha256(vmerge_pkg),
        target={"table": 1, "column": 1},
        structure_fingerprint_sha256="a" * 64,
    )
    assert author_table_column(req).code == "table_column_merged_unsupported"

    # 2. gridBefore in row
    gb_body = table(
        row(
            cell(paragraph("A")) + cell(paragraph("B")),
            properties="<w:gridBefore w:val='1'/>",
        )
        + row(cell(paragraph("C")) + cell(paragraph("D"))),
        grid=table_grid,
    )
    gb_pkg = write_package(tmp_path / "gb.docx", gb_body)
    req = TableColumnRequest(
        source=gb_pkg,
        destination=dest,
        source_sha256=_sha256(gb_pkg),
        target={"table": 1, "column": 1},
        structure_fingerprint_sha256="a" * 64,
    )
    assert author_table_column(req).code == "table_column_grid_before_after_unsupported"

    # 3. Non-rectangular (cardinality mismatch)
    nr_body = table(
        row(cell(paragraph("A")) + cell(paragraph("B")))
        + row(cell(paragraph("C"))),  # 1 cell instead of 2
        grid=table_grid,
    )
    nr_pkg = write_package(tmp_path / "nr.docx", nr_body)
    req = TableColumnRequest(
        source=nr_pkg,
        destination=dest,
        source_sha256=_sha256(nr_pkg),
        target={"table": 1, "column": 1},
        structure_fingerprint_sha256="a" * 64,
    )
    assert author_table_column(req).code == "table_column_grid_mismatch"

    # 4. Nested table hazard
    nested_tbl = table(row(cell(paragraph("inner"))), grid=grid([1000]))
    nested_body = table(
        row(cell(paragraph("A") + nested_tbl) + cell(paragraph("B"))),
        grid=table_grid,
    )
    nested_pkg = write_package(tmp_path / "nested.docx", nested_body)
    req = TableColumnRequest(
        source=nested_pkg,
        destination=dest,
        source_sha256=_sha256(nested_pkg),
        target={"table": 1, "column": 1},
        structure_fingerprint_sha256="a" * 64,
    )
    assert author_table_column(req).code == "table_column_structure_unsupported"


def test_author_table_column_refusals_payload(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    discovery = discover_table_column_targets(source)
    t = discovery["targets"][0]

    # Cell texts length mismatch (2 texts for 3 rows)
    req_mismatch = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
        cell_texts=("Only 1", "Only 2"),
    )
    assert author_table_column(req_mismatch).code == "table_column_cell_texts_mismatch"

    # Invalid XML text (null character) raises ValueError in request creation
    with pytest.raises(ValueError, match="table_column_cell_text_invalid"):
        TableColumnRequest(
            source=source,
            destination=dest,
            source_sha256=discovery["source_sha256"],
            target=t["target"],
            structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
            cell_texts=("ok", "bad \x00 text", "ok"),
        )

    # Invalid width (negative or 0)
    with pytest.raises(ValueError, match="table_column_width_invalid"):
        TableColumnRequest(
            source=source,
            destination=dest,
            source_sha256=discovery["source_sha256"],
            target=t["target"],
            structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
            width=-50,
        )

    # Invalid position
    with pytest.raises(ValueError, match="table_column_position_invalid"):
        TableColumnRequest(
            source=source,
            destination=dest,
            source_sha256=discovery["source_sha256"],
            target=t["target"],
            structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
            position="middle",
        )

    # Delete request with extra fields
    with pytest.raises(ValueError, match="table_column_request_fields_invalid"):
        TableColumnRequest(
            source=source,
            destination=dest,
            source_sha256=discovery["source_sha256"],
            target=t["target"],
            structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
            action="delete",
            position="after",
        )


def test_table_column_request_parsing_and_loading(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    discovery = discover_table_column_targets(source)
    t = discovery["targets"][0]

    raw = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": discovery["source_sha256"],
        "target": {"table": 1, "column": 1},
        "structure_fingerprint_sha256": t["structure_fingerprint_sha256"],
        "position": "after",
        "width": 2000,
        "cell_texts": ["A", "B", "C"],
        "action": "insert",
    }
    req = parse_table_column_request(raw)
    assert req.table == 1
    assert req.column == 1
    assert req.width == 2000
    assert req.cell_texts == ("A", "B", "C")

    # Round-trip to_dict
    serialized = req.to_dict()
    req2 = parse_table_column_request(serialized)
    assert req2.table == req.table
    assert req2.column == req.column
    assert req2.width == req.width
    assert req2.cell_texts == req.cell_texts

    # load_table_column_request from file
    req_file = tmp_path / "request.json"
    req_file.write_text(json.dumps(raw), encoding="utf-8")
    req3 = load_table_column_request(req_file)
    assert req3.table == 1
    assert req3.column == 1


def test_unrelated_parts_preservation(tmp_path: Path) -> None:
    # Package with custom styles and extra parts
    source = _make_sample_package(
        tmp_path / "source.docx",
        tbl_properties="",
    )
    dest = tmp_path / "dest.docx"
    discovery = discover_table_column_targets(source)
    t = discovery["targets"][0]

    req = TableColumnRequest(
        source=source,
        destination=dest,
        source_sha256=discovery["source_sha256"],
        target=t["target"],
        structure_fingerprint_sha256=t["structure_fingerprint_sha256"],
        position="after",
        width=1500,
        cell_texts=("X1", "X2", "X3"),
    )
    report = author_table_column(req)
    assert report.passed is True
    assert report.non_target_parts_preserved is True

    # Byte-level check on unrelated parts
    with zipfile.ZipFile(source) as src_zip, zipfile.ZipFile(dest) as dst_zip:
        for item in src_zip.infolist():
            if item.filename != _DOCUMENT:
                assert src_zip.read(item.filename) == dst_zip.read(item.filename)


def test_table_column_discovery_publication_and_occupied_output(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx", grid_widths=[2000, 3000, 4000])
    discovery = discover_table_column_targets(source)
    assert discovery["passed"] is True
    assert len(discovery["targets"]) == 3

    selected = discovery["targets"][0]
    destination = tmp_path / "published.docx"
    request = TableColumnRequest(
        source=source,
        destination=destination,
        source_sha256=discovery["source_sha256"],
        target=selected["target"],
        structure_fingerprint_sha256=selected["structure_fingerprint_sha256"],
        position="after",
        width=1500,
        cell_texts=("New H", "New D1", "New D2"),
    )
    request = parse_table_column_request(request.to_dict())

    published = author_table_column(request)
    assert published.passed is True
    assert destination.is_file()
    published_bytes = destination.read_bytes()

    occupied = author_table_column(request)
    assert occupied.passed is False
    assert occupied.code == "table_column_destination_occupied"
    assert destination.read_bytes() == published_bytes
