"""Focused tests for source-bound DOCX body table deletion and target discovery."""

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
from wolfdocx.inspect import Q, inspect_docx
from wolfdocx.package import PackageGraph
from wolfdocx.table_delete import (
    TableDeleteReport,
    TableDeleteRequest,
    TableDeleteTarget,
    author_table_delete,
    discover_table_delete_targets,
    load_table_delete_request,
    parse_table_delete_request,
    table_delete_structure_fingerprint,
)

_DOCUMENT = "word/document.xml"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_table_doc(
    path: Path,
    *,
    paragraphs_before: list[str] | None = None,
    tables: list[tuple[int, int]] | None = None,
    paragraphs_after: list[str] | None = None,
    extra_parts: dict[str, str | bytes] | None = None,
    extra_rels: str = "",
    sect_properties: str | None = None,
) -> Path:
    """Create a test package with configurable paragraphs and rectangular tables."""
    parts: list[str] = []
    before = (
        paragraphs_before
        if paragraphs_before is not None
        else ["Leading paragraph one.", "Leading paragraph two."]
    )
    for p_text in before:
        parts.append(paragraph(p_text))

    table_specs = tables if tables is not None else [(2, 2)]
    for rows_cnt, cols_cnt in table_specs:
        col_widths = [2000] * cols_cnt
        t_grid = grid(col_widths)
        t_rows: list[str] = []
        for r_idx in range(rows_cnt):
            t_cells = [cell(paragraph(f"R{r_idx + 1}C{c_idx + 1}")) for c_idx in range(cols_cnt)]
            t_rows.append(row("".join(t_cells)))
        parts.append(table("".join(t_rows), grid=t_grid))

    after = paragraphs_after if paragraphs_after is not None else ["Trailing paragraph one."]
    for p_text in after:
        parts.append(paragraph(p_text))

    body = "".join(parts)
    return write_package(
        path,
        body,
        parts=extra_parts,
        relationships=extra_rels,
        sect_properties=sect_properties,
    )


def test_discover_table_delete_targets(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "multi_table.docx", tables=[(2, 2), (3, 2)])
    source_sha = _sha256(source)

    discovery = discover_table_delete_targets(source)
    assert discovery["passed"] is True
    assert discovery["operation"] == "table-delete-targets"
    assert discovery["status"] == "discovered"
    assert discovery["code"] is None
    assert discovery["source_sha256"] == source_sha
    assert discovery["part"] == _DOCUMENT

    targets = discovery["targets"]
    assert len(targets) == 2

    t1 = targets[0]
    assert t1["table"] == 1
    assert t1["rows_count"] == 2
    assert t1["columns_count"] == 2
    assert "preview_text" in t1
    assert len(t1["structure_fingerprint_sha256"]) == 64

    target_obj = TableDeleteTarget(
        table=t1["table"],
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
        part=t1["part"],
        rows_count=t1["rows_count"],
        columns_count=t1["columns_count"],
        style=t1["style"],
        preview_text=t1["preview_text"],
    )
    assert target_obj.to_dict() == t1

    t2 = targets[1]
    assert t2["table"] == 2
    assert t2["rows_count"] == 3
    assert t2["columns_count"] == 2

    # Structural fingerprint helper matches discovered fingerprints
    fp1 = table_delete_structure_fingerprint(source, table=1)
    fp2 = table_delete_structure_fingerprint(source, table=2)
    assert fp1 == t1["structure_fingerprint_sha256"]
    assert fp2 == t2["structure_fingerprint_sha256"]


def test_discover_table_delete_targets_refusals(tmp_path: Path) -> None:
    # 1. Missing source
    res_missing = discover_table_delete_targets(tmp_path / "nonexistent.docx")
    assert res_missing["passed"] is False
    assert res_missing["code"] == "table_delete_source_missing"

    # 2. Corrupt source
    bad_file = tmp_path / "corrupt.docx"
    bad_file.write_bytes(b"not a zip file")
    res_corrupt = discover_table_delete_targets(bad_file)
    assert res_corrupt["passed"] is False
    assert res_corrupt["code"] == "table_delete_package_invalid"

    # 3. Unsupported part
    source = _make_table_doc(tmp_path / "sample.docx")
    res_part = discover_table_delete_targets(source, part="word/header1.xml")
    assert res_part["passed"] is False
    assert res_part["code"] == "table_delete_part_unsupported"


def test_table_delete_structure_fingerprint_edge_cases(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "fp_doc.docx", tables=[(2, 2)])

    # Valid table
    fp = table_delete_structure_fingerprint(source, table=1)
    assert fp is not None and len(fp) == 64

    # Out of bounds
    assert table_delete_structure_fingerprint(source, table=0) is None
    assert table_delete_structure_fingerprint(source, table=2) is None

    # Unsupported part
    assert table_delete_structure_fingerprint(source, table=1, part="word/footer.xml") is None

    # Missing file
    assert table_delete_structure_fingerprint(tmp_path / "absent.docx", table=1) is None


def test_author_table_delete_single_table_success(tmp_path: Path) -> None:
    source = _make_table_doc(
        tmp_path / "single_table.docx",
        paragraphs_before=["Intro 1", "Intro 2"],
        tables=[(2, 3)],
        paragraphs_after=["Outro 1"],
    )
    source_sha = _sha256(source)
    destination = tmp_path / "single_table_deleted.docx"

    fp = table_delete_structure_fingerprint(source, table=1)
    assert fp is not None

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1, "part": _DOCUMENT},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == source_sha
    assert report.destination_sha256 == _sha256(destination)
    assert destination.is_file()

    # Verify target change report and deleted table metrics
    assert report.deleted_table is not None
    assert report.deleted_table["table_index"] == 1
    assert report.deleted_table["rows_deleted"] == 2
    assert report.deleted_table["columns_deleted"] == 3
    assert report.deleted_table["remaining_tables"] == 0
    assert report.non_target_parts_preserved is True

    # Read back and prove table count delta and paragraph preservation
    s_graph = PackageGraph.from_file(destination)
    s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
    s_tables = list(s_root.iter(Q(_W, "tbl")))
    assert len(s_tables) == 0

    s_body = s_root.find(Q(_W, "body"))
    assert s_body is not None
    s_body_paragraphs = [c for c in s_body if c.tag == Q(_W, "p")]
    assert len(s_body_paragraphs) == 3

    p_texts = ["".join(t.text or "" for t in p.iter(Q(_W, "t"))) for p in s_body_paragraphs]
    assert p_texts == ["Intro 1", "Intro 2", "Outro 1"]

    # Package inspection valid
    insp = inspect_docx(destination)
    assert insp.valid is True


def test_author_table_delete_middle_table_success(tmp_path: Path) -> None:
    source = _make_table_doc(
        tmp_path / "three_tables.docx",
        paragraphs_before=["Head"],
        tables=[(1, 2), (2, 2), (3, 2)],
        paragraphs_after=["Tail"],
    )
    source_sha = _sha256(source)
    destination = tmp_path / "middle_deleted.docx"

    # Delete table 2 (the middle table)
    fp2 = table_delete_structure_fingerprint(source, table=2)
    assert fp2 is not None

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 2},
        structure_fingerprint_sha256=fp2,
    )

    report = author_table_delete(request)
    assert report.passed is True
    assert report.code is None
    assert report.deleted_table["table_index"] == 2
    assert report.deleted_table["rows_deleted"] == 2
    assert report.deleted_table["remaining_tables"] == 2

    # Verification of remaining tables
    s_graph = PackageGraph.from_file(destination)
    s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
    s_tables = list(s_root.iter(Q(_W, "tbl")))
    assert len(s_tables) == 2

    # First remaining table has 1 row; second has 3 rows
    assert len(s_tables[0].findall(Q(_W, "tr"))) == 1
    assert len(s_tables[1].findall(Q(_W, "tr"))) == 3


def test_author_table_delete_refusal_destination_occupied(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "occupied.docx"
    destination.write_text("existing content", encoding="utf-8")
    occupied_sha_before = _sha256(destination)

    fp = table_delete_structure_fingerprint(source, table=1)
    assert fp is not None

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_delete_destination_occupied"

    # Occupied destination remains completely unaltered
    assert _sha256(destination) == occupied_sha_before
    assert destination.read_text(encoding="utf-8") == "existing content"


def test_author_table_delete_refusal_source_hash_mismatch(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    destination = tmp_path / "out.docx"
    fp = table_delete_structure_fingerprint(source, table=1)
    assert fp is not None

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256="0" * 64,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_source_hash_mismatch"
    assert not destination.exists()


def test_author_table_delete_refusal_source_missing(tmp_path: Path) -> None:
    source = tmp_path / "missing.docx"
    destination = tmp_path / "out.docx"

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256="a" * 64,
        target={"table": 1},
        structure_fingerprint_sha256="b" * 64,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_source_missing"
    assert not destination.exists()


def test_author_table_delete_refusal_target_missing(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx", tables=[(1, 1)])
    source_sha = _sha256(source)
    destination = tmp_path / "out.docx"
    fp = table_delete_structure_fingerprint(source, table=1)
    assert fp is not None

    # Table 5 does not exist (only 1 table in document)
    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 5},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_target_missing"
    assert not destination.exists()


def test_author_table_delete_refusal_target_stale(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx", tables=[(2, 2)])
    source_sha = _sha256(source)
    destination = tmp_path / "out.docx"

    # Mismatched fingerprint
    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256="f" * 64,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_target_stale"
    assert not destination.exists()


def test_author_table_delete_refusal_nested_table(tmp_path: Path) -> None:
    # Construct a document where table 2 is nested inside a cell of table 1
    nested_table_xml = table(row(cell(paragraph("Nested Cell"))), grid=grid([1000]))
    outer_cell_with_nested = cell(paragraph("Outer Cell") + nested_table_xml)
    outer_table_xml = table(row(outer_cell_with_nested), grid=grid([2000]))
    body = paragraph("Intro") + outer_table_xml + paragraph("Outro")

    source = tmp_path / "nested_doc.docx"
    write_package(source, body)
    source_sha = _sha256(source)
    destination = tmp_path / "nested_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    assert len(tables) == 2

    # 1. Attempting to delete the nested table directly (table 2)
    fp_t2 = _table_structure_fingerprint(tables[1])
    req2 = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 2},
        structure_fingerprint_sha256=fp_t2,
    )
    rep2 = author_table_delete(req2)
    assert rep2.passed is False
    assert rep2.code == "table_delete_target_nested"
    assert not destination.exists()

    # 2. Attempting to delete the outer table (table 1) which contains a nested table
    fp_t1 = _table_structure_fingerprint(tables[0])
    req1 = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp_t1,
    )
    rep1 = author_table_delete(req1)
    assert rep1.passed is False
    assert rep1.code == "table_delete_target_nested"
    assert not destination.exists()


def test_author_table_delete_refusal_section_hazard_sectpr_in_table(
    tmp_path: Path,
) -> None:
    # Construct a table containing an illegal w:sectPr inside a cell
    sect_in_cell = "<w:pPr><w:sectPr/></w:pPr>"
    bad_cell = cell(f"<w:p>{sect_in_cell}</w:p>")
    bad_table_xml = table(row(bad_cell), grid=grid([2000]))
    body = paragraph("Intro") + bad_table_xml + paragraph("Outro")

    source = tmp_path / "sect_hazard.docx"
    write_package(source, body)
    source_sha = _sha256(source)
    destination = tmp_path / "sect_hazard_out.docx"

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    from wolfdocx.table_delete import _table_structure_fingerprint

    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_section_hazard"
    assert not destination.exists()


def test_author_table_delete_refusal_section_hazard_empty_body(tmp_path: Path) -> None:
    # A document where table is the only content (no body paragraphs left after deletion)
    tbl_xml = table(row(cell(paragraph("Only cell"))), grid=grid([2000]))
    body = tbl_xml

    source = tmp_path / "only_table.docx"
    write_package(source, body)
    source_sha = _sha256(source)
    destination = tmp_path / "only_table_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_section_hazard"
    assert not destination.exists()


def test_author_table_delete_refusal_bookmark_intersection(tmp_path: Path) -> None:
    # Bookmark start before table, bookmark end inside table
    bm_start = '<w:bookmarkStart w:id="10" w:name="CrossBookmark"/>'
    bm_end = '<w:bookmarkEnd w:id="10"/>'

    p_before = f"<w:p>{bm_start}<w:r><w:t>Before</w:t></w:r></w:p>"
    c_inside = cell(f"<w:p><w:r><w:t>Inside</w:t></w:r>{bm_end}</w:p>")
    tbl_xml = table(row(c_inside), grid=grid([2000]))
    p_after = paragraph("After")

    source = tmp_path / "bm_cross.docx"
    write_package(source, p_before + tbl_xml + p_after)
    source_sha = _sha256(source)
    destination = tmp_path / "bm_cross_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_bookmark_intersection_unsupported"
    assert not destination.exists()


def test_author_table_delete_refusal_bookmark_enclosing_table(tmp_path: Path) -> None:
    # Bookmark starts before table and ends after table (encloses table)
    bm_start = '<w:bookmarkStart w:id="20" w:name="EnclosingBookmark"/>'
    bm_end = '<w:bookmarkEnd w:id="20"/>'

    p_before = f"<w:p>{bm_start}<w:r><w:t>Before</w:t></w:r></w:p>"
    tbl_xml = table(row(cell(paragraph("Inside"))), grid=grid([2000]))
    p_after = f"<w:p><w:r><w:t>After</w:t></w:r>{bm_end}</w:p>"

    source = tmp_path / "bm_enclosing.docx"
    write_package(source, p_before + tbl_xml + p_after)
    source_sha = _sha256(source)
    destination = tmp_path / "bm_enclosing_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_bookmark_intersection_unsupported"
    assert not destination.exists()


def test_author_table_delete_refusal_bookmark_inside_referenced_outside(
    tmp_path: Path,
) -> None:
    # Bookmark start and end are inside table, but referenced outside via hyperlink anchor
    bm_start = '<w:bookmarkStart w:id="30" w:name="TableTarget"/>'
    bm_end = '<w:bookmarkEnd w:id="30"/>'
    c_inside = cell(f"<w:p>{bm_start}<w:r><w:t>Inside</w:t></w:r>{bm_end}</w:p>")
    tbl_xml = table(row(c_inside), grid=grid([2000]))

    # Outer paragraph contains hyperlink to bookmark
    hl_outside = (
        '<w:p><w:hyperlink w:anchor="TableTarget"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>'
    )
    body = hl_outside + tbl_xml + paragraph("After")

    source = tmp_path / "bm_ref_outside.docx"
    write_package(source, body)
    source_sha = _sha256(source)
    destination = tmp_path / "bm_ref_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_bookmark_intersection_unsupported"
    assert not destination.exists()


def test_author_table_delete_refusal_annotation_hazard(tmp_path: Path) -> None:
    # Table containing comment reference
    comment_ref = (
        '<w:commentRangeStart w:id="5"/><w:commentReference w:id="5"/><w:commentRangeEnd w:id="5"/>'
    )
    c_comment = cell(f"<w:p><w:r><w:t>Commented</w:t></w:r>{comment_ref}</w:p>")
    tbl_xml = table(row(c_comment), grid=grid([2000]))
    body = paragraph("Before") + tbl_xml + paragraph("After")

    source = tmp_path / "annot_hazard.docx"
    write_package(source, body)
    source_sha = _sha256(source)
    destination = tmp_path / "annot_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_annotation_intersection_unsupported"
    assert not destination.exists()


def test_author_table_delete_relationship_handling(tmp_path: Path) -> None:
    # 1. Shared relationship: rIdShared is referenced inside table AND outside table
    # Deleting table must NOT delete rIdShared from document.xml.rels
    hl_shared_table = (
        '<w:p><w:hyperlink r:id="rIdShared"><w:r><w:t>Table Link</w:t></w:r></w:hyperlink></w:p>'
    )
    hl_shared_outside = (
        '<w:p><w:hyperlink r:id="rIdShared"><w:r><w:t>Outside Link</w:t></w:r></w:hyperlink></w:p>'
    )

    tbl_xml = table(row(cell(hl_shared_table)), grid=grid([2000]))
    body = hl_shared_outside + tbl_xml + paragraph("Trailing")

    rels_extra = (
        f'<Relationship Id="rIdShared" Type="{_R}/hyperlink" '
        f'Target="https://example.com/shared" TargetMode="External"/>'
    )

    source = tmp_path / "shared_rel.docx"
    write_package(source, body, relationships=rels_extra)
    source_sha = _sha256(source)
    destination = tmp_path / "shared_rel_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is True
    assert report.code is None

    # Check that rIdShared is still present in document.xml.rels
    s_graph = PackageGraph.from_file(destination)
    rels_bytes = s_graph.read_part_bytes("word/_rels/document.xml.rels")
    assert b"rIdShared" in rels_bytes
    assert b"https://example.com/shared" in rels_bytes

    # 2. Solely-owned external relationship: rIdSole is referenced ONLY inside table
    # Deleting table safely prunes rIdSole from document.xml.rels
    hl_sole_table = (
        '<w:p><w:hyperlink r:id="rIdSole"><w:r><w:t>Sole Link</w:t></w:r></w:hyperlink></w:p>'
    )
    tbl_sole = table(row(cell(hl_sole_table)), grid=grid([2000]))
    body_sole = paragraph("Leading") + tbl_sole + paragraph("Trailing")

    rels_sole = (
        f'<Relationship Id="rIdSole" Type="{_R}/hyperlink" '
        f'Target="https://example.com/sole" TargetMode="External"/>'
    )

    source_sole = tmp_path / "sole_rel.docx"
    write_package(source_sole, body_sole, relationships=rels_sole)
    source_sole_sha = _sha256(source_sole)
    dest_sole = tmp_path / "sole_rel_out.docx"

    graph_sole = PackageGraph.from_file(source_sole)
    root_sole = ET.fromstring(graph_sole.read_part_bytes(_DOCUMENT))
    tables_sole = list(root_sole.iter(Q(_W, "tbl")))
    fp_sole = _table_structure_fingerprint(tables_sole[0])

    req_sole = TableDeleteRequest(
        source=source_sole,
        destination=dest_sole,
        source_sha256=source_sole_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp_sole,
    )

    rep_sole = author_table_delete(req_sole)
    assert rep_sole.passed is True
    assert rep_sole.code is None

    s_graph_sole = PackageGraph.from_file(dest_sole)
    rels_sole_bytes = s_graph_sole.read_part_bytes("word/_rels/document.xml.rels")
    assert b"rIdSole" not in rels_sole_bytes


def test_author_table_delete_refusal_relationship_hazard_missing_rel(
    tmp_path: Path,
) -> None:
    # Table references rIdMissing which is NOT defined in document.xml.rels
    hl_missing = (
        '<w:p><w:hyperlink r:id="rIdMissing"><w:r><w:t>Bad Link</w:t></w:r></w:hyperlink></w:p>'
    )
    tbl_xml = table(row(cell(hl_missing)), grid=grid([2000]))
    body = paragraph("Leading") + tbl_xml + paragraph("Trailing")

    source = tmp_path / "missing_rel.docx"
    write_package(source, body)
    source_sha = _sha256(source)
    destination = tmp_path / "missing_rel_out.docx"

    from wolfdocx.table_delete import _table_structure_fingerprint

    graph = PackageGraph.from_file(source)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    tables = list(root.iter(Q(_W, "tbl")))
    fp = _table_structure_fingerprint(tables[0])

    request = TableDeleteRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target={"table": 1},
        structure_fingerprint_sha256=fp,
    )

    report = author_table_delete(request)
    assert report.passed is False
    assert report.code == "table_delete_relationship_hazard"
    assert not destination.exists()


def test_table_delete_discovery_publication_and_occupied_output(tmp_path: Path) -> None:
    source = _make_table_doc(tmp_path / "source.docx")
    discovery = discover_table_delete_targets(source)
    selected = discovery["targets"][0]
    destination = tmp_path / "deleted.docx"
    request = parse_table_delete_request(
        TableDeleteRequest(
            source=source,
            destination=destination,
            source_sha256=discovery["source_sha256"],
            target={"table": selected["table"], "part": _DOCUMENT},
            structure_fingerprint_sha256=selected["structure_fingerprint_sha256"],
        ).to_dict()
    )
    assert author_table_delete(request).passed is True
    published = destination.read_bytes()
    assert author_table_delete(request).code == "table_delete_destination_occupied"
    assert destination.read_bytes() == published


def test_parse_and_load_table_delete_request(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    destination = tmp_path / "dest.docx"

    raw_payload = {
        "source": str(source),
        "destination": str(destination),
        "source_sha256": "a" * 64,
        "target": {"table": 2, "part": "word/document.xml"},
        "structure_fingerprint_sha256": "b" * 64,
    }

    req = parse_table_delete_request(raw_payload)
    assert req.source == source
    assert req.destination == destination
    assert req.table == 2
    assert req.part == "word/document.xml"
    assert req.to_dict() == raw_payload

    # Save to JSON file and test load_table_delete_request
    req_file = tmp_path / "request.json"
    req_file.write_text(json.dumps(raw_payload), encoding="utf-8")

    loaded_req = load_table_delete_request(req_file)
    assert loaded_req == req

    # Validation errors
    with pytest.raises(TypeError):
        parse_table_delete_request("not-a-dict")

    # Extra disallowed field
    with pytest.raises(ValueError):
        parse_table_delete_request({**raw_payload, "extra": "invalid"})

    # Invalid table index (negative)
    with pytest.raises(ValueError):
        parse_table_delete_request({**raw_payload, "target": {"table": -1}})

    # Invalid table index (boolean)
    with pytest.raises(ValueError):
        parse_table_delete_request({**raw_payload, "target": {"table": True}})

    # Invalid SHA256 length
    with pytest.raises(ValueError):
        parse_table_delete_request({**raw_payload, "source_sha256": "too-short"})


def test_table_delete_report_contract() -> None:
    report = TableDeleteReport(
        status="published",
        code=None,
        source_sha256="1" * 64,
        destination_sha256="2" * 64,
        target={"table": 1, "part": "word/document.xml"},
        deleted_table={"table_index": 1, "rows_deleted": 2},
        target_change_report={"table_index": 1, "rows_deleted": 2},
        non_target_parts_preserved=True,
    )

    assert report.passed is True
    data = report.to_dict()
    assert data["schema_version"] == 1
    assert data["operation"] == "table-delete"
    assert data["passed"] is True
    assert data["status"] == "published"
    assert data["code"] is None
    assert data["source_sha256"] == "1" * 64
    assert data["destination_sha256"] == "2" * 64
    assert data["target"] == {"table": 1, "part": "word/document.xml"}
    assert data["deleted_table"] == {"table_index": 1, "rows_deleted": 2}
    assert data["non_target_parts_preserved"] is True
