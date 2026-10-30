"""Tests for source-bound table insertion and anchor discovery."""

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
from wolfdocx.table_insert import (
    TableInsertReport,
    TableInsertRequest,
    TableInsertTarget,
    author_table_insert,
    discover_table_insert_targets,
    load_table_insert_request,
    parse_table_insert_request,
    table_insert_structure_fingerprint,
)
from wolfdocx.table_structure import discover_table_row_targets

_DOCUMENT = "word/document.xml"

_STYLES_WITH_TABLE_GRID = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/></w:style>'
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>'
    "</w:styles>"
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_sample_package(
    path: Path,
    *,
    paragraphs: list[str] | None = None,
    include_table: bool = False,
    styles_xml: str | None = None,
) -> Path:
    """Create a test package with known body paragraphs and optional table."""
    texts = (
        paragraphs
        if paragraphs is not None
        else ["First paragraph.", "Second paragraph.", "Third paragraph."]
    )
    parts: list[str] = [paragraph(t) for t in texts]
    if include_table:
        table_grid = grid([2500, 2500])
        rows_xml = [
            row(cell(paragraph("Cell 1")) + cell(paragraph("Cell 2"))),
            row(cell(paragraph("Cell 3")) + cell(paragraph("Cell 4"))),
        ]
        parts.append(table("".join(rows_xml), grid=table_grid))

    body = "".join(parts)
    extra_parts = {"word/styles.xml": styles_xml} if styles_xml is not None else None
    return write_package(path, body, parts=extra_parts)


def test_discover_targets_empty_document(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample.docx")
    source_sha = _sha256(source)

    discovery = discover_table_insert_targets(source)
    assert discovery["passed"] is True
    assert discovery["operation"] == "table-insert-targets"
    assert discovery["status"] == "discovered"
    assert discovery["code"] is None
    assert discovery["source_sha256"] == source_sha
    assert discovery["part"] == _DOCUMENT
    assert len(discovery["targets"]) == 3

    t1 = discovery["targets"][0]
    assert t1["paragraph"] == 1
    assert t1["text"] == "First paragraph."
    assert t1["empty"] is False
    assert t1["supported"] is True
    assert len(t1["structure_fingerprint_sha256"]) == 64
    assert t1["target"] == {"part": _DOCUMENT, "paragraph": 1}

    target_obj = TableInsertTarget(
        part=_DOCUMENT,
        paragraph=1,
        text="First paragraph.",
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
    )
    assert target_obj.to_dict() == t1


def test_discover_targets_unsupported_part(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample.docx")
    discovery = discover_table_insert_targets(source, part="word/header1.xml")
    assert discovery["passed"] is False
    assert discovery["status"] == "refused"
    assert discovery["code"] == "table_insert_part_unsupported"
    assert discovery["targets"] == []


def test_discover_targets_missing_file(tmp_path: Path) -> None:
    discovery = discover_table_insert_targets(tmp_path / "nonexistent.docx")
    assert discovery["passed"] is False
    assert discovery["status"] == "refused"
    assert discovery["code"] == "table_insert_source_missing"
    assert discovery["targets"] == []


def test_author_table_insert_after_paragraph(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "output.docx"

    discovery = discover_table_insert_targets(source)
    target = discovery["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="after",
        rows=[
            ["H1", "H2"],
            ["D1", "D2"],
        ],
        grid_widths=[3500, 4500],
    )

    report = author_table_insert(request)
    assert isinstance(report, TableInsertReport)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == source_sha
    assert report.destination_sha256 == _sha256(destination)
    assert report.non_target_parts_preserved is True
    assert report.target == {"part": _DOCUMENT, "paragraph": 1, "position": "after"}
    assert report.inserted_table is not None
    assert report.inserted_table["table_index"] == 1
    assert report.inserted_table["rows"] == 2
    assert report.inserted_table["cols"] == 2
    assert report.inserted_table["grid_widths"] == [3500, 4500]

    # Verify destination document XML structure, preservation, and adjacency
    graph = PackageGraph.from_file(destination)
    doc_bytes = graph.read_part_bytes(_DOCUMENT)
    root = ET.fromstring(doc_bytes)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None

    children = list(body)
    # Expected order: p1 ("First paragraph."), tbl, p2 ("Second paragraph."), p3 ("Third paragraph."), sectPr
    tags = [c.tag.rsplit("}", 1)[-1] for c in children]
    assert tags == ["p", "tbl", "p", "p", "sectPr"]

    p1 = children[0]
    assert "".join(t.text or "" for t in p1.iter(Q(ns, "t"))) == "First paragraph."

    tbl = children[1]
    tbl_rows = tbl.findall(Q(ns, "tr"))
    assert len(tbl_rows) == 2

    # Check cell texts and widths
    r0_cells = tbl_rows[0].findall(Q(ns, "tc"))
    assert len(r0_cells) == 2
    assert "".join(t.text or "" for t in r0_cells[0].iter(Q(ns, "t"))) == "H1"
    assert "".join(t.text or "" for t in r0_cells[1].iter(Q(ns, "t"))) == "H2"
    assert r0_cells[0].find(Q(ns, "tcPr")).find(Q(ns, "tcW")).get(Q(ns, "w")) == "3500"
    assert r0_cells[1].find(Q(ns, "tcPr")).find(Q(ns, "tcW")).get(Q(ns, "w")) == "4500"

    r1_cells = tbl_rows[1].findall(Q(ns, "tc"))
    assert len(r1_cells) == 2
    assert "".join(t.text or "" for t in r1_cells[0].iter(Q(ns, "t"))) == "D1"
    assert "".join(t.text or "" for t in r1_cells[1].iter(Q(ns, "t"))) == "D2"


def test_author_table_insert_before_paragraph(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "output.docx"

    discovery = discover_table_insert_targets(source)
    target = discovery["targets"][1]  # Second paragraph

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="before",
        rows=[["Alpha", "Beta"]],
        grid_widths=[4000, 4000],
    )

    report = author_table_insert(request)
    assert report.passed is True
    assert destination.is_file()

    graph = PackageGraph.from_file(destination)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None
    tags = [c.tag.rsplit("}", 1)[-1] for c in body]
    # Before second paragraph: p1, tbl, p2, p3, sectPr
    assert tags == ["p", "tbl", "p", "p", "sectPr"]


def test_author_table_insert_document_with_no_tables_rediscovery(
    tmp_path: Path,
) -> None:
    source = _make_sample_package(tmp_path / "no_tables.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "has_table.docx"

    initial_inspection = inspect_docx(source)
    assert initial_inspection.valid is True
    assert len(initial_inspection.fingerprint.tables) == 0

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        position="after",
        rows=[
            ["Metric", "Value", "Notes"],
            ["Throughput", "95.5", "Nominal"],
            ["Latency", "12.2", "P99"],
        ],
        grid_widths=[3000, 2500, 3500],
    )

    report = author_table_insert(request)
    assert report.passed is True

    # Validate with inspect_docx using actual public inspection schema
    post_inspection = inspect_docx(destination)
    assert post_inspection.valid is True
    assert len(post_inspection.fingerprint.tables) == 1

    # Validate that table_structure can now discover prototype rows
    row_discovery = discover_table_row_targets(destination)
    assert row_discovery["passed"] is True
    assert len(row_discovery["targets"]) == 3
    assert row_discovery["targets"][0]["cell_texts"] == ["Metric", "Value", "Notes"]


def test_author_table_insert_with_table_style(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "styled.docx", styles_xml=_STYLES_WITH_TABLE_GRID)
    source_sha = _sha256(source)
    destination = tmp_path / "styled_out.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[["Item", "Price"], ["Apple", "$1.00"]],
        grid_widths=[4000, 4000],
        style="TableGrid",
    )

    report = author_table_insert(request)
    assert report.passed is True
    assert report.inserted_table["style"] == "TableGrid"

    graph = PackageGraph.from_file(destination)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl_pr = root.find(f".//{Q(ns, 'tblPr')}")
    assert tbl_pr is not None
    style_elem = tbl_pr.find(Q(ns, "tblStyle"))
    assert style_elem is not None
    assert style_elem.get(Q(ns, "val")) == "TableGrid"


def test_author_table_insert_missing_style_refused(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "styled.docx", styles_xml=_STYLES_WITH_TABLE_GRID)
    source_sha = _sha256(source)
    destination = tmp_path / "missing_style_out.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[["A", "B"]],
        grid_widths=[3000, 3000],
        style="NonExistentStyle",
    )

    report = author_table_insert(request)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_style_missing"
    assert not destination.exists()


def test_author_table_insert_incompatible_style_refused(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "styled.docx", styles_xml=_STYLES_WITH_TABLE_GRID)
    source_sha = _sha256(source)
    destination = tmp_path / "incompatible_style_out.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    # "Heading" is a paragraph style, not a table style
    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[["A", "B"]],
        grid_widths=[3000, 3000],
        style="Heading",
    )

    report = author_table_insert(request)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_properties_style_type_incompatible"
    assert not destination.exists()


def test_author_table_insert_repeat_header(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "repeat_header.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[["H1", "H2"], ["D1", "D2"]],
        grid_widths=[4000, 4000],
        repeat_header=True,
    )

    report = author_table_insert(request)
    assert report.passed is True
    assert report.inserted_table["repeat_header"] is True

    graph = PackageGraph.from_file(destination)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = root.find(f".//{Q(ns, 'tbl')}")
    assert tbl is not None
    rows = tbl.findall(Q(ns, "tr"))
    # First row has tblHeader
    r0_pr = rows[0].find(Q(ns, "trPr"))
    assert r0_pr is not None
    assert r0_pr.find(Q(ns, "tblHeader")) is not None
    # Second row does not have tblHeader
    r1_pr = rows[1].find(Q(ns, "trPr"))
    assert r1_pr is None or r1_pr.find(Q(ns, "tblHeader")) is None


def test_author_table_insert_xml_text_escaping(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "escaping.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    special_text_1 = "Tom & Jerry <cartoon> \"classic\" '1940'"
    special_text_2 = "  leading and trailing spaces  "

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[[special_text_1, special_text_2]],
        grid_widths=[4500, 4500],
    )

    report = author_table_insert(request)
    assert report.passed is True

    # Read back XML and verify preservation
    graph = PackageGraph.from_file(destination)
    raw_xml = graph.read_part_bytes(_DOCUMENT).decode("utf-8")
    assert "&amp;" in raw_xml
    assert "&lt;cartoon&gt;" in raw_xml

    root = ET.fromstring(raw_xml.encode("utf-8"))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    cells = root.findall(f".//{Q(ns, 'tc')}")
    assert len(cells) == 2

    txt1 = "".join(t.text or "" for t in cells[0].iter(Q(ns, "t")))
    assert txt1 == special_text_1

    txt2 = "".join(t.text or "" for t in cells[1].iter(Q(ns, "t")))
    assert txt2 == special_text_2

    t2 = next(iter(cells[1].iter(Q(ns, "t"))))
    assert t2.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"


def test_author_table_insert_occupied_destination_refused(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "occupied.docx"
    destination.write_bytes(b"EXISTING_CONTENT")

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[["A", "B"]],
        grid_widths=[3000, 3000],
    )

    report = author_table_insert(request)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_insert_destination_occupied"
    assert destination.read_bytes() == b"EXISTING_CONTENT"


def test_author_table_insert_source_hash_mismatch_refused(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    destination = tmp_path / "output.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256="0" * 64,  # Stale hash
        target=target["target"],
        structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
        rows=[["A", "B"]],
        grid_widths=[3000, 3000],
    )

    report = author_table_insert(request)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_insert_source_hash_mismatch"
    assert not destination.exists()


def test_author_table_insert_fingerprint_mismatch_refused(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    destination = tmp_path / "output.docx"

    target = discover_table_insert_targets(source)["targets"][0]

    request = TableInsertRequest(
        source=source,
        destination=destination,
        source_sha256=source_sha,
        target=target["target"],
        structure_fingerprint_sha256="f" * 64,  # Mismatched fingerprint
        rows=[["A", "B"]],
        grid_widths=[3000, 3000],
    )

    report = author_table_insert(request)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "table_insert_structure_fingerprint_mismatch"
    assert not destination.exists()


def test_author_table_insert_non_rectangular_rows_refused(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    target = discover_table_insert_targets(source)["targets"][0]

    with pytest.raises(ValueError, match="table_insert_rows_non_rectangular"):
        TableInsertRequest(
            source=source,
            destination=tmp_path / "out.docx",
            source_sha256=source_sha,
            target=target["target"],
            structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
            rows=[["A", "B"], ["C"]],
            grid_widths=[3000, 3000],
        )


def test_author_table_insert_grid_widths_count_mismatch(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    target = discover_table_insert_targets(source)["targets"][0]

    # Construction validation: 2 cols, 3 widths
    with pytest.raises(ValueError, match="table_insert_grid_widths_count_mismatch"):
        TableInsertRequest(
            source=source,
            destination=tmp_path / "out.docx",
            source_sha256=source_sha,
            target=target["target"],
            structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
            rows=[["A", "B"]],
            grid_widths=[3000, 3000, 3000],
        )


def test_author_table_insert_grid_widths_invalid_cases(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    target = discover_table_insert_targets(source)["targets"][0]

    # Empty grid_widths
    with pytest.raises(ValueError, match="table_insert_grid_widths_invalid"):
        TableInsertRequest(
            source=source,
            destination=tmp_path / "out.docx",
            source_sha256=source_sha,
            target=target["target"],
            structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
            rows=[["A", "B"]],
            grid_widths=[],
        )

    # Non-positive grid_widths
    with pytest.raises(ValueError, match="table_insert_grid_widths_invalid"):
        TableInsertRequest(
            source=source,
            destination=tmp_path / "out.docx",
            source_sha256=source_sha,
            target=target["target"],
            structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
            rows=[["A", "B"]],
            grid_widths=[3000, 0],
        )


def test_parse_and_load_table_insert_request(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)
    target = discover_table_insert_targets(source)["targets"][0]

    request_payload = {
        "source": str(source),
        "destination": str(tmp_path / "parsed_out.docx"),
        "source_sha256": source_sha,
        "target": target["target"],
        "structure_fingerprint_sha256": target["structure_fingerprint_sha256"],
        "position": "after",
        "rows": [["X", "Y"], ["1", "2"]],
        "grid_widths": [3000, 4000],
        "repeat_header": True,
    }

    parsed = parse_table_insert_request(request_payload)
    assert parsed.paragraph == 1
    assert parsed.part == _DOCUMENT
    assert parsed.position == "after"
    assert parsed.rows == (("X", "Y"), ("1", "2"))
    assert parsed.grid_widths == (3000, 4000)
    assert parsed.repeat_header is True

    json_file = tmp_path / "request.json"
    json_file.write_text(json.dumps(request_payload), encoding="utf-8")
    loaded = load_table_insert_request(json_file)
    assert loaded.to_dict() == parsed.to_dict()

    # Invalid schema: non-mapping
    with pytest.raises(TypeError):
        parse_table_insert_request("not-a-mapping")

    # Unknown / conflicting keys
    with pytest.raises(ValueError, match="table_insert_request_fields_invalid"):
        parse_table_insert_request(dict(request_payload, table_data=[["A"]]))

    with pytest.raises(ValueError, match="table_insert_request_fields_invalid"):
        parse_table_insert_request(dict(request_payload, column_widths=[3000, 4000]))

    with pytest.raises(ValueError, match="table_insert_request_fields_invalid"):
        parse_table_insert_request(dict(request_payload, unknown_field="invalid"))

    # Missing required grid_widths
    missing_widths = dict(request_payload)
    del missing_widths["grid_widths"]
    with pytest.raises(ValueError, match="table_insert_request_fields_invalid"):
        parse_table_insert_request(missing_widths)

    # Invalid position
    with pytest.raises(ValueError, match="table_insert_position_invalid"):
        parse_table_insert_request(dict(request_payload, position="middle"))


def test_table_insert_structure_fingerprint(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    fp1 = table_insert_structure_fingerprint(source, paragraph=1)
    assert fp1 is not None and len(fp1) == 64

    # Out of range paragraph returns None
    assert table_insert_structure_fingerprint(source, paragraph=999) is None

    # Unsupported part returns None
    assert table_insert_structure_fingerprint(source, paragraph=1, part="word/header1.xml") is None

    # Missing file returns None
    assert table_insert_structure_fingerprint(tmp_path / "missing.docx", paragraph=1) is None
