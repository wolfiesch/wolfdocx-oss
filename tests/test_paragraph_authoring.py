"""Tests for source-bound body paragraph authoring and target discovery."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import (
    R,
    W,
    cell,
    footer_reference,
    grid,
    header_reference,
    paragraph,
    region_relationship,
    row,
    sect,
    table,
    write_package,
)
from wolfdocx import review
from wolfdocx.inspect import Q, _clear_inspection_memo
from wolfdocx.package import PackageGraph
from wolfdocx.paragraph_authoring import (
    ParagraphReport,
    ParagraphRequest,
    ParagraphTarget,
    author_paragraph,
    discover_paragraph_targets,
    load_paragraph_request,
    paragraph_authoring_structure_fingerprint,
    parse_paragraph_request,
)

_DOCUMENT = "word/document.xml"

_STYLES_WITH_HEADING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/></w:style>'
    '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/></w:style>'
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
    extra_parts: dict[str, str | bytes] | None = None,
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
    package_parts: dict[str, str | bytes] = {}
    if styles_xml is not None:
        package_parts["word/styles.xml"] = styles_xml
    if extra_parts is not None:
        package_parts.update(extra_parts)

    return write_package(
        path,
        body,
        parts=package_parts if package_parts else None,
    )


def test_discover_paragraph_targets_basic(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample.docx")
    source_sha = _sha256(source)

    discovery = discover_paragraph_targets(source)
    assert discovery["passed"] is True
    assert discovery["operation"] == "paragraph-targets"
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

    target_obj = ParagraphTarget(
        part=_DOCUMENT,
        paragraph=1,
        text="First paragraph.",
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
    )
    assert target_obj.to_dict()["paragraph"] == 1
    assert target_obj.to_dict()["text"] == "First paragraph."


def test_discover_paragraph_targets_errors(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample.docx")

    # Missing file
    missing = discover_paragraph_targets(tmp_path / "nonexistent.docx")
    assert missing["passed"] is False
    assert missing["status"] == "refused"
    assert missing["code"] == "paragraph_source_missing"

    # Unsupported part
    bad_part = discover_paragraph_targets(source, part="word/header1.xml")
    assert bad_part["passed"] is False
    assert bad_part["status"] == "refused"
    assert bad_part["code"] == "paragraph_part_unsupported"

    # Corrupt package
    corrupt_file = tmp_path / "corrupt.docx"
    corrupt_file.write_bytes(b"not a valid zip")
    bad_pkg = discover_paragraph_targets(corrupt_file)
    assert bad_pkg["passed"] is False
    assert bad_pkg["status"] == "refused"
    assert bad_pkg["code"] == "paragraph_package_invalid"


def test_discover_paragraph_targets_nested_and_hazards(tmp_path: Path) -> None:
    # Body with table (containing nested cell paragraphs) and a revision paragraph
    p1 = paragraph("Body paragraph 1.")
    p_nested = table(row(cell(paragraph("Nested in table"))), grid=grid([2500]))
    p_rev = '<w:p><w:ins w:id="1" w:author="User" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Tracked text</w:t></w:r></w:ins></w:p>'
    p_txbx = (
        '<w:p><w:r><w:pict><v:textbox xmlns:v="urn:schemas-microsoft-com:vml">'
        "<w:txbxContent><w:p><w:r><w:t>Textbox nested</w:t></w:r></w:p></w:txbxContent>"
        "</v:textbox></w:pict></w:r></w:p>"
    )
    p2 = paragraph("Body paragraph 2.")

    body = p1 + p_nested + p_rev + p_txbx + p2
    source = write_package(tmp_path / "hazards.docx", body)

    discovery = discover_paragraph_targets(source)
    assert discovery["passed"] is True
    # Table cell paragraph (index 2) is admitted and appears as a target.
    # p1 (index 1), p_nested cell (index 2), p_rev (index 3), p_txbx outer (index 4), p2 (index 6)
    # The inner textbox paragraph (index 5) is non-cell unsupported nesting and does NOT appear in targets.
    target_indices = [t["paragraph"] for t in discovery["targets"]]
    assert 1 in target_indices
    assert 2 in target_indices
    assert 3 in target_indices
    assert 4 in target_indices
    assert 5 not in target_indices
    assert 6 in target_indices

    # p_nested cell paragraph should be supported
    p_nested_target = next(t for t in discovery["targets"] if t["paragraph"] == 2)
    assert p_nested_target["text"] == "Nested in table"
    assert p_nested_target["supported"] is True
    assert p_nested_target.get("unsupported_reason") is None

    # p_rev should be marked supported=False with unsupported_reason="paragraph_revision_unsupported"
    p_rev_target = next(t for t in discovery["targets"] if t["paragraph"] == 3)
    assert p_rev_target["supported"] is False
    assert p_rev_target["unsupported_reason"] == "paragraph_revision_unsupported"


def test_paragraph_authoring_structure_fingerprint(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample.docx")

    fp1 = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp1 is not None
    assert len(fp1) == 64

    # Target discovery matches the direct fingerprint helper
    discovery = discover_paragraph_targets(source)
    assert discovery["targets"][0]["structure_fingerprint_sha256"] == fp1

    # Out of bounds
    assert paragraph_authoring_structure_fingerprint(source, paragraph=0) is None
    assert paragraph_authoring_structure_fingerprint(source, paragraph=999) is None

    # Unsupported part
    assert (
        paragraph_authoring_structure_fingerprint(source, paragraph=1, part="word/other.xml")
        is None
    )

    # Missing file
    assert (
        paragraph_authoring_structure_fingerprint(tmp_path / "nonexistent.docx", paragraph=1)
        is None
    )


def test_parse_and_load_request(tmp_path: Path) -> None:
    source = tmp_path / "src.docx"
    dest = tmp_path / "dst.docx"
    fake_sha = "a" * 64
    fake_fp = "b" * 64

    # Valid insert request
    req_dict = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": fake_sha,
        "structure_fingerprint_sha256": fake_fp,
        "target": {"paragraph": 1, "part": _DOCUMENT},
        "action": "insert",
        "position": "after",
        "text": "Hello world",
        "style": "Heading",
    }
    req = parse_paragraph_request(req_dict)
    assert req.action == "insert"
    assert req.position == "after"
    assert req.text == "Hello world"
    assert req.style == "Heading"
    assert req.paragraph == 1
    assert req.part == _DOCUMENT
    assert req.to_dict()["action"] == "insert"

    # Valid delete request
    del_dict = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": fake_sha,
        "structure_fingerprint_sha256": fake_fp,
        "target": {"paragraph": 2},
        "action": "delete",
    }
    del_req = parse_paragraph_request(del_dict)
    assert del_req.action == "delete"
    assert del_req.position is None
    assert del_req.text is None
    assert del_req.style is None
    assert del_req.paragraph == 2

    # Roundtrip via file
    req_path = tmp_path / "req.json"
    req_path.write_text(json.dumps(req_dict), encoding="utf-8")
    loaded = load_paragraph_request(req_path)
    assert loaded.action == req.action
    assert loaded.text == req.text

    # Delete request with extra insert fields raises ValueError
    bad_del = dict(del_dict, text="Not allowed on delete")
    with pytest.raises(ValueError, match="paragraph_request_fields_invalid"):
        parse_paragraph_request(bad_del)

    # Invalid action raises ValueError
    bad_action = dict(req_dict, action="replace")
    with pytest.raises(ValueError, match="paragraph_action_invalid"):
        parse_paragraph_request(bad_action)


def test_author_paragraph_insert_after(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        position="after",
        text="Inserted after first.",
    )
    report = author_paragraph(req)
    assert isinstance(report, ParagraphReport)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.source_sha256 == source_sha
    assert report.destination_sha256 is not None
    assert report.target_change_report is not None
    assert report.target_change_report["action"] == "insert"
    assert report.target_change_report["position"] == "after"
    assert report.target_change_report["anchor_paragraph"] == 1
    assert report.target_change_report["inserted_paragraph_index"] == 2
    assert report.target_change_report["text_characters"] == len("Inserted after first.")
    assert (
        report.target_change_report["text_sha256"]
        == hashlib.sha256(b"Inserted after first.").hexdigest()
    )
    assert "text" not in report.target_change_report
    assert report.non_target_parts_preserved is True
    assert dest.is_file()

    # Read back document and verify paragraph sequence
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    body = root.find(Q("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "body"))
    assert body is not None
    p_elements = [
        c
        for c in body
        if c.tag == Q("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p")
    ]
    assert len(p_elements) == 4

    p_texts = [
        "".join(
            t.text or ""
            for t in p.iter(Q("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "t"))
        )
        for p in p_elements
    ]
    assert p_texts == [
        "First paragraph.",
        "Inserted after first.",
        "Second paragraph.",
        "Third paragraph.",
    ]


def test_author_paragraph_insert_before(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        position="before",
        text="Inserted before first.",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"
    assert dest.is_file()

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None
    p_elements = [c for c in body if c.tag == Q(ns, "p")]
    assert len(p_elements) == 4
    p_texts = ["".join(t.text or "" for t in p.iter(Q(ns, "t"))) for p in p_elements]
    assert p_texts[0] == "Inserted before first."
    assert p_texts[1] == "First paragraph."


def test_author_paragraph_insert_with_style(tmp_path: Path) -> None:
    source = _make_sample_package(
        tmp_path / "source.docx",
        styles_xml=_STYLES_WITH_HEADING,
    )
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=2)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=fp,
        action="insert",
        position="after",
        text="Styled Heading Paragraph",
        style="Heading",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None
    p_elements = [c for c in body if c.tag == Q(ns, "p")]
    inserted_p = p_elements[2]
    ppr = inserted_p.find(Q(ns, "pPr"))
    assert ppr is not None
    pstyle = ppr.find(Q(ns, "pStyle"))
    assert pstyle is not None
    assert pstyle.get(Q(ns, "val")) == "Heading"


def test_author_paragraph_insert_whitespace_preservation(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        position="after",
        text="   Leading and trailing spaces   ",
    )
    report = author_paragraph(req)
    assert report.passed is True

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None
    p_elements = [c for c in body if c.tag == Q(ns, "p")]
    t_elem = p_elements[1].find(f".//{Q(ns, 't')}")
    assert t_elem is not None
    assert t_elem.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert t_elem.text == "   Leading and trailing spaces   "


def test_author_paragraph_delete_middle(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=2)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.target_change_report is not None
    assert report.target_change_report["action"] == "delete"
    assert report.target_change_report["deleted_paragraph_index"] == 2
    assert report.target_change_report["deleted_text_characters"] == len("Second paragraph.")
    assert (
        report.target_change_report["deleted_text_sha256"]
        == hashlib.sha256(b"Second paragraph.").hexdigest()
    )
    assert "deleted_text" not in report.target_change_report
    # Verify paragraph 2 is removed and 1 and 3 remain
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None
    p_elements = [c for c in body if c.tag == Q(ns, "p")]
    assert len(p_elements) == 2
    p_texts = ["".join(t.text or "" for t in p.iter(Q(ns, "t"))) for p in p_elements]
    assert p_texts == ["First paragraph.", "Third paragraph."]


def test_author_paragraph_delete_first(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"

    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None
    p_elements = [c for c in body if c.tag == Q(ns, "p")]
    assert len(p_elements) == 2
    p_texts = ["".join(t.text or "" for t in p.iter(Q(ns, "t"))) for p in p_elements]
    assert p_texts == ["Second paragraph.", "Third paragraph."]


def test_refusal_final_paragraph_deletion(tmp_path: Path) -> None:
    # 3-paragraph document: attempting to delete paragraph 3 (the final body paragraph)
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=3)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 3},
        structure_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_final_paragraph_deletion"
    assert not dest.exists()

    # 1-paragraph document: attempting to delete only paragraph
    single_source = _make_sample_package(tmp_path / "single.docx", paragraphs=["Only one."])
    single_sha = _sha256(single_source)
    single_fp = paragraph_authoring_structure_fingerprint(single_source, paragraph=1)
    assert single_fp is not None

    single_req = ParagraphRequest(
        source=single_source,
        destination=tmp_path / "single_dest.docx",
        source_sha256=single_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=single_fp,
        action="delete",
    )
    single_report = author_paragraph(single_req)
    assert single_report.passed is False
    assert single_report.status == "refused"
    assert single_report.code == "paragraph_final_paragraph_deletion"
    assert not (tmp_path / "single_dest.docx").exists()


def test_refusal_section_boundary_deletion(tmp_path: Path) -> None:
    # Paragraph carrying sectPr in pPr (section break)
    p1 = paragraph("Section 1 body.")
    p_sect = (
        '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
        "<w:r><w:t>Section 1 terminating</w:t></w:r></w:p>"
    )
    p2 = paragraph("Section 2 body.")
    body = p1 + p_sect + p2
    source = write_package(tmp_path / "multi_section.docx", body)
    source_sha = _sha256(source)
    dest = tmp_path / "dest.docx"

    fp = paragraph_authoring_structure_fingerprint(source, paragraph=2)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_section_boundary_deletion"
    assert not dest.exists()


def test_refusal_source_hash_mismatch(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256="0" * 64,  # Stale hash
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        text="Should fail",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_source_hash_mismatch"
    assert not dest.exists()


def test_refusal_destination_occupied(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "occupied.docx"
    dest.write_text("existing content", encoding="utf-8")
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        text="Should fail",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_destination_occupied"
    assert dest.read_text(encoding="utf-8") == "existing content"


def test_refusal_structure_fingerprint_mismatch(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256="f" * 64,  # wrong fingerprint
        action="insert",
        text="Should fail",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_structure_fingerprint_mismatch"
    assert not dest.exists()


def test_refusal_target_missing(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 99},  # Out of range
        structure_fingerprint_sha256="a" * 64,
        action="insert",
        text="Should fail",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_target_missing"
    assert not dest.exists()


def test_refusal_target_nested(tmp_path: Path) -> None:
    # Text-box nested paragraph remains refused with paragraph_target_nested
    p1 = paragraph("Body 1")
    p_txbx = (
        '<w:p><w:r><w:pict><v:textbox xmlns:v="urn:schemas-microsoft-com:vml">'
        "<w:txbxContent><w:p><w:r><w:t>Textbox nested</w:t></w:r></w:p></w:txbxContent>"
        "</v:textbox></w:pict></w:r></w:p>"
    )
    p2 = paragraph("Body 2")
    source = write_package(tmp_path / "nested.docx", p1 + p_txbx + p2)
    source_sha = _sha256(source)
    dest = tmp_path / "dest.docx"

    # Paragraph 2 is the outer p with drawing, paragraph 3 is inside txbxContent
    fp3 = paragraph_authoring_structure_fingerprint(source, paragraph=3)
    assert fp3 is not None
    assert len(fp3) == 64

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 3},
        structure_fingerprint_sha256=fp3,
        action="insert",
        text="Should fail",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_target_nested"
    assert not dest.exists()


def test_admitted_target_table_cell(tmp_path: Path) -> None:
    p1 = paragraph("Body 1")
    tbl = table(row(cell(paragraph("Nested cell p"))), grid=grid([2500]))
    p2 = paragraph("Body 2")
    source = write_package(tmp_path / "cell_admitted.docx", p1 + tbl + p2)
    source_sha = _sha256(source)
    dest = tmp_path / "dest.docx"

    # Paragraph 2 is inside table cell and is now admitted
    fp2 = paragraph_authoring_structure_fingerprint(source, paragraph=2)
    assert fp2 is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=fp2,
        action="insert",
        position="after",
        text="Admitted cell insert",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"
    assert dest.exists()


def test_refusal_style_missing(tmp_path: Path) -> None:
    source = _make_sample_package(
        tmp_path / "source.docx",
        styles_xml=_STYLES_WITH_HEADING,
    )
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        text="Test",
        style="NonExistentStyle",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_style_missing"
    assert not dest.exists()


def test_refusal_invalid_xml_text(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    # Null character \x00 or formfeed \x0c
    with pytest.raises(ValueError, match="paragraph_text_invalid"):
        ParagraphRequest(
            source=source,
            destination=dest,
            source_sha256=source_sha,
            target={"paragraph": 1},
            structure_fingerprint_sha256=fp,
            action="insert",
            text="Invalid \x00 char",
        )


def test_refusal_annotation_hazards(tmp_path: Path) -> None:
    # Bookmark hazard on delete
    p_bkmk = (
        '<w:p><w:bookmarkStart w:id="1" w:name="Bookmark1"/>'
        "<w:r><w:t>Bookmarked paragraph</w:t></w:r>"
        '<w:bookmarkEnd w:id="1"/></w:p>'
    )
    body = paragraph("P1") + p_bkmk + paragraph("P3")
    source = write_package(tmp_path / "bkmk.docx", body)
    source_sha = _sha256(source)
    dest = tmp_path / "dest.docx"

    fp = paragraph_authoring_structure_fingerprint(source, paragraph=2)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_paragraph(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "paragraph_bookmark_unsupported"
    assert not dest.exists()


def test_preservation_unrelated_parts(tmp_path: Path) -> None:
    # Create package with custom extra part
    extra_part_content = b"<customData><item>PreserveMe</item></customData>"
    source = _make_sample_package(
        tmp_path / "source.docx",
        extra_parts={"custom/part.xml": extra_part_content},
    )
    dest = tmp_path / "dest.docx"
    source_sha = _sha256(source)
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="insert",
        position="after",
        text="New middle paragraph",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.non_target_parts_preserved is True

    # Verify custom part is identical byte-for-byte
    dest_graph = PackageGraph.from_file(dest)
    assert "custom/part.xml" in dest_graph.parts
    assert dest_graph.read_part_bytes("custom/part.xml") == extra_part_content


def test_cross_paragraph_bookmark_interior_and_boundary_refused(tmp_path: Path) -> None:
    # Paragraph 1 has bookmarkStart, paragraph 2 is interior, paragraph 3 has bookmarkEnd
    body = (
        '<w:p><w:bookmarkStart w:id="9" w:name="Range"/><w:r><w:t>Start</w:t></w:r></w:p>'
        + paragraph("Middle")
        + '<w:p><w:r><w:t>End</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>'
    )
    source = write_package(tmp_path / "cross_bkmk.docx", body)
    source_sha = _sha256(source)

    discovery = discover_paragraph_targets(source)
    assert discovery["passed"] is True
    assert len(discovery["targets"]) == 3

    # All three paragraphs (boundary start, interior, boundary end) are marked unsupported
    t1, t2, t3 = discovery["targets"]
    assert t1["supported"] is False
    assert t1["unsupported_reason"] == "paragraph_bookmark_unsupported"
    assert t2["supported"] is False
    assert t2["unsupported_reason"] == "paragraph_bookmark_unsupported"
    assert t3["supported"] is False
    assert t3["unsupported_reason"] == "paragraph_bookmark_unsupported"

    # Deletion of middle (interior) paragraph is refused
    dest_del_interior = tmp_path / "del_interior.docx"
    report_del_interior = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=dest_del_interior,
            source_sha256=source_sha,
            target=t2["target"],
            structure_fingerprint_sha256=t2["structure_fingerprint_sha256"],
            action="delete",
        )
    )
    assert report_del_interior.passed is False
    assert report_del_interior.status == "refused"
    assert report_del_interior.code == "paragraph_bookmark_unsupported"
    assert not dest_del_interior.exists()

    # Insertion at middle (interior) paragraph is also refused
    dest_ins_interior = tmp_path / "ins_interior.docx"
    report_ins_interior = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=dest_ins_interior,
            source_sha256=source_sha,
            target=t2["target"],
            structure_fingerprint_sha256=t2["structure_fingerprint_sha256"],
            action="insert",
            text="Hazard insertion",
        )
    )
    assert report_ins_interior.passed is False
    assert report_ins_interior.status == "refused"
    assert report_ins_interior.code == "paragraph_bookmark_unsupported"
    assert not dest_ins_interior.exists()


def test_content_safe_report_never_leaks_source_text(tmp_path: Path) -> None:
    secret_delete = "SUPER_SECRET_PATENT_CLAIM_99"
    source = write_package(tmp_path / "secret.docx", paragraph(secret_delete) + paragraph("Keep"))
    source_sha = _sha256(source)
    discovery = discover_paragraph_targets(source)
    target = discovery["targets"][0]

    report = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=tmp_path / "deleted_secret.docx",
            source_sha256=source_sha,
            target=target["target"],
            structure_fingerprint_sha256=target["structure_fingerprint_sha256"],
            action="delete",
        )
    )
    assert report.passed is True
    report_json = json.dumps(report.to_dict())
    assert secret_delete not in report_json

    # Insert operation report content-safety
    secret_insert = "NEW_CLASSIFIED_SUBSECTION"
    target_keep = discovery["targets"][1]
    ins_report = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=tmp_path / "inserted_secret.docx",
            source_sha256=source_sha,
            target=target_keep["target"],
            structure_fingerprint_sha256=target_keep["structure_fingerprint_sha256"],
            action="insert",
            position="before",
            text=secret_insert,
        )
    )
    assert ins_report.passed is True
    ins_report_json = json.dumps(ins_report.to_dict())
    assert secret_insert not in ins_report_json


def test_paragraph_authoring_request_lifecycle(tmp_path: Path) -> None:
    # 1. Discover targets
    source = _make_sample_package(tmp_path / "source.docx")
    source_sha = _sha256(source)

    discovery = discover_paragraph_targets(source)
    assert discovery["passed"] is True
    assert discovery["operation"] == "paragraph-targets"
    assert len(discovery["targets"]) == 3
    t1 = discovery["targets"][0]
    assert t1["paragraph"] == 1
    assert t1["supported"] is True

    # 2. Write request JSON
    dest_pub = tmp_path / "published.docx"
    req_data = {
        "source": str(source),
        "destination": str(dest_pub),
        "source_sha256": source_sha,
        "target": {"paragraph": 1},
        "structure_fingerprint_sha256": t1["structure_fingerprint_sha256"],
        "action": "insert",
        "position": "after",
        "text": "Inserted paragraph text.",
    }
    req_file = tmp_path / "insert_req.json"
    req_file.write_text(json.dumps(req_data), encoding="utf-8")

    # Verify parse and load helpers agree
    parsed_req = parse_paragraph_request(req_data)
    loaded_req = load_paragraph_request(req_file)
    assert parsed_req == loaded_req
    assert isinstance(loaded_req, ParagraphRequest)

    # 3. Publish the loaded request
    report_dict = author_paragraph(loaded_req).to_dict()
    assert report_dict["passed"] is True
    assert report_dict["status"] == "published"
    assert report_dict["code"] is None
    assert dest_pub.is_file()

    # 4. Occupied destination refusal
    dest_occupied = tmp_path / "occupied.docx"
    dest_occupied.write_bytes(b"EXISTING_CONTENT")

    req_occ_data = dict(req_data)
    req_occ_data["destination"] = str(dest_occupied)
    occ_report_dict = author_paragraph(parse_paragraph_request(req_occ_data)).to_dict()
    assert occ_report_dict["passed"] is False
    assert occ_report_dict["status"] == "refused"
    assert occ_report_dict["code"] == "paragraph_destination_occupied"
    assert dest_occupied.read_bytes() == b"EXISTING_CONTENT"


def _make_header_footer_package(
    path: Path,
    *,
    header_paragraphs: list[str] | None = None,
    footer_paragraphs: list[str] | None = None,
    extra_parts: dict[str, str | bytes] | None = None,
    extra_relationships: str = "",
    extra_overrides: str = "",
    multi_section: bool = False,
) -> Path:
    """Create a package with referenced header and footer parts."""
    h_texts = (
        header_paragraphs if header_paragraphs is not None else ["Header First.", "Header Second."]
    )
    f_texts = (
        footer_paragraphs if footer_paragraphs is not None else ["Footer First.", "Footer Second."]
    )

    h_body = "".join(f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>" for t in h_texts)
    f_body = "".join(f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>" for t in f_texts)

    parts: dict[str, str | bytes] = {
        "word/header1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W}">{h_body}</w:hdr>',
        "word/footer1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W}">{f_body}</w:ftr>',
    }
    if extra_parts:
        parts.update(extra_parts)

    rels = (
        region_relationship("rIdH1", "header", "header1.xml")
        + region_relationship("rIdF1", "footer", "footer1.xml")
        + extra_relationships
    )
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        + extra_overrides
    )

    body_text = "<w:p><w:r><w:t>Document body paragraph.</w:t></w:r></w:p>"
    if multi_section:
        sec1_pr = sect(references=header_reference("rIdH1") + footer_reference("rIdF1"))
        body_text += f"<w:p><w:pPr>{sec1_pr}</w:pPr><w:r><w:t>Section 1 boundary.</w:t></w:r></w:p>"
        body_text += "<w:p><w:r><w:t>Section 2 paragraph.</w:t></w:r></w:p>"

    sec_final = sect(references=header_reference("rIdH1") + footer_reference("rIdF1"))

    return write_package(
        path,
        body=body_text,
        sect_properties=sec_final,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )


def test_header_footer_paragraph_targets_discovery(tmp_path: Path) -> None:
    source = _make_header_footer_package(tmp_path / "header_footer.docx")
    source_sha = _sha256(source)

    h_discovery = discover_paragraph_targets(source, part="word/header1.xml")
    assert h_discovery["passed"] is True
    assert h_discovery["operation"] == "paragraph-targets"
    assert h_discovery["status"] == "discovered"
    assert h_discovery["source_sha256"] == source_sha
    assert h_discovery["part"] == "word/header1.xml"
    assert len(h_discovery["targets"]) == 2
    t1 = h_discovery["targets"][0]
    assert t1["paragraph"] == 1
    assert t1["part"] == "word/header1.xml"
    assert t1["text"] == "Header First."
    assert t1["supported"] is True
    assert t1["empty"] is False

    f_discovery = discover_paragraph_targets(source, part="word/footer1.xml")
    assert f_discovery["passed"] is True
    assert f_discovery["part"] == "word/footer1.xml"
    assert len(f_discovery["targets"]) == 2
    assert f_discovery["targets"][1]["text"] == "Footer Second."

    # Fingerprint helper on header
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1, part="word/header1.xml")
    assert fp == t1["structure_fingerprint_sha256"]

    # Refusal on unreferenced header part
    unref_source = write_package(
        tmp_path / "unref.docx",
        body="<w:p><w:r><w:t>Body.</w:t></w:r></w:p>",
        parts={
            "word/header2.xml": f'<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Orphan</w:t></w:r></w:p></w:hdr>'
        },
        overrides='<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    unref_disc = discover_paragraph_targets(unref_source, part="word/header2.xml")
    assert unref_disc["passed"] is False
    assert unref_disc["status"] == "refused"
    assert unref_disc["code"] == "paragraph_part_unsupported"


def test_header_footer_paragraph_insert_and_preservation(tmp_path: Path) -> None:
    source = _make_header_footer_package(tmp_path / "insert_source.docx")
    source_sha = _sha256(source)
    dest = tmp_path / "insert_output.docx"

    disc = discover_paragraph_targets(source, part="word/header1.xml")
    t2 = disc["targets"][1]

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"part": t2["part"], "paragraph": t2["paragraph"]},
        structure_fingerprint_sha256=t2["structure_fingerprint_sha256"],
        action="insert",
        position="before",
        text="Inserted Header Middle.",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.non_target_parts_preserved is True
    assert dest.is_file()

    # Verify staged/published header content
    graph = PackageGraph.from_file(dest)
    h_root = ET.fromstring(graph.read_part_bytes("word/header1.xml"))
    h_paras = [
        "".join(t.text or "" for t in p.iter(Q(W, "t"))) for p in h_root.findall(f"./{Q(W, 'p')}")
    ]
    assert h_paras == ["Header First.", "Inserted Header Middle.", "Header Second."]

    # Verify document.xml and footer1.xml are byte-preserved
    src_graph = PackageGraph.from_file(source)
    assert graph.read_part_bytes("word/document.xml") == src_graph.read_part_bytes(
        "word/document.xml"
    )
    assert graph.read_part_bytes("word/footer1.xml") == src_graph.read_part_bytes(
        "word/footer1.xml"
    )


def test_header_footer_paragraph_delete_safety(tmp_path: Path) -> None:
    source = _make_header_footer_package(tmp_path / "del_source.docx")
    source_sha = _sha256(source)
    dest = tmp_path / "del_output.docx"

    disc = discover_paragraph_targets(source, part="word/header1.xml")
    t1 = disc["targets"][0]

    # 1. Delete paragraph 1 of 2
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"part": t1["part"], "paragraph": t1["paragraph"]},
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
        action="delete",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.target_change_report is not None
    assert report.target_change_report["action"] == "delete"
    assert report.target_change_report["deleted_paragraph_index"] == 1

    # 2. Refuse deleting the final remaining paragraph in header
    dest_single = tmp_path / "del_single_fail.docx"
    disc_dest = discover_paragraph_targets(dest, part="word/header1.xml")
    assert len(disc_dest["targets"]) == 1
    t_only = disc_dest["targets"][0]
    req_fail = ParagraphRequest(
        source=dest,
        destination=dest_single,
        source_sha256=_sha256(dest),
        target={"part": t_only["part"], "paragraph": t_only["paragraph"]},
        structure_fingerprint_sha256=t_only["structure_fingerprint_sha256"],
        action="delete",
    )
    report_fail = author_paragraph(req_fail)
    assert report_fail.passed is False
    assert report_fail.status == "refused"
    assert report_fail.code == "paragraph_final_paragraph_deletion"
    assert not dest_single.exists()


def test_shared_header_across_sections_preserved(tmp_path: Path) -> None:
    source = _make_header_footer_package(tmp_path / "shared_source.docx", multi_section=True)
    source_sha = _sha256(source)
    dest = tmp_path / "shared_output.docx"

    disc = discover_paragraph_targets(source, part="word/header1.xml")
    t1 = disc["targets"][0]

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"part": t1["part"], "paragraph": t1["paragraph"]},
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Inserted in Shared Header.",
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"

    # Verify section relationships in document.xml are preserved
    graph = PackageGraph.from_file(dest)
    doc_root = ET.fromstring(graph.read_part_bytes("word/document.xml"))
    hdr_refs = doc_root.findall(f".//{Q(W, 'headerReference')}")
    assert len(hdr_refs) == 2
    for ref in hdr_refs:
        assert ref.get(Q(R, "id")) == "rIdH1"


def test_range_hazard_sibling_marker_refused(tmp_path: Path) -> None:
    h_content = (
        '<w:p><w:r><w:t>P1 start</w:t></w:r><w:bookmarkStart w:id="10" w:name="SpanMark"/></w:p>'
        "<w:p><w:r><w:t>P2 interior</w:t></w:r></w:p>"
        '<w:bookmarkEnd w:id="10"/>'
        "<w:p><w:r><w:t>P3 clean</w:t></w:r></w:p>"
    )
    extra_parts = {
        "word/header1.xml": f'<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="{W}">{h_content}</w:hdr>'
    }
    source = _make_header_footer_package(tmp_path / "hazard_source.docx", extra_parts=extra_parts)
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source, part="word/header1.xml")
    assert disc["passed"] is True
    targets = disc["targets"]
    assert len(targets) == 3
    assert targets[0]["supported"] is False
    assert targets[0]["unsupported_reason"] == "paragraph_bookmark_unsupported"
    assert targets[1]["supported"] is False
    assert targets[1]["unsupported_reason"] == "paragraph_bookmark_unsupported"
    assert targets[2]["supported"] is True

    # Refusal when authoring target 2
    dest = tmp_path / "hazard_out.docx"
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"part": targets[1]["part"], "paragraph": targets[1]["paragraph"]},
        structure_fingerprint_sha256=targets[1]["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Should fail.",
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "paragraph_bookmark_unsupported"
    assert not dest.exists()


def test_header_paragraph_request_lifecycle(tmp_path: Path) -> None:
    source = _make_header_footer_package(tmp_path / "header_request.docx")
    source_sha = _sha256(source)

    # 1. Discover header targets
    disc = discover_paragraph_targets(source, part="word/header1.xml")
    assert disc["passed"] is True
    assert disc["part"] == "word/header1.xml"
    assert len(disc["targets"]) == 2
    t1 = disc["targets"][0]

    # 2. Insert through a parsed request
    dest = tmp_path / "header_request_out.docx"
    req_data = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": source_sha,
        "target": {"part": "word/header1.xml", "paragraph": 1},
        "structure_fingerprint_sha256": t1["structure_fingerprint_sha256"],
        "action": "insert",
        "position": "after",
        "text": "Header Inserted.",
    }
    rep = author_paragraph(parse_paragraph_request(req_data)).to_dict()
    assert rep["passed"] is True
    assert rep["status"] == "published"
    assert dest.is_file()


def test_table_cell_paragraph_insert_before_and_after(tmp_path: Path) -> None:
    p1 = paragraph("Body Before Table.")
    c1 = cell(paragraph("Cell 1 Alpha."))
    c2 = cell(paragraph("Cell 2 Alpha."))
    tbl = table(row(c1 + c2), grid=grid([2500, 2500]))
    p2 = paragraph("Body After Table.")
    source = write_package(tmp_path / "cell_insert_source.docx", p1 + tbl + p2)
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 4
    # 1: p1, 2: c1, 3: c2, 4: p2
    t_c1 = disc["targets"][1]
    assert t_c1["paragraph"] == 2
    assert t_c1["text"] == "Cell 1 Alpha."
    assert t_c1["supported"] is True

    # Insert after Cell 1 Alpha
    dest_after = tmp_path / "cell_insert_after.docx"
    req_after = ParagraphRequest(
        source=source,
        destination=dest_after,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_c1["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Cell 1 Beta.",
    )
    rep_after = author_paragraph(req_after)
    assert rep_after.passed is True
    assert rep_after.status == "published"
    assert dest_after.is_file()

    # Verify XML structure in dest_after: Cell 1 has both paragraphs, tcPr preserved
    g = PackageGraph.from_file(dest_after)
    root = ET.fromstring(g.read_part_bytes(_DOCUMENT))
    cells = list(root.iter(Q(W, "tc")))
    assert len(cells) == 2
    c1_children = list(cells[0])
    assert c1_children[0].tag == Q(W, "tcPr")
    assert c1_children[1].tag == Q(W, "p")
    assert "".join(t.text or "" for t in c1_children[1].iter(Q(W, "t"))) == "Cell 1 Alpha."
    assert c1_children[2].tag == Q(W, "p")
    assert "".join(t.text or "" for t in c1_children[2].iter(Q(W, "t"))) == "Cell 1 Beta."

    # Insert before Cell 1 Alpha
    dest_before = tmp_path / "cell_insert_before.docx"
    req_before = ParagraphRequest(
        source=source,
        destination=dest_before,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_c1["structure_fingerprint_sha256"],
        action="insert",
        position="before",
        text="Cell 1 Intro.",
    )
    rep_before = author_paragraph(req_before)
    assert rep_before.passed is True
    assert rep_before.status == "published"

    g = PackageGraph.from_file(dest_before)
    root = ET.fromstring(g.read_part_bytes(_DOCUMENT))
    cells = list(root.iter(Q(W, "tc")))
    c1_children = list(cells[0])
    assert c1_children[0].tag == Q(W, "tcPr")
    assert c1_children[1].tag == Q(W, "p")
    assert "".join(t.text or "" for t in c1_children[1].iter(Q(W, "t"))) == "Cell 1 Intro."
    assert c1_children[2].tag == Q(W, "p")
    assert "".join(t.text or "" for t in c1_children[2].iter(Q(W, "t"))) == "Cell 1 Alpha."


def test_table_cell_paragraph_styled_insert(tmp_path: Path) -> None:
    p1 = paragraph("Body 1")
    c = cell(paragraph("Cell content"))
    tbl = table(row(c), grid=grid([2500]))
    source = write_package(
        tmp_path / "styled_source.docx",
        p1 + tbl,
        parts={"word/styles.xml": _STYLES_WITH_HEADING},
    )
    source_sha = _sha256(source)
    dest = tmp_path / "styled_dest.docx"

    disc = discover_paragraph_targets(source)
    t_cell = disc["targets"][1]

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_cell["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Cell Header Paragraph",
        style="Heading",
    )
    rep = author_paragraph(req)
    assert rep.passed is True
    assert rep.status == "published"

    g = PackageGraph.from_file(dest)
    root = ET.fromstring(g.read_part_bytes(_DOCUMENT))
    p_inserted = list(root.iter(Q(W, "p")))[2]
    pstyle = p_inserted.find(f".//{Q(W, 'pStyle')}")
    assert pstyle is not None
    assert pstyle.get(Q(W, "val")) == "Heading"


def test_table_cell_paragraph_delete(tmp_path: Path) -> None:
    # Cell with two paragraphs: delete the first, preserve the second
    c_mult = cell(paragraph("Keep 1") + paragraph("Delete me"))
    tbl = table(row(c_mult), grid=grid([2500]))
    p_end = paragraph("End paragraph")
    source = write_package(tmp_path / "cell_delete_source.docx", tbl + p_end)
    source_sha = _sha256(source)
    dest = tmp_path / "cell_delete_dest.docx"

    disc = discover_paragraph_targets(source)
    assert len(disc["targets"]) == 3
    t_del = disc["targets"][1]
    assert t_del["text"] == "Delete me"

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_del["structure_fingerprint_sha256"],
        action="delete",
    )
    rep = author_paragraph(req)
    assert rep.passed is True
    assert rep.status == "published"

    g = PackageGraph.from_file(dest)
    root = ET.fromstring(g.read_part_bytes(_DOCUMENT))
    tc = next(iter(root.iter(Q(W, "tc"))))
    tc_paras = [c for c in tc if c.tag == Q(W, "p")]
    assert len(tc_paras) == 1
    assert "".join(t.text or "" for t in tc_paras[0].iter(Q(W, "t"))) == "Keep 1"


def test_table_cell_paragraph_final_safety_empty_cell(tmp_path: Path) -> None:
    # Cell with only one paragraph: deleting it must be refused to avoid empty cell
    c_single = cell(paragraph("Sole paragraph"))
    tbl = table(row(c_single), grid=grid([2500]))
    p_after = paragraph("Trailing body paragraph")
    source = write_package(tmp_path / "empty_cell_refusal.docx", tbl + p_after)
    source_sha = _sha256(source)
    dest = tmp_path / "should_not_exist.docx"

    disc = discover_paragraph_targets(source)
    t_sole = disc["targets"][0]
    assert t_sole["paragraph"] == 1

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=t_sole["structure_fingerprint_sha256"],
        action="delete",
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "paragraph_final_paragraph_deletion"
    assert not dest.exists()


def test_table_cell_paragraph_final_safety_nested_table(tmp_path: Path) -> None:
    # Cell containing nested table followed by terminal paragraph:
    # deleting terminal paragraph must be refused because a cell cannot end in tbl
    inner_tbl = table(row(cell(paragraph("Inner cell"))), grid=grid([1000]))
    c_nested = f"<w:tc><w:p><w:r><w:t>Before nested tbl</w:t></w:r></w:p>{inner_tbl}<w:p><w:r><w:t>Terminal paragraph</w:t></w:r></w:p></w:tc>"
    outer_tbl = (
        f'<w:tbl><w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid><w:tr>{c_nested}</w:tr></w:tbl>'
    )
    p_body = paragraph("Body paragraph")
    source = write_package(tmp_path / "nested_tbl_cell.docx", outer_tbl + p_body)
    source_sha = _sha256(source)
    dest = tmp_path / "nested_refused.docx"

    disc = discover_paragraph_targets(source)
    # Paras: 1: Before nested tbl, 2: Inner cell, 3: Terminal paragraph, 4: Body paragraph
    t_terminal = disc["targets"][2]
    assert t_terminal["paragraph"] == 3
    assert t_terminal["text"] == "Terminal paragraph"

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 3},
        structure_fingerprint_sha256=t_terminal["structure_fingerprint_sha256"],
        action="delete",
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "paragraph_final_paragraph_deletion"
    assert not dest.exists()

    # But deleting the paragraph BEFORE the nested table is allowed (terminal still exists)
    t_before = disc["targets"][0]
    dest_ok = tmp_path / "delete_before_nested_ok.docx"
    req_ok = ParagraphRequest(
        source=source,
        destination=dest_ok,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=t_before["structure_fingerprint_sha256"],
        action="delete",
    )
    rep_ok = author_paragraph(req_ok)
    assert rep_ok.passed is True
    assert rep_ok.status == "published"
    assert dest_ok.is_file()


def test_table_cell_paragraph_referenced_header_and_footer(tmp_path: Path) -> None:
    # Table inside referenced header
    tbl_hdr = table(
        row(cell(paragraph("Header Cell P1")) + cell(paragraph("Header Cell P2"))),
        grid=grid([2500, 2500]),
    )
    tbl_ftr = table(row(cell(paragraph("Footer Cell P1"))), grid=grid([2500]))
    source = _make_header_footer_package(
        tmp_path / "hf_table.docx",
        header_paragraphs=["Normal header p"],
        footer_paragraphs=["Normal footer p"],
        extra_parts={
            "word/header1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Normal header p</w:t></w:r></w:p>{tbl_hdr}</w:hdr>',
            "word/footer1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Normal footer p</w:t></w:r></w:p>{tbl_ftr}</w:ftr>',
        },
    )
    source_sha = _sha256(source)

    # Discover in header
    disc_h = discover_paragraph_targets(source, part="word/header1.xml")
    assert disc_h["passed"] is True
    # Target 1: Normal header p, Target 2: Header Cell P1, Target 3: Header Cell P2
    assert len(disc_h["targets"]) == 3
    t_h2 = disc_h["targets"][1]
    assert t_h2["text"] == "Header Cell P1"

    # Insert into header table cell
    dest_h = tmp_path / "header_cell_insert.docx"
    req_h = ParagraphRequest(
        source=source,
        destination=dest_h,
        source_sha256=source_sha,
        target={"part": "word/header1.xml", "paragraph": 2},
        structure_fingerprint_sha256=t_h2["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Header Cell P1 Sub",
    )
    rep_h = author_paragraph(req_h)
    assert rep_h.passed is True
    assert rep_h.status == "published"
    assert dest_h.is_file()

    # Discover in footer
    disc_f = discover_paragraph_targets(source, part="word/footer1.xml")
    assert disc_f["passed"] is True
    assert len(disc_f["targets"]) == 2
    t_f2 = disc_f["targets"][1]
    assert t_f2["text"] == "Footer Cell P1"


def test_table_cell_paragraph_refusals_source_fingerprint_occupied(
    tmp_path: Path,
) -> None:
    tbl = table(row(cell(paragraph("Cell P"))), grid=grid([2500]))
    source = write_package(tmp_path / "cell_refusals.docx", paragraph("Body") + tbl)
    source_sha = _sha256(source)
    dest = tmp_path / "occupied.docx"
    dest.write_bytes(b"OCCUPIED_SENTINEL")

    disc = discover_paragraph_targets(source)
    t_cell = disc["targets"][1]

    # Destination occupied
    req_occ = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_cell["structure_fingerprint_sha256"],
        action="insert",
        text="Insert text",
    )
    rep_occ = author_paragraph(req_occ)
    assert rep_occ.passed is False
    assert rep_occ.code == "paragraph_destination_occupied"
    assert dest.read_bytes() == b"OCCUPIED_SENTINEL"

    # Stale source hash
    dest_fresh = tmp_path / "fresh.docx"
    req_hash = ParagraphRequest(
        source=source,
        destination=dest_fresh,
        source_sha256="0" * 64,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_cell["structure_fingerprint_sha256"],
        action="insert",
        text="Insert text",
    )
    rep_hash = author_paragraph(req_hash)
    assert rep_hash.passed is False
    assert rep_hash.code == "paragraph_source_hash_mismatch"

    # Stale fingerprint
    req_fp = ParagraphRequest(
        source=source,
        destination=dest_fresh,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256="f" * 64,
        action="insert",
        text="Insert text",
    )
    rep_fp = author_paragraph(req_fp)
    assert rep_fp.passed is False
    assert rep_fp.code == "paragraph_structure_fingerprint_mismatch"


def test_table_cell_paragraph_hazards_crossing_boundaries(tmp_path: Path) -> None:
    # Bookmark starts in body paragraph and ends in table cell
    p1 = '<w:p><w:bookmarkStart w:id="10" w:name="CrossBM"/><w:r><w:t>P1</w:t></w:r></w:p>'
    c1 = '<w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p></w:tc>'
    tbl = f'<w:tbl><w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid><w:tr>{c1}</w:tr></w:tbl>'
    p2 = paragraph("P2 clean")
    source = write_package(tmp_path / "cross_hazard.docx", p1 + tbl + p2)
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    t_p1 = disc["targets"][0]
    t_c1 = disc["targets"][1]
    t_p2 = disc["targets"][2]

    assert t_p1["supported"] is False
    assert t_p1["unsupported_reason"] == "paragraph_bookmark_unsupported"
    assert t_c1["supported"] is False
    assert t_c1["unsupported_reason"] == "paragraph_bookmark_unsupported"
    assert t_p2["supported"] is True

    # Authoring on cell paragraph 2 is refused
    dest = tmp_path / "refuse_hazard.docx"
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=t_c1["structure_fingerprint_sha256"],
        action="insert",
        text="Fails",
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.code == "paragraph_bookmark_unsupported"
    assert not dest.exists()


def test_table_cell_paragraph_multiple_tables_and_cells_preservation(
    tmp_path: Path,
) -> None:
    # 2 tables, each 2x2. Modify cell (1, 1), verify all other 7 cells and tables preserved byte-for-byte.
    t1_rows = [
        row(cell(paragraph("T1 R1 C1")) + cell(paragraph("T1 R1 C2"))),
        row(cell(paragraph("T1 R2 C1")) + cell(paragraph("T1 R2 C2"))),
    ]
    t2_rows = [
        row(cell(paragraph("T2 R1 C1")) + cell(paragraph("T2 R1 C2"))),
        row(cell(paragraph("T2 R2 C1")) + cell(paragraph("T2 R2 C2"))),
    ]
    tbl1 = table("".join(t1_rows), grid=grid([2500, 2500]))
    tbl2 = table("".join(t2_rows), grid=grid([2500, 2500]))
    p_mid = paragraph("Between tables")
    source = write_package(tmp_path / "multitbl.docx", tbl1 + p_mid + tbl2)
    source_sha = _sha256(source)
    dest = tmp_path / "multitbl_out.docx"

    disc = discover_paragraph_targets(source)
    assert len(disc["targets"]) == 9
    t_target = disc["targets"][0]  # T1 R1 C1

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=t_target["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="T1 R1 C1 Added",
    )
    rep = author_paragraph(req)
    assert rep.passed is True
    assert rep.status == "published"

    g = PackageGraph.from_file(dest)
    root = ET.fromstring(g.read_part_bytes(_DOCUMENT))
    cells = list(root.iter(Q(W, "tc")))
    assert len(cells) == 8
    # Modified cell has 2 paragraphs
    assert len([c for c in cells[0] if c.tag == Q(W, "p")]) == 2
    # All other 7 cells each have 1 paragraph
    for other_tc in cells[1:]:
        assert len([c for c in other_tc if c.tag == Q(W, "p")]) == 1


def test_table_cell_ancestor_row_revisions_refused(tmp_path: Path) -> None:
    # Rows with w:trPr/w:ins, w:trPr/w:del, and w:trPr/w:trPrChange
    r_ins = (
        f'<w:tr><w:trPr><w:ins w:id="1" w:author="User" w:date="2026-01-01T00:00:00Z"/></w:trPr>'
        f"{cell(paragraph('Row Inserted Cell'))}</w:tr>"
    )
    r_del = (
        f'<w:tr><w:trPr><w:del w:id="2" w:author="User" w:date="2026-01-01T00:00:00Z"/></w:trPr>'
        f"{cell(paragraph('Row Deleted Cell'))}</w:tr>"
    )
    r_chg = (
        f'<w:tr><w:trPr><w:trPrChange w:id="3" w:author="User" w:date="2026-01-01T00:00:00Z"><w:trPr/></w:trPrChange></w:trPr>'
        f"{cell(paragraph('Row Changed Cell'))}</w:tr>"
    )
    tbl = table(r_ins + r_del + r_chg, grid=grid([2500]))
    p_body = paragraph("Clean body paragraph")
    source = write_package(tmp_path / "row_revs.docx", tbl + p_body)
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 4
    for idx in (0, 1, 2):
        t = disc["targets"][idx]
        assert t["supported"] is False
        assert t["unsupported_reason"] == "paragraph_revision_unsupported"

    t_clean = disc["targets"][3]
    assert t_clean["supported"] is True
    assert t_clean.get("unsupported_reason") is None

    # Insert on row 1 cell paragraph must refuse with paragraph_revision_unsupported
    dest_ins = tmp_path / "refuse_row_ins.docx"
    req_ins = ParagraphRequest(
        source=source,
        destination=dest_ins,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=disc["targets"][0]["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Should fail",
    )
    rep_ins = author_paragraph(req_ins)
    assert rep_ins.passed is False
    assert rep_ins.status == "refused"
    assert rep_ins.code == "paragraph_revision_unsupported"
    assert not dest_ins.exists()
    assert _sha256(source) == source_sha

    # Delete on row 2 cell paragraph must refuse with paragraph_revision_unsupported
    dest_del = tmp_path / "refuse_row_del.docx"
    req_del = ParagraphRequest(
        source=source,
        destination=dest_del,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=disc["targets"][1]["structure_fingerprint_sha256"],
        action="delete",
    )
    rep_del = author_paragraph(req_del)
    assert rep_del.passed is False
    assert rep_del.status == "refused"
    assert rep_del.code == "paragraph_revision_unsupported"
    assert not dest_del.exists()
    assert _sha256(source) == source_sha


def test_table_cell_ancestor_cell_revisions_refused(tmp_path: Path) -> None:
    # Cells with w:tcPr/w:cellIns, w:cellDel, w:cellMerge, and w:tcPrChange
    c_ins = (
        f'<w:tc><w:tcPr><w:cellIns w:id="10" w:author="User" w:date="2026-01-01T00:00:00Z"/></w:tcPr>'
        f"{paragraph('CellIns Para')}</w:tc>"
    )
    c_del = (
        f'<w:tc><w:tcPr><w:cellDel w:id="11" w:author="User" w:date="2026-01-01T00:00:00Z"/></w:tcPr>'
        f"{paragraph('CellDel Para')}</w:tc>"
    )
    c_mrg = (
        f'<w:tc><w:tcPr><w:cellMerge w:id="12" w:author="User" w:date="2026-01-01T00:00:00Z"/></w:tcPr>'
        f"{paragraph('CellMerge Para')}</w:tc>"
    )
    c_chg = (
        f'<w:tc><w:tcPr><w:tcPrChange w:id="13" w:author="User" w:date="2026-01-01T00:00:00Z"><w:tcPr/></w:tcPrChange></w:tcPr>'
        f"{paragraph('TcPrChange Para')}</w:tc>"
    )
    r1 = row(c_ins + c_del)
    r2 = row(c_mrg + c_chg)
    tbl = table(r1 + r2, grid=grid([2500, 2500]))
    source = write_package(tmp_path / "cell_revs.docx", tbl + paragraph("Body end"))
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 5
    for idx in range(4):
        t = disc["targets"][idx]
        assert t["supported"] is False
        assert t["unsupported_reason"] == "paragraph_revision_unsupported"

    # Insert on cellIns cell paragraph refused
    dest_ins = tmp_path / "refuse_cins.docx"
    req_ins = ParagraphRequest(
        source=source,
        destination=dest_ins,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=disc["targets"][0]["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Fail",
    )
    rep_ins = author_paragraph(req_ins)
    assert rep_ins.passed is False
    assert rep_ins.code == "paragraph_revision_unsupported"
    assert not dest_ins.exists()
    assert _sha256(source) == source_sha

    # Delete on tcPrChange cell paragraph refused
    dest_del = tmp_path / "refuse_cchg.docx"
    req_del = ParagraphRequest(
        source=source,
        destination=dest_del,
        source_sha256=source_sha,
        target={"paragraph": 4},
        structure_fingerprint_sha256=disc["targets"][3]["structure_fingerprint_sha256"],
        action="delete",
    )
    rep_del = author_paragraph(req_del)
    assert rep_del.passed is False
    assert rep_del.code == "paragraph_revision_unsupported"
    assert not dest_del.exists()
    assert _sha256(source) == source_sha


def test_table_cell_ancestor_table_revisions_refused(tmp_path: Path) -> None:
    # Table 1 has w:tblPrChange, Table 2 has w:tblGridChange
    t1_pr = '<w:tblPrChange w:id="20" w:author="User" w:date="2026-01-01T00:00:00Z"><w:tblPr/></w:tblPrChange>'
    tbl1 = table(row(cell(paragraph("Tbl1 Cell"))), properties=t1_pr, grid=grid([2500]))

    t2_grid = (
        '<w:tblGrid><w:gridCol w:w="2500"/>'
        '<w:tblGridChange w:id="21"><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid></w:tblGridChange>'
        "</w:tblGrid>"
    )
    tbl2 = f"<w:tbl><w:tblPr/>{t2_grid}<w:tr>{cell(paragraph('Tbl2 Cell'))}</w:tr></w:tbl>"

    source = write_package(tmp_path / "tbl_revs.docx", tbl1 + paragraph("Between") + tbl2)
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 3
    # Target 1 (Tbl1 Cell) and Target 3 (Tbl2 Cell) are both unsupported
    assert disc["targets"][0]["supported"] is False
    assert disc["targets"][0]["unsupported_reason"] == "paragraph_revision_unsupported"
    assert disc["targets"][1]["supported"] is True
    assert disc["targets"][1].get("unsupported_reason") is None
    assert disc["targets"][2]["supported"] is False
    assert disc["targets"][2]["unsupported_reason"] == "paragraph_revision_unsupported"

    # Insert on Tbl1 cell refused
    dest1 = tmp_path / "refuse_tbl1.docx"
    req1 = ParagraphRequest(
        source=source,
        destination=dest1,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=disc["targets"][0]["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Fail",
    )
    rep1 = author_paragraph(req1)
    assert rep1.passed is False
    assert rep1.code == "paragraph_revision_unsupported"
    assert not dest1.exists()

    # Delete on Tbl2 cell refused
    dest2 = tmp_path / "refuse_tbl2.docx"
    req2 = ParagraphRequest(
        source=source,
        destination=dest2,
        source_sha256=source_sha,
        target={"paragraph": 3},
        structure_fingerprint_sha256=disc["targets"][2]["structure_fingerprint_sha256"],
        action="delete",
    )
    rep2 = author_paragraph(req2)
    assert rep2.passed is False
    assert rep2.code == "paragraph_revision_unsupported"
    assert not dest2.exists()


def test_table_cell_untracked_row_in_partially_tracked_table_admitted(
    tmp_path: Path,
) -> None:
    # Row 1 is tracked (w:trPr/w:ins), Row 2 has 2 paragraphs and no revisions
    r1 = (
        f'<w:tr><w:trPr><w:ins w:id="30" w:author="User" w:date="2026-01-01T00:00:00Z"/></w:trPr>'
        f"{cell(paragraph('Row 1 Tracked'))}</w:tr>"
    )
    r2 = row(cell(paragraph("Row 2 Para 1") + paragraph("Row 2 Para 2")))
    tbl = table(r1 + r2, grid=grid([2500]))
    source = write_package(tmp_path / "partial_tracked.docx", tbl + paragraph("Body tail"))
    source_sha = _sha256(source)

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 4
    assert disc["targets"][0]["supported"] is False
    assert disc["targets"][0]["unsupported_reason"] == "paragraph_revision_unsupported"
    assert disc["targets"][1]["supported"] is True
    assert disc["targets"][1].get("unsupported_reason") is None
    assert disc["targets"][2]["supported"] is True
    assert disc["targets"][2].get("unsupported_reason") is None
    assert disc["targets"][3]["supported"] is True

    # Authoring insert on Row 2 Para 1 succeeds
    dest_ok = tmp_path / "partial_tracked_insert.docx"
    req_ok = ParagraphRequest(
        source=source,
        destination=dest_ok,
        source_sha256=source_sha,
        target={"paragraph": 2},
        structure_fingerprint_sha256=disc["targets"][1]["structure_fingerprint_sha256"],
        action="insert",
        position="after",
        text="Row 2 Inserted",
    )
    rep_ok = author_paragraph(req_ok)
    assert rep_ok.passed is True
    assert rep_ok.status == "published"
    assert dest_ok.is_file()


def test_split_paragraph_basic_across_text(tmp_path: Path) -> None:
    source = _make_sample_package(
        tmp_path / "sample_split.docx",
        paragraphs=["Hello world, this is a split test.", "Second paragraph."],
    )
    source_sha = _sha256(source)
    dest = tmp_path / "split_out.docx"

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=11,  # Split at "Hello world" | ", this is a split test."
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert dest.is_file()

    # Verify reopened target count and text
    disc2 = discover_paragraph_targets(dest)
    assert disc2["passed"] is True
    assert len(disc2["targets"]) == 3
    assert disc2["targets"][0]["text"] == "Hello world"
    assert disc2["targets"][1]["text"] == ", this is a split test."
    assert disc2["targets"][2]["text"] == "Second paragraph."


def test_split_paragraph_across_multiple_runs(tmp_path: Path) -> None:
    # Paragraph with 3 runs: run1: "Bold ", run2: "Italic ", run3: "Plain"
    r1 = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Bold </w:t></w:r>'
    r2 = '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Italic </w:t></w:r>'
    r3 = "<w:r><w:t>Plain</w:t></w:r>"
    p_xml = f"<w:p>{r1}{r2}{r3}</w:p>"
    source = write_package(tmp_path / "runs.docx", p_xml + paragraph("Tail"))
    source_sha = _sha256(source)
    dest = tmp_path / "runs_split.docx"

    disc = discover_paragraph_targets(source)
    assert disc["passed"] is True
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    # Split inside run 2: "Bold " (5 chars) + "Ita" (3 chars) = offset 8
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=8,
    )
    report = author_paragraph(req)
    assert report.passed is True
    assert dest.is_file()

    disc2 = discover_paragraph_targets(dest)
    assert len(disc2["targets"]) == 3
    assert disc2["targets"][0]["text"] == "Bold Ita"
    assert disc2["targets"][1]["text"] == "lic Plain"

    # Verify run formatting preserved in output XML
    graph = PackageGraph.from_file(dest)
    root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    paragraphs = list(root.iter(Q(W, "p")))
    p1_runs = list(paragraphs[0].iter(Q(W, "r")))
    p2_runs = list(paragraphs[1].iter(Q(W, "r")))

    # p1 should have 2 runs: bold "Bold ", italic "Ita"
    assert len(p1_runs) == 2
    assert p1_runs[0].find(f"{Q(W, 'rPr')}/{Q(W, 'b')}") is not None
    assert p1_runs[1].find(f"{Q(W, 'rPr')}/{Q(W, 'i')}") is not None

    # p2 should have 2 runs: italic "lic ", plain "Plain"
    assert len(p2_runs) == 2
    assert p2_runs[0].find(f"{Q(W, 'rPr')}/{Q(W, 'i')}") is not None
    assert p2_runs[1].find(f"{Q(W, 'rPr')}/{Q(W, 'i')}") is None


def test_split_paragraph_boundary_offsets(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample_boundary.docx", paragraphs=["Alpha Beta"])
    source_sha = _sha256(source)
    disc = discover_paragraph_targets(source)
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    # Split at offset 0: paragraph 1 is empty, paragraph 2 is "Alpha Beta"
    dest_zero = tmp_path / "dest_zero.docx"
    req_zero = ParagraphRequest(
        source=source,
        destination=dest_zero,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=0,
    )
    rep_zero = author_paragraph(req_zero)
    assert rep_zero.passed is True
    disc_zero = discover_paragraph_targets(dest_zero)
    assert len(disc_zero["targets"]) == 2
    assert disc_zero["targets"][0]["text"] == ""
    assert disc_zero["targets"][1]["text"] == "Alpha Beta"

    # Split at offset len(text): paragraph 1 is "Alpha Beta", paragraph 2 is empty
    dest_end = tmp_path / "dest_end.docx"
    req_end = ParagraphRequest(
        source=source,
        destination=dest_end,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=10,  # len("Alpha Beta") == 10
    )
    rep_end = author_paragraph(req_end)
    assert rep_end.passed is True
    disc_end = discover_paragraph_targets(dest_end)
    assert len(disc_end["targets"]) == 2
    assert disc_end["targets"][0]["text"] == "Alpha Beta"
    assert disc_end["targets"][1]["text"] == ""


def test_split_paragraph_unicode_offsets(tmp_path: Path) -> None:
    # "Café ☕ 世界" -> len is 9
    unicode_text = "Café ☕ 世界"
    source = _make_sample_package(tmp_path / "sample_unicode.docx", paragraphs=[unicode_text])
    source_sha = _sha256(source)
    disc = discover_paragraph_targets(source)
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    # Split at offset 6: "Café ☕" (6 chars) | " 世界" (3 chars)
    dest = tmp_path / "dest_unicode.docx"
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=6,
    )
    report = author_paragraph(req)
    assert report.passed is True

    disc2 = discover_paragraph_targets(dest)
    assert len(disc2["targets"]) == 2
    assert disc2["targets"][0]["text"] == "Café ☕"
    assert disc2["targets"][1]["text"] == " 世界"


def test_split_paragraph_json_request_lifecycle(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "sample_json.docx", paragraphs=["JSON split test"])
    source_sha = _sha256(source)
    dest = tmp_path / "dest_json.docx"
    disc = discover_paragraph_targets(source)
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    req_payload = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": source_sha,
        "target": {"paragraph": 1},
        "structure_fingerprint_sha256": fp,
        "action": "split",
        "offset": 4,
    }
    req_file = tmp_path / "split_req.json"
    req_file.write_text(json.dumps(req_payload), encoding="utf-8")

    req = load_paragraph_request(req_file)
    assert req.action == "split"
    assert req.offset == 4

    report = author_paragraph(req)
    assert report.passed is True
    assert report.action == "split"
    assert report.paragraph["before_text_characters"] == 4
    assert report.paragraph["after_text_characters"] == 11


def test_split_paragraph_repeated_reopen_lifecycle(tmp_path: Path) -> None:
    source = _make_sample_package(
        tmp_path / "repeat_base.docx",
        paragraphs=["First segment of paragraph text."],
    )
    dest1 = tmp_path / "repeat_step1.docx"
    dest2 = tmp_path / "repeat_step2.docx"

    # Step 1: split "First segment of paragraph text." at offset 13 -> "First segment" and " of paragraph text."
    disc1 = discover_paragraph_targets(source)
    fp1 = disc1["targets"][0]["structure_fingerprint_sha256"]
    req1 = ParagraphRequest(
        source=source,
        destination=dest1,
        source_sha256=_sha256(source),
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp1,
        action="split",
        offset=13,
    )
    rep1 = author_paragraph(req1)
    assert rep1.passed is True

    # Step 2: discover on step 1 output, then split paragraph 2 at offset 4 -> " of " and "paragraph text."
    disc2 = discover_paragraph_targets(dest1)
    assert len(disc2["targets"]) == 2
    fp2 = disc2["targets"][1]["structure_fingerprint_sha256"]
    req2 = ParagraphRequest(
        source=dest1,
        destination=dest2,
        source_sha256=_sha256(dest1),
        target={"paragraph": 2},
        structure_fingerprint_sha256=fp2,
        action="split",
        offset=4,
    )
    rep2 = author_paragraph(req2)
    assert rep2.passed is True

    # Reopen and inspect final output: should have 3 paragraphs
    disc3 = discover_paragraph_targets(dest2)
    assert len(disc3["targets"]) == 3
    assert disc3["targets"][0]["text"] == "First segment"
    assert disc3["targets"][1]["text"] == " of "
    assert disc3["targets"][2]["text"] == "paragraph text."


def test_split_paragraph_preserves_style(tmp_path: Path) -> None:
    source = write_package(
        tmp_path / "styled.docx",
        paragraph("Heading text here", style="Heading"),
        parts={"word/styles.xml": _STYLES_WITH_HEADING},
    )
    disc = discover_paragraph_targets(source)
    assert disc["targets"][0]["style"] == "Heading"
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    dest = tmp_path / "styled_split.docx"
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=7,
    )
    rep = author_paragraph(req)
    assert rep.passed is True

    disc2 = discover_paragraph_targets(dest)
    assert len(disc2["targets"]) == 2
    assert disc2["targets"][0]["style"] == "Heading"
    assert disc2["targets"][1]["style"] == "Heading"


def test_split_paragraph_refusals(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "refusals.docx", paragraphs=["Refusal test"])
    source_sha = _sha256(source)
    disc = discover_paragraph_targets(source)
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    # 1. Invalid offset: offset > len(text)
    dest1 = tmp_path / "dest1.docx"
    req_oob = ParagraphRequest(
        source=source,
        destination=dest1,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=999,
    )
    rep_oob = author_paragraph(req_oob)
    assert rep_oob.passed is False
    assert rep_oob.status == "refused"
    assert rep_oob.code == "paragraph_offset_invalid"
    assert not dest1.exists()

    # 2. Stale source hash
    dest2 = tmp_path / "dest2.docx"
    req_hash = ParagraphRequest(
        source=source,
        destination=dest2,
        source_sha256="0" * 64,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=3,
    )
    rep_hash = author_paragraph(req_hash)
    assert rep_hash.passed is False
    assert rep_hash.code == "paragraph_source_hash_mismatch"
    assert not dest2.exists()

    # 3. Stale fingerprint
    dest3 = tmp_path / "dest3.docx"
    req_fp = ParagraphRequest(
        source=source,
        destination=dest3,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256="f" * 64,
        action="split",
        offset=3,
    )
    rep_fp = author_paragraph(req_fp)
    assert rep_fp.passed is False
    assert rep_fp.code == "paragraph_structure_fingerprint_mismatch"
    assert not dest3.exists()

    # 4. Occupied destination
    dest4 = tmp_path / "dest4.docx"
    dest4.write_text("occupied", encoding="utf-8")
    req_occ = ParagraphRequest(
        source=source,
        destination=dest4,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=3,
    )
    rep_occ = author_paragraph(req_occ)
    assert rep_occ.passed is False
    assert rep_occ.code == "paragraph_destination_occupied"


def test_split_paragraph_refusal_section_property(tmp_path: Path) -> None:
    p_with_sect = f"<w:p><w:pPr>{sect()}</w:pPr><w:r><w:t>Has sectPr</w:t></w:r></w:p>"
    source = write_package(tmp_path / "sect_p.docx", p_with_sect + paragraph("After"))
    source_sha = _sha256(source)
    dest = tmp_path / "dest_sect.docx"
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=3,
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "paragraph_section_property_unsupported"
    assert not dest.exists()


def test_split_paragraph_refusal_bookmark(tmp_path: Path) -> None:
    p_with_bm = (
        '<w:p><w:bookmarkStart w:id="1" w:name="bm"/><w:r><w:t>Bookmarked text</w:t></w:r>'
        '<w:bookmarkEnd w:id="1"/></w:p>'
    )
    source = write_package(tmp_path / "bm_p.docx", p_with_bm + paragraph("Tail"))
    source_sha = _sha256(source)
    dest = tmp_path / "dest_bm.docx"
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=4,
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "paragraph_bookmark_unsupported"
    assert not dest.exists()


def test_split_paragraph_refusal_field(tmp_path: Path) -> None:
    p_with_fld = '<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'
    source = write_package(tmp_path / "fld_p.docx", p_with_fld + paragraph("Tail"))
    source_sha = _sha256(source)
    dest = tmp_path / "dest_fld.docx"
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=0,
    )
    rep = author_paragraph(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "paragraph_field_unsupported"
    assert not dest.exists()


def test_split_paragraph_in_table_cell(tmp_path: Path) -> None:
    c1 = cell(paragraph("Cell P1 to split") + paragraph("Cell P2 untouched"))
    tbl = table(row(c1), grid=grid([5000]))
    source = write_package(tmp_path / "cell_split.docx", tbl + paragraph("Body tail"))
    source_sha = _sha256(source)
    dest = tmp_path / "cell_split_out.docx"

    disc = discover_paragraph_targets(source)
    assert len(disc["targets"]) == 3
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=7,  # "Cell P1" | " to split"
    )
    rep = author_paragraph(req)
    assert rep.passed is True
    assert dest.is_file()

    disc2 = discover_paragraph_targets(dest)
    assert len(disc2["targets"]) == 4
    assert disc2["targets"][0]["text"] == "Cell P1"
    assert disc2["targets"][1]["text"] == " to split"
    assert disc2["targets"][2]["text"] == "Cell P2 untouched"
    assert disc2["targets"][3]["text"] == "Body tail"


def test_split_paragraph_inline_controls_single_run_repro(tmp_path: Path) -> None:
    # Reproduce Main's public API report: <w:t>AB</w:t><w:tab/><w:t>CD</w:t> in one run
    p_repro = "<w:p><w:r><w:t>AB</w:t><w:tab/><w:t>CD</w:t></w:r></w:p>"
    source = write_package(tmp_path / "source_repro.docx", p_repro + paragraph("Tail"))
    source_sha = _sha256(source)
    disc = discover_paragraph_targets(source)
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    # Split at offset 1: "A" | "BCD"
    dest_1 = tmp_path / "split_offset_1.docx"
    req_1 = ParagraphRequest(
        source=source,
        destination=dest_1,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=1,
    )
    rep_1 = author_paragraph(req_1)
    assert rep_1.passed is True
    assert dest_1.is_file()

    # Inspect XML of dest_1
    graph_1 = PackageGraph.from_file(dest_1)
    doc_1 = ET.fromstring(graph_1.read_part_bytes(_DOCUMENT))
    paragraphs_1 = doc_1.findall(f".//{Q(W, 'p')}")
    p1 = paragraphs_1[0]
    p2 = paragraphs_1[1]

    p1_tabs = len(p1.findall(f".//{Q(W, 'tab')}"))
    p2_tabs = len(p2.findall(f".//{Q(W, 'tab')}"))
    assert p1_tabs == 0, f"Expected 0 tabs in p1, got {p1_tabs}"
    assert p2_tabs == 1, f"Expected 1 tab in p2, got {p2_tabs}"

    # Text verification
    p1_text = "".join(t.text or "" for t in p1.iter(Q(W, "t")))
    p2_text = "".join(t.text or "" for t in p2.iter(Q(W, "t")))
    assert p1_text == "A"
    assert p2_text == "BCD"

    # Verify source order in p2: run has 'B', then tab, then 'CD'
    p2_run = p2.find(Q(W, "r"))
    assert p2_run is not None
    child_tags = [c.tag.split("}")[-1] for c in p2_run]
    assert child_tags == ["t", "tab", "t"]
    assert p2_run[0].text == "B"
    assert p2_run[2].text == "CD"

    # Split at offset 3: "ABC" | "D" (tab belongs to p1)
    dest_3 = tmp_path / "split_offset_3.docx"
    req_3 = ParagraphRequest(
        source=source,
        destination=dest_3,
        source_sha256=source_sha,
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=3,
    )
    rep_3 = author_paragraph(req_3)
    assert rep_3.passed is True
    graph_3 = PackageGraph.from_file(dest_3)
    doc_3 = ET.fromstring(graph_3.read_part_bytes(_DOCUMENT))
    paragraphs_3 = doc_3.findall(f".//{Q(W, 'p')}")
    assert len(paragraphs_3[0].findall(f".//{Q(W, 'tab')}")) == 1
    assert len(paragraphs_3[1].findall(f".//{Q(W, 'tab')}")) == 0


def test_split_paragraph_inline_controls_differently_segmented_runs(
    tmp_path: Path,
) -> None:
    # Four documents with identical text "ABCD" and one tab between "AB" and "CD":
    # Doc 1: one run containing "AB", tab, "CD"
    p_doc1 = "<w:p><w:r><w:t>AB</w:t><w:tab/><w:t>CD</w:t></w:r></w:p>"
    # Doc 2: two runs, tab in first run (trailing)
    p_doc2 = "<w:p><w:r><w:t>AB</w:t><w:tab/></w:r><w:r><w:t>CD</w:t></w:r></w:p>"
    # Doc 3: two runs, tab in second run (leading)
    p_doc3 = "<w:p><w:r><w:t>AB</w:t></w:r><w:r><w:tab/><w:t>CD</w:t></w:r></w:p>"
    # Doc 4: three runs, tab in its own isolated run
    p_doc4 = "<w:p><w:r><w:t>AB</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>CD</w:t></w:r></w:p>"

    docs = [p_doc1, p_doc2, p_doc3, p_doc4]
    for idx, doc_xml in enumerate(docs, 1):
        src = write_package(tmp_path / f"segmented_src_{idx}.docx", doc_xml + paragraph("Tail"))
        dst = tmp_path / f"segmented_dst_{idx}.docx"
        fp = paragraph_authoring_structure_fingerprint(src, paragraph=1)
        assert fp is not None

        req = ParagraphRequest(
            source=src,
            destination=dst,
            source_sha256=_sha256(src),
            target={"paragraph": 1},
            structure_fingerprint_sha256=fp,
            action="split",
            offset=2,  # Split between "AB" and "CD"
        )
        rep = author_paragraph(req)
        assert rep.passed is True
        assert dst.is_file()

        # Verify all four documents produce exactly consistent partition:
        graph = PackageGraph.from_file(dst)
        doc = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
        paragraphs = doc.findall(f".//{Q(W, 'p')}")
        p1 = paragraphs[0]
        p2 = paragraphs[1]

        p1_text = "".join(t.text or "" for t in p1.iter(Q(W, "t")))
        p2_text = "".join(t.text or "" for t in p2.iter(Q(W, "t")))
        assert p1_text == "AB"
        assert p2_text == "CD"

        # Tab is preserved exactly once in p2 across all segmentations
        p1_tabs = len(p1.findall(f".//{Q(W, 'tab')}"))
        p2_tabs = len(p2.findall(f".//{Q(W, 'tab')}"))
        assert p1_tabs == 0, f"Doc {idx}: expected 0 tabs in p1, got {p1_tabs}"
        assert p2_tabs == 1, f"Doc {idx}: expected 1 tab in p2, got {p2_tabs}"


def test_split_paragraph_mixed_controls_in_run(tmp_path: Path) -> None:
    p_mixed = (
        "<w:p><w:r>"
        "<w:t>Hello</w:t>"
        "<w:br/>"
        "<w:tab/>"
        "<w:t>World</w:t>"
        "<w:noBreakHyphen/>"
        "<w:t>Test</w:t>"
        "</w:r></w:p>"
    )
    source = write_package(tmp_path / "mixed_ctrl_src.docx", p_mixed + paragraph("Tail"))
    dest = tmp_path / "mixed_ctrl_dst.docx"
    fp = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fp is not None

    # Split at offset 5: "Hello" | "WorldTest"
    req = ParagraphRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        target={"paragraph": 1},
        structure_fingerprint_sha256=fp,
        action="split",
        offset=5,
    )
    rep = author_paragraph(req)
    assert rep.passed is True
    assert dest.is_file()

    graph = PackageGraph.from_file(dest)
    doc = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    paragraphs = doc.findall(f".//{Q(W, 'p')}")
    p1 = paragraphs[0]
    p2 = paragraphs[1]

    assert "".join(t.text or "" for t in p1.iter(Q(W, "t"))) == "Hello"
    assert "".join(t.text or "" for t in p2.iter(Q(W, "t"))) == "WorldTest"

    # p1 has no controls; p2 has br, tab, noBreakHyphen exactly once in order
    assert len(p1.findall(f".//{Q(W, 'br')}")) == 0
    assert len(p1.findall(f".//{Q(W, 'tab')}")) == 0
    assert len(p1.findall(f".//{Q(W, 'noBreakHyphen')}")) == 0

    assert len(p2.findall(f".//{Q(W, 'br')}")) == 1
    assert len(p2.findall(f".//{Q(W, 'tab')}")) == 1
    assert len(p2.findall(f".//{Q(W, 'noBreakHyphen')}")) == 1


def test_split_paragraph_unsupported_run_content_refusal(tmp_path: Path) -> None:
    cases = [
        (
            "<w:p><w:r><w:drawing/></w:r></w:p>",
            "paragraph_drawing_unsupported",
            "ref_drawing.docx",
        ),
        (
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>',
            "paragraph_field_unsupported",
            "ref_fldchar.docx",
        ),
        (
            "<w:p><w:r><w:ruby/></w:r></w:p>",
            "paragraph_complex_unsupported",
            "ref_ruby.docx",
        ),
        (
            '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
            "paragraph_complex_unsupported",
            "ref_fn.docx",
        ),
    ]

    for p_xml, expected_code, out_name in cases:
        src = write_package(tmp_path / f"src_{out_name}", p_xml + paragraph("Tail"))
        dst = tmp_path / out_name
        fp = paragraph_authoring_structure_fingerprint(src, paragraph=1)
        assert fp is not None

        req = ParagraphRequest(
            source=src,
            destination=dst,
            source_sha256=_sha256(src),
            target={"paragraph": 1},
            structure_fingerprint_sha256=fp,
            action="split",
            offset=0,
        )
        rep = author_paragraph(req)
        assert rep.passed is False, f"Expected refusal for {out_name}"
        assert rep.status == "refused"
        assert rep.code == expected_code, f"Expected {expected_code}, got {rep.code}"
        assert not dst.exists(), f"Destination should not exist for refused request: {dst}"


@pytest.mark.parametrize("action", ["delete", "insert-before", "insert-after", "split"])
@pytest.mark.parametrize("empty", [False, True])
@pytest.mark.parametrize(
    ("encoding", "bom", "declaration"),
    [
        ("utf-8", b"\xef\xbb\xbf", "UTF-8"),
        ("utf-16-le", b"\xff\xfe", "UTF-16"),
        ("utf-16-be", b"\xfe\xff", "UTF-16"),
        ("iso-8859-1", b"", "ISO-8859-1"),
    ],
)
def test_paragraph_mutations_preserve_unowned_xml_bytes(
    tmp_path: Path, action: str, empty: bool, encoding: str, bom: bytes, declaration: str
) -> None:
    target = (
        '<alias:p w:rsidR="00000001" u:stamp="quoted > delimiter"/>'
        if empty
        else '<alias:p w:rsidR="00000001" u:stamp="quoted > delimiter">'
        "<w:r><w:t>caf&#233; AB</w:t></w:r></alias:p>"
    )
    prefix = (
        f'<?xml version="1.0" encoding="{declaration}" standalone="yes"?>\r\n'
        f'<w:document xmlns:w="{W}" xmlns:alias="{W}" xmlns="urn:opaque" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:u="urn:unknown" mc:Ignorable="u" mc:ProcessContent="u:wrapper" '
        'u:stamp="literal > delimiter"><w:body>'
        "<w:p><w:r><w:t>HEAD</w:t></w:r></w:p>\r\n"
    )
    gap = "\r\n<!-- not a paragraph: <w:p/> --><?keep data?>\r\n"
    suffix = (
        "<w:p><u:wrapper><w:r><w:t><![CDATA[KEEP <&>]]></w:t></w:r></u:wrapper></w:p>"
        "<w:sectPr /></w:body></w:document>\r\n"
    )
    before = bom + (prefix + target + gap + suffix).encode(encoding)
    source = write_package(
        tmp_path / "raw.docx", "", parts={_DOCUMENT: before, "custom/raw.bin": b"\x00KEEP\xff"}
    )
    source_sha = _sha256(source)
    destination = tmp_path / "edited.docx"
    fingerprint = paragraph_authoring_structure_fingerprint(source, paragraph=2)
    assert fingerprint is not None
    fields = {}
    operation = action
    expected = ["HEAD", "" if empty else "café AB", "KEEP <&>"]
    if action.startswith("insert-"):
        operation = "insert"
        fields = {"position": action.split("-")[1], "text": " new Ω & <> "}
        expected.insert(1 if action == "insert-before" else 2, fields["text"])
        if action == "insert-before":
            suffix = target + gap + suffix
        else:
            prefix += target + gap
    elif action == "delete":
        expected.pop(1)
        suffix = gap + suffix
    else:
        fields = {"offset": 0 if empty else 4}
        expected[1:2] = ["", ""] if empty else ["café", " AB"]
        suffix = gap + suffix
    report = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=destination,
            source_sha256=source_sha,
            structure_fingerprint_sha256=fingerprint,
            paragraph=2,
            action=operation,
            **fields,
        )
    )
    assert report.passed, report.to_dict()
    original = PackageGraph.from_file(source)
    edited = PackageGraph.from_file(destination)
    after = edited.read_part_bytes(_DOCUMENT)
    raw_prefix = bom + prefix.encode(encoding)
    raw_suffix = suffix.encode(encoding)
    assert after.startswith(raw_prefix)
    assert after.endswith(raw_suffix)
    if action == "delete":
        assert after == raw_prefix + raw_suffix
    assert ["".join(p.itertext()) for p in ET.fromstring(after).iter(Q(W, "p"))] == expected
    assert original.parts.keys() == edited.parts.keys()
    for part in original.parts:
        if part != _DOCUMENT:
            assert edited.read_part_bytes(part) == original.read_part_bytes(part)
    assert _sha256(source) == source_sha


def _rewrite_zip_part(path: Path, name: str, data: bytes) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    assert name in {info.filename for info, _ in entries}
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as rewritten:
        for info, original in entries:
            rewritten.writestr(info, data if info.filename == name else original)


def _changed_members(source: Path, output: Path) -> list[str]:
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        assert before.namelist() == after.namelist()
        return [name for name in before.namelist() if before.read(name) != after.read(name)]


_VERIFY_MODE_CASES = [
    ({"paragraph": 1}, {"action": "insert", "position": "after", "text": "Inserted."}),
    ({"paragraph": 2}, {"action": "insert", "position": "before", "text": "Before."}),
    ({"paragraph": 2}, {"action": "delete"}),
    ({"paragraph": 3}, {"action": "split", "offset": 5}),
    ({"paragraph": 5}, {"action": "insert", "position": "after", "text": "In a cell."}),
]


@pytest.mark.parametrize(("target", "edit"), _VERIFY_MODE_CASES)
def test_changed_parts_verification_publishes_the_same_package_as_full(
    tmp_path: Path, target: dict[str, int], edit: dict[str, object]
) -> None:
    source = _make_sample_package(
        tmp_path / "source.docx",
        include_table=True,
        extra_parts={"custom/part.xml": b"<customData><item>PreserveMe</item></customData>"},
    )
    source_sha = _sha256(source)
    fingerprint = paragraph_authoring_structure_fingerprint(source, paragraph=target["paragraph"])
    assert fingerprint is not None
    outputs: dict[str, bytes] = {}
    for verify in ("full", "changed-parts"):
        _clear_inspection_memo()
        dest = tmp_path / f"{verify}.docx"
        report = author_paragraph(
            ParagraphRequest(
                source=source,
                destination=dest,
                source_sha256=source_sha,
                target=target,
                structure_fingerprint_sha256=fingerprint,
                **edit,  # type: ignore[arg-type]
            ),
            verify=verify,
        )
        assert report.passed is True, report.code
        assert report.destination_sha256 == _sha256(dest)
        assert _changed_members(source, dest) == [_DOCUMENT]
        outputs[verify] = dest.read_bytes()
    assert outputs["changed-parts"] == outputs["full"]
    assert _sha256(source) == source_sha


@pytest.mark.parametrize("verify", ["full", "changed-parts"])
@pytest.mark.parametrize(
    ("part", "data"),
    [
        ("custom/part.xml", b"<customData><item>Tampered</item></customData>"),
        (
            "_rels/.rels",
            (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
                b'2006/relationships/officeDocument" Target="word/document.xml"/>'
                b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/'
                b'2006/relationships/hyperlink" Target="https://example.invalid/" '
                b'TargetMode="External"/></Relationships>'
            ),
        ),
    ],
    ids=["custom-part", "package-rels"],
)
def test_staged_change_outside_the_story_part_refuses_in_every_verify_mode(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, verify: str, part: str, data: bytes
) -> None:
    source = _make_sample_package(
        tmp_path / "source.docx",
        extra_parts={"custom/part.xml": b"<customData><item>PreserveMe</item></customData>"},
    )
    source_sha = _sha256(source)
    dest = tmp_path / "dest.docx"
    fingerprint = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fingerprint is not None
    write_staged = review._write_staged

    def write_tampered(source: Path, output: Path, *args: object) -> None:
        write_staged(source, output, *args)  # type: ignore[arg-type]
        _rewrite_zip_part(output, part, data)

    monkeypatch.setattr(review, "_write_staged", write_tampered)

    report = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=dest,
            source_sha256=source_sha,
            target={"paragraph": 1},
            structure_fingerprint_sha256=fingerprint,
            action="insert",
            position="after",
            text="Never published.",
        ),
        verify=verify,
    )

    assert report.passed is False
    assert report.code == "paragraph_staged_verification_failed"
    assert not dest.exists()
    assert _sha256(source) == source_sha
    assert sorted(path.name for path in tmp_path.iterdir()) == ["source.docx"]


def test_author_paragraph_rejects_unknown_verify_mode(tmp_path: Path) -> None:
    source = _make_sample_package(tmp_path / "source.docx")
    fingerprint = paragraph_authoring_structure_fingerprint(source, paragraph=1)
    assert fingerprint is not None
    request = ParagraphRequest(
        source=source,
        destination=tmp_path / "dest.docx",
        source_sha256=_sha256(source),
        target={"paragraph": 1},
        structure_fingerprint_sha256=fingerprint,
        action="delete",
    )
    with pytest.raises(ValueError, match="verify_mode_invalid"):
        author_paragraph(request, verify="skip")
