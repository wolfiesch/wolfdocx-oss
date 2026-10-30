"""Tests for source-bound section break authoring and discovery."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import (
    footer_reference,
    header_reference,
    paragraph,
    sect,
    write_package,
)
from wolfdocx.inspect import Q, inspect_docx
from wolfdocx.section_break import (
    SectionBreakReport,
    SectionBreakRequest,
    SectionBreakTarget,
    author_section_break,
    discover_section_break_targets,
    load_section_break_request,
    parse_section_break_request,
    section_break_anchor_fingerprint,
)


def _sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def _write_package_with_regions(path: Path, body: str, *, sect_properties: str) -> Path:
    """Use real header/footer relationships when testing section semantics."""
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    return write_package(
        path,
        body,
        sect_properties=sect_properties,
        parts={
            "word/header1.xml": f'<w:hdr xmlns:w="{ns}">{paragraph("Header")}</w:hdr>',
            "word/footer1.xml": f'<w:ftr xmlns:w="{ns}">{paragraph("Footer")}</w:ftr>',
        },
        relationships=(
            f'<Relationship Id="rIdH1" Type="{rel_ns}/header" Target="header1.xml"/>'
            f'<Relationship Id="rIdF1" Type="{rel_ns}/footer" Target="footer1.xml"/>'
        ),
        overrides=(
            '<Override PartName="/word/header1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            '<Override PartName="/word/footer1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        ),
    )


# ---------------------------------------------------------------------------
# Target Discovery Tests
# ---------------------------------------------------------------------------


def test_discover_section_break_targets_single_section(tmp_path: Path) -> None:
    source = tmp_path / "single_sec.docx"
    write_package(
        source,
        f"{paragraph('Paragraph 1')}{paragraph('Paragraph 2')}{paragraph('Paragraph 3')}",
        sect_properties=sect(width=12240, height=15840),
    )

    disc = discover_section_break_targets(source)
    assert disc["passed"] is True
    assert disc["status"] == "discovered"
    assert disc["code"] is None
    assert disc["operation"] == "section-break-targets"
    assert disc["source_sha256"] == _sha256(source)
    assert len(disc["targets"]) == 3

    t0 = disc["targets"][0]
    assert t0["insert_after"] == 1
    assert t0["section"] == 1
    assert t0["part"] == "word/document.xml"
    assert t0["break_type"] == "nextPage"
    assert t0["tag"] == "p"
    assert t0["text"] == "Paragraph 1"
    assert t0["supported"] is True
    assert len(t0["anchor_fingerprint_sha256"]) == 64

    # Matches section_break_anchor_fingerprint helper
    fp = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp == t0["anchor_fingerprint_sha256"]

    t1 = disc["targets"][1]
    assert t1["insert_after"] == 2
    assert t1["section"] == 1
    assert t1["text"] == "Paragraph 2"

    t2 = disc["targets"][2]
    assert t2["insert_after"] == 3
    assert t2["section"] == 1
    assert t2["text"] == "Paragraph 3"


def test_discover_section_break_targets_multi_section(tmp_path: Path) -> None:
    source = tmp_path / "multi_sec.docx"
    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Sec1 P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
        f"{paragraph('Sec2 P2')}"
    )
    write_package(source, body, sect_properties=sec2)

    disc = discover_section_break_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 4

    # Paragraph 1 is within Section 1 (inserting after it splits Section 1)
    assert disc["targets"][0]["section"] == 1
    assert disc["targets"][0]["break_type"] == "continuous"

    # Paragraph 2 terminates Section 1 (inserting after it splits Section 2)
    assert disc["targets"][1]["section"] == 2
    assert disc["targets"][1]["break_type"] == "nextPage"

    # Paragraphs 3 and 4 are within Section 2 (inserting after them splits Section 2)
    assert disc["targets"][2]["section"] == 2
    assert disc["targets"][2]["break_type"] == "nextPage"

    assert disc["targets"][3]["section"] == 2
    assert disc["targets"][3]["break_type"] == "nextPage"


def test_discover_section_break_targets_table_anchor(tmp_path: Path) -> None:
    source = tmp_path / "table_anchor.docx"
    tbl_xml = (
        '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )
    body = f"{paragraph('Before table')}{tbl_xml}{paragraph('After table')}"
    write_package(source, body, sect_properties=sect())

    disc = discover_section_break_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 3
    assert disc["targets"][0]["tag"] == "p"
    assert disc["targets"][1]["tag"] == "tbl"
    assert disc["targets"][1]["insert_after"] == 2
    assert disc["targets"][2]["tag"] == "p"


def test_discover_section_break_targets_refusals(tmp_path: Path) -> None:
    # Non-existent file
    disc1 = discover_section_break_targets(tmp_path / "nonexistent.docx")
    assert disc1["passed"] is False
    assert disc1["status"] == "refused"
    assert disc1["code"] == "section_break_source_missing"

    # Corrupt / invalid zip package
    corrupt = tmp_path / "corrupt.docx"
    corrupt.write_bytes(b"not a valid zip file content")
    disc2 = discover_section_break_targets(corrupt)
    assert disc2["passed"] is False
    assert disc2["code"] == "section_break_package_invalid"

    # Unsupported part
    source = tmp_path / "valid.docx"
    write_package(source, paragraph("Hello"), sect_properties=sect())
    disc3 = discover_section_break_targets(source, part="word/header1.xml")
    assert disc3["passed"] is False
    assert disc3["code"] == "section_break_part_unsupported"


def test_section_break_anchor_fingerprint_helper(tmp_path: Path) -> None:
    source = tmp_path / "fp_test.docx"
    write_package(
        source,
        f"{paragraph('First')}{paragraph('Second')}",
        sect_properties=sect(),
    )

    fp1 = section_break_anchor_fingerprint(source, insert_after=1)
    fp2 = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp1 is not None and len(fp1) == 64
    assert fp2 is not None and len(fp2) == 64
    assert fp1 != fp2

    # Out-of-bounds index
    assert section_break_anchor_fingerprint(source, insert_after=0) is None
    assert section_break_anchor_fingerprint(source, insert_after=3) is None
    assert section_break_anchor_fingerprint(tmp_path / "no.docx", insert_after=1) is None


# ---------------------------------------------------------------------------
# Authoring Tests: Multi-section Insertion & OOXML Semantics
# ---------------------------------------------------------------------------


def test_author_section_break_single_section_default_next_page(tmp_path: Path) -> None:
    source = tmp_path / "single_sec.docx"
    dest = tmp_path / "dest_single.docx"
    write_package(
        source,
        f"{paragraph('P1')}{paragraph('P2')}{paragraph('P3')}",
        sect_properties=sect(width=12240, height=15840, top=1440, bottom=1440),
    )

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
        break_type="nextPage",
    )
    report = author_section_break(req)

    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.destination_sha256 == _sha256(dest)
    assert report.non_target_parts_preserved is True
    assert inspect_docx(dest).valid is True
    # Check boundaries report
    assert report.boundaries is not None
    assert report.boundaries["previous_section"]["section_index"] == 1
    assert report.boundaries["previous_section"]["break_type"] == "nextPage"
    assert report.boundaries["following_section"]["section_index"] == 2
    assert report.boundaries["following_section"]["break_type"] == "nextPage"
    assert report.boundaries["inserted_paragraph_body_index"] == 2
    assert report.boundaries["total_sections"] == 2

    # Verify destination document XML structure
    with zipfile.ZipFile(dest) as archive:
        doc_xml = archive.read("word/document.xml")
    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None

    # Body elements: P1, New P (with sectPr), P2, P3, body-final sectPr
    children = list(body)
    assert len(children) == 5
    assert children[0].find(f".//{Q(ns, 't')}").text == "P1"

    # Inserted paragraph carrying cloned sectPr
    ins_p = children[1]
    ins_ppr = ins_p.find(Q(ns, "pPr"))
    assert ins_ppr is not None
    ins_sect = ins_ppr.find(Q(ns, "sectPr"))
    assert ins_sect is not None

    # Geometry preserved
    pg_sz = ins_sect.find(Q(ns, "pgSz"))
    assert pg_sz is not None
    assert pg_sz.get(Q(ns, "w")) == "12240"
    assert pg_sz.get(Q(ns, "h")) == "15840"

    # No rsid attributes on cloned sectPr
    for attr in ins_sect.attrib:
        local = attr.rsplit("}", 1)[-1] if "}" in attr else attr
        assert not local.lower().startswith("rsid")

    # Following section terminating sectPr (body-final) has w:type nextPage
    body_sect = children[4]
    assert body_sect.tag == Q(ns, "sectPr")
    type_elem = body_sect.find(Q(ns, "type"))
    assert type_elem is not None
    assert type_elem.get(Q(ns, "val")) == "nextPage"


def test_author_section_break_multi_section_earlier_section(tmp_path: Path) -> None:
    """Ensure insertion into section 1 resolves section 1's terminating sectPr, not body-final."""
    source = tmp_path / "two_sec.docx"
    dest = tmp_path / "dest_two.docx"

    sec1 = sect(
        width=12240,
        height=15840,
        references='<w:type w:val="continuous"/>',
    )
    sec2 = sect(
        width=11906,
        height=16838,
        references='<w:type w:val="nextPage"/>',
    )
    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Sec1 P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
        f"{paragraph('Sec2 P2')}"
    )
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    # Insert after Sec1 P1 (body index 1)
    fp = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
        break_type="evenPage",
    )
    report = author_section_break(req)
    assert report.passed is True

    # Boundaries
    assert report.boundaries["previous_section"]["section_index"] == 1
    assert report.boundaries["previous_section"]["break_type"] == "continuous"  # preserved!
    assert report.boundaries["following_section"]["section_index"] == 2
    assert report.boundaries["following_section"]["break_type"] == "evenPage"
    assert report.boundaries["total_sections"] == 3

    # Verify destination document XML structure
    with zipfile.ZipFile(dest) as archive:
        doc_xml = archive.read("word/document.xml")
    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None

    # 1. Inserted paragraph has continuous type and 12240x15840 geometry from section 1
    ins_p = body[1]
    ins_sect = ins_p.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert ins_sect is not None
    ins_type = ins_sect.find(Q(ns, "type"))
    assert ins_type is not None
    assert ins_type.get(Q(ns, "val")) == "continuous"
    ins_sz = ins_sect.find(Q(ns, "pgSz"))
    assert ins_sz.get(Q(ns, "w")) == "12240"

    # 2. Section 1's original terminating sectPr (on Sec1 P2) has type evenPage
    sec1_p2 = body[2]
    sec1_term = sec1_p2.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert sec1_term is not None
    sec1_type = sec1_term.find(Q(ns, "type"))
    assert sec1_type is not None
    assert sec1_type.get(Q(ns, "val")) == "evenPage"

    # 3. Section 2's body-final sectPr is completely untouched (still 11906x16838, nextPage)
    final_sect = body[-1]
    assert final_sect.tag == Q(ns, "sectPr")
    final_type = final_sect.find(Q(ns, "type"))
    assert final_type.get(Q(ns, "val")) == "nextPage"
    final_sz = final_sect.find(Q(ns, "pgSz"))
    assert final_sz.get(Q(ns, "w")) == "11906"


def test_author_section_break_multi_section_later_section(tmp_path: Path) -> None:
    """Ensure insertion into section 2 resolves section 2's terminating sectPr, leaving section 1 untouched."""
    source = tmp_path / "two_sec_later.docx"
    dest = tmp_path / "dest_later.docx"

    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Sec1 P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
        f"{paragraph('Sec2 P2')}"
    )
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    # Insert after Sec2 P1 (body index 3)
    fp = section_break_anchor_fingerprint(source, insert_after=3)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        insert_after=3,
        anchor_fingerprint_sha256=fp,
        break_type="oddPage",
    )
    report = author_section_break(req)
    assert report.passed is True
    assert report.boundaries["total_sections"] == 3
    assert report.boundaries["previous_section"]["section_index"] == 2
    assert report.boundaries["following_section"]["section_index"] == 3

    with zipfile.ZipFile(dest) as archive:
        doc_xml = archive.read("word/document.xml")
    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))

    # Section 1 (on Sec1 P2, body index 1) is completely untouched
    sec1_p2 = body[1]
    sec1_type = sec1_p2.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}/{Q(ns, 'type')}")
    assert sec1_type.get(Q(ns, "val")) == "continuous"

    # Inserted paragraph has 11906x16838 geometry from section 2
    ins_p = body[3]
    ins_sect = ins_p.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    ins_sz = ins_sect.find(Q(ns, "pgSz"))
    assert ins_sz.get(Q(ns, "w")) == "11906"

    # Final sectPr has oddPage
    final_type = body[-1].find(Q(ns, "type"))
    assert final_type.get(Q(ns, "val")) == "oddPage"


def test_author_section_break_at_existing_section_boundary(tmp_path: Path) -> None:
    """Ensure inserting after an anchor carrying pPr/sectPr preserves the anchor's sectPr byte-for-byte,
    splits the following section, and clones the following section's properties.
    """
    source = tmp_path / "boundary_anchor.docx"
    dest = tmp_path / "dest_boundary_anchor.docx"

    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Sec1 P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
        f"{paragraph('Sec2 P2')}"
    )
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    # Anchor is Sec1 P2 (body index 2), which terminates Section 1
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        insert_after=2,
        anchor_fingerprint_sha256=fp,
        break_type="evenPage",
    )
    report = author_section_break(req)
    assert report.passed is True

    # The insertion point splits Section 2:
    assert report.boundaries["previous_section"]["section_index"] == 2
    assert (
        report.boundaries["previous_section"]["break_type"] == "nextPage"
    )  # original Section 2 type!
    assert report.boundaries["following_section"]["section_index"] == 3
    assert report.boundaries["following_section"]["break_type"] == "evenPage"  # requested type!
    assert report.boundaries["total_sections"] == 3

    with zipfile.ZipFile(dest) as archive:
        doc_xml = archive.read("word/document.xml")
    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))
    assert body is not None

    # Section 1's terminating sectPr (on Sec1 P2, body index 1) is preserved byte-for-byte:
    sec1_p2 = body[1]
    sec1_sect = sec1_p2.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert sec1_sect is not None
    assert sec1_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "continuous"
    assert sec1_sect.find(Q(ns, "pgSz")).get(Q(ns, "w")) == "12240"
    assert sec1_sect.find(Q(ns, "pgSz")).get(Q(ns, "h")) == "15840"

    # Inserted paragraph (body index 2) carries a clone of Section 2's properties (11906x16838, nextPage)
    ins_p = body[2]
    ins_sect = ins_p.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert ins_sect is not None
    assert ins_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "nextPage"
    assert ins_sect.find(Q(ns, "pgSz")).get(Q(ns, "w")) == "11906"
    assert ins_sect.find(Q(ns, "pgSz")).get(Q(ns, "h")) == "16838"

    # Body-final sectPr (terminating Section 3) receives the requested break_type: evenPage
    final_sect = body[-1]
    assert final_sect.tag == Q(ns, "sectPr")
    assert final_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "evenPage"
    assert final_sect.find(Q(ns, "pgSz")).get(Q(ns, "w")) == "11906"


def test_author_section_break_preserves_geometry_and_headers_footers(
    tmp_path: Path,
) -> None:
    source = tmp_path / "headers_footers.docx"
    dest = tmp_path / "dest_headers_footers.docx"

    refs = f"{header_reference('rIdH1', 'default')}{footer_reference('rIdF1', 'default')}"
    sec_pr = sect(
        width=12240,
        height=15840,
        top=1000,
        bottom=1100,
        left=1200,
        right=1300,
        columns=2,
        references=refs,
    )
    parts = {
        "word/header1.xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header Text</w:t></w:r></w:p></w:hdr>',
        "word/footer1.xml": '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer Text</w:t></w:r></w:p></w:ftr>',
        "custom/audit.json": b'{"purpose": "preserve unrelated parts bytewise"}',
    }
    relationships = (
        '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        '<Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    )
    write_package(
        source,
        f"{paragraph('P1')}{paragraph('P2')}",
        sect_properties=sec_pr,
        parts=parts,
        relationships=relationships,
    )

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
        break_type="continuous",
    )
    report = author_section_break(req)
    assert report.passed is True
    assert report.non_target_parts_preserved is True

    # Verify both sections retain exact header/footer references and geometry
    with zipfile.ZipFile(dest) as archive:
        doc_xml = archive.read("word/document.xml")
        # Verify custom/audit.json and header/footer files are identical
        assert archive.read("custom/audit.json") == parts["custom/audit.json"]
        assert archive.read("word/header1.xml") == parts["word/header1.xml"].encode("utf-8")
        assert archive.read("word/footer1.xml") == parts["word/footer1.xml"].encode("utf-8")

    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = root.find(Q(ns, "body"))

    ins_sect = body[1].find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    final_sect = body[-1]

    for s in (ins_sect, final_sect):
        h_ref = s.find(Q(ns, "headerReference"))
        assert h_ref is not None
        assert (
            h_ref.get(
                Q(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "id",
                )
            )
            == "rIdH1"
        )

        f_ref = s.find(Q(ns, "footerReference"))
        assert f_ref is not None
        assert (
            f_ref.get(
                Q(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "id",
                )
            )
            == "rIdF1"
        )

        pg_sz = s.find(Q(ns, "pgSz"))
        assert pg_sz.get(Q(ns, "w")) == "12240"
        assert pg_sz.get(Q(ns, "h")) == "15840"

        pg_mar = s.find(Q(ns, "pgMar"))
        assert pg_mar.get(Q(ns, "top")) == "1000"
        assert pg_mar.get(Q(ns, "bottom")) == "1100"
        assert pg_mar.get(Q(ns, "left")) == "1200"
        assert pg_mar.get(Q(ns, "right")) == "1300"

        cols = s.find(Q(ns, "cols"))
        assert cols.get(Q(ns, "num")) == "2"


def test_author_section_break_all_valid_break_types(tmp_path: Path) -> None:
    """Verify all 4 valid OOXML break types are authored without refusal."""
    for b_type in ("nextPage", "continuous", "evenPage", "oddPage"):
        source = tmp_path / f"src_{b_type}.docx"
        dest = tmp_path / f"dest_{b_type}.docx"
        write_package(source, f"{paragraph('A')}{paragraph('B')}", sect_properties=sect())

        src_sha = _sha256(source)
        fp = section_break_anchor_fingerprint(source, insert_after=1)
        assert fp is not None

        req = SectionBreakRequest(
            source=source,
            destination=dest,
            source_sha256=src_sha,
            insert_after=1,
            anchor_fingerprint_sha256=fp,
            break_type=b_type,
        )
        report = author_section_break(req)
        assert report.passed is True
        assert report.status == "published"


# ---------------------------------------------------------------------------
# Strict Request Parsing and Dataclass Validation Tests
# ---------------------------------------------------------------------------


def test_strict_request_parsing_invalid_inputs(tmp_path: Path) -> None:
    valid_data = {
        "source": str(tmp_path / "s.docx"),
        "destination": str(tmp_path / "d.docx"),
        "source_sha256": "a" * 64,
        "insert_after": 1,
        "anchor_fingerprint_sha256": "b" * 64,
        "break_type": "nextPage",
    }

    # Non-mapping
    with pytest.raises(TypeError, match="section_break_request_invalid"):
        parse_section_break_request("not a dict")

    # Missing required field
    for req_field in (
        "source",
        "destination",
        "source_sha256",
        "insert_after",
        "anchor_fingerprint_sha256",
    ):
        d = dict(valid_data)
        del d[req_field]
        with pytest.raises(ValueError, match="section_break_request_fields_invalid"):
            parse_section_break_request(d)

    # Extra unknown field
    d = dict(valid_data)
    d["extra_field"] = "bad"
    with pytest.raises(ValueError, match="section_break_request_fields_invalid"):
        parse_section_break_request(d)

    # Invalid source
    d = dict(valid_data)
    d["source"] = ""
    with pytest.raises(ValueError, match="section_break_source_invalid"):
        parse_section_break_request(d)

    # Invalid destination
    d = dict(valid_data)
    d["destination"] = ""
    with pytest.raises(ValueError, match="section_break_destination_invalid"):
        parse_section_break_request(d)

    # Invalid source_sha256
    d = dict(valid_data)
    d["source_sha256"] = "short"
    with pytest.raises(ValueError, match="section_break_source_sha256_invalid"):
        parse_section_break_request(d)

    d = dict(valid_data)
    d["source_sha256"] = "Z" * 64  # non-hex
    with pytest.raises(ValueError, match="section_break_source_sha256_invalid"):
        parse_section_break_request(d)

    # Invalid anchor_fingerprint_sha256
    d = dict(valid_data)
    d["anchor_fingerprint_sha256"] = "123"
    with pytest.raises(ValueError, match="section_break_anchor_fingerprint_sha256_invalid"):
        parse_section_break_request(d)

    # Invalid insert_after
    for bad_ia in (0, -1, True, False, 1.5, "1"):
        d = dict(valid_data)
        d["insert_after"] = bad_ia
        with pytest.raises(ValueError, match="section_break_insert_after_invalid"):
            parse_section_break_request(d)

    # Invalid break_type
    for bad_bt in ("invalid", "col", "page", 123):
        d = dict(valid_data)
        d["break_type"] = bad_bt
        with pytest.raises(ValueError, match="section_break_break_type_invalid"):
            parse_section_break_request(d)


def test_direct_dataclass_validation(tmp_path: Path) -> None:
    """Direct SectionBreakRequest instantiation must validate as strictly as parse_section_break_request."""
    s = tmp_path / "s.docx"
    d = tmp_path / "d.docx"
    sha = "a" * 64
    fp = "b" * 64

    # Valid construction
    req = SectionBreakRequest(
        source=s,
        destination=d,
        source_sha256=sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    assert req.break_type == "nextPage"
    assert req.part == "word/document.xml"

    # String paths auto-converted to Path
    req2 = SectionBreakRequest(
        source=str(s),
        destination=str(d),
        source_sha256=sha,
        insert_after=2,
        anchor_fingerprint_sha256=fp,
    )
    assert isinstance(req2.source, Path)
    assert isinstance(req2.destination, Path)

    # Invalid types / values
    with pytest.raises(TypeError, match="section_break_source_invalid"):
        SectionBreakRequest(
            source=123,
            destination=d,
            source_sha256=sha,
            insert_after=1,
            anchor_fingerprint_sha256=fp,
        )

    with pytest.raises(TypeError, match="section_break_destination_invalid"):
        SectionBreakRequest(
            source=s,
            destination=123,
            source_sha256=sha,
            insert_after=1,
            anchor_fingerprint_sha256=fp,
        )

    with pytest.raises(ValueError, match="section_break_source_sha256_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256="bad",
            insert_after=1,
            anchor_fingerprint_sha256=fp,
        )

    with pytest.raises(ValueError, match="section_break_anchor_fingerprint_sha256_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            insert_after=1,
            anchor_fingerprint_sha256="bad",
        )

    with pytest.raises(ValueError, match="section_break_insert_after_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            insert_after=0,
            anchor_fingerprint_sha256=fp,
        )

    with pytest.raises(ValueError, match="section_break_insert_after_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            insert_after=True,
            anchor_fingerprint_sha256=fp,
        )

    with pytest.raises(ValueError, match="section_break_break_type_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            insert_after=1,
            anchor_fingerprint_sha256=fp,
            break_type="wrong",
        )


def test_load_section_break_request(tmp_path: Path) -> None:
    req_file = tmp_path / "request.json"
    req_data = {
        "source": str(tmp_path / "s.docx"),
        "destination": str(tmp_path / "d.docx"),
        "source_sha256": "f" * 64,
        "insert_after": 2,
        "anchor_fingerprint_sha256": "e" * 64,
        "break_type": "continuous",
    }
    req_file.write_text(json.dumps(req_data), encoding="utf-8")

    loaded = load_section_break_request(req_file)
    assert loaded.insert_after == 2
    assert loaded.break_type == "continuous"
    assert loaded.to_dict()["insert_after"] == 2


# ---------------------------------------------------------------------------
# Refusal Tests
# ---------------------------------------------------------------------------


def test_refusals_source_missing_and_hash_mismatch(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Text"), sect_properties=sect())

    real_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=1)

    # Missing source file
    req_missing = SectionBreakRequest(
        source=tmp_path / "nonexistent.docx",
        destination=dest,
        source_sha256=real_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    rep_missing = author_section_break(req_missing)
    assert rep_missing.passed is False
    assert rep_missing.code == "section_break_source_missing"

    # Source hash mismatch
    fake_sha = "0" * 64
    req_mismatch = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=fake_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    rep_mismatch = author_section_break(req_mismatch)
    assert rep_mismatch.passed is False
    assert rep_mismatch.code == "section_break_source_hash_mismatch"


def test_refusals_destination_occupied_and_invalid(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Text"), sect_properties=sect())

    real_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=1)

    # Destination occupied (already exists)
    dest.write_text("occupied")
    req_occ = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=real_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    rep_occ = author_section_break(req_occ)
    assert rep_occ.passed is False
    assert rep_occ.code == "section_break_destination_occupied"

    # Destination in non-existent directory
    req_inv = SectionBreakRequest(
        source=source,
        destination=tmp_path / "missing_dir" / "dest.docx",
        source_sha256=real_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    rep_inv = author_section_break(req_inv)
    assert rep_inv.passed is False
    assert rep_inv.code == "section_break_destination_invalid"

    # Destination == Source
    req_same = SectionBreakRequest(
        source=source,
        destination=source,
        source_sha256=real_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    rep_same = author_section_break(req_same)
    assert rep_same.passed is False
    assert rep_same.code == "section_break_destination_invalid"


def test_refusals_anchor_missing_and_stale(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, f"{paragraph('A')}{paragraph('B')}", sect_properties=sect())

    real_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=1)

    # Anchor missing (insert_after > 2)
    req_missing = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=real_sha,
        insert_after=99,
        anchor_fingerprint_sha256=fp,
    )
    rep_missing = author_section_break(req_missing)
    assert rep_missing.passed is False
    assert rep_missing.code == "section_break_anchor_missing"

    # Anchor stale (fingerprint does not match body item 1)
    stale_fp = "9" * 64
    req_stale = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=real_sha,
        insert_after=1,
        anchor_fingerprint_sha256=stale_fp,
    )
    rep_stale = author_section_break(req_stale)
    assert rep_stale.passed is False
    assert rep_stale.code == "section_break_anchor_stale"


def test_refusals_context_unsupported(tmp_path: Path) -> None:
    # 1. Anchor paragraph containing revision tag <w:ins>
    source1 = tmp_path / "ins_ctx.docx"
    ins_p = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:ins w:id="1" w:author="User" w:date="2026-09-17T00:00:00Z">'
        "<w:r><w:t>Tracked text</w:t></w:r>"
        "</w:ins>"
        "</w:p>"
    )
    write_package(source1, ins_p, sect_properties=sect())
    fp1 = section_break_anchor_fingerprint(source1, insert_after=1)
    req1 = SectionBreakRequest(
        source=source1,
        destination=tmp_path / "d1.docx",
        source_sha256=_sha256(source1),
        insert_after=1,
        anchor_fingerprint_sha256=fp1,
    )
    assert author_section_break(req1).code == "section_break_context_unsupported"

    # 2. Section properties containing <w:sectPrChange>
    source2 = tmp_path / "sect_change.docx"
    change_sect = sect(
        extra='<w:sectPrChange w:id="2" w:author="User" w:date="2026-09-17T00:00:00Z"/>'
    )
    write_package(source2, paragraph("Normal"), sect_properties=change_sect)
    fp2 = section_break_anchor_fingerprint(source2, insert_after=1)
    req2 = SectionBreakRequest(
        source=source2,
        destination=tmp_path / "d2.docx",
        source_sha256=_sha256(source2),
        insert_after=1,
        anchor_fingerprint_sha256=fp2,
    )
    assert author_section_break(req2).code == "section_break_context_unsupported"

    # 3. Duplicate <w:type> elements in section properties
    source3 = tmp_path / "dup_type.docx"
    dup_type_sect = sect(extra='<w:type w:val="nextPage"/><w:type w:val="continuous"/>')
    write_package(source3, paragraph("Normal"), sect_properties=dup_type_sect)
    fp3 = section_break_anchor_fingerprint(source3, insert_after=1)
    req3 = SectionBreakRequest(
        source=source3,
        destination=tmp_path / "d3.docx",
        source_sha256=_sha256(source3),
        insert_after=1,
        anchor_fingerprint_sha256=fp3,
    )
    assert author_section_break(req3).code == "section_break_context_unsupported"


def test_refusal_section_missing(tmp_path: Path) -> None:
    """Document with no sectPr anywhere returns section_break_section_missing."""
    source = tmp_path / "no_sect.docx"
    write_package(source, paragraph("No section properties"), sect_properties="")

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=tmp_path / "d_no_sect.docx",
        source_sha256=src_sha,
        insert_after=1,
        anchor_fingerprint_sha256=fp,
    )
    rep = author_section_break(req)
    assert rep.passed is False
    assert rep.code == "section_break_section_missing"


def test_report_and_target_dataclasses() -> None:
    tgt = SectionBreakTarget(
        insert_after=1,
        anchor_fingerprint_sha256="a" * 64,
        section=1,
        part="word/document.xml",
        break_type="nextPage",
        tag="p",
        text="hello",
        supported=True,
    )
    assert tgt.target == {"insert_after": 1, "part": "word/document.xml"}
    d = tgt.to_dict()
    assert d["insert_after"] == 1

    rep = SectionBreakReport(
        status="published",
        code=None,
        source_sha256="s" * 64,
        destination_sha256="d" * 64,
        target={"insert_after": 1, "part": "word/document.xml"},
    )
    assert rep.passed is True
    rep_dict = rep.to_dict()
    assert rep_dict["operation"] == "section-break"
    assert rep_dict["passed"] is True
    assert rep_dict["destination_sha256"] == "d" * 64


# ---------------------------------------------------------------------------
# Intermediate Section Break Deletion Tests
# ---------------------------------------------------------------------------


def test_discover_section_break_targets_with_deletable_break(tmp_path: Path) -> None:
    source = tmp_path / "deletable_sec.docx"
    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = f"{paragraph('Sec1 P1')}{paragraph('Sec1 P2', properties=sec1)}{paragraph('Sec2 P1')}"
    write_package(source, body, sect_properties=sec2)

    disc = discover_section_break_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 3

    t0 = disc["targets"][0]
    assert t0["insert_after"] == 1
    assert t0["has_break"] is False
    assert t0["can_delete"] is False
    assert t0["actions"] == ["insert"]

    t1 = disc["targets"][1]
    assert t1["insert_after"] == 2
    assert t1["has_break"] is True
    assert t1["can_delete"] is True
    assert t1["can_update"] is True
    assert t1["actions"] == ["insert", "delete", "update"]

    t2 = disc["targets"][2]
    assert t2["insert_after"] == 3
    assert t2["has_break"] is False
    assert t2["can_delete"] is False
    assert t2["actions"] == ["insert"]


def test_author_section_break_delete_success(tmp_path: Path) -> None:
    source = tmp_path / "two_sec.docx"
    dest = tmp_path / "dest_merged.docx"

    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Terminating Paragraph P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
    )
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        insert_after=2,
        anchor_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_section_break(req)

    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.action == "delete"
    assert report.destination_sha256 == _sha256(dest)
    assert report.non_target_parts_preserved is True
    assert inspect_docx(dest).valid is True

    # Boundaries check
    assert report.boundaries is not None
    assert report.boundaries["deleted_section_break_paragraph_index"] == 2
    assert report.boundaries["sections_before"] == 2
    assert report.boundaries["total_sections"] == 1
    assert report.boundaries["merged_section_index"] == 1

    # Verify destination document XML structure
    with zipfile.ZipFile(dest) as archive:
        doc_xml = archive.read("word/document.xml")
    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body_elem = root.find(Q(ns, "body"))
    assert body_elem is not None

    # Number of body items is PRESERVED (paragraph was NOT deleted!)
    items = [node for node in body_elem if node.tag in {Q(ns, "p"), Q(ns, "tbl")}]
    assert len(items) == 3

    # Terminating paragraph stays intact with its text and runs
    p2 = items[1]
    assert p2.tag == Q(ns, "p")
    t_elem = p2.find(f".//{Q(ns, 't')}")
    assert t_elem is not None
    assert t_elem.text == "Terminating Paragraph P2"

    # sectPr was removed from p2's pPr
    p2_ppr = p2.find(Q(ns, "pPr"))
    if p2_ppr is not None:
        assert p2_ppr.find(Q(ns, "sectPr")) is None

    # Only 1 section remains (the body-final sectPr from section 2)
    sects = [node for node in body_elem if node.tag == Q(ns, "sectPr")]
    assert len(sects) == 1
    body_sect = sects[0]

    # Succeeding section's geometry was preserved
    pg_sz = body_sect.find(Q(ns, "pgSz"))
    assert pg_sz is not None
    assert pg_sz.get(Q(ns, "w")) == "11906"
    assert pg_sz.get(Q(ns, "h")) == "16838"


def test_author_section_break_delete_with_paragraph_field_and_json(
    tmp_path: Path,
) -> None:
    source = tmp_path / "req_parse_src.docx"
    dest = tmp_path / "req_parse_dst.docx"
    req_file = tmp_path / "req.json"

    sec1 = sect(width=12240, height=15840)
    sec2 = sect(width=11906, height=16838)
    body = f"{paragraph('P1')}{paragraph('P2 with break', properties=sec1)}{paragraph('P3')}"
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    payload = {
        "source": str(source),
        "destination": str(dest),
        "source_sha256": src_sha,
        "paragraph": 2,
        "anchor_fingerprint_sha256": fp,
        "action": "delete",
    }
    req_file.write_text(json.dumps(payload), encoding="utf-8")

    parsed_req = load_section_break_request(req_file)
    assert parsed_req.action == "delete"
    assert parsed_req.paragraph == 2
    assert parsed_req.insert_after == 2

    report = author_section_break(parsed_req)
    assert report.passed is True
    assert report.status == "published"
    assert inspect_docx(dest).valid is True


def test_author_section_break_delete_multi_section_three_sections(
    tmp_path: Path,
) -> None:
    source = tmp_path / "three_sec.docx"
    dest = tmp_path / "dest_two_sec.docx"

    sec1 = sect(width=10000, height=14000)
    sec2 = sect(width=11000, height=15000)
    sec3 = sect(width=12000, height=16000)

    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Sec1 P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
        f"{paragraph('Sec2 P2', properties=sec2)}"
        f"{paragraph('Sec3 P1')}"
    )
    write_package(source, body, sect_properties=sec3)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp,
        action="delete",
    )
    report = author_section_break(req)
    assert report.passed is True
    assert report.boundaries["sections_before"] == 3
    assert report.boundaries["total_sections"] == 2

    with zipfile.ZipFile(dest) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body_elem = root.find(Q(ns, "body"))
    assert body_elem is not None

    # Sec1 P2 text is preserved and its sectPr is gone
    items = [node for node in body_elem if node.tag in {Q(ns, "p"), Q(ns, "tbl")}]
    assert len(items) == 5
    assert "".join(items[1].itertext()) == "Sec1 P2"
    assert items[1].find(f".//{Q(ns, 'sectPr')}") is None

    # Sec2 P2 still has its break (sec2 geometry 11000x15000)
    assert items[3].find(f".//{Q(ns, 'sectPr')}") is not None
    p4_sz = items[3].find(f".//{Q(ns, 'pgSz')}")
    assert p4_sz is not None
    assert p4_sz.get(Q(ns, "w")) == "11000"


def test_author_section_break_delete_refusals(tmp_path: Path) -> None:
    source = tmp_path / "refusal_src.docx"
    sec1 = sect(width=12240, height=15840)
    sec2 = sect(width=11906, height=16838)
    body = (
        f"{paragraph('Normal P1')}{paragraph('Break P2', properties=sec1)}{paragraph('Normal P3')}"
    )
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp_p2 = section_break_anchor_fingerprint(source, insert_after=2)
    fp_p1 = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp_p2 is not None and fp_p1 is not None

    # 1. Target paragraph has NO break (P1 has no sectPr)
    dest1 = tmp_path / "d1.docx"
    req_no_break = SectionBreakRequest(
        source=source,
        destination=dest1,
        source_sha256=src_sha,
        paragraph=1,
        anchor_fingerprint_sha256=fp_p1,
        action="delete",
    )
    rep1 = author_section_break(req_no_break)
    assert rep1.passed is False
    assert rep1.code == "section_break_target_missing"
    assert not dest1.exists()

    # 2. Stale anchor fingerprint
    dest2 = tmp_path / "d2.docx"
    req_stale = SectionBreakRequest(
        source=source,
        destination=dest2,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256="0" * 64,
        action="delete",
    )
    rep2 = author_section_break(req_stale)
    assert rep2.passed is False
    assert rep2.code == "section_break_anchor_stale"
    assert not dest2.exists()

    # 3. Destination already occupied
    dest3 = tmp_path / "d3.docx"
    dest3.write_text("occupied", encoding="utf-8")
    req_occ = SectionBreakRequest(
        source=source,
        destination=dest3,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_p2,
        action="delete",
    )
    rep3 = author_section_break(req_occ)
    assert rep3.passed is False
    assert rep3.code == "section_break_destination_occupied"

    # 4. Out-of-bounds paragraph index
    dest4 = tmp_path / "d4.docx"
    req_oob = SectionBreakRequest(
        source=source,
        destination=dest4,
        source_sha256=src_sha,
        paragraph=99,
        anchor_fingerprint_sha256=fp_p2,
        action="delete",
    )
    rep4 = author_section_break(req_oob)
    assert rep4.passed is False
    assert rep4.code == "section_break_anchor_missing"
    assert not dest4.exists()

    # 5. Stale source hash
    dest5 = tmp_path / "d5.docx"
    req_src_stale = SectionBreakRequest(
        source=source,
        destination=dest5,
        source_sha256="f" * 64,
        paragraph=2,
        anchor_fingerprint_sha256=fp_p2,
        action="delete",
    )
    rep5 = author_section_break(req_src_stale)
    assert rep5.passed is False
    assert rep5.code == "section_break_source_hash_mismatch"
    assert not dest5.exists()


def test_author_section_break_delete_header_footer_conflict_refusal(
    tmp_path: Path,
) -> None:
    source = tmp_path / "header_conflict.docx"
    dest = tmp_path / "d_hdr.docx"

    # Section 1 defines a header reference, but Section 2 does NOT (inherited-header context)
    sec1 = sect(
        width=12240,
        height=15840,
        references='<w:headerReference w:type="default" r:id="rIdH1"/>',
    )
    sec2 = sect(width=11906, height=16838)
    body = f"{paragraph('P1')}{paragraph('P2', properties=sec1)}{paragraph('P3')}"
    _write_package_with_regions(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp,
        action="delete",
    )
    rep = author_section_break(req)
    assert rep.passed is False
    assert rep.code == "section_break_context_unsupported"
    assert not dest.exists()


def test_direct_section_break_request_contradictory_aliases_rejected(
    tmp_path: Path,
) -> None:
    """Direct SectionBreakRequest construction rejects contradictory target/paragraph/insert_after aliases."""
    s = tmp_path / "s.docx"
    d = tmp_path / "d.docx"
    sha = "a" * 64
    fp = "b" * 64

    # 1. Contradictory insert_after and paragraph (action="delete")
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            insert_after=2,
            paragraph=4,
            action="delete",
        )

    # 2. Contradictory insert_after and paragraph (action="insert")
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            insert_after=2,
            paragraph=4,
            action="insert",
        )

    # 3. Contradictory target dict with both paragraph and insert_after
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            target={"paragraph": 4, "insert_after": 2},
            action="delete",
        )

    # 4. Contradictory target['paragraph'] and top-level insert_after
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            target={"paragraph": 4},
            insert_after=2,
            action="delete",
        )

    # 5. Contradictory target['paragraph'] and top-level paragraph
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            target={"paragraph": 4},
            paragraph=2,
            action="delete",
        )

    # 6. Contradictory target['insert_after'] and top-level paragraph
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            target={"insert_after": 4},
            paragraph=2,
            action="delete",
        )

    # 7. Contradictory target['insert_after'] and top-level insert_after
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            target={"insert_after": 4},
            insert_after=2,
            action="delete",
        )

    # 8. Non-mapping target
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        SectionBreakRequest(
            source=s,
            destination=d,
            source_sha256=sha,
            anchor_fingerprint_sha256=fp,
            target="not_a_mapping",
            action="delete",
        )


def test_section_break_request_json_parity_and_canonicalization(
    tmp_path: Path,
) -> None:
    """JSON parsing and direct construction produce identical canonical requests and reject contradictory aliases."""
    s = str(tmp_path / "s.docx")
    d = str(tmp_path / "d.docx")
    sha = "a" * 64
    fp = "b" * 64

    base_json = {
        "source": s,
        "destination": d,
        "source_sha256": sha,
        "anchor_fingerprint_sha256": fp,
    }

    # JSON route rejects contradictory aliases
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        parse_section_break_request(
            {**base_json, "insert_after": 2, "paragraph": 4, "action": "delete"}
        )

    with pytest.raises(ValueError, match="section_break_target_invalid"):
        parse_section_break_request(
            {**base_json, "insert_after": 2, "paragraph": 4, "action": "insert"}
        )

    with pytest.raises(ValueError, match="section_break_target_invalid"):
        parse_section_break_request(
            {
                **base_json,
                "target": {"paragraph": 4, "insert_after": 2},
                "action": "delete",
            }
        )

    with pytest.raises(ValueError, match="section_break_target_invalid"):
        parse_section_break_request(
            {
                **base_json,
                "target": {"paragraph": 4},
                "insert_after": 2,
                "action": "delete",
            }
        )

    with pytest.raises(ValueError, match="section_break_target_invalid"):
        parse_section_break_request({**base_json, "target": "not_a_mapping", "action": "delete"})

    # Matching aliases canonicalize identically across direct and JSON routes
    req_d_ia = SectionBreakRequest(
        source=Path(s),
        destination=Path(d),
        source_sha256=sha,
        anchor_fingerprint_sha256=fp,
        insert_after=2,
        action="delete",
    )
    assert req_d_ia.insert_after == 2
    assert req_d_ia.paragraph == 2

    req_d_p = SectionBreakRequest(
        source=Path(s),
        destination=Path(d),
        source_sha256=sha,
        anchor_fingerprint_sha256=fp,
        paragraph=2,
        action="delete",
    )
    assert req_d_p.insert_after == 2
    assert req_d_p.paragraph == 2

    req_d_both = SectionBreakRequest(
        source=Path(s),
        destination=Path(d),
        source_sha256=sha,
        anchor_fingerprint_sha256=fp,
        insert_after=2,
        paragraph=2,
        action="delete",
    )
    assert req_d_both.insert_after == 2
    assert req_d_both.paragraph == 2

    req_d_tgt = SectionBreakRequest(
        source=Path(s),
        destination=Path(d),
        source_sha256=sha,
        anchor_fingerprint_sha256=fp,
        target={"paragraph": 2},
        action="delete",
    )
    assert req_d_tgt.insert_after == 2
    assert req_d_tgt.paragraph == 2
    assert req_d_tgt.target == {"paragraph": 2, "insert_after": 2}

    # JSON route
    req_j_ia = parse_section_break_request({**base_json, "insert_after": 2, "action": "delete"})
    assert req_j_ia.insert_after == 2
    assert req_j_ia.paragraph == 2

    req_j_p = parse_section_break_request({**base_json, "paragraph": 2, "action": "delete"})
    assert req_j_p.insert_after == 2
    assert req_j_p.paragraph == 2

    req_j_both = parse_section_break_request(
        {**base_json, "insert_after": 2, "paragraph": 2, "action": "delete"}
    )
    assert req_j_both.insert_after == 2
    assert req_j_both.paragraph == 2

    # Full parity between direct and JSON
    assert req_d_both.to_dict() == req_j_both.to_dict()
    assert req_d_both.target_dict() == req_j_both.target_dict()
    assert req_d_ia.target_dict() == req_j_ia.target_dict()
    assert req_d_p.target_dict() == req_j_p.target_dict()


def test_author_section_break_delete_title_pg_effective_equality_regression(
    tmp_path: Path,
) -> None:
    """Effective titlePg equality handles absent, default true, explicit true, and explicit false."""
    hdr_ref = header_reference("rIdH1")

    # 1. Negative regression: section 1 explicit false (val="0"), section 2 explicit true (val="1")
    # Must refuse section_break_context_unsupported and leave source and destination unchanged
    src_neg1 = tmp_path / "title_pg_neg1.docx"
    dst_neg1 = tmp_path / "dst_neg1.docx"
    sec_false = sect(references=hdr_ref, extra='<w:titlePg w:val="0"/>')
    sec_true = sect(references=hdr_ref, extra='<w:titlePg w:val="1"/>')
    body_neg1 = f"{paragraph('P1')}{paragraph('P2 break', properties=sec_false)}{paragraph('P3')}"
    _write_package_with_regions(src_neg1, body_neg1, sect_properties=sec_true)
    src_neg1_sha = _sha256(src_neg1)
    fp_neg1 = section_break_anchor_fingerprint(src_neg1, insert_after=2)
    assert fp_neg1 is not None

    disc_neg1 = discover_section_break_targets(src_neg1)
    assert disc_neg1["targets"][1]["can_delete"] is False
    assert disc_neg1["targets"][1]["can_update"] is True
    assert disc_neg1["targets"][1]["actions"] == ["insert", "update"]

    req_neg1 = SectionBreakRequest(
        source=src_neg1,
        destination=dst_neg1,
        source_sha256=src_neg1_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_neg1,
        action="delete",
    )
    rep_neg1 = author_section_break(req_neg1)
    assert rep_neg1.passed is False
    assert rep_neg1.code == "section_break_context_unsupported"
    assert not dst_neg1.exists()
    assert _sha256(src_neg1) == src_neg1_sha

    # 2. Negative regression: section 1 explicit true (val="true"), section 2 explicit false (val="false")
    src_neg2 = tmp_path / "title_pg_neg2.docx"
    dst_neg2 = tmp_path / "dst_neg2.docx"
    sec_true_named = sect(references=hdr_ref, extra='<w:titlePg w:val="true"/>')
    sec_false_named = sect(references=hdr_ref, extra='<w:titlePg w:val="false"/>')
    body_neg2 = (
        f"{paragraph('P1')}{paragraph('P2 break', properties=sec_true_named)}{paragraph('P3')}"
    )
    _write_package_with_regions(src_neg2, body_neg2, sect_properties=sec_false_named)
    src_neg2_sha = _sha256(src_neg2)
    fp_neg2 = section_break_anchor_fingerprint(src_neg2, insert_after=2)
    assert fp_neg2 is not None

    disc_neg2 = discover_section_break_targets(src_neg2)
    assert disc_neg2["targets"][1]["can_delete"] is False

    req_neg2 = SectionBreakRequest(
        source=src_neg2,
        destination=dst_neg2,
        source_sha256=src_neg2_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_neg2,
        action="delete",
    )
    rep_neg2 = author_section_break(req_neg2)
    assert rep_neg2.passed is False
    assert rep_neg2.code == "section_break_context_unsupported"
    assert not dst_neg2.exists()
    assert _sha256(src_neg2) == src_neg2_sha

    # 3. Negative regression: section 1 default true (<w:titlePg/>), section 2 explicit false (val="0")
    src_neg3 = tmp_path / "title_pg_neg3.docx"
    dst_neg3 = tmp_path / "dst_neg3.docx"
    sec_default_true = sect(references=hdr_ref, title_page=True)
    body_neg3 = (
        f"{paragraph('P1')}{paragraph('P2 break', properties=sec_default_true)}{paragraph('P3')}"
    )
    _write_package_with_regions(src_neg3, body_neg3, sect_properties=sec_false)
    src_neg3_sha = _sha256(src_neg3)
    fp_neg3 = section_break_anchor_fingerprint(src_neg3, insert_after=2)
    assert fp_neg3 is not None

    disc_neg3 = discover_section_break_targets(src_neg3)
    assert disc_neg3["targets"][1]["can_delete"] is False

    req_neg3 = SectionBreakRequest(
        source=src_neg3,
        destination=dst_neg3,
        source_sha256=src_neg3_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_neg3,
        action="delete",
    )
    rep_neg3 = author_section_break(req_neg3)
    assert rep_neg3.passed is False
    assert rep_neg3.code == "section_break_context_unsupported"
    assert not dst_neg3.exists()
    assert _sha256(src_neg3) == src_neg3_sha

    # 4. Positive case: section 1 explicit false (val="0"), section 2 absent titlePg (title_page=False)
    # Both are effectively disabled (False == False) -> deletion permitted
    src_pos1 = tmp_path / "title_pg_pos1.docx"
    dst_pos1 = tmp_path / "dst_pos1.docx"
    sec_absent = sect(references=hdr_ref, title_page=False)
    body_pos1 = f"{paragraph('P1')}{paragraph('P2 break', properties=sec_false)}{paragraph('P3')}"
    _write_package_with_regions(src_pos1, body_pos1, sect_properties=sec_absent)
    src_pos1_sha = _sha256(src_pos1)
    fp_pos1 = section_break_anchor_fingerprint(src_pos1, insert_after=2)
    assert fp_pos1 is not None

    disc_pos1 = discover_section_break_targets(src_pos1)
    assert disc_pos1["targets"][1]["can_delete"] is True
    assert disc_pos1["targets"][1]["can_update"] is True
    assert disc_pos1["targets"][1]["actions"] == ["insert", "delete", "update"]

    req_pos1 = SectionBreakRequest(
        source=src_pos1,
        destination=dst_pos1,
        source_sha256=src_pos1_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_pos1,
        action="delete",
    )
    rep_pos1 = author_section_break(req_pos1)
    assert rep_pos1.passed is True
    assert rep_pos1.status == "published"
    assert dst_pos1.exists()

    # 5. Positive case: section 1 default true (<w:titlePg/>), section 2 explicit true (val="1")
    # Both are effectively enabled (True == True) -> deletion permitted
    src_pos2 = tmp_path / "title_pg_pos2.docx"
    dst_pos2 = tmp_path / "dst_pos2.docx"
    body_pos2 = (
        f"{paragraph('P1')}{paragraph('P2 break', properties=sec_default_true)}{paragraph('P3')}"
    )
    _write_package_with_regions(src_pos2, body_pos2, sect_properties=sec_true)
    src_pos2_sha = _sha256(src_pos2)
    fp_pos2 = section_break_anchor_fingerprint(src_pos2, insert_after=2)
    assert fp_pos2 is not None

    disc_pos2 = discover_section_break_targets(src_pos2)
    assert disc_pos2["targets"][1]["can_delete"] is True
    assert disc_pos2["targets"][1]["can_update"] is True
    assert disc_pos2["targets"][1]["actions"] == ["insert", "delete", "update"]

    req_pos2 = SectionBreakRequest(
        source=src_pos2,
        destination=dst_pos2,
        source_sha256=src_pos2_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_pos2,
        action="delete",
    )
    rep_pos2 = author_section_break(req_pos2)
    assert rep_pos2.passed is True
    assert rep_pos2.status == "published"
    assert dst_pos2.exists()


# ---------------------------------------------------------------------------
# Section Break Update Tests (commercial-docx-section-break-update-20260918)
# ---------------------------------------------------------------------------


def test_discover_section_break_targets_update_candidates(tmp_path: Path) -> None:
    """Discovery inventories can_update and actions correctly for break-bearing and break-free targets."""
    source = tmp_path / "disc_update.docx"
    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = f"{paragraph('P1')}{paragraph('P2', properties=sec1)}{paragraph('P3')}"
    write_package(source, body, sect_properties=sec2)

    disc = discover_section_break_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 3

    # P1: No break -> can_update=False, can_delete=False, actions=["insert"]
    t0 = disc["targets"][0]
    assert t0["has_break"] is False
    assert t0["can_update"] is False
    assert t0["can_delete"] is False
    assert t0["actions"] == ["insert"]

    # P2: Has break -> can_update=True, can_delete=True, actions=["insert", "delete", "update"]
    t1 = disc["targets"][1]
    assert t1["has_break"] is True
    assert t1["can_update"] is True
    assert t1["can_delete"] is True
    assert t1["actions"] == ["insert", "delete", "update"]

    # P3: No break -> can_update=False, can_delete=False, actions=["insert"]
    t2 = disc["targets"][2]
    assert t2["has_break"] is False
    assert t2["can_update"] is False
    assert t2["can_delete"] is False
    assert t2["actions"] == ["insert"]


def test_author_section_break_update_direct_and_json_canonicalization(
    tmp_path: Path,
) -> None:
    """Update works identically via direct SectionBreakRequest and parse_section_break_request with target aliases."""
    source = tmp_path / "update_canon_src.docx"
    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = f"{paragraph('P1')}{paragraph('P2', properties=sec1)}{paragraph('P3')}"
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    # 1. Direct construction with paragraph alias
    dest_direct = tmp_path / "dest_direct.docx"
    req_direct = SectionBreakRequest(
        source=source,
        destination=dest_direct,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp,
        break_type="nextPage",
        action="update",
    )
    assert req_direct.action == "update"
    assert req_direct.paragraph == 2
    assert req_direct.insert_after == 2
    assert req_direct.target_dict() == {
        "paragraph": 2,
        "insert_after": 2,
        "part": "word/document.xml",
    }

    rep_direct = author_section_break(req_direct)
    assert rep_direct.passed is True
    assert rep_direct.action == "update"
    assert rep_direct.boundaries["previous_break_type"] == "continuous"
    assert rep_direct.boundaries["break_type"] == "nextPage"
    assert rep_direct.boundaries["total_sections"] == 2
    assert rep_direct.target_change_report["action"] == "update"
    assert rep_direct.target_change_report["break_type"] == "nextPage"
    assert inspect_docx(dest_direct).valid is True

    # 2. JSON request with insert_after alias and target dict
    dest_json = tmp_path / "dest_json.docx"
    payload = {
        "source": str(source),
        "destination": str(dest_json),
        "source_sha256": src_sha,
        "target": {"paragraph": 2, "insert_after": 2},
        "anchor_fingerprint_sha256": fp,
        "break_type": "evenPage",
        "action": "update",
    }
    req_json = parse_section_break_request(payload)
    assert req_json.action == "update"
    assert req_json.paragraph == 2
    assert req_json.insert_after == 2
    assert req_json.break_type == "evenPage"

    rep_json = author_section_break(req_json)
    assert rep_json.passed is True
    assert rep_json.action == "update"
    assert rep_json.boundaries["previous_break_type"] == "continuous"
    assert rep_json.boundaries["break_type"] == "evenPage"
    assert inspect_docx(dest_json).valid is True

    # Verify evenPage in dest_json XML
    with zipfile.ZipFile(dest_json) as archive:
        doc_xml = archive.read("word/document.xml")
    root = ET.fromstring(doc_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    p2_sect = root.find(f".//{Q(ns, 'p')}[2]/{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert p2_sect is not None
    type_elem = p2_sect.find(Q(ns, "type"))
    assert type_elem is not None
    assert type_elem.get(Q(ns, "val")) == "evenPage"


def test_author_section_break_update_all_admitted_types(tmp_path: Path) -> None:
    """Update successfully transitions across all four admitted break types."""
    admitted = ["nextPage", "continuous", "evenPage", "oddPage"]
    current_src = tmp_path / "cycle_src.docx"
    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = f"{paragraph('Cycle P1')}{paragraph('Cycle P2', properties=sec1)}{paragraph('Cycle P3')}"
    write_package(current_src, body, sect_properties=sec2)

    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    for target_type in admitted:
        dest = tmp_path / f"dest_{target_type}.docx"
        src_sha = _sha256(current_src)
        fp = section_break_anchor_fingerprint(current_src, insert_after=2)
        assert fp is not None

        req = SectionBreakRequest(
            source=current_src,
            destination=dest,
            source_sha256=src_sha,
            paragraph=2,
            anchor_fingerprint_sha256=fp,
            break_type=target_type,
            action="update",
        )
        rep = author_section_break(req)
        assert rep.passed is True, f"Failed updating to {target_type}: {rep.code}"
        assert rep.code is None
        assert rep.action == "update"
        assert rep.boundaries["break_type"] == target_type
        assert inspect_docx(dest).valid is True

        # Verify exact XML type
        with zipfile.ZipFile(dest) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        p2_sect = root.find(f".//{Q(ns, 'p')}[2]/{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
        assert p2_sect is not None
        type_elem = p2_sect.find(Q(ns, "type"))
        assert type_elem is not None
        assert type_elem.get(Q(ns, "val")) == target_type

        # Advance to next iteration
        current_src = dest


def test_author_section_break_update_reopen_reupdate_lifecycle(tmp_path: Path) -> None:
    """Full update-reopen-reupdate lifecycle preserves paragraph/section counts and non-target properties."""
    doc0 = tmp_path / "doc0_initial.docx"
    hdr_ref = header_reference("rIdH1")
    ftr_ref = footer_reference("rIdF1")
    sec1 = sect(
        width=12240,
        height=15840,
        references=f'{hdr_ref}{ftr_ref}<w:type w:val="continuous"/>',
    )
    sec2 = sect(
        width=11906,
        height=16838,
        references=f'{hdr_ref}{ftr_ref}<w:type w:val="nextPage"/>',
    )
    body = f"{paragraph('Paragraph One')}{paragraph('Paragraph Two Break', properties=sec1)}{paragraph('Paragraph Three')}"
    _write_package_with_regions(doc0, body, sect_properties=sec2)

    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Step 1: Update continuous -> nextPage
    doc1 = tmp_path / "doc1_nextPage.docx"
    disc0 = discover_section_break_targets(doc0)
    assert disc0["passed"] is True
    fp0 = disc0["targets"][1]["anchor_fingerprint_sha256"]

    req1 = SectionBreakRequest(
        source=doc0,
        destination=doc1,
        source_sha256=_sha256(doc0),
        paragraph=2,
        anchor_fingerprint_sha256=fp0,
        break_type="nextPage",
        action="update",
    )
    rep1 = author_section_break(req1)
    assert rep1.passed is True
    assert rep1.boundaries["previous_break_type"] == "continuous"
    assert rep1.boundaries["break_type"] == "nextPage"
    assert rep1.boundaries["total_sections"] == 2
    assert inspect_docx(doc1).valid is True

    # Step 2: Reopen doc1, discover targets, update nextPage -> oddPage
    doc2 = tmp_path / "doc2_oddPage.docx"
    disc1 = discover_section_break_targets(doc1)
    assert disc1["passed"] is True
    assert disc1["targets"][1]["has_break"] is True
    assert disc1["targets"][1]["can_update"] is True
    fp1 = disc1["targets"][1]["anchor_fingerprint_sha256"]

    req2 = SectionBreakRequest(
        source=doc1,
        destination=doc2,
        source_sha256=_sha256(doc1),
        paragraph=2,
        anchor_fingerprint_sha256=fp1,
        break_type="oddPage",
        action="update",
    )
    rep2 = author_section_break(req2)
    assert rep2.passed is True
    assert rep2.boundaries["previous_break_type"] == "nextPage"
    assert rep2.boundaries["break_type"] == "oddPage"
    assert rep2.boundaries["total_sections"] == 2
    assert inspect_docx(doc2).valid is True

    # Step 3: Reopen doc2, discover targets, update oddPage -> evenPage
    doc3 = tmp_path / "doc3_evenPage.docx"
    disc2 = discover_section_break_targets(doc2)
    assert disc2["passed"] is True
    fp2 = disc2["targets"][1]["anchor_fingerprint_sha256"]

    req3 = SectionBreakRequest(
        source=doc2,
        destination=doc3,
        source_sha256=_sha256(doc2),
        paragraph=2,
        anchor_fingerprint_sha256=fp2,
        break_type="evenPage",
        action="update",
    )
    rep3 = author_section_break(req3)
    assert rep3.passed is True
    assert rep3.boundaries["previous_break_type"] == "oddPage"
    assert rep3.boundaries["break_type"] == "evenPage"
    assert rep3.boundaries["total_sections"] == 2
    assert inspect_docx(doc3).valid is True

    # Verify exact invariants in final doc3:
    with zipfile.ZipFile(doc3) as archive:
        root3 = ET.fromstring(archive.read("word/document.xml"))
    body3 = root3.find(Q(ns, "body"))
    assert body3 is not None

    # 3 paragraphs preserved with exact text
    p_nodes = body3.findall(Q(ns, "p"))
    assert len(p_nodes) == 3
    assert "".join(p_nodes[0].itertext()) == "Paragraph One"
    assert "".join(p_nodes[1].itertext()) == "Paragraph Two Break"
    assert "".join(p_nodes[2].itertext()) == "Paragraph Three"

    # Section 1 sectPr on P2:
    p2_sect = p_nodes[1].find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert p2_sect is not None
    assert p2_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "evenPage"
    p2_sz = p2_sect.find(Q(ns, "pgSz"))
    assert p2_sz.get(Q(ns, "w")) == "12240"
    assert p2_sz.get(Q(ns, "h")) == "15840"
    p2_hdr = p2_sect.find(Q(ns, "headerReference"))
    rels_id = (
        p2_hdr.get(
            Q(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "id",
            )
        )
        or p2_hdr.get(Q(ns, "id"))
        or p2_hdr.get("id")
    )
    assert rels_id is not None

    # Section 2 sectPr at body final:
    final_sect = body3.find(Q(ns, "sectPr"))
    assert final_sect is not None
    assert final_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "nextPage"
    final_sz = final_sect.find(Q(ns, "pgSz"))
    assert final_sz.get(Q(ns, "w")) == "11906"
    assert final_sz.get(Q(ns, "h")) == "16838"


def test_author_section_break_update_missing_type_element_initially(
    tmp_path: Path,
) -> None:
    """Updating a section break that initially lacks a w:type element inserts w:type in schema order."""
    source = tmp_path / "no_type_src.docx"
    dest = tmp_path / "with_type_dest.docx"
    # sec1 has NO w:type element; by ECMA-376 it defaults to nextPage
    sec1 = sect(width=12240, height=15840)
    sec2 = sect(width=11906, height=16838)
    body = f"{paragraph('P1')}{paragraph('P2', properties=sec1)}{paragraph('P3')}"
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp,
        break_type="continuous",
        action="update",
    )
    rep = author_section_break(req)
    assert rep.passed is True
    assert rep.boundaries["previous_break_type"] == "nextPage"
    assert rep.boundaries["break_type"] == "continuous"
    assert inspect_docx(dest).valid is True

    # Verify w:type was inserted and schema order is monotonic
    with zipfile.ZipFile(dest) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    p2_sect = root.find(f".//{Q(ns, 'p')}[2]/{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert p2_sect is not None
    type_elem = p2_sect.find(Q(ns, "type"))
    assert type_elem is not None
    assert type_elem.get(Q(ns, "val")) == "continuous"

    # Verify monotonic tag order in sectPr
    order_map = {
        "headerReference": 0,
        "footerReference": 0,
        "type": 3,
        "pgSz": 4,
        "pgMar": 5,
    }
    tags = [c.tag.rsplit("}", 1)[-1] for c in p2_sect]
    ranks = [order_map.get(t, 99) for t in tags]
    assert ranks == sorted(ranks)


def test_author_section_break_update_preserves_surrounding_sections(
    tmp_path: Path,
) -> None:
    """Updating section 1 break type preserves section 2 and section 3 exactly."""
    source = tmp_path / "three_sec_src.docx"
    dest = tmp_path / "three_sec_dest.docx"

    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="oddPage"/>')
    sec3 = sect(width=10000, height=14000, references='<w:type w:val="nextPage"/>')
    body = (
        f"{paragraph('Sec1 P1')}"
        f"{paragraph('Sec1 P2', properties=sec1)}"
        f"{paragraph('Sec2 P1')}"
        f"{paragraph('Sec2 P2', properties=sec2)}"
        f"{paragraph('Sec3 P1')}"
    )
    write_package(source, body, sect_properties=sec3)

    src_sha = _sha256(source)
    fp = section_break_anchor_fingerprint(source, insert_after=2)
    assert fp is not None

    req = SectionBreakRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp,
        break_type="evenPage",
        action="update",
    )
    rep = author_section_break(req)
    assert rep.passed is True
    assert rep.boundaries["section_index"] == 1
    assert rep.boundaries["break_type"] == "evenPage"
    assert rep.boundaries["total_sections"] == 3
    assert inspect_docx(dest).valid is True

    with zipfile.ZipFile(dest) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Total paragraphs = 5
    body_elem = root.find(Q(ns, "body"))
    assert body_elem is not None
    assert len(body_elem.findall(Q(ns, "p"))) == 5

    # Sec1 terminating sectPr on P2: type evenPage
    p2_sect = body_elem[1].find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert p2_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "evenPage"
    assert p2_sect.find(Q(ns, "pgSz")).get(Q(ns, "w")) == "12240"

    # Sec2 terminating sectPr on P4: type oddPage untouched
    p4_sect = body_elem[3].find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
    assert p4_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "oddPage"
    assert p4_sect.find(Q(ns, "pgSz")).get(Q(ns, "w")) == "11906"

    # Sec3 body-final sectPr: type nextPage untouched
    final_sect = body_elem.find(Q(ns, "sectPr"))
    assert final_sect.find(Q(ns, "type")).get(Q(ns, "val")) == "nextPage"
    assert final_sect.find(Q(ns, "pgSz")).get(Q(ns, "w")) == "10000"


def test_author_section_break_update_refusals_and_atomicity(tmp_path: Path) -> None:
    """Atomic refusals for update: stale hash, stale anchor, occupied dest, missing target, invalid break type."""
    source = tmp_path / "refusal_update_src.docx"
    sec1 = sect(width=12240, height=15840, references='<w:type w:val="continuous"/>')
    sec2 = sect(width=11906, height=16838, references='<w:type w:val="nextPage"/>')
    body = f"{paragraph('P1')}{paragraph('P2', properties=sec1)}{paragraph('P3')}"
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp_p2 = section_break_anchor_fingerprint(source, insert_after=2)
    fp_p1 = section_break_anchor_fingerprint(source, insert_after=1)
    assert fp_p2 is not None and fp_p1 is not None

    # 1. Target paragraph has NO break (P1 has no sectPr)
    d1 = tmp_path / "d1_ref.docx"
    req1 = SectionBreakRequest(
        source=source,
        destination=d1,
        source_sha256=src_sha,
        paragraph=1,
        anchor_fingerprint_sha256=fp_p1,
        break_type="nextPage",
        action="update",
    )
    rep1 = author_section_break(req1)
    assert rep1.passed is False
    assert rep1.code == "section_break_target_missing"
    assert not d1.exists()

    # 2. Stale anchor fingerprint
    d2 = tmp_path / "d2_ref.docx"
    req2 = SectionBreakRequest(
        source=source,
        destination=d2,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256="0" * 64,
        break_type="nextPage",
        action="update",
    )
    rep2 = author_section_break(req2)
    assert rep2.passed is False
    assert rep2.code == "section_break_anchor_stale"
    assert not d2.exists()

    # 3. Destination already occupied
    d3 = tmp_path / "d3_ref.docx"
    d3.write_text("occupied", encoding="utf-8")
    req3 = SectionBreakRequest(
        source=source,
        destination=d3,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_p2,
        break_type="nextPage",
        action="update",
    )
    rep3 = author_section_break(req3)
    assert rep3.passed is False
    assert rep3.code == "section_break_destination_occupied"

    # 4. Out-of-bounds paragraph index
    d4 = tmp_path / "d4_ref.docx"
    req4 = SectionBreakRequest(
        source=source,
        destination=d4,
        source_sha256=src_sha,
        paragraph=99,
        anchor_fingerprint_sha256=fp_p2,
        break_type="nextPage",
        action="update",
    )
    rep4 = author_section_break(req4)
    assert rep4.passed is False
    assert rep4.code == "section_break_anchor_missing"
    assert not d4.exists()

    # 5. Stale source hash
    d5 = tmp_path / "d5_ref.docx"
    req5 = SectionBreakRequest(
        source=source,
        destination=d5,
        source_sha256="e" * 64,
        paragraph=2,
        anchor_fingerprint_sha256=fp_p2,
        break_type="nextPage",
        action="update",
    )
    rep5 = author_section_break(req5)
    assert rep5.passed is False
    assert rep5.code == "section_break_source_hash_mismatch"
    assert not d5.exists()

    # 6. Invalid break type in SectionBreakRequest
    with pytest.raises(ValueError, match="section_break_break_type_invalid"):
        SectionBreakRequest(
            source=source,
            destination=tmp_path / "bad_type.docx",
            source_sha256=src_sha,
            paragraph=2,
            anchor_fingerprint_sha256=fp_p2,
            break_type="bogusType",
            action="update",
        )

    # 7. Destination invalid (same as source)
    req7 = SectionBreakRequest(
        source=source,
        destination=source,
        source_sha256=src_sha,
        paragraph=2,
        anchor_fingerprint_sha256=fp_p2,
        break_type="nextPage",
        action="update",
    )
    rep7 = author_section_break(req7)
    assert rep7.passed is False
    assert rep7.code == "section_break_destination_invalid"

    # 8. Contradictory target aliases in JSON
    with pytest.raises(ValueError, match="section_break_target_invalid"):
        parse_section_break_request(
            {
                "source": str(source),
                "destination": str(tmp_path / "bad_alias.docx"),
                "source_sha256": src_sha,
                "insert_after": 2,
                "paragraph": 3,
                "anchor_fingerprint_sha256": fp_p2,
                "break_type": "nextPage",
                "action": "update",
            }
        )

    # 9. Source file unchanged after all refusals
    assert _sha256(source) == src_sha
