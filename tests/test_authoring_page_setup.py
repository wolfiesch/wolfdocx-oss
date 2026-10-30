"""Tests for source-bound section page setup authoring and discovery."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import (
    paragraph,
    sect,
    write_package,
)
from wolfdocx.inspect import Q
from wolfdocx.page_setup import (
    ColumnsSpec,
    MarginsSpec,
    PageNumberingSpec,
    PageSetupReport,
    PageSetupRequest,
    PageSizeSpec,
    author_page_setup,
    discover_page_setup_targets,
    load_page_setup_request,
    page_setup_structure_fingerprint,
    parse_page_setup_request,
)


def _sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def test_discover_page_setup_targets_single_section(tmp_path: Path) -> None:
    source = tmp_path / "single_sec.docx"
    write_package(
        source,
        paragraph("Hello world"),
        sect_properties=sect(width=12240, height=15840),
    )

    disc = discover_page_setup_targets(source)
    assert disc["passed"] is True
    assert disc["status"] == "discovered"
    assert disc["code"] is None
    assert disc["source_sha256"] == _sha256(source)
    assert len(disc["targets"]) == 1

    t0 = disc["targets"][0]
    assert t0["section"] == 1
    assert t0["part"] == "word/document.xml"
    assert t0["page_size"] == {"width": 12240, "height": 15840, "code": None}
    assert t0["orientation"] == "portrait"
    assert t0["margins"]["top"] == 1440
    assert t0["margins"]["bottom"] == 1440
    assert t0["margins"]["left"] == 1440
    assert t0["margins"]["right"] == 1440
    assert t0["columns"]["count"] == 1
    assert t0["columns"]["equal_width"] is True
    assert t0["usable_area"] == {"width": 12240 - 2880, "height": 15840 - 2880}
    assert t0["supported"] is True
    assert len(t0["structure_fingerprint_sha256"]) == 64


def test_discover_page_setup_targets_multi_section(tmp_path: Path) -> None:
    source = tmp_path / "multi_sec.docx"
    sec1 = sect(width=12240, height=15840, columns=1)
    sec2 = sect(width=11906, height=16838, columns=2)
    # Put section 1 in pPr/sectPr, section 2 as body sectPr
    body = f"{paragraph('Section 1 text', properties=sec1)}{paragraph('Section 2 text')}"
    write_package(source, body, sect_properties=sec2)

    disc = discover_page_setup_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 2

    t1 = disc["targets"][0]
    assert t1["section"] == 1
    assert t1["page_size"]["width"] == 12240
    assert t1["columns"]["count"] == 1
    t2 = disc["targets"][1]
    assert t2["section"] == 2
    assert t2["page_size"]["width"] == 11906
    assert t2["columns"]["count"] == 2


def test_page_setup_structure_fingerprint(tmp_path: Path) -> None:
    source1 = tmp_path / "source1.docx"
    write_package(source1, paragraph("Body copy"), sect_properties=sect(width=12240, height=15840))

    fp1 = page_setup_structure_fingerprint(source1, section=1)
    assert fp1 is not None
    assert len(fp1) == 64

    # Modifying paragraph text alone does NOT change section structure fingerprint
    source2 = tmp_path / "source2.docx"
    write_package(
        source2,
        paragraph("Completely different body text"),
        sect_properties=sect(width=12240, height=15840),
    )
    fp2 = page_setup_structure_fingerprint(source2, section=1)
    assert fp1 == fp2

    # Modifying sectPr attributes DOES change the fingerprint
    source3 = tmp_path / "source3.docx"
    write_package(
        source3,
        paragraph("Body copy"),
        sect_properties=sect(width=11906, height=16838),
    )
    fp3 = page_setup_structure_fingerprint(source3, section=1)
    assert fp3 != fp1

    # Out of bounds section
    assert page_setup_structure_fingerprint(source1, section=2) is None
    assert page_setup_structure_fingerprint(source1, section=0) is None
    assert page_setup_structure_fingerprint(tmp_path / "nonexistent.docx", section=1) is None


def test_author_page_setup_page_size_and_orientation_success(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(
        source,
        paragraph("Document text"),
        sect_properties=sect(width=12240, height=15840),
    )

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        page_size=PageSizeSpec(width=16838, height=11906),
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert isinstance(report, PageSetupReport)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True
    assert report.target_change_report is not None
    assert report.target_change_report["orientation"] == "landscape"
    assert report.target_change_report["page_size"]["width"] == 16838
    assert report.target_change_report["page_size"]["height"] == 11906
    assert dest.is_file()

    # Independent XML inspection
    with zipfile.ZipFile(dest) as archive:
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(doc_ns, 'sectPr')}")
        assert sect_pr is not None
        pg_sz = sect_pr.find(Q(doc_ns, "pgSz"))
        assert pg_sz is not None
        assert pg_sz.get(Q(doc_ns, "w")) == "16838"
        assert pg_sz.get(Q(doc_ns, "h")) == "11906"
        # w:orient is an ATTRIBUTE of w:pgSz, never a child element
        assert pg_sz.get(Q(doc_ns, "orient")) == "landscape"
        assert sect_pr.find(f".//{Q(doc_ns, 'orient')}") is None


def test_author_page_setup_margins_success(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(
        source,
        paragraph("Document text"),
        sect_properties=sect(top=1440, bottom=1440, left=1440, right=1440, header=720, footer=720),
    )

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        margins=MarginsSpec(top=1080, bottom=1080, left=1080, right=1080),
    )

    report = author_page_setup(req)
    assert report.passed is True
    assert report.status == "published"

    # Independent XML inspection: margins updated, header/footer preserved
    with zipfile.ZipFile(dest) as archive:
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(doc_ns, 'sectPr')}")
        assert sect_pr is not None
        pg_mar = sect_pr.find(Q(doc_ns, "pgMar"))
        assert pg_mar is not None
        assert pg_mar.get(Q(doc_ns, "top")) == "1080"
        assert pg_mar.get(Q(doc_ns, "bottom")) == "1080"
        assert pg_mar.get(Q(doc_ns, "left")) == "1080"
        assert pg_mar.get(Q(doc_ns, "right")) == "1080"
        # Preserved unspecified header and footer margins
        assert pg_mar.get(Q(doc_ns, "header")) == "720"
        assert pg_mar.get(Q(doc_ns, "footer")) == "720"


def test_author_page_setup_columns_success(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(
        source,
        paragraph("Document text"),
        sect_properties=sect(width=12240, height=15840, columns=1),
    )

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        columns=ColumnsSpec(count=2, spacing=720),
    )

    report = author_page_setup(req)
    assert report.passed is True
    assert report.status == "published"
    with zipfile.ZipFile(dest) as archive:
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(doc_ns, 'sectPr')}")
        assert sect_pr is not None
        cols = sect_pr.find(Q(doc_ns, "cols"))
        assert cols is not None
        assert cols.get(Q(doc_ns, "num")) == "2"
        assert cols.get(Q(doc_ns, "space")) == "720"
        assert cols.get(Q(doc_ns, "equalWidth")) == "1"


def test_author_page_setup_selected_section_isolation(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    sec1 = sect(width=12240, height=15840)
    sec2 = sect(width=11906, height=16838)
    body = f"{paragraph('Section 1 text', properties=sec1)}{paragraph('Section 2 text')}"
    write_package(source, body, sect_properties=sec2)

    src_sha = _sha256(source)
    fp1 = page_setup_structure_fingerprint(source, section=1)
    fp2_before = page_setup_structure_fingerprint(source, section=2)
    assert fp1 is not None
    assert fp2_before is not None

    # Mutate section 1 only
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp1,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is True

    # Section 2 structure must be completely unchanged in destination
    fp2_after = page_setup_structure_fingerprint(dest, section=2)
    assert fp2_after == fp2_before

    with zipfile.ZipFile(dest) as archive:
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        p_sect = doc_root.find(f".//{Q(doc_ns, 'pPr')}/{Q(doc_ns, 'sectPr')}")
        assert p_sect is not None
        assert p_sect.find(Q(doc_ns, "pgSz")).get(Q(doc_ns, "orient")) == "landscape"


def test_author_page_setup_source_staleness(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Initial text"), sect_properties=sect())

    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256="0" * 64,  # Stale hash
        structure_fingerprint_sha256=fp,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_source_hash_mismatch"
    assert not dest.exists()


def test_author_page_setup_occupied_output(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Initial text"), sect_properties=sect())
    dest.write_bytes(b"Occupied content")

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_destination_occupied"
    # Destination must NOT be clobbered
    assert dest.read_bytes() == b"Occupied content"


def test_author_page_setup_unsupported_context(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    # Section break in paragraph enclosed in an active comment range
    sect_xml = sect()
    body = (
        '<w:commentRangeStart w:id="1"/>'
        f"{paragraph('Text inside comment', properties=sect_xml)}"
        '<w:commentRangeEnd w:id="1"/>'
    )
    write_package(source, body)

    src_sha = _sha256(source)
    # Context check should prevent fingerprinting
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is None

    disc = discover_page_setup_targets(source)
    assert disc["passed"] is True
    assert disc["targets"][0]["supported"] is False

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256="a" * 64,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_context_unsupported"
    assert not dest.exists()

    # Also test tracked change on section itself (w:sectPrChange)
    source_rev = tmp_path / "source_rev.docx"
    rev_sect = (
        "<w:sectPr>"
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
        '<w:cols w:num="1"/>'
        '<w:sectPrChange w:id="2" w:author="editor" w:date="2026-09-17T00:00:00Z"/>'
        "</w:sectPr>"
    )
    write_package(source_rev, paragraph("Tracked change section"), sect_properties=rev_sect)
    assert page_setup_structure_fingerprint(source_rev, section=1) is None
    disc_rev = discover_page_setup_targets(source_rev)
    assert disc_rev["targets"][0]["supported"] is False


def test_author_page_setup_missing_explicit_geometry(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    # Section missing w:pgMar
    bad_sect = "<w:sectPr><w:pgSz w:w='12240' w:h='15840'/></w:sectPr>"
    write_package(source, paragraph("Body copy"), sect_properties=bad_sect)

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    # Request only updates orientation, so margins must be resolved from source
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_geometry_missing"
    assert not dest.exists()


def test_author_page_setup_unequal_columns_refusal(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    # Section with valid unequal-width columns (single cols element with unequal col children)
    unequal_sect = (
        "<w:sectPr>"
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
        '<w:cols w:equalWidth="0" w:num="2" w:space="720">'
        '<w:col w:w="4000" w:space="720"/>'
        '<w:col w:w="5360"/>'
        "</w:cols>"
        "</w:sectPr>"
    )
    write_package(source, paragraph("Body copy"), sect_properties=unequal_sect)
    before = source.read_bytes()
    with zipfile.ZipFile(source) as package:
        original_root = ET.fromstring(package.read("word/document.xml"))
    original_columns = ET.tostring(original_root.find(f".//{{{W}}}cols"))

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.non_target_parts_preserved is True
    assert source.read_bytes() == before
    with zipfile.ZipFile(dest) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    assert ET.tostring(root.find(f".//{{{W}}}cols")) == original_columns
    page_size = root.find(f".//{{{W}}}pgSz")
    assert page_size.get(f"{{{W}}}orient") == "landscape"
    assert page_size.get(f"{{{W}}}w") == "15840"
    assert page_size.get(f"{{{W}}}h") == "12240"


def test_author_page_setup_duplicate_geometry_refusal(tmp_path: Path) -> None:
    source = tmp_path / "source_dup.docx"
    dest = tmp_path / "dest_dup.docx"
    # Section with invalid duplicate cols element
    dup_cols_sect = (
        "<w:sectPr>"
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
        '<w:cols w:num="1"/>'
        '<w:cols w:num="2"/>'
        "</w:sectPr>"
    )
    write_package(source, paragraph("Body copy"), sect_properties=dup_cols_sect)

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_columns_unsupported"
    assert not dest.exists()


def test_author_page_setup_invalid_usable_area(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Body copy"), sect_properties=sect(width=12240, height=15840))

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    # Margins exceed page width (7000 + 7000 = 14000 > 12240)
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        margins=MarginsSpec(left=7000, right=7000),
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_usable_area_invalid"
    assert not dest.exists()


def test_author_page_setup_columns_exceed_usable_width(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    # Usable width = 12240 - 2880 = 9360 twips
    write_package(source, paragraph("Body copy"), sect_properties=sect(width=12240, height=15840))

    src_sha = _sha256(source)
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    # 3 columns with spacing 5000: (3 - 1) * 5000 = 10000 > 9360 usable width
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        columns=ColumnsSpec(count=3, spacing=5000),
    )

    report = author_page_setup(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "page_setup_columns_exceed_usable_width"
    assert not dest.exists()


def test_author_page_setup_source_byte_preservation_on_refusal(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Body copy"), sect_properties=sect())
    before_bytes = source.read_bytes()

    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256="f" * 64,  # will refuse
        structure_fingerprint_sha256="0" * 64,
        section=1,
        orientation="landscape",
    )

    report = author_page_setup(req)
    assert report.passed is False
    # Source bytes must be bit-for-bit preserved
    assert source.read_bytes() == before_bytes
    assert not dest.exists()


def test_strict_request_errors() -> None:
    # Non-mapping
    with pytest.raises(TypeError, match="page_setup_request_invalid"):
        parse_page_setup_request("not-a-mapping")

    # Missing required fields
    with pytest.raises(ValueError, match="page_setup_request_fields_invalid"):
        parse_page_setup_request({"source": "a.docx"})

    # Unknown fields / aliases rejected
    valid_base = {
        "source": "a.docx",
        "destination": "b.docx",
        "source_sha256": "a" * 64,
        "structure_fingerprint_sha256": "b" * 64,
        "section": 1,
        "orientation": "landscape",
    }
    with pytest.raises(ValueError, match="page_setup_request_fields_invalid"):
        parse_page_setup_request({**valid_base, "unknown_field": True})

    # Booleans strictly rejected
    with pytest.raises(ValueError, match="page_setup_section_invalid"):
        parse_page_setup_request({**valid_base, "section": True})

    with pytest.raises(ValueError, match="page_setup_width_invalid"):
        parse_page_setup_request({**valid_base, "page_size": {"width": True, "height": 10000}})

    with pytest.raises(ValueError, match="page_setup_top_margin_invalid"):
        parse_page_setup_request({**valid_base, "margins": {"top": False}})

    with pytest.raises(ValueError, match="page_setup_column_count_invalid"):
        parse_page_setup_request({**valid_base, "columns": {"count": True}})

    # Coordinates bounds
    with pytest.raises(ValueError, match="page_setup_width_invalid"):
        parse_page_setup_request(
            {**valid_base, "page_size": {"width": 100, "height": 10000}}
        )  # < 144 twips

    with pytest.raises(ValueError, match="page_setup_top_margin_invalid"):
        parse_page_setup_request({**valid_base, "margins": {"top": -50}})

    # Invalid orientation
    with pytest.raises(ValueError, match="page_setup_orientation_invalid"):
        parse_page_setup_request({**valid_base, "orientation": "diagonal"})

    # Empty request (no domain fields)
    empty_req = {
        "source": "a.docx",
        "destination": "b.docx",
        "source_sha256": "a" * 64,
        "structure_fingerprint_sha256": "b" * 64,
        "section": 1,
    }
    with pytest.raises(ValueError, match="page_setup_request_empty"):
        parse_page_setup_request(empty_req)


def test_load_page_setup_request(tmp_path: Path) -> None:
    req_file = tmp_path / "request.json"
    data = {
        "source": "source.docx",
        "destination": "dest.docx",
        "source_sha256": "a" * 64,
        "structure_fingerprint_sha256": "b" * 64,
        "section": 2,
        "orientation": "landscape",
        "page_size": {"width": 16838, "height": 11906},
        "margins": {"top": 1440, "bottom": 1440, "left": 1440, "right": 1440},
        "columns": {"count": 2, "spacing": 720},
    }
    req_file.write_text(json.dumps(data), encoding="utf-8")

    loaded = load_page_setup_request(req_file)
    assert loaded.source == Path("source.docx")
    assert loaded.destination == Path("dest.docx")
    assert loaded.section == 2
    assert loaded.orientation == "landscape"
    assert loaded.page_size == PageSizeSpec(width=16838, height=11906)
    assert loaded.margins == MarginsSpec(top=1440, bottom=1440, left=1440, right=1440)
    assert loaded.columns == ColumnsSpec(count=2, spacing=720)


def test_page_setup_reports_publication_and_occupied_refusal(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.docx"
    destination = tmp_path / "output.docx"
    write_package(
        source,
        paragraph("Preserved document content"),
        sect_properties=sect(width=12240, height=15840),
    )
    fingerprint = page_setup_structure_fingerprint(source, section=1)
    assert fingerprint is not None
    request = PageSetupRequest(
        source=source,
        destination=destination,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=fingerprint,
        page_size=PageSizeSpec(width=16838, height=11906),
        orientation="landscape",
    )
    request = parse_page_setup_request(request.to_dict())

    report = author_page_setup(request).to_dict()
    assert report["status"] == "published"
    assert report["passed"] is True
    assert report["code"] is None
    assert report["destination_sha256"] == _sha256(destination)
    assert report["target_change_report"]["orientation"] == "landscape"
    assert report["non_target_parts_preserved"] is True

    published_bytes = destination.read_bytes()
    refusal = author_page_setup(request).to_dict()
    assert refusal["status"] == "refused"
    assert refusal["passed"] is False
    assert refusal["code"] == "page_setup_destination_occupied"
    assert refusal["destination_sha256"] is None
    assert destination.read_bytes() == published_bytes


def test_page_numbering_spec_validation_and_parsing() -> None:
    valid = PageNumberingSpec(start=5, format="lowerRoman")
    assert valid.to_dict() == {"start": 5, "format": "lowerRoman"}

    clear_spec = PageNumberingSpec(clear=True)
    assert clear_spec.to_dict() == {"clear": True}

    clear_sub = PageNumberingSpec(clear_start=True, clear_format=True)
    assert clear_sub.to_dict() == {"clear_start": True, "clear_format": True}

    with pytest.raises(ValueError, match="page_numbering_start_invalid"):
        PageNumberingSpec(start=0)
    with pytest.raises(ValueError, match="page_numbering_start_invalid"):
        PageNumberingSpec(start=32768)
    with pytest.raises(ValueError, match="page_numbering_format_invalid"):
        PageNumberingSpec(format="invalid_fmt")
    with pytest.raises(ValueError, match="page_numbering_clear_conflict"):
        PageNumberingSpec(clear=True, start=1)
    with pytest.raises(ValueError, match="page_numbering_clear_conflict"):
        PageNumberingSpec(clear=True, format="decimal")
    with pytest.raises(ValueError, match="page_numbering_clear_conflict"):
        PageNumberingSpec(clear_start=True, start=1)
    with pytest.raises(ValueError, match="page_numbering_clear_conflict"):
        PageNumberingSpec(clear_format=True, format="decimal")


def test_author_page_setup_sets_and_clears_page_numbering(tmp_path: Path) -> None:
    source = tmp_path / "source_num.docx"
    write_package(
        source,
        paragraph("Section 1 body"),
        sect_properties=sect(width=12240, height=15840),
    )
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None

    # 1. Set start=10 and format=upperRoman
    dest1 = tmp_path / "dest1.docx"
    req1 = PageSetupRequest(
        source=source,
        destination=dest1,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=fp,
        section=1,
        page_numbering=PageNumberingSpec(start=10, format="upperRoman"),
    )
    rep1 = author_page_setup(req1)
    assert rep1.passed is True
    assert rep1.status == "published"
    assert rep1.page_setup is not None
    assert rep1.page_setup["page_numbering"] == {"start": 10, "format": "upperRoman"}

    # Verify XML content in dest1
    with zipfile.ZipFile(dest1, "r") as zf:
        root1 = ET.fromstring(zf.read("word/document.xml"))
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        sect_pr1 = root1.find(f".//{{{ns}}}sectPr")
        assert sect_pr1 is not None
        pg_num1 = sect_pr1.find(f"{{{ns}}}pgNumType")
        assert pg_num1 is not None
        assert pg_num1.get(f"{{{ns}}}start") == "10"
        assert pg_num1.get(f"{{{ns}}}fmt") == "upperRoman"

    # 2. Clear format only (clear_format=True, keeping start)
    fp1 = page_setup_structure_fingerprint(dest1, section=1)
    dest2 = tmp_path / "dest2.docx"
    req2 = PageSetupRequest(
        source=dest1,
        destination=dest2,
        source_sha256=_sha256(dest1),
        structure_fingerprint_sha256=fp1,
        section=1,
        page_numbering=PageNumberingSpec(clear_format=True),
    )
    rep2 = author_page_setup(req2)
    assert rep2.passed is True
    assert rep2.page_setup is not None
    assert rep2.page_setup["page_numbering"] == {"start": 10, "format": None}

    with zipfile.ZipFile(dest2, "r") as zf:
        root2 = ET.fromstring(zf.read("word/document.xml"))
        sect_pr2 = root2.find(f".//{{{ns}}}sectPr")
        pg_num2 = sect_pr2.find(f"{{{ns}}}pgNumType")
        assert pg_num2 is not None
        assert pg_num2.get(f"{{{ns}}}start") == "10"
        assert pg_num2.get(f"{{{ns}}}fmt") is None

    # 3. Clear entirely (clear=True)
    fp2 = page_setup_structure_fingerprint(dest2, section=1)
    dest3 = tmp_path / "dest3.docx"
    req3 = PageSetupRequest(
        source=dest2,
        destination=dest3,
        source_sha256=_sha256(dest2),
        structure_fingerprint_sha256=fp2,
        section=1,
        page_numbering=PageNumberingSpec(clear=True),
    )
    rep3 = author_page_setup(req3)
    assert rep3.passed is True
    assert rep3.page_setup is not None
    assert rep3.page_setup["page_numbering"] is None

    with zipfile.ZipFile(dest3, "r") as zf:
        root3 = ET.fromstring(zf.read("word/document.xml"))
        sect_pr3 = root3.find(f".//{{{ns}}}sectPr")
        assert sect_pr3.find(f"{{{ns}}}pgNumType") is None


def test_author_page_setup_schema_order_with_cols(tmp_path: Path) -> None:
    source = tmp_path / "source_cols.docx"
    write_package(
        source,
        paragraph("Cols body"),
        sect_properties=sect(width=12240, height=15840),
    )
    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None
    dest = tmp_path / "dest_cols.docx"
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=fp,
        section=1,
        columns=ColumnsSpec(count=2, spacing=720),
        page_numbering=PageNumberingSpec(start=1, format="decimal"),
    )
    rep = author_page_setup(req)
    assert rep.passed is True

    # Verify child order in sectPr: pgNumType must precede cols (rank 9 vs 10)
    with zipfile.ZipFile(dest, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        sect_pr = root.find(f".//{{{ns}}}sectPr")
        tags = [c.tag.split("}")[-1] for c in sect_pr]
        assert "pgNumType" in tags
        assert "cols" in tags
        assert tags.index("pgNumType") < tags.index("cols")


def test_author_page_setup_refuses_duplicate_pg_num_type(tmp_path: Path) -> None:
    source = tmp_path / "source_dup.docx"
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    sect_xml = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f'  <w:pgNumType w:start="1"/>'
        f'  <w:pgNumType w:start="2"/>'
        f"</w:sectPr>"
    )
    write_package(
        source,
        paragraph("Duplicate pgNumType"),
        sect_properties=sect_xml,
    )
    targets = discover_page_setup_targets(source)
    assert targets["passed"] is True
    assert targets["targets"][0]["supported"] is False

    fp = page_setup_structure_fingerprint(source, section=1)
    assert fp is not None
    dest = tmp_path / "dest_dup.docx"
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=fp,
        section=1,
        page_numbering=PageNumberingSpec(start=5),
    )
    rep = author_page_setup(req)
    assert rep.passed is False
    assert rep.status == "refused"
    assert rep.code == "page_setup_page_numbering_unsupported"


def test_discover_page_setup_targets_different_first_page(tmp_path: Path) -> None:
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # 1. Without titlePg -> different_first_page is False
    source_no = tmp_path / "source_no_title_pg.docx"
    write_package(
        source_no,
        paragraph("No title page"),
        sect_properties=sect(width=12240, height=15840),
    )
    disc_no = discover_page_setup_targets(source_no)
    assert disc_no["passed"] is True
    assert disc_no["targets"][0]["different_first_page"] is False

    # 2. With normative <w:titlePg/> -> different_first_page is True
    source_yes = tmp_path / "source_yes_title_pg.docx"
    sect_yes = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f"  <w:titlePg/>"
        f"</w:sectPr>"
    )
    write_package(
        source_yes,
        paragraph("Yes title page"),
        sect_properties=sect_yes,
    )
    disc_yes = discover_page_setup_targets(source_yes)
    assert disc_yes["passed"] is True
    assert disc_yes["targets"][0]["different_first_page"] is True

    # 3. With <w:titlePg w:val="0"/> -> different_first_page is False
    source_off = tmp_path / "source_off_title_pg.docx"
    sect_off = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f'  <w:titlePg w:val="0"/>'
        f"</w:sectPr>"
    )
    write_package(
        source_off,
        paragraph("Off title page"),
        sect_properties=sect_off,
    )
    disc_off = discover_page_setup_targets(source_off)
    assert disc_off["passed"] is True
    assert disc_off["targets"][0]["different_first_page"] is False

    # 4. With duplicate titlePg -> supported is False
    source_dup = tmp_path / "source_dup_title_pg.docx"
    sect_dup = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f"  <w:titlePg/>"
        f"  <w:titlePg/>"
        f"</w:sectPr>"
    )
    write_package(
        source_dup,
        paragraph("Dup title page"),
        sect_properties=sect_dup,
    )
    disc_dup = discover_page_setup_targets(source_dup)
    assert disc_dup["passed"] is True
    assert disc_dup["targets"][0]["supported"] is False


def test_author_page_setup_different_first_page_enable_disable_lifecycle(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source_lifecycle.docx"
    write_package(
        source,
        paragraph("Lifecycle test"),
        sect_properties=sect(width=12240, height=15840),
    )

    # Step 1: Enable different_first_page
    disc1 = discover_page_setup_targets(source)
    assert disc1["passed"] is True
    t1 = disc1["targets"][0]
    assert t1["different_first_page"] is False

    dest1 = tmp_path / "dest_enabled.docx"
    req1 = PageSetupRequest(
        source=source,
        destination=dest1,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
        section=1,
        different_first_page=True,
    )
    rep1 = author_page_setup(req1)
    assert rep1.passed is True
    assert rep1.status == "published"
    assert rep1.code is None
    assert rep1.destination_sha256 == _sha256(dest1)
    assert rep1.target_change_report is not None
    assert rep1.target_change_report["different_first_page"] is True

    # Inspect XML of dest1
    with zipfile.ZipFile(dest1, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        doc_ns = root.tag[1:].split("}", 1)[0]
        sect_pr = root.find(f".//{{{doc_ns}}}sectPr")
        assert sect_pr is not None
        title_pg = sect_pr.find(f"{{{doc_ns}}}titlePg")
        assert title_pg is not None
        assert not title_pg.attrib  # Normative: no attributes
        # Verify schema ordering: titlePg (rank 14) follows pgMar (5)
        tags = [c.tag.split("}")[-1] for c in sect_pr]
        assert tags.index("pgMar") < tags.index("titlePg")

    # Step 2: Reopen dest1 and disable different_first_page
    disc2 = discover_page_setup_targets(dest1)
    assert disc2["passed"] is True
    t2 = disc2["targets"][0]
    assert t2["different_first_page"] is True
    assert t2["structure_fingerprint_sha256"] != t1["structure_fingerprint_sha256"]

    dest2 = tmp_path / "dest_disabled.docx"
    req2 = PageSetupRequest(
        source=dest1,
        destination=dest2,
        source_sha256=_sha256(dest1),
        structure_fingerprint_sha256=t2["structure_fingerprint_sha256"],
        section=1,
        different_first_page=False,
    )
    rep2 = author_page_setup(req2)
    assert rep2.passed is True
    assert rep2.status == "published"
    assert rep2.code is None
    assert rep2.destination_sha256 == _sha256(dest2)
    assert rep2.target_change_report is not None
    assert rep2.target_change_report["different_first_page"] is False

    # Inspect XML of dest2: titlePg completely removed
    with zipfile.ZipFile(dest2, "r") as zf:
        root2 = ET.fromstring(zf.read("word/document.xml"))
        doc_ns2 = root2.tag[1:].split("}", 1)[0]
        sect_pr2 = root2.find(f".//{{{doc_ns2}}}sectPr")
        assert sect_pr2 is not None
        assert sect_pr2.find(f"{{{doc_ns2}}}titlePg") is None

    # Step 3: Discover on dest2
    disc3 = discover_page_setup_targets(dest2)
    assert disc3["passed"] is True
    assert disc3["targets"][0]["different_first_page"] is False


def test_author_page_setup_different_first_page_omission_preserves_state(
    tmp_path: Path,
) -> None:
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    source = tmp_path / "source_preserves.docx"
    sect_xml = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f"  <w:titlePg/>"
        f"</w:sectPr>"
    )
    write_package(
        source,
        paragraph("Omission preserves titlePg"),
        sect_properties=sect_xml,
    )
    disc = discover_page_setup_targets(source)
    assert disc["targets"][0]["different_first_page"] is True
    fp = disc["targets"][0]["structure_fingerprint_sha256"]

    # Update only orientation, omit different_first_page
    dest = tmp_path / "dest_preserved.docx"
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=fp,
        section=1,
        orientation="landscape",
        different_first_page=None,
    )
    rep = author_page_setup(req)
    assert rep.passed is True
    assert rep.target_change_report["different_first_page"] is True
    assert rep.target_change_report["orientation"] == "landscape"

    # Verify titlePg remains present in dest
    with zipfile.ZipFile(dest, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        doc_ns = root.tag[1:].split("}", 1)[0]
        sect_pr = root.find(f".//{{{doc_ns}}}sectPr")
        assert sect_pr.find(f"{{{doc_ns}}}titlePg") is not None


def test_author_page_setup_different_first_page_preserves_headers_footers_and_non_target(
    tmp_path: Path,
) -> None:
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    source = tmp_path / "source_multi.docx"

    # Section 1 with header reference, Section 2 without
    sec1_xml = (
        f'<w:sectPr xmlns:w="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'  <w:headerReference r:id="rIdHeader1" w:type="default"/>'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f"</w:sectPr>"
    )
    sec2_xml = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/>'
        f"</w:sectPr>"
    )

    header_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{ns}"><w:p><w:r><w:t>Header Text</w:t></w:r></w:p></w:hdr>'
    )
    body = f"<w:p><w:pPr>{sec1_xml}</w:pPr><w:r><w:t>Section 1</w:t></w:r></w:p><w:p><w:r><w:t>Section 2</w:t></w:r></w:p>"
    rel = '<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    override = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'

    write_package(
        source,
        body,
        sect_properties=sec2_xml,
        parts={"word/header1.xml": header_content},
        relationships=rel,
        overrides=override,
    )

    disc = discover_page_setup_targets(source)
    assert disc["passed"] is True
    assert len(disc["targets"]) == 2
    t1 = disc["targets"][0]
    t2 = disc["targets"][1]
    assert t1["different_first_page"] is False
    assert t2["different_first_page"] is False

    # Mutate section 1 to enable different_first_page
    dest = tmp_path / "dest_multi.docx"
    req = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=_sha256(source),
        structure_fingerprint_sha256=t1["structure_fingerprint_sha256"],
        section=1,
        different_first_page=True,
    )
    rep = author_page_setup(req)
    assert rep.passed is True
    assert rep.non_target_parts_preserved is True

    # Re-discover on dest
    disc_dest = discover_page_setup_targets(dest)
    assert disc_dest["targets"][0]["different_first_page"] is True
    assert disc_dest["targets"][1]["different_first_page"] is False
    # Section 2 fingerprint is identical
    assert (
        disc_dest["targets"][1]["structure_fingerprint_sha256"]
        == t2["structure_fingerprint_sha256"]
    )

    # Verify header1.xml and relationships are completely preserved
    with zipfile.ZipFile(dest, "r") as zf:
        assert zf.read("word/header1.xml") == header_content.encode("utf-8")
        doc_rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
        assert "rIdHeader1" in doc_rels
        # Verify section 1 has headerReference before titlePg
        root = ET.fromstring(zf.read("word/document.xml"))
        doc_ns = root.tag[1:].split("}", 1)[0]
        sec1_elem = root.findall(f".//{{{doc_ns}}}sectPr")[0]
        sec1_tags = [c.tag.split("}")[-1] for c in sec1_elem]
        assert "headerReference" in sec1_tags
        assert "titlePg" in sec1_tags
        assert sec1_tags.index("headerReference") < sec1_tags.index("titlePg")


def test_author_page_setup_different_first_page_refusals(tmp_path: Path) -> None:
    source = tmp_path / "source_refusals.docx"
    write_package(
        source,
        paragraph("Refusals test"),
        sect_properties=sect(width=12240, height=15840),
    )
    disc = discover_page_setup_targets(source)
    fp = disc["targets"][0]["structure_fingerprint_sha256"]
    src_sha = _sha256(source)
    dest = tmp_path / "dest_refused.docx"

    # 1. Invalid boolean in PageSetupRequest -> refused
    req_invalid_bool = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        different_first_page="yes",  # type: ignore[arg-type]
    )
    rep_invalid_bool = author_page_setup(req_invalid_bool)
    assert rep_invalid_bool.passed is False
    assert rep_invalid_bool.status == "refused"
    assert rep_invalid_bool.code == "page_setup_different_first_page_invalid"
    assert not dest.exists()

    # 2. Stale source hash -> refused
    req_stale = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256="0" * 64,
        structure_fingerprint_sha256=fp,
        section=1,
        different_first_page=True,
    )
    rep_stale = author_page_setup(req_stale)
    assert rep_stale.passed is False
    assert rep_stale.status == "refused"
    assert rep_stale.code == "page_setup_source_hash_mismatch"
    assert not dest.exists()

    # 3. Occupied destination -> refused
    dest.write_bytes(b"Occupied destination bytes")
    req_occ = PageSetupRequest(
        source=source,
        destination=dest,
        source_sha256=src_sha,
        structure_fingerprint_sha256=fp,
        section=1,
        different_first_page=True,
    )
    rep_occ = author_page_setup(req_occ)
    assert rep_occ.passed is False
    assert rep_occ.status == "refused"
    assert rep_occ.code == "page_setup_destination_occupied"
    assert dest.read_bytes() == b"Occupied destination bytes"
    dest.unlink()

    # 4. Duplicate titlePg in source -> refused
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    source_dup = tmp_path / "source_dup_title.docx"
    sect_dup = (
        f'<w:sectPr xmlns:w="{ns}">'
        f'  <w:pgSz w:w="12240" w:h="15840"/>'
        f'  <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        f"  <w:titlePg/>"
        f"  <w:titlePg/>"
        f"</w:sectPr>"
    )
    write_package(
        source_dup,
        paragraph("Duplicate titlePg"),
        sect_properties=sect_dup,
    )
    fp_dup = page_setup_structure_fingerprint(source_dup, section=1)
    assert fp_dup is not None
    req_dup = PageSetupRequest(
        source=source_dup,
        destination=dest,
        source_sha256=_sha256(source_dup),
        structure_fingerprint_sha256=fp_dup,
        section=1,
        different_first_page=True,
    )
    rep_dup = author_page_setup(req_dup)
    assert rep_dup.passed is False
    assert rep_dup.status == "refused"
    assert rep_dup.code == "page_setup_section_unsupported"
    assert not dest.exists()


def test_strict_request_different_first_page(tmp_path: Path) -> None:
    valid_base = {
        "source": "a.docx",
        "destination": "b.docx",
        "source_sha256": "a" * 64,
        "structure_fingerprint_sha256": "b" * 64,
        "section": 1,
    }

    # Invalid bool types in JSON parser
    with pytest.raises(ValueError, match="page_setup_different_first_page_invalid"):
        parse_page_setup_request({**valid_base, "different_first_page": "true"})

    with pytest.raises(ValueError, match="page_setup_different_first_page_invalid"):
        parse_page_setup_request({**valid_base, "different_first_page": 1})

    with pytest.raises(ValueError, match="page_setup_different_first_page_invalid"):
        parse_page_setup_request({**valid_base, "different_first_page": 0})

    # Valid boolean
    parsed = parse_page_setup_request({**valid_base, "different_first_page": True})
    assert parsed.different_first_page is True
    assert parsed.to_dict()["different_first_page"] is True

    parsed_false = parse_page_setup_request({**valid_base, "different_first_page": False})
    assert parsed_false.different_first_page is False
    assert parsed_false.to_dict()["different_first_page"] is False

    # JSON file load
    req_path = tmp_path / "req.json"
    req_path.write_text(
        json.dumps({**valid_base, "different_first_page": True}),
        encoding="utf-8",
    )
    loaded = load_page_setup_request(req_path)
    assert loaded.different_first_page is True
    assert loaded.to_dict()["different_first_page"] is True
