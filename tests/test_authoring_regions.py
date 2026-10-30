"""Tests for source-bound section header and footer region discovery and authoring."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import (
    R,
    W,
    footer_reference,
    header_reference,
    paragraph,
    region_relationship,
    sect,
    settings_xml,
    write_package,
)
from wolfdocx.inspect import Q
from wolfdocx.regions import (
    RegionParagraph,
    RegionRequest,
    _find_sections,
    author_region,
    discover_region_targets,
    load_region_request,
    parse_region_request,
)


def _sha256(path: Path) -> str:
    import hashlib

    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def test_author_header_default_single_section(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Body copy"))

    disc_before = discover_region_targets(source)
    assert disc_before["passed"] is True
    assert disc_before["sections_count"] == 1
    sec = disc_before["sections"][0]
    assert sec["index"] == 1
    assert len(sec["references"]) == 0
    assert {"kind": "header", "variant": "default"} in sec["free_variants"]

    req = RegionRequest(
        source=source,
        source_sha256=disc_before["source_sha256"],
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Company Confidential", style="Heading")],
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.code is None
    assert report.destination_sha256 is not None
    assert report.non_target_parts_preserved is True
    assert dest.is_file()

    # Verify created part and XML content
    with zipfile.ZipFile(dest) as archive:
        assert "word/header1.xml" in archive.namelist()
        header_root = ET.fromstring(archive.read("word/header1.xml"))
        h_ns = header_root.tag[1:].split("}", 1)[0]
        assert "Company Confidential" in archive.read("word/header1.xml").decode("utf-8")
        style_elem = header_root.find(f".//{Q(h_ns, 'pStyle')}")
        assert style_elem is not None
        assert style_elem.get(Q(h_ns, "val")) == "Heading"

        # Check content type override
        types_xml = archive.read("[Content_Types].xml").decode("utf-8")
        assert 'PartName="/word/header1.xml"' in types_xml
        assert (
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"'
            in types_xml
        )

        # Check relationship
        rels_xml = archive.read("word/_rels/document.xml.rels").decode("utf-8")
        assert 'Target="header1.xml"' in rels_xml

        # Check sectPr reference
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(doc_ns, 'sectPr')}")
        assert sect_pr is not None
        ref = sect_pr.find(Q(doc_ns, "headerReference"))
        assert ref is not None
        assert ref.get(Q(doc_ns, "type")) == "default"

    # Verify discovery on staged output
    disc_after = discover_region_targets(dest)
    assert disc_after["passed"] is True
    sec_after = disc_after["sections"][0]
    assert len(sec_after["references"]) == 1
    ref = sec_after["references"][0]
    assert ref["kind"] == "header"
    assert ref["variant"] == "default"
    assert ref["part"] == "word/header1.xml"
    assert {"kind": "header", "variant": "default"} in sec_after["occupied_variants"]
    assert {"kind": "header", "variant": "default"} not in sec_after["free_variants"]


def test_author_footer_default_single_section(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Main text"))

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="footer",
        variant="default",
        paragraphs=[RegionParagraph(text="Page 1 of 10")],
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"
    assert dest.is_file()

    with zipfile.ZipFile(dest) as archive:
        assert "word/footer1.xml" in archive.namelist()
        footer_xml = archive.read("word/footer1.xml").decode("utf-8")
        assert "Page 1 of 10" in footer_xml

        types_xml = archive.read("[Content_Types].xml").decode("utf-8")
        assert 'PartName="/word/footer1.xml"' in types_xml
        assert (
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"'
            in types_xml
        )

        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(doc_ns, 'sectPr')}")
        assert sect_pr is not None
        ref = sect_pr.find(Q(doc_ns, "footerReference"))
        assert ref is not None
        assert ref.get(Q(doc_ns, "type")) == "default"


def test_author_first_page_variant_enables_title_pg(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    # Section without titlePg
    write_package(source, paragraph("Report Cover"), sect_properties=sect(title_page=False))

    disc_before = discover_region_targets(source)
    assert disc_before["sections"][0]["title_page"] is False

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="first",
        paragraphs=[RegionParagraph(text="First Page Only Header")],
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"

    with zipfile.ZipFile(dest) as archive:
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(ns, 'sectPr')}")
        assert sect_pr is not None

        # Assert w:titlePg is present in sectPr
        title_pg = sect_pr.find(Q(ns, "titlePg"))
        assert title_pg is not None

        # Assert w:headerReference with w:type="first" is present
        header_refs = sect_pr.findall(Q(ns, "headerReference"))
        assert any(ref.get(Q(ns, "type")) == "first" for ref in header_refs)

    disc_after = discover_region_targets(dest)
    assert disc_after["sections"][0]["title_page"] is True


def test_author_even_variant_enables_even_and_odd_settings(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    # Document with settings having evenAndOddHeaders=False
    write_package(
        source,
        paragraph("Body copy"),
        settings=settings_xml(evenAndOddHeaders=False),
    )

    disc_before = discover_region_targets(source)
    assert disc_before["even_and_odd_headers"] is False

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="footer",
        variant="even",
        paragraphs=[RegionParagraph(text="Even Page Footer")],
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"

    with zipfile.ZipFile(dest) as archive:
        settings_xml_content = archive.read("word/settings.xml").decode("utf-8")
        assert "evenAndOddHeaders" in settings_xml_content
        assert 'w:val="false"' not in settings_xml_content

        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(doc_ns, 'sectPr')}")
        assert sect_pr is not None
        ref = sect_pr.find(Q(doc_ns, "footerReference"))
        assert ref is not None
        assert ref.get(Q(doc_ns, "type")) == "even"

    disc_after = discover_region_targets(dest)
    assert disc_after["even_and_odd_headers"] is True


def test_multiple_sections_inheritance_and_independent_authoring(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    # Create 2 sections:
    # Section 1 has headerReference to rIdH1
    # Section 2 has no explicit headerReference (inherits from Section 1)
    header1_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Section 1 Header</w:t></w:r></w:p></w:hdr>'
    )
    s1_sect = sect(references=header_reference("rIdH1", "default"))
    s2_sect = sect()

    body = (
        f"<w:p><w:pPr>{s1_sect}</w:pPr><w:r><w:t>Section 1 content</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Section 2 content</w:t></w:r></w:p>"
    )
    parts = {"word/header1.xml": header1_content}
    rels = region_relationship("rIdH1", "header", "header1.xml")
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )

    write_package(
        source,
        body,
        sect_properties=s2_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )

    disc_before = discover_region_targets(source)
    assert disc_before["passed"] is True
    assert disc_before["sections_count"] == 2
    sec1 = disc_before["sections"][0]
    sec2 = disc_before["sections"][1]

    # Section 1 has explicit default header
    assert len(sec1["references"]) == 1
    assert sec1["references"][0]["variant"] == "default"
    assert sec1["references"][0]["part"] == "word/header1.xml"

    # Section 2 inherits default header from Section 1
    assert len(sec2["references"]) == 0
    assert len(sec2["inherited_references"]) == 1
    assert sec2["inherited_references"][0]["variant"] == "default"
    assert sec2["inherited_references"][0]["inherited_from_section"] == 1
    assert {"kind": "header", "variant": "default"} in sec2["free_variants"]

    # Author an independent default header on Section 2
    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=2,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Section 2 Custom Header")],
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"

    disc_after = discover_region_targets(dest)
    assert disc_after["passed"] is True
    assert disc_after["sections_count"] == 2

    # Section 1 retains header1.xml
    sec1_after = disc_after["sections"][0]
    assert sec1_after["references"][0]["part"] == "word/header1.xml"

    # Section 2 now has its own header2.xml and is unlinked
    sec2_after = disc_after["sections"][1]
    assert len(sec2_after["references"]) == 1
    assert sec2_after["references"][0]["part"] == "word/header2.xml"
    assert len(sec2_after["inherited_references"]) == 0


def test_refusal_when_variant_already_occupied(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    header_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Existing</w:t></w:r></w:p></w:hdr>'
    )
    write_package(
        source,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdH1", "default")),
        parts={"word/header1.xml": header_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides=(
            '<Override PartName="/word/header1.xml" ContentType="'
            'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        ),
    )

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Attempted Overwrite")],
    )

    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "region_variant_occupied"
    assert not dest.exists()


def test_refusal_when_style_missing(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Body text"))

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Header", style="NonExistentStyle404")],
    )

    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "region_style_missing"
    assert not dest.exists()


def test_refusal_when_source_hash_mismatch(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Body text"))

    req = RegionRequest(
        source=source,
        source_sha256="0" * 64,
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Header text")],
    )

    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "source_hash_mismatch"
    assert not dest.exists()


def test_refusal_when_destination_occupied(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Body text"))
    dest.write_bytes(b"EXISTING_CONTENT")

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Header text")],
    )

    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "region_destination_occupied"
    assert dest.read_bytes() == b"EXISTING_CONTENT"


def test_refusal_when_section_out_of_range(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"
    write_package(source, paragraph("Single section doc"))

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=5,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Header text")],
    )

    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "region_section_not_found"
    assert not dest.exists()


def test_byte_preservation_of_unrelated_parts(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    unrelated_payload = b"CUSTOM_SECRET_METADATA_UNTOUCHED_12345"
    parts = {"custom/unrelated.bin": unrelated_payload}
    write_package(source, paragraph("Preservation test"), parts=parts)

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Header text")],
    )

    report = author_region(req)
    assert report.passed is True
    assert report.non_target_parts_preserved is True

    # Check byte-for-byte identity of unowned part
    with zipfile.ZipFile(source) as src_zip, zipfile.ZipFile(dest) as dest_zip:
        assert dest_zip.read("custom/unrelated.bin") == unrelated_payload
        assert src_zip.read("custom/unrelated.bin") == dest_zip.read("custom/unrelated.bin")


def test_sect_pr_schema_ordering(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    # Start with a sectPr that has pgSz, pgMar, cols, docGrid
    custom_sect = (
        "<w:sectPr>"
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
        "<w:cols/>"
        '<w:docGrid w:linePitch="360"/>'
        "</w:sectPr>"
    )
    write_package(source, paragraph("Schema order check"), sect_properties=custom_sect)

    # Author first page header (which inserts both headerReference and titlePg)
    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="first",
        paragraphs=[RegionParagraph(text="First Header")],
    )

    report = author_region(req)
    assert report.passed is True

    with zipfile.ZipFile(dest) as archive:
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{Q(ns, 'sectPr')}")
        assert sect_pr is not None

        tags = [child.tag.rsplit("}", 1)[-1] for child in sect_pr]

        # In OOXML CT_SectPr:
        # headerReference MUST precede pgSz, pgMar, cols, titlePg, docGrid
        # titlePg MUST precede docGrid and follow cols
        hdr_idx = tags.index("headerReference")
        pg_sz_idx = tags.index("pgSz")
        cols_idx = tags.index("cols")
        title_pg_idx = tags.index("titlePg")
        doc_grid_idx = tags.index("docGrid")

        assert hdr_idx < pg_sz_idx
        assert cols_idx < title_pg_idx
        assert title_pg_idx < doc_grid_idx


def test_parse_and_load_region_request(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    write_package(source, paragraph("Parser test"))
    digest = _sha256(source)

    # Valid canonical dictionary with "section" positive integer and object paragraphs
    raw_dict = {
        "source": str(source),
        "source_sha256": digest,
        "destination": str(tmp_path / "out.docx"),
        "section": 1,
        "kind": "header",
        "variant": "default",
        "paragraphs": [{"text": "Hello", "style": "Heading"}],
    }
    req1 = parse_region_request(raw_dict)
    assert req1.section == 1
    assert req1.section_index == 1
    assert req1.kind == "header"
    assert req1.variant == "default"
    assert len(req1.paragraphs) == 1
    assert req1.paragraphs[0].text == "Hello"
    assert req1.paragraphs[0].style == "Heading"

    # Reject section object synonym
    with pytest.raises(TypeError, match="region_section_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": {"index": 1},
            }
        )

    # Reject section_index alias synonym
    with pytest.raises(ValueError, match="region_request_fields_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section_index": 1,
            }
        )

    # Reject unknown top-level keys
    with pytest.raises(ValueError, match="region_request_fields_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "extra_field": "disallowed",
            }
        )

    # Reject string shorthand in paragraphs
    with pytest.raises(TypeError, match="region_paragraph_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "paragraphs": ["string shorthand"],
            }
        )

    # Reject unknown keys in paragraph objects
    with pytest.raises(ValueError, match="region_paragraph_fields_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "paragraphs": [{"text": "Hello", "unsupported": 123}],
            }
        )

    # Reject missing paragraph text
    with pytest.raises(ValueError, match="region_paragraph_fields_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "paragraphs": [{"style": "Heading"}],
            }
        )

    # Invalid non-dict
    with pytest.raises(TypeError, match="region_request_invalid"):
        parse_region_request("not a mapping")

    # Missing required section field
    with pytest.raises(ValueError, match="region_request_fields_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
            }
        )

    # Bad section value (<= 0)
    with pytest.raises(ValueError, match="region_section_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 0,
            }
        )

    # Bad kind
    with pytest.raises(ValueError, match="region_kind_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "kind": "sidebar",
            }
        )

    # Bad variant
    with pytest.raises(ValueError, match="region_variant_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "kind": "header",
                "variant": "watermark",
            }
        )

    # Load from disk
    json_file = tmp_path / "request.json"
    json_file.write_text(json.dumps(raw_dict), encoding="utf-8")
    loaded_req = load_region_request(json_file)
    assert loaded_req.section == 1
    assert loaded_req.section_index == 1
    assert loaded_req.kind == "header"


def test_replace_shared_header_allocates_isolated_story_and_preserves_sibling(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    # Two sections share word/header1.xml
    header1_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Shared Header Text</w:t></w:r></w:p></w:hdr>'
    )
    s1_sect = sect(references=header_reference("rIdH1", "default"))
    s2_sect = sect(references=header_reference("rIdH1", "default"))

    body = (
        f"<w:p><w:pPr>{s1_sect}</w:pPr><w:r><w:t>Section 1 body</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Section 2 body</w:t></w:r></w:p>"
    )
    parts = {"word/header1.xml": header1_content}
    rels = region_relationship("rIdH1", "header", "header1.xml")
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )

    write_package(
        source,
        body,
        sect_properties=s2_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )

    disc_before = discover_region_targets(source)
    assert disc_before["sections"][0]["references"][0]["part"] == "word/header1.xml"
    assert disc_before["sections"][1]["references"][0]["part"] == "word/header1.xml"

    # Replace Section 2 default header
    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=2,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="Section 2 Isolated Header", style="Heading")],
        action="replace",
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.region is not None
    assert report.region["isolated"] is True
    assert report.region["part"] == "word/header2.xml"

    with zipfile.ZipFile(dest) as archive:
        # Original shared part must remain unchanged
        assert archive.read("word/header1.xml") == header1_content.encode("utf-8")
        # New isolated part contains replacement paragraphs
        h2_content = archive.read("word/header2.xml").decode("utf-8")
        assert "Section 2 Isolated Header" in h2_content

        # Section 1 points to header1, Section 2 points to header2
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sects = _find_sections(doc_root, doc_ns)
        assert len(sects) == 2

        # Verify relationships
        rels_root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        r_ns = rels_root.tag[1:].split("}", 1)[0]
        rels_map = {
            r.get("Id"): r.get("Target") for r in rels_root.findall(f"{{{r_ns}}}Relationship")
        }
        s1_ref = sects[0].find(f"{{{doc_ns}}}headerReference")
        s2_ref = sects[1].find(f"{{{doc_ns}}}headerReference")
        assert s1_ref is not None and s2_ref is not None
        assert rels_map[s1_ref.get(f"{{{R}}}id")] == "header1.xml"
        assert rels_map[s2_ref.get(f"{{{R}}}id")] == "header2.xml"


def test_replace_unshared_header_in_place(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    header_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Old Header</w:t></w:r></w:p></w:hdr>'
    )
    s1_sect = sect(references=header_reference("rIdH1", "default"))
    body = "<w:p><w:r><w:t>Body</w:t></w:r></w:p>"
    parts = {"word/header1.xml": header_content}
    rels = region_relationship("rIdH1", "header", "header1.xml")
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )

    write_package(
        source,
        body,
        sect_properties=s1_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="New Updated Header", style="Heading")],
        action="replace",
    )

    report = author_region(req)
    assert report.passed is True
    assert report.region is not None
    assert report.region["isolated"] is False
    assert report.region["part"] == "word/header1.xml"
    assert report.region["relationship_id"] == "rIdH1"

    with zipfile.ZipFile(dest) as archive:
        content = archive.read("word/header1.xml").decode("utf-8")
        assert "New Updated Header" in content
        assert "Old Header" not in content


def test_replace_footer_first_variant(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    footer_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Original First Footer</w:t></w:r></w:p></w:ftr>'
    )
    s1_sect = sect(
        references=footer_reference("rIdF1", "first"),
        title_page=True,
    )
    body = "<w:p><w:r><w:t>Body</w:t></w:r></w:p>"
    parts = {"word/footer1.xml": footer_content}
    rels = region_relationship("rIdF1", "footer", "footer1.xml")
    overrides = (
        '<Override PartName="/word/footer1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    )

    write_package(
        source,
        body,
        sect_properties=s1_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="footer",
        variant="first",
        paragraphs=[RegionParagraph(text="Updated First Footer", style=None)],
        action="replace",
    )

    report = author_region(req)
    assert report.passed is True
    assert report.region is not None
    assert report.region["kind"] == "footer"
    assert report.region["variant"] == "first"

    with zipfile.ZipFile(dest) as archive:
        content = archive.read("word/footer1.xml").decode("utf-8")
        assert "Updated First Footer" in content
        doc_root = ET.fromstring(archive.read("word/document.xml"))
        doc_ns = doc_root.tag[1:].split("}", 1)[0]
        sect_pr = doc_root.find(f".//{{{doc_ns}}}sectPr")
        assert sect_pr is not None
        assert sect_pr.find(f"{{{doc_ns}}}titlePg") is not None


def test_remove_restores_word_inheritance(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    h1_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:hdr>'
    )
    h2_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Header 2</w:t></w:r></w:p></w:hdr>'
    )

    s1_sect = sect(references=header_reference("rIdH1", "default"))
    s2_sect = sect(references=header_reference("rIdH2", "default"))
    s3_sect = sect()

    body = (
        f"<w:p><w:pPr>{s1_sect}</w:pPr><w:r><w:t>S1</w:t></w:r></w:p>"
        f"<w:p><w:pPr>{s2_sect}</w:pPr><w:r><w:t>S2</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>S3</w:t></w:r></w:p>"
    )
    parts = {"word/header1.xml": h1_content, "word/header2.xml": h2_content}
    rels = (
        f"{region_relationship('rIdH1', 'header', 'header1.xml')}"
        f"{region_relationship('rIdH2', 'header', 'header2.xml')}"
    )
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/header2.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )

    write_package(
        source,
        body,
        sect_properties=s3_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )

    disc_before = discover_region_targets(source)
    assert len(disc_before["sections"][1]["references"]) == 1
    assert disc_before["sections"][2]["inherited_references"][0]["inherited_from_section"] == 2

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=2,
        kind="header",
        variant="default",
        action="remove",
    )

    report = author_region(req)
    assert report.passed is True
    assert report.status == "published"
    assert report.retained_unreferenced_parts == ("word/header2.xml",)
    assert report.region is not None
    assert report.region["retained_unreferenced_parts"] == ["word/header2.xml"]

    with zipfile.ZipFile(dest) as archive:
        assert "word/header2.xml" in archive.namelist()
        assert archive.read("word/header2.xml") == h2_content.encode("utf-8")

    disc_after = discover_region_targets(dest)
    assert len(disc_after["sections"][1]["references"]) == 0
    assert disc_after["sections"][1]["inherited_references"][0]["inherited_from_section"] == 1
    assert disc_after["sections"][2]["inherited_references"][0]["inherited_from_section"] == 1


def test_remove_shared_header_retains_for_sharer(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest = tmp_path / "dest.docx"

    h1_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Shared Header</w:t></w:r></w:p></w:hdr>'
    )
    s1_sect = sect(references=header_reference("rIdH1", "default"))
    s2_sect = sect(references=header_reference("rIdH1", "default"))

    body = (
        f"<w:p><w:pPr>{s1_sect}</w:pPr><w:r><w:t>S1</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>S2</w:t></w:r></w:p>"
    )
    parts = {"word/header1.xml": h1_content}
    rels = region_relationship("rIdH1", "header", "header1.xml")
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )

    write_package(
        source,
        body,
        sect_properties=s2_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
    )

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=2,
        kind="header",
        variant="default",
        action="remove",
    )

    report = author_region(req)
    assert report.passed is True
    assert report.retained_unreferenced_parts == ()
    assert report.region is not None
    assert report.region["retained_unreferenced_parts"] == []

    disc_after = discover_region_targets(dest)
    assert len(disc_after["sections"][0]["references"]) == 1
    assert disc_after["sections"][0]["references"][0]["part"] == "word/header1.xml"
    assert len(disc_after["sections"][1]["references"]) == 0
    assert disc_after["sections"][1]["inherited_references"][0]["inherited_from_section"] == 1


def test_preserve_title_pg_and_even_odd_settings_on_lifecycle(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    dest1 = tmp_path / "dest1.docx"
    dest2 = tmp_path / "dest2.docx"

    h1_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>First Header</w:t></w:r></w:p></w:hdr>'
    )
    h_even_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Even Header</w:t></w:r></w:p></w:hdr>'
    )

    s1_sect = sect(
        references=header_reference("rIdH1", "first") + header_reference("rIdH2", "even"),
        title_page=True,
    )
    parts = {
        "word/header1.xml": h1_content,
        "word/header2.xml": h_even_content,
    }
    rels = (
        f"{region_relationship('rIdH1', 'header', 'header1.xml')}"
        f"{region_relationship('rIdH2', 'header', 'header2.xml')}"
    )
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/header2.xml" ContentType="'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )
    settings = settings_xml(evenAndOddHeaders=True)

    write_package(
        source,
        "<w:p><w:r><w:t>Text</w:t></w:r></w:p>",
        sect_properties=s1_sect,
        parts=parts,
        relationships=rels,
        overrides=overrides,
        settings=settings,
    )

    # Prove source settings are enabled before authoring
    disc_source = discover_region_targets(source)
    assert disc_source["even_and_odd_headers"] is True
    assert disc_source["sections"][0]["title_page"] is True
    req_replace = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest1,
        section=1,
        kind="header",
        variant="even",
        paragraphs=[RegionParagraph(text="New Even Header")],
        action="replace",
    )
    rep1 = author_region(req_replace)
    assert rep1.passed is True

    disc1 = discover_region_targets(dest1)
    assert disc1["even_and_odd_headers"] is True
    assert disc1["sections"][0]["title_page"] is True

    req_remove = RegionRequest(
        source=dest1,
        source_sha256=_sha256(dest1),
        destination=dest2,
        section=1,
        kind="header",
        variant="first",
        action="remove",
    )
    rep2 = author_region(req_remove)
    assert rep2.passed is True

    disc2 = discover_region_targets(dest2)
    assert disc2["even_and_odd_headers"] is True
    assert disc2["sections"][0]["title_page"] is True


def test_remove_rejects_nonempty_paragraphs(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    write_package(source, paragraph("Hello"))

    with pytest.raises(ValueError, match="region_remove_paragraphs_prohibited"):
        RegionRequest(
            source=source,
            source_sha256=_sha256(source),
            destination=tmp_path / "out.docx",
            section=1,
            paragraphs=[RegionParagraph(text="Forbidden")],
            action="remove",
        )

    with pytest.raises(ValueError, match="region_remove_paragraphs_prohibited"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": _sha256(source),
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "action": "remove",
                "paragraphs": [{"text": "Forbidden"}],
            }
        )


def test_refusal_missing_explicit_target(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    write_package(source, paragraph("Hello"))

    req_rep = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=tmp_path / "out1.docx",
        section=1,
        kind="header",
        variant="default",
        paragraphs=[RegionParagraph(text="New")],
        action="replace",
    )
    rep_rep = author_region(req_rep)
    assert rep_rep.passed is False
    assert rep_rep.code == "region_target_missing"
    assert not (tmp_path / "out1.docx").exists()

    req_rem = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=tmp_path / "out2.docx",
        section=1,
        kind="header",
        variant="default",
        action="remove",
    )
    rep_rem = author_region(req_rem)
    assert rep_rem.passed is False
    assert rep_rem.code == "region_target_missing"
    assert not (tmp_path / "out2.docx").exists()


def test_refusal_duplicate_and_foreign_namespace_references(tmp_path: Path) -> None:
    source_dup = tmp_path / "source_dup.docx"
    s_dup_sect = sect(
        references=header_reference("rIdH1", "default") + header_reference("rIdH1", "default")
    )
    h_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>'
    )
    write_package(
        source_dup,
        paragraph("Body"),
        sect_properties=s_dup_sect,
        parts={"word/header1.xml": h_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    req_dup = RegionRequest(
        source=source_dup,
        source_sha256=_sha256(source_dup),
        destination=tmp_path / "out_dup.docx",
        section=1,
        kind="header",
        variant="default",
        action="remove",
    )
    report_dup = author_region(req_dup)
    assert report_dup.passed is False
    assert report_dup.code == "region_reference_duplicate"

    source_foreign = tmp_path / "source_foreign.docx"
    foreign_ref = '<f:headerReference xmlns:f="urn:foreign" w:type="default" r:id="rIdH1"/>'
    s_foreign_sect = sect(references=foreign_ref)
    write_package(
        source_foreign,
        paragraph("Body"),
        sect_properties=s_foreign_sect,
        parts={"word/header1.xml": h_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    req_foreign = RegionRequest(
        source=source_foreign,
        source_sha256=_sha256(source_foreign),
        destination=tmp_path / "out_foreign.docx",
        section=1,
        kind="header",
        variant="default",
        action="remove",
    )
    report_foreign = author_region(req_foreign)
    assert report_foreign.passed is False
    assert report_foreign.code == "region_reference_foreign_namespace"


def test_refusal_unsafe_and_malformed_relationships(tmp_path: Path) -> None:
    h_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>'
    )
    source_ext = tmp_path / "source_ext.docx"
    s_sect = sect(references=header_reference("rIdExt", "default"))
    rels_ext = f'<Relationship Id="rIdExt" Type="{R}/header" Target="http://evil.com" TargetMode="External"/>'
    write_package(
        source_ext,
        paragraph("Body"),
        sect_properties=s_sect,
        relationships=rels_ext,
    )
    req_ext = RegionRequest(
        source=source_ext,
        source_sha256=_sha256(source_ext),
        destination=tmp_path / "out_ext.docx",
        section=1,
        action="remove",
    )
    rep_ext = author_region(req_ext)
    assert rep_ext.passed is False
    assert rep_ext.code == "region_relationship_external_unsupported"

    source_trav = tmp_path / "source_trav.docx"
    rels_trav = f'<Relationship Id="rIdTrav" Type="{R}/header" Target="../passwd.xml"/>'
    write_package(
        source_trav,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdTrav", "default")),
        relationships=rels_trav,
    )
    req_trav = RegionRequest(
        source=source_trav,
        source_sha256=_sha256(source_trav),
        destination=tmp_path / "out_trav.docx",
        section=1,
        action="remove",
    )
    rep_trav = author_region(req_trav)
    assert rep_trav.passed is False
    assert rep_trav.code == "region_relationship_unsafe"

    source_miss = tmp_path / "source_miss.docx"
    write_package(
        source_miss,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdGhost", "default")),
    )
    req_miss = RegionRequest(
        source=source_miss,
        source_sha256=_sha256(source_miss),
        destination=tmp_path / "out_miss.docx",
        section=1,
        action="remove",
    )
    rep_miss = author_region(req_miss)
    assert rep_miss.passed is False
    assert rep_miss.code == "region_relationship_missing"

    source_type = tmp_path / "source_type.docx"
    rels_wrong_type = f'<Relationship Id="rIdH1" Type="{R}/footer" Target="header1.xml"/>'
    write_package(
        source_type,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdH1", "default")),
        parts={"word/header1.xml": h_content},
        relationships=rels_wrong_type,
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    req_type = RegionRequest(
        source=source_type,
        source_sha256=_sha256(source_type),
        destination=tmp_path / "out_type.docx",
        section=1,
        kind="header",
        action="remove",
    )
    rep_type = author_region(req_type)
    assert rep_type.passed is False
    assert rep_type.code == "region_relationship_type_invalid"


def test_refusal_wrong_story_roots_and_types(tmp_path: Path) -> None:
    source_root = tmp_path / "source_root.docx"
    ftr_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Ftr</w:t></w:r></w:p></w:ftr>'
    )
    write_package(
        source_root,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdH1", "default")),
        parts={"word/header1.xml": ftr_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    req_root = RegionRequest(
        source=source_root,
        source_sha256=_sha256(source_root),
        destination=tmp_path / "out_root.docx",
        section=1,
        kind="header",
        action="remove",
    )
    rep_root = author_region(req_root)
    assert rep_root.passed is False
    assert rep_root.code == "region_story_root_invalid"

    source_ct = tmp_path / "source_ct.docx"
    hdr_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>'
    )
    write_package(
        source_ct,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdH1", "default")),
        parts={"word/header1.xml": hdr_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    )
    req_ct = RegionRequest(
        source=source_ct,
        source_sha256=_sha256(source_ct),
        destination=tmp_path / "out_ct.docx",
        section=1,
        kind="header",
        action="remove",
    )
    rep_ct = author_region(req_ct)
    assert rep_ct.passed is False
    assert rep_ct.code == "region_story_type_invalid"

    source_miss_part = tmp_path / "source_miss_part.docx"
    write_package(
        source_miss_part,
        paragraph("Body"),
        sect_properties=sect(references=header_reference("rIdH1", "default")),
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    req_miss_part = RegionRequest(
        source=source_miss_part,
        source_sha256=_sha256(source_miss_part),
        destination=tmp_path / "out_miss_part.docx",
        section=1,
        kind="header",
        action="remove",
    )
    rep_miss_part = author_region(req_miss_part)
    assert rep_miss_part.passed is False
    assert rep_miss_part.code == "region_story_part_missing"


def test_parse_and_load_request_actions(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    write_package(source, paragraph("Parser test"))
    digest = _sha256(source)

    req_default = parse_region_request(
        {
            "source": str(source),
            "source_sha256": digest,
            "destination": str(tmp_path / "out.docx"),
            "section": 1,
        }
    )
    assert req_default.action == "create"

    req_replace = parse_region_request(
        {
            "source": str(source),
            "source_sha256": digest,
            "destination": str(tmp_path / "out.docx"),
            "section": 1,
            "action": "replace",
            "paragraphs": [{"text": "Hello"}],
        }
    )
    assert req_replace.action == "replace"

    req_remove = parse_region_request(
        {
            "source": str(source),
            "source_sha256": digest,
            "destination": str(tmp_path / "out.docx"),
            "section": 1,
            "action": "remove",
        }
    )
    assert req_remove.action == "remove"

    with pytest.raises(ValueError, match="region_action_invalid"):
        parse_region_request(
            {
                "source": str(source),
                "source_sha256": digest,
                "destination": str(tmp_path / "out.docx"),
                "section": 1,
                "action": "delete",
            }
        )

    d = req_replace.to_dict()
    assert d["action"] == "replace"
    assert d["section"] == 1
    assert len(d["paragraphs"]) == 1


def test_refusal_foreign_document_root(tmp_path: Path) -> None:
    source = tmp_path / "source_foreign_doc.docx"
    dest = tmp_path / "out_foreign_doc.docx"

    h_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>'
    )
    s_sect = sect(references=header_reference("rIdH1", "default"))
    write_package(
        source,
        paragraph("Body"),
        sect_properties=s_sect,
        parts={"word/header1.xml": h_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )

    with zipfile.ZipFile(source, "r") as zin:
        parts_data = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    doc_elem = ET.fromstring(parts_data["word/document.xml"])
    doc_elem.tag = "{urn:foreign}document"
    parts_data["word/document.xml"] = ET.tostring(doc_elem, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, fbytes in parts_data.items():
            zout.writestr(fname, fbytes)

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        action="remove",
    )
    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "region_document_root_invalid"
    assert not dest.exists()


def test_refusal_duplicate_content_type_override(tmp_path: Path) -> None:
    source = tmp_path / "source_dup_ct.docx"
    dest = tmp_path / "out_dup_ct.docx"

    h_content = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>'
    )
    s_sect = sect(references=header_reference("rIdH1", "default"))
    overrides = (
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/xml"/>'
    )
    write_package(
        source,
        paragraph("Body"),
        sect_properties=s_sect,
        parts={"word/header1.xml": h_content},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        overrides=overrides,
    )

    req = RegionRequest(
        source=source,
        source_sha256=_sha256(source),
        destination=dest,
        section=1,
        kind="header",
        variant="default",
        action="remove",
    )
    report = author_region(req)
    assert report.passed is False
    assert report.status == "refused"
    assert report.code == "region_content_type_duplicate"
    assert not dest.exists()


def test_refusal_malformed_relationships_and_types_roots(tmp_path: Path) -> None:
    source_ct = tmp_path / "source_bad_ct.docx"
    write_package(source_ct, paragraph("Body"))
    with zipfile.ZipFile(source_ct, "r") as zin:
        parts_data = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    ct_elem = ET.fromstring(parts_data["[Content_Types].xml"])
    ct_elem.tag = "{urn:foreign}Types"
    parts_data["[Content_Types].xml"] = ET.tostring(ct_elem, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(source_ct, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, fbytes in parts_data.items():
            zout.writestr(fname, fbytes)

    req_ct = RegionRequest(
        source=source_ct,
        source_sha256=_sha256(source_ct),
        destination=tmp_path / "out1.docx",
        section=1,
        action="create",
    )
    rep_ct = author_region(req_ct)
    assert rep_ct.passed is False
    assert rep_ct.code == "source_package_invalid"

    source_rels = tmp_path / "source_bad_rels.docx"
    write_package(source_rels, paragraph("Body"))
    with zipfile.ZipFile(source_rels, "r") as zin:
        parts_data = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    rels_elem = ET.fromstring(parts_data["word/_rels/document.xml.rels"])
    rels_elem.tag = "{urn:foreign}Relationships"
    parts_data["word/_rels/document.xml.rels"] = ET.tostring(
        rels_elem, encoding="utf-8", xml_declaration=True
    )
    with zipfile.ZipFile(source_rels, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, fbytes in parts_data.items():
            zout.writestr(fname, fbytes)

    req_rels = RegionRequest(
        source=source_rels,
        source_sha256=_sha256(source_rels),
        destination=tmp_path / "out2.docx",
        section=1,
        action="create",
    )
    rep_rels = author_region(req_rels)
    assert rep_rels.passed is False
    assert rep_rels.code == "source_package_invalid"
