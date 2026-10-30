from __future__ import annotations

import hashlib
import json
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

from wolfdocx import PackageLimits
from wolfdocx.inspect import inspect_docx
from wolfdocx.inspect import (
    _parse_bounded_xml,
    _XmlSafetyError,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)

from .docx_fixtures import R, W, cell, declare_untyped_parts, paragraph, row, table, write_package


def write_fixture(path, *, dangling_relationship=False):
    """Generate a three-paragraph package with a two-cell table."""
    body = paragraph("Baseline paragraph") + table(
        row(cell(paragraph("First cell")) + cell(paragraph("Second cell"))),
        grid='<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>',
    )
    relationships = (
        f'<Relationship Id="rIdMissing" Type="{R}/header" Target="missing.xml"/>'
        if dangling_relationship else ""
    )
    return write_package(path, body, relationships=relationships)


def add_package_part(path, name, value):
    with zipfile.ZipFile(path) as archive:
        parts = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    parts[name] = value
    types = declare_untyped_parts(parts["[Content_Types].xml"].decode(), parts)
    parts["[Content_Types].xml"] = types.encode()
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for part, data in parts.items():
            archive.writestr(part, data)


def _replace_zip_part(document, part_name: str, value: bytes) -> None:
    replacement = document.with_suffix(".replacement.docx")
    with (
        zipfile.ZipFile(document) as source,
        zipfile.ZipFile(replacement, "w", compression=zipfile.ZIP_DEFLATED) as target,
    ):
        for entry in source.infolist():
            target.writestr(
                entry,
                value if entry.filename == part_name else source.read(entry.filename),
            )
    replacement.replace(document)


def _add_zip_part(document, part_name: str, value: bytes) -> None:
    add_package_part(document, part_name, value)


def test_serialization_preserves_package_default_namespaces_and_fingerprints():
    for namespace, name, child in (
        (
            "http://schemas.openxmlformats.org/package/2006/content-types",
            "Types",
            '<Default Extension="xml" ContentType="application/xml"/>',
        ),
        (
            "http://schemas.openxmlformats.org/package/2006/relationships",
            "Relationships",
            '<Relationship Id="rId1" Type="urn:example" Target="document.xml"/>',
        ),
    ):
        original = f'<{name} xmlns="{namespace}">{child}</{name}>'.encode()
        root = parse_xml_with_preserved_qnames(original)
        fingerprint_input = ET.tostring(root)
        output = serialize_xml_with_preserved_qnames(root)
        declarations = dict(
            value for _, value in ET.iterparse(BytesIO(output), events=("start-ns",))
        )
        assert declarations[""] == namespace
        assert ET.tostring(ET.fromstring(output)) == fingerprint_input
        assert ET.tostring(root) == fingerprint_input


def test_serialization_retains_ignorable_namespace_aliases_without_global_mutation():
    original = b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body/></w:document>'
    root = parse_xml_with_preserved_qnames(original)
    fingerprint_input = ET.tostring(root)
    output = serialize_xml_with_preserved_qnames(root)
    declarations = dict(value for _, value in ET.iterparse(BytesIO(output), events=("start-ns",)))
    assert declarations["w14"] == "http://schemas.microsoft.com/office/word/2010/wordml"
    assert ET.tostring(root) == fingerprint_input


def test_new_opc_nodes_preserve_default_namespace_and_qualified_attributes():
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    root = ET.Element(f"{{{namespace}}}Types")
    ET.SubElement(
        root,
        f"{{{namespace}}}Default",
        {"Extension": "xml", f"{{{namespace}}}extension": "qualified"},
    )
    ET.SubElement(root, "unqualified")
    output = serialize_xml_with_preserved_qnames(root)
    declarations = dict(
        value for _, value in ET.iterparse(BytesIO(output), events=("start-ns",)) if value[1]
    )
    assert declarations[""] == namespace
    parsed = ET.fromstring(output)
    assert parsed[0].tag == f"{{{namespace}}}Default"
    assert parsed[0].attrib == {
        "Extension": "xml",
        f"{{{namespace}}}extension": "qualified",
    }
    assert parsed[1].tag == "unqualified"


@pytest.mark.parametrize(
    ("limits", "code"),
    [
        (PackageLimits(max_xml_elements=1), "xml_element_limit"),
        (PackageLimits(max_total_xml_elements=1), "xml_total_element_limit"),
        (PackageLimits(max_xml_depth=1), "xml_depth_limit"),
    ],
)
def test_bounded_parser_refuses_before_consuming_malformed_suffix(limits, code):
    with pytest.raises(_XmlSafetyError) as error:
        _parse_bounded_xml(
            b"<root><child/>" + b" " * (64 * 1024) + b"<malformed",
            limits=limits,
            total_elements=[0],
        )
    assert error.value.code == code


def test_bounded_parser_carries_inherited_bindings_into_detached_subtrees():
    root = _parse_bounded_xml(
        b'<root xmlns:k="urn:retained"><child opaque="k:Type"/></root>',
        limits=PackageLimits(max_xml_elements=2, max_total_xml_elements=2, max_xml_depth=2),
        total_elements=[0],
    )
    child = root[0]
    root.remove(child)
    output = serialize_xml_with_preserved_qnames(child)
    declarations = dict(value for _, value in ET.iterparse(BytesIO(output), events=("start-ns",)))
    assert declarations["k"] == "urn:retained"
    assert ET.fromstring(output).get("opaque") == "k:Type"


def test_inspects_minimal_docx_package(tmp_path):
    document = write_fixture(tmp_path / "baseline.docx")

    report = inspect_docx(document)

    assert report.valid
    assert report.issues == []
    assert report.fingerprint.paragraph_count == 3
    assert report.fingerprint.table_count == 1
    assert report.fingerprint.row_count == 1
    assert report.fingerprint.cell_count == 2
    assert report.fingerprint.feature_counts["sections"] == 1
    assert report.active_content["present"] is False
    assert {part.name for part in report.parts} >= {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
        "word/styles.xml",
    }


def test_inspects_iso_strict_wordprocessingml_semantics(tmp_path):
    document = tmp_path / "strict.docx"
    strict_word = "http://purl.oclc.org/ooxml/wordprocessingml/main"
    drawings = "<w:r><w:drawing/></w:r>" * 3
    body = (
        '<w:p><w:bookmarkStart w:id="1" w:name="Anchor"/>'
        '<w:r><w:t>Strict paragraph</w:t></w:r>'
        + drawings + '<w:bookmarkEnd w:id="1"/></w:p>'
        + paragraph("Second strict paragraph") + paragraph("Third strict paragraph")
    )
    styles = f'<w:styles xmlns:w="{W}">' + "".join(
        f'<w:style w:type="paragraph" w:styleId="Style{index}">'
        f'<w:name w:val="Style {index}"/></w:style>' for index in range(6)
    ) + "</w:styles>"
    write_package(document, body, parts={"word/styles.xml": styles})
    with zipfile.ZipFile(document) as archive:
        parts = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    with zipfile.ZipFile(document, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload.replace(W.encode(), strict_word.encode()).replace(
                R.encode(), b"http://purl.oclc.org/ooxml/officeDocument/relationships"
            ))

    report = inspect_docx(document)

    assert report.valid
    assert report.fingerprint.paragraph_count == 3
    assert len(report.fingerprint.styles) == 6
    assert report.fingerprint.feature_counts["bookmarks"] == 1
    assert report.fingerprint.feature_counts["drawings"] == 3
    assert report.fingerprint.feature_counts["sections"] == 1


def test_reports_dangling_internal_relationship(tmp_path):
    document = write_fixture(
        tmp_path / "dangling.docx",
        dangling_relationship=True,
    )

    report = inspect_docx(document)

    assert not report.valid
    assert any(issue.code == "dangling_relationship" for issue in report.issues)


def test_reports_non_zip_input(tmp_path):
    document = tmp_path / "not-a-docx.docx"
    document.write_text("not a package")

    report = inspect_docx(document)

    assert not report.valid
    assert [issue.code for issue in report.issues] == ["not_zip"]


def test_stops_before_reading_package_over_uncompressed_limit(tmp_path):
    document = write_fixture(tmp_path / "large.docx")

    report = inspect_docx(
        document,
        limits=PackageLimits(max_uncompressed_size=1),
    )

    assert not report.valid
    assert [issue.code for issue in report.issues] == ["package_size_limit"]
    assert report.parts == []


def test_rejects_dtd_before_parsing_ooxml(tmp_path):
    document = write_fixture(tmp_path / "dtd.docx")
    with zipfile.ZipFile(document) as archive:
        original = archive.read("word/document.xml")
    declaration_end = original.find(b"?>") + 2
    _replace_zip_part(
        document,
        "word/document.xml",
        original[:declaration_end]
        + b'<!DOCTYPE document [<!ENTITY x "unsafe">]>'
        + original[declaration_end:],
    )

    report = inspect_docx(document)

    assert not report.valid
    issue = next(issue for issue in report.issues if issue.code == "xml_dtd_forbidden")
    assert issue.part == "word/document.xml"
    assert report.fingerprint.paragraph_count == 0


def test_dtd_text_inside_xml_comment_is_not_misclassified(tmp_path):
    document = write_fixture(tmp_path / "dtd-comment.docx")
    with zipfile.ZipFile(document) as archive:
        original = archive.read("word/document.xml")
    declaration_end = original.find(b"?>") + 2
    _replace_zip_part(
        document,
        "word/document.xml",
        original[:declaration_end]
        + b"<!-- documentation mentioning <!DOCTYPE but declaring nothing -->"
        + original[declaration_end:],
    )

    report = inspect_docx(document)

    assert report.valid
    assert report.fingerprint.paragraph_count == 3


def test_rejects_utf16_encoded_dtd_before_parsing_ooxml(tmp_path):
    document = write_fixture(tmp_path / "utf16-dtd.docx")
    with zipfile.ZipFile(document) as archive:
        original = archive.read("word/document.xml").decode("utf-8")
    declaration_end = original.find("?>") + 2
    unsafe = (
        original[:declaration_end]
        + '<!DOCTYPE document [<!ENTITY x "unsafe">]>'
        + original[declaration_end:]
    ).encode("utf-16")
    _replace_zip_part(document, "word/document.xml", unsafe)

    report = inspect_docx(document)

    assert not report.valid
    assert any(issue.code == "xml_dtd_forbidden" for issue in report.issues)


def test_rejects_xml_part_over_size_limit_before_parsing(tmp_path):
    document = write_fixture(tmp_path / "xml-size.docx")

    report = inspect_docx(
        document,
        limits=PackageLimits(max_xml_part_size=1),
    )

    assert not report.valid
    assert any(issue.code == "xml_size_limit" for issue in report.issues)
    assert report.fingerprint.paragraph_count == 0


def test_rejects_xml_part_over_element_limit_before_fingerprinting(tmp_path):
    document = write_fixture(tmp_path / "xml-elements.docx")

    report = inspect_docx(
        document,
        limits=PackageLimits(max_xml_elements=1),
    )

    assert not report.valid
    assert any(issue.code == "xml_element_limit" for issue in report.issues)
    assert report.fingerprint.paragraph_count == 0


def test_rejects_package_over_aggregate_xml_element_limit(tmp_path):
    document = write_fixture(tmp_path / "xml-total-elements.docx")

    report = inspect_docx(
        document,
        limits=PackageLimits(
            max_xml_elements=1_000_000,
            max_total_xml_elements=1,
        ),
    )

    assert not report.valid
    assert any(issue.code == "xml_total_element_limit" for issue in report.issues)
    assert report.fingerprint.paragraph_count == 0


def test_rejects_xml_over_nesting_depth_limit(tmp_path):
    document = write_fixture(tmp_path / "xml-depth.docx")

    report = inspect_docx(
        document,
        limits=PackageLimits(max_xml_depth=1),
    )

    assert not report.valid
    assert any(issue.code == "xml_depth_limit" for issue in report.issues)
    assert report.fingerprint.paragraph_count == 0


def test_inventories_active_content_without_echoing_targets_or_field_text(tmp_path):
    document = write_fixture(tmp_path / "active-content.docx")
    external_target = "https://private.example.invalid/confidential"
    field_target = "https://private.example.invalid/source.docx"
    with zipfile.ZipFile(document) as archive:
        relationships = archive.read("word/_rels/document.xml.rels")
        document_xml = archive.read("word/document.xml")
    _replace_zip_part(
        document,
        "word/_rels/document.xml.rels",
        relationships.replace(
            b"</Relationships>",
            (
                '<Relationship Id="rIdExternal" '
                'Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/hyperlink" '
                f'Target="{external_target}" TargetMode="External"/>'
                "</Relationships>"
            ).encode(),
        ),
    )
    _replace_zip_part(
        document,
        "word/document.xml",
        document_xml.replace(
            b"<w:sectPr>",
            (
                f'<w:p><w:fldSimple w:instr="INCLUDETEXT {field_target}">'
                "<w:r><w:t>external</w:t></w:r></w:fldSimple></w:p>"
                '<w:altChunk r:id="rIdExternal"/>'
                "<w:sectPr>"
            ).encode(),
        ),
    )
    _add_zip_part(document, "word/vbaProject.bin", b"synthetic macro marker")
    _add_zip_part(document, "word/embeddings/object1.bin", b"synthetic object")

    report = inspect_docx(document)
    inventory = report.active_content
    encoded = json.dumps(report.to_dict(full=True), sort_keys=True)

    assert report.valid
    assert inventory["present"] is True
    assert inventory["macro_parts"] == [
        {
            "part": "word/vbaProject.bin",
            "kind": "vba",
            "content_sha256": hashlib.sha256(b"synthetic macro marker").hexdigest(),
        }
    ]
    assert inventory["embedded_parts"] == [
        {
            "part": "word/embeddings/object1.bin",
            "content_sha256": hashlib.sha256(b"synthetic object").hexdigest(),
        }
    ]
    assert inventory["altchunk_parts"] == {"word/document.xml": 1}
    assert inventory["external_field_counts"] == {"INCLUDETEXT": 1}
    assert len(inventory["external_relationships"]) == 1
    relationship = inventory["external_relationships"][0]
    assert relationship["kind"] == "hyperlink"
    assert relationship["source_part"] == "word/document.xml"
    assert relationship["target_scheme"] == "https"
    assert len(relationship["target_sha256"]) == 64
    assert external_target not in encoded
    assert field_target not in encoded


def _with_footnote_reference(document, *, include_body: bool) -> None:
    with zipfile.ZipFile(document) as archive:
        parts = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    _replace_zip_part(
        document,
        "[Content_Types].xml",
        parts["[Content_Types].xml"].replace(
            b"</Types>",
            b'<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
            b"</Types>",
        ),
    )
    _replace_zip_part(
        document,
        "word/_rels/document.xml.rels",
        parts["word/_rels/document.xml.rels"].replace(
            b"</Relationships>",
            b'<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
            b"</Relationships>",
        ),
    )
    _replace_zip_part(
        document,
        "word/document.xml",
        parts["word/document.xml"].replace(
            b"<w:sectPr",
            b'<w:p><w:r><w:t>Clause with note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr',
            1,
        ),
    )
    if include_body:
        _add_zip_part(
            document,
            "word/footnotes.xml",
            b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>Source note body</w:t></w:r></w:p></w:footnote></w:footnotes>',
        )


def test_inspection_requires_footnote_part_and_referenced_body(tmp_path):
    valid = write_fixture(tmp_path / "footnote-valid.docx")
    _with_footnote_reference(valid, include_body=True)
    assert inspect_docx(valid).valid

    dangling = write_fixture(tmp_path / "footnote-dangling.docx")
    _with_footnote_reference(dangling, include_body=False)
    report = inspect_docx(dangling)
    assert not report.valid
    assert any(issue.code == "dangling_footnote_reference" for issue in report.issues)
