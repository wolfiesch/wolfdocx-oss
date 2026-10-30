from __future__ import annotations

import hashlib
import zipfile
from collections.abc import Callable
from functools import partial
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import write_styled_table_package
from wolfdocx import authoring
from wolfdocx.authoring import (
    CellRequestLimitError,
    author_cell_text,
    cell_structure_fingerprint,
    parse_cell_text_request,
)
from wolfdocx.inspect import Q, _clear_inspection_memo, inspect_docx


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _request(source: Path, destination: Path, *, text: str = "7,500") -> dict[str, object]:
    fingerprint = cell_structure_fingerprint(source, table=1, row=2, cell=2)
    assert fingerprint is not None
    return {
        "source": str(source),
        "destination": str(destination),
        "source_sha256": _sha256(source),
        "expected_old_text": "bar",
        "structure_fingerprint_sha256": fingerprint,
        "target": {"table": 1, "row": 2, "cell": 2},
        "text": text,
    }


def _cell_text(path: Path, *, table: int, row: int, cell: int) -> str:
    report = inspect_docx(path)
    return report.fingerprint.tables[table - 1]["rows"][row - 1]["cells"][cell - 1]["text"]


def _fixture(destination: Path) -> Path:
    return write_styled_table_package(destination)


def _document_xml(path: Path) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml")


def _replace_document_xml(path: Path, replace: Callable[[bytes], bytes]) -> None:
    with zipfile.ZipFile(path) as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as destination:
        for info, data in entries:
            if info.filename == "word/document.xml":
                data = replace(data)
            destination.writestr(info, data)


def _prepend_unrelated_wrapped_text_before_table(data: bytes) -> bytes:
    table = b"<w:tbl>"
    wrapped_text = (
        b"<w:p>"
        b"<w:hyperlink><w:r><w:t>unrelated hyperlink</w:t></w:r></w:hyperlink>"
        b'<w:fldSimple w:instr=" DATE "><w:r><w:t>unrelated field</w:t></w:r></w:fldSimple>'
        b"</w:p>"
    )
    assert table in data
    return data.replace(table, wrapped_text + table, 1)


def _add_target_grid_span(data: bytes) -> bytes:
    root = ET.fromstring(data)
    namespace = root.tag[1:].split("}", 1)[0]
    word = partial(Q, namespace)
    target = next(iter(root.iter(word("tbl")))).findall(word("tr"))[1].findall(word("tc"))[1]
    properties = target.find(word("tcPr"))
    if properties is None:
        properties = ET.Element(word("tcPr"))
        target.insert(0, properties)
    properties.append(ET.Element(word("gridSpan"), {word("val"): "2"}))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _wrap_target_table_in_revision(data: bytes) -> bytes:
    root = ET.fromstring(data)
    namespace = root.tag[1:].split("}", 1)[0]
    word = partial(Q, namespace)
    body = root.find(word("body"))
    assert body is not None
    table = next(child for child in body if child.tag == word("tbl"))
    position = list(body).index(table)
    body.remove(table)
    revision = ET.Element(word("ins"), {word("id"): "1"})
    revision.append(table)
    body.insert(position, revision)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def test_publishes_one_bound_simple_cell_and_preserves_all_other_part_bytes(tmp_path):
    source = _fixture(tmp_path / "source.docx")
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"existing destination")
    source_sha256 = _sha256(source)
    request = parse_cell_text_request(_request(source, destination))

    report = author_cell_text(request)

    assert report.passed is True
    assert report.code is None
    assert report.non_target_parts_preserved is True
    assert report.target_change_report is not None
    assert report.target_change_report["passed"] is True
    assert _sha256(source) == source_sha256
    assert _cell_text(destination, table=1, row=2, cell=2) == "7,500"
    source_document_xml = _document_xml(source)
    candidate_document_xml = _document_xml(destination)
    expected_document_xml = source_document_xml.replace(b"<w:t>bar</w:t>", b"<w:t>7,500</w:t>", 1)
    assert expected_document_xml != source_document_xml
    assert b"xmlns:w14=" in source_document_xml
    assert b"xmlns:wp14=" in source_document_xml
    assert b'mc:Ignorable="w14 wp14"' in source_document_xml
    assert candidate_document_xml == expected_document_xml
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(destination) as candidate:
        assert set(baseline.namelist()) == set(candidate.namelist())
        for name in baseline.namelist():
            if name != "word/document.xml":
                assert baseline.read(name) == candidate.read(name)


@pytest.mark.parametrize(
    ("ancestor_space", "node_space", "old_text", "text", "expected_code"),
    [
        (None, "preserve", " bar ", " 7,500 ", None),
        ("preserve", None, " bar ", " 7,500 ", None),
        ("default", "preserve", " bar ", " 7,500 ", None),
        (None, "preserve", "bar", " 7,500 ", None),
        (None, "preserve", " bar ", "7,500", None),
        (None, None, " bar ", " 7,500 ", "cell_structure_unsupported"),
        ("preserve", "default", " bar ", " 7,500 ", "cell_structure_unsupported"),
        (None, None, "bar", " 7,500 ", "replacement_whitespace_unsupported"),
        ("preserve", "default", "bar", " 7,500 ", "replacement_whitespace_unsupported"),
    ],
)
def test_cell_padding_requires_effective_preserve_without_changing_markup(
    tmp_path, ancestor_space, node_space, old_text, text, expected_code
):
    source = _fixture(tmp_path / "source.docx")
    space_attribute = f' xml:space="{node_space}"' if node_space else ""
    old_node = f"<w:t{space_attribute}>{old_text}</w:t>".encode()

    def transform(data):
        data = data.replace(b"<w:t>bar</w:t>", old_node, 1)
        if ancestor_space:
            data = data.replace(
                b"<w:document ",
                f'<w:document xml:space="{ancestor_space}" '.encode(),
                1,
            )
        return data

    _replace_document_xml(source, transform)
    source_before = source.read_bytes()
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    request_data = _request(source, destination, text=text)
    request_data["expected_old_text"] = old_text

    report = author_cell_text(parse_cell_text_request(request_data))

    assert source.read_bytes() == source_before
    assert report.code == expected_code
    if expected_code:
        assert report.passed is False
        assert destination.read_bytes() == b"do not replace"
    else:
        assert report.passed is True
        assert _cell_text(destination, table=1, row=2, cell=2) == text
        assert _document_xml(destination) == _document_xml(source).replace(
            old_node, f"<w:t{space_attribute}>{text}</w:t>".encode(), 1
        )
        with (
            zipfile.ZipFile(source) as baseline,
            zipfile.ZipFile(destination) as candidate,
        ):
            assert baseline.namelist() == candidate.namelist()
            for name in baseline.namelist():
                if name != "word/document.xml":
                    assert baseline.read(name) == candidate.read(name)


def test_publishes_simple_cell_after_unrelated_wrapped_text_and_preserves_it(tmp_path):
    source = _fixture(tmp_path / "wrapped-before-target.docx")
    _replace_document_xml(source, _prepend_unrelated_wrapped_text_before_table)
    destination = tmp_path / "output.docx"
    request = parse_cell_text_request(_request(source, destination))

    report = author_cell_text(request)

    assert report.passed is True
    assert report.code is None
    assert report.non_target_parts_preserved is True
    assert report.target_change_report is not None
    assert report.target_change_report["passed"] is True
    assert _cell_text(destination, table=1, row=2, cell=2) == "7,500"
    source_document_xml = _document_xml(source)
    candidate_document_xml = _document_xml(destination)
    expected_document_xml = source_document_xml.replace(b"<w:t>bar</w:t>", b"<w:t>7,500</w:t>", 1)
    assert b"<w:hyperlink><w:r><w:t>unrelated hyperlink</w:t>" in source_document_xml
    assert b'<w:fldSimple w:instr=" DATE "><w:r><w:t>unrelated field</w:t>' in source_document_xml
    assert source_document_xml.index(b"unrelated hyperlink") < source_document_xml.index(
        b"<w:t>bar</w:t>"
    )
    assert candidate_document_xml == expected_document_xml
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(destination) as candidate:
        for name in baseline.namelist():
            if name != "word/document.xml":
                assert baseline.read(name) == candidate.read(name)


@pytest.mark.parametrize(
    ("field", "value", "expected_code"),
    [
        ("source_sha256", "0" * 64, "source_hash_mismatch"),
        ("expected_old_text", "not the source text", "expected_old_text_mismatch"),
    ],
)
def test_mismatches_refuse_without_changing_source_or_existing_destination(
    tmp_path, field, value, expected_code
):
    source = _fixture(tmp_path / "source.docx")
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    source_before = source.read_bytes()
    destination_before = destination.read_bytes()
    request_data = _request(source, destination)
    request_data[field] = value

    report = author_cell_text(parse_cell_text_request(request_data))

    assert report.passed is False
    assert report.code == expected_code
    assert source.read_bytes() == source_before
    assert destination.read_bytes() == destination_before


@pytest.mark.parametrize(
    ("transform", "expected_code"),
    [
        (
            lambda data: data.replace(
                b"<w:t>bar</w:t>",
                b'<w:fldSimple w:instr=" DATE "><w:r><w:t>bar</w:t></w:r></w:fldSimple>',
                1,
            ),
            "cell_structure_unsupported",
        ),
        (_add_target_grid_span, "cell_structure_unsupported"),
        (_wrap_target_table_in_revision, "cell_context_unsupported"),
    ],
)
def test_complex_selected_cells_refuse_before_replacing_existing_destination(
    tmp_path, transform, expected_code
):
    source = _fixture(tmp_path / "complex.docx")
    _replace_document_xml(source, transform)
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    request = parse_cell_text_request(_request(source, destination))

    report = author_cell_text(request)

    assert report.passed is False
    assert report.code == expected_code
    assert destination.read_bytes() == b"do not replace"


def test_cell_request_rejects_unknown_fields():
    with pytest.raises(ValueError, match="cell_request_fields_invalid"):
        parse_cell_text_request({"source": "source.docx"})


def test_text_character_limit_identifies_requested_length_and_recovery(tmp_path):
    source = _fixture(tmp_path / "source.docx")
    request_data = _request(source, tmp_path / "output.docx")
    request_data["text"] = "x" * 32_769

    with pytest.raises(CellRequestLimitError) as error:
        parse_cell_text_request(request_data)

    assert error.value.to_dict()["limit"] == {
        "field": "text",
        "maximum_characters": 32_768,
        "requested_characters": 32_769,
        "recovery": "Shorten the value or choose a smaller span.",
    }


def _rewrite_part(path: Path, name: str, replace: Callable[[bytes], bytes]) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as rewritten:
        for info, data in entries:
            rewritten.writestr(info, replace(data) if info.filename == name else data)


def _add_style(data: bytes) -> bytes:
    assert data.count(b"</w:styles>") == 1
    return data.replace(
        b"</w:styles>", b'<w:style w:type="paragraph" w:styleId="Tampered"/></w:styles>'
    )


def _add_relationship(data: bytes) -> bytes:
    assert data.count(b"</Relationships>") == 1
    return data.replace(
        b"</Relationships>",
        b'<Relationship Id="rIdTampered" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" '
        b'TargetMode="External"/></Relationships>',
    )


@pytest.mark.parametrize("warm", [False, True])
def test_changed_parts_verification_publishes_the_same_package_as_full(tmp_path, warm):
    source = _fixture(tmp_path / "source.docx")
    source_sha256 = _sha256(source)
    outputs = {}
    for verify in ("full", "changed-parts"):
        if not warm:
            _clear_inspection_memo()
        destination = tmp_path / f"{verify}.docx"
        report = author_cell_text(
            parse_cell_text_request(_request(source, destination)), verify=verify
        )
        assert report.passed is True, report.code
        assert report.non_target_parts_preserved is True
        assert report.target_change_report is not None
        assert report.target_change_report["passed"] is True
        assert report.destination_sha256 == _sha256(destination)
        outputs[verify] = destination.read_bytes()
    assert outputs["changed-parts"] == outputs["full"]
    assert _sha256(source) == source_sha256
    destination = tmp_path / "changed-parts.docx"
    assert _cell_text(destination, table=1, row=2, cell=2) == "7,500"
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(destination) as candidate:
        assert baseline.namelist() == candidate.namelist()
        changed = [
            name for name in baseline.namelist() if baseline.read(name) != candidate.read(name)
        ]
    assert changed == ["word/document.xml"]


@pytest.mark.parametrize("verify", ["full", "changed-parts"])
@pytest.mark.parametrize(
    ("part", "tamper"),
    [
        ("word/styles.xml", _add_style),
        ("word/_rels/document.xml.rels", _add_relationship),
        ("_rels/.rels", _add_relationship),
    ],
)
def test_staged_change_outside_the_target_part_refuses_in_every_verify_mode(
    tmp_path, monkeypatch, verify, part, tamper
):
    source = _fixture(tmp_path / "source.docx")
    source_sha256 = _sha256(source)
    destination = tmp_path / "output.docx"
    request = parse_cell_text_request(_request(source, destination))
    write_staged_package = authoring._write_staged_package

    def write_tampered(source: Path, staged: Path, document_xml: bytes) -> None:
        write_staged_package(source, staged, document_xml)
        _rewrite_part(staged, part, tamper)

    monkeypatch.setattr(authoring, "_write_staged_package", write_tampered)

    report = author_cell_text(request, verify=verify)

    assert report.passed is False
    assert report.code == "staged_verification_failed"
    assert not destination.exists()
    assert _sha256(source) == source_sha256
    assert not list(tmp_path.glob(".output.docx.*"))


def test_unknown_verify_mode_is_rejected(tmp_path):
    source = _fixture(tmp_path / "source.docx")
    request = parse_cell_text_request(_request(source, tmp_path / "output.docx"))
    with pytest.raises(ValueError, match="verify_mode_invalid"):
        author_cell_text(request, verify="none")
