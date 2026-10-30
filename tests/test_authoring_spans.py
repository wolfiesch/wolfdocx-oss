from __future__ import annotations

import hashlib
import zipfile
from collections.abc import Callable
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import write_styled_table_package
from wolfdocx.authoring import (
    SpanRequestLimitError,
    author_cell_text,
    author_span_text,
    cell_structure_fingerprint,
    paragraph_structure_fingerprint,
    parse_cell_text_request,
    parse_span_text_request,
)
from wolfdocx.inspect import Q

_SOURCE_RUN = b"<w:r><w:t>foobar</w:t></w:r>"
_SPLIT_RUN_CONTENT = (
    b'<w:r><w:rPr><w:color w:val="4472C4"/></w:rPr><w:t xml:space="preserve">head </w:t></w:r>'
    b"<w:r><w:rPr><w:b/></w:rPr><w:t>pre alpha</w:t></w:r>"
    b'<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">beta post</w:t></w:r>'
    b'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> tail</w:t></w:r>'
)
_SPLIT_RUNS = b'<w:pPr><w:jc w:val="right"/></w:pPr>' + _SPLIT_RUN_CONTENT


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


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


def _split_run_fixture(destination: Path) -> Path:
    write_styled_table_package(destination)
    _replace_document_xml(destination, lambda data: data.replace(_SOURCE_RUN, _SPLIT_RUNS, 1))
    return destination


def _request(
    source: Path, destination: Path, *, text: str = "replacement", paragraph: int = 1
) -> dict[str, object]:
    fingerprint = paragraph_structure_fingerprint(source, paragraph=paragraph)
    assert fingerprint is not None
    return {
        "source": str(source),
        "destination": str(destination),
        "source_sha256": _sha256(source),
        "expected_old_text": "alphabeta",
        "structure_fingerprint_sha256": fingerprint,
        "paragraph": paragraph,
        "text": text,
    }


def _paragraph_text(path: Path, *, paragraph: int) -> str:
    root = ET.fromstring(_document_xml(path))
    namespace = root.tag[1:].split("}", 1)[0]
    target = list(root.iter(Q(namespace, "p")))[paragraph - 1]
    return "".join(text.text or "" for text in target.iter(Q(namespace, "t")))


def test_publishes_one_bound_split_run_span_and_preserves_every_other_byte(tmp_path):
    source = _split_run_fixture(tmp_path / "source.docx")
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"existing destination")
    source_sha256 = _sha256(source)
    request = parse_span_text_request(_request(source, destination))

    report = author_span_text(request)

    assert report.passed is True
    assert report.code is None
    assert report.target_change_report is not None
    assert report.target_change_report["passed"] is True
    assert report.non_target_document_bytes_preserved is True
    assert report.non_target_parts_preserved is True
    assert _sha256(source) == source_sha256
    assert _paragraph_text(destination, paragraph=1) == "head pre replacement post tail"
    source_document_xml = _document_xml(source)
    expected_document_xml = source_document_xml.replace(
        b"<w:t>pre alpha</w:t>", b"<w:t>pre replacement</w:t>", 1
    ).replace(
        b'<w:t xml:space="preserve">beta post</w:t>',
        b'<w:t xml:space="preserve"> post</w:t>',
        1,
    )
    assert _document_xml(destination) == expected_document_xml
    assert b"<w:rPr><w:b/></w:rPr><w:t>pre replacement</w:t>" in expected_document_xml
    assert b'<w:pPr><w:jc w:val="right"/></w:pPr>' in expected_document_xml
    assert (
        b'<w:rPr><w:color w:val="4472C4"/></w:rPr><w:t xml:space="preserve">head </w:t>'
        in expected_document_xml
    )
    assert (
        b'<w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> tail</w:t>'
        in expected_document_xml
    )
    assert (
        paragraph_structure_fingerprint(destination, paragraph=1)
        == request.structure_fingerprint_sha256
    )
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(destination) as candidate:
        assert set(baseline.namelist()) == set(candidate.namelist())
        for name in baseline.namelist():
            if name != "word/document.xml":
                assert baseline.read(name) == candidate.read(name)


def _make_ambiguous(data: bytes) -> bytes:
    return data.replace(
        _SPLIT_RUNS,
        (
            b'<w:pPr><w:jc w:val="right"/></w:pPr>'
            b'<w:r><w:rPr><w:color w:val="4472C4"/></w:rPr><w:t xml:space="preserve">head </w:t></w:r>'
            b"<w:r><w:rPr><w:b/></w:rPr><w:t>pre alpha</w:t></w:r>"
            b"<w:r><w:rPr><w:i/></w:rPr><w:t>beta alpha</w:t></w:r>"
            b'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">beta post</w:t></w:r>'
            b'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> tail</w:t></w:r>'
        ),
        1,
    )


def _insert_field(data: bytes) -> bytes:
    return data.replace(
        _SPLIT_RUNS,
        (
            b'<w:pPr><w:jc w:val="right"/></w:pPr>'
            b'<w:fldSimple w:instr=" DATE ">' + _SPLIT_RUN_CONTENT + b"</w:fldSimple>"
        ),
        1,
    )


def _insert_revision(data: bytes) -> bytes:
    return data.replace(
        _SPLIT_RUNS,
        b'<w:pPr><w:jc w:val="right"/></w:pPr><w:ins w:id="1">' + _SPLIT_RUN_CONTENT + b"</w:ins>",
        1,
    )


def _insert_comment(data: bytes) -> bytes:
    return data.replace(
        _SPLIT_RUNS,
        b'<w:pPr><w:jc w:val="right"/></w:pPr>'
        + _SPLIT_RUN_CONTENT.replace(
            b"</w:r><w:r><w:rPr><w:i/>",
            b'</w:r><w:commentRangeStart w:id="1"/><w:r><w:rPr><w:i/>',
            1,
        ),
        1,
    )


def _insert_control(data: bytes) -> bytes:
    return data.replace(
        _SPLIT_RUNS,
        b'<w:pPr><w:jc w:val="right"/></w:pPr><w:sdt><w:sdtContent>'
        + _SPLIT_RUN_CONTENT
        + b"</w:sdtContent></w:sdt>",
        1,
    )


def test_refuses_suffix_whitespace_without_existing_space_preservation(tmp_path):
    source = _split_run_fixture(tmp_path / "source.docx")
    _replace_document_xml(
        source,
        lambda data: data.replace(
            b'<w:t xml:space="preserve">beta post</w:t>', b"<w:t>beta post</w:t>", 1
        ),
    )
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    source_before = source.read_bytes()
    destination_before = destination.read_bytes()

    report = author_span_text(parse_span_text_request(_request(source, destination)))

    assert report.passed is False
    assert report.code == "span_whitespace_unsupported"
    assert source.read_bytes() == source_before
    assert destination.read_bytes() == destination_before


def _surround_document_with_range(data: bytes, start: bytes, end: bytes) -> bytes:
    return data.replace(b"<w:body>", b"<w:body><w:p>" + start + b"</w:p>", 1).replace(
        b"<w:sectPr", b"<w:p>" + end + b"</w:p><w:sectPr", 1
    )


def _insert_closed_range_before_document(data: bytes, start: bytes, end: bytes) -> bytes:
    return data.replace(b"<w:body>", b"<w:body><w:p>" + start + end + b"</w:p>", 1)


@pytest.mark.parametrize(
    ("start", "end"),
    [
        (b'<w:commentRangeStart w:id="17"/>', b'<w:commentRangeEnd w:id="17"/>'),
        (
            b'<w:fldChar w:fldCharType="begin"/>',
            b'<w:fldChar w:fldCharType="end"/>',
        ),
        (b'<w:moveFromRangeStart w:id="17"/>', b'<w:moveFromRangeEnd w:id="17"/>'),
        (b'<w:moveToRangeStart w:id="17"/>', b'<w:moveToRangeEnd w:id="17"/>'),
    ],
)
def test_active_document_range_context_refuses_span_before_writing(tmp_path, start, end):
    source = _split_run_fixture(tmp_path / "source.docx")
    _replace_document_xml(source, lambda data: _surround_document_with_range(data, start, end))
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    source_before = source.read_bytes()
    destination_before = destination.read_bytes()

    report = author_span_text(parse_span_text_request(_request(source, destination, paragraph=2)))

    assert report.passed is False
    assert report.code == "paragraph_context_unsupported"
    assert source.read_bytes() == source_before
    assert destination.read_bytes() == destination_before


def test_closed_unrelated_document_range_allows_span(tmp_path):
    source = _split_run_fixture(tmp_path / "source.docx")
    _replace_document_xml(
        source,
        lambda data: _insert_closed_range_before_document(
            data,
            b'<w:commentRangeStart w:id="17"/>',
            b'<w:commentRangeEnd w:id="17"/>',
        ),
    )
    destination = tmp_path / "existing-output.docx"

    report = author_span_text(parse_span_text_request(_request(source, destination, paragraph=2)))

    assert report.passed is True
    assert report.code is None
    assert _paragraph_text(destination, paragraph=2) == "head pre replacement post tail"


def _cell_request(source: Path, destination: Path) -> dict[str, object]:
    fingerprint = cell_structure_fingerprint(source, table=1, row=2, cell=2)
    assert fingerprint is not None
    return {
        "source": str(source),
        "destination": str(destination),
        "source_sha256": _sha256(source),
        "expected_old_text": "bar",
        "structure_fingerprint_sha256": fingerprint,
        "target": {"table": 1, "row": 2, "cell": 2},
        "text": "7,500",
    }


def test_active_document_range_context_refuses_shared_cell_gate(tmp_path):
    source = _split_run_fixture(tmp_path / "source.docx")
    _replace_document_xml(
        source,
        lambda data: _surround_document_with_range(
            data,
            b'<w:commentRangeStart w:id="17"/>',
            b'<w:commentRangeEnd w:id="17"/>',
        ),
    )
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    source_before = source.read_bytes()
    destination_before = destination.read_bytes()

    report = author_cell_text(parse_cell_text_request(_cell_request(source, destination)))

    assert report.passed is False
    assert report.code == "cell_context_unsupported"
    assert source.read_bytes() == source_before
    assert destination.read_bytes() == destination_before


@pytest.mark.parametrize(
    ("transform", "expected_code"),
    [
        (_make_ambiguous, "span_occurrence_ambiguous"),
        (_insert_field, "span_structure_unsupported"),
        (_insert_revision, "span_structure_unsupported"),
        (_insert_comment, "span_structure_unsupported"),
        (_insert_control, "span_structure_unsupported"),
    ],
)
def test_ambiguous_or_unsupported_span_refuses_before_writing(tmp_path, transform, expected_code):
    source = _split_run_fixture(tmp_path / "source.docx")
    _replace_document_xml(source, transform)
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    source_before = source.read_bytes()
    destination_before = destination.read_bytes()

    report = author_span_text(parse_span_text_request(_request(source, destination)))

    assert report.passed is False
    assert report.code == expected_code
    assert source.read_bytes() == source_before
    assert destination.read_bytes() == destination_before


@pytest.mark.parametrize(
    ("field", "value", "expected_code"),
    [
        ("source_sha256", "0" * 64, "source_hash_mismatch"),
        ("expected_old_text", "not the source text", "expected_old_text_mismatch"),
    ],
)
def test_stale_or_wrong_span_request_refuses_without_replacing_destination(
    tmp_path, field, value, expected_code
):
    source = _split_run_fixture(tmp_path / "source.docx")
    destination = tmp_path / "existing-output.docx"
    destination.write_bytes(b"do not replace")
    source_before = source.read_bytes()
    destination_before = destination.read_bytes()
    request_data = _request(source, destination)
    request_data[field] = value

    report = author_span_text(parse_span_text_request(request_data))

    assert report.passed is False
    assert report.code == expected_code
    assert source.read_bytes() == source_before
    assert destination.read_bytes() == destination_before


def test_span_request_rejects_unknown_fields():
    with pytest.raises(ValueError, match="span_request_fields_invalid"):
        parse_span_text_request({"source": "source.docx"})


def test_span_request_rejects_empty_old_text_and_reports_its_own_limit(tmp_path):
    source = _split_run_fixture(tmp_path / "source.docx")
    request_data = _request(source, tmp_path / "output.docx")
    request_data["expected_old_text"] = ""
    with pytest.raises(ValueError, match="span_expected_old_text_empty"):
        parse_span_text_request(request_data)

    request_data = _request(source, tmp_path / "output.docx")
    request_data["text"] = "x" * 32_769
    with pytest.raises(SpanRequestLimitError) as error:
        parse_span_text_request(request_data)
    assert error.value.to_dict() == {
        "schema_version": 1,
        "operation": "span",
        "passed": False,
        "status": "refused",
        "code": "span_text_character_limit",
        "limit": {
            "field": "text",
            "maximum_characters": 32_768,
            "requested_characters": 32_769,
            "recovery": "Shorten the value or choose a smaller span.",
        },
    }
