from __future__ import annotations

import hashlib
import importlib
import struct
import zipfile
from pathlib import Path

import pytest

from tests.docx_fixtures import cell, paragraph, row, table, write_package
from wolfdocx import Document
from wolfdocx.authoring import CellTextRequest, author_cell_text, cell_structure_fingerprint
from wolfdocx.find_replace import FindReplaceRequest, find_replace
from wolfdocx.inspect import _inspect_bytes, _inspect_docx
from wolfdocx.package import PackageFormatError, PackageGraph
from wolfdocx.paragraph_authoring import (
    ParagraphRequest,
    author_paragraph,
    paragraph_authoring_structure_fingerprint,
)

# ``wolfdocx.find_replace`` the attribute is the function; patch the module itself.
find_replace_module = importlib.import_module("wolfdocx.find_replace")


def _request(tmp_path: Path) -> tuple[FindReplaceRequest, Path]:
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Hello"))
    destination = tmp_path / "dest.docx"
    return (
        FindReplaceRequest(
            source=source, destination=destination, pattern="Hello", replacement="Hi"
        ),
        destination,
    )


def test_find_replace_never_clobbers_destination_created_after_precheck(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    request, destination = _request(tmp_path)
    write_package_original = find_replace_module._write_package

    def stage_then_occupy(source: Path, output: Path, replacements: object) -> None:
        write_package_original(source, output, replacements)
        destination.write_bytes(b"written by another process")

    monkeypatch.setattr(find_replace_module, "_write_package", stage_then_occupy)

    report = find_replace(request)

    assert report.passed is False
    assert report.code == "destination_exists"
    assert destination.read_bytes() == b"written by another process"
    assert list(tmp_path.glob(".dest.docx.*")) == []


def test_find_replace_fails_closed_without_hard_links(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    request, destination = _request(tmp_path)

    def no_links(*_args: object) -> None:
        raise OSError("hard links unsupported")

    monkeypatch.setattr("wolfdocx.review.os.link", no_links)

    report = find_replace(request)

    assert report.passed is False
    assert report.code == "find_replace_publication_unavailable"
    assert not destination.exists()
    assert list(tmp_path.glob(".dest.docx.*")) == []


def test_package_graph_from_bytes_reads_only_the_held_bytes(tmp_path: Path) -> None:
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Held"))
    data = source.read_bytes()
    from_file = PackageGraph.from_file(source)

    graph = PackageGraph.from_bytes(data, path=source)
    source.write_bytes(b"replaced on disk")

    assert graph.file_sha256 == from_file.file_sha256
    assert set(graph.parts) == set(from_file.parts)
    assert b"Held" in graph.read_part_bytes("word/document.xml")
    with pytest.raises(PackageFormatError):
        PackageGraph.from_bytes(b"not a zip", path=source)


def test_inspect_bytes_reports_what_inspecting_the_file_reports(tmp_path: Path) -> None:
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Same"))

    from_file = _inspect_docx(source)
    from_bytes = _inspect_bytes(source.read_bytes(), display_path=str(source))

    assert from_bytes.valid and from_file.valid
    assert from_bytes.report.to_dict() == from_file.report.to_dict()


# The module implementing ``Document.save`` owns its staged-change verification.
document_module = importlib.import_module(Document.save.__module__)
STORY = "word/document.xml"
# Last direct body paragraph in ``root.iter(w:p)`` order: two cell paragraphs precede it.
LAST_BODY_PARAGRAPH = 4


def _session_source(tmp_path: Path) -> Path:
    plain = write_package(
        tmp_path / "plain.docx",
        paragraph("Hello world")
        + table(row(cell(paragraph("A1")) + cell(paragraph("B1"))))
        + paragraph("Tail"),
    )
    # Deflate styles at level 1: recompressing it at the default level would change its
    # compressed bytes, so only a raw member copy preserves them.
    source = tmp_path / "source.docx"
    with zipfile.ZipFile(plain) as archive, zipfile.ZipFile(source, "w") as rewritten:
        for info in archive.infolist():
            rewritten.writestr(
                info,
                archive.read(info),
                compresslevel=1 if info.filename == "word/styles.xml" else None,
            )
    plain.unlink()
    return source


def _members(path: Path) -> dict[str, tuple[bytes, bytes, zipfile.ZipInfo]]:
    """Each member's decompressed bytes, compressed bytes read independently of the
    library under test, and entry metadata."""
    data = path.read_bytes()
    members = {}
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            offset = info.header_offset
            name_length, extra_length = struct.unpack_from("<HH", data, offset + 26)
            start = offset + 30 + name_length + extra_length
            members[info.filename] = (
                archive.read(info),
                data[start : start + info.compress_size],
                info,
            )
    return members


def _stateless_cell(source: Path, destination: Path) -> None:
    report = author_cell_text(
        CellTextRequest(
            source=source,
            destination=destination,
            source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
            expected_old_text="A1",
            structure_fingerprint_sha256=cell_structure_fingerprint(source, table=1, row=1, cell=1)
            or "",
            table=1,
            row=1,
            cell=1,
            text="Cell",
        )
    )
    assert report.passed, report.code


def _stateless_append(source: Path, destination: Path) -> None:
    fingerprint = paragraph_authoring_structure_fingerprint(source, paragraph=LAST_BODY_PARAGRAPH)
    assert fingerprint is not None
    report = author_paragraph(
        ParagraphRequest(
            source=source,
            destination=destination,
            source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
            structure_fingerprint_sha256=fingerprint,
            paragraph=LAST_BODY_PARAGRAPH,
            action="insert",
            position="after",
            text="Appended",
        )
    )
    assert report.passed, report.code


def _stateless_replace(source: Path, destination: Path) -> None:
    report = find_replace(
        FindReplaceRequest(source, destination, pattern="world", replacement="there")
    )
    assert report.passed, report.code


@pytest.mark.parametrize(
    ("edit", "stateless"),
    [
        (lambda doc: doc.append_paragraph("Appended"), _stateless_append),
        (lambda doc: doc.set_cell_text(table=1, row=1, cell=1, text="Cell"), _stateless_cell),
        (lambda doc: doc.find_replace("world", "there"), _stateless_replace),
    ],
    ids=["append_paragraph", "set_cell_text", "find_replace"],
)
def test_session_edit_matches_stateless_output_and_copies_other_members_raw(
    tmp_path: Path, edit, stateless
) -> None:
    source = _session_source(tmp_path)
    stateless(source, tmp_path / "stateless.docx")

    with Document(source) as document:
        report = edit(document)
        assert report.passed and report.status == "applied"
        saved = document.save(tmp_path / "session.docx")

    assert saved.passed and saved.changed_parts == (STORY,)
    original = _members(source)
    session = _members(tmp_path / "session.docx")
    expected = _members(tmp_path / "stateless.docx")
    assert list(session) == list(original)
    assert session[STORY][0] == expected[STORY][0] != original[STORY][0]
    for name, (data, compressed, info) in original.items():
        if name == STORY:
            continue
        copied = session[name]
        assert copied[0] == data
        assert copied[1] == compressed
        assert (copied[2].compress_type, copied[2].date_time, copied[2].CRC) == (
            info.compress_type,
            info.date_time,
            info.CRC,
        )
    assert (
        saved.destination_sha256
        == hashlib.sha256((tmp_path / "session.docx").read_bytes()).hexdigest()
    )


def test_several_session_edits_publish_once_as_the_stateless_chain_would(tmp_path: Path) -> None:
    source = _session_source(tmp_path)
    _stateless_cell(source, tmp_path / "one.docx")
    _stateless_append(tmp_path / "one.docx", tmp_path / "two.docx")
    _stateless_replace(tmp_path / "two.docx", tmp_path / "three.docx")

    document = Document(source)
    assert document.set_cell_text(table=1, row=1, cell=1, text="Cell").passed
    assert document.append_paragraph("Appended").passed
    replaced = document.find_replace("world", "there")
    saved = document.save(tmp_path / "session.docx")

    assert replaced.replacement_count == 1
    assert saved.passed and saved.changed_parts == (STORY,)
    assert (
        _members(tmp_path / "session.docx")[STORY][0] == _members(tmp_path / "three.docx")[STORY][0]
    )


def test_session_refuses_unsupported_targets_without_changing_anything(tmp_path: Path) -> None:
    source = _session_source(tmp_path)
    document = Document(source)

    assert document.append_paragraph("x", part="word/styles.xml").code == (
        "paragraph_part_unsupported"
    )
    assert document.set_cell_text(table=9, row=1, cell=1, text="x").code == (
        "cell_coordinate_missing"
    )
    assert document.set_cell_text(
        table=1, row=1, cell=1, text="x", expected_old_text="stale"
    ).code == ("expected_old_text_mismatch")
    assert document.find_replace("x", "y", part="word/missing.xml").code == (
        "find_replace_part_unsupported"
    )
    assert document.find_replace("(", "y").code == "find_replace_pattern_invalid"
    assert document.changed_parts == ()


def test_save_never_replaces_an_existing_destination(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _session_source(tmp_path)
    document = Document(source)
    assert document.find_replace("world", "there").passed
    occupied = tmp_path / "occupied.docx"
    occupied.write_bytes(b"already here")

    assert document.save(occupied).code == "document_destination_exists"
    assert occupied.read_bytes() == b"already here"

    # A destination created after the pre-check is preserved too.
    late = tmp_path / "late.docx"
    verify = document_module._verify_staged_changes

    def verify_then_occupy(*args: object, **kwargs: object) -> object:
        result = verify(*args, **kwargs)
        late.write_bytes(b"written by another process")
        return result

    monkeypatch.setattr(document_module, "_verify_staged_changes", verify_then_occupy)
    assert document.save(late).code == "document_destination_exists"
    assert late.read_bytes() == b"written by another process"
    assert list(tmp_path.glob(".*.tmp")) == []

    document.close()
    with pytest.raises(ValueError, match="document_closed"):
        document.save(tmp_path / "closed.docx")
