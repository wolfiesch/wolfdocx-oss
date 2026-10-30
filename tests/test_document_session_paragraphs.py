"""Paragraph sessions use the shared planner without staging a package per edit."""

from __future__ import annotations

import hashlib
import importlib
import struct
import zipfile
from collections.abc import Mapping
from dataclasses import replace
from pathlib import Path
from typing import Any

import pytest

from tests.docx_fixtures import (
    R,
    W,
    cell,
    footer_reference,
    grid,
    header_reference,
    paragraph,
    row,
    sect,
    styles_xml,
    table,
    write_package,
)
from wolfdocx import Document
from wolfdocx.paragraph_authoring import (
    ParagraphRequest,
    author_paragraph,
    paragraph_authoring_structure_fingerprint,
)
from wolfdocx.regions import RegionRequest, author_region
from wolfdocx.session_ops._bridge import operation_detail

# The module implementing ``Document.save`` also plans in-memory paragraph edits.
document_module = importlib.import_module(Document.save.__module__)
STORY = "word/document.xml"


def _source(tmp_path: Path, *, body: str | None = None, regions: bool = True) -> Path:
    if body is None:
        body = (
            paragraph("Alpha Beta", run_properties="<w:b/>")
            + paragraph("Middle", style="Normal")
            + table(row(cell(paragraph("Cell"))), grid=grid([2400]))
            + paragraph("Tail")
        )
    parts: dict[str, str] = {
        "customXml/item1.xml": '<opaque xmlns="urn:session-test">  untouched  </opaque>',
        "word/styles.xml": styles_xml().replace(
            '<w:name w:val="Heading"/>', '<w:name w:val="Section heading"/>'
        ),
    }
    relationships = overrides = references = ""
    if regions:
        for kind, tag in (("header", "hdr"), ("footer", "ftr")):
            parts[f"word/{kind}1.xml"] = (
                f'<w:{tag} xmlns:w="{W}">{paragraph(kind.title())}{paragraph("Tail")}</w:{tag}>'
            )
            relationships += (
                f'<Relationship Id="rId{kind}" Type="{R}/{kind}" Target="{kind}1.xml"/>'
            )
            overrides += (
                f'<Override PartName="/word/{kind}1.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
            )
        references = header_reference("rIdheader") + footer_reference("rIdfooter")
    seed = write_package(
        tmp_path / "seed.docx",
        body,
        parts=parts,
        relationships=relationships,
        overrides=overrides,
        sect_properties=sect(references=references),
    )
    source = tmp_path / "source.docx"
    # An unusual compression level and metadata make recompression observable.
    with zipfile.ZipFile(seed) as archive, zipfile.ZipFile(source, "w") as output:
        output.comment = b"session paragraph fixture"
        for info in archive.infolist():
            info.comment = b"preserve member metadata"
            info.external_attr = 0o640 << 16
            output.writestr(info, archive.read(info), compresslevel=1)
    seed.unlink()
    return source


def _parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()}


def _current(document: Document) -> dict[str, bytes]:
    return {name: document.part_bytes(name) for name in document.part_names}


def _raw(path: Path) -> dict[str, tuple[bytes, tuple[Any, ...]]]:
    """Read compressed members independently of Document's raw-copy implementation."""
    data = path.read_bytes()
    members = {}
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            name_length, extra_length = struct.unpack_from("<HH", data, info.header_offset + 26)
            start = info.header_offset + 30 + name_length + extra_length
            members[info.filename] = (
                data[start : start + info.compress_size],
                (
                    info.compress_type,
                    info.date_time,
                    info.CRC,
                    info.file_size,
                    info.compress_size,
                    info.external_attr,
                    info.internal_attr,
                    info.create_system,
                    info.extra,
                    info.comment,
                ),
            )
    return members


def _stateless(source: Path, destination: Path, index: int, **fields: Any) -> Any:
    part = fields.get("part", STORY)
    return author_paragraph(
        ParagraphRequest(
            source=source,
            destination=destination,
            source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
            structure_fingerprint_sha256=(
                paragraph_authoring_structure_fingerprint(source, paragraph=index, part=part)
                or "0" * 64
            ),
            paragraph=index,
            **fields,
        )
    )


def _edit(document: Document, index: int, fields: Mapping[str, Any]) -> Any:
    fields = dict(fields)
    action = fields.pop("action")
    return getattr(document, f"{action}_paragraph")(index, **fields)


def _disable_bridge(monkeypatch: pytest.MonkeyPatch, source: Path) -> None:
    def forbidden(*_args: object, **_kwargs: object) -> Any:
        raise AssertionError("a paragraph session edit used the disk bridge")

    monkeypatch.setattr(Document, "_materialize", forbidden)
    monkeypatch.setattr(Document, "_apply_operation", forbidden)
    read_bytes = Path.read_bytes

    def held_bytes(path: Path) -> bytes:
        if path.resolve() == source.resolve():
            raise AssertionError("a paragraph session edit reread its source")
        return read_bytes(path)

    monkeypatch.setattr(Path, "read_bytes", held_bytes)


@pytest.mark.parametrize(
    ("index", "fields"),
    [
        (1, {"action": "insert", "position": "before", "text": "Before"}),
        (1, {"action": "insert", "text": "After", "style": "Section heading"}),
        (2, {"action": "delete"}),
        (1, {"action": "split", "offset": 6}),
        (3, {"action": "insert", "text": "In cell"}),
        (1, {"action": "insert", "text": "Header edit", "part": "word/header1.xml"}),
        (1, {"action": "split", "offset": 3, "part": "word/footer1.xml"}),
        (1, {"action": "delete", "part": "word/header1.xml"}),
    ],
    ids=[
        "before",
        "style-name",
        "delete",
        "formatted-split",
        "cell",
        "header",
        "footer",
        "header-delete",
    ],
)
def test_paragraph_planner_matches_stateless_parts_and_report_detail(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, index: int, fields: Mapping[str, Any]
) -> None:
    source = _source(tmp_path)
    expected = tmp_path / "expected.docx"
    stateless = _stateless(source, expected, index, **fields)
    assert stateless.passed, stateless.code
    expected_parts = _parts(expected)
    document = Document(source)
    _disable_bridge(monkeypatch, source)

    report = _edit(document, index, fields)

    assert report.passed and report.status == "applied"
    assert report.changed_parts == (fields.get("part", STORY),)
    assert report.removed_parts == ()
    assert report.detail == operation_detail(stateless)
    assert _current(document) == expected_parts
    assert document.save(tmp_path / "session.docx").passed
    assert _parts(tmp_path / "session.docx") == expected_parts
    document.close()


def test_chained_paragraph_edits_save_once_and_preserve_untouched_compressed_members(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _source(tmp_path)
    original = _raw(source)
    edits = [
        (1, {"action": "insert", "text": "New text", "position": "before"}),
        (2, {"action": "split", "offset": 6}),
        (4, {"action": "delete"}),
    ]
    expected = source
    expected_reports = []
    expected_parts = []
    for step, (index, fields) in enumerate(edits):
        destination = tmp_path / f"step-{step}.docx"
        report = _stateless(expected, destination, index, **fields)
        assert report.passed, report.code
        expected_reports.append(report)
        expected_parts.append(_parts(destination))
        expected = destination
    document = Document(source)
    source_hash = document.source_sha256
    _disable_bridge(monkeypatch, source)
    destination = tmp_path / "session.docx"

    for (index, fields), stateless, parts in zip(
        edits, expected_reports, expected_parts, strict=True
    ):
        report = _edit(document, index, fields)
        assert report.passed and report.detail == operation_detail(stateless)
        assert _current(document) == parts
        assert not destination.exists()
    saved = document.save(destination)

    assert saved.passed and saved.source_sha256 == source_hash
    assert saved.changed_parts == (STORY,)
    assert saved.added_parts == saved.removed_parts == ()
    assert [entry["operation"] for entry in saved.edits] == ["paragraph"] * len(edits)
    assert saved.raw_copied_parts == len(original) - 1
    assert _parts(destination) == expected_parts[-1]
    published = _raw(destination)
    for name in original.keys() - {STORY}:
        assert published[name] == original[name], name
    with zipfile.ZipFile(destination) as archive:
        assert archive.comment == b"session paragraph fixture"
    document.close()


@pytest.mark.parametrize(
    ("body", "index", "fields"),
    [
        (None, 999, {"action": "insert", "text": "Missing"}),
        (None, 1, {"action": "insert", "text": "Missing region", "part": "word/header999.xml"}),
        (None, 1, {"action": "insert", "text": "Wrong part", "part": "word/styles.xml"}),
        (None, 3, {"action": "delete"}),
        (
            '<w:p><w:bookmarkStart w:id="0" w:name="Anchor"/><w:r><w:t>Alpha</w:t></w:r>'
            '<w:bookmarkEnd w:id="0"/></w:p>' + paragraph("Tail"),
            1,
            {"action": "split", "offset": 2},
        ),
        (
            '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>Date</w:t></w:r></w:fldSimple></w:p>'
            + paragraph("Tail"),
            1,
            {"action": "split", "offset": 2},
        ),
        (None, 1, {"action": "split", "offset": 999}),
    ],
    ids=[
        "missing",
        "missing-part",
        "unsupported-part",
        "sole-cell-paragraph",
        "bookmark",
        "field",
        "offset",
    ],
)
def test_refusals_keep_current_parts_and_receipts_with_stateless_detail(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    body: str | None,
    index: int,
    fields: Mapping[str, Any],
) -> None:
    source = _source(tmp_path, body=body)
    initial_index = 4 if body is None else 2
    previous = tmp_path / "previous.docx"
    assert _stateless(source, previous, initial_index, action="insert", text="Accepted").passed
    stateless = _stateless(previous, tmp_path / "refused.docx", index, **fields)
    assert not stateless.passed
    document = Document(source)
    _disable_bridge(monkeypatch, source)
    assert document.insert_paragraph(initial_index, "Accepted").passed
    before = _current(document)
    changed = document.changed_parts

    report = _edit(document, index, fields)

    assert not report.passed and report.code == stateless.code
    assert report.detail == operation_detail(stateless)
    assert report.changed_parts == report.removed_parts == ()
    assert _current(document) == before and document.changed_parts == changed
    saved = document.save(tmp_path / "session.docx")
    assert saved.passed and [edit["operation"] for edit in saved.edits] == ["paragraph"]
    document.close()


@pytest.mark.parametrize(
    ("index", "fields", "exception", "message"),
    [
        (True, {"action": "insert", "text": "x"}, ValueError, "paragraph_paragraph_invalid"),
        (
            1,
            {"action": "insert", "text": "x", "position": "inside"},
            ValueError,
            "paragraph_position_invalid",
        ),
        (1, {"action": "insert", "text": 10}, TypeError, "paragraph_text_invalid"),
        (1, {"action": "split", "offset": -1}, ValueError, "paragraph_offset_invalid"),
        (1, {"action": "delete", "part": ""}, ValueError, "paragraph_part_invalid"),
    ],
)
def test_request_validation_errors_leave_session_unchanged(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    index: int,
    fields: Mapping[str, Any],
    exception: type[Exception],
    message: str,
) -> None:
    source = _source(tmp_path)
    document = Document(source)
    _disable_bridge(monkeypatch, source)
    assert document.insert_paragraph(4, "Accepted").passed
    before = _current(document)
    changed = document.changed_parts
    with pytest.raises(exception, match=message):
        _edit(document, index, fields)
    assert _current(document) == before and document.changed_parts == changed
    saved = document.save(tmp_path / "session.docx")
    assert saved.passed and len(saved.edits) == 1
    document.close()


@pytest.mark.parametrize("malformed", [False, True], ids=["failed-verification", "invalid-xml"])
def test_failed_staged_paragraph_verification_does_not_adopt_mutated_tree(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, malformed: bool
) -> None:
    source = _source(tmp_path)
    document = Document(source)
    assert document.insert_paragraph(4, "Accepted").passed
    before = _current(document)
    changed = document.changed_parts
    planner = document_module._plan_paragraph_edit

    def bad_plan(*args: Any, **kwargs: Any) -> Any:
        plan = planner(*args, **kwargs)
        assert not isinstance(plan, str)
        return (
            replace(plan, serialized=b"not XML")
            if malformed
            else replace(plan, verify_story=lambda _: False)
        )

    _disable_bridge(monkeypatch, source)
    monkeypatch.setattr(document_module, "_plan_paragraph_edit", bad_plan)
    report = document.split_paragraph(1, 3)
    assert not report.passed and report.code == "paragraph_staged_verification_failed"
    assert report.detail["target"] == {"part": STORY, "paragraph": 1, "offset": 3}
    assert report.detail["action"] == "split"
    assert _current(document) == before and document.changed_parts == changed
    saved = document.save(tmp_path / "session.docx")
    assert saved.passed and len(saved.edits) == 1
    document.close()


def test_inverse_paragraph_edits_restore_original_raw_member(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    seed = _source(tmp_path)
    inserted = tmp_path / "normalized-insert.docx"
    source = tmp_path / "normalized.docx"
    # Start with the stateless planner's lexical form so the inverse restores exact bytes.
    assert _stateless(seed, inserted, 1, action="insert", text="Normalize").passed
    assert _stateless(inserted, source, 2, action="delete").passed
    original_parts, original_raw = _parts(source), _raw(source)
    document = Document(source)
    _disable_bridge(monkeypatch, source)

    assert document.insert_paragraph(1, "Temporary").passed
    assert document.changed_parts == (STORY,)
    assert document.delete_paragraph(2).passed
    assert _current(document) == original_parts and document.changed_parts == ()
    destination = tmp_path / "session.docx"
    saved = document.save(destination)

    assert saved.passed and saved.changed_parts == saved.added_parts == saved.removed_parts == ()
    assert saved.raw_copied_parts == len(original_raw)
    assert len(saved.edits) == 2 and _raw(destination) == original_raw
    document.close()


def test_paragraph_planner_edits_regions_added_by_existing_session_bridge(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _source(tmp_path, regions=False)
    previous = source
    for kind in ("header", "footer"):
        destination = tmp_path / f"create-{kind}.docx"
        report = author_region(
            RegionRequest(
                source=previous,
                destination=destination,
                source_sha256=hashlib.sha256(previous.read_bytes()).hexdigest(),
                kind=kind,
                action="create",
                paragraphs=[{"text": kind.title()}],
            )
        )
        assert report.passed, report.code
        previous = destination
    edits = [
        (1, {"action": "insert", "text": "Added header paragraph", "part": "word/header1.xml"}),
        (1, {"action": "split", "offset": 3, "part": "word/footer1.xml"}),
    ]
    reports = []
    for step, (index, fields) in enumerate(edits):
        destination = tmp_path / f"region-edit-{step}.docx"
        report = _stateless(previous, destination, index, **fields)
        assert report.passed, report.code
        reports.append(report)
        previous = destination
    expected = _parts(previous)
    document = Document(source)
    for kind in ("header", "footer"):
        report = document.edit_header_footer(kind, "create", paragraphs=[{"text": kind.title()}])
        assert report.passed, report.code
    _disable_bridge(monkeypatch, source)
    for (index, fields), expected_report in zip(edits, reports, strict=True):
        report = _edit(document, index, fields)
        assert report.passed and report.detail == operation_detail(expected_report)
    assert _current(document) == expected
    saved = document.save(tmp_path / "session.docx")
    assert saved.passed and set(saved.added_parts) == {"word/header1.xml", "word/footer1.xml"}
    assert [entry["operation"] for entry in saved.edits] == ["header-footer"] * 2 + [
        "paragraph"
    ] * 2
    assert _parts(tmp_path / "session.docx") == expected
    document.close()
