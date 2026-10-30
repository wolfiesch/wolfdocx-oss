"""Addressed edits preserve mixed-content runs and refuse without mutation."""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from tests.docx_fixtures import R, W, body_runs, cell, grid, row, table, write_package
from wolfdocx import Document
from wolfdocx.run_text import RunTextEditError

_DOCUMENT = "word/document.xml"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
_TEXT_TAGS = {"t", "tab", "br", "cr", "noBreakHyphen"}


def _parts(source: Path | io.BytesIO) -> dict[str, bytes]:
    with zipfile.ZipFile(source) as package:
        return {name: package.read(name) for name in package.namelist()}


def _signature(element: ET.Element) -> tuple:
    return (
        element.tag,
        sorted(element.items()),
        element.text,
        element.tail,
        tuple(_signature(child) for child in element),
    )


def _runs(root: ET.Element, ns: str = W) -> list[list[ET.Element]]:
    body = root.find(f"{{{ns}}}body")
    assert body is not None
    return [p.findall(f"{{{ns}}}r") for p in body.findall(f"{{{ns}}}p")]


def _save(document: Document) -> io.BytesIO:
    output = io.BytesIO()
    report = document.save(output)
    assert report.passed, report.to_dict()
    return output


def _assert_only_target_changed(
    source: Path, output: io.BytesIO, paragraph: int, run: int, ns: str = W
) -> ET.Element:
    before, after = _parts(source), _parts(output)
    assert before.keys() == after.keys()
    assert [name for name in before if before[name] != after[name]] == [_DOCUMENT]
    old_root, new_root = ET.fromstring(before[_DOCUMENT]), ET.fromstring(after[_DOCUMENT])
    old_run = _runs(old_root, ns)[paragraph][run]
    new_run = _runs(new_root, ns)[paragraph][run]
    assert old_run.attrib == new_run.attrib
    tags = {f"{{{ns}}}{name}" for name in _TEXT_TAGS}
    assert [_signature(c) for c in old_run if c.tag not in tags] == [
        _signature(c) for c in new_run if c.tag not in tags
    ]
    # Mask exactly the addressed run. Compare every sibling, descendant, property,
    # wrapper, other paragraph and section, rather than just their visible text.
    saved = ET.fromstring(ET.tostring(new_run))
    old_run.clear()
    new_run.clear()
    assert _signature(old_root) == _signature(new_root)
    return saved


def test_direct_body_and_run_addressing_preserves_duplicate_text(tmp_path: Path) -> None:
    repeated = '<w:p><w:r w:rsidR="00112233"><w:t>same</w:t></w:r></w:p>'
    body = (
        repeated
        + table(row(cell(repeated)), grid=grid([3000]))
        + "<w:sdt><w:sdtContent>"
        + repeated
        + "</w:sdtContent></w:sdt>"
        + "<w:p><w:hyperlink><w:r><w:t>same</w:t></w:r></w:hyperlink>"
        + '<w:ins w:id="1"><w:r><w:t>same</w:t></w:r></w:ins>'
        + '<w:r w:rsidR="44556677"><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r>'
        + '<w:r w:rsidR="8899AABB"><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>'
    )
    source = write_package(tmp_path / "source.docx", body)
    with Document(source) as document:
        document.set_run_text(1, 1, "exact target", expected_old_text="same")
        output = _save(document)
    selected = _assert_only_target_changed(source, output, 1, 1)
    assert selected.find(f"{{{W}}}t").text == "exact target"
    view = body_runs(output)
    assert [[r.text for r in p] for p in view] == [
        ["same"],
        ["same", "exact target"],
    ]
    assert view[1][1].bold is True


def test_mixed_field_drawing_object_and_note_children_survive(tmp_path: Path) -> None:
    body = (
        '<w:p><w:r w:rsidR="AABBCCDD" w:rsidRPr="10203040">'
        '<w:rPr><w:i/><w:color w:val="123456"/></w:rPr>'
        '<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF name </w:instrText>'
        '<w:fldChar w:fldCharType="separate"/><w:t>old</w:t>'
        '<w:drawing><w:opaque w:val="retained"/></w:drawing><w:tab/><w:t>tail</w:t>'
        '<w:object><w:opaque/></w:object><w:footnoteReference w:id="1"/>'
        '<w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>keep</w:t></w:r></w:p>'
    )
    notes = f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>'
    source = write_package(
        tmp_path / "source.docx",
        body,
        parts={"word/footnotes.xml": notes},
        relationships=f'<Relationship Id="rNote" Type="{R}/footnotes" Target="footnotes.xml"/>',
        overrides='<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    )
    with Document(source) as document:
        document.set_run_text(0, 0, "new\tvalue\nline", expected_old_text="old\ttail")
        output = _save(document)
    selected = _assert_only_target_changed(source, output, 0, 0)
    assert [c.tag.rsplit("}", 1)[-1] for c in selected] == [
        "rPr",
        "fldChar",
        "instrText",
        "fldChar",
        "t",
        "tab",
        "t",
        "br",
        "t",
        "drawing",
        "object",
        "footnoteReference",
        "fldChar",
    ]
    assert body_runs(output)[0][0].text == "new\tvalue\nline"


def test_controls_guard_and_emission(tmp_path: Path) -> None:
    source = write_package(
        tmp_path / "source.docx",
        "<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t>"
        "<w:cr/><w:t>D</w:t><w:noBreakHyphen/><w:t>E</w:t>"
        '<w:br w:type="page"/><w:br w:type="column"/></w:r></w:p>',
    )
    assert body_runs(source)[0][0].text == "A\tB\nC\nD-E"
    with Document(source) as document:
        document.set_run_text(0, 0, "x\ty\nz\rw\u2011end", expected_old_text="A\tB\nC\nD-E")
        output = _save(document)
    selected = _assert_only_target_changed(source, output, 0, 0)
    assert [c.tag.rsplit("}", 1)[-1] for c in selected] == [
        "t",
        "tab",
        "t",
        "br",
        "t",
        "cr",
        "t",
        "noBreakHyphen",
        "t",
    ]
    assert body_runs(output)[0][0].text == "x\ty\nz\nw-end"


@pytest.mark.parametrize("text", [" leading", "trailing ", " both ", " \t x \n end ", "A<&Ω😀"])
def test_space_and_xml_escaping(tmp_path: Path, text: str) -> None:
    source = write_package(tmp_path / "source.docx", "<w:p><w:r><w:t>old</w:t></w:r></w:p>")
    with Document(source) as document:
        document.set_run_text(0, 0, text)
        output = _save(document)
    selected = _assert_only_target_changed(source, output, 0, 0)
    assert body_runs(output)[0][0].text == text
    for node in selected.findall(f"{{{W}}}t"):
        preserve = node.text[0].isspace() or node.text[-1].isspace()
        assert node.get("{http://www.w3.org/XML/1998/namespace}space") == (
            "preserve" if preserve else None
        )


@pytest.mark.parametrize("old_children", ["<w:t>old</w:t>", "<w:tab/><w:br/>", ""])
def test_empty_text_and_empty_guard(tmp_path: Path, old_children: str) -> None:
    source = write_package(
        tmp_path / "source.docx", f"<w:p><w:r><w:rPr><w:b/></w:rPr>{old_children}</w:r></w:p>"
    )
    with Document(source) as document:
        document.set_run_text(0, 0, "")
        assert (
            ET.fromstring(document.part_bytes(_DOCUMENT)).find(f".//{{{W}}}r")[0].tag
            == f"{{{W}}}rPr"
        )
        document.set_run_text(0, 0, "added", expected_old_text="")
        output = _save(document)
    selected = _assert_only_target_changed(source, output, 0, 0)
    assert selected.find(f"{{{W}}}t").text == "added"


@pytest.mark.parametrize(
    ("paragraph", "run", "text", "guard", "code"),
    [
        (1, 0, "new", None, "paragraph_out_of_range"),
        (0, 1, "new", None, "run_out_of_range"),
        (-1, 0, "new", None, "paragraph_index_invalid"),
        (0, -1, "new", None, "run_index_invalid"),
        (True, 0, "new", None, "paragraph_index_invalid"),
        (0, 0.0, "new", None, "run_index_invalid"),
        (0, 0, "new", "wrong", "guard_mismatch"),
        (0, 0, "new", "", "guard_mismatch"),
        (0, 0, None, None, "text_invalid"),
        (0, 0, "bad\x00", None, "text_invalid"),
        (0, 0, "bad\ud800", None, "text_invalid"),
        (0, 0, "new", 42, "expected_old_text_invalid"),
        (0, 0, "x" * 32769, None, "text_limit_exceeded"),
        (0, 0, "new", "x" * 32769, "expected_old_text_limit_exceeded"),
    ],
)
def test_refusals_leave_session_and_filesystem_unchanged(
    tmp_path: Path, paragraph, run, text, guard, code: str
) -> None:
    source = write_package(tmp_path / "source.docx", "<w:p><w:r><w:t>old</w:t></w:r></w:p>")
    original = source.read_bytes()
    with Document(source) as document:
        before = document.part_bytes(_DOCUMENT)
        with pytest.raises(RunTextEditError) as failure:
            document.set_run_text(paragraph, run, text, expected_old_text=guard)
        assert failure.value.code == "run_text_" + code
        assert document.part_bytes(_DOCUMENT) == before
        assert source.read_bytes() == original
        assert list(tmp_path.iterdir()) == [source]
        assert _parts(_save(document)) == _parts(source)


def test_guard_uses_current_session_and_refusal_preserves_prior_edit(tmp_path: Path) -> None:
    source = write_package(tmp_path / "source.docx", "<w:p><w:r><w:t>old</w:t></w:r></w:p>")
    with Document(source) as document:
        document.set_run_text(0, 0, "first", expected_old_text="old")
        before = document.part_bytes(_DOCUMENT)
        with pytest.raises(RunTextEditError, match="run_text_guard_mismatch"):
            document.set_run_text(0, 0, "wrong", expected_old_text="old")
        assert document.part_bytes(_DOCUMENT) == before
        document.set_run_text(0, 0, "final", expected_old_text="first")
        output = _save(document)
    assert body_runs(output)[0][0].text == "final"


def test_strict_document_namespace_and_other_parts_preserved(tmp_path: Path) -> None:
    transitional = write_package(
        tmp_path / "transitional.docx",
        '<w:p><w:r w:rsidR="12345678"><w:t>keep</w:t></w:r>'
        '<w:r w:rsidR="87654321"><w:rPr><w:b/></w:rPr><w:t>old</w:t></w:r></w:p>',
    )
    source = tmp_path / "strict.docx"
    with zipfile.ZipFile(source, "w") as output:
        for name, data in _parts(transitional).items():
            if name.endswith((".xml", ".rels")):
                data = data.replace(W.encode(), _STRICT_W.encode()).replace(
                    R.encode(), _STRICT_R.encode()
                )
            output.writestr(name, data)
    with Document(source) as document:
        document.set_run_text(0, 1, " strict\tvalue ", expected_old_text="old")
        output = _save(document)
    selected = _assert_only_target_changed(source, output, 0, 1, _STRICT_W)
    assert [(n.tag, n.text) for n in selected if n.tag.endswith(("}t", "}tab"))] == [
        (f"{{{_STRICT_W}}}t", " strict"),
        (f"{{{_STRICT_W}}}tab", None),
        (f"{{{_STRICT_W}}}t", "value "),
    ]
    assert W.encode() not in _parts(output)[_DOCUMENT]
