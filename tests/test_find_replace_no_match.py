"""Standalone find and replace: zero-match refusal and element-preserving rebuilds.

A request that replaces nothing refuses with ``find_replace_no_match`` and publishes
nothing, as ``Document.find_replace`` does. A paragraph rebuilt by a replacement keeps
its ``w:tab``, ``w:br``, ``w:cr`` and ``w:noBreakHyphen`` children as elements, and every
paragraph a replacement does not touch serializes exactly as it would with no edit.
"""

from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.parsers import expat

import pytest

from tests.docx_fixtures import cell, grid, row, table, write_package
from wolfdocx import Document
from wolfdocx.find_replace import FindReplaceOptions, FindReplaceRequest, find_replace
from wolfdocx.inspect import (
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DOCUMENT = "word/document.xml"

# Editor-written XML: compact empty elements (``<w:b/>``), which the part serializer
# rewrites as ``<w:b />`` whether or not anything is edited.
_KEEP_BEFORE = (
    '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/></w:rPr>'
    "<w:t>Keep</w:t><w:tab/><w:t>tab</w:t><w:br/><w:t>break</w:t></w:r></w:p>"
)
_EDITED = (
    "<w:p><w:r><w:t>Alpha</w:t><w:tab/><w:t>Beta</w:t>"
    '<w:br w:type="textWrapping" w:clear="all"/><w:t>Gamma</w:t><w:cr/><w:t>Delta</w:t>'
    "<w:noBreakHyphen/><w:t>Eps</w:t></w:r>"
    '<w:r><w:br w:type="page"/></w:r>'
    "<w:r><w:rPr><w:i/></w:rPr><w:t>Omega</w:t></w:r></w:p>"
)
_KEEP_AFTER = '<w:p><w:r><w:t xml:space="preserve">Tail </w:t><w:tab/><w:t>end</w:t></w:r></w:p>'
_KEEP_CELL = (
    "<w:p><w:r><w:t>Cell</w:t><w:br/><w:t>kept</w:t><w:noBreakHyphen/><w:t>x</w:t></w:r></w:p>"
)
_EDITED_INDEX = 1


def _source(path: Path) -> Path:
    body = _KEEP_BEFORE + _EDITED + _KEEP_AFTER + table(row(cell(_KEEP_CELL)), grid=grid([3000]))
    return write_package(path, body)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _part(path: Path) -> bytes:
    with zipfile.ZipFile(path) as package:
        return package.read(_DOCUMENT)


def _rewrite_part(source: Path, destination: Path, data: bytes) -> Path:
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(destination, "w") as dst:
        for info in src.infolist():
            dst.writestr(info, data if info.filename == _DOCUMENT else src.read(info))
    return destination


def _paragraph_slices(data: bytes) -> tuple[list[tuple[int, int]], list[bytes]]:
    """Byte ranges and bytes of each outermost ``w:p`` in a serialized part."""
    parser = expat.ParserCreate(namespace_separator=" ")
    ranges: list[tuple[int, int]] = []
    starts: list[int] = []
    p_name = f"{_W} p"

    def start(name: str, _attrs: dict[str, str]) -> None:
        if name == p_name:
            starts.append(parser.CurrentByteIndex)

    def end(name: str) -> None:
        if name != p_name:
            return
        begin = starts.pop()
        index = parser.CurrentByteIndex
        closer = b">" if data.startswith(b"</", index) else b"/>"
        if not starts:
            ranges.append((begin, data.index(closer, index) + len(closer)))

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    return ranges, [data[begin:finish] for begin, finish in ranges]


def _without(data: bytes, span: tuple[int, int]) -> bytes:
    return data[: span[0]] + data[span[1] :]


def _request(source: Path, destination: Path, pattern: str, replacement: str, **options):
    return FindReplaceRequest(
        source=source,
        destination=destination,
        pattern=pattern,
        replacement=replacement,
        source_sha256=_sha256(source),
        options=FindReplaceOptions(**options) if options else None,
    )


def _edited_paragraph(path: Path) -> ET.Element:
    body = ET.fromstring(_part(path)).find(f"{{{_W}}}body")
    assert body is not None
    return [child for child in body if child.tag == f"{{{_W}}}p"][_EDITED_INDEX]


def _run_children(paragraph: ET.Element) -> list[tuple]:
    """Each run's content children in order as (local name, attributes, text)."""
    children = []
    for run in paragraph.iter(f"{{{_W}}}r"):
        for child in run:
            local = child.tag.rsplit("}", 1)[1]
            if local == "rPr":
                continue
            attributes = {name.rsplit("}", 1)[-1]: value for name, value in child.items()}
            children.append((local, attributes, child.text or ""))
    return children


@pytest.mark.parametrize("verify", ["full", "changed-parts"])
@pytest.mark.parametrize(
    ("pattern", "options"),
    [("Absent", {}), ("Beta", {"max_replacements": 0})],
    ids=["no-match", "zero-budget"],
)
def test_standalone_zero_replacement_refuses_and_publishes_nothing(
    tmp_path: Path, pattern: str, options: dict, verify: str
) -> None:
    source = _source(tmp_path / "source.docx")
    digest = _sha256(source)
    destination = tmp_path / "out.docx"

    report = find_replace(_request(source, destination, pattern, "x", **options), verify=verify)

    assert not report.passed
    assert (report.status, report.code) == ("refused", "find_replace_no_match")
    assert (report.match_count, report.replacement_count, report.matches) == (0, 0, [])
    assert report.source_sha256 == digest
    assert sorted(path.name for path in tmp_path.iterdir()) == ["source.docx"]
    assert _sha256(source) == digest

    session = Document(source).find_replace(
        pattern, "x", options=FindReplaceOptions(**options) if options else None
    )
    session_view = (session.status, session.code, session.match_count, session.matches)
    assert session_view == (
        report.status,
        report.code,
        report.match_count,
        report.matches,
    )


def test_rebuilt_paragraph_keeps_break_tab_and_hyphen_elements(tmp_path: Path) -> None:
    source = _source(tmp_path / "source.docx")
    destination = tmp_path / "out.docx"

    report = find_replace(_request(source, destination, "Beta", "BETA"))

    assert report.status == "published" and report.replacement_count == 1
    paragraph = _edited_paragraph(destination)
    assert _run_children(paragraph) == [
        ("t", {}, "Alpha"),
        ("tab", {}, ""),
        ("t", {}, "BETA"),
        ("br", {"type": "textWrapping", "clear": "all"}, ""),
        ("t", {}, "Gamma"),
        ("cr", {}, ""),
        ("t", {}, "Delta"),
        ("noBreakHyphen", {}, ""),
        ("t", {}, "Eps"),
        ("br", {"type": "page"}, ""),
        ("t", {}, "Omega"),
    ]
    italic = [run for run in paragraph.iter(f"{{{_W}}}r") if run.find(f".//{{{_W}}}i") is not None]
    assert [child.text for run in italic for child in run.iter(f"{{{_W}}}t")] == ["Omega"]


def test_replacement_spanning_a_tab_emits_a_tab_element(tmp_path: Path) -> None:
    source = _source(tmp_path / "source.docx")
    destination = tmp_path / "out.docx"

    report = find_replace(_request(source, destination, "Alpha\tBeta", "A\tB"))

    assert report.status == "published" and report.replacement_count == 1
    children = _run_children(_edited_paragraph(destination))
    assert children[:3] == [("t", {}, "A"), ("tab", {}, ""), ("t", {}, "B")]
    assert all("\t" not in text and "\n" not in text for _, _, text in children)


def test_untouched_paragraphs_serialize_as_with_no_edit(tmp_path: Path) -> None:
    source = _source(tmp_path / "source.docx")
    destination = tmp_path / "out.docx"
    unedited = serialize_xml_with_preserved_qnames(parse_xml_with_preserved_qnames(_part(source)))

    assert find_replace(_request(source, destination, "Beta", "BETA")).passed

    output = _part(destination)
    reference_ranges, reference = _paragraph_slices(unedited)
    output_ranges, edited = _paragraph_slices(output)
    assert len(edited) == len(reference) == 4
    for index in (0, 2, 3):
        assert edited[index] == reference[index]
    assert edited[_EDITED_INDEX] != reference[_EDITED_INDEX]
    assert _without(output, output_ranges[_EDITED_INDEX]) == _without(
        unedited, reference_ranges[_EDITED_INDEX]
    )


def test_untouched_paragraphs_keep_canonical_source_bytes(tmp_path: Path) -> None:
    word_source = _source(tmp_path / "word.docx")
    canonical = serialize_xml_with_preserved_qnames(
        parse_xml_with_preserved_qnames(_part(word_source))
    )
    source = _rewrite_part(word_source, tmp_path / "source.docx", canonical)
    destination = tmp_path / "out.docx"

    assert find_replace(_request(source, destination, "Beta", "BETA")).passed

    output = _part(destination)
    source_ranges, original = _paragraph_slices(canonical)
    output_ranges, edited = _paragraph_slices(output)
    for index in (0, 2, 3):
        assert edited[index] == original[index]
    assert _without(output, output_ranges[_EDITED_INDEX]) == _without(
        canonical, source_ranges[_EDITED_INDEX]
    )
