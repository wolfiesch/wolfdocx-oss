"""Pin the node contract and ECMA-376 w:t-only lexical edit preservation."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.docx_fixtures import R, W, cell, paragraph, row, table, write_package
from wolfdocx.document import Document
from wolfdocx.model import (
    DocumentVisitor,
    Node,
    NodeEditError,
    NodeType,
    StoryType,
    VisitorAction,
)

C = "application/vnd.openxmlformats-officedocument.wordprocessingml."


def _fixture(tmp_path: Path) -> Path:
    stories = {
        "header2.xml": ("hdr", "H2", "header"),
        "footer1.xml": ("ftr", "F1", "footer"),
        "header1.xml": ("hdr", "H1", "header"),
        "footnotes.xml": ("footnotes", "FN", "footnotes"),
        "endnotes.xml": ("endnotes", "EN", "endnotes"),
        "comments.xml": ("comments", "CM", "comments"),
    }
    parts = {}
    for name, (tag, text, kind) in stories.items():
        content = paragraph(text)
        if kind in {"footnotes", "endnotes", "comments"}:
            entry = {
                "footnotes": "footnote",
                "endnotes": "endnote",
                "comments": "comment",
            }[kind]
            content = f'<w:{entry} w:id="4">{content}</w:{entry}>'
            if kind == "footnotes":
                content = '<w:footnote w:id="-1" w:type="separator"><w:p/></w:footnote>' + content
        parts["word/" + name] = f'<w:{tag} xmlns:w="{W}">{content}</w:{tag}>'
    parts["custom/payload.bin"] = b"unclassified extension\x00\xff"
    relationships = "".join(
        f'<Relationship Id="s{index}" Type="{R}/{kind}" Target="{name}"/>'
        for index, (name, (_, _, kind)) in enumerate(stories.items())
    )
    overrides = "".join(
        f'<Override PartName="/word/{name}" ContentType="{C}{kind}+xml"/>'
        for name, (_, _, kind) in stories.items()
    )
    # First section reverses package/relationship order. The second shares F1.
    section1 = '<w:sectPr><w:footerReference w:type="default" r:id="s1"/><w:headerReference w:type="default" r:id="s2"/></w:sectPr>'
    section2 = '<w:sectPr><w:headerReference w:type="default" r:id="s0"/><w:footerReference w:type="default" r:id="s1"/></w:sectPr>'
    anchor = (
        '<w:p><w:r><w:t xml:space="preserve">Anchor</w:t>'
        '<w:tab/><w:br w:type="page"/><w:ptab w:alignment="right" w:relativeTo="margin" w:leader="none"/>'
        "<w:pict><w:txbxContent>" + paragraph("BOX") + "</w:txbxContent></w:pict></w:r></w:p>"
    )
    body = paragraph("Body") + anchor + table(row(cell(paragraph("Cell"))))
    body += f"<w:p><w:pPr>{section1}</w:pPr><w:r><w:t>Tail</w:t></w:r></w:p>"
    return write_package(
        tmp_path / "source.docx",
        body,
        sect_properties=section2,
        relationships=relationships,
        overrides=overrides,
        parts=parts,
        content_types_extra='<Default Extension="bin" ContentType="application/octet-stream"/>',
    )


class Recorder(DocumentVisitor):
    def __init__(self) -> None:
        self.events: list[tuple[str, Node]] = []

    def visit_node_start(self, node: Node) -> VisitorAction:
        self.events.append(("start", node))
        return VisitorAction.CONTINUE

    def visit_node_end(self, node: Node) -> VisitorAction:
        self.events.append(("end", node))
        return VisitorAction.CONTINUE


def test_document_order_and_typed_balanced_callbacks(tmp_path: Path) -> None:
    document = Document(_fixture(tmp_path))
    visitor = Recorder()
    assert document.accept(visitor) == VisitorAction.CONTINUE
    starts = [node for phase, node in visitor.events if phase == "start"]
    assert [node.text for node in starts if node.node_type == NodeType.RUN] == [
        "Body",
        "Anchor",
        "BOX",
        "Cell",
        "Tail",
        "F1",
        "H1",
        "H2",
        "FN",
        "EN",
        "CM",
    ]
    assert [node.story_type for node in starts if node.node_type == NodeType.STORY] == [
        StoryType.BODY,
        StoryType.TEXT_BOX,
        StoryType.FOOTER,
        StoryType.HEADER,
        StoryType.HEADER,
        StoryType.FOOTNOTE_SEPARATOR,
        StoryType.FOOTNOTE,
        StoryType.ENDNOTE,
        StoryType.COMMENT,
    ]
    stack = []
    for phase, node in visitor.events:
        if phase == "start":
            stack.append(node)
        else:
            assert stack.pop() is node
    assert not stack
    assert len(starts) == len({id(node) for node in starts})
    assert [
        node.special_character for node in starts if node.node_type == NodeType.SPECIAL_CHARACTER
    ] == ["tab", "br", "ptab"]
    assert any(
        node.node_type == NodeType.SHAPE and node.xml_name == f"{{{W}}}pict" for node in starts
    )


def test_navigation_collections_and_deep_queries(tmp_path: Path) -> None:
    root = Document(_fixture(tmp_path)).node
    body = root.children[0]
    assert root.node_type == NodeType.DOCUMENT and root.parent is None
    assert root.previous_sibling is None and root.next_sibling is None
    assert body.parent is root and body.previous_sibling is None
    assert body.next_sibling is root.children[1]
    assert root.children[-1].next_sibling is None
    assert root.children[-1].previous_sibling is root.children[-2]
    assert len(body.paragraphs) == 3 and len(body.tables) == 1 and not body.runs
    assert [node.text for node in body.get_child_nodes(NodeType.RUN, True)] == [
        "Body",
        "Anchor",
        "BOX",
        "Cell",
        "Tail",
    ]
    assert len(body.get_child_nodes(NodeType.PARAGRAPH, True)) == 5
    assert body.paragraphs[0].runs[0].parent is body.paragraphs[0]
    with pytest.raises(TypeError, match="node_type_invalid"):
        body.get_child_nodes("run")  # type: ignore[arg-type]


@pytest.mark.parametrize("stop_at_end", [False, True])
def test_stop_is_global_without_synthetic_ends(tmp_path: Path, stop_at_end: bool) -> None:
    class Stopper(Recorder):
        def visit_run_start(self, node: Node) -> VisitorAction:
            self.events.append(("start", node))
            return VisitorAction.CONTINUE if stop_at_end else VisitorAction.STOP

        def visit_run_end(self, node: Node) -> VisitorAction:
            self.events.append(("end", node))
            return VisitorAction.STOP

    visitor = Stopper()
    assert Document(_fixture(tmp_path)).accept(visitor) == VisitorAction.STOP
    assert visitor.events[-1][0] == ("end" if stop_at_end else "start")
    assert visitor.events[-1][1].text == "Body"


def test_skip_children_still_ends_and_visits_later_siblings(tmp_path: Path) -> None:
    class Skipper(Recorder):
        def visit_table_start(self, node: Node) -> VisitorAction:
            self.events.append(("start", node))
            return VisitorAction.SKIP_CHILDREN

        def visit_run_end(self, node: Node) -> VisitorAction:
            self.events.append(("end", node))
            return VisitorAction.SKIP_CHILDREN

    visitor = Skipper()
    assert Document(_fixture(tmp_path)).accept(visitor) == VisitorAction.CONTINUE
    assert "Cell" not in [
        node.text
        for phase, node in visitor.events
        if phase == "start" and node.node_type == NodeType.RUN
    ]
    assert any(
        phase == "end" and node.node_type == NodeType.TABLE for phase, node in visitor.events
    )
    assert visitor.events[-1][1].node_type == NodeType.DOCUMENT


def test_invalid_callback_action_refused(tmp_path: Path) -> None:
    class Invalid(DocumentVisitor):
        def visit_document_start(self, node: Node) -> VisitorAction:
            return None  # type: ignore[return-value]

    with pytest.raises(TypeError, match="visitor_action_invalid"):
        Document(_fixture(tmp_path)).accept(Invalid())


def _members(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_run_rewrite_preserves_every_other_part_and_every_other_xml_byte(
    tmp_path: Path,
) -> None:
    source = _fixture(tmp_path)
    before = _members(source)
    document = Document(source)
    replacements = {
        "Body": "Rewritten & <body>",
        "BOX": "Box replacement",
        "H1": "Header replacement",
        "FN": "Note replacement",
        "CM": "Comment replacement",
    }

    class Rewriter(DocumentVisitor):
        def visit_run_start(self, node: Node) -> VisitorAction:
            if node.text in replacements:
                node.text = replacements[node.text]
            return VisitorAction.CONTINUE

    assert document.accept(Rewriter()) == VisitorAction.CONTINUE
    expected_parts = {
        "word/document.xml",
        "word/header1.xml",
        "word/footnotes.xml",
        "word/comments.xml",
    }
    assert set(document.changed_parts) == expected_parts
    destination = tmp_path / "rewritten.docx"
    assert document.save(destination).passed
    after = _members(destination)
    assert before.keys() == after.keys()
    for name in before:
        if name not in expected_parts:
            assert before[name] == after[name], name
        else:
            expected = before[name]
            for old, new in replacements.items():
                escaped = new.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                expected = expected.replace(f">{old}</w:t>".encode(), f">{escaped}</w:t>".encode())
            assert after[name] == expected, name
            mask = rb"(<w:t\b[^>]*>).*?(</w:t>)"
            assert re.sub(mask, rb"\1\2", before[name]) == re.sub(mask, rb"\1\2", after[name])
    first = destination.read_bytes()
    second = tmp_path / "again.docx"
    assert document.save(second).passed
    assert second.read_bytes() == first


def _replace_part(source: Path, name: str, data: bytes) -> None:
    members = _members(source)
    members[name] = data
    with zipfile.ZipFile(source, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for key, value in members.items():
            archive.writestr(key, value)


@pytest.mark.parametrize("encoding", ["utf-8", "utf-16-le", "utf-16-be"])
def test_multisegment_rewrites_encoding_entities_and_specials(
    tmp_path: Path, encoding: str
) -> None:
    source = write_package(
        tmp_path / "encoded.docx",
        '<w:p><w:hyperlink><w:r><w:t xml:space="preserve">A&amp;B</w:t><w:tab/><w:t xml:space="preserve">C</w:t></w:r></w:hyperlink></w:p>',
    )
    declaration = "UTF-8" if encoding == "utf-8" else "UTF-16"
    bom = {"utf-8": b"", "utf-16-le": b"\xff\xfe", "utf-16-be": b"\xfe\xff"}[encoding]
    xml = _members(source)["word/document.xml"].decode().replace("UTF-8", declaration)
    _replace_part(source, "word/document.xml", bom + xml.encode(encoding))
    document = Document(source)
    run = document.node.get_child_nodes(NodeType.RUN, True)[0]
    assert run.text == "A&BC"
    run.text = "<&\r tail"
    run.text = "<&\r tail!"
    actual = document.part_bytes("word/document.xml")
    expected = xml.replace("A&amp;B", "&lt;&amp;&#13;").replace(">C</w:t>", "> tail!</w:t>")
    assert actual == bom + expected.encode(encoding)
    assert run.text == "<&\r tail!"
    assert run.children[0].special_character == "tab"
    parsed = ET.fromstring(actual)
    assert "".join(item.text or "" for item in parsed.iter(f"{{{W}}}t")) == run.text


@pytest.mark.parametrize(
    "content,new_text,code",
    [
        ("<w:t/>", "new", "run_text_empty_element"),
        ("<w:tab/>", "new", "run_text_target_missing"),
        ("<w:t>old</w:t>", " leading", "run_text_space_not_preserved"),
        ("<w:t>old</w:t>", "bad\x00", "run_text_invalid"),
        ("<w:t>old<!--annotation--></w:t>", "new", "run_text_markup_not_preserved"),
        ("<w:t>old<?annotation keep?></w:t>", "new", "run_text_markup_not_preserved"),
        ("<w:t><![CDATA[old]]></w:t>", "new", "run_text_markup_not_preserved"),
    ],
)
def test_typed_refusals_are_atomic(tmp_path: Path, content: str, new_text: str, code: str) -> None:
    source = write_package(tmp_path / "refusal.docx", f"<w:p><w:r>{content}</w:r></w:p>")
    document = Document(source)
    before = document.part_bytes("word/document.xml")
    run = document.node.get_child_nodes(NodeType.RUN, True)[0]
    with pytest.raises(NodeEditError) as caught:
        run.text = new_text
    assert caught.value.code == code
    assert document.part_bytes("word/document.xml") == before
    assert not document.changed_parts


def test_noop_and_stale_snapshot(tmp_path: Path) -> None:
    document = Document(_fixture(tmp_path))
    run = document.node.get_child_nodes(NodeType.RUN, True)[0]
    run.text = "Body"
    assert not document.changed_parts
    assert document.append_paragraph("Added").passed
    with pytest.raises(NodeEditError, match="node_snapshot_stale"):
        run.text = "Other"
    document.close()
    with pytest.raises(ValueError, match="document_closed"):
        document.accept(Recorder())


def test_strict_word_namespace_and_noop_accept(tmp_path: Path) -> None:
    source = write_package(tmp_path / "strict.docx", paragraph("Strict"))
    before = _members(source)["word/document.xml"].replace(
        W.encode(), b"http://purl.oclc.org/ooxml/wordprocessingml/main"
    )
    _replace_part(source, "word/document.xml", before)
    document = Document(source)
    assert document.accept(DocumentVisitor()) == VisitorAction.CONTINUE
    assert document.part_bytes("word/document.xml") == before and not document.changed_parts
    run = document.node.get_child_nodes(NodeType.RUN, True)[0]
    run.text = "Updated"
    assert document.part_bytes("word/document.xml") == before.replace(
        b">Strict</w:t>", b">Updated</w:t>"
    )
