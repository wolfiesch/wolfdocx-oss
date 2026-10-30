"""Tests for cross-run and cross-paragraph regex find and replace in WolfDocx."""

import importlib
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from tests.docx_fixtures import (
    paragraph,
    region_relationship,
    sect,
    write_package,
)
from wolfdocx.find_replace import (
    FindReplaceOptions,
    FindReplaceRequest,
    discover_matches,
    find_replace,
    load_find_replace_request,
    parse_find_replace_request,
)
from wolfdocx.inspect import _clear_inspection_memo, inspect_docx

find_replace_module = importlib.import_module("wolfdocx.find_replace")

_WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"


def _run(text: str, *, rpr: str = "") -> str:
    rpr_xml = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _p(*runs: str, properties: str = "") -> str:
    ppr = f"<w:pPr>{properties}</w:pPr>" if properties else ""
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def _doc_text(path: Path, part: str = "word/document.xml") -> str:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read(part))
    texts = [node.text for node in root.iter(f"{{{_WORD}}}t") if node.text is not None]
    return "".join(texts)


def _doc_paragraphs(path: Path, part: str = "word/document.xml") -> list[str]:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read(part))
    result: list[str] = []
    for p in root.iter(f"{{{_WORD}}}p"):
        t_nodes = [t.text for t in p.iter(f"{{{_WORD}}}t") if t.text is not None]
        result.append("".join(t_nodes))
    return result


# ===========================================================================
# 1. Single-Paragraph Intra-Run and Cross-Run Replacement
# ===========================================================================


def test_single_paragraph_intra_run_replace(tmp_path: Path) -> None:
    """Replace substring within a single run, preserving surrounding text."""
    source = tmp_path / "intra_run.docx"
    body = _p(_run("Hello wonderful world"))
    write_package(source, body)

    dest = tmp_path / "intra_run_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"wonderful",
        replacement="glorious",
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.match_count == 1
    assert report.replacement_count == 1
    assert _doc_text(dest) == "Hello glorious world"
    assert inspect_docx(dest).valid is True


def test_single_paragraph_cross_run_formatting_preservation(tmp_path: Path) -> None:
    """Match spanning across three distinct runs preserves formatting on unreplaced segments."""
    source = tmp_path / "cross_run.docx"
    body = _p(
        _run("The quick ", rpr="<w:b/>"),
        _run("brown fox ", rpr="<w:i/>"),
        _run("jumps over the lazy dog", rpr='<w:color w:val="FF0000"/>'),
    )
    write_package(source, body)

    dest = tmp_path / "cross_run_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"quick brown fox jumps",
        replacement="sleek wolf leaped",
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.match_count == 1
    assert report.replacement_count == 1
    assert _doc_text(dest) == "The sleek wolf leaped over the lazy dog"

    # Inspect XML to verify formatting preservation
    with zipfile.ZipFile(dest) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    runs = list(root.iter(f"{{{_WORD}}}r"))

    # First run "The " should still have bold
    assert runs[0].find(f"{{{_WORD}}}rPr/{{{_WORD}}}b") is not None
    t0 = runs[0].find(f"{{{_WORD}}}t")
    assert t0 is not None and t0.text == "The "

    # Replacement run should inherit first run's formatting (bold)
    assert runs[1].find(f"{{{_WORD}}}rPr/{{{_WORD}}}b") is not None
    t1 = runs[1].find(f"{{{_WORD}}}t")
    assert t1 is not None and "sleek wolf leaped" in t1.text

    # Final run " over the lazy dog" should retain red color
    last_run = runs[-1]
    color_node = last_run.find(f"{{{_WORD}}}rPr/{{{_WORD}}}color")
    assert color_node is not None
    assert color_node.get(f"{{{_WORD}}}val") == "FF0000"
    t_last = last_run.find(f"{{{_WORD}}}t")
    assert t_last is not None and t_last.text == " over the lazy dog"


# ===========================================================================
# 2. Cross-Paragraph Matching Across Consecutive Paragraphs
# ===========================================================================


def test_cross_paragraph_replace_merges_consecutively(tmp_path: Path) -> None:
    """Regex matching across paragraph newline boundary merges paragraphs atomically."""
    source = tmp_path / "cross_para.docx"
    body = _p(_run("First paragraph ending with alpha.")) + _p(
        _run("Beta begins second paragraph.")
    )
    write_package(source, body)

    dest = tmp_path / "cross_para_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"alpha\.\nBeta",
        replacement="omega connects",
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.match_count == 1
    assert report.replacement_count == 1

    paras = _doc_paragraphs(dest)
    # Consecutive paragraphs merged into one
    assert len(paras) == 1
    assert paras[0] == "First paragraph ending with omega connects begins second paragraph."
    assert inspect_docx(dest).valid is True


def test_cross_paragraph_preserves_boundary_properties_keep_first(
    tmp_path: Path,
) -> None:
    """Cross-paragraph match with keep_first retains first paragraph's pPr."""
    source = tmp_path / "merge_keep_first.docx"
    body = _p(_run("Section Header"), properties='<w:jc w:val="center"/>') + _p(
        _run("Body line content"), properties='<w:jc w:val="right"/>'
    )
    write_package(source, body)

    dest = tmp_path / "merge_keep_first_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"Header\nBody",
        replacement="Content",
        options=FindReplaceOptions(paragraph_merge_policy="keep_first"),
    )
    report = find_replace(req)

    assert report.passed is True
    with zipfile.ZipFile(dest) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    p_nodes = list(root.iter(f"{{{_WORD}}}p"))
    assert len(p_nodes) == 1
    jc = p_nodes[0].find(f"{{{_WORD}}}pPr/{{{_WORD}}}jc")
    assert jc is not None and jc.get(f"{{{_WORD}}}val") == "center"


def test_cross_paragraph_preserves_boundary_properties_keep_last(
    tmp_path: Path,
) -> None:
    """Cross-paragraph match with keep_last adopts last paragraph's pPr."""
    source = tmp_path / "merge_keep_last.docx"
    body = _p(_run("Section Header"), properties='<w:jc w:val="center"/>') + _p(
        _run("Body line content"), properties='<w:jc w:val="right"/>'
    )
    write_package(source, body)

    dest = tmp_path / "merge_keep_last_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"Header\nBody",
        replacement="Content",
        options=FindReplaceOptions(paragraph_merge_policy="keep_last"),
    )
    report = find_replace(req)

    assert report.passed is True
    with zipfile.ZipFile(dest) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    p_nodes = list(root.iter(f"{{{_WORD}}}p"))
    assert len(p_nodes) == 1
    jc = p_nodes[0].find(f"{{{_WORD}}}pPr/{{{_WORD}}}jc")
    assert jc is not None and jc.get(f"{{{_WORD}}}val") == "right"


def test_cross_paragraph_merge_policy_combines_properties(tmp_path: Path) -> None:
    """Cross-paragraph match with merge policy combines distinct property elements."""
    source = tmp_path / "merge_props.docx"
    body = _p(_run("Para 1"), properties='<w:jc w:val="center"/>') + _p(
        _run("Para 2"), properties='<w:spacing w:after="240"/>'
    )
    write_package(source, body)

    dest = tmp_path / "merge_props_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"Para 1\nPara 2",
        replacement="Combined",
        options=FindReplaceOptions(paragraph_merge_policy="merge"),
    )
    report = find_replace(req)

    assert report.passed is True
    with zipfile.ZipFile(dest) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    p_nodes = list(root.iter(f"{{{_WORD}}}p"))
    assert len(p_nodes) == 1
    ppr = p_nodes[0].find(f"{{{_WORD}}}pPr")
    assert ppr is not None
    assert ppr.find(f"{{{_WORD}}}jc") is not None
    assert ppr.find(f"{{{_WORD}}}spacing") is not None


# ===========================================================================
# 3. Newline in Replacement Text Splits Paragraphs
# ===========================================================================


def test_replacement_with_newline_splits_into_multiple_paragraphs(
    tmp_path: Path,
) -> None:
    """Inserting newlines in replacement splits text into new consecutive paragraphs."""
    source = tmp_path / "split_repl.docx"
    body = _p(_run("Title line: content to expand here."))
    write_package(source, body)

    dest = tmp_path / "split_repl_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"content to expand here\.",
        replacement="paragraph one.\nparagraph two.",
    )
    report = find_replace(req)

    assert report.passed is True
    paras = _doc_paragraphs(dest)
    assert len(paras) == 2
    assert paras[0] == "Title line: paragraph one."
    assert paras[1] == "paragraph two."
    assert inspect_docx(dest).valid is True


# ===========================================================================
# 4. Regex Backreferences and Groups
# ===========================================================================


def test_regex_backreferences_and_capture_groups(tmp_path: Path) -> None:
    """Regex capture groups \\1, \\2 expand correctly across matched runs."""
    source = tmp_path / "groups.docx"
    body = _p(_run("Hello [USER_123], role is [ROLE_ADMIN]."))
    write_package(source, body)

    dest = tmp_path / "groups_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"\[([A-Z]+)_([A-Z0-9]+)\]",
        replacement=r"<\2:\1>",
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.match_count == 2
    assert _doc_text(dest) == "Hello <123:USER>, role is <ADMIN:ROLE>."


def test_callable_replacement_function(tmp_path: Path) -> None:
    """Callable replacement function dynamically transforms match objects."""
    source = tmp_path / "callable.docx"
    body = _p(_run("small lowercase words"))
    write_package(source, body)

    dest = tmp_path / "callable_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"\b[a-z]+\b",
        replacement=lambda m: m.group(0).upper(),
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.match_count == 3
    assert _doc_text(dest) == "SMALL LOWERCASE WORDS"


# ===========================================================================
# 5. Preservation of Untouched Structures and Tables
# ===========================================================================


def test_preserves_untouched_tables_and_metadata(tmp_path: Path) -> None:
    """Untouched paragraphs, tables, and document metadata remain identical byte-for-byte."""
    source = tmp_path / "table_doc.docx"
    tbl = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    p1 = _p(_run("Paragraph before table"))
    p2 = _p(_run("Paragraph after table with TARGET text"))
    body = p1 + tbl + p2
    write_package(source, body)

    dest = tmp_path / "table_doc_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern=r"TARGET",
        replacement="REPLACED",
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.unselected_parts_preserved is True

    # Verify table is still intact
    with zipfile.ZipFile(dest) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    tables = list(root.iter(f"{{{_WORD}}}tbl"))
    assert len(tables) == 1
    cell_text = "".join(t.text for t in tables[0].iter(f"{{{_WORD}}}t") if t.text)
    assert cell_text == "Cell 1"


def test_preserves_root_namespace_declarations_and_ignorable_prefixes(
    tmp_path: Path,
) -> None:
    """Unused root xmlns declarations survive, so every mc:Ignorable prefix stays declared."""
    namespaces = {
        "w": _WORD,
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "o": "urn:schemas-microsoft-com:office:office",
        "v": "urn:schemas-microsoft-com:vml",
        "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    }
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in namespaces.items())
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {declarations} mc:Ignorable="w14 w15 w16se wp14">'
        f"<w:body>{_p(_run('Street address line'))}</w:body></w:document>"
    )
    source = write_package(tmp_path / "namespaces.docx", "", parts={"word/document.xml": document})
    dest = tmp_path / "namespaces_out.docx"

    report = find_replace(
        FindReplaceRequest(source=source, destination=dest, pattern="ddress", replacement="ZZ")
    )

    assert report.passed is True
    assert report.replacement_count == 1
    with zipfile.ZipFile(dest) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    root_tag = re.search(r"<w:document\b[^>]*>", xml)
    assert root_tag is not None
    declared = dict(re.findall(r'xmlns:(\w+)="([^"]*)"', root_tag.group(0)))
    assert declared == namespaces
    ignorable = re.search(r'mc:Ignorable="([^"]*)"', root_tag.group(0))
    assert ignorable is not None
    assert set(ignorable.group(1).split()) <= declared.keys()
    assert _doc_text(dest) == "Street aZZ line"


# ===========================================================================
# 6. Header/Footer Story Part Support
# ===========================================================================


def test_header_story_part_find_replace(tmp_path: Path) -> None:
    """Find and replace works on reachable story parts such as word/header1.xml."""
    source = tmp_path / "header_doc.docx"
    header_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{_WORD}">{_p(_run("Confidential Header DRAFT"))}</w:hdr>'
    )
    write_package(
        source,
        _p(_run("Main document body")),
        parts={"word/header1.xml": header_xml},
        relationships=region_relationship("rIdH1", "header", "header1.xml"),
        sect_properties=sect(extra='<w:headerReference w:type="default" r:id="rIdH1"/>'),
        overrides=f'<Override PartName="/word/header1.xml" ContentType="{_HEADER_CT}"/>',
    )

    dest = tmp_path / "header_doc_out.docx"
    req = FindReplaceRequest(
        source=source,
        destination=dest,
        part="word/header1.xml",
        pattern=r"DRAFT",
        replacement="FINAL",
    )
    report = find_replace(req)

    assert report.passed is True
    assert report.match_count == 1
    assert _doc_text(dest, part="word/header1.xml") == "Confidential Header FINAL"
    # Main document body untouched
    assert _doc_text(dest, part="word/document.xml") == "Main document body"


# ===========================================================================
# 7. Discover Matches (Read-Only)
# ===========================================================================


def test_discover_matches_read_only(tmp_path: Path) -> None:
    """discover_matches returns matches across runs without modifying disk."""
    source = tmp_path / "discover.docx"
    body = _p(_run("Alpha pattern ")) + _p(_run("Beta pattern ")) + _p(_run("Gamma end"))
    write_package(source, body)

    result = discover_matches(source, r"\w+\s+pattern")
    assert result["passed"] is True
    assert result["match_count"] == 2
    assert len(result["matches"]) == 2
    assert result["matches"][0]["matched_text"] == "Alpha pattern"
    assert result["matches"][1]["matched_text"] == "Beta pattern"


# ===========================================================================
# 8. Atomic Refusals
# ===========================================================================


def test_refusal_on_missing_source(tmp_path: Path) -> None:
    """Missing source returns find_replace_source_missing refusal."""
    req = FindReplaceRequest(
        source=tmp_path / "nonexistent.docx",
        destination=tmp_path / "out.docx",
        pattern="test",
        replacement="repl",
    )
    report = find_replace(req)
    assert report.passed is False
    assert report.code == "find_replace_source_missing"
    assert not (tmp_path / "out.docx").exists()


def test_refusal_on_source_hash_changed(tmp_path: Path) -> None:
    """Mismatched source_sha256 returns source_hash_changed refusal."""
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Hello"))

    req = FindReplaceRequest(
        source=source,
        destination=tmp_path / "out.docx",
        pattern="Hello",
        replacement="Hi",
        source_sha256="0" * 64,
    )
    report = find_replace(req)
    assert report.passed is False
    assert report.code == "source_hash_changed"
    assert not (tmp_path / "out.docx").exists()


def test_refusal_on_occupied_destination(tmp_path: Path) -> None:
    """Occupied destination returns destination_exists refusal without overwriting."""
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Hello"))

    dest = tmp_path / "dest.docx"
    dest.write_text("existing content", encoding="utf-8")

    req = FindReplaceRequest(
        source=source,
        destination=dest,
        pattern="Hello",
        replacement="Hi",
    )
    report = find_replace(req)
    assert report.passed is False
    assert report.code == "destination_exists"
    assert dest.read_text(encoding="utf-8") == "existing content"


def test_refusal_on_invalid_regex_pattern(tmp_path: Path) -> None:
    """Uncompilable regex pattern returns find_replace_pattern_invalid refusal."""
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Hello"))

    req = FindReplaceRequest(
        source=source,
        destination=tmp_path / "out.docx",
        pattern="[unclosed group",
        replacement="Hi",
    )
    report = find_replace(req)
    assert report.passed is False
    assert report.code == "find_replace_pattern_invalid"
    assert not (tmp_path / "out.docx").exists()


def test_refusal_on_unsupported_story_part(tmp_path: Path) -> None:
    """Nonexistent story part returns find_replace_part_unsupported refusal."""
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Hello"))

    req = FindReplaceRequest(
        source=source,
        destination=tmp_path / "out.docx",
        part="word/header99.xml",
        pattern="Hello",
        replacement="Hi",
    )
    report = find_replace(req)
    assert report.passed is False
    assert report.code == "find_replace_part_unsupported"
    assert not (tmp_path / "out.docx").exists()


# ===========================================================================
# 9. Request Parsing and Serialization
# ===========================================================================


def test_parse_and_load_find_replace_request(tmp_path: Path) -> None:
    """Verify parse_find_replace_request and load_find_replace_request."""
    source = tmp_path / "doc.docx"
    write_package(source, paragraph("Hello"))

    dest = tmp_path / "dest.docx"
    payload = {
        "source": str(source),
        "destination": str(dest),
        "pattern": r"Hello",
        "replacement": "Hi",
        "part": "word/document.xml",
        "options": {
            "ignore_case": True,
            "paragraph_merge_policy": "keep_last",
        },
    }

    req = parse_find_replace_request(payload)
    assert req.pattern == "Hello"
    assert req.replacement == "Hi"
    assert req.options.ignore_case is True
    assert req.options.paragraph_merge_policy == "keep_last"
    assert req.output == dest

    req_path = tmp_path / "request.json"
    req_path.write_text(json.dumps(payload), encoding="utf-8")
    loaded_req = load_find_replace_request(req_path)
    assert loaded_req.pattern == "Hello"
    assert loaded_req.destination == dest


_CUSTOM_PART = b"<customData><item>PreserveMe</item></customData>"


def _verify_mode_source(path: Path) -> Path:
    body = _p(_run("Alpha TARGET "), _run("beta")) + _p(_run("gamma TARGET delta"))
    return write_package(path, body, parts={"custom/part.xml": _CUSTOM_PART})


def _changed_members(source: Path, output: Path) -> list[str]:
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        assert before.namelist() == after.namelist()
        return [name for name in before.namelist() if before.read(name) != after.read(name)]


def test_changed_parts_verification_publishes_the_same_package_as_full(
    tmp_path: Path,
) -> None:
    source = _verify_mode_source(tmp_path / "source.docx")
    outputs: dict[str, bytes] = {}
    for verify in ("full", "changed-parts"):
        _clear_inspection_memo()
        dest = tmp_path / f"{verify}.docx"
        report = find_replace(
            FindReplaceRequest(source, dest, pattern="TARGET", replacement="DONE"),
            verify=verify,
        )
        assert report.passed is True, report.code
        assert report.replacement_count == 2
        assert _changed_members(source, dest) == ["word/document.xml"]
        outputs[verify] = dest.read_bytes()
    assert outputs["changed-parts"] == outputs["full"]
    assert _doc_paragraphs(tmp_path / "changed-parts.docx") == [
        "Alpha DONE beta",
        "gamma DONE delta",
    ]


@pytest.mark.parametrize("verify", ["full", "changed-parts"])
@pytest.mark.parametrize(
    ("part", "data"),
    [
        ("custom/part.xml", b"<customData><item>Tampered</item></customData>"),
        (
            "_rels/.rels",
            (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
                b'2006/relationships/officeDocument" Target="word/document.xml"/>'
                b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/'
                b'2006/relationships/hyperlink" Target="https://example.invalid/" '
                b'TargetMode="External"/></Relationships>'
            ),
        ),
    ],
    ids=["custom-part", "package-rels"],
)
def test_staged_change_outside_the_edited_part_refuses_in_both_verify_modes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, verify: str, part: str, data: bytes
) -> None:
    source = _verify_mode_source(tmp_path / "source.docx")
    source_bytes = source.read_bytes()
    dest = tmp_path / "dest.docx"
    write_package_parts = find_replace_module._write_package

    def write_tampered(source: Path, output: Path, replacements: dict[str, bytes]) -> None:
        write_package_parts(source, output, replacements)
        with zipfile.ZipFile(output) as archive:
            entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as rewritten:
            for info, original in entries:
                rewritten.writestr(info, data if info.filename == part else original)

    monkeypatch.setattr(find_replace_module, "_write_package", write_tampered)

    report = find_replace(
        FindReplaceRequest(source, dest, pattern="TARGET", replacement="DONE"), verify=verify
    )

    assert report.passed is False
    assert report.code == "find_replace_staged_verification_failed"
    assert not dest.exists()
    assert source.read_bytes() == source_bytes
    assert sorted(path.name for path in tmp_path.iterdir()) == ["source.docx"]


def test_find_replace_rejects_unknown_verify_mode(tmp_path: Path) -> None:
    source = _verify_mode_source(tmp_path / "source.docx")
    request = FindReplaceRequest(source, tmp_path / "dest.docx", pattern="x", replacement="y")
    with pytest.raises(ValueError, match="verify_mode_invalid"):
        find_replace(request, verify="partial")
