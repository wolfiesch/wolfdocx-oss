"""Paragraph-anchored find and replace: the portfolio targeted and bulk edit shapes.

``^`` and ``$`` anchor at every paragraph, as they do on each paragraph's text, no
matter how paragraphs group around tables, and a session replace that changes
nothing refuses instead of reporting an applied edit.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

import wolfdocx
from tests.docx_fixtures import cell, grid, paragraph, row, table, write_package
from wolfdocx.find_replace import (
    FindReplaceOptions,
    FindReplaceRequest,
    discover_matches,
    find_replace,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _split_run_paragraph(*pieces: str) -> str:
    """One paragraph whose text is split across differently formatted runs."""
    runs = "".join(
        f'<w:r><w:rPr><w:sz w:val="{20 + 2 * index}"/></w:rPr>'
        f'<w:t xml:space="preserve">{piece}</w:t></w:r>'
        for index, piece in enumerate(pieces)
    )
    return f"<w:p>{runs}</w:p>"


def _portfolio_package(path: Path) -> Path:
    """Body paragraphs, a 2x2 table, then more body paragraphs (three paragraph groups)."""
    before = "".join(paragraph(f"Paragraph {n:05d} portfolio pending.") for n in (1, 2, 3))
    cells = "".join(
        row("".join(cell(paragraph(f"Cell {r:05d} c{c} pending")) for c in range(2)))
        for r in range(2)
    )
    after = _split_run_paragraph("Para", "graph 00004 port", "folio pending.") + paragraph(
        "Paragraph 00005 portfolio pending."
    )
    return write_package(path, before + table(cells, grid=grid([3000, 3000])) + after)


def _logical(path: Path) -> dict[str, list]:
    """Body paragraph texts and table cell texts, in document order."""
    with zipfile.ZipFile(path) as package:
        body = ET.fromstring(package.read("word/document.xml")).find(f"{{{_W}}}body")
    assert body is not None

    def text(element: ET.Element) -> str:
        return "".join(node.text or "" for node in element.iter(f"{{{_W}}}t"))

    paragraphs = [text(child) for child in body if child.tag == f"{{{_W}}}p"]
    tables = [
        [[text(tc) for tc in tr.iter(f"{{{_W}}}tc")] for tr in tbl.iter(f"{{{_W}}}tr")]
        for tbl in body
        if tbl.tag == f"{{{_W}}}tbl"
    ]
    return {"paragraphs": paragraphs, "tables": tables}


def test_bulk_anchor_replaces_every_body_paragraph_and_table_cell(
    tmp_path: Path,
) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    document = wolfdocx.Document(source)

    report = document.find_replace(r"^(Paragraph|Cell)", r"Bulk: \1")

    assert report.status == "applied"
    assert report.match_count == report.replacement_count == 5 + 4
    assert document.save(tmp_path / "out.docx").passed
    assert _logical(tmp_path / "out.docx") == {
        "paragraphs": [
            "Bulk: Paragraph 00001 portfolio pending.",
            "Bulk: Paragraph 00002 portfolio pending.",
            "Bulk: Paragraph 00003 portfolio pending.",
            "Bulk: Paragraph 00004 portfolio pending.",
            "Bulk: Paragraph 00005 portfolio pending.",
        ],
        "tables": [
            [
                ["Bulk: Cell 00000 c0 pending", "Bulk: Cell 00000 c1 pending"],
                ["Bulk: Cell 00001 c0 pending", "Bulk: Cell 00001 c1 pending"],
            ]
        ],
    }


@pytest.mark.parametrize("number", [1, 2, 4, 5])
def test_whole_paragraph_anchor_selects_one_paragraph_inside_a_group(
    tmp_path: Path, number: int
) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    expected = _logical(source)
    target = f"Paragraph {number:05d} portfolio pending."
    expected["paragraphs"][number - 1] = "Targeted: " + target
    document = wolfdocx.Document(source)

    report = document.find_replace(f"^{re.escape(target)}$", "Targeted: " + target)

    assert report.status == "applied"
    assert report.replacement_count == 1
    assert report.matches[0]["matched_text"] == target
    assert document.save(tmp_path / "out.docx").passed
    assert _logical(tmp_path / "out.docx") == expected


def test_split_run_anchor_match_keeps_the_neighbouring_paragraphs(
    tmp_path: Path,
) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    document = wolfdocx.Document(source)

    report = document.find_replace(r"^Paragraph 00004 portfolio pending\.$", "Edited 00004.")

    assert report.replacement_count == 1
    assert document.save(tmp_path / "out.docx").passed
    paragraphs = _logical(tmp_path / "out.docx")["paragraphs"]
    assert paragraphs[3:] == ["Edited 00004.", "Paragraph 00005 portfolio pending."]
    with zipfile.ZipFile(tmp_path / "out.docx") as package:
        root = ET.fromstring(package.read("word/document.xml"))
    edited = [p for p in root.iter(f"{{{_W}}}p") if "".join(p.itertext()) == "Edited 00004."]
    size = edited[0].find(f"{{{_W}}}r/{{{_W}}}rPr/{{{_W}}}sz")
    assert size is not None and size.get(f"{{{_W}}}val") == "20"


def test_targeted_replace_then_cell_edit_matches_the_portfolio_expectation(
    tmp_path: Path,
) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    expected = _logical(source)
    expected["paragraphs"][0] = "Targeted: Paragraph 00001 portfolio pending."
    expected["tables"][0][0][0] = "Targeted: Cell 00000 c0 pending"
    document = wolfdocx.Document(source)

    replaced = document.find_replace(
        r"^Paragraph 00001 portfolio pending\.$",
        "Targeted: Paragraph 00001 portfolio pending.",
    )
    cell_report = document.set_cell_text(
        table=1,
        row=1,
        cell=1,
        text="Targeted: Cell 00000 c0 pending",
        expected_old_text="Cell 00000 c0 pending",
    )

    assert replaced.replacement_count == 1 and cell_report.passed
    assert document.save(tmp_path / "out.docx").passed
    assert _logical(tmp_path / "out.docx") == expected


def test_anchor_also_starts_a_line_after_a_soft_break(tmp_path: Path) -> None:
    source = write_package(
        tmp_path / "source.docx",
        "<w:p><w:r><w:t>Alpha</w:t><w:br/><w:t>Beta</w:t></w:r></w:p>" + paragraph("Gamma"),
    )

    discovery = discover_matches(source, r"^\w+$")

    assert [match["matched_text"] for match in discovery["matches"]] == ["Alpha", "Beta", "Gamma"]


def test_no_multiline_anchors_only_at_paragraph_group_edges(tmp_path: Path) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    destination = tmp_path / "out.docx"

    report = find_replace(
        FindReplaceRequest(
            source,
            destination,
            pattern=r"^(Paragraph|Cell)",
            replacement=r"Bulk: \1",
            options=FindReplaceOptions(multiline=False),
        )
    )

    # The body before the table, each table cell, and the body after the table.
    assert report.status == "published" and report.replacement_count == 1 + 4 + 1
    assert [text.startswith("Bulk: ") for text in _logical(destination)["paragraphs"]] == [
        True,
        False,
        False,
        True,
        False,
    ]


def test_session_replace_without_a_match_refuses_and_changes_nothing(
    tmp_path: Path,
) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    document = wolfdocx.Document(source)

    missed = document.find_replace(r"^portfolio pending\.$", "never")

    assert missed.status == "refused" and missed.code == "find_replace_no_match"
    assert not missed.passed and missed.replacement_count == 0
    assert document.changed_parts == ()
    assert document.find_replace(r"^Cell 00001 c1 pending$", "Last cell").replacement_count == 1
    assert document.changed_parts == ("word/document.xml",)
    assert document.save(tmp_path / "out.docx").passed
    assert _logical(tmp_path / "out.docx")["tables"][0][1][1] == "Last cell"


def test_session_replace_with_a_zero_budget_refuses(tmp_path: Path) -> None:
    source = _portfolio_package(tmp_path / "source.docx")
    document = wolfdocx.Document(source)

    report = document.find_replace("Paragraph", "P", options={"max_replacements": 0})

    assert report.code == "find_replace_no_match"
    assert document.changed_parts == ()


def test_discovery_defaults_to_paragraph_anchors_and_groups_without_multiline(
    tmp_path: Path,
) -> None:
    source = _portfolio_package(tmp_path / "source.docx")

    default = discover_matches(source, "^Cell")
    grouped = discover_matches(source, "^Paragraph", options={"multiline": False})

    assert default["match_count"] == 4
    assert grouped["match_count"] == 2
