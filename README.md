# WolfDocx

**Edit existing Word `.docx` files from Python and keep every part you did not touch byte-for-byte identical. A pure-Python, MIT-licensed python-docx alternative for documents that already exist.**

WolfDocx opens a `.docx`, reads its paragraphs, runs, and tables, applies edits in
memory, and saves a new file. Each edit rewrites only the parts it changes; every
other part of the package is copied from the source unchanged. An edit that cannot be
applied safely is refused with a code and leaves the document as it was. WolfDocx
does not create documents from scratch, render pages, or convert to PDF.

[![PyPI](https://img.shields.io/pypi/v/wolfdocx)](https://pypi.org/project/wolfdocx/)
[![Python versions](https://img.shields.io/pypi/pyversions/wolfdocx)](https://pypi.org/project/wolfdocx/)
[![CI](https://github.com/wolfiesch/wolfdocx-oss/actions/workflows/ci.yml/badge.svg)](https://github.com/wolfiesch/wolfdocx-oss/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/pypi/l/wolfdocx)](https://github.com/wolfiesch/wolfdocx-oss/blob/main/LICENSE)

[Quickstart](https://github.com/wolfiesch/wolfdocx-oss#quickstart) ·
[Usage](https://github.com/wolfiesch/wolfdocx-oss#usage) ·
[Coming from python-docx](https://github.com/wolfiesch/wolfdocx-oss#coming-from-python-docx) ·
[Community and Commercial](https://github.com/wolfiesch/wolfdocx-oss#what-the-mit-package-does-not-include) ·
[Docs](https://wolfdocx.com/docs) ·
[Changelog](https://github.com/wolfiesch/wolfdocx-oss/blob/main/CHANGELOG.md) ·
[Issues](https://github.com/wolfiesch/wolfdocx-oss/issues) ·
[wolfdocx.com](https://wolfdocx.com/)

```bash
pip install wolfdocx
```

Replace a phrase that Word split across two runs, update a table cell, and save a
new file:

```python
from wolfdocx import Document

with Document("contract.docx") as doc:
    doc.find_replace(r"ACME Corp\.", "Acme Corporation")  # matches across runs
    doc.set_cell_text(table=1, row=2, cell=2, text="$7,500")
    report = doc.save("contract-final.docx")

print(report.status, report.changed_parts, report.raw_copied_parts)
```

```text
published ('word/document.xml',) 16
```

Only `word/document.xml` was rewritten. The other 16 parts of the package, including
styles, numbering, settings, theme, fonts, and relationships, were copied from the
source with their original compressed bytes. In `contract.docx`, "ACME" is a bold run
and " Corp." a plain one; the replacement takes the run formatting of the first
matched character that has any, so "Acme Corporation" comes out bold.

## Why WolfDocx

Use WolfDocx when a script has to change a document someone else maintains, such as a
contract, a report, or a template, and the rest of that file must stay exactly as it
was.

- **Edits stay local.** Each edit rewrites only the parts it needs, usually
  `word/document.xml`. Run properties, attributes, and neighboring markup stay where
  they were.
- **Untouched parts are copied, not re-serialized.** On save, every part you did not
  change keeps its compressed bytes and ZIP entry metadata. Opening a document and
  saving it without edits produces a byte-identical file.
- **Refusals instead of guesses.** An edit that cannot be applied safely (a stale
  `expected_old_text`, a missing target, an unsupported structure, a pattern with no
  match) returns `status == "refused"` and a `code`, and the document is unchanged.
  `set_run_text` raises `RunTextEditError` instead.
- **No-clobber, verified saves.** `save` stages the output, checks it, and only then
  publishes it. It never replaces an existing file.
- **Pure Python, no runtime dependencies.** The standard library is enough for
  everything except XPath selection, which needs `lxml`. No Word or LibreOffice
  install is needed.

## Capabilities

All edits are methods on `Document` and are held in memory until `save`.

| Area | API |
| --- | --- |
| Open and save | `Document(path_or_stream)`, `Document.from_bytes(data, path=...)`, `save(path_or_stream)`, `close()` |
| Read | `node`, `stories`, `get_story(StoryType...)`, `Node.paragraphs`, `runs`, `tables`, `get_child_nodes(NodeType..., deep=True)`, `to_string("txt")`, `accept(DocumentVisitor)` |
| Text | `find_replace` (regex, across runs and paragraphs), `set_run_text`, `Node.text = ...` on a run, `insert_run`, `delete_run` |
| Paragraphs | `append_paragraph`, `insert_paragraph`, `delete_paragraph`, `split_paragraph` |
| Formatting | `format_text` (bold, italic, underline, strike, color, size, character style), `format_paragraph` (style, alignment, spacing, indent, keep with next, borders, tabs) |
| Tables | `set_cell_text`, `insert_table`, `delete_table`, `insert_table_row`, `delete_table_row`, `insert_table_column`, `delete_table_column`, `merge_table_cells`, `split_table_cells` |
| Lists | `apply_list`, `remove_list`, `restart_list`, `insert_list_item` (with a numbering definition the document already has) |
| Sections | `set_page_setup` (page size, orientation, margins, columns), `insert_section_break`, `delete_section_break`, `edit_header_footer` |
| Definitions | `styles`, `lists`, `background` as editable copies you assign back |
| Approved-value edits | `wolfdocx.reporting`: discover editable cells and paragraphs, then apply edits bound to the source file's SHA-256 |

## Quickstart

Requirements: Python 3.11 or newer. WolfDocx has no runtime dependencies, and Word
and LibreOffice are not needed.

```bash
pip install wolfdocx
```

The examples in this README use `contract.docx`, a small document with four body
paragraphs and one table. Print its paragraphs and table rows:

```python
from wolfdocx import Document, NodeType, StoryType

with Document("contract.docx") as doc:
    body = doc.get_story(StoryType.BODY)

    for paragraph in body.paragraphs:
        print(paragraph.to_string("txt").rstrip("\n"))

    for table in body.tables:
        for row in table.get_child_nodes(NodeType.ROW):
            print([cell.to_string("txt").strip() for cell in row.get_child_nodes(NodeType.CELL)])
```

```text
Services Agreement
This agreement is between ACME Corp. and the Client.
Status: Draft
Payment is due within 30 days.
['Item', 'Amount']
['Setup fee', '$0']
['Monthly fee', '$1,200']
```

`Document` accepts a path or an open, seekable binary stream. Use it as a context
manager, or call `close()`, to release the source snapshot. `doc.node` and the story
nodes are snapshots: read them again after an edit to see the updated tree.

## Usage

### How targets are addressed

Read this before the edit methods, because two conventions are in use:

- `set_run_text(paragraph, run, ...)` takes **zero-based** indices that count only the
  direct paragraphs of the body and the direct runs of that paragraph. Paragraphs
  inside tables do not count.
- Every other edit method takes **1-based** numbers. `paragraph` counts every
  paragraph of the part in document order, including paragraphs inside table cells,
  and `table`, `row`, `cell`, and `column` count from 1.

In `contract.docx`, the last body paragraph ("Payment is due...") is paragraph 3 for
`set_run_text` and paragraph 10 for the other methods, because the six table cells
each hold one paragraph.

### Replace text and keep run formatting

`find_replace` runs a Python regular expression over the text of one part
(`word/document.xml` by default), so a phrase Word split across several runs still
matches. The replacement takes the run formatting of the first matched character that
has any.
The replacement can be a string or a function of the match.

```python
from wolfdocx import Document

with Document("contract.docx") as doc:
    report = doc.find_replace(
        r"acme corp\.",
        "Acme Corporation",
        options={"ignore_case": True, "max_replacements": 1},
    )
    print(report.status, report.match_count, report.replacement_count)

    report = doc.find_replace(r"\$([\d,]+)", lambda m: f"USD {m.group(1)}")
    print(report.status, report.replacement_count)

    report = doc.find_replace("Net 60", "Net 30")
    print(report.status, report.code)

    print(doc.save("contract-replaced.docx").status)
```

```text
applied 1 1
applied 2
refused find_replace_no_match
published
```

Pass `part="word/header1.xml"` (or any other part name from `doc.part_names`) to
search a header, footer, or other story. Matches inside tracked revisions and field
instructions are refused with `find_replace_protected_text`.

To change the text of one run and keep its run properties, use `set_run_text`. The
optional `expected_old_text` guard refuses the edit if the run no longer holds the
text you expect:

```python
from wolfdocx import Document
from wolfdocx.run_text import RunTextEditError

with Document("contract.docx") as doc:
    # Zero-based: body paragraph 1, run 1 is the bold "ACME" run.
    doc.set_run_text(1, 1, "Acme", expected_old_text="ACME")

    try:
        doc.set_run_text(2, 0, "Status: Final", expected_old_text="Status: Signed")
    except RunTextEditError as error:
        print("refused:", error.code)

    body = doc.node.children[0]  # read the snapshot again after the edit
    print(body.paragraphs[1].to_string("txt").rstrip("\n"))
    print(doc.save("contract-renamed.docx").status)
```

```text
refused: run_text_guard_mismatch
This agreement is between Acme Corp. and the Client.
published
```

The saved run is still `<w:r><w:rPr><w:b /></w:rPr><w:t>Acme</w:t></w:r>`.
Assigning `run.text` on a run node from the tree does the same thing, in the style of
python-docx:

```python
from wolfdocx import Document, NodeType

with Document("contract.docx") as doc:
    body = doc.node.children[0]
    for run in body.get_child_nodes(NodeType.RUN, deep=True):
        if "ACME" in run.text:
            run.text = run.text.replace("ACME", "Acme")

    report = doc.save("contract-runs.docx")
    print(report.status, report.changed_parts)
```

```text
published ('word/document.xml',)
```

### Update table cells and rows

```python
from wolfdocx import Document, NodeType

with Document("contract.docx") as doc:
    # Tables, rows, and cells are 1-based.
    print(doc.set_cell_text(table=1, row=2, cell=2, text="$7,500", expected_old_text="$0").status)
    print(doc.set_cell_text(table=1, row=3, cell=2, text="$900", expected_old_text="$1,000").code)

    print(doc.insert_table_row(1, 3, cell_texts=["Support", "$300"]).status)
    print(doc.delete_table_row(1, 2).status)

    table = doc.node.children[0].tables[0]
    for row in table.get_child_nodes(NodeType.ROW):
        print([cell.to_string("txt").strip() for cell in row.get_child_nodes(NodeType.CELL)])

    print(doc.save("contract-table.docx").status)
```

```text
applied
expected_old_text_mismatch
applied
applied
['Item', 'Amount']
['Monthly fee', '$1,200']
['Support', '$300']
published
```

`insert_table_row` clones the structure of the row it is inserted next to.
`insert_table(paragraph, rows, grid_widths)` adds a new table after (or before) a
paragraph, with column widths in twentieths of a point:

```python
doc.insert_table(10, [["Milestone", "Date"], ["Kickoff", "2026-12-01"]], [4320, 4320], repeat_header=True)
```

A `style` passed to `insert_table` must be a table style the document defines and
WolfDocx supports; otherwise the edit is refused (for example
`table_properties_style_unsupported`). `merge_table_cells` refuses a merge that would
drop text from the merged cells.

### Add paragraphs, format text, and change the page

```python
from wolfdocx import Document

with Document("contract.docx") as doc:
    doc.format_text(1, 0, 8, color="C00000")  # characters 0-8 of "Services Agreement"
    doc.format_paragraph(1, alignment="center")
    doc.insert_paragraph(3, "Effective date: 1 December 2026")  # after "Status: Draft"
    doc.append_paragraph("Signed electronically.", style="Heading2")
    doc.edit_header_footer("footer", "create", paragraphs=[{"text": "Confidential"}])
    doc.set_page_setup(orientation="landscape")

    report = doc.save("contract-extended.docx")
    print(report.status)
    print(report.changed_parts)
    print(report.added_parts)
```

```text
published
('[Content_Types].xml', 'word/_rels/document.xml.rels', 'word/document.xml')
('word/footer1.xml',)
```

`style` takes a style ID that already exists in the document (`Heading2` here); an
unknown style is refused with `paragraph_style_missing`. `alignment` accepts Word's
`ST_Jc` values such as `left`, `center`, `right`, and `both`, and python-docx
`WD_PARAGRAPH_ALIGNMENT` names or members.

### Apply approved values bound to a reviewed file

`wolfdocx.reporting` is a request-based API for pipelines that may change only
values someone approved. `discover_report_targets` lists each editable table cell
and paragraph with its current text and a structure fingerprint, plus the SHA-256 of
the source. `author_report` applies all edits or none, and writes the destination
only when every edit verifies.

```python
from wolfdocx.reporting import author_report, discover_report_targets, parse_report_request

found = discover_report_targets("contract.docx")
cells = {t["expected_old_text"]: t for t in found["targets"] if t["kind"] == "cell"}
texts = {t["expected_old_text"]: t for t in found["targets"] if t["kind"] == "text"}

report = author_report(parse_report_request({
    "source": "contract.docx",
    "destination": "contract-approved.docx",
    "source_sha256": found["source_sha256"],
    "edits": [
        {**cells["$0"], "text": "$7,500"},
        {**texts["Status: Draft"], "expected_old_text": "Draft", "text": "Final"},
    ],
}))
print(report.status, report.code, report.unselected_parts_preserved)
```

```text
published None True
```

If the source changed after discovery the request is refused with
`source_hash_mismatch`; if one edit's `expected_old_text` no longer matches, the
whole request is refused with `expected_old_text_mismatch`. In both cases no
destination file is written. Walkthroughs:
[replace text in a Word document](https://wolfdocx.com/replace-text-docx-python) and
[update table cells](https://wolfdocx.com/update-docx-table-python).

### Save safely and verify what changed

```python
import io
import zipfile

from wolfdocx import Document

with zipfile.ZipFile("contract.docx") as before, zipfile.ZipFile("contract-final.docx") as after:
    changed = [n for n in before.namelist() if before.read(n) != after.read(n)]
    print(len(before.namelist()), "parts, changed:", changed)

with Document("contract.docx") as doc:
    doc.find_replace("Draft", "Final")
    refused = doc.save("contract-final.docx")  # this file already exists
    print(refused.status, refused.code)

    buffer = io.BytesIO()
    print(doc.save(buffer).status)

with Document.from_bytes(buffer.getvalue(), path="in-memory.docx") as copy:
    print(copy.node.children[0].paragraphs[2].to_string("txt").rstrip("\n"))
```

```text
17 parts, changed: ['word/document.xml']
refused document_destination_exists
published
Status: Final
```

The save report lists `changed_parts`, `added_parts`, `removed_parts`,
`raw_copied_parts`, `destination_sha256`, and the `edits` applied, and
`report.to_dict()` returns them as a dictionary. A stream destination must be
seekable and empty. Saving a macro-enabled document under a `.docx` name is refused
with `document_destination_extension_mismatch`.

### Walk the document tree

`Document.accept` visits stories, paragraphs, runs, tables, rows, cells, fields,
bookmarks, comments, shapes, and other nodes in document order. Override the
`visit_<type>_start` and `visit_<type>_end` methods you need:

```python
from wolfdocx import Document, DocumentVisitor, VisitorAction


class Counter(DocumentVisitor):
    def __init__(self):
        self.paragraphs = 0
        self.tables = 0

    def visit_paragraph_start(self, node):
        self.paragraphs += 1
        return VisitorAction.CONTINUE

    def visit_table_start(self, node):
        self.tables += 1
        return VisitorAction.SKIP_CHILDREN  # do not count paragraphs inside tables


with Document("contract.docx") as doc:
    counter = Counter()
    doc.accept(counter)
    print(counter.paragraphs, counter.tables)
```

```text
4 1
```

`Node.select_nodes(xpath)` runs XPath 1.0 queries over the tree. It is the one
feature that needs a third-party package: install `lxml` to use it.

## Coming from python-docx

WolfDocx is not a drop-in replacement for python-docx. Its API is built around an
editing session and explicit targets. Common python-docx calls map like this:

| python-docx | WolfDocx |
| --- | --- |
| `Document("in.docx")` | `Document("in.docx")`, preferably as a context manager |
| `Document()` (new document from the default template) | Not available. Start from an existing `.docx`. |
| `doc.paragraphs`, `paragraph.text` | `doc.get_story(StoryType.BODY).paragraphs`, `paragraph.to_string("txt")` |
| `doc.tables`, `table.rows`, `row.cells`, `cell.text` | `body.tables`, `table.get_child_nodes(NodeType.ROW)`, `row.get_child_nodes(NodeType.CELL)`, `cell.to_string("txt")` |
| `run.text = "..."` | `run.text = "..."` on a run node, or `doc.set_run_text(paragraph, run, "...")` |
| `cell.text = "..."` | `doc.set_cell_text(table=1, row=2, cell=2, text="...")` |
| `doc.add_paragraph(text, style)` | `doc.append_paragraph(text, style=...)` or `doc.insert_paragraph(n, text)` |
| `doc.add_table(rows, cols)` | `doc.insert_table(paragraph, rows, grid_widths)` |
| `paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER` | `doc.format_paragraph(n, alignment="center")` |
| `section.orientation = WD_ORIENT.LANDSCAPE` | `doc.set_page_setup(orientation="landscape")` |
| `doc.save("out.docx")` | `doc.save("out.docx")`, which refuses to replace an existing file |

Differences to plan for:

- Edits go through `Document` methods and return a report. Check `report.passed` or
  `report.status` instead of assuming the edit landed.
- `doc.node` and the nodes under it are snapshots. Read them again after an edit
  method.
- Paragraph numbers are 1-based and include paragraphs in table cells, except in
  `set_run_text` (see [How targets are addressed](https://github.com/wolfiesch/wolfdocx-oss#how-targets-are-addressed)).
- python-docx has no built-in find and replace; `find_replace` handles phrases split
  across runs.
- Keep python-docx for building new documents from code. A
  [comparison](https://wolfdocx.com/python-docx-alternative) and a
  [migration checklist](https://wolfdocx.com/python-docx-migration) for the
  Commercial API are on wolfdocx.com.

## What the MIT package does not include

This repository is WolfDocx Community, the MIT-licensed read and edit core published
on PyPI. It does not:

- create a new document without a source `.docx`;
- lay out pages, render, or convert to PDF or images;
- refresh fields or rebuild a table of contents;
- run mail merge over `MERGEFIELD` templates or repeat template table rows per record;
- merge several documents into one;
- open legacy binary `.doc` files (refused with `document_format_unsupported`).

WolfDocx Commercial, an organization-scoped subscription installed from an
authenticated package index, and Wolf Runtime, licensed under a written Order, cover
the tasks in this table. Guides and access are on [wolfdocx.com](https://wolfdocx.com/):

| Task | Guide | Edition |
| --- | --- | --- |
| Fill a Word template table with rows | [fill-word-template-python](https://wolfdocx.com/fill-word-template-python) | Commercial |
| Mail merge with `MERGEFIELD` templates | [mail-merge-docx-python](https://wolfdocx.com/mail-merge-docx-python) | Commercial |
| Merge `.docx` files and keep numbering and styles | [merge-docx-python](https://wolfdocx.com/merge-docx-python) | Commercial |
| Update a table of contents without Word | [update-toc-docx-python](https://wolfdocx.com/update-toc-docx-python) | Commercial |
| Convert `.docx` to PDF without Word or LibreOffice | [convert-docx-to-pdf-python](https://wolfdocx.com/convert-docx-to-pdf-python) | [Wolf Runtime](https://wolfsdk.com/runtime) |
| WolfDocx Commercial Python API and CLI reference | [api-reference](https://wolfdocx.com/api-reference) | Commercial |

Commercial source and releases are maintained separately and are not part of this
repository.

## Compatibility

| | |
| --- | --- |
| Python | 3.11 or newer (`requires-python = ">=3.11"`). CI runs 3.11, 3.12, and 3.13 on Ubuntu. |
| Platforms | Pure-Python wheel (`py3-none-any`). No compiled extensions. |
| Dependencies | None at runtime. `lxml` only for `Node.select_nodes`. |
| Input | `.docx` packages in Transitional or Strict Open XML. Legacy binary `.doc` is refused. |
| Output | A new file or an empty seekable stream. Existing files are never replaced. |
| Office apps | Not required. WolfDocx does not call Word or LibreOffice. |

## Documentation and support

- Guides live at [wolfdocx.com/docs](https://wolfdocx.com/docs): getting started,
  [reading content](https://wolfdocx.com/docs/reading-content),
  [editing text and tables](https://wolfdocx.com/docs/editing-text-and-tables), and
  [reports and refusals](https://wolfdocx.com/docs/reports-and-cli).
- API documentation lives in the docstrings of `src/wolfdocx`; start with
  [`document.py`](https://github.com/wolfiesch/wolfdocx-oss/blob/main/src/wolfdocx/document.py),
  [`session_ops/`](https://github.com/wolfiesch/wolfdocx-oss/tree/main/src/wolfdocx/session_ops),
  and [`model.py`](https://github.com/wolfiesch/wolfdocx-oss/blob/main/src/wolfdocx/model.py).
- The [tests](https://github.com/wolfiesch/wolfdocx-oss/tree/main/tests) show each
  operation on small fixtures, including the refusal codes.
- Bugs and questions: [GitHub issues](https://github.com/wolfiesch/wolfdocx-oss/issues).
  Include the wolfdocx and Python versions, and a minimal `.docx` when you can share
  one.
- Security reports: privately, as described in
  [SECURITY.md](https://github.com/wolfiesch/wolfdocx-oss/blob/main/SECURITY.md).
- Commercial evaluation, pricing, and support: [wolfdocx.com](https://wolfdocx.com/).

## Status and versioning

WolfDocx 0.1.0, released 2026-10-29, was the first public release. 0.1.1 updates the
documentation and package metadata and leaves the library code unchanged. The PyPI
classifier is `Development Status :: 4 - Beta`. The project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html), under which a 0.x minor
release may change the public API; pin a version in production. Changes are recorded
in the [changelog](https://github.com/wolfiesch/wolfdocx-oss/blob/main/CHANGELOG.md).
The code was extracted from the author's private source repository and reduced to
this pure-Python subset; see
[PROVENANCE.md](https://github.com/wolfiesch/wolfdocx-oss/blob/main/PROVENANCE.md).

## Security

Report vulnerabilities privately through
[GitHub Security Advisories](https://github.com/wolfiesch/wolfdocx-oss/security/advisories/new),
not in a public issue. The 0.1.x line receives security fixes. See
[SECURITY.md](https://github.com/wolfiesch/wolfdocx-oss/blob/main/SECURITY.md).

When opening a package, WolfDocx enforces limits on archive size, entry count, entry
size, total uncompressed size, compression ratio, XML part size, element count, and
nesting depth, and refuses XML that declares a DTD or entities. The defaults are in
`PackageLimits`; pass `Document(path, limits=PackageLimits(...))` to change them.
WolfDocx makes no network requests and starts no processes.

Some operations keep an inspection cache, a SQLite database under
`WOLFDOCX_CACHE_DIR` or the platform user cache directory. Its stored values can
include facts derived from document parts, including part text. Set
`WOLFDOCX_INSPECTION_CACHE=off` to disable it.

## Related libraries

The same MIT-licensed family covers the other Office formats:

- [WolfXL](https://github.com/SynthGL/wolfxl-oss): read, write, and edit Excel
  `.xlsx` and `.xlsm` workbooks, an openpyxl alternative (`pip install wolfxl`).
- [WolfPPT](https://github.com/wolfiesch/wolfppt-oss): edit PowerPoint `.pptx`
  decks, a python-pptx alternative (`pip install wolfppt`).

## Contributing

```bash
python -m pip install -e ".[dev]"
ruff check .
python -m pytest
```

Open an [issue](https://github.com/wolfiesch/wolfdocx-oss/issues) before a large
change. CI runs the same lint and test commands on every push and pull request.

## License

MIT. See [LICENSE](https://github.com/wolfiesch/wolfdocx-oss/blob/main/LICENSE).
