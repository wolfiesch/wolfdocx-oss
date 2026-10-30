# wolfdocx

Edit .docx files in Python. wolfdocx is a python-docx alternative that opens an
existing Word document, applies your edits in place, and saves a new package where
every part you did not touch is copied byte-for-byte from the original.

- Pure Python, no runtime dependencies.
- Python 3.11 or newer.
- MIT licensed.

## Install

```bash
pip install wolfdocx
```

## Quickstart

Open an existing `.docx`, read its paragraphs and tables, make a few edits, and save
the result to a new file.

```python
from wolfdocx import Document

with Document("contract.docx") as doc:
    # The first story under doc.node is the main document body.
    body = doc.node.children[0]

    for paragraph in body.paragraphs:
        print(paragraph.to_string("txt").rstrip("\n"))

    for table in body.tables:
        print(table.to_string("txt"))

    # Replace the text of one run. Paragraph and run indices are zero-based and
    # count the direct paragraphs of the body and the direct runs of that paragraph.
    first_run = body.paragraphs[0].runs[0]
    doc.set_run_text(0, 0, "Master Services Agreement", expected_old_text=first_run.text)

    # Regular-expression find and replace across runs.
    replaced = doc.find_replace(r"ACME Corp\.?", "Acme Corporation")
    print(replaced.status, replaced.replacement_count)

    # Add a plain paragraph after the last body paragraph.
    added = doc.append_paragraph("Signed electronically.")
    print(added.status)

    report = doc.save("contract-edited.docx")
    print(report.status, report.changed_parts, report.raw_copied_parts)
```

Edits are held in memory until `save(path)` writes a new file, and `doc.node` is a
snapshot: read it again after an edit to see the updated tree.

## When to use WolfDocx instead of python-docx

Use WolfDocx when you need targeted edits to an existing document and want every
untouched package part to remain byte-for-byte identical.

- **In-place edits.** Each edit rewrites only the XML part it targets. Run
  properties, attributes, and neighboring markup stay where they were.
- **Byte-for-byte preservation.** On save, every package part you did not change is
  copied from the source archive unchanged. `report.changed_parts` lists the parts
  that were rewritten and `report.raw_copied_parts` counts the parts copied as-is.
- **Verified, no-clobber saves.** The output is staged and checked before it is
  published. `save` refuses to replace an existing file and reports
  `status == "refused"` with a `code` explaining why.
- **Explicit refusals.** An edit that cannot be applied safely, such as a
  `find_replace` pattern with no match, returns a report with `status == "refused"`
  and leaves the document unchanged. `set_run_text` raises `RunTextEditError` for
  bad indices or an `expected_old_text` mismatch.

wolfdocx works on `.docx` packages. It does not render pages, convert to PDF, or
read legacy binary `.doc` files.

## Commercial capabilities

Commercial capabilities are available at https://wolfdocx.com.

## Contributing

```bash
python -m pip install -e ".[dev]"
ruff check .
python -m pytest
```

Report bugs at https://github.com/wolfiesch/wolfdocx-oss/issues. For security
issues, see [SECURITY.md](SECURITY.md).

## License

MIT. See [LICENSE](LICENSE).
