# Changelog

All notable changes to this project are documented in this file. The format is
based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.2] - 2026-11-18

### Changed

- The package is regenerated from one source commit. Every module is an unmodified
  copy of that commit's file, and `PROVENANCE.md` names the commit and lists the
  SHA-256 of each module.
- The distribution version is read from `wolfdocx/_version.py`, so the installed
  version and `wolfdocx.__version__` always agree.
- Requesting a capability this package does not ship, such as PDF output through
  `SaveOptions.create("pdf")`, raises `ModuleNotFoundError` with the message
  `commercial_feature_unavailable`.

### Added

- The modules the regenerated core imports: load options, contract records,
  effective paragraph formatting, style resolution, OOXML child ordering, web
  extension part preservation, picture session edits, and the optional-feature
  loader.
- `wolfdocx.builder.DocumentBuilder` for new documents with paragraphs, runs, tables
  and sections, and `wolfdocx.comments` for comment threads, so the `wolfdocx-docx`
  import alias runs on this package alone. Builder pictures, shapes, hyperlinks,
  document settings and fields raise
  `ModuleNotFoundError("commercial_feature_unavailable")`.

## [0.1.1] - 2026-11-09

Documentation and package metadata only. The library code is unchanged from 0.1.0.

### Changed

- README rewritten with a runnable example for each operation, a python-docx
  comparison, the list of tasks the MIT package does not cover, and links to the
  guides at [wolfdocx.com/docs](https://wolfdocx.com/docs).
- Package description and keywords no longer suggest that WolfDocx creates
  documents from scratch; it reads and edits existing `.docx` files.

## [0.1.0] - 2026-10-29

### Added

- First public release under the MIT license.
- `Document` session for opening an existing `.docx`, with in-memory edits
  published only by `save`.
- Node navigation through `Document.node` for reading stories, paragraphs, runs,
  and tables.
- Edits: `set_run_text`, `find_replace`, `append_paragraph`, and related session
  operations.
- Verified, no-clobber saves that copy every untouched package part byte-for-byte
  and report changed and raw-copied parts.

[0.1.2]: https://github.com/wolfiesch/wolfdocx-oss/releases/tag/v0.1.2
[0.1.1]: https://github.com/wolfiesch/wolfdocx-oss/releases/tag/v0.1.1
[0.1.0]: https://github.com/wolfiesch/wolfdocx-oss/releases/tag/v0.1.0
