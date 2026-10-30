# Changelog

All notable changes to this project are documented in this file. The format is
based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[0.1.0]: https://github.com/wolfiesch/wolfdocx-oss/releases/tag/v0.1.0
