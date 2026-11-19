"""One DOCX package opened once, edited in memory, and published without clobbering.

:class:`Document` validates the package graph, member CRCs and bounded main-story
XML before opening. Explicit load options require whole-package admission or
verified recovery unless the caller requests deferred part validation for
diagnosis or opaque preservation. Other parts are parsed when an edit needs them.
Every edit reuses its stateless counterpart; a refusal leaves the session
unchanged. :meth:`Document.append_paragraph`, :meth:`Document.set_cell_text`,
:meth:`Document.find_replace`, and the paragraph and run edits plan in memory with
the stateless planning step. Table, picture, section, page-setup, header and footer,
list numbering, and styling edits (:mod:`wolfdocx.session_ops`), and the comment,
footnote and endnote, bookmark, hyperlink, field and table-of-contents, content
control, watermark, and text box edits (:class:`RichObjectEdits`) run the stateless
operation itself against the session's current package in a private scratch directory
and adopt its output, including parts it adds or removes.

:meth:`Document.save` copies every unchanged ZIP member's compressed bytes and
entry metadata from the source, writes only the changed and added parts, and
verifies only those: the staged package must validate as a package graph, every
unchanged member must hold the source's compressed bytes, every changed
relationship or content-type part must hold its intended bytes, each changed part
must re-parse under the package limits, and note references are re-checked when
the story or a notes part changed or a part was removed. Unchanged parts are never
parsed, so the aggregate element limit covers the changed parts only. The staged
file is published with a hard link, which never replaces an existing destination.
"""

from __future__ import annotations

import contextlib
import mmap
import shutil
import tempfile
import weakref
import zipfile
from pathlib import Path
from types import TracebackType
from typing import TYPE_CHECKING, Any, BinaryIO, Self

from .admission import _CFB_SIGNATURE, _admit, _Damage
from .package import PackageError, PackageLimits

if TYPE_CHECKING:
    from xml.etree import ElementTree as ET

    from .load_options import LoadError, LoadOptions, LoadReport
    from .render_warnings import WarningCollection
    from .session_ops import SessionEditReport
    from .session_ops.rich_objects import SessionSaveReport

__all__ = ["Document", "SessionEditReport", "SessionSaveReport"]


def _load_refusal(code: str, message: str, entry_name: str | None = None) -> LoadError:
    """Construct the public refusal type only on the refusal path."""
    from .load_options import LoadError

    return LoadError(code, message, entry_name)


class _SourceSnapshot:
    """A stream's private disk copy and the read-only map the session reads it through.

    Windows cannot delete a mapped file, so :meth:`release` closes the map before it
    removes the directory. The owning :class:`Document` registers it as a finalizer,
    so it runs once: on :meth:`Document.close`, when the session is garbage-collected,
    or at interpreter exit while the session is still open.
    """

    def __init__(self, directory: Path) -> None:
        self.directory = directory
        self.source = directory / "source.docx"
        self.mapping: mmap.mmap | None = None

    def release(self) -> None:
        if self.mapping is not None:
            self.mapping.close()
            self.mapping = None
        with contextlib.suppress(FileNotFoundError):
            shutil.rmtree(self.directory)


class _EditMethod:
    """Resolve a session operation on first use, then retain its original descriptor."""

    def __init__(self, name: str) -> None:
        self.name = name

    def _resolve(self) -> Any:
        from .session_ops.rich_objects import DocumentOperations

        value = getattr(DocumentOperations, self.name)
        setattr(Document, self.name, value)
        return value

    def __get__(self, instance: Any, owner: type) -> Any:
        return self._resolve().__get__(instance, owner)


class _EditProperty(_EditMethod):
    def __set__(self, instance: Any, value: Any) -> None:
        self._resolve().__set__(instance, value)


class Document:
    'One DOCX package opened once and edited in memory until :meth:`save`.\n\nOpen paths or open, seekable binary streams; streams are read from the start,\nremain caller-owned and are copied to a private disk snapshot. Closed,\nnon-seekable and read-failure inputs raise a named\n:class:`~wolfdocx.package.PackageError`. Default admission validates the main\nstory and internal targets. For diagnostic callers, explicitly defer part\nvalidation through :class:`~wolfdocx.load_options.LoadOptions`.\n:meth:`append_paragraph`, :meth:`set_cell_text`, and :meth:`find_replace` return\nthe report types of the stateless APIs with status ``"applied"``; every other\nedit returns a :class:`~wolfdocx.session_ops.SessionEditReport`. Nothing is\nwritten until :meth:`save`, which may be called more than once.\n\n:attr:`warnings` is cumulative diagnostic history until ``clear()``. Clearing\nit changes neither package bytes nor load, save, or render receipts.\n:meth:`render` renders current edits without publishing an intermediate DOCX.'

    def __init__(
        self,
        path: str | Path | BinaryIO,
        *,
        limits: PackageLimits | None = None,
        load_options: LoadOptions | None = None,
    ) -> None:
        self._source_snapshot: _SourceSnapshot | None = None
        if hasattr(path, "read"):
            self._open_stream(path, Path("document.docx").resolve(), limits, load_options)
        else:
            resolved = Path(path).resolve()
            with resolved.open("rb") as stream:
                self._open_stream(stream, resolved, limits, load_options)

    @classmethod
    def from_bytes(
        cls,
        data: bytes | BinaryIO,
        *,
        path: str | Path,
        limits: PackageLimits | None = None,
        load_options: LoadOptions | None = None,
    ) -> Self:
        """Open package bytes or a binary stream; ``path`` is the report identity.

        Streams are copied incrementally into a private, immutable disk snapshot.
        """
        document = cls.__new__(cls)
        document._source_snapshot = None
        if hasattr(data, "read"):
            document._open_stream(data, Path(path).resolve(), limits, load_options)
        else:
            document._open(bytes(data), Path(path).resolve(), limits, load_options)
        return document

    def _open_stream(
        self,
        stream: BinaryIO,
        path: Path,
        limits: PackageLimits | None,
        load_options: LoadOptions | None,
    ) -> None:
        if stream.closed:
            raise _load_refusal("stream_closed", "Load refused: stream_closed.")
        if not stream.seekable():
            raise _load_refusal("stream_not_seekable", "Load refused: stream_not_seekable.")
        try:
            stream.seek(0)
        except (OSError, ValueError):
            raise _load_refusal("stream_read_failed", "Load refused: stream_read_failed.") from None
        snapshot = _SourceSnapshot(
            Path(
                tempfile.mkdtemp(
                    prefix="wolfdocx-session-",
                    dir=None if load_options is None else load_options.temp_folder,
                )
            )
        )
        self._source_snapshot = snapshot
        self._release_snapshot = weakref.finalize(self, snapshot.release)
        bound = (limits or PackageLimits()).max_archive_size
        size = 0
        try:
            try:
                with snapshot.source.open("wb") as output:
                    while chunk := stream.read(64 * 1024):
                        size += len(chunk)
                        if size > bound:
                            raise _load_refusal(
                                "archive_size_limit", "Load refused: archive_size_limit."
                            )
                        output.write(chunk)
            except (OSError, ValueError):
                raise _load_refusal(
                    "stream_read_failed", "Load refused: stream_read_failed."
                ) from None
            self._open(snapshot.source, path, limits, load_options)
        except BaseException:
            self._close_snapshot()
            raise

    def _open(
        self,
        data: bytes | Path,
        path: Path,
        limits: PackageLimits | None,
        load_options: LoadOptions | None,
    ) -> None:
        self.path = path
        self._warnings: WarningCollection | None = None
        if isinstance(data, Path):
            with data.open("rb") as stream:
                signature = stream.read(len(_CFB_SIGNATURE))
        else:
            signature = data[: len(_CFB_SIGNATURE)]
        if load_options is None and signature == _CFB_SIGNATURE:
            # Use the encryption loader's named format/password refusal. Default
            # sessions cannot decrypt without explicit password options.
            from .load_options import load_package

            load_package(data, limits=limits, path=self.path)
        # Default sessions retain their absent load receipt and lazy ancillary XML
        # semantics. Explicit options choose full recovery or deferred diagnosis.
        if load_options is None:
            try:
                self._graph = _admit(data, limits or PackageLimits(), self.path, whole=False)
            except _Damage as damage:
                raise _load_refusal(
                    damage.code, f"Load refused: {damage.code}.", damage.target
                ) from None
            self.load_report = None
        else:
            from .load_options import load_package

            loaded = load_package(data, options=load_options, limits=limits, path=self.path)
            data = loaded.data
            self.load_report = loaded.report
            self._graph = loaded.graph
        self.source_sha256 = self._graph.file_sha256
        if isinstance(data, Path):
            with data.open("rb") as stream:
                self._data = mmap.mmap(stream.fileno(), 0, access=mmap.ACCESS_READ)
            if self._source_snapshot is not None:
                self._source_snapshot.mapping = self._data
        else:
            self._data = data
        # Keep the admitted source graph for unchanged part bytes; later operation
        # edits replace only the current graph.
        self._source_graph = self._graph
        self._signature_source_graph = self._source_graph
        with self._graph._open_archive() as archive:
            self._infos = {info.filename: info for info in archive.infolist()}
            self._comment = archive.comment
        # The current members in package order: source entries, then added parts.
        self._entries: dict[str, zipfile.ZipInfo] = dict(self._infos)
        # Bytes of every source part an edit changed and of every added part.
        self._changed: dict[str, bytes] = {}
        self._trees: dict[str, ET.Element] = {}
        # One read-only tree per part with the exact bytes and element count it was
        # parsed from under the session limits: an edit's verification parse, reused
        # by later edits and the save gate for identical bytes. Nothing mutates them.
        self._checked: dict[str, tuple[bytes, ET.Element, int]] = {}
        self._edits: list[dict[str, Any]] = []
        self._load_preferences = load_options
        if load_options is not None:
            self._apply_initial_load_options(load_options)

    @property
    def load_report(self) -> LoadReport | None:
        """The original load receipt, independent of clearable warning history."""
        return self._load_report

    @load_report.setter
    def load_report(self, report: LoadReport | None) -> None:
        previous = getattr(self, "_load_report", None)
        self._load_report = report
        if report is not None:
            self._notify_load_report(report, previous)

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        """Release package bytes, disk snapshot, and parsed trees."""
        if isinstance(self._data, mmap.mmap):
            self._data.close()
        self._data = None
        self._close_snapshot()
        self._changed.clear()
        self._trees.clear()
        self._checked.clear()

    def _close_snapshot(self) -> None:
        if self._source_snapshot is not None:
            self._source_snapshot = None
            self._release_snapshot()

    @property
    def changed_parts(self) -> tuple[str, ...]:
        """The source parts an edit changed and the parts edits added."""
        return tuple(sorted(self._changed))

    def part_bytes(self, name: str) -> bytes:
        """The current bytes of one part: its last edit, else the source bytes."""
        self._source()
        if name in self._changed:
            return self._changed[name]
        if name not in self._entries:
            raise PackageError("missing_part", f"Part {name!r} not in package", name)
        return self._source_graph.read_part_bytes(name)

    def _source(self) -> bytes | mmap.mmap:
        if self._data is None:
            raise ValueError("document_closed")
        return self._data


for _name in (
    "_adopt",
    "_apply_initial_load_options",
    "_apply_operation",
    "_apply_paragraph_request",
    "_apply_shared_load_options",
    "_apply_steps",
    "_bind_imported_node",
    "_checked_tree",
    "_column",
    "_current_tree",
    "_definition_part",
    "_edit_paragraph",
    "_edit_part",
    "_editable",
    "_load_dirty_fields",
    "_load_shape_math",
    "_load_unlink_pictures",
    "_materialize",
    "_merge",
    "_note",
    "_note_parts",
    "_notify_load_report",
    "_numbering",
    "_paragraph",
    "_paragraph_request",
    "_parse",
    "_parse_checked",
    "_preference_story",
    "_record",
    "_record_save_report",
    "_replace_definitions",
    "_restore",
    "_row",
    "_run_edit",
    "_section_break",
    "_set_part",
    "_snapshot",
    "_take_editable",
    "_unchanged_raw",
    "_unchanged_raw_bytes",
    "_with_load_preferences",
    "_write",
    "_write_to",
    "accept",
    "add_text_watermark",
    "append_paragraph",
    "apply_list",
    "delete_paragraph",
    "delete_picture",
    "delete_run",
    "delete_section_break",
    "delete_table",
    "delete_table_column",
    "delete_table_row",
    "edit_header_footer",
    "find_replace",
    "format_paragraph",
    "format_text",
    "get_story",
    "insert_bookmark",
    "insert_comment",
    "insert_content_control",
    "insert_endnote",
    "insert_field",
    "insert_footnote",
    "insert_hyperlink",
    "insert_list_item",
    "insert_paragraph",
    "insert_picture",
    "insert_run",
    "insert_section_break",
    "insert_story",
    "insert_table",
    "insert_table_column",
    "insert_table_of_contents",
    "insert_table_row",
    "insert_text_box",
    "merge_table_cells",
    "refresh_fields",
    "remove_list",
    "render",
    "reply_to_comment",
    "restart_list",
    "save",
    "save_text",
    "set_cell_text",
    "set_page_setup",
    "set_run_text",
    "split_paragraph",
    "split_table_cells",
):
    setattr(Document, _name, _EditMethod(_name))

for _name in (
    "background",
    "background_xml",
    "editing_language",
    "lists",
    "load_recovery",
    "node",
    "part_names",
    "stories",
    "styles",
    "warnings",
    "web_add_ins",
):
    setattr(Document, _name, _EditProperty(_name))

del _name


def __getattr__(name: str) -> Any:
    if name == "SessionEditReport":
        from .session_ops import SessionEditReport

        return SessionEditReport
    if name in {
        "_raw_member",
        "_convert_ooxml",
        "_without_zip64_extra",
        "_structural",
        "_parse_part",
        "_span",
        "_copy_member_raw",
        "_save_ooxml",
        "_legacy_control_ooxml",
        "_xml_utf8",
        "_main_part",
        "_validate_node_structure",
        "_patch_node_texts",
        "RichObjectEdits",
        "_first_edition_ooxml",
        "_added_entry",
        "_paragraph_target",
        "_PART_ERRORS",
        "_save_text",
        "SessionSaveReport",
        "_OOXML_NAMESPACES",
        "_shape_text",
        "_plan_paragraph_edit",
    }:
        from .session_ops import rich_objects

        value = getattr(rich_objects, name)
        globals()[name] = value
        return value
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
