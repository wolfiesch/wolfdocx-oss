"""Open a DOCX once, edit in memory, and preserve untouched ZIP members."""

from __future__ import annotations

import mmap
import tempfile
import zipfile
from pathlib import Path
from types import TracebackType
from typing import TYPE_CHECKING, Any, BinaryIO, Self

from .admission import _CFB_SIGNATURE, _admit, _Damage
from .package import PackageError, PackageLimits

if TYPE_CHECKING:
    from xml.etree import ElementTree as ET

    from .render_warnings import WarningCollection

__all__ = ["Document"]


class _EditMethod:
    """Resolve an editing operation on first use and retain its descriptor."""

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
    """A validated DOCX package with in-memory editing and no-clobber saves.

    Accept a path or an open, seekable binary stream. Streams remain caller-owned
    and are copied to a private snapshot. Changes are written only by ``save``.
    Untouched members retain their compressed bytes and ZIP entry metadata.
    Use the context manager or ``close`` to release the source snapshot.
    """

    def __init__(
        self, path: str | Path | BinaryIO, *, limits: PackageLimits | None = None
    ) -> None:
        self._source_snapshot: tempfile.TemporaryDirectory | None = None
        if hasattr(path, "read"):
            self._open_stream(path, Path("document.docx").resolve(), limits)
        else:
            resolved = Path(path).resolve()
            with resolved.open("rb") as stream:
                self._open_stream(stream, resolved, limits)

    @classmethod
    def from_bytes(
        cls, data: bytes | BinaryIO, *, path: str | Path,
        limits: PackageLimits | None = None,
    ) -> Self:
        """Open package bytes or a stream; ``path`` identifies it in reports."""
        document = cls.__new__(cls)
        document._source_snapshot = None
        if hasattr(data, "read"):
            document._open_stream(data, Path(path).resolve(), limits)
        else:
            document._open(bytes(data), Path(path).resolve(), limits)
        return document

    def _open_stream(
        self, stream: BinaryIO, path: Path, limits: PackageLimits | None
    ) -> None:
        if stream.closed:
            raise PackageError("stream_closed", "Load refused: stream_closed.")
        if not stream.seekable():
            raise PackageError("stream_not_seekable", "Load refused: stream_not_seekable.")
        try:
            stream.seek(0)
        except (OSError, ValueError):
            raise PackageError("stream_read_failed", "Load refused: stream_read_failed.") from None
        self._source_snapshot = tempfile.TemporaryDirectory(prefix="wolfdocx-session-")
        snapshot = Path(self._source_snapshot.name) / "source.docx"
        bound = (limits or PackageLimits()).max_archive_size
        size = 0
        try:
            try:
                with snapshot.open("wb") as output:
                    while chunk := stream.read(64 * 1024):
                        size += len(chunk)
                        if size > bound:
                            raise PackageError("archive_size_limit", "Load refused: archive_size_limit.")
                        output.write(chunk)
            except (OSError, ValueError):
                raise PackageError("stream_read_failed", "Load refused: stream_read_failed.") from None
            self._open(snapshot, path, limits)
        except BaseException:
            self._source_snapshot.cleanup()
            self._source_snapshot = None
            raise

    def _open(self, data: bytes | Path, path: Path, limits: PackageLimits | None) -> None:
        self.path = path
        self._warnings: WarningCollection | None = None
        if isinstance(data, Path):
            with data.open("rb") as stream:
                signature = stream.read(len(_CFB_SIGNATURE))
        else:
            signature = data[:len(_CFB_SIGNATURE)]
        if signature == _CFB_SIGNATURE:
            raise PackageError("document_format_unsupported", "A DOCX ZIP package is required.")
        try:
            self._graph = _admit(data, limits or PackageLimits(), self.path, whole=False)
        except _Damage as damage:
            raise PackageError(damage.code, f"Load refused: {damage.code}.", damage.target) from None
        self.source_sha256 = self._graph.file_sha256
        if isinstance(data, Path):
            with data.open("rb") as stream:
                self._data = mmap.mmap(stream.fileno(), 0, access=mmap.ACCESS_READ)
        else:
            self._data = data
        self._source_graph = self._graph
        with self._graph._open_archive() as archive:
            self._infos = {info.filename: info for info in archive.infolist()}
            self._comment = archive.comment
        self._entries: dict[str, zipfile.ZipInfo] = dict(self._infos)
        self._changed: dict[str, bytes] = {}
        self._trees: dict[str, ET.Element] = {}
        self._checked: dict[str, tuple[bytes, ET.Element, int]] = {}
        self._edits: list[dict[str, Any]] = []

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self, exc_type: type[BaseException] | None,
        exc: BaseException | None, traceback: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        """Release package bytes, disk snapshot, and parsed trees."""
        if isinstance(self._data, mmap.mmap):
            self._data.close()
        self._data = None
        if self._source_snapshot is not None:
            self._source_snapshot.cleanup()
            self._source_snapshot = None
        self._changed.clear()
        self._trees.clear()
        self._checked.clear()

    @property
    def changed_parts(self) -> tuple[str, ...]:
        """Source parts changed and parts added by this session."""
        return tuple(sorted(self._changed))

    def part_bytes(self, name: str) -> bytes:
        """Current part bytes, including pending edits."""
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
    "_adopt", "_apply_operation", "_apply_paragraph_request", "_apply_steps",
    "_bind_imported_node", "_checked_tree", "_column", "_current_tree",
    "_definition_part", "_edit_paragraph", "_edit_part", "_editable", "_materialize",
    "_merge", "_note_parts", "_numbering", "_paragraph", "_paragraph_request",
    "_parse", "_parse_checked", "_record", "_record_save_report", "_replace_definitions",
    "_restore", "_row", "_run_edit", "_section_break", "_set_part", "_snapshot",
    "_take_editable", "_unchanged_raw", "_unchanged_raw_bytes", "_write", "_write_to",
    "accept", "append_paragraph", "apply_list", "delete_paragraph", "delete_run",
    "delete_section_break", "delete_table", "delete_table_column", "delete_table_row",
    "edit_header_footer", "find_replace", "format_paragraph", "format_text", "get_story",
    "insert_list_item", "insert_paragraph", "insert_run", "insert_section_break",
    "insert_story", "insert_table", "insert_table_column", "insert_table_row",
    "merge_table_cells", "remove_list", "restart_list", "save", "set_cell_text",
    "set_page_setup", "set_run_text", "split_paragraph", "split_table_cells",
):
    setattr(Document, _name, _EditMethod(_name))

for _name in ("background", "background_xml", "lists", "node", "part_names", "stories", "styles", "warnings"):
    setattr(Document, _name, _EditProperty(_name))
del _name


def __getattr__(name: str) -> Any:
    if name == "SessionEditReport":
        from .session_ops import SessionEditReport
        return SessionEditReport
    if name in {
        "_raw_member", "_without_zip64_extra", "_structural", "_parse_part",
        "_copy_member_raw", "_main_part", "_validate_node_structure", "_patch_node_texts",
        "_added_entry", "_PART_ERRORS", "SessionSaveReport", "_plan_paragraph_edit",
    }:
        from .session_ops import rich_objects
        return getattr(rich_objects, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
