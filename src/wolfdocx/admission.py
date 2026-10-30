"""Shared DOCX safety admission, independent of optional import and recovery policies."""

from __future__ import annotations

import hashlib
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET

from .inspect import PackageLimits, _parse_bounded_xml, _XmlSafetyError
from .package import PackageError, PackageGraph

_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
_OFFICE_DOCUMENT = f"{_R}/officeDocument"
_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_MAIN_CONTENT_TYPES = frozenset(
    {
        f"{_WML}.document.main+xml",
        f"{_WML}.template.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    }
)
_CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class _Damage(Exception):
    def __init__(self, code: str, target: str | None = None) -> None:
        super().__init__(code)
        self.code = code
        self.target = target


# -- package loading -----------------------------------------------------------------


class _DiskPackageGraph(PackageGraph):
    """Read lazy parts from the admitted snapshot, independent of the display path."""

    def _open_archive(self) -> zipfile.ZipFile:
        return self._open_zip(self._snapshot_path)

    def _mutation_source(self) -> bytes:
        from .package import PackageFormatError, PackageSafetyError

        try:
            data = self._snapshot_path.read_bytes()
        except OSError as exc:
            raise PackageFormatError(
                "package_read_failed", "Cannot read immutable package source."
            ) from exc
        if hashlib.sha256(data).hexdigest() != self.file_sha256:
            raise PackageSafetyError(
                "package_source_changed",
                "Source bytes no longer match the immutable package graph.",
            )
        return data

    @classmethod
    def from_snapshot(cls, source: Path, path: str | Path, limits: PackageLimits) -> PackageGraph:
        graph = cls.from_file(source, limits=limits)
        graph._snapshot_path = source
        graph.path = Path(path)
        return graph


def _admit(
    data: bytes | Path,
    limits: PackageLimits,
    path: str | Path,
    *,
    whole: bool = True,
    defer_parts: bool = False,
) -> PackageGraph:
    """Check member integrity and main XML; explicit recovery checks every XML part."""
    try:
        graph = (
            _DiskPackageGraph.from_snapshot(data, path, limits)
            if isinstance(data, Path)
            else PackageGraph.from_bytes(data, path=path, limits=limits)
        )
    except PackageError as exc:
        if not whole:
            raise
        raise _Damage(exc.code, exc.entry_name) from None
    try:
        with graph._open_archive() as archive:
            for info in archive.infolist():
                try:
                    with archive.open(info) as member:
                        while member.read(64 * 1024):
                            pass
                except zipfile.BadZipFile:
                    raise _Damage("member_crc_invalid", info.filename) from None
    except (zipfile.BadZipFile, zlib.error, EOFError, OSError, NotImplementedError):
        raise _Damage("member_corrupt") from None
    if whole and not defer_parts:
        for name in graph.parts:
            extension = name.rpartition(".")[2].lower()
            if (
                name not in graph.content_type_overrides
                and extension not in graph.content_type_defaults
            ):
                raise _Damage("part_content_type_missing", name)
    main = _main_part_name(graph)
    if main is None:
        raise _Damage("main_document_missing")
    if defer_parts:
        return graph
    try:
        graph.validate_relationship_graph()
    except PackageError as exc:
        if not whole:
            raise
        raise _Damage(exc.code, exc.entry_name) from None
    total = [0]
    for name, part in graph.parts.items():
        if not whole and name != main:
            continue
        if not _is_xml_part(name, part.content_type):
            continue
        try:
            root = _parse_bounded_xml(
                graph.read_part_bytes(name),
                limits=limits,
                total_elements=total,
                preserve_qnames=False,
            )
        except _XmlSafetyError as exc:
            raise _Damage(exc.code, name) from None
        except (ET.ParseError, PackageError):
            raise _Damage("malformed_xml", name) from None
        if not whole or name.endswith(".rels"):
            continue
        identifiers = {relationship.id for relationship in graph.relationships_for(name)}
        for element in root.iter():
            for key, value in element.attrib.items():
                if (
                    key.startswith((f"{{{_R}}}", f"{{{_STRICT_R}}}"))
                    and value
                    and value not in identifiers
                ):
                    raise _Damage("relationship_reference_dangling", name)
    return graph


def _main_part_name(graph: PackageGraph) -> str | None:
    for relationship in graph.relationships:
        if (
            relationship.type in {_OFFICE_DOCUMENT, f"{_STRICT_R}/officeDocument"}
            and not relationship.is_external
        ):
            target = relationship.resolved_target()
            part = graph.parts.get(target)
            if part is not None and part.content_type in _MAIN_CONTENT_TYPES:
                return target
    return None


def _is_xml_part(name: str, content_type: str | None) -> bool:
    return name.endswith((".xml", ".rels")) or (content_type or "").endswith("xml")
