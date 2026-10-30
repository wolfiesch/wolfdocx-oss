"""Immutable, reusable PackageGraph for OOXML / DOCX packages.

Opens and validates a DOCX ZIP container once against structural and safety limits,
preserves part, content-type, and relationship identity, avoids eager decompression
of media and embeddings, and provides thread-safe, immutable semantic projections.
"""

from __future__ import annotations

import copy
import hashlib
import io
import posixpath
import stat
import threading
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Callable, Iterator, Mapping, Sequence
from contextlib import AbstractContextManager, contextmanager
from dataclasses import dataclass, field, replace
from pathlib import Path
from types import MappingProxyType
from typing import Self
from urllib.parse import quote, unquote, urlsplit, urlunsplit
from xml.parsers import expat

from .canonical import semantic_sha256, sha256_bytes, sha256_file
from .inspect import (
    CONTENT_TYPES,
    PKG_REL,
    PackageLimits,
    _parse_bounded_xml,
    _safe_member_name,
    _XmlSafetyError,
    serialize_xml_with_preserved_qnames,
)
from .package_profile import PackageProfile

Q_CONTENT_TYPES_DEFAULT = f"{{{CONTENT_TYPES}}}Default"
Q_CONTENT_TYPES_OVERRIDE = f"{{{CONTENT_TYPES}}}Override"
Q_PKG_RELATIONSHIP = f"{{{PKG_REL}}}Relationship"

_WORDPROCESSINGML_FEATURE_TAGS = (
    ("sections", "sectPr"),
    ("comments", "commentReference"),
    ("tracked_insertions", "ins"),
    ("tracked_deletions", "del"),
    ("bookmarks", "bookmarkStart"),
    ("hyperlinks", "hyperlink"),
    ("content_controls", "sdt"),
    ("drawings", "drawing"),
)
_WORDPROCESSINGML_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    }
)

# ISO/IEC 29500-1: custom XML data properties (§22.5), schema library
# (§23.2), and WordprocessingML custom XML/smart tags (§17.5.1).
_CUSTOM_XML_NAMESPACES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    "http://purl.oclc.org/ooxml/officeDocument/customXml",
)
_SCHEMA_LIBRARY_NAMESPACES = (
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://purl.oclc.org/ooxml/schemaLibrary/main",
)
_SCHEMA_REFERENCE_TAGS = frozenset(
    [f"{{{ns}}}schemaRef" for ns in _CUSTOM_XML_NAMESPACES]
    + [f"{{{ns}}}schema" for ns in _SCHEMA_LIBRARY_NAMESPACES]
    + [f"{{{ns}}}attachedSchema" for ns in _WORDPROCESSINGML_NAMESPACES]
)
_ELEMENT_PROPERTY_TAGS = {
    f"{{{ns}}}{kind}": f"{{{ns}}}{kind}Pr"
    for ns in _WORDPROCESSINGML_NAMESPACES
    for kind in ("smartTag", "customXml")
}
_RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"


class PackageError(Exception):
    """Base exception for package format or safety failures."""

    def __init__(self, code: str, message: str, entry_name: str | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.entry_name = entry_name


class PackageSafetyError(PackageError):
    """Raised when package structure violates configured safety limits."""


class PackageFormatError(PackageError):
    """Raised when package is not a valid OPC/ZIP container."""


_ASCII_CASE_FOLD = str.maketrans("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")
_URI_UNRESERVED = frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")
_URI_SEGMENT_CHARACTERS = _URI_UNRESERVED | frozenset("!$&'()*+,;=:@")
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_XML = "http://www.w3.org/XML/1998/namespace"
_MC_ATTRIBUTES = frozenset(
    {"Ignorable", "MustUnderstand", "ProcessContent", "PreserveElements", "PreserveAttributes"}
)


def validate_part_names(names: Sequence[str]) -> None:
    """Refuse invalid OPC ZIP names, ASCII-equivalent names, and segment prefixes.

    The content-types stream is ZIP metadata, not an OPC part, and is the sole
    grammar exception. URI escapes must be well formed and must not encode
    unreserved characters or separators. No caller's names are normalized.
    """
    equivalent: dict[str, tuple[str, int]] = {}
    for ordinal, name in enumerate(names):
        valid = isinstance(name, str) and bool(name)
        if valid and name != "[Content_Types].xml":
            for segment in name.split("/"):
                if not segment or segment.endswith(".") or not segment.strip("."):
                    valid = False
                    break
                index = 0
                while index < len(segment):
                    character = segment[index]
                    if character == "%":
                        escape = segment[index + 1 : index + 3]
                        if len(escape) != 2 or any(
                            digit not in "0123456789abcdefABCDEF" for digit in escape
                        ):
                            valid = False
                            break
                        decoded = chr(int(escape, 16))
                        if decoded in _URI_UNRESERVED or decoded in "/\\":
                            valid = False
                            break
                        index += 3
                    elif character in _URI_SEGMENT_CHARACTERS:
                        index += 1
                    else:
                        valid = False
                        break
                if not valid:
                    break
        if not valid:
            raise PackageFormatError(
                "part_name_invalid", "Package part name violates OPC URI grammar.", name
            )
        key = name.translate(_ASCII_CASE_FOLD)
        if key in equivalent:
            raise PackageFormatError(
                "part_name_equivalent",
                f"Package part name is ASCII-equivalent to {equivalent[key][0]!r}.",
                name,
            )
        equivalent[key] = (name, ordinal)
    # Walk segment ancestors, not string prefixes (a.xml and a.xml2 are distinct).
    for key, (name, ordinal) in equivalent.items():
        prefix = key
        while "/" in prefix:
            prefix = prefix.rpartition("/")[0]
            if prefix in equivalent:
                ancestor, ancestor_ordinal = equivalent[prefix]
                raise PackageFormatError(
                    "part_name_prefix_collision",
                    f"Package parts {name!r} and {ancestor!r} have a segment-prefix collision.",
                    name if ordinal > ancestor_ordinal else ancestor,
                )


def _package_entry_is_part(entry: zipfile.ZipInfo) -> bool:
    """Validate ZIP metadata before distinguishing OPC parts from preserved baggage."""
    name = entry.filename
    if not _safe_member_name(name):
        raise PackageSafetyError("unsafe_path", "Package entry uses an unsafe path.", name)
    kind = stat.S_IFMT(entry.external_attr >> 16)
    directory = name.endswith("/")
    if (directory and (entry.file_size != 0 or kind not in (0, stat.S_IFDIR))) or (
        not directory and kind not in (0, stat.S_IFREG)
    ):
        raise PackageSafetyError(
            "directory_entry", "Package contains non-regular directory entry.", name
        )
    # Only the legacy top-level trash segment is exceptional. Its suffix still
    # obeys OPC grammar, and all ZIP byte/count/ratio limits apply to its entries.
    legacy_trash = name.startswith("[trash]/")
    candidate = name[len("[trash]/") :] if legacy_trash else name
    if directory:
        candidate = candidate[:-1]
    if (directory or legacy_trash) and (candidate or not legacy_trash):
        validate_part_names((candidate,))
    return not directory and not legacy_trash


def _validate_content_type_part_names(root: ET.Element) -> None:
    for child in root:
        if child.tag == Q_CONTENT_TYPES_OVERRIDE:
            validate_part_names((child.get("PartName", "").lstrip("/"),))


def _validate_legacy_trash_target(source_part: str, target: str) -> None:
    try:
        uri = urlsplit(target)
    except ValueError:
        return  # Retain the existing malformed-target diagnostic path.
    if uri.scheme or uri.netloc:
        return
    path = unquote(uri.path)
    resolved = posixpath.normpath(
        path.lstrip("/")
        if path.startswith("/")
        else posixpath.join(posixpath.dirname(source_part), path)
    )
    if resolved.translate(_ASCII_CASE_FOLD).startswith("[trash]/"):
        validate_part_names((resolved,))


def is_relationships_part(name: str) -> bool:
    """Recognize the reserved OPC relationships-part layout without case aliases."""
    folded = name.translate(_ASCII_CASE_FOLD)
    directory, basename = posixpath.split(folded)
    return (
        posixpath.basename(directory) == "_rels"
        and basename.endswith(".rels")
        and (basename != ".rels" or directory == "_rels")
    )


def validate_relationships_part(root: ET.Element, rels_name: str) -> None:
    """Validate relationships markup after scope-aware MCE processing.

    Ignorable namespaces are inherited as expanded namespace identities, not
    lexical prefixes. Parsed elements retain inspect's lexical QName scopes.
    Ignorable extension subtrees/attributes remain opaque and are never removed,
    so subsequent package edits preserve their declarations and payloads.
    """
    if root.tag != f"{{{PKG_REL}}}Relationships":
        raise PackageFormatError(
            "relationships_root_invalid", "Unexpected relationships-part root.", rels_name
        )
    understood = {PKG_REL, _MC, _XML}
    pending = [(root, frozenset())]
    while pending:
        element, inherited = pending.pop()
        scope = dict(getattr(element, "_namespace_scope", ()))
        # Plain authored elements may declare prefixes explicitly.
        for attribute, value in element.attrib.items():
            if attribute.startswith("xmlns:"):
                scope[attribute[6:]] = value
        ignorable = set(inherited)
        for local in ("Ignorable", "MustUnderstand"):
            for prefix in element.get(f"{{{_MC}}}{local}", "").split():
                namespace = scope.get(prefix)
                if not namespace:
                    raise PackageFormatError(
                        "relationships_mce_prefix_invalid",
                        "Relationships MCE attribute names an unbound prefix.",
                        rels_name,
                    )
                if local == "Ignorable":
                    ignorable.add(namespace)
                elif namespace not in understood:
                    raise PackageFormatError(
                        "relationships_must_understand",
                        "Relationships part requires an unsupported namespace.",
                        rels_name,
                    )
        allowed = (
            {"Id", "Type", "Target", "TargetMode"} if element.tag == Q_PKG_RELATIONSHIP else set()
        )
        for attribute in element.attrib:
            if attribute == "xmlns" or attribute.startswith("xmlns:"):
                continue
            if attribute in allowed:
                continue
            if attribute.startswith("{"):
                namespace, local = attribute[1:].split("}", 1)
                if namespace == _MC and local in _MC_ATTRIBUTES:
                    continue
                if namespace not in understood and namespace in ignorable:
                    continue
            raise PackageFormatError(
                "relationships_attribute_invalid",
                "Relationships part contains a non-ignorable or unknown attribute.",
                rels_name,
            )
        if element.text and element.text.strip():
            raise PackageFormatError(
                "relationships_content_invalid",
                "Relationships part contains unexpected text.",
                rels_name,
            )
        for child in reversed(element):
            namespace = child.tag[1:].split("}", 1)[0] if child.tag.startswith("{") else ""
            if namespace not in understood and namespace in ignorable:
                continue
            if element is not root or child.tag != Q_PKG_RELATIONSHIP:
                raise PackageFormatError(
                    "relationships_content_invalid",
                    "Relationships part contains a non-ignorable or unknown element.",
                    rels_name,
                )
            pending.append((child, frozenset(ignorable)))


@dataclass(slots=True)
class PackageCounters:
    """Per-instance instrumentation counters for verification and profiling."""

    zip_open_count: int = 0
    part_bytes_read_count: int = 0
    part_bytes_streamed_count: int = 0
    xml_parse_count: int = 0
    xml_stream_count: int = 0
    total_bytes_decompressed: int = 0
    total_bytes_streamed: int = 0
    current_materialized_bytes: int = 0
    peak_materialized_bytes: int = 0
    total_xml_elements_parsed: int = 0
    relationship_type_index_build_count: int = 0
    relationship_target_index_build_count: int = 0
    parts_by_content_type_index_build_count: int = 0
    wordprocessingml_feature_counts_build_count: int = 0
    canonical_xml_digest_build_count: int = 0

    def snapshot(self) -> dict[str, int]:
        return {
            "zip_open_count": self.zip_open_count,
            "part_bytes_read_count": self.part_bytes_read_count,
            "part_bytes_streamed_count": self.part_bytes_streamed_count,
            "xml_parse_count": self.xml_parse_count,
            "xml_stream_count": self.xml_stream_count,
            "total_bytes_decompressed": self.total_bytes_decompressed,
            "total_bytes_streamed": self.total_bytes_streamed,
            "current_materialized_bytes": self.current_materialized_bytes,
            "peak_materialized_bytes": self.peak_materialized_bytes,
            "total_xml_elements_parsed": self.total_xml_elements_parsed,
            "relationship_type_index_build_count": self.relationship_type_index_build_count,
            "relationship_target_index_build_count": self.relationship_target_index_build_count,
            "parts_by_content_type_index_build_count": self.parts_by_content_type_index_build_count,
            "wordprocessingml_feature_counts_build_count": (
                self.wordprocessingml_feature_counts_build_count
            ),
            "canonical_xml_digest_build_count": self.canonical_xml_digest_build_count,
        }


@dataclass(frozen=True, slots=True)
class WordprocessingMlFeatureCounts:
    """Compact immutable counts for the modeled WordprocessingML feature subset."""

    sections: int = 0
    comments: int = 0
    tracked_insertions: int = 0
    tracked_deletions: int = 0
    bookmarks: int = 0
    fields: int = 0
    hyperlinks: int = 0
    content_controls: int = 0
    drawings: int = 0
    headers: int = 0
    footers: int = 0

    def to_dict(self) -> dict[str, int]:
        """Return field names compatible with inspection feature-count semantics."""
        return {
            "sections": self.sections,
            "comments": self.comments,
            "tracked_insertions": self.tracked_insertions,
            "tracked_deletions": self.tracked_deletions,
            "bookmarks": self.bookmarks,
            "fields": self.fields,
            "hyperlinks": self.hyperlinks,
            "content_controls": self.content_controls,
            "drawings": self.drawings,
            "headers": self.headers,
            "footers": self.footers,
        }


@dataclass(frozen=True, slots=True)
class PackageXmlEvent:
    """One immutable event from a bounded, non-tree XML traversal."""

    kind: str
    tag: str = ""
    attrib: Mapping[str, str] = field(default_factory=lambda: MappingProxyType({}))
    text: str | None = None


@dataclass(frozen=True, slots=True)
class PackageRelationship:
    """Immutable representation of an OPC relationship."""

    id: str
    type: str
    target: str
    target_mode: str = "Internal"
    source_part: str = ""

    @property
    def is_external(self) -> bool:
        return self.target_mode == "External"

    def resolved_target(self) -> str:
        """Resolve internal target path relative to source part."""
        if self.is_external:
            return self.target
        if self.target.startswith("/"):
            return posixpath.normpath(self.target.lstrip("/"))
        source_dir = posixpath.dirname(self.source_part)
        return posixpath.normpath(posixpath.join(source_dir, self.target))


@dataclass(frozen=True, slots=True)
class PackageXmlElement:
    """Immutable semantic projection of an XML element.

    Exposes read-only access to tag, attributes, text, tail, and children,
    preventing mutable element tree poisoning across shared consumers.
    """

    tag: str
    attrib: Mapping[str, str] = field(default_factory=lambda: MappingProxyType({}))
    text: str | None = None
    tail: str | None = None
    children: tuple[PackageXmlElement, ...] = ()

    @classmethod
    def from_etree(cls, element: ET.Element) -> PackageXmlElement:
        """Construct an immutable element tree from a standard Element."""
        attrib = MappingProxyType(dict(element.attrib))
        children = tuple(cls.from_etree(child) for child in element)
        return cls(
            tag=element.tag,
            attrib=attrib,
            text=element.text,
            tail=element.tail,
            children=children,
        )

    def get(self, key: str, default: str | None = None) -> str | None:
        return self.attrib.get(key, default)

    def find(self, tag: str) -> PackageXmlElement | None:
        """Find a direct child or slash-delimited direct-child path."""
        segments: list[str] = []
        start = 0
        namespace_depth = 0
        for index, character in enumerate(tag):
            if character == "{":
                namespace_depth += 1
            elif character == "}":
                namespace_depth -= 1
            elif character == "/" and namespace_depth == 0:
                segments.append(tag[start:index])
                start = index + 1
        segments.append(tag[start:])

        current: PackageXmlElement | None = self
        for segment in segments:
            if current is None:
                return None
            current = next(
                (child for child in current.children if child.tag == segment),
                None,
            )
        return current

    def findall(self, tag: str) -> list[PackageXmlElement]:
        return [child for child in self.children if child.tag == tag]

    def iter(self, tag: str | None = None) -> Iterator[PackageXmlElement]:
        if tag is None or self.tag == tag:
            yield self
        for child in self.children:
            yield from child.iter(tag)

    def __iter__(self) -> Iterator[PackageXmlElement]:
        return iter(self.children)

    def __len__(self) -> int:
        return len(self.children)

    def to_mutable_element(self) -> ET.Element:
        """Return an isolated, defensive copy as a mutable standard Element."""
        element = ET.Element(self.tag, attrib=dict(self.attrib))
        element.text = self.text
        element.tail = self.tail
        for child in self.children:
            element.append(child.to_mutable_element())
        return element


@dataclass(frozen=True, slots=True)
class PackageSchemaReference:
    """One existing ds:schemaRef, sl:schema, or w:attachedSchema.

    ``element_path`` is a tuple of zero-based child indexes from the XML root;
    ``()`` identifies the root. ``element`` retains actual expanded names and
    attributes, including unknown extensions. Listing neither fetches schemas nor
    infers a schema location from a namespace URI.
    """

    part_name: str
    element_path: tuple[int, ...]
    element: PackageXmlElement


@dataclass(frozen=True, slots=True)
class PackageElementProperties:
    """Explicit identity and properties of one w:smartTag or w:customXml.

    The full immutable owner attributes and optional properties subtree are
    exposed, rather than inventing a flattened schema for their extension data.
    Child indexes include all XML children, as for PackageSchemaReference.
    """

    part_name: str
    element_path: tuple[int, ...]
    owner_tag: str
    owner_attributes: Mapping[str, str]
    properties: PackageXmlElement | None


class _PackageArchiveSession(AbstractContextManager["_PackageArchiveSession"]):
    """One scoped reuse of a graph's ZIP handle."""

    __slots__ = ("_closed", "_entered", "_graph")

    def __init__(self, graph: PackageGraph) -> None:
        self._graph = graph
        self._entered = False
        self._closed = False

    def __enter__(self) -> Self:
        if self._closed:
            raise RuntimeError("package read session is closed")
        if not self._entered:
            self._graph._enter_archive_session()
            self._entered = True
        return self

    def close(self) -> None:
        if self._entered and not self._closed:
            self._graph._exit_archive_session()
            self._closed = True

    def __exit__(self, *args: object) -> None:
        self.close()


class PackagePart:
    """Immutable descriptor of an OPC package part with thread-safe lazy loading."""

    __slots__ = (
        "_cached_bytes",
        "_cached_sha256",
        "_cached_xml",
        "_graph_ref",
        "_lock",
        "compressed_size",
        "content_type",
        "name",
        "relationships",
        "size",
    )

    def __init__(
        self,
        *,
        name: str,
        content_type: str,
        size: int,
        compressed_size: int,
        relationships: Sequence[PackageRelationship] = (),
        graph: PackageGraph,
        initial_bytes: bytes | None = None,
        initial_sha256: str | None = None,
        initial_xml: PackageXmlElement | None = None,
    ) -> None:
        if initial_bytes is not None and initial_sha256 is not None:
            raise ValueError("initial bytes and SHA-256 are mutually exclusive")
        self.name = name
        self.content_type = content_type
        self.size = size
        self.compressed_size = compressed_size
        self.relationships = tuple(relationships)
        self._graph_ref = graph
        self._lock = threading.RLock()
        self._cached_bytes: bytes | None = initial_bytes
        self._cached_sha256: str | None = (
            hashlib.sha256(initial_bytes).hexdigest()
            if initial_bytes is not None
            else initial_sha256
        )
        self._cached_xml = initial_xml

    def __repr__(self) -> str:
        return f"<PackagePart name={self.name!r} type={self.content_type!r} size={self.size}>"

    @property
    def sha256(self) -> str:
        """Return SHA256 without retaining bytes when only identity is requested."""
        if self._cached_sha256 is None:
            with self._lock:
                if self._cached_sha256 is None:
                    self._cached_sha256 = self._graph_ref._stream_member_sha256(
                        self.name,
                        self.size,
                    )
        return self._cached_sha256

    def read_bytes(self) -> bytes:
        """Lazily read and memoize part bytes with thread-safe synchronization."""
        if self._cached_bytes is None:
            with self._lock:
                if self._cached_bytes is None:
                    raw = self._graph_ref._read_member_bytes(self.name, self.size)
                    self._cached_bytes = raw
                    self._cached_sha256 = hashlib.sha256(raw).hexdigest()
        return self._cached_bytes

    def iter_xml_events(self) -> Iterator[PackageXmlEvent]:
        """Yield immutable events without constructing an XML element tree."""
        return self._graph_ref._iter_member_xml_events(self.name, self.size)

    def parsed_xml(self) -> PackageXmlElement:
        """Lazily parse and memoize an immutable semantic XML projection."""
        if self._cached_xml is None:
            with self._lock:
                if self._cached_xml is None:
                    transient_bytes = (
                        self._cached_bytes is None and self._graph_ref._is_archive_session_active()
                    )
                    data = (
                        self._graph_ref._read_member_bytes(self.name, self.size)
                        if transient_bytes
                        else self.read_bytes()
                    )
                    try:
                        raw_root = self._graph_ref._parse_member_xml(self.name, data)
                        projection = PackageXmlElement.from_etree(raw_root)
                    finally:
                        if transient_bytes:
                            self._graph_ref._release_materialization(self.size)
                    with self._graph_ref._lock:
                        self._graph_ref.counters.xml_parse_count += 1
                    self._cached_xml = projection
        return self._cached_xml


class PackageGraph:
    """Thread-safe, immutable representation of an entire OOXML / DOCX container."""

    def __init__(
        self,
        *,
        path: Path,
        file_sha256: str,
        limits: PackageLimits,
        parts: Mapping[str, PackagePart],
        relationships: Sequence[PackageRelationship],
        content_type_defaults: Mapping[str, str],
        content_type_overrides: Mapping[str, str],
        entry_count: int,
        total_uncompressed_size: int,
        initial_materialized_bytes: int,
        materialization_limit: int,
        xml_element_counts: Mapping[str, int],
        counters: PackageCounters,
    ) -> None:
        self.path = path
        self.file_sha256 = file_sha256
        self.limits = limits
        self.parts = MappingProxyType(dict(parts))
        self.relationships = tuple(relationships)
        self.content_type_defaults = MappingProxyType(dict(content_type_defaults))
        self.content_type_overrides = MappingProxyType(dict(content_type_overrides))
        self.entry_count = entry_count
        self.total_uncompressed_size = total_uncompressed_size
        self._materialization_limit = materialization_limit
        self.counters = counters
        self.counters.current_materialized_bytes = initial_materialized_bytes
        self.counters.peak_materialized_bytes = initial_materialized_bytes
        self.counters.total_xml_elements_parsed = sum(xml_element_counts.values())
        self._lock = threading.RLock()
        self._archive: zipfile.ZipFile | None = None
        self._archive_session_depth = 0
        self._total_xml_elements = [self.counters.total_xml_elements_parsed]
        self._xml_element_counts = dict(xml_element_counts)
        self._relationships_by_target: Mapping[str, tuple[PackageRelationship, ...]] | None = None
        self._relationships_by_type: Mapping[str, tuple[PackageRelationship, ...]] | None = None
        self._parts_by_content_type: Mapping[str, tuple[PackagePart, ...]] | None = None
        self._wordprocessingml_feature_counts: WordprocessingMlFeatureCounts | None = None
        self._canonical_xml_digests: dict[str, str] | None = None
        # Set only by ``from_bytes``: members are then read from these exact bytes,
        # never from ``path``, so ``file_sha256`` always names what is read.
        self._data: bytes | None = None

    def __repr__(self) -> str:
        return (
            f"<PackageGraph path={self.path!r} parts={len(self.parts)} "
            f"size={self.total_uncompressed_size}>"
        )

    @classmethod
    def from_file(
        cls,
        path: str | Path,
        *,
        limits: PackageLimits | None = None,
        profile: PackageProfile | None = None,
    ) -> PackageGraph:
        """Open and validate a DOCX package once, constructing an immutable PackageGraph."""
        doc_path = Path(path).resolve()
        limits, materialization_limit = cls._effective_limits(limits, profile)

        if not doc_path.is_file():
            raise PackageError("file_not_found", f"File not found: {doc_path}")

        file_size = doc_path.stat().st_size
        if file_size > limits.max_archive_size:
            raise PackageSafetyError(
                "archive_size_limit", "Archive size exceeds configured safety limit."
            )

        if not zipfile.is_zipfile(doc_path):
            raise PackageFormatError("not_zip", "File is not a valid ZIP/OPC archive.")

        return cls._build(
            lambda: cls._open_zip(doc_path),
            path=doc_path,
            file_sha256=sha256_file(doc_path),
            limits=limits,
            materialization_limit=materialization_limit,
        )

    @classmethod
    def from_bytes(
        cls,
        data: bytes,
        *,
        path: str | Path,
        limits: PackageLimits | None = None,
        profile: PackageProfile | None = None,
        file_sha256: str | None = None,
    ) -> PackageGraph:
        """Validate one in-memory package exactly as ``from_file`` validates a file.

        Every member is read from ``data``; ``path`` only names the package.
        ``file_sha256`` may pass the digest a caller already computed for ``data``.
        """
        limits, materialization_limit = cls._effective_limits(limits, profile)
        if len(data) > limits.max_archive_size:
            raise PackageSafetyError(
                "archive_size_limit", "Archive size exceeds configured safety limit."
            )
        if not zipfile.is_zipfile(io.BytesIO(data)):
            raise PackageFormatError("not_zip", "File is not a valid ZIP/OPC archive.")
        graph = cls._build(
            lambda: zipfile.ZipFile(io.BytesIO(data), "r"),
            path=Path(path),
            file_sha256=file_sha256 or sha256_bytes(data),
            limits=limits,
            materialization_limit=materialization_limit,
        )
        graph._data = data
        return graph

    @classmethod
    def _effective_limits(
        cls, limits: PackageLimits | None, profile: PackageProfile | None
    ) -> tuple[PackageLimits, int]:
        limits = limits or PackageLimits()
        materialization_limit = limits.max_uncompressed_size
        if profile is not None:
            if not isinstance(profile, PackageProfile):
                raise TypeError("profile must be a PackageProfile")
            materialization_limit = profile.memory.max_materialized_bytes
            limits = cls._limits_for_profile(limits, profile)
        return limits, materialization_limit

    @classmethod
    def _build(
        cls,
        open_archive: Callable[[], zipfile.ZipFile],
        *,
        path: Path,
        file_sha256: str,
        limits: PackageLimits,
        materialization_limit: int,
    ) -> PackageGraph:
        counters = PackageCounters()
        with open_archive() as archive:
            counters.zip_open_count += 1
            entries = archive.infolist()

            if len(entries) > limits.max_entries:
                raise PackageSafetyError(
                    "entry_limit",
                    f"Package has {len(entries)} entries; limit is {limits.max_entries}.",
                )

            names = [entry.filename for entry in entries]
            if len(names) != len(set(names)):
                raise PackageFormatError(
                    "duplicate_entry", "Package contains duplicate ZIP entries."
                )

            total_uncompressed = sum(entry.file_size for entry in entries)
            if total_uncompressed > limits.max_uncompressed_size:
                raise PackageSafetyError(
                    "package_size_limit",
                    "Package uncompressed size exceeds safety limit.",
                )

            # Validate safety of all entry headers
            entry_map: dict[str, zipfile.ZipInfo] = {}
            for entry in entries:
                is_part = _package_entry_is_part(entry)
                if entry.file_size > limits.max_entry_size:
                    raise PackageSafetyError(
                        "entry_size_limit",
                        "Package entry exceeds size limit.",
                        entry.filename,
                    )
                ratio = entry.file_size / max(entry.compress_size, 1)
                if ratio > limits.max_compression_ratio:
                    raise PackageSafetyError(
                        "compression_ratio_limit",
                        "Compression ratio exceeds safety limit.",
                        entry.filename,
                    )
                if is_part:
                    entry_map[entry.filename] = entry
            validate_part_names(tuple(entry_map))
            validate_part_names(
                tuple(
                    name[len("[trash]/") :]
                    for name in names
                    if name.startswith("[trash]/") and not name.endswith("/")
                )
            )

            if "[Content_Types].xml" not in entry_map:
                raise PackageFormatError(
                    "missing_required_part",
                    "Missing required [Content_Types].xml part.",
                )

            # Read [Content_Types].xml and .rels files to establish package topology
            ct_entry = entry_map["[Content_Types].xml"]
            ct_data = archive.read(ct_entry.filename)
            counters.part_bytes_read_count += 1
            counters.total_bytes_decompressed += len(ct_data)

            total_xml_elements = [0]
            xml_element_counts: dict[str, int] = {}
            initial_xml: dict[str, PackageXmlElement] = {}
            try:
                ct_root = _parse_bounded_xml(
                    ct_data, limits=limits, total_elements=total_xml_elements
                )
            except _XmlSafetyError as exc:
                raise PackageSafetyError(exc.code, exc.message, "[Content_Types].xml") from exc
            except ET.ParseError as exc:
                raise PackageFormatError(
                    "malformed_xml",
                    f"Malformed [Content_Types].xml: {exc}",
                    "[Content_Types].xml",
                ) from exc
            counters.xml_parse_count += 1
            xml_element_counts["[Content_Types].xml"] = total_xml_elements[0]
            initial_xml["[Content_Types].xml"] = PackageXmlElement.from_etree(ct_root)

            _validate_content_type_part_names(ct_root)
            defaults, overrides = cls._parse_content_types(ct_root)

            # Read and parse relationship parts (.rels) to establish package topology.
            rels_by_source: dict[str, list[PackageRelationship]] = {}
            root_relationships: list[PackageRelationship] = []

            for name in entry_map:
                if name.endswith(".rels"):
                    rel_data = archive.read(name)
                    counters.part_bytes_read_count += 1
                    counters.total_bytes_decompressed += len(rel_data)
                    try:
                        rel_root = _parse_bounded_xml(
                            rel_data, limits=limits, total_elements=total_xml_elements
                        )
                    except _XmlSafetyError as exc:
                        raise PackageSafetyError(exc.code, exc.message, name) from exc
                    except ET.ParseError as exc:
                        raise PackageFormatError(
                            "malformed_xml",
                            f"Malformed relationships {name}: {exc}",
                            name,
                        ) from exc
                    counters.xml_parse_count += 1
                    xml_element_counts[name] = total_xml_elements[0] - sum(
                        xml_element_counts.values()
                    )
                    initial_xml[name] = PackageXmlElement.from_etree(rel_root)

                    source_part = cls._source_part_for_rels(name)
                    rels_list = cls._parse_relationships(rel_root, source_part=source_part)
                    if source_part == "":
                        root_relationships = rels_list
                    else:
                        rels_by_source[source_part] = rels_list

        # Construct PackageGraph and PackagePart descriptors
        graph = cls(
            path=path,
            file_sha256=file_sha256,
            limits=limits,
            parts={},
            relationships=root_relationships,
            content_type_defaults=defaults,
            content_type_overrides=overrides,
            entry_count=len(entries),
            total_uncompressed_size=total_uncompressed,
            initial_materialized_bytes=0,
            materialization_limit=materialization_limit,
            xml_element_counts=xml_element_counts,
            counters=counters,
        )

        parts: dict[str, PackagePart] = {}
        for name, entry in entry_map.items():
            ctype = cls._resolve_content_type(name, defaults, overrides)
            initial_sha256 = (
                hashlib.sha256(ct_data).hexdigest() if name == "[Content_Types].xml" else None
            )
            part_rels = rels_by_source.get(name, ())
            parts[name] = PackagePart(
                name=name,
                content_type=ctype,
                size=entry.file_size,
                compressed_size=entry.compress_size,
                relationships=part_rels,
                graph=graph,
                initial_sha256=initial_sha256,
                initial_xml=initial_xml.get(name),
            )

        graph.parts = MappingProxyType(parts)
        return graph

    @staticmethod
    def _limits_for_profile(
        limits: PackageLimits,
        profile: PackageProfile,
    ) -> PackageLimits:
        """Constrain individual materialization boundaries by a package profile."""
        bound = profile.memory.max_materialized_bytes
        return replace(
            limits,
            max_entry_size=min(limits.max_entry_size, bound),
            max_xml_part_size=min(limits.max_xml_part_size, bound),
        )

    @staticmethod
    def _open_zip(path: Path) -> zipfile.ZipFile:
        return zipfile.ZipFile(path, "r")

    def _open_archive(self) -> zipfile.ZipFile:
        if self._data is not None:
            return zipfile.ZipFile(io.BytesIO(self._data), "r")
        return self._open_zip(self.path)

    def archive_session(self) -> _PackageArchiveSession:
        """Scope one reusable ZIP archive across a higher-level graph operation."""
        return _PackageArchiveSession(self)

    def _enter_archive_session(self) -> None:
        with self._lock:
            self._archive_session_depth += 1

    def _exit_archive_session(self) -> None:
        with self._lock:
            if self._archive_session_depth <= 0:
                raise RuntimeError("package archive session depth underflow")
            self._archive_session_depth -= 1
            if self._archive_session_depth == 0:
                archive = self._archive
                self._archive = None
                if archive is not None:
                    archive.close()

    def _is_archive_session_active(self) -> bool:
        with self._lock:
            return self._archive_session_depth > 0

    @contextmanager
    def _archive_for_read(self) -> Iterator[zipfile.ZipFile]:
        """Serialize access to the session archive or retain legacy direct behavior."""
        with self._lock:
            if self._archive is not None:
                yield self._archive
                return
            if self._archive_session_depth:
                self._archive = self._open_archive()
                self.counters.zip_open_count += 1
                yield self._archive
                return
            with self._open_archive() as archive:
                self.counters.zip_open_count += 1
                yield archive

    def _reserve_materialization(self, expected_size: int, member_name: str) -> None:
        """Account for retained bytes before decompression, failing closed at the bound."""
        if expected_size > self._materialization_limit:
            raise PackageSafetyError(
                "materialization_limit",
                "Part would exceed the active materialization bound.",
                member_name,
            )
        with self._lock:
            current = self.counters.current_materialized_bytes
            if current + expected_size > self._materialization_limit:
                raise PackageSafetyError(
                    "materialization_limit",
                    "Materializing this part would exceed the active aggregate bound.",
                    member_name,
                )
            self.counters.current_materialized_bytes = current + expected_size
            self.counters.peak_materialized_bytes = max(
                self.counters.peak_materialized_bytes,
                self.counters.current_materialized_bytes,
            )

    def _release_materialization(self, expected_size: int) -> None:
        with self._lock:
            self.counters.current_materialized_bytes -= expected_size

    def _read_member_bytes(self, member_name: str, expected_size: int) -> bytes:
        """Internal helper to decompress and retain a single member on demand."""
        self._reserve_materialization(expected_size, member_name)
        try:
            with self._archive_for_read() as archive:
                data = archive.read(member_name)
        except Exception:
            with self._lock:
                self.counters.current_materialized_bytes -= expected_size
            raise
        if len(data) != expected_size:
            with self._lock:
                self.counters.current_materialized_bytes -= expected_size
            raise PackageFormatError(
                "entry_size_mismatch",
                f"Part size changed during read: expected {expected_size}, got {len(data)}.",
                member_name,
            )
        with self._lock:
            self.counters.part_bytes_read_count += 1
            self.counters.total_bytes_decompressed += len(data)
        return data

    def _stream_member_sha256(self, member_name: str, expected_size: int) -> str:
        """Hash one part incrementally without retaining its decompressed bytes."""
        digest = hashlib.sha256()
        total = 0
        with self._archive_for_read() as archive, archive.open(member_name) as member:
            for chunk in iter(lambda: member.read(64 * 1024), b""):
                digest.update(chunk)
                total += len(chunk)
        if total != expected_size:
            raise PackageFormatError(
                "entry_size_mismatch",
                f"Part size changed during stream: expected {expected_size}, got {total}.",
                member_name,
            )
        with self._lock:
            self.counters.part_bytes_streamed_count += 1
            self.counters.total_bytes_streamed += total
            self.counters.total_bytes_decompressed += total
        return digest.hexdigest()

    def _record_xml_element_count(self, member_name: str, element_count: int) -> None:
        with self._lock:
            previous = self._xml_element_counts.get(member_name, 0)
            self._xml_element_counts[member_name] = element_count
            self._total_xml_elements[0] += element_count - previous
            self.counters.total_xml_elements_parsed = self._total_xml_elements[0]

    def _parse_member_xml(self, member_name: str, data: bytes) -> ET.Element:
        with self._lock:
            previous = self._xml_element_counts.get(member_name, 0)
            total_elements = [self._total_xml_elements[0] - previous]
            try:
                root = _parse_bounded_xml(
                    data,
                    limits=self.limits,
                    total_elements=total_elements,
                )
            except _XmlSafetyError as exc:
                raise PackageSafetyError(exc.code, exc.message, member_name) from exc
            except ET.ParseError as exc:
                raise PackageFormatError(
                    "malformed_xml",
                    f"XML parse error: {exc}",
                    member_name,
                ) from exc
            self._record_xml_element_count(
                member_name,
                total_elements[0] - (self._total_xml_elements[0] - previous),
            )
            return root

    def _iter_member_xml_events(
        self,
        member_name: str,
        expected_size: int,
    ) -> Iterator[PackageXmlEvent]:
        """Stream scalar XML events with the same XML safety limits as tree parsing."""
        if expected_size > self.limits.max_xml_part_size:
            raise PackageSafetyError(
                "xml_size_limit",
                "XML part exceeds the configured safety limit.",
                member_name,
            )

        events: list[PackageXmlEvent] = []
        elements = 0
        depth = 0
        previous = self._xml_element_counts.get(member_name, 0)
        base_total = self._total_xml_elements[0] - previous
        parser = expat.ParserCreate(namespace_separator="}")

        def qname(name: str) -> str:
            namespace, separator, local_name = name.rpartition("}")
            return f"{{{namespace}}}{local_name}" if separator else name

        def start(name: str, attrs: dict[str, str]) -> None:
            nonlocal depth, elements
            depth += 1
            if depth > self.limits.max_xml_depth:
                raise PackageSafetyError(
                    "xml_depth_limit",
                    "XML part exceeds the configured nesting-depth limit.",
                    member_name,
                )
            elements += 1
            if elements > self.limits.max_xml_elements:
                raise PackageSafetyError(
                    "xml_element_limit",
                    "XML part exceeds the configured element limit.",
                    member_name,
                )
            if base_total + elements > self.limits.max_total_xml_elements:
                raise PackageSafetyError(
                    "xml_total_element_limit",
                    "Package XML exceeds configured aggregate element limit.",
                    member_name,
                )
            events.append(
                PackageXmlEvent(
                    kind="start",
                    tag=qname(name),
                    attrib=MappingProxyType({qname(key): value for key, value in attrs.items()}),
                )
            )

        def end(name: str) -> None:
            nonlocal depth
            events.append(PackageXmlEvent(kind="end", tag=qname(name)))
            depth -= 1

        def text(value: str) -> None:
            if value:
                events.append(PackageXmlEvent(kind="text", text=value))

        def forbidden(*_args: object) -> None:
            raise PackageSafetyError(
                "xml_dtd_forbidden",
                "XML part contains a forbidden DTD or entity declaration.",
                member_name,
            )

        parser.StartElementHandler = start
        parser.EndElementHandler = end
        parser.CharacterDataHandler = text
        parser.StartDoctypeDeclHandler = forbidden
        parser.EntityDeclHandler = forbidden

        total = 0
        try:
            with (
                self._archive_for_read() as archive,
                archive.open(member_name) as member,
            ):
                for chunk in iter(lambda: member.read(64 * 1024), b""):
                    total += len(chunk)
                    parser.Parse(chunk, False)
                    yield from events
                    events.clear()
                parser.Parse(b"", True)
        except PackageSafetyError:
            raise
        except expat.ExpatError as exc:
            raise PackageFormatError(
                "malformed_xml",
                f"XML parse error: {exc}",
                member_name,
            ) from exc
        if total != expected_size:
            raise PackageFormatError(
                "entry_size_mismatch",
                f"Part size changed during stream: expected {expected_size}, got {total}.",
                member_name,
            )
        yield from events
        self._record_xml_element_count(member_name, elements)
        with self._lock:
            self.counters.xml_stream_count += 1
            self.counters.part_bytes_streamed_count += 1
            self.counters.total_bytes_streamed += total
            self.counters.total_bytes_decompressed += total

    @staticmethod
    def _source_part_for_rels(rels_name: str) -> str:
        if rels_name == "_rels/.rels":
            return ""
        if rels_name.startswith("_rels/") and rels_name.endswith(".rels"):
            return rels_name[len("_rels/") : -len(".rels")]
        marker = "/_rels/"
        if marker not in rels_name or not rels_name.endswith(".rels"):
            return ""
        prefix, rel_file = rels_name.split(marker, 1)
        return posixpath.join(prefix, rel_file[: -len(".rels")])

    @staticmethod
    def _parse_content_types(root: ET.Element) -> tuple[dict[str, str], dict[str, str]]:
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for child in root:
            if child.tag == Q_CONTENT_TYPES_DEFAULT:
                ext = child.get("Extension")
                ct = child.get("ContentType")
                if ext and ct:
                    defaults[ext.lower()] = ct
            elif child.tag == Q_CONTENT_TYPES_OVERRIDE:
                part = child.get("PartName")
                ct = child.get("ContentType")
                if part and ct:
                    overrides[part.lstrip("/")] = ct
        return defaults, overrides

    @staticmethod
    def _resolve_content_type(
        name: str,
        defaults: Mapping[str, str],
        overrides: Mapping[str, str],
    ) -> str:
        if name in overrides:
            return overrides[name]
        extension = name.rpartition(".")[2].lower()
        return defaults.get(extension, "application/octet-stream")

    @staticmethod
    def _parse_relationships(root: ET.Element, source_part: str) -> list[PackageRelationship]:
        rels_name = PackageGraph._rels_part_name(source_part)
        validate_relationships_part(root, rels_name)
        relationships: list[PackageRelationship] = []
        seen_ids: set[str] = set()
        for elem in root.findall(Q_PKG_RELATIONSHIP):
            rel_id = elem.get("Id", "")
            rel_type = elem.get("Type", "")
            target = elem.get("Target", "")
            mode = elem.get("TargetMode", "Internal")
            if rel_id in seen_ids:
                raise PackageFormatError(
                    "duplicate_relationship_id", f"Duplicate relationship ID: {rel_id}"
                )
            seen_ids.add(rel_id)
            relationship = PackageRelationship(
                id=rel_id,
                type=rel_type,
                target=target,
                target_mode=mode,
                source_part=source_part,
            )
            if not relationship.is_external and target:
                _validate_legacy_trash_target(source_part, target)
                # Keep other malformed-target diagnostics in graph validation.
                try:
                    uri = urlsplit(target)
                except ValueError:
                    relationships.append(relationship)
                    continue
                path = unquote(uri.path)
                resolved = posixpath.normpath(
                    path.lstrip("/")
                    if path.startswith("/")
                    else posixpath.join(posixpath.dirname(source_part), path)
                )
                if not uri.scheme and not uri.netloc and is_relationships_part(resolved):
                    raise PackageFormatError(
                        "relationship_to_relationships_part",
                        "Internal relationship targets a relationships part.",
                        rels_name,
                    )
            relationships.append(relationship)
        return relationships

    def part(self, name: str) -> PackagePart | None:
        """Look up a package part by normalized name."""
        return self.parts.get(name)

    def has_part(self, name: str) -> bool:
        """Return True if part exists in graph."""
        return name in self.parts

    @property
    def relationships_by_type(self) -> Mapping[str, tuple[PackageRelationship, ...]]:
        """Return the lazily built, immutable relationship-type index."""
        with self._lock:
            if self._relationships_by_type is None:
                grouped: dict[str, list[PackageRelationship]] = {}
                for relationship in self._iter_relationships():
                    grouped.setdefault(relationship.type, []).append(relationship)
                self._relationships_by_type = MappingProxyType(
                    {key: tuple(value) for key, value in grouped.items()}
                )
                self.counters.relationship_type_index_build_count += 1
            return self._relationships_by_type

    @property
    def relationships_by_target(self) -> Mapping[str, tuple[PackageRelationship, ...]]:
        """Return the lazily built, immutable raw-relationship-target index."""
        with self._lock:
            if self._relationships_by_target is None:
                grouped: dict[str, list[PackageRelationship]] = {}
                for relationship in self._iter_relationships():
                    grouped.setdefault(relationship.target, []).append(relationship)
                self._relationships_by_target = MappingProxyType(
                    {key: tuple(value) for key, value in grouped.items()}
                )
                self.counters.relationship_target_index_build_count += 1
            return self._relationships_by_target

    @property
    def parts_by_content_type(self) -> Mapping[str, tuple[PackagePart, ...]]:
        """Return the lazily built, immutable content-type index."""
        with self._lock:
            if self._parts_by_content_type is None:
                grouped: dict[str, list[PackagePart]] = {}
                for part in self.parts.values():
                    grouped.setdefault(part.content_type, []).append(part)
                self._parts_by_content_type = MappingProxyType(
                    {key: tuple(value) for key, value in grouped.items()}
                )
                self.counters.parts_by_content_type_index_build_count += 1
            return self._parts_by_content_type

    @property
    def wordprocessingml_feature_counts(self) -> WordprocessingMlFeatureCounts:
        """Return lazy counts through a scalar stream, without retaining an XML tree."""
        with self._lock:
            if self._wordprocessingml_feature_counts is None:
                document = self.parts.get("word/document.xml")
                counted = {name: 0 for name, _ in _WORDPROCESSINGML_FEATURE_TAGS}
                field_count = 0
                if document is not None:
                    local_to_name = {
                        local_name: name for name, local_name in _WORDPROCESSINGML_FEATURE_TAGS
                    }
                    for event in document.iter_xml_events():
                        if event.kind != "start" or not event.tag.startswith("{"):
                            continue
                        namespace, local_name = event.tag[1:].split("}", 1)
                        if namespace not in _WORDPROCESSINGML_NAMESPACES:
                            continue
                        feature_name = local_to_name.get(local_name)
                        if feature_name is not None:
                            counted[feature_name] += 1
                        if local_name in {"fldChar", "fldSimple"}:
                            field_count += 1
                    counts = WordprocessingMlFeatureCounts(
                        **counted,
                        fields=field_count,
                        headers=sum(
                            name.startswith("word/header") and name.endswith(".xml")
                            for name in self.parts
                        ),
                        footers=sum(
                            name.startswith("word/footer") and name.endswith(".xml")
                            for name in self.parts
                        ),
                    )
                else:
                    counts = WordprocessingMlFeatureCounts()
                self._wordprocessingml_feature_counts = counts
                self.counters.wordprocessingml_feature_counts_build_count += 1
            return self._wordprocessingml_feature_counts

    def _iter_relationships(self) -> Iterator[PackageRelationship]:
        yield from self.relationships
        for part in self.parts.values():
            yield from part.relationships

    def parts_by_type(self, content_type: str) -> tuple[PackagePart, ...]:
        """Return all parts matching exact MIME content type."""
        return self.parts_by_content_type.get(content_type, ())

    def canonical_xml_digest(self, name: str) -> str:
        """Return a lazy semantic digest for one immutable XML part projection."""
        with self._lock:
            if self._canonical_xml_digests is not None:
                digest = self._canonical_xml_digests.get(name)
                if digest is not None:
                    return digest

            part = self.parts.get(name)
            if part is None:
                raise PackageError("missing_part", f"Part {name!r} not in package", name)
            digest = semantic_sha256(self._canonical_xml_projection(part.parsed_xml()))
            if self._canonical_xml_digests is None:
                self._canonical_xml_digests = {}
            self._canonical_xml_digests[name] = digest
            self.counters.canonical_xml_digest_build_count += 1
            return digest

    @staticmethod
    def _canonical_xml_projection(element: PackageXmlElement) -> tuple[object, ...]:
        return (
            element.tag,
            tuple(sorted(element.attrib.items())),
            element.text,
            element.tail,
            tuple(PackageGraph._canonical_xml_projection(child) for child in element.children),
        )

    def relationships_for(self, source_part: str) -> tuple[PackageRelationship, ...]:
        """Return relationships originating from source_part (empty string for package root)."""
        if source_part == "":
            return self.relationships
        part = self.parts.get(source_part)
        return part.relationships if part is not None else ()

    def validate_relationship_graph(self) -> None:
        """Raise when an internal relationship cannot resolve inside this package."""
        names = {name.translate(_ASCII_CASE_FOLD) for name in self.parts}
        for relationship in self._iter_relationships():
            if not relationship.id or not relationship.type:
                raise PackageFormatError(
                    "relationship_invalid",
                    "Package relationship is missing an ID or type.",
                    relationship.source_part or "_rels/.rels",
                )
            if relationship.is_external:
                continue
            if not relationship.target:
                raise PackageFormatError(
                    "relationship_target_missing",
                    "Internal package relationship has no target.",
                    relationship.source_part or "_rels/.rels",
                )
            target = relationship.resolved_target()
            if target in {"", ".", ".."} or target.startswith("../"):
                raise PackageFormatError(
                    "relationship_target_invalid",
                    "Internal package relationship escapes the package.",
                    relationship.source_part or "_rels/.rels",
                )
            if is_relationships_part(target):
                raise PackageFormatError(
                    "relationship_to_relationships_part",
                    "Internal relationship targets a relationships part.",
                    self._rels_part_name(relationship.source_part),
                )
            if target.translate(_ASCII_CASE_FOLD) not in names:
                raise PackageFormatError(
                    "relationship_target_missing",
                    "Internal package relationship targets a missing part.",
                    relationship.source_part or "_rels/.rels",
                )

    def read_part_bytes(self, name: str) -> bytes:
        """Read and return part bytes, using part cache."""
        p = self.parts.get(name)
        if p is None:
            raise PackageError("missing_part", f"Part {name!r} not in package", name)
        return p.read_bytes()

    def parse_part_xml(self, name: str) -> PackageXmlElement:
        """Parse and return immutable XML projection for part."""
        p = self.parts.get(name)
        if p is None:
            raise PackageError("missing_part", f"Part {name!r} not in package", name)
        return p.parsed_xml()

    def schema_references(self, part_name: str | None = None) -> tuple[PackageSchemaReference, ...]:
        """List existing schema references with exact XML part/path identity.

        With an explicit part, bounded XML parsing is required; without one, scan
        XML-named/XML-typed parts in archive order. ds:schemaRef (§22.5.2),
        sl:schema (§23.2.1), and w:attachedSchema (§17.15.1) are returned without
        resolving URLs, validating schemas, or collapsing duplicate references.
        Malformed/unsafe XML raises the normal typed package errors.
        """
        names = (
            (part_name,)
            if part_name is not None
            else tuple(
                name
                for name, part in self.parts.items()
                if self._is_xml_part(name, part.content_type)
            )
        )
        found: list[PackageSchemaReference] = []
        with self.archive_session():
            for name in names:
                for path, element in self._xml_paths(self.parse_part_xml(name)):
                    if element.tag in _SCHEMA_REFERENCE_TAGS:
                        found.append(PackageSchemaReference(name, path, element))
        return tuple(found)

    def element_properties(self, part_name: str) -> tuple[PackageElementProperties, ...]:
        """List smart-tag/custom-XML owners in one explicitly named XML part.

        The property containers are w:smartTagPr and w:customXmlPr (ISO/IEC
        29500-1 §17.5.1). Unknown attributes/children are retained. Duplicate
        property containers are refused rather than silently selecting one.
        """
        found: list[PackageElementProperties] = []
        for path, owner in self._xml_paths(self.parse_part_xml(part_name)):
            property_tag = _ELEMENT_PROPERTY_TAGS.get(owner.tag)
            if property_tag is None:
                continue
            properties = tuple(child for child in owner if child.tag == property_tag)
            if len(properties) > 1:
                raise PackageFormatError(
                    "duplicate_element_properties",
                    "XML owner has multiple properties containers.",
                    part_name,
                )
            found.append(
                PackageElementProperties(
                    part_name,
                    path,
                    owner.tag,
                    owner.attrib,
                    properties[0] if properties else None,
                )
            )
        return tuple(found)

    def edit_element_properties(
        self,
        part_name: str,
        element_path: tuple[int, ...],
        properties: bytes | None,
        *,
        expected: PackageElementProperties | None = None,
    ) -> bytes:
        """Replace/remove one owner's properties; return a new package's bytes.

        ``element_path`` is the exact zero-based child-index path from listing,
        not XPath or an ordinal across parts. ``properties`` is a complete XML
        w:smartTagPr/w:customXmlPr subtree in the owner's namespace, or None to
        remove it. Parsing and serialization preserve lexical namespace bindings.
        No schema is invented: supplied property/extension content is preserved,
        not schema-validated. Owner attributes and content remain unchanged.
        ``expected`` optionally refuses a stale listing snapshot. Equal expanded
        values AND namespace scopes are a byte-identical no-op. The graph and
        source are unchanged.
        """
        self._custom_part_name(part_name)
        if not isinstance(element_path, tuple) or any(
            type(index) is not int or index < 0 for index in element_path
        ):
            raise PackageSafetyError(
                "element_path_invalid", "Invalid XML child-index path.", part_name
            )
        root = self._parse_member_xml(part_name, self.read_part_bytes(part_name))
        owner = root
        for index in element_path:
            if index >= len(owner):
                raise PackageFormatError(
                    "element_path_missing",
                    "XML element path does not exist.",
                    part_name,
                )
            owner = owner[index]
        property_tag = _ELEMENT_PROPERTY_TAGS.get(owner.tag)
        if property_tag is None:
            raise PackageFormatError(
                "element_kind_invalid",
                "Path is not a smart-tag/custom-XML owner.",
                part_name,
            )
        existing = [child for child in owner if child.tag == property_tag]
        if len(existing) > 1:
            raise PackageFormatError(
                "duplicate_element_properties",
                "XML owner has multiple properties containers.",
                part_name,
            )
        old = existing[0] if existing else None
        snapshot = PackageElementProperties(
            part_name,
            element_path,
            owner.tag,
            MappingProxyType(dict(owner.attrib)),
            PackageXmlElement.from_etree(old) if old is not None else None,
        )
        if expected is not None and expected != snapshot:
            raise PackageFormatError(
                "element_properties_stale",
                "Properties identity/snapshot no longer matches.",
                part_name,
            )
        new = None
        if properties is not None:
            self._part_payload(part_name, properties)
            new = self._parse_mutation_xml(part_name, properties)
            if new.tag != property_tag:
                raise PackageFormatError(
                    "element_properties_kind_invalid",
                    "Properties root does not match owner.",
                    part_name,
                )
            # Container-tail whitespace belongs to the containing document.
            new.tail = old.tail if old is not None else None
        projected = PackageXmlElement.from_etree(new) if new is not None else None
        same_scope = (
            new is None
            and old is None
            or (
                new is not None
                and old is not None
                and getattr(new, "_namespace_scope", {}) == getattr(old, "_namespace_scope", {})
            )
        )
        if projected == snapshot.properties and same_scope:
            return self._mutation_source()
        if old is not None:
            index = list(owner).index(old)
            if new is None:
                # Retain the removed container's tail as surrounding content.
                if index:
                    owner[index - 1].tail = (owner[index - 1].tail or "") + (old.tail or "")
                else:
                    owner.text = (owner.text or "") + (old.tail or "")
            owner.remove(old)
        else:
            index = 0
        if new is not None:
            owner.insert(index, new)
        return self._write_mutation({part_name: serialize_xml_with_preserved_qnames(root)}, set())

    def add_custom_part(
        self,
        name: str,
        data: bytes,
        content_type: str,
        *,
        relationships: Sequence[PackageRelationship] = (),
    ) -> bytes:
        """Add a caller-named part and explicit inbound/outbound relationships.

        Names are canonical ZIP member names (no leading slash), not filesystem
        paths. Existing parts are never overwritten. Each supplied relationship
        must either originate from this new part or internally target it; sources
        are canonical part names or "" for the package root. Targets are OPC URI
        references relative to their source, or package-absolute references.
        IDs are source-scoped and collisions are refused. No relationship type,
        custom XML datastore identity, or schema is inferred.

        An exact content-type override and required relationship-part content
        types are maintained without modifying defaults for unrelated parts.
        Returns validated package bytes; the immutable graph/source is untouched.
        """
        self._custom_part_name(name)
        self._part_payload(name, data)
        self._content_type_value(name, content_type)
        if name in self.parts:
            raise PackageFormatError(
                "part_exists", "Adding a part would overwrite an existing part.", name
            )
        changed = {name: data}
        final_names = set(self.parts) | {name}
        validate_part_names((*self.parts, name))
        roots: dict[str, ET.Element] = {}
        for relationship in relationships:
            self._validate_mutation_relationship(relationship, final_names)
            if relationship.source_part != name and (
                relationship.is_external or self._relationship_part_target(relationship) != name
            ):
                raise PackageFormatError(
                    "relationship_unrelated",
                    "Added relationship does not involve the new part.",
                    name,
                )
            rels_name = self._rels_part_name(relationship.source_part)
            root = roots.get(rels_name)
            if root is None:
                root = self._editable_metadata(rels_name, f"{{{PKG_REL}}}Relationships")
                roots[rels_name] = root
            if any(
                child.get("Id") == relationship.id
                for child in root
                if child.tag == Q_PKG_RELATIONSHIP
            ):
                raise PackageFormatError(
                    "duplicate_relationship_id",
                    "Relationship ID already exists at this source.",
                    rels_name,
                )
            attrs = {
                "Id": relationship.id,
                "Type": relationship.type,
                "Target": relationship.target,
            }
            if relationship.is_external:
                attrs["TargetMode"] = "External"
            ET.SubElement(root, Q_PKG_RELATIONSHIP, attrs)
        for rels_name, root in roots.items():
            changed[rels_name] = serialize_xml_with_preserved_qnames(root)
        self._update_mutation_content_types(changed, set(), {name: content_type})
        return self._write_mutation(changed, set())

    def replace_custom_part(
        self,
        name: str,
        data: bytes,
        *,
        content_type: str | None = None,
        relationships: Sequence[PackageRelationship] | None = None,
    ) -> bytes:
        """Replace an existing explicitly named part; never upsert.

        Incoming relationships retain identity. None preserves outgoing
        relationships; an explicit sequence replaces all outgoing edges and must
        use ``source_part=name`` (empty sequence removes them). Unknown metadata
        children are preserved. None preserves the content type; an explicit
        MIME value sets an override. Returns new bytes without changing this graph.
        """
        self._custom_part_name(name)
        self._part_payload(name, data)
        if name not in self.parts:
            raise PackageFormatError("missing_part", "Replacement requires an existing part.", name)
        changed = {name: data} if data != self.read_part_bytes(name) else {}
        removed: set[str] = set()
        if relationships is not None and tuple(relationships) != self.relationships_for(name):
            rels_name = self._rels_part_name(name)
            root = self._editable_metadata(rels_name, f"{{{PKG_REL}}}Relationships")
            for child in list(root):
                if child.tag == Q_PKG_RELATIONSHIP:
                    root.remove(child)
            seen: set[str] = set()
            for relationship in relationships:
                self._validate_mutation_relationship(relationship, set(self.parts))
                if relationship.source_part != name:
                    raise PackageFormatError(
                        "relationship_source_invalid",
                        "Replacement relationships must originate at the replaced part.",
                        name,
                    )
                if relationship.id in seen:
                    raise PackageFormatError(
                        "duplicate_relationship_id",
                        "Duplicate outgoing relationship ID.",
                        rels_name,
                    )
                seen.add(relationship.id)
                attrs = {
                    "Id": relationship.id,
                    "Type": relationship.type,
                    "Target": relationship.target,
                }
                if relationship.is_external:
                    attrs["TargetMode"] = "External"
                ET.SubElement(root, Q_PKG_RELATIONSHIP, attrs)
            if len(root) or root.attrib or (root.text and root.text.strip()):
                changed[rels_name] = serialize_xml_with_preserved_qnames(root)
            elif rels_name in self.parts:
                removed.add(rels_name)
        overrides = {}
        if content_type is not None:
            self._content_type_value(name, content_type)
            if content_type != self.parts[name].content_type:
                overrides[name] = content_type
                if self._is_xml_part(name, content_type):
                    self._parse_mutation_xml(name, data)
        self._update_mutation_content_types(changed, removed, overrides)
        return self._write_mutation(changed, removed)

    def remove_custom_part(self, name: str) -> bytes:
        """Remove one named part, its outgoing metadata, and every inbound edge.

        Relative, absolute, percent-encoded and fragment-bearing internal targets
        are resolved against their actual source. External edges are not mistaken
        for package targets. All .rels parts are examined, including root edges.
        Content-type overrides for removed parts/relationship metadata are deleted;
        shared defaults and unrelated parts stay intact. Target payloads of outgoing
        edges are NOT recursively deleted. XML references such as r:id are outside
        this package-level API and must be removed by the owning document API.
        Missing parts are errors; returns bytes without changing graph/source.
        """
        self._custom_part_name(name)
        if name not in self.parts:
            raise PackageFormatError("missing_part", "Removal requires an existing part.", name)
        removed = {name}
        outgoing = self._rels_part_name(name)
        if outgoing in self.parts:
            removed.add(outgoing)
        changed: dict[str, bytes] = {}
        for rels_name in self.parts:
            if not rels_name.endswith(".rels") or rels_name in removed:
                continue
            root = self._editable_metadata(rels_name, f"{{{PKG_REL}}}Relationships")
            source = self._source_part_for_rels(rels_name)
            touched = False
            for child in list(root):
                if child.tag != Q_PKG_RELATIONSHIP:
                    continue
                relationship = PackageRelationship(
                    child.get("Id", ""),
                    child.get("Type", ""),
                    child.get("Target", ""),
                    child.get("TargetMode", "Internal"),
                    source,
                )
                if (
                    not relationship.is_external
                    and self._relationship_part_target(relationship) in removed
                ):
                    root.remove(child)
                    touched = True
            if touched:
                # Keep even an empty inbound .rels part: other incoming edges may
                # reference its identity. Only the removed source's metadata dies.
                changed[rels_name] = serialize_xml_with_preserved_qnames(root)
        self._update_mutation_content_types(changed, removed, {})
        return self._write_mutation(changed, removed)

    def rename_custom_part(self, old_name: str, new_name: str) -> bytes:
        """Atomically move one custom part while retaining every relationship edge.

        The associated .rels member moves with its source. Incoming edges retain
        ID/type/mode and URI fragments; outgoing relative paths are rebased to
        keep their original destinations. Absolute and external targets keep
        their spelling unless they target the moved part. Content-type overrides
        follow both moved members. Unrelated members are raw-copied, including
        opaque ignorable extensions. Reserved metadata and the main document are
        protected. Returns validated bytes without writing or mutating the graph.
        """
        self._custom_part_name(old_name)
        self._custom_part_name(new_name)
        if old_name not in self.parts:
            raise PackageFormatError("missing_part", "Rename requires an existing part.", old_name)
        protected = {"word/document.xml"}
        for relationship in self.relationships:
            if (
                not relationship.is_external
                and relationship.type.rsplit("/", 1)[-1] == "officeDocument"
            ):
                protected.add(
                    self._relationship_part_target(relationship).translate(_ASCII_CASE_FOLD)
                )
        for name in (old_name, new_name):
            if name.translate(_ASCII_CASE_FOLD) in protected:
                raise PackageSafetyError(
                    "custom_part_name_invalid", "The main document part is protected.", name
                )
        if old_name == new_name:
            return self._mutation_source()
        if new_name in self.parts:
            raise PackageFormatError(
                "part_exists", "Renaming a part would overwrite an existing part.", new_name
            )
        old_rels = self._rels_part_name(old_name)
        new_rels = self._rels_part_name(new_name)
        moved = {old_name: new_name}
        if old_rels in self.parts:
            if new_rels in self.parts and new_rels != old_rels:
                raise PackageFormatError(
                    "part_exists", "Renaming would overwrite relationship metadata.", new_rels
                )
            moved[old_rels] = new_rels
        removed = set(moved)
        validate_part_names(
            (*[name for name in self.parts if name not in removed], *moved.values())
        )
        changed = {new_name: self.read_part_bytes(old_name)}
        old_key = old_name.translate(_ASCII_CASE_FOLD)
        for rels_name in self.parts:
            if not rels_name.endswith(".rels"):
                continue
            source = self._source_part_for_rels(rels_name)
            new_source = new_name if rels_name == old_rels else source
            root = self._editable_metadata(rels_name, f"{{{PKG_REL}}}Relationships")
            touched = False
            for child in root:
                if child.tag != Q_PKG_RELATIONSHIP:
                    continue
                relationship = PackageRelationship(
                    child.get("Id", ""),
                    child.get("Type", ""),
                    child.get("Target", ""),
                    child.get("TargetMode", "Internal"),
                    source,
                )
                if relationship.is_external:
                    continue
                target = self._relationship_part_target(relationship)
                inbound = target.translate(_ASCII_CASE_FOLD) == old_key
                uri = urlsplit(relationship.target)
                rebase = posixpath.dirname(new_source) != posixpath.dirname(
                    source
                ) and not uri.path.startswith("/")
                if not inbound and not rebase:
                    continue
                destination = new_name if inbound else target
                if uri.path.startswith("/"):
                    path = "/" + destination
                elif not uri.path and destination == new_source:
                    path = ""
                else:
                    path = posixpath.relpath(destination, posixpath.dirname(new_source) or ".")
                    # A colon in the first segment must not become a URI scheme.
                    if ":" in path.split("/", 1)[0]:
                        path = "./" + path
                rewritten = urlunsplit(
                    ("", "", quote(path, safe="/!$&'()*+,;=:@-._~%"), "", uri.fragment)
                )
                if rewritten != relationship.target:
                    child.set("Target", rewritten)
                    touched = True
            destination_rels = moved.get(rels_name, rels_name)
            if touched:
                changed[destination_rels] = serialize_xml_with_preserved_qnames(root)
            elif rels_name == old_rels:
                changed[destination_rels] = self.read_part_bytes(rels_name)
        ct_root = self._editable_metadata("[Content_Types].xml", f"{{{CONTENT_TYPES}}}Types")
        overridden: set[str] = set()
        ct_touched = False
        for child in ct_root:
            if child.tag != Q_CONTENT_TYPES_OVERRIDE:
                continue
            name = child.get("PartName", "").lstrip("/")
            if name in moved:
                child.set("PartName", "/" + moved[name])
                overridden.add(moved[name])
                ct_touched = True
        for old, new in moved.items():
            if (
                new not in overridden
                and self._resolve_content_type(
                    new, self.content_type_defaults, self.content_type_overrides
                )
                != self.parts[old].content_type
            ):
                ET.SubElement(
                    ct_root,
                    Q_CONTENT_TYPES_OVERRIDE,
                    {"PartName": "/" + new, "ContentType": self.parts[old].content_type},
                )
                ct_touched = True
        if ct_touched:
            changed["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(ct_root)
        return self._write_mutation(changed, removed)

    @staticmethod
    def _is_xml_part(name: str, content_type: str) -> bool:
        return (
            name.endswith(".xml")
            or content_type.endswith("+xml")
            or content_type in {"application/xml", "text/xml"}
        )

    @staticmethod
    def _xml_paths(
        root: PackageXmlElement,
    ) -> Iterator[tuple[tuple[int, ...], PackageXmlElement]]:
        pending = [((), root)]
        while pending:
            path, element = pending.pop()
            yield path, element
            pending.extend(
                (path + (index,), child) for index, child in reversed(tuple(enumerate(element)))
            )

    @staticmethod
    def _custom_part_name(name: str) -> None:
        if (
            not isinstance(name, str)
            or not name
            or name in {".", ".."}
            or not _safe_member_name(name)
            or posixpath.normpath(name) != name
            or name.translate(_ASCII_CASE_FOLD).endswith(("/", ".rels"))
            or name.translate(_ASCII_CASE_FOLD) == "[content_types].xml"
            or "_rels" in name.translate(_ASCII_CASE_FOLD).split("/")
            or any(char in name for char in ("?", "#", "%", "\x00"))
            or any(ord(char) < 32 for char in name)
        ):
            raise PackageSafetyError(
                "custom_part_name_invalid",
                "Custom part name is unsafe, noncanonical, or reserved.",
                name,
            )
        validate_part_names((name,))

    def _part_payload(self, name: str, data: bytes) -> None:
        if not isinstance(data, bytes):
            raise PackageFormatError("part_payload_invalid", "Part payload must be bytes.", name)
        if len(data) > min(self.limits.max_entry_size, self._materialization_limit):
            raise PackageSafetyError(
                "entry_size_limit",
                "Mutation payload exceeds the active size limit.",
                name,
            )

    @staticmethod
    def _content_type_value(name: str, content_type: str) -> None:
        token_punctuation = "!#$%&'*+-.^_`|~"
        if not isinstance(content_type, str):
            valid = False
        else:
            components = content_type.split("/")
            valid = len(components) == 2 and all(
                component
                and all(
                    char.isascii() and (char.isalnum() or char in token_punctuation)
                    for char in component
                )
                for component in components
            )
        if not valid:
            raise PackageFormatError(
                "content_type_invalid",
                "A nonempty MIME content type is required.",
                name,
            )

    @staticmethod
    def _rels_part_name(source_part: str) -> str:
        if not source_part:
            return "_rels/.rels"
        directory, basename = posixpath.split(source_part)
        return posixpath.join(directory, "_rels", basename + ".rels")

    @staticmethod
    def _relationship_part_target(relationship: PackageRelationship) -> str:
        if not isinstance(relationship.target, str) or any(
            character == "%"
            and (
                len(relationship.target[index + 1 : index + 3]) != 2
                or any(
                    digit not in "0123456789abcdefABCDEF"
                    for digit in relationship.target[index + 1 : index + 3]
                )
            )
            for index, character in enumerate(relationship.target)
        ):
            raise PackageSafetyError(
                "relationship_target_invalid",
                "Malformed internal target URI.",
                relationship.source_part,
            )
        try:
            uri = urlsplit(relationship.target)
            target = unquote(uri.path, errors="strict")
        except (ValueError, UnicodeError) as exc:
            raise PackageSafetyError(
                "relationship_target_invalid",
                "Malformed internal target URI.",
                relationship.source_part,
            ) from exc
        if uri.scheme or uri.netloc or uri.query or "\\" in target or "\x00" in target:
            raise PackageSafetyError(
                "relationship_target_invalid",
                "Internal relationship target is not a package URI.",
                relationship.source_part,
            )
        if not target:
            target = relationship.source_part if uri.fragment else ""
        elif target.startswith("/"):
            target = target[1:]
        else:
            target = posixpath.join(posixpath.dirname(relationship.source_part), target)
        target = posixpath.normpath(target)
        if target in {"", ".", ".."} or not _safe_member_name(target):
            raise PackageSafetyError(
                "relationship_target_invalid",
                "Internal relationship target escapes the package.",
                relationship.source_part,
            )
        return target

    def _validate_mutation_relationship(
        self, relationship: PackageRelationship, names: set[str]
    ) -> None:
        if not isinstance(relationship, PackageRelationship):
            raise PackageFormatError("relationship_invalid", "Expected a PackageRelationship.")
        if (
            not all(
                isinstance(value, str)
                for value in (
                    relationship.id,
                    relationship.type,
                    relationship.target,
                    relationship.target_mode,
                    relationship.source_part,
                )
            )
            or not relationship.id
            or not relationship.type
            or not relationship.target
            or relationship.target_mode not in {"Internal", "External"}
        ):
            raise PackageFormatError(
                "relationship_invalid",
                "Relationship requires ID, type, target and a valid target mode.",
                relationship.source_part,
            )
        if relationship.source_part:
            self._custom_part_name(relationship.source_part)
            if relationship.source_part not in names:
                raise PackageFormatError(
                    "relationship_source_missing",
                    "Relationship source part does not exist.",
                    relationship.source_part,
                )
        if not relationship.is_external:
            target = self._relationship_part_target(relationship)
            if is_relationships_part(target):
                raise PackageFormatError(
                    "relationship_to_relationships_part",
                    "Internal relationship targets a relationships part.",
                    relationship.source_part,
                )
            target_key = target.translate(_ASCII_CASE_FOLD)
            if target not in names and not any(
                name.translate(_ASCII_CASE_FOLD) == target_key for name in names
            ):
                raise PackageFormatError(
                    "relationship_target_missing",
                    "Relationship targets a missing package part.",
                    relationship.source_part,
                )

    def _parse_mutation_xml(self, name: str, data: bytes) -> ET.Element:
        try:
            return _parse_bounded_xml(data, limits=self.limits, total_elements=[0])
        except _XmlSafetyError as exc:
            raise PackageSafetyError(exc.code, exc.message, name) from exc
        except ET.ParseError as exc:
            raise PackageFormatError(
                "malformed_xml", f"Malformed mutation XML: {exc}", name
            ) from exc

    def _editable_metadata(self, name: str, root_tag: str) -> ET.Element:
        root = (
            self._parse_mutation_xml(name, self.read_part_bytes(name))
            if name in self.parts
            else ET.Element(root_tag)
        )
        if root.tag != root_tag:
            raise PackageFormatError("metadata_root_invalid", "Unexpected OPC metadata root.", name)
        return root

    def _update_mutation_content_types(
        self, changed: dict[str, bytes], removed: set[str], overrides: Mapping[str, str]
    ) -> None:
        root = self._editable_metadata("[Content_Types].xml", f"{{{CONTENT_TYPES}}}Types")
        touched = False
        requested = dict(overrides)
        for name in changed:
            if (
                name.endswith(".rels")
                and self._resolve_content_type(
                    name, self.content_type_defaults, self.content_type_overrides
                )
                != _RELATIONSHIPS_CONTENT_TYPE
            ):
                requested[name] = _RELATIONSHIPS_CONTENT_TYPE
        for child in list(root):
            if child.tag == Q_CONTENT_TYPES_OVERRIDE:
                name = child.get("PartName", "").lstrip("/")
                if name in removed or name in requested:
                    root.remove(child)
                    touched = True
        for name, content_type in requested.items():
            ET.SubElement(
                root,
                Q_CONTENT_TYPES_OVERRIDE,
                {"PartName": "/" + name, "ContentType": content_type},
            )
            touched = True
        if touched:
            changed["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(root)

    def _mutation_source(self) -> bytes:
        try:
            data = self._data if self._data is not None else self.path.read_bytes()
        except OSError as exc:
            raise PackageFormatError(
                "package_read_failed", f"Cannot read immutable package source: {exc}"
            ) from exc
        if sha256_bytes(data) != self.file_sha256:
            raise PackageSafetyError(
                "package_source_changed",
                "Source bytes no longer match the immutable package graph.",
            )
        return data

    def _write_mutation(self, changed: Mapping[str, bytes], removed: set[str]) -> bytes:
        """Reuse DocumentSession's raw member writer; never rewrite untouched data."""
        # A local import avoids the document -> package import cycle. These are
        # the existing raw-copy primitives, not a second ZIP writer convention.
        from .document import _copy_member_raw, _raw_member

        source = self._mutation_source()
        if not changed and not removed:
            return source
        names = (set(self.parts) - removed) | set(changed)
        validate_part_names(
            (
                *[name for name in self.parts if name not in removed and name not in changed],
                *changed,
            )
        )
        if len(names) > self.limits.max_entries:
            raise PackageSafetyError("entry_limit", "Mutation exceeds the package entry limit.")
        total_size = sum(
            len(changed[name]) if name in changed else self.parts[name].size for name in names
        )
        if total_size > self.limits.max_uncompressed_size:
            raise PackageSafetyError(
                "package_size_limit", "Mutation exceeds the package size limit."
            )
        if "[Content_Types].xml" in changed:
            defaults, overrides = self._parse_content_types(
                self._parse_mutation_xml("[Content_Types].xml", changed["[Content_Types].xml"])
            )
        else:
            defaults, overrides = (
                self.content_type_defaults,
                self.content_type_overrides,
            )
        total_elements = [0]
        for name, data in changed.items():
            self._part_payload(name, data)
            content_type = self._resolve_content_type(name, defaults, overrides)
            if self._is_xml_part(name, content_type) or name.endswith(".rels"):
                try:
                    _parse_bounded_xml(data, limits=self.limits, total_elements=total_elements)
                except _XmlSafetyError as exc:
                    raise PackageSafetyError(exc.code, exc.message, name) from exc
                except ET.ParseError as exc:
                    raise PackageFormatError(
                        "malformed_xml", f"Malformed mutation XML: {exc}", name
                    ) from exc
        buffer = io.BytesIO()
        try:
            with (
                zipfile.ZipFile(io.BytesIO(source)) as archive,
                zipfile.ZipFile(buffer, "w") as output,
            ):
                output.comment = archive.comment
                for entry in archive.infolist():
                    if entry.filename in removed:
                        continue
                    if entry.filename in changed:
                        output.writestr(copy.copy(entry), changed[entry.filename])
                    else:
                        _copy_member_raw(output, entry, _raw_member(source, entry))
                for name in sorted(set(changed) - set(self.parts)):
                    entry = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                    entry.compress_type = zipfile.ZIP_DEFLATED
                    entry.external_attr = 0o600 << 16
                    output.writestr(entry, changed[name])
        except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
            raise PackageFormatError(
                "package_write_failed", f"Cannot write package mutation: {exc}"
            ) from exc
        data = buffer.getvalue()
        graph = PackageGraph.from_bytes(data, path=self.path, limits=self.limits)
        for name, payload in changed.items():
            if self._is_xml_part(name, graph.parts[name].content_type):
                graph._parse_member_xml(name, payload)
        graph_names = set(graph.parts)
        # Inspect the actual relationship XML, including otherwise orphaned .rels.
        for name in graph.parts:
            if not name.endswith(".rels"):
                continue
            source_part = graph._source_part_for_rels(name)
            if source_part and source_part not in graph_names:
                raise PackageFormatError(
                    "relationship_source_missing",
                    "Relationship metadata belongs to a missing source part.",
                    name,
                )
            if graph._rels_part_name(source_part) != name:
                raise PackageFormatError(
                    "relationship_part_invalid",
                    "Relationship part has no canonical source path.",
                    name,
                )
            root = graph._editable_metadata(name, f"{{{PKG_REL}}}Relationships")
            for relationship in graph._parse_relationships(root, source_part):
                graph._validate_mutation_relationship(relationship, graph_names)
        return data
