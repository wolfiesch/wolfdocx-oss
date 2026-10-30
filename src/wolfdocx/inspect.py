"""Bounded DOCX inspection, legacy fingerprints, and document statistics.

``document_statistics`` reads story text independently of the legacy report and
fingerprint text inventories. Its documented English-only readability heuristic does
not require layout, fonts, or a pronunciation dictionary.

OPC part-name and relationships-part admission share the PackageGraph validators;
inspection retains its bounded parsing, per-part diagnostics, and reuse behavior.
Content-type coverage/uniqueness and relationship references in every retained MCE
branch are diagnosed without changing package bytes.
"""

from __future__ import annotations

import posixpath
import re
import threading
import unicodedata
import zipfile
from collections import OrderedDict
from collections.abc import Callable, Collection, Iterable, Iterator, Mapping
from copy import deepcopy
from dataclasses import asdict, dataclass, field, replace
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import TYPE_CHECKING, Any, cast
from xml.etree import ElementTree as ET

from . import _inspection_cache
from .canonical import (
    semantic_json_bytes,
    semantic_sha256,
    sha256_bytes,
    sha256_file,
)
from .model import DocumentFingerprint, Issue, PartInfo, StyleInfo
from .package_profile import PackageProfile


def xml_escape(data: str, entities: Mapping[str, str] | None = None) -> str:
    """``xml.sax.saxutils.escape``, byte for byte, without importing :mod:`xml.sax`.

    ``xml.sax.saxutils`` imports :mod:`urllib.request` and with it ``http``,
    ``email``, ``ssl``, and ``socket``, none of which serialization uses; they
    would otherwise stay resident in every process that inspects a package.
    """
    data = data.replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;")
    for key, value in (entities or {}).items():
        data = data.replace(key, value)
    return data


def xml_quoteattr(data: str) -> str:
    """``xml.sax.saxutils.quoteattr`` (no extra entities), byte for byte."""
    data = xml_escape(data, {"\n": "&#10;", "\r": "&#13;", "\t": "&#9;"})
    if '"' not in data:
        return f'"{data}"'
    if "'" not in data:
        return f"'{data}'"
    return '"{}"'.format(data.replace('"', "&quot;"))


TRANSITIONAL_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STRICT_WORD_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main"
SUPPORTED_WORD_NAMESPACES = {TRANSITIONAL_WORD_NS, STRICT_WORD_NS}
W = TRANSITIONAL_WORD_NS
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
EXTERNAL_FIELD_CODES = {
    "DATABASE",
    "DDE",
    "DDEAUTO",
    "HYPERLINK",
    "INCLUDEPICTURE",
    "INCLUDETEXT",
    "LINK",
    "RD",
}


def Q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


_XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"
_BASE_NAMESPACE_SCOPE = (("xml", _XML_NAMESPACE), ("", ""))


class _QNameElement(ET.Element):
    """Keep lexical scope out of XML attributes and structural fingerprints."""

    __slots__ = ("_namespace_scope",)

    def __init__(self, tag: str, attrib: dict[str, str]) -> None:
        super().__init__(tag, attrib)
        self._namespace_scope = _BASE_NAMESPACE_SCOPE

    def __copy__(self) -> _QNameElement:
        clone = _QNameElement(self.tag, self.attrib)
        clone._namespace_scope = self._namespace_scope
        clone.text, clone.tail = self.text, self.tail
        clone.extend(self)
        return clone

    def __deepcopy__(self, memo: dict[int, Any]) -> _QNameElement:
        clone = _QNameElement(self.tag, {})
        memo[id(self)] = clone
        clone.tag = deepcopy(self.tag, memo)
        clone.attrib = deepcopy(self.attrib, memo)
        clone._namespace_scope = self._namespace_scope
        clone.text = deepcopy(self.text, memo)
        clone.tail = deepcopy(self.tail, memo)
        clone.extend(deepcopy(child, memo) for child in self)
        return clone


class _QNameTreeBuilder:
    """Consume namespace/start/end callbacks, refusing limits before allocation."""

    def __init__(
        self,
        limits: PackageLimits | None = None,
        total_elements: list[int] | None = None,
    ) -> None:
        self.builder = ET.TreeBuilder(element_factory=_QNameElement)
        self.scopes = [_BASE_NAMESPACE_SCOPE]
        self.pending: dict[str, str] = {}
        self.limits = limits
        self.total_elements = total_elements if total_elements is not None else [0]
        self.part_elements = 0

    def start_ns(self, prefix: str | None, uri: str | None) -> None:
        self.pending[prefix or ""] = uri or ""

    def start(self, tag: str, attrs: dict[str, str]) -> ET.Element:
        self.part_elements += 1
        self.total_elements[0] += 1
        if self.limits is not None:
            if self.part_elements > self.limits.max_xml_elements:
                raise _XmlSafetyError(
                    "xml_element_limit",
                    "XML part exceeds the configured element-count limit.",
                )
            if self.total_elements[0] > self.limits.max_total_xml_elements:
                raise _XmlSafetyError(
                    "xml_total_element_limit",
                    "Package XML exceeds the configured aggregate element-count limit.",
                )
            if len(self.scopes) > self.limits.max_xml_depth:
                raise _XmlSafetyError(
                    "xml_depth_limit",
                    "XML part exceeds the configured nesting-depth limit.",
                )
        scope = self.scopes[-1]
        if self.pending:
            bindings = dict(scope)
            bindings.update(self.pending)
            scope = tuple(bindings.items())
            self.pending.clear()
        self.scopes.append(scope)
        element = self.builder.start(tag, attrs)
        element._namespace_scope = scope
        return element

    def end(self, tag: str) -> ET.Element:
        element = self.builder.end(tag)
        self.scopes.pop()
        return element

    def data(self, text: str) -> None:
        self.builder.data(text)

    def doctype(self, name: str, pubid: str | None, system: str | None) -> None:
        raise ET.ParseError("DTD and entity declarations are forbidden")

    def close(self) -> ET.Element:
        return self.builder.close()


def parse_xml_with_preserved_qnames(data: bytes | str) -> ET.Element:
    """Parse editable XML, retaining each element's inherited lexical bindings.

    Copy/deepcopy and subtree moves retain scope. Attribute strings remain opaque.
    Package admission must still use ``_parse_bounded_xml`` for resource limits.
    """
    parser = ET.XMLParser(target=_QNameTreeBuilder())
    for offset in range(0, len(data), 64 * 1024):
        parser.feed(data[offset : offset + 64 * 1024])
    return parser.close()


_TEXT_ESCAPED = re.compile(r"[&<>\r]")
_ATTRIBUTE_ESCAPED = re.compile("[&<>\n\r\t\"']")


def _escape_text(text: str) -> str:
    """``xml_escape(text, {"\\r": "&#13;"})``, skipping the replacements when none apply."""
    if type(text) is str and _TEXT_ESCAPED.search(text) is None:
        return text
    return xml_escape(text, {"\r": "&#13;"})


def _quote_attribute(value: str) -> str:
    """:func:`xml_quoteattr`, skipping the replacements when none apply."""
    if type(value) is str and _ATTRIBUTE_ESCAPED.search(value) is None:
        return f'"{value}"'
    return xml_quoteattr(value)


def serialize_xml_with_preserved_qnames(root: ET.Element) -> bytes:
    """Serialize the current tree without losing lexical namespace provenance.

    Retain all source bindings, including aliases used only in opaque values.
    Allocate prefixes only for new expanded names, never through the global
    ElementTree namespace registry. Plain, newly authored elements inherit their
    destination scope; parsed subtrees carry their own scope when moved or copied.

    A scope handed down to children is never mutated again, so an element whose
    scope and names resolve without binding anything reuses its parent's scope and
    the cached merges and name resolutions of that scope. Any element that would
    bind a namespace takes the general path, which copies and extends its scope.
    """
    chunks = ["<?xml version='1.0' encoding='utf-8'?>\n"]
    write = chunks.append
    next_prefix = 0
    # Caches below are keyed by object identity; ``alive`` keeps every keyed scope
    # object alive for the whole call so no identity is reused.
    alive: list[object] = []
    merged: dict[tuple[int, int], tuple[dict[str, str], str]] = {}
    resolved: dict[tuple[int, object, bool], str | None] = {}

    def resolve(scope: dict[str, str], name: str | ET.QName, attribute: bool) -> str | None:
        """The lexical form of ``name`` in ``scope``, or None if it needs a new binding."""
        key = (id(scope), name, attribute)
        if key in resolved:
            return resolved[key]
        text = name.text if isinstance(name, ET.QName) else name
        result: str | None = None
        if not text.startswith("{"):
            if attribute or ":" in text or scope.get("") == "":
                result = text
        else:
            uri, local = text[1:].rsplit("}", 1)
            if not uri:
                if attribute or scope.get("") == "":
                    result = local
            elif not attribute and scope.get("") == uri:
                result = local
            else:
                for prefix, bound_uri in scope.items():
                    if prefix and bound_uri == uri:
                        result = f"{prefix}:{local}"
                        break
        resolved[key] = result
        return result

    def write_general(element: ET.Element, inherited: dict[str, str]) -> None:
        nonlocal next_prefix
        scope = inherited.copy()
        alive.append(scope)
        if isinstance(element, _QNameElement):
            scope.update(element._namespace_scope)
        elif element is root and element.tag in {
            Q(CONTENT_TYPES, "Types"),
            Q(PKG_REL, "Relationships"),
        }:
            # Newly authored OPC metadata uses the package default namespace.
            scope[""] = element.tag[1:].split("}", 1)[0]
        for name, value in element.items():
            if name == "xmlns":
                scope[""] = value
            elif isinstance(name, str) and name.startswith("xmlns:"):
                scope[name[6:]] = value

        def qname(name: str | ET.QName, *, attribute: bool = False) -> str:
            nonlocal next_prefix
            name = name.text if isinstance(name, ET.QName) else name
            if not name.startswith("{"):
                if not attribute and ":" not in name:
                    scope[""] = ""
                return name
            uri, local = name[1:].rsplit("}", 1)
            if not uri:
                if not attribute:
                    scope[""] = ""
                return local
            if not attribute and scope.get("") == uri:
                return local
            for prefix, bound_uri in scope.items():
                if prefix and bound_uri == uri:
                    return f"{prefix}:{local}"
            while f"ns{next_prefix}" in scope:
                next_prefix += 1
            prefix = f"ns{next_prefix}"
            next_prefix += 1
            scope[prefix] = uri
            return f"{prefix}:{local}"

        tag = qname(element.tag)
        attributes = []
        for name, value in element.items():
            if name == "xmlns" or (isinstance(name, str) and name.startswith("xmlns:")):
                continue
            attribute_name = qname(name, attribute=True)
            if isinstance(value, ET.QName):
                value = qname(value, attribute=True)
            attributes.append((attribute_name, value))
        write("<" + tag)
        for prefix, uri in scope.items():
            if inherited.get(prefix) != uri:
                name = f"xmlns:{prefix}" if prefix else "xmlns"
                write(f" {name}={xml_quoteattr(uri)}")
        for name, value in attributes:
            write(f" {name}={xml_quoteattr(value)}")
        write_content(element, tag, scope)

    def write_content(element: ET.Element, tag: str, scope: dict[str, str]) -> None:
        text = element.text
        if text or len(element):
            write(">")
            if text:
                write(_escape_text(text))
            for child in element:
                write_element(child, scope)
            write(f"</{tag}>")
        else:
            write(" />")

    def write_element(element: ET.Element, inherited: dict[str, str]) -> None:
        tag = element.tag
        if tag is ET.Comment:
            write(f"<!--{element.text or ''}-->")
        elif tag is ET.ProcessingInstruction:
            write(f"<?{element.text or ''}?>")
        elif tag is None:
            if element.text:
                write(_escape_text(element.text))
            for child in element:
                write_element(child, inherited)
        else:
            scope, declarations = inherited, ""
            general = False
            if isinstance(element, _QNameElement):
                bindings = element._namespace_scope
                key = (id(inherited), id(bindings))
                entry = merged.get(key)
                if entry is None:
                    candidate = inherited.copy()
                    candidate.update(bindings)
                    if candidate == inherited:
                        entry = (inherited, "")
                    else:
                        entry = (
                            candidate,
                            "".join(
                                f" {f'xmlns:{prefix}' if prefix else 'xmlns'}={xml_quoteattr(uri)}"
                                for prefix, uri in candidate.items()
                                if inherited.get(prefix) != uri
                            ),
                        )
                    alive.append(candidate)
                    alive.append(bindings)
                    merged[key] = entry
                scope, declarations = entry
            elif element is root:
                general = True
            items = element.items()
            qualified = None if general else resolve(scope, tag, False)
            attributes = []
            if qualified is not None:
                for name, value in items:
                    if type(name) is not str or name == "xmlns" or name.startswith("xmlns:"):
                        qualified = None
                        break
                    attribute_name = resolve(scope, name, True)
                    if attribute_name is None:
                        qualified = None
                        break
                    if isinstance(value, ET.QName):
                        value = resolve(scope, value, True)
                        if value is None:
                            qualified = None
                            break
                    attributes.append((attribute_name, value))
            if qualified is None:
                write_general(element, inherited)
            else:
                write("<" + qualified + declarations)
                for name, value in attributes:
                    write(f" {name}={_quote_attribute(value)}")
                write_content(element, qualified, scope)
        if element.tail:
            write(_escape_text(element.tail))

    base = dict(_BASE_NAMESPACE_SCOPE)
    alive.append(base)
    write_element(root, base)
    return "".join(chunks).encode("utf-8")


REQUIRED_PARTS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
}
MODELED_DOCUMENT_ELEMENT_NAMES = {
    "pPr",
    "rPr",
    "tblPr",
    "tblGrid",
    "trPr",
    "tcPr",
    "sectPr",
    "sdtPr",
    "fldChar",
    "fldSimple",
    "instrText",
    "drawing",
}


@dataclass(frozen=True)
class PackageLimits:
    max_archive_size: int = 128 * 1024 * 1024
    max_entries: int = 5_000
    max_entry_size: int = 64 * 1024 * 1024
    max_uncompressed_size: int = 256 * 1024 * 1024
    max_compression_ratio: float = 500.0
    max_xml_part_size: int = 32 * 1024 * 1024
    max_xml_elements: int = 500_000
    max_total_xml_elements: int = 1_000_000
    max_xml_depth: int = 256

    def to_dict(self) -> dict[str, int | float]:
        return asdict(self)


class _XmlSafetyError(Exception):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


@dataclass
class InspectionReport:
    path: str
    file_sha256: str
    limits: dict[str, int | float] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)
    parts: list[PartInfo] = field(default_factory=list)
    fingerprint: DocumentFingerprint = field(default_factory=DocumentFingerprint)
    active_content: dict[str, Any] = field(default_factory=dict)
    # Read-only LabelInfo inventories; attributes retain their original lexical values.
    sensitivity_labels: list[dict[str, Any]] = field(default_factory=list)
    # Retained only in memory for scoped assertions; never serialized into evidence.
    part_texts: dict[str, str] = field(default_factory=dict, repr=False)

    @property
    def valid(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)

    def to_dict(self, *, full: bool = False) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "path": self.path,
            "file_sha256": self.file_sha256,
            "limits": self.limits,
            "valid": self.valid,
            "issues": [issue.to_dict() for issue in self.issues],
            "parts": [part.to_dict() for part in self.parts],
            "fingerprint": self.fingerprint.to_dict(full=full),
            "active_content": self.active_content,
            "sensitivity_labels": self.sensitivity_labels,
        }


def _stable_hash(value: Any) -> str:
    return semantic_sha256(value)


def _canonical_xml(
    element: ET.Element,
    *,
    blank_text_tags: set[str] | None = None,
    skip_tags: set[str] | None = None,
) -> Any:
    blank_text_tags = blank_text_tags or set()
    skip_tags = skip_tags or set()
    text = ""
    if element.tag not in blank_text_tags and element.text and element.text.strip():
        text = element.text
    return {
        "tag": element.tag,
        "attributes": sorted(element.attrib.items()),
        "text": text,
        "children": [
            _canonical_xml(
                child,
                blank_text_tags=blank_text_tags,
                skip_tags=skip_tags,
            )
            for child in element
            if child.tag not in skip_tags
        ],
    }


def _element_namespace(element: ET.Element | None) -> str | None:
    if element is None or not isinstance(element.tag, str) or not element.tag.startswith("{"):
        return None
    return element.tag[1:].split("}", 1)[0]


def _attr(element: ET.Element | None, local: str, namespace: str) -> str | None:
    if element is None:
        return None
    return element.get(Q(namespace, local))


def _safe_member_name(name: str) -> bool:
    path = PurePosixPath(name)
    return not path.is_absolute() and ".." not in path.parts and "\\" not in name


def _xml_tree_shape(root: ET.Element) -> tuple[int, int]:
    """Return the element nesting depth of ``root`` (1 for a childless root) and its
    element count, from one level-by-level walk that visits each element once."""

    height = count = 0
    level = [root]
    while level:
        height += 1
        count += len(level)
        below: list[ET.Element] = []
        for element in level:
            below.extend(element)
        level = below
    return height, count


def _parse_bounded_xml(
    data: bytes,
    *,
    limits: PackageLimits,
    total_elements: list[int],
    preserve_qnames: bool = True,
) -> ET.Element:
    """Parse one package XML part under the configured resource limits.

    When a byte-level upper bound already proves the element-count limits hold and
    the part cannot contain a document type declaration, the C tree builder builds
    the tree: plain elements for read-only consumers (``preserve_qnames=False``) that
    never re-serialize it, scoped elements otherwise. Otherwise it takes the
    element-by-element path, which refuses limits before allocation. Both paths
    refuse the same inputs.
    """
    if len(data) > limits.max_xml_part_size:
        raise _XmlSafetyError(
            "xml_size_limit",
            "XML part exceeds the configured safety limit.",
        )

    # DTDs and custom entities are not part of the OOXML contract WolfDocx
    # needs. Removing NULs makes the conservative check cover UTF-16/32 XML
    # declarations without decoding attacker-controlled text.
    unpadded = data.replace(b"\x00", b"")
    markup_declaration = b"<!" in unpadded
    if markup_declaration:
        declaration_scan = unpadded.upper()
        declaration_scan = re.sub(
            rb"<!--.*?-->",
            b"",
            declaration_scan,
            flags=re.DOTALL,
        )
        declaration_scan = re.sub(
            rb"<\?.*?\?>",
            b"",
            declaration_scan,
            flags=re.DOTALL,
        )
        if b"<!DOCTYPE" in declaration_scan or b"<!ENTITY" in declaration_scan:
            raise _XmlSafetyError(
                "xml_dtd_forbidden",
                "XML part contains a forbidden DTD or entity declaration.",
            )

    if len(unpadded) == len(data):
        # NUL-free bytes are UTF-8 or an ASCII-compatible single-byte encoding, where
        # byte 0x3C only ever encodes ``<``. Every element has exactly one start tag
        # beginning with a ``<`` not followed by ``/``, so this bound over-counts
        # (comments, PIs and CDATA only add to it). A document type declaration always
        # contains a literal ``<!D``; its absence makes the tree builder's missing
        # DOCTYPE hook irrelevant.
        element_bound = data.count(b"<") - data.count(b"</")
        if (
            element_bound <= limits.max_xml_elements
            and total_elements[0] + element_bound <= limits.max_total_xml_elements
            and not (markup_declaration and (b"<!D" in data or b"<!d" in data))
        ):
            try:
                root = _parse_scoped_xml(data) if preserve_qnames else _parse_plain_xml(data)
            except ET.ParseError:
                root = None
            # Malformed or too-deep parts re-run element by element so refusal codes
            # and aggregate element accounting match that path exactly.
            if root is not None:
                height, count = _xml_tree_shape(root)
                if height <= limits.max_xml_depth:
                    total_elements[0] += count
                    return root

    parser = ET.XMLParser(target=_QNameTreeBuilder(limits, total_elements))
    for offset in range(0, len(data), 64 * 1024):
        parser.feed(data[offset : offset + 64 * 1024])
    return parser.close()


def _parse_plain_xml(data: bytes) -> ET.Element:
    """Plain elements from the C tree builder, for read-only consumers."""
    parser = ET.XMLParser()
    parser.feed(data)
    return parser.close()


def _parse_scoped_xml(data: bytes) -> ET.Element:
    """The tree :class:`_QNameTreeBuilder` builds, without a Python call per element.

    The C tree builder creates the elements; its namespace and start/end events then
    assign each element the lexical scope the callback builder would have given it,
    sharing one scope tuple across elements that declare nothing.
    """
    parser = ET.XMLParser(target=ET.TreeBuilder(element_factory=_QNameElement))
    events: list[tuple[str, Any]] = []
    parser._setevents(events, ("start-ns", "start", "end"))  # type: ignore[attr-defined]
    parser.feed(data)
    root = parser.close()
    scopes = [_BASE_NAMESPACE_SCOPE]
    pending: dict[str, str] = {}
    for event, item in events:
        if event == "start":
            scope = scopes[-1]
            if pending:
                bindings = dict(scope)
                bindings.update(pending)
                scope = tuple(bindings.items())
                pending = {}
            item._namespace_scope = scope
            scopes.append(scope)
        elif event == "end":
            scopes.pop()
        else:
            prefix, uri = item
            pending[prefix or ""] = uri or ""
    return root


def _parse_content_types(root: ET.Element) -> tuple[dict[str, str], dict[str, str]]:
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        if child.tag == Q(CONTENT_TYPES, "Default"):
            extension = child.get("Extension")
            content_type = child.get("ContentType")
            if extension and content_type:
                defaults[extension.lower()] = content_type
        elif child.tag == Q(CONTENT_TYPES, "Override"):
            part_name = child.get("PartName")
            content_type = child.get("ContentType")
            if part_name and content_type:
                overrides[part_name.lstrip("/")] = content_type
    return defaults, overrides


def _content_type_for(
    name: str, defaults: Mapping[str, str], overrides: Mapping[str, str]
) -> str | None:
    if name in overrides:
        return overrides[name]
    extension = name.rpartition(".")[2].lower()
    return defaults.get(extension)


def _source_part_for_relationship_part(rels_name: str) -> str | None:
    if rels_name == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in rels_name or not rels_name.endswith(".rels"):
        return None
    prefix, rel_file = rels_name.split(marker, 1)
    return posixpath.join(prefix, rel_file[: -len(".rels")])


def _relationship_part_for_source(source_part: str) -> str:
    if not source_part:
        return "_rels/.rels"
    parent, filename = posixpath.split(source_part)
    rels_directory = posixpath.join(parent, "_rels") if parent else "_rels"
    return posixpath.join(rels_directory, f"{filename}.rels")


def _content_type_issues(root: ET.Element, part_names: Collection[str]) -> list[Issue]:
    """Check declaration uniqueness and coverage using OPC case-insensitive keys."""
    issues: list[Issue] = []
    defaults: set[str] = set()
    overrides: set[str] = set()
    for child in root:
        if child.tag == Q(CONTENT_TYPES, "Default"):
            key = child.get("Extension", "").lower()
            seen, code = defaults, "duplicate_content_type_default"
        elif child.tag == Q(CONTENT_TYPES, "Override"):
            key = child.get("PartName", "").lstrip("/").lower()
            seen, code = overrides, "duplicate_content_type_override"
        else:
            continue
        if key in seen:
            issues.append(
                Issue(
                    code,
                    "error",
                    f"Duplicate content-type declaration: {key}",
                    "[Content_Types].xml",
                )
            )
        if child.get("ContentType"):
            seen.add(key)
    for name in sorted(part_names):
        if name == "[Content_Types].xml":
            continue
        if name.lower() not in overrides and name.rpartition(".")[2].lower() not in defaults:
            issues.append(
                Issue("missing_content_type", "error", "OPC part has no content type.", name)
            )
    return issues


def _relationship_reference_issues(parsed_xml: Mapping[str, ET.Element]) -> list[Issue]:
    """Check all retained branches, as assembly does, against the owning part's IDs."""
    prefixes = (
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}",
        "{http://purl.oclc.org/ooxml/officeDocument/relationships}",
    )
    issues: list[Issue] = []
    for name in parsed_xml:
        if name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        root = parsed_xml[name]
        relationships = parsed_xml.get(_relationship_part_for_source(name))
        declared = (
            {node.get("Id") for node in relationships.findall(Q(PKG_REL, "Relationship"))}
            if relationships is not None
            else set()
        )
        for element in root.iter():
            for attribute, identifier in element.attrib.items():
                if not attribute.startswith(prefixes) or identifier in declared:
                    continue
                # CT_Shape permits an empty optional r:blip for a layout shape
                # without an image. It is not a reference to an undeclared ID.
                if identifier == "" and (element.tag, attribute) in {
                    (
                        "{http://schemas.openxmlformats.org/drawingml/2006/diagram}shape",
                        prefixes[0] + "blip",
                    ),
                    (
                        "{http://purl.oclc.org/ooxml/drawingml/diagram}shape",
                        prefixes[1] + "blip",
                    ),
                }:
                    continue
                issues.append(
                    Issue(
                        "dangling_relationship_reference",
                        "error",
                        f"Relationship reference {identifier!r} has no declaration in "
                        f"{_relationship_part_for_source(name)}.",
                        name,
                    )
                )
    return issues


def _relationship_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def _relationship_kind(value: str | None) -> str:
    if not value:
        return "unknown"
    candidate = value.rstrip("/").rsplit("/", 1)[-1].lower()
    if re.fullmatch(r"[a-z0-9._-]{1,64}", candidate):
        return candidate
    return "other"


def _note_reference_issues(tree: Callable[[str], ET.Element | None]) -> list[Issue]:
    """Check footnote and endnote references; ``tree`` returns a part's parsed root."""
    document = tree("word/document.xml")
    namespace = _element_namespace(document)
    if document is None or namespace not in SUPPORTED_WORD_NAMESPACES:
        return []
    relationships_namespace = (
        "http://purl.oclc.org/ooxml/officeDocument/relationships"
        if namespace == STRICT_WORD_NS
        else "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    issues: list[Issue] = []
    relationship_nodes: list[ET.Element] | None = None
    for kind, plural in (("footnote", "footnotes"), ("endnote", "endnotes")):
        references = list(document.iter(Q(namespace, f"{kind}Reference")))
        if not references:
            continue
        if relationship_nodes is None:
            # Only a story with note references needs its relationships.
            relationships = tree("word/_rels/document.xml.rels")
            relationship_nodes = (
                relationships.findall(Q(PKG_REL, "Relationship"))
                if relationships is not None
                else []
            )
        relationship = next(
            (
                node
                for node in relationship_nodes
                if node.get("Type") == f"{relationships_namespace}/{plural}"
                and node.get("TargetMode") is None
            ),
            None,
        )
        target = relationship.get("Target") if relationship is not None else None
        part = _relationship_target("word/document.xml", target) if target is not None else None
        notes = tree(part) if part is not None else None
        if notes is None or notes.tag != Q(namespace, plural):
            issues.append(
                Issue(
                    f"dangling_{kind}_reference",
                    "error",
                    f"{kind.capitalize()} references require a valid {plural} part.",
                    "word/document.xml",
                )
            )
            continue
        identifiers: set[str] = set()
        for note in notes.findall(Q(namespace, kind)):
            note_id = _attr(note, "id", namespace)
            if note_id is None:
                issues.append(
                    Issue(
                        f"malformed_{kind}",
                        "error",
                        f"{kind.capitalize()} has no identity.",
                        part,
                    )
                )
            elif note_id in identifiers:
                issues.append(
                    Issue(
                        f"duplicate_{kind}_identity",
                        "error",
                        f"{kind.capitalize()} identity {note_id} is duplicated.",
                        part,
                    )
                )
            else:
                identifiers.add(note_id)
        for reference in references:
            note_id = _attr(reference, "id", namespace)
            if note_id is None or note_id not in identifiers:
                issues.append(
                    Issue(
                        f"dangling_{kind}_reference",
                        "error",
                        f"{kind.capitalize()} reference {note_id or '<missing>'} has no body.",
                        "word/document.xml",
                    )
                )
    return issues


def _note_check_inputs(
    parsed_xml: Mapping[str, ET.Element | _DeferredXml], part_sha256: Mapping[str, str]
) -> str | None:
    """Name and digest of every part ``_note_reference_issues`` can read.

    Covers the main story, its relationships, and the target of every internal
    footnotes or endnotes relationship under either namespace (a superset of what the
    check resolves). ``None`` when any of them is missing or failed to parse, because
    the check's issues then depend on more than the parts' bytes.
    """
    rels_name = "word/_rels/document.xml.rels"
    rels = parsed_xml.get(rels_name)
    if rels is None or type(rels) is _DeferredXml:
        return None
    names = ["word/document.xml", rels_name]
    for node in rels.findall(Q(PKG_REL, "Relationship")):
        target = node.get("Target")
        if (
            target is not None
            and node.get("TargetMode") is None
            and (node.get("Type") or "").endswith(("/footnotes", "/endnotes"))
        ):
            names.append(_relationship_target("word/document.xml", target))
    inputs = []
    for name in names:
        if name not in parsed_xml:
            return None
        inputs.append(f"{name}\0{part_sha256[name]}")
    return "\0\0".join(inputs)


def _external_target_scheme(value: str) -> str:
    match = re.match(r"^([A-Za-z][A-Za-z0-9+.-]*):", value)
    return match.group(1).lower() if match else "relative"


def _attribute_by_local_name(element: ET.Element, local_name: str) -> str | None:
    for name, value in element.attrib.items():
        if name.rsplit("}", 1)[-1] == local_name:
            return value
    return None


def _field_instructions(root: ET.Element) -> list[str]:
    instructions: list[str] = []
    stack: list[list[str]] = []
    for element in root.iter():
        local = element.tag.rsplit("}", 1)[-1] if isinstance(element.tag, str) else ""
        if local == "fldSimple":
            value = _attribute_by_local_name(element, "instr")
            if value:
                instructions.append(value)
        elif local == "fldChar":
            field_type = _attribute_by_local_name(element, "fldCharType")
            if field_type == "begin":
                stack.append([])
            elif field_type == "end" and stack:
                value = "".join(stack.pop())
                if value:
                    instructions.append(value)
        elif local == "instrText" and element.text:
            if stack:
                stack[-1].append(element.text)
            else:
                instructions.append(element.text)
    instructions.extend("".join(fragments) for fragments in stack if fragments)
    return instructions


if TYPE_CHECKING:
    from .package import PackageGraph


def _part_active_facts(root: ET.Element) -> tuple[int, tuple[str, ...]]:
    """Return one part's ``altChunk`` count and external field codes, in document order."""
    altchunk_count = sum(
        1
        for element in root.iter()
        if isinstance(element.tag, str) and element.tag.rsplit("}", 1)[-1] == "altChunk"
    )
    codes: list[str] = []
    for instruction in _field_instructions(root):
        match = re.match(r"\s*([A-Za-z]+)", instruction)
        if match is None:
            continue
        code = match.group(1).upper()
        if code in EXTERNAL_FIELD_CODES:
            codes.append(code)
    return altchunk_count, tuple(codes)


def _part_text(root: ET.Element) -> str:
    """Join one part's Word text, deleted-text and field-instruction node values."""
    return "\n".join(
        element.text
        for element in root.iter()
        if isinstance(element.tag, str)
        and element.tag.startswith("{")
        and element.tag[1:].split("}", 1)[0] in SUPPORTED_WORD_NAMESPACES
        and element.tag.rsplit("}", 1)[-1] in {"t", "delText", "instrText"}
        and element.text
    )


_ACTIVE_RELATIONSHIP_KINDS = {
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject": "vba_project",
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility": "custom_ui_2006",
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility": "custom_ui_2009",
    "http://schemas.microsoft.com/office/2011/relationships/webextension": "web_extension",
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes": "task_panes",
}
_CUSTOM_UI_NAMESPACE_SCHEMAS = {
    "http://schemas.microsoft.com/office/2006/01/customui": "2006",
    "http://schemas.microsoft.com/office/2009/07/customui": "2009",
}
_CUSTOM_UI_RELATIONSHIP_SCHEMAS = {"custom_ui_2006": "2006", "custom_ui_2009": "2009"}
_WEB_EXTENSION_CONTENT_TYPE_KINDS = {
    "application/vnd.ms-office.webextension+xml": "web_extension",
    "application/vnd.ms-office.webextensiontaskpanes+xml": "task_panes",
}


def _active_relationship_kind(relationship_type: str | None, source_part: str) -> str | None:
    """Classify an internal relationship the active-content inventory and macro check read.

    ``source_part`` is empty for a package relationship. The main document
    relationship is recognised under the transitional and strict namespaces alike.
    """
    if not relationship_type:
        return None
    kind = _ACTIVE_RELATIONSHIP_KINDS.get(relationship_type)
    if (
        kind is None
        and not source_part
        and relationship_type.rsplit("/", 1)[-1] == "officeDocument"
    ):
        return "office_document"
    return kind


def _is_vba_part(name: str, content_type: str | None) -> bool:
    return "vba" in (content_type or "").lower() or "vbaproject" in name.lower()


def _macro_content_type_issues(
    part_names: Collection[str],
    content_type_defaults: Mapping[str, str],
    content_type_overrides: Mapping[str, str],
    active_relationships: list[dict[str, str]],
) -> list[Issue]:
    """Diagnose a main document content type that disagrees with the package's macros.

    A macro-enabled main type (docm, dotm) needs a VBA project part or relationship, and a
    macro-free main type (docx, dotx) must not carry one. The main part is the target of
    the package's ``officeDocument`` relationship; nothing is diagnosed without it or
    without a content type for it. Diagnostics are warnings: the package still opens.
    """
    main = next(
        (item["target"] for item in active_relationships if item["kind"] == "office_document"),
        None,
    )
    if main is None or main not in part_names:
        return []
    main_type = _content_type_for(main, content_type_defaults, content_type_overrides)
    if not main_type:
        return []
    has_macros = any(item["kind"] == "vba_project" for item in active_relationships) or any(
        _is_vba_part(
            name,
            _content_type_for(name, content_type_defaults, content_type_overrides),
        )
        for name in part_names
    )
    macro_enabled = "macroenabled" in main_type.lower()
    if has_macros and not macro_enabled:
        return [
            Issue(
                "macro_parts_under_macro_free_content_type",
                "warning",
                "The package holds a VBA project but its main document content type is not "
                "macro-enabled.",
                main,
            )
        ]
    if macro_enabled and not has_macros:
        return [
            Issue(
                "macro_enabled_content_type_without_vba_project",
                "warning",
                "The main document content type is macro-enabled but the package holds no "
                "VBA project part or relationship.",
                main,
            )
        ]
    return []


# Per-part cap on each schema-member list (global elements, child elements, attributes).
MAX_SCHEMA_MEMBERS = 1024


def _member_name(tag: str) -> tuple[str, str]:
    """Split a Clark-form tag or attribute key into ``(namespace, local name)``.

    A name with no namespace has the empty namespace, which sorts first.
    """
    if tag.startswith("{"):
        namespace, _, local = tag[1:].partition("}")
        return namespace, local
    return "", tag


def _schema_members(root: ET.Element | None) -> dict[str, Any]:
    """List the schema-member names one CustomUI or web-extension part uses.

    Each name is an object ``{"namespace", "name"}``: the bare local ``name`` as the
    part spells it and its ``namespace`` URI (``None`` for an attribute or element with
    no namespace). The root element is the part's ``global_elements`` member; every
    other element name is a ``child_elements`` member. Names are distinct and sorted
    by namespace then name, and every one is reported whether or not the schema defines
    it. Each list is cut at ``MAX_SCHEMA_MEMBERS`` names and ``truncated`` says whether
    any was. ``status`` is ``unparsed`` when the part held no parsable XML tree; the
    part's ``malformed_xml`` (or XML-safety) issue names the cause and no members are
    listed.
    """
    if root is None:
        return {
            "status": "unparsed",
            "limit": MAX_SCHEMA_MEMBERS,
            "truncated": False,
            "global_elements": [],
            "child_elements": [],
            "attributes": [],
        }
    children: set[tuple[str, str]] = set()
    attributes: set[tuple[str, str]] = set()
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        if element is not root:
            children.add(_member_name(element.tag))
        attributes.update(_member_name(key) for key in element.attrib)
    lists = {
        "global_elements": [_member_name(root.tag)] if isinstance(root.tag, str) else [],
        "child_elements": sorted(children),
        "attributes": sorted(attributes),
    }
    truncated = any(len(names) > MAX_SCHEMA_MEMBERS for names in lists.values())
    return {
        "status": "ok",
        "limit": MAX_SCHEMA_MEMBERS,
        "truncated": truncated,
        **{
            key: [
                {"namespace": namespace or None, "name": name}
                for namespace, name in names[:MAX_SCHEMA_MEMBERS]
            ]
            for key, names in lists.items()
        },
    }


def _member_inventories(
    *,
    parsed_xml: Mapping[str, ET.Element],
    part_hashes: Mapping[str, str],
    content_type_defaults: Mapping[str, str],
    content_type_overrides: Mapping[str, str],
    active_relationships: list[dict[str, str]],
) -> dict[str, list[dict[str, Any]]]:
    """Inventory CustomUI (2006/2009) and web-extension/task-pane parts.

    A part is listed when a relationship targets it, or by its package location
    (``customUI/*.xml``) or content type. Each entry names the part and every
    relationship source reaching it. Only the few candidate parts are parsed. A
    ``schema_members`` object adds the part's schema-member names (:func:`_schema_members`).
    """
    sources: dict[str, dict[str, set[tuple[str, str]]]] = {"custom_ui": {}, "web_extensions": {}}
    kinds: dict[str, str] = {}
    for item in active_relationships:
        kind = item["kind"]
        if kind in _CUSTOM_UI_RELATIONSHIP_SCHEMAS:
            family = "custom_ui"
        elif kind in ("web_extension", "task_panes"):
            family = "web_extensions"
        else:
            continue
        sources[family].setdefault(item["target"], set()).add((item["source_part"], item["type"]))
        kinds.setdefault(item["target"], kind)
    for name in part_hashes:
        lowered = name.lower()
        if lowered.startswith("customui/") and lowered.endswith(".xml") and "/" not in name[9:]:
            sources["custom_ui"].setdefault(name, set())
        content_type = _content_type_for(name, content_type_defaults, content_type_overrides)
        web_kind = _WEB_EXTENSION_CONTENT_TYPE_KINDS.get(content_type or "")
        if web_kind is not None:
            sources["web_extensions"].setdefault(name, set())
            kinds.setdefault(name, web_kind)

    inventories: dict[str, list[dict[str, Any]]] = {}
    for family, targets in sources.items():
        entries: list[dict[str, Any]] = []
        for name in sorted(targets):
            if name not in part_hashes:
                # A relationship to a missing part is already a dangling_relationship issue.
                continue
            content_type = _content_type_for(name, content_type_defaults, content_type_overrides)
            entry: dict[str, Any] = {"part": name}
            if family == "custom_ui":
                namespace = None
                root = parsed_xml.get(name)
                if root is not None and root.tag.startswith("{"):
                    namespace = root.tag[1:].split("}", 1)[0]
                entry["schema"] = _CUSTOM_UI_RELATIONSHIP_SCHEMAS.get(
                    kinds.get(name, ""),
                    _CUSTOM_UI_NAMESPACE_SCHEMAS.get(namespace or "", "unknown"),
                )
                entry["namespace"] = namespace
            else:
                entry["kind"] = kinds[name]
            entry["content_type"] = content_type
            entry["content_sha256"] = part_hashes[name]
            entry["relationship_sources"] = [
                {"source_part": source or "package", "relationship_type": relationship_type}
                for source, relationship_type in sorted(targets[name])
            ]
            entry["schema_members"] = _schema_members(parsed_xml.get(name))
            entries.append(entry)
        if entries:
            inventories[family] = entries
    return inventories


def _active_content_inventory(
    *,
    parsed_xml: Mapping[str, ET.Element],
    parts: list[PartInfo],
    content_type_defaults: Mapping[str, str],
    content_type_overrides: Mapping[str, str],
    external_relationships: list[dict[str, str]],
    active_relationships: list[dict[str, str]],
    part_facts: Mapping[str, tuple[int, tuple[str, ...]]] | None = None,
) -> dict[str, Any]:
    part_hashes = {part.name: part.sha256 for part in parts}
    macro_parts: list[dict[str, str]] = []
    embedded_parts: list[dict[str, str]] = []
    for name in sorted(part_hashes):
        lowered = name.lower()
        content_type = (
            _content_type_for(
                name,
                content_type_defaults,
                content_type_overrides,
            )
            or ""
        ).lower()
        is_vba_part = _is_vba_part(name, content_type)
        is_macro_content_type = "macroenabled" in content_type
        if is_vba_part or is_macro_content_type:
            item = {
                "part": name,
                "kind": ("vba" if is_vba_part else "macro_enabled_content_type"),
            }
            if lowered.endswith(".bin"):
                item["content_sha256"] = part_hashes[name]
            macro_parts.append(item)
        if "/embeddings/" in lowered or "/activex/" in lowered:
            embedded_parts.append(
                {
                    "part": name,
                    "content_sha256": part_hashes[name],
                }
            )

    altchunk_parts: dict[str, int] = {}
    external_field_counts: dict[str, int] = {}
    for name in sorted(parsed_xml):
        facts = part_facts.get(name) if part_facts is not None else None
        altchunk_count, codes = facts if facts is not None else _part_active_facts(parsed_xml[name])
        if altchunk_count:
            altchunk_parts[name] = altchunk_count
        for code in codes:
            external_field_counts[code] = external_field_counts.get(code, 0) + 1

    core = {
        "macro_parts": macro_parts,
        "embedded_parts": embedded_parts,
        "external_relationships": sorted(
            external_relationships,
            key=lambda item: (
                item["source_part"],
                item["kind"],
                item["target_sha256"],
            ),
        ),
        "altchunk_parts": altchunk_parts,
        "external_field_counts": dict(sorted(external_field_counts.items())),
        # Present only when the package holds such parts, so the digest of a package
        # without them is unchanged.
        **_member_inventories(
            parsed_xml=parsed_xml,
            part_hashes=part_hashes,
            content_type_defaults=content_type_defaults,
            content_type_overrides=content_type_overrides,
            active_relationships=active_relationships,
        ),
    }
    return {
        "schema_version": 1,
        "present": any(core.values()),
        "inventory_sha256": _stable_hash(core),
        **core,
    }


def _canonical_properties(element: ET.Element | None) -> dict[str, str]:
    if element is None:
        return {}
    result: dict[str, str] = {}
    for child in element:
        local = child.tag.rsplit("}", 1)[-1]
        attributes = {
            key.rsplit("}", 1)[-1]: value
            for key, value in sorted(child.attrib.items(), key=lambda item: item[0])
        }
        if attributes:
            rendered = ",".join(f"{key}={value}" for key, value in sorted(attributes.items()))
        else:
            rendered = "present"
        if len(child):
            rendered = semantic_json_bytes(_canonical_xml(child)).decode("utf-8")
        result[local] = rendered
    return result


def _style_catalog(
    styles_root: ET.Element | None,
    word_namespace: str,
) -> tuple[dict[str, StyleInfo], dict[str, str]]:
    if styles_root is None:
        return {}, {}

    def word(local):
        return Q(word_namespace, local)

    defaults_rpr = styles_root.find(f"{word('docDefaults')}/{word('rPrDefault')}/{word('rPr')}")
    defaults = _canonical_properties(defaults_rpr)
    styles: dict[str, StyleInfo] = {}
    for style in styles_root.findall(word("style")):
        style_id = _attr(style, "styleId", word_namespace)
        if not style_id:
            continue
        style_info = StyleInfo(
            style_id=style_id,
            style_type=_attr(style, "type", word_namespace),
            name=_attr(style.find(word("name")), "val", word_namespace),
            based_on=_attr(style.find(word("basedOn")), "val", word_namespace),
            run_properties=_canonical_properties(style.find(word("rPr"))),
        )
        styles[style_id] = style_info
    return styles, defaults


def _resolved_style_properties(
    style_id: str | None,
    styles: dict[str, StyleInfo],
    defaults: dict[str, str],
) -> dict[str, str]:
    resolved = dict(defaults)
    chain: list[StyleInfo] = []
    seen: set[str] = set()
    current = style_id
    while current and current not in seen and current in styles:
        seen.add(current)
        info = styles[current]
        chain.append(info)
        current = info.based_on
    for info in reversed(chain):
        resolved.update(info.run_properties)
    return resolved


def _paragraph_record(
    paragraph: ET.Element,
    styles: dict[str, StyleInfo],
    defaults: dict[str, str],
    word_namespace: str,
) -> dict[str, Any]:
    def word(local):
        return Q(word_namespace, local)

    ppr = paragraph.find(word("pPr"))
    paragraph_style = _attr(
        ppr.find(word("pStyle")) if ppr is not None else None,
        "val",
        word_namespace,
    )
    num_pr = ppr.find(word("numPr")) if ppr is not None else None
    numbering = {
        "num_id": _attr(
            num_pr.find(word("numId")) if num_pr is not None else None,
            "val",
            word_namespace,
        ),
        "level": _attr(
            num_pr.find(word("ilvl")) if num_pr is not None else None,
            "val",
            word_namespace,
        ),
    }

    text_chunks: list[str] = []
    runs: list[dict[str, Any]] = []
    for run in paragraph.iter(word("r")):
        run_text = "".join(node.text or "" for node in run.iter(word("t")))
        text_chunks.append(run_text)
        rpr = run.find(word("rPr"))
        character_style = _attr(
            rpr.find(word("rStyle")) if rpr is not None else None,
            "val",
            word_namespace,
        )
        resolved = _resolved_style_properties(paragraph_style, styles, defaults)
        resolved.update(_resolved_style_properties(character_style, styles, {}))
        resolved.update(_canonical_properties(rpr))
        runs.append(
            {
                "text_length": len(run_text),
                "character_style": character_style,
                "resolved_properties": resolved,
            }
        )
    return {
        "text": "".join(text_chunks),
        "paragraph_style": paragraph_style,
        "numbering": numbering,
        "paragraph_properties": _canonical_properties(ppr),
        "runs": runs,
    }


def _border_record(parent: ET.Element | None) -> dict[str, dict[str, str]]:
    if parent is None:
        return {}
    result: dict[str, dict[str, str]] = {}
    for border in parent:
        local = border.tag.rsplit("}", 1)[-1]
        result[local] = {
            key.rsplit("}", 1)[-1]: value
            for key, value in sorted(border.attrib.items(), key=lambda item: item[0])
        }
    return result


def _table_records(
    document_root: ET.Element,
    styles: dict[str, StyleInfo],
    defaults: dict[str, str],
    word_namespace: str,
) -> list[dict[str, Any]]:
    def word(local):
        return Q(word_namespace, local)

    tables: list[dict[str, Any]] = []
    for table_index, table in enumerate(document_root.iter(word("tbl"))):
        tbl_pr = table.find(word("tblPr"))
        grid = table.find(word("tblGrid"))
        table_record: dict[str, Any] = {
            "index": table_index,
            "properties": _canonical_properties(tbl_pr),
            "style": _attr(
                tbl_pr.find(word("tblStyle")) if tbl_pr is not None else None,
                "val",
                word_namespace,
            ),
            "width": _canonical_properties(tbl_pr).get("tblW"),
            "borders": _border_record(
                tbl_pr.find(word("tblBorders")) if tbl_pr is not None else None
            ),
            "grid_widths": [
                _attr(column, "w", word_namespace) for column in grid.findall(word("gridCol"))
            ]
            if grid is not None
            else [],
            "rows": [],
        }
        for row_index, row in enumerate(table.findall(word("tr"))):
            row_record: dict[str, Any] = {
                "index": row_index,
                "properties": _canonical_properties(row.find(word("trPr"))),
                "cells": [],
            }
            for cell_index, cell in enumerate(row.findall(word("tc"))):
                tc_pr = cell.find(word("tcPr"))
                paragraphs = [
                    _paragraph_record(
                        paragraph,
                        styles,
                        defaults,
                        word_namespace,
                    )
                    for paragraph in cell.findall(word("p"))
                ]
                row_record["cells"].append(
                    {
                        "index": cell_index,
                        "text": "\n".join(item["text"] for item in paragraphs),
                        "properties": _canonical_properties(tc_pr),
                        "borders": _border_record(
                            tc_pr.find(word("tcBorders")) if tc_pr is not None else None
                        ),
                        "paragraphs": paragraphs,
                    }
                )
            table_record["rows"].append(row_record)
        tables.append(table_record)
    return tables


def _feature_counts(document_root: ET.Element, word_namespace: str) -> dict[str, int]:
    def word(local):
        return Q(word_namespace, local)

    return {
        "sections": len(list(document_root.iter(word("sectPr")))),
        "comments": len(list(document_root.iter(word("commentReference")))),
        "tracked_insertions": len(list(document_root.iter(word("ins")))),
        "tracked_deletions": len(list(document_root.iter(word("del")))),
        "bookmarks": len(list(document_root.iter(word("bookmarkStart")))),
        "fields": len(list(document_root.iter(word("fldChar"))))
        + len(list(document_root.iter(word("fldSimple")))),
        "hyperlinks": len(list(document_root.iter(word("hyperlink")))),
        "content_controls": len(list(document_root.iter(word("sdt")))),
        "drawings": len(list(document_root.iter(word("drawing")))),
        "headers": 0,
        "footers": 0,
    }


def _part_collection_payload(
    parsed_xml: Mapping[str, ET.Element],
    *,
    prefix: str,
) -> list[dict[str, Any]]:
    return [
        {
            "part": name,
            "xml": _canonical_xml(parsed_xml[name]),
        }
        for name in sorted(parsed_xml)
        if name.startswith(prefix) and name.endswith(".xml")
    ]


def _fingerprint(
    parsed_xml: Mapping[str, ET.Element],
    part_names: set[str],
) -> DocumentFingerprint:
    document_root = parsed_xml.get("word/document.xml")
    if document_root is None:
        return DocumentFingerprint()
    word_namespace = _element_namespace(document_root)
    if word_namespace not in SUPPORTED_WORD_NAMESPACES:
        return DocumentFingerprint()

    def word(local):
        return Q(word_namespace, local)

    styles_by_id, defaults = _style_catalog(
        parsed_xml.get("word/styles.xml"),
        word_namespace,
    )
    paragraphs = [
        _paragraph_record(paragraph, styles_by_id, defaults, word_namespace)
        for paragraph in document_root.iter(word("p"))
    ]
    tables = _table_records(
        document_root,
        styles_by_id,
        defaults,
        word_namespace,
    )
    feature_counts = _feature_counts(document_root, word_namespace)
    feature_counts["headers"] = sum(
        name.startswith("word/header") and name.endswith(".xml") for name in part_names
    )
    feature_counts["footers"] = sum(
        name.startswith("word/footer") and name.endswith(".xml") for name in part_names
    )

    text_payload = [paragraph["text"] for paragraph in paragraphs]
    structure_payload = [
        {
            "paragraph_style": paragraph["paragraph_style"],
            "numbering": paragraph["numbering"],
        }
        for paragraph in paragraphs
    ]
    formatting_payload = [
        {
            "paragraph_properties": paragraph["paragraph_properties"],
            "runs": [
                {
                    "character_style": run["character_style"],
                    "resolved_properties": run["resolved_properties"],
                }
                for run in paragraph["runs"]
            ],
        }
        for paragraph in paragraphs
    ]
    table_payload = [
        {
            "properties": table["properties"],
            "style": table["style"],
            "width": table["width"],
            "borders": table["borders"],
            "grid_widths": table["grid_widths"],
            "rows": [
                {
                    "properties": row["properties"],
                    "cells": [
                        {
                            "properties": cell["properties"],
                            "borders": cell["borders"],
                            "paragraph_formatting": [
                                {
                                    "paragraph_style": paragraph["paragraph_style"],
                                    "paragraph_properties": paragraph["paragraph_properties"],
                                    "runs": [
                                        {
                                            "character_style": run["character_style"],
                                            "resolved_properties": run["resolved_properties"],
                                        }
                                        for run in paragraph["runs"]
                                    ],
                                }
                                for paragraph in cell["paragraphs"]
                            ],
                        }
                        for cell in row["cells"]
                    ],
                }
                for row in table["rows"]
            ],
        }
        for table in tables
    ]
    styles_payload = [
        style.to_dict() for style in sorted(styles_by_id.values(), key=lambda item: item.style_id)
    ]
    styles_root = parsed_xml.get("word/styles.xml")
    numbering_root = parsed_xml.get("word/numbering.xml")
    sections_payload = [_canonical_xml(section) for section in document_root.iter(word("sectPr"))]
    headers_payload = _part_collection_payload(parsed_xml, prefix="word/header")
    footers_payload = _part_collection_payload(parsed_xml, prefix="word/footer")
    comments_payload = _part_collection_payload(parsed_xml, prefix="word/comments")
    field_tags = {word("fldChar"), word("fldSimple"), word("instrText")}
    fields_payload = [
        _canonical_xml(
            element,
            blank_text_tags={word("t"), word("delText")},
        )
        for element in document_root.iter()
        if element.tag in field_tags
    ]
    drawings_payload = [
        _canonical_xml(
            drawing,
            blank_text_tags={word("t"), word("delText")},
        )
        for drawing in document_root.iter(word("drawing"))
    ]
    content_controls_payload = [
        _canonical_xml(properties) for properties in document_root.iter(word("sdtPr"))
    ]

    return DocumentFingerprint(
        document_text_sha256=_stable_hash(text_payload),
        document_nontext_sha256=_stable_hash(
            _canonical_xml(
                document_root,
                blank_text_tags={word("t"), word("delText")},
                skip_tags={word(element_name) for element_name in MODELED_DOCUMENT_ELEMENT_NAMES},
            )
        ),
        structure_sha256=_stable_hash(structure_payload),
        styles_sha256=_stable_hash(_canonical_xml(styles_root))
        if styles_root is not None
        else _stable_hash(styles_payload),
        numbering_sha256=_stable_hash(_canonical_xml(numbering_root))
        if numbering_root is not None
        else "",
        sections_sha256=_stable_hash(sections_payload),
        headers_sha256=_stable_hash(headers_payload) if headers_payload else "",
        footers_sha256=_stable_hash(footers_payload) if footers_payload else "",
        comments_sha256=_stable_hash(comments_payload) if comments_payload else "",
        fields_sha256=_stable_hash(fields_payload) if fields_payload else "",
        drawings_sha256=_stable_hash(drawings_payload) if drawings_payload else "",
        content_controls_sha256=_stable_hash(content_controls_payload)
        if content_controls_payload
        else "",
        formatting_sha256=_stable_hash(formatting_payload),
        tables_sha256=_stable_hash(table_payload),
        features_sha256=_stable_hash(feature_counts),
        paragraph_count=len(paragraphs),
        table_count=len(tables),
        row_count=sum(len(table["rows"]) for table in tables),
        cell_count=sum(len(row["cells"]) for table in tables for row in table["rows"]),
        feature_counts=feature_counts,
        paragraphs=paragraphs,
        tables=tables,
        styles=sorted(styles_by_id.values(), key=lambda item: item.style_id),
    )


def _limits_for_profile(limits: PackageLimits, profile: PackageProfile) -> PackageLimits:
    """Intersect package limits with the profile's explicit materialization bound."""

    materialized_limit = profile.memory.max_materialized_bytes
    return replace(
        limits,
        max_entry_size=min(limits.max_entry_size, materialized_limit),
        max_uncompressed_size=min(limits.max_uncompressed_size, materialized_limit),
        max_xml_part_size=min(limits.max_xml_part_size, materialized_limit),
    )


# Inspections are pure functions of (package bytes, effective limits), so a small
# process-wide memo lets request preparation and the request itself share one
# inspection of the same source, and lets a chained edit reuse the inspection of the
# package the previous edit published. Entries hold read-only parsed trees, so the
# memo is bounded both by entry count and by the XML bytes an entry may retain.
_INSPECTION_MEMO_CAPACITY = 4
_DEFAULT_MAX_XML_DEPTH = PackageLimits().max_xml_depth
_INSPECTION_MEMO_MAX_XML_BYTES = 32 * 1024 * 1024
_inspection_memo: OrderedDict[tuple[str, PackageLimits], _InspectionState] = OrderedDict()
_inspection_memo_lock = threading.Lock()


class _StoredFactMismatch(RuntimeError):
    """A deferred part's bytes no longer agree with the facts that let inspection defer it."""


class _DeferredXml:
    """An XML part whose stored facts let inspection skip parsing it until it is read.

    The part's bytes parsed cleanly under the same limits and inspection code inside a
    valid package (see :mod:`wolfdocx._inspection_cache`). The first read parses them
    with the aggregate element total the full inspection would have started this part
    at, so it takes the same parser path and builds the same tree, then re-verifies the
    stored element count. Disagreement evicts the stored row and raises.
    """

    __slots__ = (
        "_data",
        "_elements",
        "_elements_before",
        "_key",
        "_limits",
        "_lock",
        "_root",
        "_store",
        "name",
    )

    def __init__(
        self,
        name: str,
        data: bytes,
        *,
        limits: PackageLimits,
        elements_before: int,
        elements: int,
        store: _inspection_cache.PartFactStore,
        key: bytes,
    ) -> None:
        self.name = name
        self._data: bytes | None = data
        self._limits = limits
        self._elements_before = elements_before
        self._elements = elements
        self._store = store
        self._key = key
        self._root: ET.Element | None = None
        self._lock = threading.Lock()

    def resolve(self) -> ET.Element:
        root = self._root
        if root is not None:
            return root
        with self._lock:
            if self._root is None:
                assert self._data is not None
                total = [self._elements_before]
                try:
                    root = _parse_bounded_xml(
                        self._data, limits=self._limits, total_elements=total, preserve_qnames=False
                    )
                except (_XmlSafetyError, ET.ParseError):
                    root = None
                if root is None or total[0] - self._elements_before != self._elements:
                    self._store.forget(self._key)
                    raise _StoredFactMismatch(
                        f"Stored inspection facts for {self.name} disagree with its bytes."
                    )
                self._root = root
                self._data = None
            return self._root


class _PartTrees(Mapping[str, ET.Element]):
    """Read-only parsed trees by part name; a deferred part parses on first read."""

    __slots__ = ("_entries",)

    def __init__(self, entries: Mapping[str, ET.Element | _DeferredXml]) -> None:
        self._entries = entries

    def __getitem__(self, name: str) -> ET.Element:
        entry = self._entries[name]
        return entry.resolve() if type(entry) is _DeferredXml else entry

    def __contains__(self, name: object) -> bool:
        return name in self._entries

    def __iter__(self) -> Iterator[str]:
        return iter(self._entries)

    def __len__(self) -> int:
        return len(self._entries)


class _InspectionState:
    """Everything one inspection derived from one package's bytes under one limit set.

    The structural stage (entry safety, bounded XML parsing, relationship and note
    checks) runs eagerly because it alone decides validity. The descriptive stage
    (document fingerprint, part texts, active-content inventory) never adds issues,
    so it runs on first use of :attr:`report`. Parsed trees are shared and must be
    treated as read-only; copy one before mutating it.
    """

    def __init__(
        self,
        report: InspectionReport,
        limits: PackageLimits,
        *,
        parsed_xml: dict[str, ET.Element | _DeferredXml] | None = None,
        part_sha256: dict[str, str] | None = None,
        element_counts: dict[str, int] | None = None,
        content_type_defaults: dict[str, str] | None = None,
        content_type_overrides: dict[str, str] | None = None,
        active_relationships: list[dict[str, str]] | None = None,
        external_relationships: list[dict[str, str]] | None = None,
        entry_names: set[str] | None = None,
        xml_bytes: int = 0,
        reuse: _InspectionState | None = None,
        pending_facts: tuple[_inspection_cache.PartFactStore, dict[bytes, int]] | None = None,
        store: _inspection_cache.PartFactStore | None = None,
    ) -> None:
        self._report = report
        self.limits = limits
        # Raw entries are shared with inspections that reuse this one, so a deferred
        # part parses at most once however many inspections hold it.
        self._entries: dict[str, ET.Element | _DeferredXml] = parsed_xml or {}
        self.parsed_xml: Mapping[str, ET.Element] = _PartTrees(self._entries)
        self.part_sha256: Mapping[str, str] = part_sha256 or {}
        self.element_counts: Mapping[str, int] = element_counts or {}
        self.total_elements = sum(self.element_counts.values())
        self.xml_bytes = xml_bytes
        self._content_type_defaults = content_type_defaults or {}
        self._content_type_overrides = content_type_overrides or {}
        self._external_relationships = external_relationships or []
        self._active_relationships = active_relationships or []
        self._entry_names = entry_names or set()
        self._reuse = reuse
        self._pending_facts = pending_facts
        self._described = parsed_xml is None
        self._part_facts: dict[str, tuple[int, tuple[str, ...]]] = {}
        # Descriptive values are consulted only under default depth bounds: deeper custom
        # bounds may make the descriptive stage raise depending on the caller's stack.
        self._store = store if limits.max_xml_depth <= _DEFAULT_MAX_XML_DEPTH else None
        self._pending_values: dict[bytes, Any] = {}
        self._lock = threading.Lock()
        # True when the persistent store may have decided anything in this state, directly
        # or through a reuse donor; such states are not shared while the store is
        # suspended (``_inspection_cache.gate_scope``).
        self.store_backed = False

    @property
    def valid(self) -> bool:
        # Under default depth bounds the descriptive stage cannot fail, so structural
        # validity is final. Deeper custom bounds may make it raise, and the eager
        # inspection reported that as an error, so describe before answering.
        if self.limits.max_xml_depth > _DEFAULT_MAX_XML_DEPTH:
            return self.report.valid
        return self._report.valid

    @property
    def file_sha256(self) -> str:
        return self._report.file_sha256

    @property
    def parts(self) -> list[PartInfo]:
        return self._report.parts

    @property
    def report(self) -> InspectionReport:
        """Return the complete report, computing its descriptive stage once."""
        if not self._described:
            with self._lock:
                if not self._described:
                    self._describe()
                    self._described = True
                    pending, self._pending_values = self._pending_values, {}
                    if pending and self._store is not None and self._report.valid:
                        self._store.record_values(pending)
        return self._report

    def _describe(self) -> None:
        try:
            self._describe_parts()
        except RuntimeError as exc:
            # Matches the eager inspection, which reported descriptive-stage failures
            # (for example recursion on extreme custom depth limits) as read errors.
            self._report.issues.append(Issue("package_read_error", "error", str(exc)))

    def _describe_parts(self) -> None:
        report = self._report
        reuse = self._reuse if self._reuse is not None and self._reuse._described else None
        self._reuse = None
        store = self._store
        if store is not None:
            limits = self.limits.to_dict()
            word_facts = store.facts(limits, "describe-word-part")
            other_facts = store.facts(limits, "describe-part")
        part_texts: dict[str, str] = {}
        # Iterate names, not trees: a part whose description is known stays deferred.
        for name in self._entries:
            is_word = name.startswith("word/")
            donor = (
                reuse
                if reuse is not None and reuse.part_sha256.get(name) == self.part_sha256[name]
                else None
            )
            facts = donor._part_facts.get(name) if donor is not None else None
            text = donor._report.part_texts.get(name) if donor is not None and is_word else None
            if facts is None or (is_word and text is None):
                key = None
                described = None
                if store is not None:
                    key = (word_facts if is_word else other_facts).key(self.part_sha256[name])
                    described = _decode_part_description(store, key, is_word)
                if described is None:
                    root = self.parsed_xml[name]
                    if facts is None:
                        facts = _part_active_facts(root)
                    if is_word and text is None:
                        text = _part_text(root)
                    if key is not None:
                        self._pending_values[key] = [facts[0], list(facts[1]), text]
                else:
                    facts, text = described
            self._part_facts[name] = facts
            if is_word:
                part_texts[name] = text
        fingerprint = None
        fingerprint_key = None
        if store is not None:
            fingerprint_key = store.facts(limits, "fingerprint").key(
                _fingerprint_inputs(self._entries, self.part_sha256, self._entry_names)
            )
            fingerprint = _decode_fingerprint(store, fingerprint_key)
        if fingerprint is None:
            fingerprint = _fingerprint(self.parsed_xml, self._entry_names)
            if fingerprint_key is not None:
                self._pending_values[fingerprint_key] = asdict(fingerprint)
        report.fingerprint = fingerprint
        report.part_texts = part_texts
        report.active_content = _active_content_inventory(
            parsed_xml=self.parsed_xml,
            parts=report.parts,
            content_type_defaults=self._content_type_defaults,
            content_type_overrides=self._content_type_overrides,
            external_relationships=self._external_relationships,
            active_relationships=self._active_relationships,
            part_facts=self._part_facts,
        )


_FINGERPRINT_PARTS = ("word/document.xml", "word/styles.xml", "word/numbering.xml")
_FINGERPRINT_PREFIXES = ("word/header", "word/footer", "word/comments")


def _fingerprint_inputs(
    entries: Mapping[str, ET.Element | _DeferredXml],
    part_sha256: Mapping[str, str],
    entry_names: Collection[str],
) -> str:
    """Name and state of every part ``_fingerprint`` reads or counts.

    A part is its digest when parsed (or deferred), ``unparsed`` when the package holds
    it but it did not parse, and ``absent`` otherwise.
    """
    names = set(_FINGERPRINT_PARTS)
    for name in (*entries, *entry_names):
        if name.startswith(_FINGERPRINT_PREFIXES) and name.endswith(".xml"):
            names.add(name)
    inputs = []
    for name in sorted(names):
        if name in entries:
            state = part_sha256[name]
        elif name in entry_names:
            state = "unparsed"
        else:
            state = "absent"
        inputs.append(f"{name}\0{state}")
    return "\0\0".join(inputs)


def _decode_part_description(
    store: _inspection_cache.PartFactStore, key: bytes, is_word: bool
) -> tuple[tuple[int, tuple[str, ...]], str | None] | None:
    """Return a stored ``(active facts, text)`` pair; a malformed value is evicted."""
    value = store.lookup_value(key)
    if value is None:
        return None
    if (
        type(value) is list
        and len(value) == 3
        and type(value[0]) is int
        and value[0] >= 0
        and type(value[1]) is list
        and all(type(code) is str for code in value[1])
        and (type(value[2]) is str if is_word else value[2] is None)
    ):
        return (value[0], tuple(value[1])), value[2]
    store.forget_value(key)
    return None


def _decode_fingerprint(
    store: _inspection_cache.PartFactStore, key: bytes
) -> DocumentFingerprint | None:
    """Return a stored fingerprint; a value of the wrong shape is evicted."""
    value = store.lookup_value(key)
    if value is None:
        return None
    try:
        if type(value) is not dict or type(value.get("styles")) is not list:
            raise TypeError("fingerprint is not an object with a styles list")
        styles = [StyleInfo(**style) for style in value["styles"]]
        return DocumentFingerprint(**{**value, "styles": styles})
    except TypeError:
        store.forget_value(key)
        return None


class _Inspection:
    """One path's view of a shared :class:`_InspectionState`."""

    __slots__ = ("_report", "path", "state")

    def __init__(self, path: str, state: _InspectionState) -> None:
        self.path = path
        self.state = state
        self._report: InspectionReport | None = None

    @property
    def valid(self) -> bool:
        return self.state.valid

    @property
    def file_sha256(self) -> str:
        return self.state.file_sha256

    @property
    def limits(self) -> PackageLimits:
        return self.state.limits

    @property
    def parsed_xml(self) -> Mapping[str, ET.Element]:
        """Read-only parsed trees; use :meth:`mutable_root` for a private copy."""
        return self.state.parsed_xml

    @property
    def part_sha256(self) -> Mapping[str, str]:
        return self.state.part_sha256

    @property
    def report(self) -> InspectionReport:
        """The complete report for this path; shared, so callers must not mutate it."""
        if self._report is None:
            base = self.state.report
            self._report = base if base.path == self.path else replace(base, path=self.path)
        return self._report

    def mutable_root(self, part: str) -> ET.Element | None:
        root = self.state.parsed_xml.get(part)
        return None if root is None else deepcopy(root)


def _memoized_inspection(file_sha256: str, limits: PackageLimits) -> _InspectionState | None:
    with _inspection_memo_lock:
        state = _inspection_memo.get((file_sha256, limits))
        if state is not None and state.store_backed and not _inspection_cache.enabled():
            return None
        if state is not None:
            _inspection_memo.move_to_end((file_sha256, limits))
        return state


def _remember_inspection(state: _InspectionState) -> None:
    if not state.valid or state.xml_bytes > _INSPECTION_MEMO_MAX_XML_BYTES:
        return
    with _inspection_memo_lock:
        _inspection_memo[(state.file_sha256, state.limits)] = state
        _inspection_memo.move_to_end((state.file_sha256, state.limits))
        while len(_inspection_memo) > _INSPECTION_MEMO_CAPACITY:
            _inspection_memo.popitem(last=False)


def _persist_part_facts(state: _InspectionState) -> None:
    """Store the facts of parts this inspection parsed, once it is known to be valid."""
    pending, state._pending_facts = state._pending_facts, None
    if pending is not None and state.valid:
        store, facts = pending
        store.record(facts)


def _remembered_inspection(
    path: str | Path, file_sha256: str, limits: PackageLimits | None = None
) -> _Inspection | None:
    """Return the memoized inspection of the bytes ``file_sha256`` names, if any.

    ``path`` is not read: the caller has just hashed it. The result only lends parsed
    trees to parts whose bytes hash identically (see ``_inspect_docx(reuse=...)``), so
    it stays correct even if the file changes afterwards.
    """
    state = _memoized_inspection(file_sha256, limits or PackageLimits())
    return None if state is None else _Inspection(str(Path(path)), state)


def _clear_inspection_memo() -> None:
    """Forget every memoized inspection."""
    with _inspection_memo_lock:
        _inspection_memo.clear()


def _effective_limits(
    limits: PackageLimits | None, profile: PackageProfile | None
) -> PackageLimits:
    limits = limits or PackageLimits()
    if profile is not None:
        if not isinstance(profile, PackageProfile):
            raise TypeError("profile must be a PackageProfile")
        limits = _limits_for_profile(limits, profile)
    return limits


_LABEL_NAMESPACE = "http://schemas.microsoft.com/office/2020/mipLabelMetadata"
_LABEL_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels"
)
_LABEL_CONTENT_TYPE = "application/vnd.ms-office.classificationlabels+xml"
_LABEL_SITE_ID = re.compile(r"\{[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}\}")


def _inspect_sensitivity_labels(
    report: InspectionReport,
    parsed_xml: Mapping[str, ET.Element],
    defaults: dict[str, str],
    overrides: dict[str, str],
) -> None:
    """Inventory admitted LabelInfo parts without rewriting metadata or enforcing policy.

    MS-OFFCRYPTO 2.6.4/2.6.5: id and method are strings, siteId is a lowercase
    braced GUID, and contentBits is an optional unsigned DWORD, including future bits.
    Read through _PartTrees so stored valid XML still contributes its label inventory.
    """
    parts = {
        name
        for name in parsed_xml
        if _content_type_for(name, defaults, overrides) == _LABEL_CONTENT_TYPE
    }
    relationships = parsed_xml.get("_rels/.rels")
    if relationships is not None:
        for relationship in relationships.findall(Q(PKG_REL, "Relationship")):
            if (
                relationship.get("Type") == _LABEL_RELATIONSHIP
                and relationship.get("TargetMode") != "External"
                and relationship.get("Target")
            ):
                parts.add(_relationship_target("", relationship.get("Target", "")))
    for part in sorted(parts):
        root = parsed_xml.get(part)
        if root is None:
            # Malformed, bounded-out and missing parts already have admission diagnostics.
            continue
        if root.tag != Q(_LABEL_NAMESPACE, "labelList"):
            report.issues.append(
                Issue(
                    "sensitivity_label_root_invalid",
                    "error",
                    "LabelInfo requires a clbl:labelList root.",
                    part,
                )
            )
            continue
        labels = []
        seen_sites: set[str] = set()
        for index, element in enumerate(root.findall(Q(_LABEL_NAMESPACE, "label"))):
            labels.append(dict(element.attrib))
            invalid = []
            identifier = element.get("id")
            if not identifier or identifier != identifier.lower():
                invalid.append("id")
            for attribute in ("enabled", "removed"):
                if element.get(attribute, "").strip() not in {"0", "1", "false", "true"}:
                    invalid.append(attribute)
            method = element.get("method")
            if method is None or (
                element.get("removed", "").strip() in {"1", "true"} and method != ""
            ):
                invalid.append("method")
            site = " ".join(element.get("siteId", "").split())
            if _LABEL_SITE_ID.fullmatch(site) is None:
                invalid.append("siteId")
            elif site in seen_sites:
                report.issues.append(
                    Issue(
                        "sensitivity_label_site_duplicate",
                        "error",
                        f"LabelInfo label {index} repeats a siteId.",
                        part,
                    )
                )
            seen_sites.add(site)
            bits = element.get("contentBits")
            if bits is not None:
                lexical = bits.strip()
                # Compare bounded decimal strings, avoiding an unbounded int conversion.
                digits = lexical.removeprefix("+").lstrip("0") or "0"
                if (
                    re.fullmatch(r"\+?[0-9]+", lexical) is None
                    or len(digits) > 10
                    or (len(digits) == 10 and digits > "4294967295")
                ):
                    invalid.append("contentBits")
            for attribute in invalid:
                report.issues.append(
                    Issue(
                        "sensitivity_label_attribute_invalid",
                        "error",
                        f"Invalid LabelInfo label {index} attribute {attribute}.",
                        part,
                    )
                )
        report.sensitivity_labels.append({"part": part, "labels": labels})


def inspect_docx(
    path: str | Path,
    *,
    limits: PackageLimits | None = None,
    profile: PackageProfile | None = None,
) -> InspectionReport:
    return deepcopy(_inspect_docx(path, limits=limits, profile=profile).report)


_INSPECTION_EXECUTION_ISSUES = frozenset({"file_not_found", "package_read_error"})


def inspection_execution_error(report: InspectionReport) -> bool:
    """Whether inspection failed before a document-level decision was possible."""

    return any(
        issue.severity == "error" and issue.code in _INSPECTION_EXECUTION_ISSUES
        for issue in report.issues
    )


def _readability_counts(
    body: ET.Element, namespace: str
) -> tuple[int, int, int, float | None, float | None]:
    """Count main-story readability words, sentences, syllables, and Flesch scores.

    See :class:`DocumentStatistics` for the documented English-only heuristic.
    """
    words = sentences = syllables = 0
    for text in _readability_paragraph_texts(body, namespace):
        sentences += sum(bool(fragment.strip()) for fragment in re.split(r"[.?!]+", text))
        for match in re.finditer(r"[A-Za-z]+(?:['\u2019][A-Za-z]+)*", text):
            words += 1
            word = match.group().lower()
            groups = len(re.findall(r"[aeiouy]+", word))
            if (
                groups > 1
                and word.endswith("e")
                and re.search(r"[bcdfghjklmnpqrstvwxz]le$", word) is None
            ):
                groups -= 1
            syllables += max(1, groups)

    ease = grade = None
    if words and sentences:
        words_per_sentence = words / sentences
        syllables_per_word = syllables / words
        ease = 206.835 - 1.015 * words_per_sentence - 84.6 * syllables_per_word
        grade = 0.39 * words_per_sentence + 11.8 * syllables_per_word - 15.59
    return words, sentences, syllables, ease, grade


def _readability_paragraph_texts(body: ET.Element, namespace: str) -> Iterable[str]:
    """Yield visible main-story paragraphs, retaining cross-paragraph field state."""
    excluded = {
        Q(namespace, name)
        for name in (
            "del",
            "moveFrom",
            "delText",
            "delInstrText",
            "instrText",
            "rt",
            "drawing",
            "pict",
            "object",
            "txbxContent",
        )
    }
    paragraph_tag = Q(namespace, "p")
    text_tag = Q(namespace, "t")
    field_tag = Q(namespace, "fldChar")
    field_type_attr = Q(namespace, "fldCharType")
    separators = {Q(namespace, name) for name in ("tab", "br", "cr")}
    field_results: list[bool] = []

    def run_text(element: ET.Element) -> Iterable[str]:
        if element.tag in excluded:
            return
        if element.tag == field_tag:
            field_type = element.get(field_type_attr)
            if field_type == "begin":
                field_results.append(False)
            elif field_type == "separate" and field_results:
                field_results[-1] = True
            elif field_type == "end" and field_results:
                field_results.pop()
            return
        if element.tag == text_tag:
            if all(field_results):
                yield element.text or ""
            return
        if element.tag in separators:
            if all(field_results):
                yield " "
            return
        for child in element:
            yield from run_text(child)

    def paragraphs(element: ET.Element) -> Iterable[str]:
        if element.tag in excluded:
            return
        if element.tag == paragraph_tag:
            yield "".join(run_text(element))
            return
        for child in element:
            yield from paragraphs(child)

    yield from paragraphs(body)


def _inspect_docx(
    path: str | Path,
    *,
    limits: PackageLimits | None = None,
    profile: PackageProfile | None = None,
    reuse: _Inspection | None = None,
) -> _Inspection:
    """Inspect one package, sharing read-only state with equal-content inspections.

    The package is read into memory once, so the reported ``file_sha256`` is exactly
    the digest of the bytes that were validated. A memoized inspection of the same
    bytes under the same effective limits is returned without re-reading any part.
    ``reuse`` names an inspection (normally the edit's source) whose parsed tree is
    adopted for every part whose bytes hash identically; each adopted tree is still
    charged against the aggregate element limit in package order, and any part that
    would not fit is parsed again so refusals match a from-scratch inspection.
    """
    document_path = Path(path)
    limits = _effective_limits(limits, profile)
    display_path = str(document_path)
    if not document_path.is_file():
        return _Inspection(
            display_path,
            _InspectionState(
                InspectionReport(
                    path=display_path,
                    file_sha256="",
                    limits=limits.to_dict(),
                    issues=[Issue("file_not_found", "error", "The DOCX file does not exist.")],
                ),
                limits,
            ),
        )
    if document_path.stat().st_size > limits.max_archive_size:
        return _Inspection(
            display_path,
            _archive_size_state(display_path, sha256_file(document_path), limits),
        )
    return _inspect_bytes(
        document_path.read_bytes(),
        display_path=display_path,
        limits=limits,
        reuse=reuse,
    )


def _inspect_bytes(
    package_bytes: bytes,
    *,
    display_path: str,
    limits: PackageLimits | None = None,
    file_sha256: str | None = None,
    reuse: _Inspection | None = None,
) -> _Inspection:
    """Inspect package bytes already in memory, exactly as ``_inspect_docx`` would a file.

    ``file_sha256`` may pass the digest a caller already computed for these bytes.
    """
    limits = limits or PackageLimits()
    if file_sha256 is None:
        file_sha256 = sha256_bytes(package_bytes)
    if len(package_bytes) > limits.max_archive_size:
        return _Inspection(display_path, _archive_size_state(display_path, file_sha256, limits))
    state = _memoized_inspection(file_sha256, limits)
    if state is None:
        store_on = _inspection_cache.enabled()
        reuse_state = (
            reuse.state
            if reuse is not None
            and reuse.limits == limits
            and (store_on or not reuse.state.store_backed)
            else None
        )
        try:
            state = _inspect_package_bytes(
                package_bytes,
                display_path=display_path,
                file_sha256=file_sha256,
                limits=limits,
                reuse=reuse_state,
            )
            state.store_backed = store_on or (reuse_state is not None and reuse_state.store_backed)
        except _StoredFactMismatch:
            # A stored fact disagreed with its part (and was evicted). Inspect again from
            # the bytes alone so the result is exactly a from-scratch inspection.
            state = _inspect_package_bytes(
                package_bytes,
                display_path=display_path,
                file_sha256=file_sha256,
                limits=limits,
                reuse=None,
                use_store=False,
            )
        _remember_inspection(state)
        _persist_part_facts(state)
    return _Inspection(display_path, state)


VERIFY_MODES = frozenset({"full", "changed-parts"})


def _require_verify_mode(verify: str) -> str:
    if verify not in VERIFY_MODES:
        raise ValueError("verify_mode_invalid")
    return verify


def _unchanged_parts_identical(
    baseline_sha256: Mapping[str, str],
    staged_names: Iterable[str],
    staged_sha256: Callable[[str], str],
    changed: Collection[str],
) -> bool:
    """True when a staged package names exactly the baseline's parts and every part
    outside ``changed`` (relationship parts and ``[Content_Types].xml`` included)
    hashes identically to the baseline.

    ``staged_sha256`` is only called for those unchanged parts, so a caller that
    hashes lazily never digests a changed part. Both staged-verification modes share
    this: ``verify="changed-parts"`` relies on it instead of re-inspecting unchanged
    parts, and ``verify="full"`` applies it on top of the whole-package inspection,
    which judges validity but not whether an unchanged part still holds its bytes.
    """
    if set(staged_names) != baseline_sha256.keys():
        return False
    return all(
        staged_sha256(name) == digest
        for name, digest in baseline_sha256.items()
        if name not in changed
    )


def _verify_changed_parts(
    baseline: _Inspection,
    staged: Path | bytes,
    changed: Mapping[str, bytes],
) -> tuple[str, dict[str, ET.Element]] | None:
    """Verify a staged package that may differ from a valid baseline only in ``changed``.

    This is the ``verify="changed-parts"`` publication gate. It opens the staged
    package as a validated :class:`PackageGraph` under the baseline's limits, requires
    the same part names, proves every unchanged part (relationship parts and
    ``[Content_Types].xml`` included) hashes identically to the baseline, requires each
    changed part to hold exactly the intended bytes, and re-parses every changed XML
    part under the per-part and aggregate element limits. Note references are then
    re-checked over the baseline trees with the changed parts substituted. What it
    skips is re-parsing and re-fingerprinting the unchanged parts, whose validity the
    baseline inspection already established for the identical bytes.

    ``staged`` is the staged package file, or its bytes when it is still in memory.
    Returns the staged file digest and freshly parsed changed-part roots (private to
    the caller), or ``None`` when any check fails.
    """
    if not baseline.valid:
        return None
    state = baseline.state
    return _verify_staged_changes(
        staged,
        changed,
        limits=baseline.limits,
        unchanged_identical=lambda graph: _unchanged_parts_identical(
            state.part_sha256,
            graph.parts,
            lambda name: graph.parts[name].sha256,
            changed,
        ),
        unchanged_elements=state.total_elements
        - sum(state.element_counts.get(name, 0) for name in changed),
        baseline_tree=state.parsed_xml.get,
    )


def _verify_staged_changes(
    staged: Path | bytes,
    changed: Mapping[str, bytes],
    *,
    limits: PackageLimits,
    unchanged_identical: Callable[[PackageGraph], bool],
    unchanged_elements: int,
    baseline_tree: Callable[[str], ET.Element | None],
    check_notes: bool = True,
    parsed: Mapping[str, tuple[bytes, ET.Element, int]] | None = None,
) -> tuple[str, dict[str, ET.Element]] | None:
    """Verify only the ``changed`` parts of a staged package.

    Refuses a change to ``[Content_Types].xml`` or any relationship part, opens the
    staged package as a validated :class:`PackageGraph` under ``limits``, and asks
    ``unchanged_identical`` to prove every other part holds its source bytes. Each
    changed part must then hold exactly its intended bytes and re-parse under the
    per-part limits and an aggregate that starts at ``unchanged_elements``. With
    ``check_notes``, note references are re-checked over ``baseline_tree`` with the
    changed parts substituted. Returns the staged digest and the changed-part roots,
    or ``None`` when any check fails.

    ``parsed`` maps a part to an unmodified read-only tree that ``_parse_bounded_xml``
    built from exactly the given bytes under ``limits`` and a zero aggregate, with
    its element count. A changed part whose staged bytes equal those bytes reuses
    that parse instead of repeating it: the per-part limits already held, and its
    count joins the aggregate check. Reused roots are returned shared, not private.
    """
    from .package import PackageError, PackageGraph

    if any(name == "[Content_Types].xml" or name.endswith(".rels") for name in changed):
        return None
    try:
        graph = (
            PackageGraph.from_bytes(staged, path="<staged>", limits=limits)
            if isinstance(staged, bytes)
            else PackageGraph.from_file(staged, limits=limits)
        )
        if not unchanged_identical(graph):
            return None
        total_elements = [unchanged_elements]
        roots: dict[str, ET.Element] = {}
        for name in changed:
            data = graph.read_part_bytes(name)
            if data != changed[name]:
                return None
            if name.endswith(".xml"):
                known = parsed.get(name) if parsed is not None else None
                if known is not None and known[0] == data:
                    total_elements[0] += known[2]
                    if total_elements[0] > limits.max_total_xml_elements:
                        return None
                    roots[name] = known[1]
                    continue
                roots[name] = _parse_bounded_xml(
                    data,
                    limits=limits,
                    total_elements=total_elements,
                    preserve_qnames=False,
                )
        if check_notes:
            note_issues = _note_reference_issues(
                lambda name: roots[name] if name in roots else baseline_tree(name)
            )
            if any(issue.severity == "error" for issue in note_issues):
                return None
    except (
        PackageError,
        OSError,
        zipfile.BadZipFile,
        RuntimeError,
        _XmlSafetyError,
        ET.ParseError,
    ):
        return None
    return graph.file_sha256, roots


def _archive_size_state(
    display_path: str, file_sha256: str, limits: PackageLimits
) -> _InspectionState:
    return _InspectionState(
        InspectionReport(
            path=display_path,
            file_sha256=file_sha256,
            limits=limits.to_dict(),
            issues=[
                Issue(
                    "archive_size_limit",
                    "error",
                    "Archive size exceeds the configured safety limit.",
                )
            ],
        ),
        limits,
    )


def _inspect_package_bytes(
    package_bytes: bytes,
    *,
    display_path: str,
    file_sha256: str,
    limits: PackageLimits,
    reuse: _InspectionState | None,
    use_store: bool = True,
) -> _InspectionState:
    # Import lazily: package uses this module's bounded XML parser and limits.
    from .package import (
        PackageError,
        _package_entry_is_part,
        _validate_content_type_part_names,
        _validate_legacy_trash_target,
        is_relationships_part,
        validate_part_names,
        validate_relationships_part,
    )

    report = InspectionReport(
        path=display_path,
        file_sha256=file_sha256,
        limits=limits.to_dict(),
    )
    if not zipfile.is_zipfile(BytesIO(package_bytes)):
        report.issues.append(
            Issue("not_zip", "error", "The file is not a readable ZIP/OPC package.")
        )
        return _InspectionState(report, limits)

    parsed_xml: dict[str, ET.Element | _DeferredXml] = {}
    # Stored per-part facts let a part seen in an earlier valid inspection skip parsing;
    # relationship parts and [Content_Types].xml are always parsed because every
    # inspection reads them.
    store = _inspection_cache.open_store() if use_store else None
    part_facts = store.facts(limits.to_dict(), "plain") if store is not None else None
    note_facts = store.facts(limits.to_dict(), "note-check") if store is not None else None
    reference_facts = (
        store.facts(limits.to_dict(), "relationship-reference-check") if store is not None else None
    )
    new_facts: dict[bytes, int] = {}
    element_counts: dict[str, int] = {}
    part_sha256: dict[str, str] = {}
    content_type_defaults: dict[str, str] = {}
    content_type_overrides: dict[str, str] = {}
    total_xml_elements = [0]
    xml_bytes = 0
    external_relationships: list[dict[str, str]] = []
    active_relationships: list[dict[str, str]] = []

    try:
        with zipfile.ZipFile(BytesIO(package_bytes)) as archive:
            entries = archive.infolist()
            if len(entries) > limits.max_entries:
                report.issues.append(
                    Issue(
                        "entry_limit",
                        "error",
                        f"Package has {len(entries)} entries; limit is {limits.max_entries}.",
                    )
                )
                return _InspectionState(report, limits)
            names = [entry.filename for entry in entries]
            name_set = set(names)
            if len(names) != len(name_set):
                report.issues.append(
                    Issue(
                        "duplicate_entry",
                        "error",
                        "Package contains duplicate ZIP entries.",
                    )
                )
                return _InspectionState(report, limits)
            try:
                part_entries = [entry for entry in entries if _package_entry_is_part(entry)]
                validate_part_names(tuple(entry.filename for entry in part_entries))
                validate_part_names(
                    tuple(
                        name[len("[trash]/") :]
                        for name in names
                        if name.startswith("[trash]/") and not name.endswith("/")
                    )
                )
                name_set = {entry.filename for entry in part_entries}
            except PackageError as exc:
                report.issues.append(Issue(exc.code, "error", exc.message, exc.entry_name))
                return _InspectionState(report, limits)

            total_uncompressed = sum(entry.file_size for entry in entries)
            if total_uncompressed > limits.max_uncompressed_size:
                report.issues.append(
                    Issue(
                        "package_size_limit",
                        "error",
                        "Package uncompressed size exceeds the configured safety limit.",
                    )
                )
                return _InspectionState(report, limits)

            for entry in entries:
                if entry.file_size > limits.max_entry_size:
                    report.issues.append(
                        Issue(
                            "entry_size_limit",
                            "error",
                            "Package entry exceeds the configured safety limit.",
                            entry.filename,
                        )
                    )
                    continue
                ratio = entry.file_size / max(entry.compress_size, 1)
                if ratio > limits.max_compression_ratio:
                    report.issues.append(
                        Issue(
                            "compression_ratio_limit",
                            "error",
                            "Package entry exceeds the configured compression-ratio limit.",
                            entry.filename,
                        )
                    )
                    continue
                if entry.filename not in name_set:
                    continue
                data = archive.read(entry.filename)
                digest = sha256_bytes(data)
                part_sha256[entry.filename] = digest
                if not (entry.filename.endswith(".xml") or entry.filename.endswith(".rels")):
                    continue
                xml_bytes += len(data)
                if reuse is not None and reuse.part_sha256.get(entry.filename) == digest:
                    reused_root = reuse._entries.get(entry.filename)
                    reused_count = reuse.element_counts.get(entry.filename)
                    if (
                        reused_root is not None
                        and reused_count is not None
                        and total_xml_elements[0] + reused_count <= limits.max_total_xml_elements
                    ):
                        total_xml_elements[0] += reused_count
                        parsed_xml[entry.filename] = reused_root
                        element_counts[entry.filename] = reused_count
                        continue
                fact_key = (
                    part_facts.key(digest)
                    if part_facts is not None
                    and entry.filename.endswith(".xml")
                    and entry.filename != "[Content_Types].xml"
                    else None
                )
                if fact_key is not None:
                    stored = part_facts.lookup(fact_key)
                    if (
                        stored is not None
                        and total_xml_elements[0] + stored <= limits.max_total_xml_elements
                    ):
                        parsed_xml[entry.filename] = _DeferredXml(
                            entry.filename,
                            data,
                            limits=limits,
                            elements_before=total_xml_elements[0],
                            elements=stored,
                            store=part_facts.store,
                            key=fact_key,
                        )
                        element_counts[entry.filename] = stored
                        total_xml_elements[0] += stored
                        continue
                elements_before = total_xml_elements[0]
                try:
                    parsed_xml[entry.filename] = _parse_bounded_xml(
                        data,
                        limits=limits,
                        total_elements=total_xml_elements,
                        preserve_qnames=entry.filename.endswith(".rels"),
                    )
                    element_counts[entry.filename] = total_xml_elements[0] - elements_before
                    if fact_key is not None:
                        new_facts[fact_key] = element_counts[entry.filename]
                except _XmlSafetyError as exc:
                    report.issues.append(
                        Issue(
                            exc.code,
                            "error",
                            exc.message,
                            entry.filename,
                        )
                    )
                except ET.ParseError as exc:
                    report.issues.append(
                        Issue(
                            "malformed_xml",
                            "error",
                            f"XML parsing failed: {exc}",
                            entry.filename,
                        )
                    )

            missing = sorted(REQUIRED_PARTS - name_set)
            for required in missing:
                report.issues.append(
                    Issue(
                        "missing_required_part",
                        "error",
                        "Required OPC part is missing.",
                        required,
                    )
                )

            content_types_root = parsed_xml.get("[Content_Types].xml")
            if content_types_root is not None:
                try:
                    _validate_content_type_part_names(content_types_root)
                except PackageError as exc:
                    report.issues.append(Issue(exc.code, "error", exc.message, exc.entry_name))
                content_type_defaults, content_type_overrides = _parse_content_types(
                    content_types_root
                )
                report.issues.extend(_content_type_issues(content_types_root, name_set))

            for rels_name, root in parsed_xml.items():
                if not rels_name.endswith(".rels"):
                    continue
                source_part = _source_part_for_relationship_part(rels_name)
                if source_part is None:
                    continue
                try:
                    validate_relationships_part(root, rels_name)
                except PackageError as exc:
                    report.issues.append(
                        Issue(exc.code, "error", exc.message, exc.entry_name or rels_name)
                    )
                seen_ids: set[str] = set()
                for relationship in root.findall(Q(PKG_REL, "Relationship")):
                    relationship_id = relationship.get("Id")
                    if relationship_id:
                        if relationship_id in seen_ids:
                            report.issues.append(
                                Issue(
                                    "duplicate_relationship_id",
                                    "error",
                                    f"Duplicate relationship id {relationship_id}.",
                                    rels_name,
                                )
                            )
                        seen_ids.add(relationship_id)
                    if relationship.get("TargetMode") == "External":
                        target = relationship.get("Target")
                        relationship_type = relationship.get("Type")
                        if target:
                            external_relationships.append(
                                {
                                    "source_part": source_part or "package",
                                    "kind": _relationship_kind(relationship_type),
                                    "relationship_type_sha256": sha256_bytes(
                                        (relationship_type or "").encode("utf-8")
                                    ),
                                    "target_scheme": _external_target_scheme(target),
                                    "target_sha256": sha256_bytes(target.encode("utf-8")),
                                }
                            )
                        continue
                    target = relationship.get("Target")
                    if not target:
                        report.issues.append(
                            Issue(
                                "missing_relationship_target",
                                "error",
                                "Internal relationship has no target.",
                                rels_name,
                            )
                        )
                        continue
                    try:
                        _validate_legacy_trash_target(source_part, target)
                    except PackageError as exc:
                        report.issues.append(
                            Issue(exc.code, "error", exc.message, exc.entry_name or rels_name)
                        )
                    resolved = _relationship_target(source_part, target)
                    active_kind = _active_relationship_kind(relationship.get("Type"), source_part)
                    if active_kind is not None:
                        active_relationships.append(
                            {
                                "source_part": source_part,
                                "kind": active_kind,
                                "type": relationship.get("Type", ""),
                                "target": resolved,
                            }
                        )
                    if is_relationships_part(resolved):
                        report.issues.append(
                            Issue(
                                "relationship_to_relationships_part",
                                "error",
                                "Internal relationship targets a relationships part.",
                                rels_name,
                            )
                        )
                    if resolved not in name_set:
                        report.issues.append(
                            Issue(
                                "dangling_relationship",
                                "error",
                                f"Internal relationship target is missing: {resolved}",
                                rels_name,
                            )
                        )

            # Reuse only a passing check for the exact part inventory and bytes. Names
            # matter because relationship IDs belong to the source part's .rels file.
            reference_key = (
                reference_facts.key(
                    "\0\0".join(f"{name}\0{part_sha256[name]}" for name in sorted(part_sha256))
                )
                if reference_facts is not None
                else None
            )
            if reference_key is None or reference_facts.lookup(reference_key) != 0:
                reference_issues = _relationship_reference_issues(_PartTrees(parsed_xml))
                report.issues.extend(reference_issues)
                if reference_key is not None and not reference_issues:
                    new_facts[reference_key] = 0

            # An entry refused by a size or ratio limit has no recorded part, so the part
            # inventory is incomplete and the macro check cannot trust its absences.
            if name_set.issubset(part_sha256):
                report.issues.extend(
                    _macro_content_type_issues(
                        name_set,
                        content_type_defaults,
                        content_type_overrides,
                        active_relationships,
                    )
                )
            _inspect_sensitivity_labels(
                report,
                _PartTrees(parsed_xml),
                content_type_defaults,
                content_type_overrides,
            )

            # A note check that passed inside a valid package passes again for the same
            # bytes of every part it can read, so the main story may stay deferred.
            note_inputs = (
                _note_check_inputs(parsed_xml, part_sha256) if note_facts is not None else None
            )
            note_key = note_facts.key(note_inputs) if note_inputs is not None else None
            if note_key is None or note_facts.lookup(note_key) != 0:
                note_issues = _note_reference_issues(_PartTrees(parsed_xml).get)
                report.issues.extend(note_issues)
                if note_key is not None and not note_issues:
                    new_facts[note_key] = 0
            for entry in entries:
                if entry.filename.endswith("/") or entry.filename not in part_sha256:
                    continue
                report.parts.append(
                    PartInfo(
                        name=entry.filename,
                        size=entry.file_size,
                        compressed_size=entry.compress_size,
                        sha256=part_sha256[entry.filename],
                        content_type=_content_type_for(
                            entry.filename,
                            content_type_defaults,
                            content_type_overrides,
                        ),
                    )
                )
            report.parts.sort(key=lambda item: item.name)
    except _StoredFactMismatch:
        raise
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        report.issues.append(Issue("package_read_error", "error", str(exc)))
        return _InspectionState(report, limits)

    return _InspectionState(
        report,
        limits,
        parsed_xml=parsed_xml,
        part_sha256=part_sha256,
        element_counts=element_counts,
        content_type_defaults=content_type_defaults,
        content_type_overrides=content_type_overrides,
        external_relationships=external_relationships,
        active_relationships=active_relationships,
        entry_names=name_set,
        xml_bytes=xml_bytes,
        reuse=reuse,
        pending_facts=(part_facts.store, new_facts) if part_facts and new_facts else None,
        store=store,
    )


def _inspect_package_graph(
    graph: PackageGraph,
    *,
    display_path: str,
) -> InspectionReport:
    """Build an inspection report from one validated graph without reopening it."""
    from .package import PackageError

    report = InspectionReport(
        path=display_path,
        file_sha256=graph.file_sha256,
        limits=graph.limits.to_dict(),
    )
    parsed_xml: dict[str, ET.Element] = {}
    for name, part in graph.parts.items():
        if not name.endswith((".xml", ".rels")):
            continue
        try:
            parsed_xml[name] = cast(ET.Element, part.parsed_xml())
        except PackageError as exc:
            report.issues.append(
                Issue(
                    exc.code,
                    "error",
                    f"XML parsing failed: {exc.message}",
                    name,
                )
            )

    report.parts = sorted(
        [
            PartInfo(
                name=part.name,
                size=part.size,
                compressed_size=part.compressed_size,
                sha256=part.sha256,
                content_type=_content_type_for(
                    part.name,
                    graph.content_type_defaults,
                    graph.content_type_overrides,
                ),
            )
            for part in graph.parts.values()
        ],
        key=lambda item: item.name,
    )
    external_relationships: list[dict[str, str]] = []
    active_relationships: list[dict[str, str]] = []
    relationships = [
        *graph.relationships,
        *(relationship for part in graph.parts.values() for relationship in part.relationships),
    ]
    for relationship in relationships:
        source_part = relationship.source_part
        if relationship.is_external:
            external_relationships.append(
                {
                    "source_part": source_part or "package",
                    "kind": _relationship_kind(relationship.type),
                    "relationship_type_sha256": sha256_bytes(relationship.type.encode("utf-8")),
                    "target_scheme": _external_target_scheme(relationship.target),
                    "target_sha256": sha256_bytes(relationship.target.encode("utf-8")),
                }
            )
            continue
        active_kind = _active_relationship_kind(relationship.type, source_part)
        if active_kind is not None and relationship.target:
            active_relationships.append(
                {
                    "source_part": source_part,
                    "kind": active_kind,
                    "type": relationship.type,
                    "target": relationship.resolved_target(),
                }
            )
        if not relationship.target:
            report.issues.append(
                Issue(
                    "missing_relationship_target",
                    "error",
                    "Internal relationship has no target.",
                    _relationship_part_for_source(source_part),
                )
            )
        elif relationship.resolved_target() not in graph.parts:
            report.issues.append(
                Issue(
                    "dangling_relationship",
                    "error",
                    (f"Internal relationship target is missing: {relationship.resolved_target()}"),
                    _relationship_part_for_source(source_part),
                )
            )
    report.issues.extend(
        _macro_content_type_issues(
            set(graph.parts),
            graph.content_type_defaults,
            graph.content_type_overrides,
            active_relationships,
        )
    )
    content_types_root = parsed_xml.get("[Content_Types].xml")
    if content_types_root is not None:
        report.issues.extend(_content_type_issues(content_types_root, set(graph.parts)))
    report.issues.extend(_relationship_reference_issues(parsed_xml))
    report.fingerprint = _fingerprint(parsed_xml, set(graph.parts))
    report.part_texts = {
        name: _part_text(root) for name, root in parsed_xml.items() if name.startswith("word/")
    }
    report.active_content = _active_content_inventory(
        parsed_xml=parsed_xml,
        parts=report.parts,
        content_type_defaults=dict(graph.content_type_defaults),
        content_type_overrides=dict(graph.content_type_overrides),
        external_relationships=external_relationships,
        active_relationships=active_relationships,
    )
    return report


# -- document statistics -----------------------------------------------------------

_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MARKUP_COMPATIBILITY_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# Story containers whose text Word's document statistics leave out: removed revision
# content, the legacy rendering of a markup-compatibility choice, and text boxes,
# which belong to their own story rather than the main document flow.
_STATISTICS_SKIPPED = frozenset({"del", "moveFrom", "txbxContent"})
# Run children that carry no counted text.
_STATISTICS_SILENT = frozenset(
    {"instrText", "delText", "delInstrText", "rPr", "pPr", "sectPr", "rt"}
)
# Whitespace that separates words. A no-break space is still a space Word counts.
_STATISTICS_SPACES = frozenset(" \u00a0\u2002\u2003\u2009\u3000")
_STATISTICS_BREAKS = frozenset("\t\n\r")


@dataclass(frozen=True, slots=True)
class DocumentStatistics:
    """Statistics for one document, counted over its main story or over all stories.

    With ``stories="main"`` words, characters, and paragraphs are read from the main
    document story the way Word's statistics count it: deleted revision text, field
    instructions, hidden runs, text boxes, headers, footers, comments, and notes are
    left out; a word is a maximal run of non-space characters, and each East Asian
    ideograph is a word of its own; characters exclude, and characters with spaces
    include, the space and tab characters between them; a paragraph counts once it
    holds any text. With ``stories="all"`` they follow LibreOffice's document
    statistics instead (see :func:`document_statistics`).

    ``lines`` and ``pages`` come from ``wolf extract`` with system fonts, if a wolf
    executable is supplied or found on PATH; otherwise they are ``None``. Lines
    always count distinct physical-page/body-baseline bands, so simultaneous
    table-cell lines contribute once, like a vertical visual-row cursor walk.

    ``readability_words``, ``sentences``, ``syllables``, ``flesch_reading_ease``, and
    ``flesch_kincaid_grade`` are deterministic English-only estimates over the main
    story in either mode. ``readability_words`` is its own count, not ``words``.
    Paragraphs, including paragraphs in nested tables and content wrappers, are
    measured separately. Adjacent runs concatenate without introducing spaces;
    tabs and line breaks separate words. Headers, footers, notes, comments,
    drawings/text boxes, deleted or moved-from content, field instructions, and ruby
    guides are excluded. Cached field results and ruby base text are counted once.
    Complex field instructions remain excluded until their separator, even if they
    span paragraphs.

    Readability words match ``[A-Za-z]+(?:['\\u2019][A-Za-z]+)*``: ASCII alphabetic
    tokens optionally joined by interior straight or typographic apostrophes.
    Syllables are contiguous lowercase ``[aeiouy]`` groups, minus one for a trailing
    silent ``e`` when there is more than one group, except for a consonant-plus-``le``
    ending; every word contributes at least one syllable. Sentences are nonempty
    (after whitespace stripping) fragments split on ``[.?!]+`` and paragraph
    boundaries, including a trailing unpunctuated fragment.

    These are approximations, not linguistic analysis. With ``W`` readability words,
    ``S`` sentences, and ``Y`` syllables, Flesch reading ease is
    ``206.835 - 1.015 * (W / S) - 84.6 * (Y / W)`` and Flesch-Kincaid grade is
    ``0.39 * (W / S) + 11.8 * (Y / W) - 15.59``. Neither is rounded or clamped.
    Empty documents have zero counts and ``None`` scores; wordless nonempty text
    retains its sentence count but also has ``None`` scores.

    :meth:`to_dict` is the ``wolfdocx inspect --statistics`` payload; it reports the
    story, count, and layout fields and leaves the readability fields out.
    """

    words: int
    characters: int
    characters_with_spaces: int
    paragraphs: int
    readability_words: int
    sentences: int
    syllables: int
    flesch_reading_ease: float | None
    flesch_kincaid_grade: float | None
    lines: int | None = None
    pages: int | None = None
    stories: str = "main"

    def to_dict(self) -> dict[str, int | str | None]:
        return {
            "stories": self.stories,
            "words": self.words,
            "characters": self.characters,
            "characters_with_spaces": self.characters_with_spaces,
            "paragraphs": self.paragraphs,
            "lines": self.lines,
            "pages": self.pages,
        }


def _statistics_hidden(run: ET.Element, ns: str) -> bool:
    properties = run.find(Q(ns, "rPr"))
    if properties is None:
        return False
    vanish = properties.find(Q(ns, "vanish"))
    if vanish is None:
        return False
    return vanish.get(Q(ns, "val"), "true") not in {"0", "false", "off"}


def _statistics_text(paragraph: ET.Element, ns: str) -> str:
    """The counted text of one paragraph, without any nested paragraph's text."""

    pieces: list[str] = []
    fallback = Q(_MARKUP_COMPATIBILITY_NS, "Fallback")
    skipped = {Q(ns, name) for name in _STATISTICS_SKIPPED} | {fallback, Q(ns, "p")}
    silent = {Q(ns, name) for name in _STATISTICS_SILENT}
    text_tags = {Q(ns, "t"), Q(_MATH_NS, "t")}
    run_tag = Q(ns, "r")
    tab_tags = {Q(ns, "tab"), Q(ns, "ptab")}
    break_tags = {Q(ns, "br"), Q(ns, "cr")}
    symbol_tags = {Q(ns, "sym"), Q(ns, "noBreakHyphen")}

    def walk(element: ET.Element) -> None:
        for child in element:
            tag = child.tag
            if tag in skipped or tag in silent:
                continue
            if tag == run_tag and _statistics_hidden(child, ns):
                continue
            if tag in text_tags:
                pieces.append(child.text or "")
            elif tag in tab_tags:
                pieces.append("\t")
            elif tag in break_tags:
                pieces.append("\n")
            elif tag in symbol_tags:
                pieces.append("-" if tag == Q(ns, "noBreakHyphen") else "\ufffc")
            walk(child)

    walk(paragraph)
    return "".join(pieces)


def _statistics_paragraphs(body: ET.Element, ns: str) -> list[str]:
    fallback = Q(_MARKUP_COMPATIBILITY_NS, "Fallback")
    skipped = {Q(ns, name) for name in _STATISTICS_SKIPPED} | {fallback}
    paragraph_tag = Q(ns, "p")
    texts: list[str] = []

    def visit(element: ET.Element) -> None:
        for child in element:
            if child.tag in skipped:
                continue
            if child.tag == paragraph_tag:
                texts.append(_statistics_text(child, ns))
            visit(child)

    visit(body)
    return texts


def _ideograph(character: str) -> bool:
    return unicodedata.east_asian_width(character) in {"W", "F"} and (
        unicodedata.category(character).startswith("L")
    )


def _count_text(paragraphs: Iterable[str]) -> tuple[int, int, int, int]:
    words = characters = spaces = filled = 0
    for text in paragraphs:
        if text:
            filled += 1
        in_word = False
        for character in text:
            if character in _STATISTICS_SPACES or character in _STATISTICS_BREAKS:
                in_word = False
                if character != "\n" and character != "\r":
                    spaces += 1
                continue
            characters += 1
            if _ideograph(character):
                words += 1
                in_word = False
            elif not in_word:
                words += 1
                in_word = True
    return words, characters, characters + spaces, filled


# -- all-stories statistics --------------------------------------------------------

_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
# The viramas that join two consonants of an Indic script into one conjunct (UAX #29
# rule GB9c), and the Unicode blocks of the scripts whose consonants they join.
_INDIC_LINKERS = frozenset("\u094d\u09cd\u0acd\u0b4d\u0c4d\u0d4d")
_INDIC_BLOCKS = (
    (0x0900, 0x097F),
    (0x0980, 0x09FF),
    (0x0A80, 0x0AFF),
    (0x0B00, 0x0B7F),
    (0x0C00, 0x0C7F),
    (0x0D00, 0x0D7F),
)
_GRAPHEME_JOINERS = frozenset("\u200c\u200d")
# LibreOffice separates words at whitespace and at the en and em dash; the dashes
# still count as characters.
_ALL_STORIES_SPACES = frozenset(
    " \t\n\r\u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009"
    "\u200a\u202f\u205f\u3000"
)
_ALL_STORIES_DASHES = frozenset("\u2013\u2014")
_NOTE_KINDS = (
    ("footnote", "footnotes", "decimal"),
    ("endnote", "endnotes", "lowerRoman"),
)
_NOTE_SEPARATOR_TYPES = frozenset({"separator", "continuationSeparator", "continuationNotice"})
_CHICAGO_MARKS = "*\u2020\u2021\u00a7"
_ROMAN_NUMERALS = (
    (1000, "m"),
    (900, "cm"),
    (500, "d"),
    (400, "cd"),
    (100, "c"),
    (90, "xc"),
    (50, "l"),
    (40, "xl"),
    (10, "x"),
    (9, "ix"),
    (5, "v"),
    (4, "iv"),
    (1, "i"),
)


def _indic_consonant(character: str) -> bool:
    point = ord(character)
    if not any(low <= point <= high for low, high in _INDIC_BLOCKS):
        return False
    if unicodedata.category(character) != "Lo":
        return False
    name = unicodedata.name(character, "")
    return " LETTER " in name and "VOWEL" not in name


def _grapheme_extends(character: str) -> bool:
    return (
        unicodedata.category(character) in {"Mn", "Me", "Mc"}
        or character in _GRAPHEME_JOINERS
        or 0x1F3FB <= ord(character) <= 0x1F3FF
    )


def _grapheme_control(character: str) -> bool:
    category = unicodedata.category(character)
    return category in {"Cc", "Zl", "Zp"} or (
        category == "Cf" and character not in _GRAPHEME_JOINERS
    )


def _grapheme_joins(previous: str, character: str, linker_pending: bool) -> bool:
    """Whether ``character`` continues the grapheme cluster ``previous`` belongs to."""

    if previous == "\r" and character == "\n":
        return True
    if _grapheme_control(previous) or _grapheme_control(character):
        return False
    if _grapheme_extends(character):
        return True
    if previous == "\u200d" and unicodedata.category(character) == "So":
        return True
    return linker_pending and _indic_consonant(character)


def _grapheme_count(text: str) -> int:
    """Count extended grapheme clusters, the characters LibreOffice counts.

    Marks, joiners, and emoji modifiers extend the cluster before them, a zero-width
    joiner joins the pictograph after it, CR LF is one cluster, controls stand alone,
    and a virama between two consonants of an Indic script joins them into one
    conjunct (UAX #29 rule GB9c), so a Devanagari conjunct is one character.
    """

    count = 0
    previous: str | None = None
    in_conjunct = False
    linker_pending = False
    for character in text:
        joined = previous is not None and _grapheme_joins(previous, character, linker_pending)
        consonant = _indic_consonant(character)
        if not joined:
            count += 1
            in_conjunct = consonant
            linker_pending = False
        elif consonant:
            in_conjunct = True
            linker_pending = False
        elif character in _INDIC_LINKERS and in_conjunct:
            linker_pending = True
        previous = character
    return count


def _libreoffice_count(text: str) -> tuple[int, int, int]:
    """Words, characters, and characters with spaces of one text, as LibreOffice counts."""

    words = characters = 0
    start: int | None = None
    for index, character in enumerate(text):
        if character in _ALL_STORIES_SPACES or character in _ALL_STORIES_DASHES:
            if start is not None:
                words += 1
                characters += _grapheme_count(text[start:index])
                start = None
            if character in _ALL_STORIES_DASHES:
                characters += 1
        elif start is None:
            start = index
    if start is not None:
        words += 1
        characters += _grapheme_count(text[start:])
    return words, characters, _grapheme_count(text)


def _statistics_flag(element: ET.Element | None, ns: str) -> bool:
    return element is not None and element.get(Q(ns, "val"), "true") not in {
        "0",
        "false",
        "off",
    }


def _statistics_integer(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def _statistics_value(element: ET.Element | None, ns: str, name: str) -> str | None:
    child = element.find(Q(ns, name)) if element is not None else None
    return child.get(Q(ns, "val")) if child is not None else None


def _statistics_number(value: int, fmt: str) -> str:
    """Print one list or note counter; a format this count does not model prints digits."""

    if fmt in {"bullet", "none"}:
        return ""
    if fmt == "decimalZero":
        return f"{value:02d}"
    if fmt in {"lowerLetter", "upperLetter"} and value > 0:
        letters = chr(ord("a") + (value - 1) % 26) * ((value - 1) // 26 + 1)
        return letters.upper() if fmt == "upperLetter" else letters
    if fmt in {"lowerRoman", "upperRoman"} and 0 < value < 4000:
        numerals: list[str] = []
        remaining = value
        for amount, numeral in _ROMAN_NUMERALS:
            while remaining >= amount:
                numerals.append(numeral)
                remaining -= amount
        roman = "".join(numerals)
        return roman.upper() if fmt == "upperRoman" else roman
    if fmt == "chicago" and value > 0:
        return _CHICAGO_MARKS[(value - 1) % 4] * ((value - 1) // 4 + 1)
    if fmt == "ordinal":
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(value % 10, "th")
        return f"{value}{'th' if 10 <= value % 100 <= 20 else suffix}"
    return str(value)


@dataclass(frozen=True, slots=True)
class _StatisticsLevel:
    fmt: str
    text: str | None
    start: int
    restart: int | None
    legal: bool
    style: str | None


class _StatisticsNumbering:
    """The list labels LibreOffice counts with the text of the paragraphs they number."""

    def __init__(self, numbering: ET.Element | None, styles: ET.Element | None, ns: str) -> None:
        self._ns = ns
        self._abstract: dict[str, dict[int, _StatisticsLevel]] = {}
        self._links: dict[str, str] = {}
        self._instances: dict[int, tuple[str, dict[int, tuple[int | None, Any]]]] = {}
        self._styles: dict[str, tuple[str | None, int | None, int | None]] = {}
        self._numbering_styles: dict[str, int] = {}
        self._default_style: str | None = None
        self._counters: dict[str, dict[int, int]] = {}
        for style in styles.findall(Q(ns, "style")) if styles is not None else ():
            identifier = style.get(Q(ns, "styleId"))
            if not identifier:
                continue
            num_pr = style.find(f"{Q(ns, 'pPr')}/{Q(ns, 'numPr')}")
            num_id = _statistics_integer(_statistics_value(num_pr, ns, "numId"))
            if style.get(Q(ns, "type")) == "numbering" and num_id is not None:
                self._numbering_styles[identifier] = num_id
            if style.get(Q(ns, "type")) != "paragraph":
                continue
            based = _statistics_value(style, ns, "basedOn")
            ilvl = _statistics_integer(_statistics_value(num_pr, ns, "ilvl"))
            self._styles[identifier] = (based, num_id, ilvl)
            if style.get(Q(ns, "default")) in {"1", "true", "on"}:
                self._default_style = identifier
        if numbering is None:
            return
        for abstract in numbering.findall(Q(ns, "abstractNum")):
            identifier = abstract.get(Q(ns, "abstractNumId"))
            if identifier is None:
                continue
            link = _statistics_value(abstract, ns, "numStyleLink")
            if link:
                self._links[identifier] = link
            levels = (self._level(element) for element in abstract.findall(Q(ns, "lvl")))
            self._abstract[identifier] = dict(level for level in levels if level is not None)
        for instance in numbering.findall(Q(ns, "num")):
            num_id = _statistics_integer(instance.get(Q(ns, "numId")))
            abstract_id = _statistics_value(instance, ns, "abstractNumId")
            if num_id is None or abstract_id is None:
                continue
            overrides: dict[int, tuple[int | None, Any]] = {}
            for override in instance.findall(Q(ns, "lvlOverride")):
                index = _statistics_integer(override.get(Q(ns, "ilvl")))
                if index is None:
                    continue
                start = _statistics_integer(_statistics_value(override, ns, "startOverride"))
                replacement = override.find(Q(ns, "lvl"))
                level = self._level(replacement) if replacement is not None else None
                overrides[index] = (start, level[1] if level is not None else None)
            self._instances[num_id] = (abstract_id, overrides)

    def _level(self, element: ET.Element) -> tuple[int, _StatisticsLevel] | None:
        ns = self._ns
        index = _statistics_integer(element.get(Q(ns, "ilvl")))
        if index is None:
            return None
        start = _statistics_integer(_statistics_value(element, ns, "start"))
        return index, _StatisticsLevel(
            fmt=(_statistics_value(element, ns, "numFmt") or "decimal").strip(),
            text=_statistics_value(element, ns, "lvlText"),
            start=start if start is not None else 1,
            restart=_statistics_integer(_statistics_value(element, ns, "lvlRestart")),
            legal=_statistics_flag(element.find(Q(ns, "isLgl")), ns),
            style=_statistics_value(element, ns, "pStyle"),
        )

    def _abstract_levels(self, abstract_id: str) -> dict[int, _StatisticsLevel]:
        seen = {abstract_id}
        while abstract_id in self._links:
            num_id = self._numbering_styles.get(self._links[abstract_id])
            instance = self._instances.get(num_id) if num_id is not None else None
            if instance is None or instance[0] in seen:
                break
            abstract_id = instance[0]
            seen.add(abstract_id)
        return self._abstract.get(abstract_id, {})

    def _level_of(self, num_id: int, ilvl: int) -> _StatisticsLevel | None:
        instance = self._instances.get(num_id)
        if instance is None:
            return None
        abstract_id, overrides = instance
        start, replacement = overrides.get(ilvl, (None, None))
        level = (
            replacement if replacement is not None else self._abstract_levels(abstract_id).get(ilvl)
        )
        if level is not None and start is not None:
            level = replace(level, start=start)
        return level

    def _style_reference(self, style_id: str | None) -> tuple[int | None, int | None]:
        num_id = ilvl = None
        seen: set[str] = set()
        while style_id and style_id not in seen and style_id in self._styles:
            seen.add(style_id)
            style_id, style_num_id, style_ilvl = self._styles[style_id]
            num_id = style_num_id if num_id is None else num_id
            ilvl = style_ilvl if ilvl is None else ilvl
        return num_id, ilvl

    def _label_text(self, num_id: int, level: _StatisticsLevel, counters: dict[int, int]) -> str:
        text = level.text or ""
        out: list[str] = []
        position = 0
        while position < len(text):
            character = text[position]
            if character == "%" and text[position + 1 : position + 2] in set("123456789"):
                index = int(text[position + 1]) - 1
                referenced = self._level_of(num_id, index)
                fmt = "decimal" if level.legal or referenced is None else referenced.fmt
                start = referenced.start if referenced is not None else 1
                out.append(_statistics_number(counters.get(index, start), fmt))
                position += 2
                continue
            out.append(character)
            position += 1
        return "".join(out)

    def label(self, paragraph: ET.Element) -> tuple[int, int, int] | None:
        """Count the label one paragraph prints, advancing its list's counters."""

        ns = self._ns
        properties = paragraph.find(Q(ns, "pPr"))
        style_id = _statistics_value(properties, ns, "pStyle") or self._default_style
        num_id, ilvl = self._style_reference(style_id)
        num_pr = properties.find(Q(ns, "numPr")) if properties is not None else None
        direct_id = _statistics_integer(_statistics_value(num_pr, ns, "numId"))
        direct_ilvl = _statistics_integer(_statistics_value(num_pr, ns, "ilvl"))
        num_id = direct_id if direct_id is not None else num_id
        ilvl = direct_ilvl if direct_ilvl is not None else ilvl
        instance = self._instances.get(num_id) if num_id else None
        if num_id is None or instance is None:
            return None
        if ilvl is None:
            levels = self._abstract_levels(instance[0])
            linked = [
                index for index, level in levels.items() if style_id and level.style == style_id
            ]
            ilvl = min(linked) if linked else 0
        level = self._level_of(num_id, ilvl)
        if level is None:
            return None
        # Numbering continues across the instances of one definition unless an
        # instance overrides its levels, which starts a list of its own.
        counters = self._counters.setdefault(instance[0] if not instance[1] else f"#{num_id}", {})
        counters[ilvl] = counters[ilvl] + 1 if ilvl in counters else level.start
        for deeper in [index for index in counters if index > ilvl]:
            deeper_level = self._level_of(num_id, deeper)
            restart = deeper_level.restart if deeper_level is not None else None
            if restart is None or (restart != 0 and ilvl <= restart - 1):
                del counters[deeper]
        if level.fmt == "bullet":
            # LibreOffice counts any bullet as one word of one character.
            return 1, 1, 1
        if level.fmt == "none":
            return None
        text = self._label_text(num_id, level, counters)
        return _libreoffice_count(text) if text else None


class _AllStoriesCounter:
    """Accumulates LibreOffice's statistics over the stories handed to it in order."""

    def __init__(
        self,
        ns: str,
        numbering: _StatisticsNumbering,
        settings: ET.Element | None,
        sections: list[ET.Element],
    ) -> None:
        self.ns = ns
        self.numbering = numbering
        self.words = self.characters = self.characters_with_spaces = self.paragraphs = 0
        self.references: dict[str, list[str]] = {kind: [] for kind, _plural, _fmt in _NOTE_KINDS}
        self._defaults: dict[str, tuple[str, int, str]] = {}
        for kind, _plural, fmt in _NOTE_KINDS:
            policy = settings.find(Q(ns, f"{kind}Pr")) if settings is not None else None
            start = _statistics_integer(_statistics_value(policy, ns, "numStart"))
            self._defaults[kind] = (
                _statistics_value(policy, ns, "numFmt") or fmt,
                start if start is not None else 1,
                _statistics_value(policy, ns, "numRestart") or "continuous",
            )
        self._formats: dict[str, tuple[str, int]] = {}
        self._marks: dict[str, int] = {kind: 0 for kind, _plural, _fmt in _NOTE_KINDS}
        self._sections = sections
        self._section = 0
        self._enter_section(sections[0] if sections else None)

    def _enter_section(self, section: ET.Element | None) -> None:
        # Marks that restart on every page are numbered per section: which page a
        # mark falls on needs a layout.
        for kind, (fmt, start, restart) in self._defaults.items():
            policy = section.find(Q(self.ns, f"{kind}Pr")) if section is not None else None
            own_start = _statistics_integer(_statistics_value(policy, self.ns, "numStart"))
            self._formats[kind] = (
                _statistics_value(policy, self.ns, "numFmt") or fmt,
                own_start if own_start is not None else start,
            )
            if (_statistics_value(policy, self.ns, "numRestart") or restart) != "continuous":
                self._marks[kind] = 0

    def _citation(self, kind: str, reference: ET.Element) -> str:
        identifier = reference.get(Q(self.ns, "id"))
        if identifier is not None:
            self.references[kind].append(identifier)
        if reference.get(Q(self.ns, "customMarkFollows")) in {"1", "true", "on"}:
            # The mark is the run text that follows, counted where it stands.
            return ""
        fmt, start = self._formats[kind]
        self._marks[kind] += 1
        return _statistics_number(start + self._marks[kind] - 1, fmt)

    def _pieces(self, paragraph: ET.Element) -> tuple[list[str], bool]:
        """A paragraph's counted text, split where a page or column break splits it,
        and whether it anchors an inline object."""

        ns = self.ns
        skipped = {Q(ns, name) for name in _STATISTICS_SKIPPED} | {
            Q(_MARKUP_COMPATIBILITY_NS, "Fallback"),
            Q(ns, "p"),
        }
        silent = {Q(ns, name) for name in _STATISTICS_SILENT}
        characters = {
            Q(ns, "tab"): "\t",
            Q(ns, "ptab"): "\t",
            Q(ns, "cr"): "\n",
            Q(ns, "sym"): "\ufffc",
            Q(ns, "noBreakHyphen"): "\u2011",
            Q(ns, "softHyphen"): "\u00ad",
        }
        notes = {
            Q(ns, "footnoteReference"): "footnote",
            Q(ns, "endnoteReference"): "endnote",
        }
        objects = {Q(_DRAWING_NS, "inline"), Q(ns, "object"), Q(_MATH_NS, "oMath")}
        pieces: list[list[str]] = [[]]
        anchored = False

        def walk(element: ET.Element) -> None:
            nonlocal anchored
            for child in element:
                tag = child.tag
                if tag in skipped or tag in silent:
                    continue
                if tag == Q(ns, "r") and _statistics_hidden(child, ns):
                    continue
                if tag == Q(ns, "t"):
                    pieces[-1].append(child.text or "")
                elif tag in characters:
                    pieces[-1].append(characters[tag])
                elif tag == Q(ns, "br"):
                    if child.get(Q(ns, "type")) in {"page", "column"}:
                        pieces.append([])
                    else:
                        pieces[-1].append("\n")
                elif tag in notes:
                    pieces[-1].append(self._citation(notes[tag], child))
                elif tag in objects or (tag == Q(ns, "pict") and not _vml_floating(child)):
                    # Formulas become objects whose text the statistics do not read.
                    anchored = True
                    if tag == Q(_MATH_NS, "oMath"):
                        continue
                walk(child)

        walk(paragraph)
        return ["".join(piece) for piece in pieces], anchored

    def paragraph(self, paragraph: ET.Element) -> None:
        ns = self.ns
        pieces, anchored = self._pieces(paragraph)
        counted = [piece for piece in pieces if piece]
        properties = paragraph.find(Q(ns, "pPr"))
        mark = properties.find(Q(ns, "rPr")) if properties is not None else None
        if counted or not _statistics_flag(
            mark.find(Q(ns, "vanish")) if mark is not None else None, ns
        ):
            label = self.numbering.label(paragraph)
        else:
            label = None
        for text in counted:
            words, characters, with_spaces = _libreoffice_count(text)
            self.words += words
            self.characters += characters
            self.characters_with_spaces += with_spaces
        if label is not None:
            self.words += label[0]
            self.characters += label[1]
            self.characters_with_spaces += label[2]
        self.paragraphs += len(counted) or int(label is not None or anchored)
        section = properties.find(Q(ns, "sectPr")) if properties is not None else None
        if (
            section is not None
            and self._section + 1 < len(self._sections)
            and self._sections[self._section] is section
        ):
            self._section += 1
            self._enter_section(self._sections[self._section])

    def story(self, root: ET.Element) -> None:
        """Count every paragraph of one story, its text boxes included, in order."""

        fallback = Q(_MARKUP_COMPATIBILITY_NS, "Fallback")
        paragraph_tag = Q(self.ns, "p")

        def visit(element: ET.Element) -> None:
            for child in element:
                if child.tag == fallback:
                    continue
                if child.tag == paragraph_tag:
                    self.paragraph(child)
                visit(child)

        visit(root)


def _vml_floating(element: ET.Element) -> bool:
    return any(
        "position:absolute" in (node.get("style") or "").replace(" ", "") for node in element.iter()
    )


def _statistics_sections(body: ET.Element, ns: str) -> list[ET.Element]:
    """The section properties of the body, in document order."""

    sections: list[ET.Element] = []
    for child in body:
        if child.tag == Q(ns, "sectPr"):
            sections.append(child)
        elif child.tag == Q(ns, "p"):
            section = child.find(f"{Q(ns, 'pPr')}/{Q(ns, 'sectPr')}")
            if section is not None:
                sections.append(section)
    return sections


def _section_header_footer_parts(
    sections: list[ET.Element], targets: Mapping[str, str], ns: str
) -> list[str]:
    """The header and footer parts each section shows, one entry per section.

    A section without its own reference of a type shows the previous section's.
    LibreOffice imports the first-page and even-page parts into the section's page
    style whether or not ``w:titlePg`` or ``w:evenAndOddHeaders`` turns them on, and
    counts their text, so every variant a section references counts.
    """

    kinds = {Q(ns, "headerReference"): "header", Q(ns, "footerReference"): "footer"}
    shown: dict[tuple[str, str], str] = {}
    parts: list[str] = []
    for section in sections:
        for reference in section:
            kind = kinds.get(reference.tag)
            if kind is None:
                continue
            identifier = next(
                (value for name, value in reference.attrib.items() if name.endswith("}id")),
                None,
            )
            target = targets.get(identifier) if identifier is not None else None
            if target is not None:
                shown[(kind, reference.get(Q(ns, "type"), "default"))] = target
        variants = ("default", "first", "even")
        for kind in ("header", "footer"):
            parts.extend(shown[(kind, variant)] for variant in variants if (kind, variant) in shown)
    return parts


def _all_stories_counts(
    graph: Any, document_part: str, body: ET.Element, ns: str
) -> tuple[int, int, int, int]:
    relationships = [
        relationship
        for relationship in graph.relationships_for(document_part)
        if not relationship.is_external
    ]
    targets = {relationship.id: relationship.resolved_target() for relationship in relationships}
    by_type: dict[str, str] = {}
    for relationship in relationships:
        by_type.setdefault(relationship.type.rsplit("/", 1)[-1], relationship.resolved_target())
    budget = [0]
    parsed: dict[str, ET.Element | None] = {}

    def part(name: str | None) -> ET.Element | None:
        if name is None or not graph.has_part(name):
            return None
        if name not in parsed:
            parsed[name] = _parse_bounded_xml(
                graph.read_part_bytes(name),
                limits=PackageLimits(),
                total_elements=budget,
            )
        return parsed[name]

    settings = part(by_type.get("settings"))
    numbering = _StatisticsNumbering(
        part(by_type.get("numbering")), part(by_type.get("styles")), ns
    )
    sections = _statistics_sections(body, ns)
    counter = _AllStoriesCounter(ns, numbering, settings, sections)
    counter.story(body)
    for name in _section_header_footer_parts(sections, targets, ns):
        story = part(name)
        if story is not None:
            counter.story(story)
    for kind, plural, _fmt in _NOTE_KINDS:
        notes_root = part(by_type.get(plural))
        if notes_root is None:
            continue
        notes = {
            note.get(Q(ns, "id")): note
            for note in notes_root.findall(Q(ns, kind))
            if note.get(Q(ns, "type"), "normal") not in _NOTE_SEPARATOR_TYPES
        }
        for identifier in list(counter.references[kind]):
            note = notes.get(identifier)
            if note is not None:
                counter.story(note)
    return (
        counter.words,
        counter.characters,
        counter.characters_with_spaces,
        counter.paragraphs,
    )





def document_statistics(
    source: str | Path,
    *,
    stories: str = "main",
    
) -> DocumentStatistics:
    """Count text statistics without rendering or laying out the document.

    ``stories="main"`` counts Word's main story (see :class:`DocumentStatistics`).
    ``stories="all"`` counts the way LibreOffice's document statistics do: the main
    story; the headers and footers every section shows, including the ones it
    inherits from the section before it, the first-page ones under ``w:titlePg``, and
    the even-page ones under ``w:evenAndOddHeaders``; the footnotes and endnotes the
    stories reference, with their citation marks; and text boxes. Comments, deleted
    revisions, field instructions, hidden runs, and the text of formulas are left out.
    A list label counts with its paragraph and a bullet as one word of one character;
    a paragraph counts once it holds text, a label, or an inline object, and a page or
    column break splits it. Words are maximal runs between whitespace and en or em
    dashes, and characters are extended grapheme clusters, so a Devanagari conjunct
    is one character. Note marks that restart on every page are numbered per section,
    since the page a mark falls on needs a layout.


    The readability fields always measure the main story (see
    :class:`DocumentStatistics`). A missing, unreadable, or unsafe package raises
    ``ValueError("document_statistics_source_invalid")``; a document part without a
    WordprocessingML body raises ``ValueError("statistics_document_invalid")``.
    """

    from .package import PackageError, PackageGraph

    if stories not in {"main", "all"}:
        raise ValueError("statistics_stories_invalid")
    path = Path(source)
    document_part = "word/document.xml"
    try:
        graph = PackageGraph.from_file(path)
        data = graph.read_part_bytes(document_part)
        root = _parse_bounded_xml(data, limits=PackageLimits(), total_elements=[0])
    except (OSError, PackageError, _XmlSafetyError, ET.ParseError):
        raise ValueError("document_statistics_source_invalid") from None
    ns = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else ""
    if ns not in SUPPORTED_WORD_NAMESPACES:
        raise ValueError("statistics_document_invalid")
    body = root.find(Q(ns, "body"))
    if body is None:
        raise ValueError("statistics_document_invalid")
    if stories == "all":
        words, characters, with_spaces, paragraphs = _all_stories_counts(
            graph, document_part, body, ns
        )
    else:
        words, characters, with_spaces, paragraphs = _count_text(_statistics_paragraphs(body, ns))
    readability = _readability_counts(body, ns)
    return DocumentStatistics(
        words, characters, with_spaces, paragraphs, *readability, stories=stories
    )
