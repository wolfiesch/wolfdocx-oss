"""Source-bound native Word tracked-text authoring and bounded review decisions.

The editable surface is derived from the package relationships: the document
body, the header and footer parts that ``word/document.xml`` actually
references, and the relationship-reachable footnote and endnote parts. Every
target, revision, and published result is qualified by its exact source part,
so identical numeric revision IDs in different parts stay independently
selectable. Paragraph-mark revision groups, which Word writes when a whole
paragraph is inserted or deleted, are one logical revision shared by the
paragraph mark and its runs; a decision on that ID always applies to the whole
group or is refused, never to a structure-corrupting half of it. A named move,
whose ``moveFrom`` and ``moveTo`` halves are linked by range markers carrying one
``w:name``, is likewise one logical revision selected by its ``moveFrom`` ID.

Every record ``discover_review_revisions`` returns is accepted decision input:
``ReviewRevisionRef.from_discovery`` (or the record itself in a JSON decision
request) binds the decision to that exact record, so a reference whose revision
no longer matches the discovered one is refused as stale instead of applied.

Section, table-exception, cell and grid snapshots use the same accept-current /
reject-prior contract as run and paragraph properties. Cell decisions retain a
valid row; batch deletions that would empty one refuse before publication.
Numbering acceptance removes its history marker; rejection refuses because
``original`` contains a rendered label rather than prior numbering properties.
Custom XML decisions cover paired marker-only ranges. Wrapper-bearing, nested,
cross-container and unpaired ranges refuse rather than infer XML topology.
Discovery also reports the relationship-bound tracking/view settings; all
authoring and decision operations preserve that part's original bytes.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import tempfile
import zipfile
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .canonical import deterministic_zip_info
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    _inspect_docx,
    _Inspection,
    _require_verify_mode,
    _verify_changed_parts,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .reporting import _supported_part

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_CONFLICT_REVISION_TAGS = frozenset({Q(_W14, "conflictIns"), Q(_W14, "conflictDel")})
_CUSTOM_XML_CONFLICT_TAGS = frozenset(
    Q(_W14, f"customXmlConflict{kind}Range{boundary}")
    for kind in ("Ins", "Del")
    for boundary in ("Start", "End")
)
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DOCUMENT = "word/document.xml"
_HEADER_REFERENCE = "headerReference"
_FOOTER_REFERENCE = "footerReference"
_PROPERTY_CHANGE_TAGS = frozenset(
    {
        "rPrChange",
        "pPrChange",
        "tblPrChange",
        "trPrChange",
        "sectPrChange",
        "tblPrExChange",
        "tcPrChange",
        "tblGridChange",
    }
)
_CELL_REVISION_TAGS = frozenset({"cellIns", "cellDel", "cellMerge"})
_CUSTOM_XML_STARTS = frozenset(
    f"customXml{kind}RangeStart" for kind in ("Ins", "Del", "MoveFrom", "MoveTo")
)
_SELECTABLE_REVISION_TAGS = (
    frozenset({"ins", "del", "numberingChange"})
    | _PROPERTY_CHANGE_TAGS
    | _CELL_REVISION_TAGS
    | _CUSTOM_XML_STARTS
)
_MOVE_TAGS = frozenset({"moveFrom", "moveTo"})
_DISCOVERY_REQUIRED_KEYS = frozenset(
    {
        "part",
        "region",
        "revision_id",
        "kind",
        "author",
        "date",
        "paragraph",
        "text",
        "decidable",
        "decision_code",
    }
)
_DISCOVERY_OPTIONAL_KEYS = frozenset(
    {
        "paragraph_mark",
        "note_id",
        "note_type",
        "move_name",
        "move_to_revision_id",
        "move_to_paragraph",
    }
)
# Decision eligibility is re-derived from the live source; it is not identity.
_BINDING_EXCLUDED_KEYS = frozenset({"decidable", "decision_code"})
_UNSUPPORTED_REVISION_TAGS = frozenset(
    {
        "moveFrom",
        "moveTo",
        "cellIns",
        "cellDel",
        "cellMerge",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart",
        "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart",
        "customXmlMoveToRangeEnd",
    }
)
_UNSUPPORTED_REVISION_CHANGE_SUFFIX = "Change"
_REVISION_INTERSECTION_TAGS = (
    frozenset(
        {
            "commentRangeStart",
            "commentRangeEnd",
            "commentReference",
            "ins",
            "del",
            "moveFrom",
            "moveTo",
        }
    )
    | _PROPERTY_CHANGE_TAGS
)
_REVISION_ANCESTOR_TAGS = frozenset({"ins", "del", "moveFrom", "moveTo"}) | _PROPERTY_CHANGE_TAGS
_PROPERTY_CONTEXT_TAGS = frozenset(
    {
        "pPr",
        "rPr",
        "sectPr",
        "sdtEndPr",
        "sdtPr",
        "tblPr",
        "tblPrEx",
        "tcPr",
        "trPr",
    }
)
_REVISION_CONTEXT_TAGS = _REVISION_ANCESTOR_TAGS | _PROPERTY_CONTEXT_TAGS
_NOTE_KINDS: dict[str, tuple[str, str]] = {
    "footnote": ("footnotes", "footnote"),
    "endnote": ("endnotes", "endnote"),
}
_NOTE_TAGS = frozenset(element for _, element in _NOTE_KINDS.values())


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _mask_text(element: ET.Element, namespace: str) -> str:
    clone = copy.deepcopy(element)
    for node in clone.iter(Q(namespace, "t")):
        node.text = ""
    return hashlib.sha256(ET.tostring(clone, encoding="utf-8")).hexdigest()


def _word_namespace(root: ET.Element) -> str | None:
    if not root.tag.startswith("{"):
        return None
    namespace = root.tag[1:].split("}", 1)[0]
    return namespace if namespace in SUPPORTED_WORD_NAMESPACES else None


def _sha(value: Any, label: str) -> str:
    if (
        not isinstance(value, str)
        or len(value) != 64
        or any(c not in "0123456789abcdefABCDEF" for c in value)
    ):
        raise ValueError(f"{label}_invalid")
    return value.lower()


def _positive(value: Any, label: str, *, zero: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < (0 if zero else 1):
        raise ValueError(f"{label}_invalid")
    return value


def _metadata(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value or len(value) > 256:
        raise ValueError(f"{label}_invalid")
    return value


def _review_part_name(value: Any, label: str) -> str:
    """A package part path is validated as a name only; the package relationship gate decides support."""

    name = _metadata(value, label)
    if name.startswith("/") or "\\" in name or ".." in name.split("/") or not name.endswith(".xml"):
        raise ValueError(f"{label}_invalid")
    return name


def _date(value: Any) -> str:
    text = _metadata(value, "review_date")
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError as error:
        raise ValueError("review_date_invalid") from error
    if parsed.tzinfo is None:
        raise ValueError("review_date_timezone_required")
    return text


@dataclass(frozen=True)
class ReviewChange:
    part: str
    paragraph: int
    start: int
    end: int
    text: str
    structure_fingerprint_sha256: str


@dataclass(frozen=True)
class ReviewRevisionRef:
    """One part-qualified revision a decision names.

    ``revision_sha256`` optionally binds the reference to one discovered record;
    build it with ``from_discovery`` from a ``discover_review_revisions`` record.
    """

    part: str
    revision_id: int
    revision_sha256: str | None = None

    @classmethod
    def from_discovery(cls, record: Mapping[str, Any]) -> ReviewRevisionRef:
        """Bind one ``discover_review_revisions`` record as decision input.

        The record must carry the discovery keys unchanged. Its identity (part,
        region, ID, kind, author, date, paragraph, text, and any group, note, or
        move keys) is fingerprinted, and ``decide_review`` refuses the reference
        with ``review_revision_target_stale`` when the source no longer holds that
        same revision. Eligibility is re-evaluated at decision time, so an
        undecidable record is refused there with its live ``decision_code``.
        """

        if (
            not isinstance(record, Mapping)
            or not _DISCOVERY_REQUIRED_KEYS <= set(record)
            or not set(record) <= _DISCOVERY_REQUIRED_KEYS | _DISCOVERY_OPTIONAL_KEYS
        ):
            raise ValueError("review_revision_reference_invalid")
        return cls(
            _review_part_name(record["part"], "review_part"),
            _positive(record["revision_id"], "review_revision_id"),
            _revision_binding_sha256(record),
        )

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"part": self.part, "revision_id": self.revision_id}
        if self.revision_sha256 is not None:
            payload["revision_sha256"] = self.revision_sha256
        return payload


def _revision_binding_sha256(record: Mapping[str, Any]) -> str:
    """Fingerprint the identity of one discovered revision record."""

    identity = {key: value for key, value in record.items() if key not in _BINDING_EXCLUDED_KEYS}
    try:
        encoded = json.dumps(identity, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    except (TypeError, ValueError) as error:
        raise ValueError("review_revision_reference_invalid") from error
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class ReviewRequest:
    source: Path
    source_sha256: str
    output: Path
    author: str
    date: str
    changes: tuple[ReviewChange, ...]


@dataclass(frozen=True)
class ReviewDecisionRequest:
    """One accept or reject decision over named revisions or one author's revisions.

    Name revisions exactly through ``revisions``, or leave it empty and set
    ``author_filter`` to decide every discoverable revision attributed to that
    author across the body, referenced headers and footers, and note parts. An
    author-scoped decision is all-or-nothing: it refuses without writing unless
    every revision markup element the package attributes to that author is
    decidable, and publishes only after the staged package holds none of them.
    """

    source: Path
    source_sha256: str
    output: Path
    action: str
    revisions: tuple[ReviewRevisionRef, ...] = ()
    author_filter: str | None = None


@dataclass(frozen=True)
class ReviewReport:
    status: str
    code: str | None
    source_sha256: str
    output_sha256: str | None
    revisions: list[dict[str, Any]]
    changed_targets: list[dict[str, Any]]

    @property
    def passed(self) -> bool:
        return self.status == "published"

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "output_sha256": self.output_sha256,
            "revisions": self.revisions,
            "changed_targets": self.changed_targets,
        }


def _parse_change(value: Any) -> ReviewChange:
    if not isinstance(value, Mapping) or set(value) != {
        "part",
        "paragraph",
        "start",
        "end",
        "text",
        "structure_fingerprint_sha256",
    }:
        raise ValueError("review_change_invalid")
    paragraph = _positive(value["paragraph"], "review_paragraph")
    start, end = (
        _positive(value["start"], "review_start", zero=True),
        _positive(value["end"], "review_end", zero=True),
    )
    if end < start or not isinstance(value["text"], str) or len(value["text"]) > 32768:
        raise ValueError("review_change_range_invalid")
    if start == end and not value["text"]:
        raise ValueError("review_noop_invalid")
    return ReviewChange(
        _review_part_name(value["part"], "review_part"),
        paragraph,
        start,
        end,
        value["text"],
        _sha(value["structure_fingerprint_sha256"], "review_fingerprint"),
    )


def _parse_revision_ref(value: Any) -> ReviewRevisionRef:
    """Accept a bare reference, a bound reference, or one discovery record verbatim."""

    keys = set(value) if isinstance(value, Mapping) else set()
    if keys in ({"part", "revision_id"}, {"part", "revision_id", "revision_sha256"}):
        return ReviewRevisionRef(
            _review_part_name(value["part"], "review_part"),
            _positive(value["revision_id"], "review_revision_id"),
            _sha(value["revision_sha256"], "review_revision_sha256")
            if "revision_sha256" in value
            else None,
        )
    return ReviewRevisionRef.from_discovery(value)


def parse_review_request(value: Any) -> ReviewRequest:
    if not isinstance(value, Mapping) or set(value) != {
        "source",
        "source_sha256",
        "output",
        "author",
        "date",
        "changes",
    }:
        raise ValueError("review_request_invalid")
    if (
        not isinstance(value["source"], str)
        or not isinstance(value["output"], str)
        or not isinstance(value["changes"], list)
        or not value["changes"]
    ):
        raise ValueError("review_request_fields_invalid")
    changes = tuple(_parse_change(item) for item in value["changes"])
    if len({(change.part, change.paragraph) for change in changes}) != len(changes):
        raise ValueError("review_paragraph_overlap")
    return ReviewRequest(
        Path(value["source"]),
        _sha(value["source_sha256"], "review_source_sha256"),
        Path(value["output"]),
        _metadata(value["author"], "review_author"),
        _date(value["date"]),
        changes,
    )


def parse_review_decision_request(value: Any) -> ReviewDecisionRequest:
    required = {"source", "source_sha256", "output", "action"}
    if (
        not isinstance(value, Mapping)
        or not required <= set(value)
        or not set(value) <= required | {"revisions", "author_filter"}
        or not set(value) & {"revisions", "author_filter"}
    ):
        raise ValueError("review_decision_request_invalid")
    listed = value.get("revisions", [])
    if (
        not isinstance(value["source"], str)
        or not isinstance(value["output"], str)
        or value["action"] not in {"accept", "reject"}
        or not isinstance(listed, list)
        or (not listed and "author_filter" not in value)
    ):
        raise ValueError("review_decision_fields_invalid")
    author_filter = (
        _metadata(value["author_filter"], "review_author_filter")
        if "author_filter" in value
        else None
    )
    if author_filter is not None and listed:
        raise ValueError("review_decision_selection_ambiguous")
    revisions = tuple(_parse_revision_ref(item) for item in listed)
    if len({(item.part, item.revision_id) for item in revisions}) != len(revisions):
        raise ValueError("review_revision_ids_duplicate")
    return ReviewDecisionRequest(
        Path(value["source"]),
        _sha(value["source_sha256"], "review_source_sha256"),
        Path(value["output"]),
        value["action"],
        revisions,
        author_filter,
    )


def load_review_request(path: str | Path) -> ReviewRequest:
    with Path(path).open(encoding="utf-8") as file:
        return parse_review_request(json.load(file))


def load_review_decision_request(path: str | Path) -> ReviewDecisionRequest:
    with Path(path).open(encoding="utf-8") as file:
        return parse_review_decision_request(json.load(file))


@dataclass(frozen=True)
class _ReviewPart:
    """One editable Word part, qualified by the package relationship that reaches it."""

    part: str
    region: str
    region_type: str | None
    relationship_id: str | None

    def to_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "region": self.region,
            "region_type": self.region_type,
            "relationship_id": self.relationship_id,
        }


def _attribute_by_local_name(element: ET.Element, local_name: str) -> str | None:
    for name, value in element.attrib.items():
        if name.rsplit("}", 1)[-1] == local_name:
            return value
    return None


def _note_part(graph: PackageGraph, kind: str) -> tuple[str, str] | None:
    """Resolve one relationship-reachable note part and its declaring relationship ID.

    A note part is eligible only when exactly one non-external ``footnotes`` or
    ``endnotes`` relationship leaves ``word/document.xml``, its content type is
    the matching note type, and its root element is the matching note container.
    An arbitrary user path can never be selected.
    """

    plural, _ = _NOTE_KINDS[kind]
    content_type = f"application/vnd.openxmlformats-officedocument.wordprocessingml.{plural}+xml"
    matches = [
        relationship
        for relationship in graph.relationships_for(_DOCUMENT)
        if not relationship.is_external and relationship.type.rsplit("/", 1)[-1] == plural
    ]
    if len(matches) != 1:
        return None
    target = matches[0].resolved_target()
    package_part = graph.part(target)
    if package_part is None or package_part.content_type != content_type:
        return None
    try:
        root = ET.fromstring(graph.read_part_bytes(target))
    except (ET.ParseError, PackageError):
        return None
    namespace = _word_namespace(root)
    if namespace is None or root.tag != Q(namespace, plural):
        return None
    return target, matches[0].id


def _review_parts(
    graph: PackageGraph, document_root: ET.Element, namespace: str
) -> dict[str, _ReviewPart]:
    """Editable parts reachable from the document: body, referenced regions, and notes.

    Membership is decided by ``reporting._supported_part`` for the body, headers,
    and footers so the review surface shares one allowlist definition with report
    authoring: a header or footer is eligible only when a non-external document
    relationship resolves to it and ``word/document.xml`` holds the matching
    ``w:headerReference`` or ``w:footerReference``. Footnotes and endnotes are
    eligible only through ``_note_part``. An arbitrary user path can never be
    selected.
    """

    parts: dict[str, _ReviewPart] = {
        _DOCUMENT: _ReviewPart(_DOCUMENT, "body", None, None),
    }
    relationships = {
        relationship.id: relationship
        for relationship in graph.relationships_for(_DOCUMENT)
        if not relationship.is_external
    }
    for element in document_root.iter():
        local_name = _local_name(element.tag)
        if local_name not in {_HEADER_REFERENCE, _FOOTER_REFERENCE}:
            continue
        identifier = _attribute_by_local_name(element, "id")
        relationship = relationships.get(identifier) if identifier else None
        if relationship is None:
            continue
        target = relationship.resolved_target()
        if not _supported_part(graph, target):
            continue
        region = "header" if local_name == _HEADER_REFERENCE else "footer"
        parts[target] = _ReviewPart(target, region, element.get(Q(namespace, "type")), identifier)
    for kind in sorted(_NOTE_KINDS):
        resolved = _note_part(graph, kind)
        if resolved is None:
            continue
        target, relationship_id = resolved
        parts[target] = _ReviewPart(target, kind, None, relationship_id)
    return parts


def _note_identity(note: ET.Element, namespace: str) -> dict[str, Any]:
    """The part-scoped note identity: numeric note ID plus any structural note type."""

    raw = note.get(Q(namespace, "id"))
    identifier = int(raw) if raw is not None and raw.lstrip("-").isdigit() else None
    return {"note_id": identifier, "note_type": note.get(Q(namespace, "type"))}


def _note_ancestor(node: ET.Element, parents: Mapping[int, ET.Element]) -> ET.Element | None:
    """The enclosing footnote or endnote element, when the node sits in a note part."""

    ancestor: ET.Element | None = node
    while ancestor is not None:
        if _local_name(ancestor.tag) in _NOTE_TAGS:
            return ancestor
        ancestor = parents.get(id(ancestor))
    return None


def _note_structural_code(
    node: ET.Element, parents: Mapping[int, ET.Element], namespace: str
) -> str | None:
    """Refuse notes that carry structure rather than user content.

    Separator, continuation-separator, and continuation-notice notes exist for
    note layout, not for authoring; their paragraphs and revisions are reported
    as unsupported instead of being selectable or editable.
    """

    note = _note_ancestor(node, parents)
    if note is None:
        return None
    return "review_revision_note_unsupported" if note.get(Q(namespace, "type")) else None


def _ordered_parts(parts: Mapping[str, _ReviewPart]) -> list[_ReviewPart]:
    return [parts[name] for name in sorted(parts)]


def _part_context(part: str, parts: Mapping[str, _ReviewPart]) -> dict[str, Any]:
    selected = parts.get(part)
    return {
        "part": part,
        "region": selected.region if selected is not None else None,
        "region_type": selected.region_type if selected is not None else None,
        "relationship_id": selected.relationship_id if selected is not None else None,
        "allowed_parts": [item.to_dict() for item in _ordered_parts(parts)],
    }


def discover_review_targets(source: str | Path, *, part: str = _DOCUMENT) -> dict[str, Any]:
    """List source-bound paragraph targets supported by native tracked editing.

    Read-only. ``part`` selects the body or one referenced header or footer; the
    response always names the complete allowed part set so callers can reach
    every supported region.
    """

    path = Path(source)
    try:
        graph = PackageGraph.from_file(path)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return {"passed": False, "code": "review_source_package_invalid", "targets": []}
    try:
        document_root = ET.fromstring(parts[_DOCUMENT])
    except (KeyError, ET.ParseError):
        return {"passed": False, "code": "review_source_package_invalid", "targets": []}
    namespace = _word_namespace(document_root)
    if namespace is None:
        return {"passed": False, "code": "review_namespace_unsupported", "targets": []}
    allowed = _review_parts(graph, document_root, namespace)
    if part not in allowed:
        return {
            "passed": False,
            "code": "review_part_unsupported",
            "source_sha256": graph.file_sha256,
            **_part_context(part, allowed),
            "targets": [],
        }
    try:
        root = ET.fromstring(parts[part])
    except (KeyError, ET.ParseError):
        return {"passed": False, "code": "review_source_package_invalid", "targets": []}
    part_namespace = _word_namespace(root)
    if part_namespace is None:
        return {"passed": False, "code": "review_namespace_unsupported", "targets": []}
    selected = allowed[part]
    parents = {id(child): parent for parent in root.iter() for child in parent}
    targets = _paragraph_target_records(part, root, part_namespace, selected, parents)
    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        **_part_context(part, allowed),
        "targets": targets,
    }


def _paragraph_target_records(
    part: str,
    root: ET.Element,
    namespace: str,
    selected: _ReviewPart,
    parents: Mapping[int, ET.Element],
) -> list[dict[str, Any]]:
    """One part's paragraph targets: the one bound description of an editable paragraph.

    The review and comment surfaces expose the same records, so a paragraph one
    surface offers is described identically by the other: the same part, the same
    one-based position, the same text-masked structure fingerprint, and the same
    editable, protected, and annotation ranges.
    """

    note_region = selected.region in _NOTE_KINDS
    targets: list[dict[str, Any]] = []
    for index, paragraph in enumerate(root.iter(Q(namespace, "p")), start=1):
        note = _note_ancestor(paragraph, parents) if note_region else None
        note_code = _note_structural_code(paragraph, parents, namespace) if note_region else None
        target = {
            "part": part,
            "region": selected.region,
            "paragraph": index,
            "structure_fingerprint_sha256": _mask_text(paragraph, namespace),
            **(_note_identity(note, namespace) if note is not None else {}),
        }
        if note_code is not None:
            target.update(
                {
                    "text": _visible_text(paragraph, namespace),
                    "editable_ranges": [],
                    "protected_ranges": [],
                    "annotation_ranges": [],
                    "context_code": note_code,
                }
            )
            targets.append(target)
            continue
        try:
            mapping = _paragraph_map(paragraph, namespace)
        except _ReviewContextError as error:
            target.update(
                {
                    "text": _visible_text(paragraph, namespace),
                    "editable_ranges": [],
                    "protected_ranges": [],
                    "annotation_ranges": [],
                    "context_code": error.code,
                }
            )
        else:
            target.update({"text": mapping.text, **_target_ranges(mapping)})
        targets.append(target)
    return targets


def _local_name(tag: Any) -> str:
    return tag.split("}", 1)[-1] if isinstance(tag, str) else ""


def _selectable_revision_id(identifier: Any) -> int | None:
    """Return the decision ID for a raw w:id, or None when a request cannot name it exactly."""

    if not isinstance(identifier, str) or not identifier.isascii() or not identifier.isdigit():
        return None
    value = int(identifier)
    return value if value >= 1 and str(value) == identifier else None


def _revision_intersection(node: ET.Element) -> bool:
    return any(
        descendant is not node
        and (
            _local_name(descendant.tag) in _REVISION_INTERSECTION_TAGS
            or descendant.tag in _CONFLICT_REVISION_TAGS
        )
        for descendant in node.iter()
    )


def _revision_ancestor(node: ET.Element, parents: Mapping[int, ET.Element]) -> bool:
    """True when a revision sits in a nested-revision or property/mark context."""

    parent = parents.get(id(node))
    while parent is not None:
        if (
            _local_name(parent.tag) in _REVISION_CONTEXT_TAGS
            or parent.tag in _CONFLICT_REVISION_TAGS
        ):
            return True
        parent = parents.get(id(parent))
    return False


def _revision_text(node: ET.Element, namespace: str, kind: str) -> str:
    """Text a decision applies: insertion text on accept, deleted text on reject."""

    tag = Q(namespace, "t" if kind in {"ins", "conflictIns"} else "delText")
    return "".join(child.text or "" for child in node.iter(tag))


def _property_change_structural_code(
    node: ET.Element,
    kind: str,
    namespace: str,
    parents: Mapping[int, ET.Element],
) -> str | None:
    """Validate container structure and ancestry for one property change revision."""

    container_tag = kind[:-6]
    parent = parents.get(id(node))
    if parent is None or _local_name(parent.tag) != container_tag:
        return "review_revision_context_unsupported"

    grandparent = parents.get(id(parent))
    if grandparent is None:
        return "review_revision_context_unsupported"

    gp_name = _local_name(grandparent.tag)
    expected_parent = {
        "rPr": {"r"},
        "pPr": {"p"},
        "tblPr": {"tbl"},
        "trPr": {"tr"},
        "sectPr": {"body", "pPr"},
        "tblPrEx": {"tr"},
        "tcPr": {"tc"},
        "tblGrid": {"tbl"},
    }.get(container_tag, set())
    if gp_name not in expected_parent:
        return "review_revision_context_unsupported"

    prior_nodes = {id(child) for child in node.iter()}
    sibling_changes = [
        child
        for child in parent.iter()
        if id(child) not in prior_nodes and _local_name(child.tag) in _SELECTABLE_REVISION_TAGS
    ]
    if sibling_changes:
        return "review_revision_context_unsupported"

    curr: ET.Element | None = parents.get(id(grandparent)) if gp_name == "pPr" else grandparent
    while curr is not None:
        name = _local_name(curr.tag)
        if (
            name in _REVISION_ANCESTOR_TAGS
            or name in _PROPERTY_CONTEXT_TAGS
            or curr.tag in _CONFLICT_REVISION_TAGS
        ):
            return "review_revision_context_unsupported"
        curr = parents.get(id(curr))

    if len(node) != 1 or node[0].tag != Q(namespace, container_tag):
        return "review_revision_target_malformed"

    return None


def _structure_revision_code(
    entry: _IndexedRevision, namespace: str, parents: Mapping[int, ET.Element]
) -> str | None:
    node = entry.node
    parent = parents.get(id(node))
    if entry.kind == "numberingChange":
        if parent is None or parent.tag != Q(namespace, "numPr"):
            return "review_revision_context_unsupported"
        # original stores a rendered label, not the former numId/ilvl properties.
        return None
    if entry.kind in _CELL_REVISION_TAGS:
        cell = parents.get(id(parent)) if parent is not None else None
        if parent is None or parent.tag != Q(namespace, "tcPr") or cell is None:
            return "review_revision_context_unsupported"
        if cell.tag != Q(namespace, "tc") or len(node):
            return "review_revision_target_malformed"
        if (
            sum(child.tag in {Q(namespace, k) for k in _CELL_REVISION_TAGS} for child in parent)
            != 1
        ):
            return "review_revision_context_unsupported"
        if entry.kind == "cellMerge" and any(
            node.get(Q(namespace, attr)) not in {None, "rest", "cont"}
            for attr in ("vMerge", "vMergeOrig")
        ):
            return "review_revision_target_malformed"
        return None
    end_tag = Q(namespace, entry.kind.removesuffix("Start") + "End")
    root = node
    while id(root) in parents:
        root = parents[id(root)]
    ends = [n for n in root.iter(end_tag) if n.get(Q(namespace, "id")) == entry.identifier]
    if len(ends) != 1 or parent is None or parents.get(id(ends[0])) is not parent:
        return "review_revision_range_structure_unsupported"
    siblings = list(parent)
    first, last = siblings.index(node), siblings.index(ends[0])
    if first >= last or any(
        descendant.tag == Q(namespace, "customXml")
        or _local_name(descendant.tag) in _SELECTABLE_REVISION_TAGS
        or "Range" in _local_name(descendant.tag)
        for child in siblings[first + 1 : last]
        for descendant in child.iter()
    ):
        # Arbitrary custom XML wrapper topology needs a semantic reconstruction.
        return "review_revision_range_structure_unsupported"
    return None


def _revision_decision_code(
    entry: _IndexedRevision,
    namespace: str,
    parents: Mapping[int, ET.Element],
    duplicates: frozenset[str],
) -> str | None:
    """The single reason an exact-ID decision could not act on only this revision."""

    if _selectable_revision_id(entry.identifier) is None:
        return "review_revision_id_unsupported"
    if entry.identifier in duplicates:
        return "review_revision_id_ambiguous"
    note_code = _note_structural_code(entry.node, parents, namespace)
    if note_code is not None:
        return note_code
    if entry.group is not None:
        return _group_decision_code(entry.group, parents)
    if entry.move is not None:
        return _move_decision_code(entry.move, namespace, parents, duplicates)
    node = entry.node
    if _revision_intersection(node):
        return "review_revision_intersection_unsupported"
    if entry.kind in _PROPERTY_CHANGE_TAGS:
        return _property_change_structural_code(node, entry.kind, namespace, parents)
    if entry.kind in _CELL_REVISION_TAGS | _CUSTOM_XML_STARTS | {"numberingChange"}:
        return _structure_revision_code(entry, namespace, parents)
    if _revision_ancestor(node, parents):
        return "review_revision_context_unsupported"
    if node.tag in _CONFLICT_REVISION_TAGS:
        # MS-DOCX 2.6.1.3 / 2.6.1.6 permit treating inline conflicts as
        # tracked del/ins. The identically named CT_TrackChange parent-object
        # variants (2.6.1.4 / 2.6.1.5) are not run decisions.
        parent = parents.get(id(node))
        if parent is None or parent.tag != Q(namespace, "p"):
            return "review_revision_context_unsupported"
        if _paragraph_mark_revision(parent, namespace) is not None:
            return "review_revision_context_unsupported"
        word_prefix = f"{{{namespace}}}"
        ancestor: ET.Element | None = parent
        while ancestor is not None:
            if not ancestor.tag.startswith(word_prefix) or _local_name(ancestor.tag) not in {
                "p",
                "tc",
                "tr",
                "tbl",
                "body",
                "document",
                "hdr",
                "ftr",
                "footnote",
                "footnotes",
                "endnote",
                "endnotes",
            }:
                return "review_revision_context_unsupported"
            ancestor = parents.get(id(ancestor))
    if _revision_visible_text(node, namespace, entry.kind) is None:
        return "review_revision_context_unsupported"
    return None


def _move_decision_code(
    pair: _MovePair,
    namespace: str,
    parents: Mapping[int, ET.Element],
    duplicates: frozenset[str],
) -> str | None:
    """Structural preconditions for deciding one named move as a single unit.

    Each half must be one wrapper of direct text runs, placed as a paragraph child
    immediately between its own range start and end markers, in a paragraph whose
    mark is not itself revised. Both halves must carry the same text, so accepting
    or rejecting the pair can never duplicate or lose content.
    """

    destination_id = pair.destination.get(Q(namespace, "id"), "")
    if _selectable_revision_id(destination_id) is None:
        return "review_revision_id_unsupported"
    if destination_id in duplicates:
        return "review_revision_id_ambiguous"
    note_code = _note_structural_code(pair.destination, parents, namespace)
    if note_code is not None:
        return note_code
    texts: list[str] = []
    for wrapper, (start, end), text_kinds in (
        (pair.source, pair.source_range, ("del", "ins")),
        (pair.destination, pair.destination_range, ("ins",)),
    ):
        paragraph = parents.get(id(wrapper))
        if paragraph is None or paragraph.tag != Q(namespace, "p"):
            return "review_revision_move_structure_unsupported"
        siblings = list(paragraph)
        position = siblings.index(wrapper)
        if (
            position == 0
            or siblings[position - 1] is not start
            or position + 1 >= len(siblings)
            or siblings[position + 1] is not end
        ):
            return "review_revision_move_structure_unsupported"
        marks = paragraph.find(f"{Q(namespace, 'pPr')}/{Q(namespace, 'rPr')}")
        if marks is not None and any(
            _local_name(mark.tag) in {"ins", "del", *_MOVE_TAGS} for mark in marks
        ):
            return "review_revision_move_structure_unsupported"
        if _revision_intersection(wrapper):
            return "review_revision_intersection_unsupported"
        if _revision_ancestor(wrapper, parents):
            return "review_revision_context_unsupported"
        text = next(
            (
                visible
                for kind in text_kinds
                if (visible := _revision_visible_text(wrapper, namespace, kind)) is not None
            ),
            None,
        )
        if text is None:
            return "review_revision_context_unsupported"
        texts.append(text)
    if texts[0] != texts[1]:
        return "review_revision_move_structure_unsupported"
    return None


def _apply_move_decision(
    pair: _MovePair, action: str, namespace: str, parents: Mapping[int, ET.Element]
) -> None:
    """Accept keeps the destination text; reject keeps the source text.

    The other half is removed with its content, the kept half is unwrapped in
    place, and all four range markers are removed, so no orphaned move remains.
    """

    keep, drop = (
        (pair.destination, pair.source) if action == "accept" else (pair.source, pair.destination)
    )
    for node in (*pair.source_range, *pair.destination_range, drop):
        parent = parents.get(id(node))
        if parent is not None:
            parent.remove(node)
    parent = parents.get(id(keep))
    if parent is None:
        return
    index = list(parent).index(keep)
    for child in list(keep):
        if keep is pair.source:
            for text_node in child.iter(Q(namespace, "delText")):
                text_node.tag = Q(namespace, "t")
        parent.insert(index, child)
        index += 1
    parent.remove(keep)


def _group_decision_code(group: _ParagraphGroup, parents: Mapping[int, ET.Element]) -> str | None:
    """Structural preconditions for acting on one whole-paragraph revision group.

    A group changes paragraph structure, so it is decidable only when the
    paragraph is a direct body item of a part that is not itself nested in a
    revision or property context. Anything else is refused rather than applied
    partially.
    """

    if any(node.tag in _CONFLICT_REVISION_TAGS for node in group.paragraph.iter()):
        return "review_revision_intersection_unsupported"
    parent = parents.get(id(group.paragraph))
    if parent is None or _local_name(parent.tag) != "body":
        return "review_revision_group_structure_unsupported"
    if any(parents.get(id(member)) is not group.paragraph for member in group.members):
        return "review_revision_group_structure_unsupported"
    ancestor = parents.get(id(parent))
    while ancestor is not None:
        if (
            _local_name(ancestor.tag) in _REVISION_CONTEXT_TAGS
            or ancestor.tag in _CONFLICT_REVISION_TAGS
        ):
            return "review_revision_context_unsupported"
        ancestor = parents.get(id(ancestor))
    return None


def _apply_group_decision(
    group: _ParagraphGroup,
    action: str,
    namespace: str,
    parents: Mapping[int, ET.Element],
) -> None:
    """Apply one whole-paragraph decision without splitting its structure.

    Accepting an inserted paragraph or rejecting a deleted one keeps the
    paragraph, unwraps every run wrapper the group owns, and removes the mark
    revision. The opposite actions remove the whole paragraph. Either way the
    paragraph mark and its runs move together.
    """

    keep_paragraph = (action == "accept" and group.kind == "ins") or (
        action == "reject" and group.kind == "del"
    )
    paragraph_parent = parents.get(id(group.paragraph))
    if paragraph_parent is None:
        return
    if not keep_paragraph:
        paragraph_parent.remove(group.paragraph)
        return
    for member in group.members:
        parent = parents.get(id(member))
        if parent is None:
            continue
        index = list(parent).index(member)
        for child in list(member):
            if group.kind == "del":
                for text_node in child.iter(Q(namespace, "delText")):
                    text_node.tag = Q(namespace, "t")
            parent.insert(index, child)
            index += 1
        parent.remove(member)
    run_properties = parents.get(id(group.mark))
    if run_properties is None:
        return
    run_properties.remove(group.mark)
    if len(run_properties) or run_properties.attrib:
        return
    properties = parents.get(id(run_properties))
    if properties is not None:
        properties.remove(run_properties)


@dataclass(frozen=True)
class _ParagraphGroup:
    """One whole-paragraph insertion or deletion: the paragraph mark plus its run wrappers."""

    paragraph: ET.Element
    mark: ET.Element
    kind: str
    revision_id: str
    members: tuple[ET.Element, ...]


@dataclass(frozen=True)
class _MovePair:
    """One named move: its moveFrom and moveTo wrappers and their range markers."""

    name: str
    source: ET.Element
    destination: ET.Element
    source_range: tuple[ET.Element, ET.Element]
    destination_range: tuple[ET.Element, ET.Element]


def _move_ranges(
    root: ET.Element, namespace: str, kind: str, order: Mapping[int, int]
) -> dict[str, tuple[ET.Element, ET.Element, ET.Element]]:
    """Named ranges of one move half, each enclosing exactly one wrapper of that half.

    A range qualifies only when its name and ID are unique, its end marker is
    unique, and no other qualifying range claims the same wrapper.
    """

    starts: dict[str, list[ET.Element]] = {}
    start_ids: dict[str, int] = {}
    for start in root.iter(Q(namespace, f"{kind}RangeStart")):
        starts.setdefault(start.get(Q(namespace, "name")) or "", []).append(start)
        identifier = start.get(Q(namespace, "id")) or ""
        start_ids[identifier] = start_ids.get(identifier, 0) + 1
    ends: dict[str, list[ET.Element]] = {}
    for end in root.iter(Q(namespace, f"{kind}RangeEnd")):
        ends.setdefault(end.get(Q(namespace, "id")) or "", []).append(end)
    wrappers = [(order[id(node)], node) for node in root.iter(Q(namespace, kind))]
    ranges: dict[str, tuple[ET.Element, ET.Element, ET.Element]] = {}
    claims: dict[int, int] = {}
    for name, named in starts.items():
        if not name or len(named) != 1:
            continue
        start = named[0]
        identifier = start.get(Q(namespace, "id")) or ""
        matched = ends.get(identifier, [])
        if not identifier or start_ids[identifier] != 1 or len(matched) != 1:
            continue
        low, high = order[id(start)], order[id(matched[0])]
        inside = [node for position, node in wrappers if low < position < high]
        if len(inside) != 1:
            continue
        ranges[name] = (inside[0], start, matched[0])
        claims[id(inside[0])] = claims.get(id(inside[0]), 0) + 1
    return {name: value for name, value in ranges.items() if claims[id(value[0])] == 1}


def _move_pairs(root: ET.Element, namespace: str) -> dict[int, _MovePair]:
    """Pair moveFrom and moveTo halves by range name, keyed by both wrappers.

    Halves without a uniquely named partner are not moves a decision can name;
    they remain reported as unsupported revision markup.
    """

    if next(root.iter(Q(namespace, "moveFrom")), None) is None:
        return {}
    order = {id(node): position for position, node in enumerate(root.iter())}
    sources = _move_ranges(root, namespace, "moveFrom", order)
    destinations = _move_ranges(root, namespace, "moveTo", order)
    pairs: dict[int, _MovePair] = {}
    for name in sorted(sources.keys() & destinations.keys()):
        source, source_start, source_end = sources[name]
        destination, destination_start, destination_end = destinations[name]
        pair = _MovePair(
            name,
            source,
            destination,
            (source_start, source_end),
            (destination_start, destination_end),
        )
        pairs[id(source)] = pair
        pairs[id(destination)] = pair
    return pairs


def _move_text(node: ET.Element, namespace: str) -> str:
    tags = {Q(namespace, "t"), Q(namespace, "delText")}
    return "".join(child.text or "" for child in node.iter() if child.tag in tags)


@dataclass(frozen=True)
class _IndexedRevision:
    """One logical tracked revision: a single node, a paragraph-mark group, or a move."""

    identifier: str
    kind: str
    node: ET.Element
    group: _ParagraphGroup | None
    text: str
    author: str | None
    date: str | None
    move: _MovePair | None = None


def _paragraph_mark_revision(paragraph: ET.Element, namespace: str) -> ET.Element | None:
    """The paragraph's own mark revision: a direct pPr/rPr insertion or deletion."""

    for properties in paragraph:
        if properties.tag != Q(namespace, "pPr"):
            continue
        for run_properties in properties:
            if run_properties.tag != Q(namespace, "rPr"):
                continue
            for node in run_properties:
                if node.tag in {Q(namespace, "ins"), Q(namespace, "del")}:
                    return node
    return None


def _paragraph_group(
    paragraph: ET.Element,
    namespace: str,
    shared_nodes: Mapping[str, tuple[ET.Element, ...]],
) -> _ParagraphGroup | None:
    """One coherent whole-paragraph insertion or deletion, or None when this is not one.

    A group exists only when the paragraph mark carries a usable revision ID, every
    node in the part sharing that ID is the mark or a direct child wrapper of this
    same paragraph, and every content child of the paragraph is one of those
    wrappers holding at least one run. Foreign, nested, mixed-kind, or partial
    markup therefore keeps the previous per-node behavior and stays individually
    undecidable instead of editable as a half-applied structure.
    """

    mark = _paragraph_mark_revision(paragraph, namespace)
    if mark is None:
        return None
    identifier = mark.get(Q(namespace, "id"))
    if not identifier or _selectable_revision_id(identifier) is None:
        return None
    kind = "ins" if mark.tag == Q(namespace, "ins") else "del"
    tag = Q(namespace, kind)
    content = [child for child in paragraph if child.tag != Q(namespace, "pPr")]
    members = tuple(child for child in content if child.tag == tag)
    if not members or len(members) != len(content):
        return None
    shared = shared_nodes.get(identifier, ())
    if len(shared) != len(members) + 1:
        return None
    if any(node is not mark and node not in content for node in shared):
        return None
    if any(not any(node.tag == Q(namespace, "r") for node in member.iter()) for member in members):
        return None
    return _ParagraphGroup(paragraph, mark, kind, identifier, members)


def _group_revision(group: _ParagraphGroup, namespace: str) -> _IndexedRevision:
    return _IndexedRevision(
        group.revision_id,
        group.kind,
        group.mark,
        group,
        "".join(_revision_text(member, namespace, group.kind) for member in group.members),
        group.mark.get(Q(namespace, "author")),
        group.mark.get(Q(namespace, "date")),
    )


def _part_revision_index(
    root: ET.Element, namespace: str
) -> tuple[
    tuple[_IndexedRevision, ...],
    dict[str, tuple[_IndexedRevision, ...]],
    frozenset[str],
]:
    """Index logical tracked revisions in one part's document order.

    A coherent paragraph-mark group collapses to the single record at its mark
    position; every other ``w:ins``/``w:del`` node, including markup with no usable
    ID, stays its own record. The identifier map is built from logical records, so
    an ID shared inside one group is not ambiguous while duplicate IDs that are not
    one coherent group stay ``review_revision_id_ambiguous``. A named move is one
    record at its ``moveFrom`` wrapper, selected by that ID; its ``moveTo`` ID takes
    part only in the ambiguity check.
    Inline conflict wrappers share the Word ID space and remain separate logical
    revisions; their context eligibility is decided by the same decision gate.
    Decimal-equivalent spellings collide only for ambiguity checks; exact-ID
    selection and unsupported-ID diagnostics retain their lexical semantics.
    """

    indexed_tags = {
        Q(namespace, tag) for tag in _SELECTABLE_REVISION_TAGS
    } | _CONFLICT_REVISION_TAGS
    shared_nodes: dict[str, list[ET.Element]] = {}
    for node in root.iter():
        if node.tag in indexed_tags:
            shared_nodes.setdefault(node.get(Q(namespace, "id"), ""), []).append(node)
    frozen_nodes = {key: tuple(value) for key, value in shared_nodes.items()}
    groups: dict[str, _ParagraphGroup] = {}
    for paragraph in root.iter(Q(namespace, "p")):
        group = _paragraph_group(paragraph, namespace, frozen_nodes)
        if group is not None:
            groups.setdefault(group.revision_id, group)
    grouped = {id(node) for group in groups.values() for node in (group.mark, *group.members)}
    moves = _move_pairs(root, namespace)
    ordered: list[_IndexedRevision] = []
    for node in root.iter():
        move = moves.get(id(node))
        if move is not None:
            if node is move.source:
                ordered.append(
                    _IndexedRevision(
                        node.get(Q(namespace, "id"), ""),
                        "move",
                        node,
                        None,
                        _move_text(node, namespace),
                        node.get(Q(namespace, "author")),
                        node.get(Q(namespace, "date")),
                        move,
                    )
                )
            continue
        if node.tag not in indexed_tags:
            continue
        if id(node) in grouped:
            group = groups.get(node.get(Q(namespace, "id"), ""))
            if group is not None and node is group.mark:
                ordered.append(_group_revision(group, namespace))
            continue
        kind = _local_name(node.tag)
        text = (
            _revision_text(node, namespace, kind)
            if kind in {"ins", "del", "conflictIns", "conflictDel"}
            else ""
        )
        ordered.append(
            _IndexedRevision(
                node.get(Q(namespace, "id"), ""),
                kind,
                node,
                None,
                text,
                node.get(Q(namespace, "author")),
                node.get(Q(namespace, "date")),
            )
        )
    by_identifier: dict[str, list[_IndexedRevision]] = {}
    identity_counts: dict[str, int] = {}
    for entry in ordered:
        by_identifier.setdefault(entry.identifier, []).append(entry)
        identifiers = [entry.identifier]
        if entry.move is not None:
            identifiers.append(entry.move.destination.get(Q(namespace, "id"), ""))
        for identifier in identifiers:
            identity = identifier.strip(" \t\r\n").removeprefix("+")
            if identity.isascii() and identity.isdigit():
                identity = identity.lstrip("0") or "0"
            else:
                identity = identifier
            identity_counts[identity] = identity_counts.get(identity, 0) + 1
    duplicates = frozenset(identifier for identifier, count in identity_counts.items() if count > 1)
    return (
        tuple(ordered),
        {key: tuple(value) for key, value in by_identifier.items()},
        duplicates,
    )


def _revision_discovery_refusal(
    source_sha256: str | None,
    code: str,
    *,
    part: str = _DOCUMENT,
    parts: Mapping[str, _ReviewPart] = {},
) -> dict[str, Any]:
    return {
        "passed": False,
        "code": code,
        "source_sha256": source_sha256,
        **_part_context(part, parts),
        "revisions": [],
        "unsupported": [],
    }


def _unsupported_part_records(
    parts: Mapping[str, bytes],
    selected_part: str,
    allowed: Mapping[str, _ReviewPart],
) -> list[dict[str, Any]]:
    """Report revisions outside the selected scope and unsupported revision markup."""

    unsupported: list[dict[str, Any]] = []
    for name in sorted(parts):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        region = allowed[name].region if name in allowed else None
        try:
            part_root = ET.fromstring(parts[name])
        except ET.ParseError:
            unsupported.append(
                {
                    "part": name,
                    "region": region,
                    "kind": None,
                    "revision_id": None,
                    "code": "review_revision_part_unparsable",
                }
            )
            continue
        parents: Mapping[int, ET.Element] | None = None
        word_namespace = _word_namespace(part_root) or _W
        moves: Mapping[int, _MovePair] | None = None
        for node in part_root.iter():
            if not isinstance(node.tag, str) or not node.tag.startswith("{"):
                continue
            node_namespace, local_name = node.tag[1:].split("}", 1)
            conflict = node.tag in _CONFLICT_REVISION_TAGS
            custom_xml_conflict = node.tag in _CUSTOM_XML_CONFLICT_TAGS
            attribute_namespace = (
                word_namespace if conflict or custom_xml_conflict else node_namespace
            )
            if (
                not (conflict or custom_xml_conflict)
                and node_namespace not in SUPPORTED_WORD_NAMESPACES
            ):
                continue
            if custom_xml_conflict:
                code = "review_revision_kind_unsupported"
            elif conflict or local_name in _SELECTABLE_REVISION_TAGS:
                if name != selected_part:
                    code = "review_revision_scope_unsupported"
                elif local_name in _PROPERTY_CHANGE_TAGS:
                    if parents is None:
                        parents = {
                            id(child): parent for parent in part_root.iter() for child in parent
                        }
                    if (
                        _property_change_structural_code(node, local_name, node_namespace, parents)
                        is not None
                    ):
                        code = "review_revision_kind_unsupported"
                    else:
                        continue
                else:
                    continue
            elif local_name in {kind.removesuffix("Start") + "End" for kind in _CUSTOM_XML_STARTS}:
                start_tag = Q(node_namespace, local_name.removesuffix("End") + "Start")
                if any(
                    start.get(Q(node_namespace, "id")) == node.get(Q(node_namespace, "id"))
                    for start in part_root.iter(start_tag)
                ):
                    continue
                code = "review_revision_kind_unsupported"
            elif local_name in _MOVE_TAGS and id(node) in (
                moves if moves is not None else (moves := _move_pairs(part_root, word_namespace))
            ):
                if name == selected_part:
                    continue
                code = "review_revision_scope_unsupported"
            elif local_name in _UNSUPPORTED_REVISION_TAGS or local_name.endswith(
                _UNSUPPORTED_REVISION_CHANGE_SUFFIX
            ):
                code = "review_revision_kind_unsupported"
            else:
                continue
            unsupported.append(
                {
                    "part": name,
                    "region": region,
                    "kind": local_name,
                    "revision_id": _selectable_revision_id(node.get(Q(attribute_namespace, "id"))),
                    "code": code,
                    **({"namespace": node_namespace} if conflict or custom_xml_conflict else {}),
                }
            )
    return unsupported


def discover_review_revisions(source: str | Path, *, part: str = _DOCUMENT) -> dict[str, Any]:
    """List native revisions an exact part-qualified decision can name.

    Read-only: the source package is opened and hashed, never rewritten. ``part``
    selects the body, one referenced header or footer, or one relationship-reachable
    footnotes or endnotes part. Every in-scope insertion or deletion is reported in
    that part's document order with its raw ID, author, date, one-based paragraph
    position, and the text a decision applies. A coherent whole-paragraph insertion
    or deletion is reported once, at its paragraph mark, and carries
    ``paragraph_mark: True`` so callers can tell it moves paragraph structure.
    Revisions inside separator, continuation-separator, or continuation-notice notes
    are never decidable. ``decidable`` is true only when ``decide_review`` can act on
    exactly that one revision; otherwise ``decision_code`` names the shared refusal.
    Inline conflict insertions and deletions use the same record shape and tracked
    text decisions for direct paragraph runs in supported stories and tables.
    Parent-object conflicts, nested revisions, controls, and alternate-content
    contexts remain undecidable; this does not model coauthor synchronization.
    A named move is reported once, at its ``moveFrom`` wrapper, as kind ``move``
    with ``move_name``, ``move_to_revision_id``, and ``move_to_paragraph``; accepting
    it keeps the destination text and rejecting it keeps the source text. Unpaired
    move halves stay unsupported. Each record is itself valid decision input, via
    ``ReviewRevisionRef.from_discovery`` or verbatim in a JSON decision request.
    Revisions in other parts, unsupported revision kinds, and unusable revision
    markup are reported under ``unsupported`` and are never advertised as selectable.
    Custom XML conflict range starts and ends remain unsupported and are reported
    individually with their Word-namespace IDs, including unpaired markers.
    """

    path = Path(source)
    source_sha256: str | None = None
    try:
        source_sha256 = _sha256(path)
        graph = PackageGraph.from_file(path)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return _revision_discovery_refusal(source_sha256, "review_source_package_invalid")
    try:
        document_root = ET.fromstring(parts[_DOCUMENT])
    except (KeyError, ET.ParseError):
        return _revision_discovery_refusal(source_sha256, "review_source_package_invalid")
    namespace = _word_namespace(document_root)
    if namespace is None:
        return _revision_discovery_refusal(source_sha256, "review_namespace_unsupported")
    allowed = _review_parts(graph, document_root, namespace)
    if part not in allowed:
        return _revision_discovery_refusal(
            source_sha256, "review_part_unsupported", part=part, parts=allowed
        )
    try:
        root = ET.fromstring(parts[part])
    except (KeyError, ET.ParseError):
        return _revision_discovery_refusal(source_sha256, "review_source_package_invalid")
    part_namespace = _word_namespace(root)
    if part_namespace is None:
        return _revision_discovery_refusal(source_sha256, "review_namespace_unsupported")
    selected = allowed[part]
    parents = {id(child): parent for parent in root.iter() for child in parent}
    ordered, _, duplicates = _part_revision_index(root, part_namespace)
    positions = _paragraph_positions(root, part_namespace)
    revisions = [
        _revision_record(
            entry, part, selected.region, part_namespace, parents, positions, duplicates
        )
        for entry in ordered
    ]
    return {
        "passed": True,
        "code": None,
        "source_sha256": source_sha256,
        **_part_context(part, allowed),
        "revisions": revisions,
        "unsupported": _unsupported_part_records(parts, part, allowed),
        **_revision_settings(graph, parts),
    }


def _revision_settings(graph: PackageGraph, parts: Mapping[str, bytes]) -> dict[str, Any]:
    relationships = [
        rel
        for rel in graph.relationships_for(_DOCUMENT)
        if rel.type.rsplit("/", 1)[-1] == "settings" and not rel.is_external
    ]
    if len(relationships) != 1 or relationships[0].resolved_target() not in parts:
        return {"track_revisions": None, "revision_view": None}
    try:
        root = ET.fromstring(parts[relationships[0].resolved_target()])
    except ET.ParseError:
        return {"track_revisions": None, "revision_view": None}
    namespace = _word_namespace(root)
    if namespace is None:
        return {"track_revisions": None, "revision_view": None}
    track = root.find(Q(namespace, "trackRevisions"))
    view = root.find(Q(namespace, "revisionView"))

    def enabled(value: str | None) -> bool:
        return value not in {"0", "false", "off"}

    return {
        "track_revisions": track is not None and enabled(track.get(Q(namespace, "val"))),
        "revision_view": None
        if view is None
        else {
            name: enabled(view.get(Q(namespace, name)))
            for name in ("markup", "comments", "insDel", "formatting", "inkAnnotations")
        },
    }


def _paragraph_positions(root: ET.Element, namespace: str) -> dict[int, int]:
    """One-based position of the first paragraph enclosing each node of one part."""

    positions: dict[int, int] = {}
    for position, paragraph in enumerate(root.iter(Q(namespace, "p")), start=1):
        for node in paragraph.iter():
            positions.setdefault(id(node), position)
    return positions


def _revision_record(
    entry: _IndexedRevision,
    part: str,
    region: str,
    namespace: str,
    parents: Mapping[int, ET.Element],
    positions: Mapping[int, int],
    duplicates: frozenset[str],
) -> dict[str, Any]:
    """The one discovery record for a logical revision; also the stale-binding identity."""

    code = _revision_decision_code(entry, namespace, parents, duplicates)
    note = _note_ancestor(entry.node, parents) if region in _NOTE_KINDS else None
    move = entry.move
    return {
        "part": part,
        "region": region,
        "revision_id": _selectable_revision_id(entry.identifier),
        "kind": entry.kind,
        "author": entry.author,
        "date": entry.date,
        "paragraph": positions.get(id(entry.node)),
        "text": entry.text,
        **(
            {
                "move_name": move.name,
                "move_to_revision_id": _selectable_revision_id(
                    move.destination.get(Q(namespace, "id"))
                ),
                "move_to_paragraph": positions.get(id(move.destination)),
            }
            if move is not None
            else {}
        ),
        "decidable": code is None,
        "decision_code": code,
        **({"paragraph_mark": True} if entry.group is not None else {}),
        **(_note_identity(note, namespace) if note is not None else {}),
    }


def _refusal(source_sha256: str, code: str) -> ReviewReport:
    return ReviewReport("refused", code, source_sha256, None, [], [])


def _write_staged(
    source: Path,
    output: Path,
    replacements: Mapping[str, bytes],
    additions: Mapping[str, bytes] = {},
) -> None:
    """Rewrite the named parts of one package and append the parts it does not have yet.

    Every part that is neither replaced nor added is copied from the source archive
    entry for entry, so its bytes are the source bytes. An addition that would
    duplicate an existing entry is refused instead of being written twice; the
    caller validates additions against the source parts before staging.
    New entries use fixed ZIP metadata; source entries retain their original metadata.
    """

    with (
        zipfile.ZipFile(source) as archive,
        zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as rewritten,
    ):
        seen: set[str] = set()
        for info in archive.infolist():
            if info.filename in seen:
                continue
            seen.add(info.filename)
            rewritten.writestr(info, replacements.get(info.filename, archive.read(info.filename)))
        for name, data in additions.items():
            if name in seen:
                raise ValueError("review_addition_conflict")
            seen.add(name)
            rewritten.writestr(deterministic_zip_info(name), data)


def _text_node(namespace: str, text: str, *, deleted: bool = False) -> ET.Element:
    node = ET.Element(Q(namespace, "delText" if deleted else "t"))
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set(_XML_SPACE, "preserve")
    return node


@dataclass(frozen=True)
class _TextRun:
    element: ET.Element
    text: str
    start: int
    end: int


@dataclass(frozen=True)
class _ProtectedRange:
    start: int
    end: int
    kind: str
    revision_id: str | None = None


@dataclass(frozen=True)
class _ParagraphMap:
    text: str
    runs: tuple[_TextRun, ...]
    editable_ranges: tuple[tuple[int, int], ...]
    protected_ranges: tuple[_ProtectedRange, ...]
    endpoint_positions: frozenset[int]
    annotation_ranges: tuple[dict[str, Any], ...]


class _ReviewContextError(Exception):
    def __init__(self, code: str) -> None:
        self.code = code


def _direct_text_run(
    node: ET.Element, namespace: str, text_tag: str
) -> tuple[ET.Element, str] | None:
    if node.tag != Q(namespace, "r") or any(
        child.tag not in {Q(namespace, "rPr"), Q(namespace, text_tag)} for child in node
    ):
        return None
    texts = node.findall(Q(namespace, text_tag))
    if len(texts) != 1:
        return None
    return texts[0], texts[0].text or ""


def _revision_visible_text(node: ET.Element, namespace: str, kind: str) -> str | None:
    text_tag = "t" if kind in {"ins", "conflictIns"} else "delText"
    if any(
        descendant is not node
        and descendant.tag
        in (
            {
                Q(namespace, "ins"),
                Q(namespace, "del"),
                Q(namespace, "moveFrom"),
                Q(namespace, "moveTo"),
                Q(namespace, "bookmarkStart"),
                Q(namespace, "bookmarkEnd"),
                Q(namespace, "commentRangeStart"),
                Q(namespace, "commentRangeEnd"),
                Q(namespace, "commentReference"),
            }
            | {Q(namespace, tag) for tag in _PROPERTY_CHANGE_TAGS}
            | _CONFLICT_REVISION_TAGS
        )
        for descendant in node.iter()
    ):
        return None
    texts: list[str] = []
    for child in node:
        direct = _direct_text_run(child, namespace, text_tag)
        if direct is None:
            return None
        _, text = direct
        texts.append(text)
    return "".join(texts)


def _paragraph_map(paragraph: ET.Element, namespace: str) -> _ParagraphMap:
    if any(node.tag in _CONFLICT_REVISION_TAGS for node in paragraph.iter()):
        raise _ReviewContextError("review_revision_context_unsupported")
    text_parts: list[str] = []
    runs: list[_TextRun] = []
    protected: list[_ProtectedRange] = []
    endpoint_positions: set[int] = set()
    annotation_ranges: list[dict[str, Any]] = []
    annotation_stack: list[tuple[str, str, int]] = []
    offset = 0

    for child in paragraph:
        if child.tag == Q(namespace, "pPr"):
            continue
        normal = _direct_text_run(child, namespace, "t")
        if normal is not None:
            _, text = normal
            rPr = child.find(Q(namespace, "rPr"))
            rPrChange = rPr.find(Q(namespace, "rPrChange")) if rPr is not None else None
            if rPrChange is not None:
                rev_id = rPrChange.get(Q(namespace, "id"))
                protected.append(_ProtectedRange(offset, offset + len(text), "revision", rev_id))
            runs.append(_TextRun(child, text, offset, offset + len(text)))
            text_parts.append(text)
            offset += len(text)
            continue
        if child.tag in {Q(namespace, "ins"), Q(namespace, "del")}:
            kind = child.tag.rsplit("}", 1)[-1]
            value = _revision_visible_text(child, namespace, kind)
            revision_id = child.get(Q(namespace, "id"))
            if value is None or not revision_id:
                raise _ReviewContextError("review_revision_context_unsupported")
            start = offset
            if kind == "ins":
                text_parts.append(value)
                offset += len(value)
            protected.append(_ProtectedRange(start, offset, "revision", revision_id))
            continue
        local_name = child.tag.rsplit("}", 1)[-1]
        if local_name in {"bookmarkStart", "commentRangeStart"}:
            identifier = child.get(Q(namespace, "id"))
            if not identifier:
                raise _ReviewContextError("review_annotation_context_unsupported")
            kind = "bookmark" if local_name == "bookmarkStart" else "comment"
            annotation_stack.append((kind, identifier, offset))
            endpoint_positions.add(offset)
            continue
        if local_name in {"bookmarkEnd", "commentRangeEnd"}:
            identifier = child.get(Q(namespace, "id"))
            kind = "bookmark" if local_name == "bookmarkEnd" else "comment"
            if (
                not identifier
                or not annotation_stack
                or annotation_stack[-1][:2] != (kind, identifier)
            ):
                raise _ReviewContextError("review_annotation_context_unsupported")
            _, _, start = annotation_stack.pop()
            endpoint_positions.add(offset)
            annotation_ranges.append(
                {"kind": kind, "id": identifier, "start": start, "end": offset}
            )
            continue
        if (
            local_name == "r"
            and sum(node.tag == Q(namespace, "commentReference") for node in child) == 1
            and all(
                node.tag in {Q(namespace, "rPr"), Q(namespace, "commentReference")}
                for node in child
            )
        ):
            endpoint_positions.add(offset)
            continue
        raise _ReviewContextError("review_target_context_unsupported")

    if annotation_stack:
        raise _ReviewContextError("review_annotation_context_unsupported")

    editable_ranges: list[tuple[int, int]] = []
    for run in runs:
        if not run.text:
            continue
        if (
            editable_ranges
            and editable_ranges[-1][1] == run.start
            and run.start not in endpoint_positions
        ):
            editable_ranges[-1] = (editable_ranges[-1][0], run.end)
        else:
            editable_ranges.append((run.start, run.end))
    return _ParagraphMap(
        "".join(text_parts),
        tuple(runs),
        tuple(editable_ranges),
        tuple(protected),
        frozenset(endpoint_positions),
        tuple(annotation_ranges),
    )


def _visible_text(paragraph: ET.Element, namespace: str) -> str:
    def text(node: ET.Element) -> str:
        if node.tag == Q(namespace, "del"):
            return ""
        if node.tag == Q(namespace, "t"):
            return node.text or ""
        return "".join(text(child) for child in node)

    return text(paragraph)


def _target_ranges(mapping: _ParagraphMap) -> dict[str, list[dict[str, Any]]]:
    return {
        "editable_ranges": [{"start": start, "end": end} for start, end in mapping.editable_ranges],
        "protected_ranges": [
            {
                "start": item.start,
                "end": item.end,
                "kind": item.kind,
                **({"revision_id": item.revision_id} if item.revision_id is not None else {}),
            }
            for item in mapping.protected_ranges
        ],
        "annotation_ranges": list(mapping.annotation_ranges),
    }


def _change_context_code(mapping: _ParagraphMap, change: ReviewChange) -> str | None:
    if change.end > len(mapping.text):
        return "review_target_text_unsupported"
    if change.start == change.end:
        if change.start in mapping.endpoint_positions or any(
            change.start in {item.start, item.end} for item in mapping.protected_ranges
        ):
            return "review_insertion_affinity_ambiguous"
        if any(start <= change.start <= end for start, end in mapping.editable_ranges):
            return None
        return "review_target_text_unsupported"
    if any(
        item.start < change.end
        and change.start < item.end
        or item.start == item.end
        and change.start < item.start < change.end
        for item in mapping.protected_ranges
    ):
        return "review_revision_intersection_unsupported"
    if any(change.start < position < change.end for position in mapping.endpoint_positions):
        return "review_annotation_boundary_unsupported"
    if any(start <= change.start and change.end <= end for start, end in mapping.editable_ranges):
        return None
    return "review_target_text_unsupported"


def _clone_run(run: ET.Element, namespace: str, text: str, *, deleted: bool = False) -> ET.Element:
    clone = copy.deepcopy(run)
    for child in list(clone):
        if child.tag == Q(namespace, "t"):
            clone.remove(child)
    clone.append(_text_node(namespace, text, deleted=deleted))
    return clone


def _revision(
    namespace: str,
    kind: str,
    revision_id: int,
    author: str,
    date: str,
    children: list[ET.Element],
) -> ET.Element:
    wrapper = ET.Element(
        Q(namespace, kind),
        {
            Q(namespace, "id"): str(revision_id),
            Q(namespace, "author"): author,
            Q(namespace, "date"): date,
        },
    )
    wrapper.extend(children)
    return wrapper


@dataclass(frozen=True)
class _AuthorEdit:
    """One ordered, non-overlapping text edit with the revision IDs allocated for it."""

    start: int
    end: int
    text: str
    deletion_id: int | None
    insertion_id: int | None


def _run_edit_span(
    run: _TextRun, edit: _AuthorEdit, position: int, emitted: frozenset[int]
) -> str | None:
    """How one edit meets one run: a caret insertion, a range replacement, or nothing."""

    if edit.start == edit.end:
        if position in emitted or not run.start <= edit.start <= run.end:
            return None
        return "insert"
    if run.end <= edit.start or run.start >= edit.end:
        return None
    return "replace"


def _revised_children(
    paragraph: ET.Element,
    mapping: _ParagraphMap,
    edits: tuple[_AuthorEdit, ...],
    namespace: str,
    author: str,
    date: str,
) -> list[ET.Element]:
    """Rewrite one paragraph's runs as native insertions and deletions.

    ``edits`` are ordered and non-overlapping over the paragraph's editable text.
    Each edit keeps its own revision IDs, so independently authored edits stay
    independently decidable. An insertion is emitted once, inside the first run that
    can carry it in document order, which matches the single-edit placement. Runs
    that no edit touches are copied byte for byte. The return value is the
    paragraph's replacement children, including its copied property element, so a
    caller that publishes a paragraph must rewrap them in the paragraph element.
    """

    runs = {id(run.element): run for run in mapping.runs}
    result: list[ET.Element] = []
    emitted: set[int] = set()
    for child in paragraph:
        run = runs.get(id(child))
        if run is None:
            result.append(copy.deepcopy(child))
            continue
        spans = [
            (position, edit, kind)
            for position, edit in enumerate(edits)
            if (kind := _run_edit_span(run, edit, position, frozenset(emitted))) is not None
        ]
        if not spans:
            result.append(copy.deepcopy(child))
            continue
        cursor = run.start
        for position, edit, kind in spans:
            if kind == "insert":
                offset = edit.start - run.start
                if offset > cursor - run.start:
                    result.append(
                        _clone_run(
                            run.element,
                            namespace,
                            run.text[cursor - run.start : offset],
                        )
                    )
                if edit.insertion_id is not None:
                    result.append(
                        _revision(
                            namespace,
                            "ins",
                            edit.insertion_id,
                            author,
                            date,
                            [_clone_run(run.element, namespace, edit.text)],
                        )
                    )
                    emitted.add(position)
                cursor = edit.start
                continue
            overlap_start, overlap_end = (
                max(run.start, edit.start),
                min(run.end, edit.end),
            )
            if overlap_start > cursor:
                result.append(
                    _clone_run(
                        run.element,
                        namespace,
                        run.text[cursor - run.start : overlap_start - run.start],
                    )
                )
            if edit.deletion_id is not None:
                result.append(
                    _revision(
                        namespace,
                        "del",
                        edit.deletion_id,
                        author,
                        date,
                        [
                            _clone_run(
                                run.element,
                                namespace,
                                run.text[overlap_start - run.start : overlap_end - run.start],
                                deleted=True,
                            )
                        ],
                    )
                )
            if edit.insertion_id is not None and position not in emitted:
                result.append(
                    _revision(
                        namespace,
                        "ins",
                        edit.insertion_id,
                        author,
                        date,
                        [_clone_run(run.element, namespace, edit.text)],
                    )
                )
                emitted.add(position)
            cursor = overlap_end
        if cursor < run.end:
            result.append(_clone_run(run.element, namespace, run.text[cursor - run.start :]))
    for position, edit in enumerate(edits):
        if edit.insertion_id is None or position in emitted or not mapping.runs:
            continue
        result.append(
            _revision(
                namespace,
                "ins",
                edit.insertion_id,
                author,
                date,
                [_clone_run(mapping.runs[-1].element, namespace, edit.text)],
            )
        )
    return result


def _all_revision_ids(parts: Mapping[str, bytes]) -> set[int]:
    ids: set[int] = set()
    for name, data in parts.items():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        for node in root.iter():
            if not isinstance(node.tag, str) or not node.tag.startswith("{"):
                continue
            namespace, local_name = node.tag[1:].split("}", 1)
            conflict = node.tag in _CONFLICT_REVISION_TAGS or node.tag in _CUSTOM_XML_CONFLICT_TAGS
            if not conflict and (
                namespace not in SUPPORTED_WORD_NAMESPACES
                or (
                    local_name not in _SELECTABLE_REVISION_TAGS
                    and local_name not in _UNSUPPORTED_REVISION_TAGS
                    and not local_name.endswith("Change")
                    and local_name
                    not in {
                        "moveFromRangeStart",
                        "moveFromRangeEnd",
                        "moveToRangeStart",
                        "moveToRangeEnd",
                    }
                )
            ):
                continue
            attribute_namespace = (_word_namespace(root) or _W) if conflict else namespace
            value = node.get(Q(attribute_namespace, "id"))
            if value is not None and value.isascii() and value.strip().removeprefix("+").isdigit():
                ids.add(int(value))
    return ids


def _apply_property_decision(
    node: ET.Element,
    kind: str,
    action: str,
    namespace: str,
    parents: Mapping[int, ET.Element],
) -> None:
    """Apply an accept or reject decision to one property container revision."""

    container_tag = kind[:-6]
    parent = parents.get(id(node))
    if parent is None:
        return
    prior = node.find(Q(namespace, container_tag))
    if action == "accept":
        parent.remove(node)
        if (
            len(parent) == 0
            and not parent.attrib
            and container_tag not in {"tblPr", "tblGrid", "tcPr", "sectPr"}
        ):
            grandparent = parents.get(id(parent))
            if grandparent is not None:
                grandparent.remove(parent)
    elif action == "reject":
        parent.remove(node)
        parent.clear()
        if prior is not None:
            parent.attrib.update(prior.attrib)
            for child in prior:
                parent.append(copy.deepcopy(child))
        if (
            len(parent) == 0
            and not parent.attrib
            and container_tag not in {"tblPr", "tblGrid", "tcPr", "sectPr"}
        ):
            grandparent = parents.get(id(parent))
            if grandparent is not None:
                grandparent.remove(parent)


@dataclass(frozen=True)
class _StagedPackage:
    """The staged package a ``verify_staged`` hook re-checks after the shared gate.

    ``part_sha256`` names every staged part and its digest. ``parsed_xml`` maps XML
    part names to read-only trees parsed from the staged bytes under the package
    limits; in ``changed-parts`` mode, parts outside ``replacements`` are the baseline
    trees of byte-identical source parts. Copy a tree before mutating it.
    """

    path: Path
    file_sha256: str
    part_sha256: Mapping[str, str]
    parsed_xml: Mapping[str, ET.Element]


def _stage_verified(
    staged: Path,
    replacements: Mapping[str, bytes],
    baseline: _Inspection | None,
    verify_mode: str,
    preserved_missing_target: str | None = None,
) -> _StagedPackage | None:
    """Run the shared staged-package gate once and describe what it validated."""

    if verify_mode == "full":
        candidate = _inspect_docx(staged, reuse=baseline)
        if not candidate.valid:
            if baseline is None or preserved_missing_target is None:
                return None
            errors = [issue for issue in candidate.report.issues if issue.severity == "error"]
            original_errors = [
                issue for issue in baseline.report.issues if issue.severity == "error"
            ]
            if (
                not errors
                or errors != original_errors
                or any(
                    issue.code != "dangling_relationship"
                    or issue.message
                    != f"Internal relationship target is missing: {preserved_missing_target}"
                    for issue in errors
                )
                or candidate.part_sha256.keys() != baseline.part_sha256.keys()
                or any(
                    digest != baseline.part_sha256[name]
                    for name, digest in candidate.part_sha256.items()
                    if name not in replacements
                )
                or any(
                    candidate.part_sha256.get(name) != hashlib.sha256(data).hexdigest()
                    for name, data in replacements.items()
                )
            ):
                return None
        return _StagedPackage(
            staged, candidate.file_sha256, candidate.part_sha256, candidate.parsed_xml
        )
    assert baseline is not None
    verified = _verify_changed_parts(baseline, staged, replacements)
    if verified is None:
        return None
    file_sha256, roots = verified
    return _StagedPackage(
        staged,
        file_sha256,
        {
            **baseline.part_sha256,
            **{name: hashlib.sha256(data).hexdigest() for name, data in replacements.items()},
        },
        {**baseline.parsed_xml, **roots},
    )


def _link_no_clobber(
    staged: Path, output: Path, *, exists_code: str, unavailable_code: str
) -> str | None:
    """Publish ``staged`` at ``output`` only while ``output`` does not exist.

    ``os.link`` never replaces an existing file, so a destination created by any
    writer before the link is preserved and reported with ``exists_code``. A
    filesystem that cannot hard-link fails closed with ``unavailable_code`` instead of
    falling back to a clobbering rename. Returns ``None`` once published.
    """

    try:
        os.link(staged, output)
    except FileExistsError:
        return exists_code
    except (AttributeError, NotImplementedError, OSError):
        return unavailable_code
    return None


def _publish(
    source: Path,
    output: Path,
    source_sha256: str,
    replacements: Mapping[str, bytes],
    revisions: list[dict[str, Any]],
    targets: list[dict[str, Any]],
    *,
    bindings: Mapping[Path, str] = {},
    additions: Mapping[str, bytes] = {},
    verify: Callable[[Path], str | None] | None = None,
    verify_staged: Callable[[_StagedPackage], str | None] | None = None,
    baseline: _Inspection | None = None,
    verify_mode: str = "full",
    preserved_missing_target: str | None = None,
    failure_code: str = "review_staged_verification_failed",
    exists_code: str = "review_destination_invalid",
    unavailable_code: str = "review_publication_unavailable",
) -> ReviewReport:
    """Stage, verify, and publish one rewritten package without ever clobbering.

    Publication re-reads the source and every bound input after staging; a digest
    that has moved by then abandons the staged file, so no change visible before
    that read is ever published against a stale receipt. A write concurrent with
    the read is not covered. The staged file is then linked into place
    with ``os.link``, which only creates the destination while it does not exist, so
    a writer that creates the destination at any point is preserved and reported
    with ``exists_code``. A filesystem that cannot hard-link fails closed with
    ``unavailable_code`` instead of falling back to a clobbering rename.

    The staged package is validated once. With ``verify_mode="full"`` (the default)
    it is inspected whole; a ``baseline`` inspection of the source lends its parsed
    trees to byte-identical parts. ``verify_mode="changed-parts"`` requires a valid
    ``baseline`` and no ``additions``: every part outside ``replacements`` must hash
    identically to the baseline, each replaced part must hold exactly its intended
    bytes, and only the replaced XML parts are re-parsed under the package limits.

    A caller that authors a surface this module does not model passes its own
    hook and the parts it creates in ``additions``. ``verify`` re-reads the staged
    package from its path; ``verify_staged`` instead receives the already validated
    :class:`_StagedPackage`, so it re-checks its invariant without inspecting the
    package again. Either returns the refusal code for its own invariant, so a
    publication decision is still made once, on one staged file, after the shared
    source and package checks. The source and every bound input are read once more
    after the hook returns, so no callback can publish against a receipt that moved.
    A write concurrent with that last read is not covered.

    ``preserved_missing_target`` admits only an existing dangling relationship to
    that exact absent part. Full bounded inspection still runs; its errors must
    equal the source's errors, all must name that target, and every unedited part
    must hash identically. Relationships, content types and additions cannot change.
    """

    _require_verify_mode(verify_mode)
    if verify is not None and verify_staged is not None:
        raise ValueError("publish_verify_hooks_exclusive")
    if verify_mode == "changed-parts" and (baseline is None or additions):
        raise ValueError("publish_changed_parts_requires_baseline_without_additions")
    if preserved_missing_target is not None:
        if (
            verify_mode != "full"
            or additions
            or any(name.endswith(".rels") or name == "[Content_Types].xml" for name in replacements)
        ):
            raise ValueError("publish_missing_target_requires_data_only_full_verification")
        baseline = baseline or _inspect_docx(source)
        if baseline.file_sha256 != source_sha256:
            return _refusal(source_sha256, failure_code)
    staged: Path | None = None
    try:
        if output.exists() or output.is_symlink():
            return _refusal(source_sha256, exists_code)
        with tempfile.NamedTemporaryFile(
            dir=output.parent, prefix=f".{output.name}.", suffix=".tmp", delete=False
        ) as file:
            staged = Path(file.name)
        _write_staged(source, staged, replacements, additions)
        candidate = _stage_verified(
            staged, replacements, baseline, verify_mode, preserved_missing_target
        )
        if (
            candidate is None
            or _sha256(source) != source_sha256
            or any(_sha256(path) != digest for path, digest in bindings.items())
        ):
            return _refusal(source_sha256, failure_code)
        if verify is not None or verify_staged is not None:
            if verify is not None:
                verification = verify(staged)
            else:
                assert verify_staged is not None
                verification = verify_staged(candidate)
            if verification is not None:
                return _refusal(source_sha256, verification)
            # The hook is the last code that runs before publication, so the source
            # and every bound input are read once more after it returns: a hook that
            # moved a receipt the caller holds cannot reach the destination
            # undetected. A change concurrent with this last read is not covered.
            if _sha256(source) != source_sha256 or any(
                _sha256(path) != digest for path, digest in bindings.items()
            ):
                return _refusal(source_sha256, failure_code)
        linked = _link_no_clobber(
            staged, output, exists_code=exists_code, unavailable_code=unavailable_code
        )
        if linked is not None:
            return _refusal(source_sha256, linked)
        return ReviewReport(
            "published",
            None,
            source_sha256,
            candidate.file_sha256,
            revisions,
            targets,
        )
    finally:
        if staged is not None:
            staged.unlink(missing_ok=True)


def _load_part(
    parts: Mapping[str, bytes],
    roots: dict[str, ET.Element],
    namespaces: dict[str, str],
    part: str,
) -> tuple[ET.Element, str] | None:
    """Parse and namespace-check one part once, remembering the result per request."""

    if part in roots:
        return roots[part], namespaces[part]
    try:
        root = parse_xml_with_preserved_qnames(parts[part])
    except (KeyError, ET.ParseError):
        return None
    namespace = _word_namespace(root)
    if namespace is None:
        return None
    roots[part] = root
    namespaces[part] = namespace
    return root, namespace


def author_review(request: ReviewRequest) -> ReviewReport:
    source, output = request.source.resolve(), request.output.resolve()
    if source == output or not source.is_file() or not output.parent.is_dir():
        return _refusal(request.source_sha256, "review_destination_invalid")
    if _sha256(source) != request.source_sha256:
        return _refusal(request.source_sha256, "review_source_hash_mismatch")
    try:
        graph = PackageGraph.from_file(source)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
        document_root = ET.fromstring(parts[_DOCUMENT])
        if _word_namespace(document_root) is None:
            return _refusal(request.source_sha256, "review_namespace_unsupported")
        allowed = _review_parts(graph, document_root, _word_namespace(document_root) or "")
        roots: dict[str, ET.Element] = {}
        namespaces: dict[str, str] = {}
        paragraphs_by_part: dict[str, list[ET.Element]] = {}
        resolved: list[tuple[ET.Element, ReviewChange, _ParagraphMap]] = []
        for change in request.changes:
            if change.part not in allowed:
                return _refusal(request.source_sha256, "review_part_unsupported")
            loaded = _load_part(parts, roots, namespaces, change.part)
            if loaded is None:
                return _refusal(request.source_sha256, "review_package_invalid")
            root, namespace = loaded
            paragraphs = paragraphs_by_part.setdefault(
                change.part, list(root.iter(Q(namespace, "p")))
            )
            if change.paragraph > len(paragraphs):
                return _refusal(request.source_sha256, "review_target_missing")
            paragraph = paragraphs[change.paragraph - 1]
            if _mask_text(paragraph, namespace) != change.structure_fingerprint_sha256:
                return _refusal(request.source_sha256, "review_target_stale")
            try:
                mapping = _paragraph_map(paragraph, namespace)
            except _ReviewContextError as error:
                return _refusal(request.source_sha256, error.code)
            code = _change_context_code(mapping, change)
            if code is not None:
                return _refusal(request.source_sha256, code)
            resolved.append((paragraph, change, mapping))
        next_id = max(_all_revision_ids(parts), default=0) + 1
        revisions: list[dict[str, Any]] = []
        parents_by_part = {
            part: {id(child): parent for parent in root.iter() for child in parent}
            for part, root in roots.items()
        }
        for paragraph, change, mapping in resolved:
            namespace = namespaces[change.part]
            deletion_id = next_id if change.end > change.start else None
            next_id += 1 if deletion_id is not None else 0
            insertion_id = next_id if change.text else None
            next_id += 1 if insertion_id is not None else 0
            edits = (
                _AuthorEdit(
                    change.start,
                    change.end,
                    change.text,
                    deletion_id,
                    insertion_id,
                ),
            )
            children = _revised_children(
                paragraph,
                mapping,
                edits,
                namespace,
                request.author,
                request.date,
            )
            parent = parents_by_part[change.part][id(paragraph)]
            index = list(parent).index(paragraph)
            replacement = copy.copy(paragraph)
            del replacement[:]
            replacement.text = None
            replacement.tail = None
            replacement.extend(children)
            parent.remove(paragraph)
            parent.insert(index, replacement)
            for identifier in (deletion_id, insertion_id):
                if identifier is not None:
                    revisions.append({"part": change.part, "revision_id": identifier})
        replacements = {
            part: serialize_xml_with_preserved_qnames(root) for part, root in roots.items()
        }
        return _publish(
            source,
            output,
            request.source_sha256,
            replacements,
            revisions,
            [
                {
                    "part": change.part,
                    "paragraph": change.paragraph,
                    "start": change.start,
                    "end": change.end,
                }
                for _, change, _ in resolved
            ],
        )
    except (OSError, PackageError, ET.ParseError):
        return _refusal(request.source_sha256, "review_package_invalid")


def _author_markup(root: ET.Element, author: str) -> list[ET.Element]:
    """Every revision markup element one part attributes to ``author``.

    Any Word-namespace element carrying that ``w:author`` counts, so revision
    kinds this module cannot decide are found too; comments carry an author but
    are not revisions.
    """

    word_namespace = _word_namespace(root) or _W
    found: list[ET.Element] = []
    for node in root.iter():
        if node.tag in _CONFLICT_REVISION_TAGS:
            attribute_namespace = word_namespace
        elif isinstance(node.tag, str) and node.tag.startswith("{"):
            attribute_namespace, local_name = node.tag[1:].split("}", 1)
            if attribute_namespace not in SUPPORTED_WORD_NAMESPACES or local_name == "comment":
                continue
        else:
            continue
        if node.get(Q(attribute_namespace, "author")) == author:
            found.append(node)
    return found


def _author_markup_uncovered(
    parts: Mapping[str, bytes],
    author: str,
    roots: Mapping[str, ET.Element],
    covered: frozenset[int],
) -> bool:
    """True when any Word part holds author markup outside ``covered``.

    ``roots`` supplies parts already parsed for selection so node identity
    matches ``covered``; every other part is parsed here. An unparsable part
    raises ``ET.ParseError`` because it could hide that author's revisions.
    """

    for name in sorted(parts):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        root = roots[name] if name in roots else ET.fromstring(parts[name])
        if any(id(node) not in covered for node in _author_markup(root, author)):
            return True
    return False


def _indexed_revision_nodes(entry: _IndexedRevision) -> tuple[ET.Element, ...]:
    """Every markup element one logical revision's decision removes or unwraps."""

    if entry.group is not None:
        return (entry.group.mark, *entry.group.members)
    if entry.move is not None:
        move = entry.move
        return (
            move.source,
            move.destination,
            *move.source_range,
            *move.destination_range,
        )
    return (entry.node,)


def _author_selection(
    author: str,
    parts: Mapping[str, bytes],
    allowed: Mapping[str, _ReviewPart],
    roots: dict[str, ET.Element],
    namespaces: dict[str, str],
    parents_by_part: dict[str, dict[int, ET.Element]],
) -> tuple[list[tuple[str, _IndexedRevision]], tuple[ReviewRevisionRef, ...]] | str:
    """Select every discoverable revision by one author, or the refusal blocking all of them.

    Parts are visited in name order and revisions in document order, exactly as
    ``discover_review_revisions`` reports them; each selection is bound to its
    live discovery record. The first undecidable match refuses the whole request
    with its ``decision_code``, and author markup that no selected revision owns
    (unsupported kinds, unpaired moves, revisions outside reviewable parts)
    refuses with ``review_revision_author_scope_unsupported``. Only parts with a
    match are recorded in ``roots``, so untouched parts are never reserialized.
    """

    selected: list[tuple[str, _IndexedRevision]] = []
    references: list[ReviewRevisionRef] = []
    covered: set[int] = set()
    scanned: dict[str, ET.Element] = {}
    scanned_namespaces: dict[str, str] = {}
    for name in sorted(allowed):
        loaded = _load_part(parts, scanned, scanned_namespaces, name)
        if loaded is None:
            return "review_package_invalid"
        root, namespace = loaded
        ordered, _, duplicates = _part_revision_index(root, namespace)
        matched = [entry for entry in ordered if entry.author == author]
        if not matched:
            continue
        parents = {id(child): parent for parent in root.iter() for child in parent}
        positions = _paragraph_positions(root, namespace)
        for entry in matched:
            record = _revision_record(
                entry,
                name,
                allowed[name].region,
                namespace,
                parents,
                positions,
                duplicates,
            )
            if record["decision_code"] is not None:
                return record["decision_code"]
            selected.append((name, entry))
            references.append(
                ReviewRevisionRef(name, record["revision_id"], _revision_binding_sha256(record))
            )
            covered.update(id(node) for node in _indexed_revision_nodes(entry))
        roots[name], namespaces[name], parents_by_part[name] = root, namespace, parents
    if not selected:
        return "review_revision_author_unmatched"
    if _author_markup_uncovered(parts, author, scanned, frozenset(covered)):
        return "review_revision_author_scope_unsupported"
    return selected, tuple(references)


def _author_residual_verifier(author: str) -> Callable[[Path], str | None]:
    """Staged-package check: no markup attributed to ``author`` survives the decision."""

    def verify(staged: Path) -> str | None:
        try:
            graph = PackageGraph.from_file(staged)
            parts = {name: graph.read_part_bytes(name) for name in graph.parts}
            if _author_markup_uncovered(parts, author, {}, frozenset()):
                return "review_revision_author_residual"
        except (OSError, PackageError, ET.ParseError):
            return "review_staged_verification_failed"
        return None

    return verify


def decide_review(request: ReviewDecisionRequest) -> ReviewReport:
    source, output = request.source.resolve(), request.output.resolve()
    if source == output or not source.is_file() or not output.parent.is_dir():
        return _refusal(request.source_sha256, "review_destination_invalid")
    if _sha256(source) != request.source_sha256:
        return _refusal(request.source_sha256, "review_source_hash_mismatch")
    if request.author_filter is None:
        if not request.revisions:
            return _refusal(request.source_sha256, "review_decision_selection_empty")
    elif request.revisions:
        return _refusal(request.source_sha256, "review_decision_selection_ambiguous")
    elif (
        not isinstance(request.author_filter, str)
        or not request.author_filter
        or len(request.author_filter) > 256
    ):
        return _refusal(request.source_sha256, "review_author_filter_invalid")
    try:
        graph = PackageGraph.from_file(source)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
        document_root = ET.fromstring(parts[_DOCUMENT])
        document_namespace = _word_namespace(document_root)
        if document_namespace is None:
            return _refusal(request.source_sha256, "review_namespace_unsupported")
        allowed = _review_parts(graph, document_root, document_namespace)
        roots: dict[str, ET.Element] = {}
        namespaces: dict[str, str] = {}
        indexes: dict[str, tuple[Mapping[str, tuple[_IndexedRevision, ...]], frozenset[str]]] = {}
        parents_by_part: dict[str, dict[int, ET.Element]] = {}
        selected: list[tuple[str, _IndexedRevision]] = []
        positions_by_part: dict[str, dict[int, int]] = {}
        references = request.revisions
        verify: Callable[[Path], str | None] | None = None
        if request.author_filter is not None:
            chosen = _author_selection(
                request.author_filter,
                parts,
                allowed,
                roots,
                namespaces,
                parents_by_part,
            )
            if isinstance(chosen, str):
                return _refusal(request.source_sha256, chosen)
            selected, references = chosen
            verify = _author_residual_verifier(request.author_filter)
        if len({(item.part, item.revision_id) for item in request.revisions}) != len(
            request.revisions
        ):
            return _refusal(request.source_sha256, "review_revision_ids_duplicate")
        for reference in request.revisions:
            if reference.part not in allowed:
                return _refusal(request.source_sha256, "review_part_unsupported")
            loaded = _load_part(parts, roots, namespaces, reference.part)
            if loaded is None:
                return _refusal(request.source_sha256, "review_package_invalid")
            root, namespace = loaded
            if reference.part not in indexes:
                _, matches, duplicates = _part_revision_index(root, namespace)
                indexes[reference.part] = (matches, duplicates)
                parents_by_part[reference.part] = {
                    id(child): parent for parent in root.iter() for child in parent
                }
            matches, duplicates = indexes[reference.part]
            entries = matches.get(str(reference.revision_id), ())
            if not entries:
                return _refusal(request.source_sha256, "review_revision_target_missing")
            code = _revision_decision_code(
                entries[0],
                namespaces[reference.part],
                parents_by_part[reference.part],
                duplicates,
            )
            if code is not None:
                return _refusal(request.source_sha256, code)
            entry = entries[0]
            if entry.kind == "numberingChange" and request.action == "reject":
                return _refusal(
                    request.source_sha256, "review_revision_numbering_original_unavailable"
                )
            if entry.kind in {"cellIns", "cellDel"} and (
                (entry.kind == "cellIns") == (request.action == "reject")
            ):
                parent = parents_by_part[reference.part][id(entry.node)]
                cell = parents_by_part[reference.part][id(parent)]
                row = parents_by_part[reference.part].get(id(cell))
                if row is None or len(row.findall(Q(namespace, "tc"))) < 2:
                    return _refusal(
                        request.source_sha256, "review_revision_cell_structure_unsupported"
                    )
                if any(
                    n is not entry.node and _local_name(n.tag) in _SELECTABLE_REVISION_TAGS
                    for n in cell.iter()
                ):
                    return _refusal(
                        request.source_sha256, "review_revision_intersection_unsupported"
                    )
            if reference.revision_sha256 is not None:
                if reference.part not in positions_by_part:
                    positions_by_part[reference.part] = _paragraph_positions(root, namespace)
                live = _revision_record(
                    entries[0],
                    reference.part,
                    allowed[reference.part].region,
                    namespace,
                    parents_by_part[reference.part],
                    positions_by_part[reference.part],
                    duplicates,
                )
                if _revision_binding_sha256(live) != reference.revision_sha256:
                    return _refusal(request.source_sha256, "review_revision_target_stale")
            selected.append((reference.part, entries[0]))
        removed_by_row: dict[tuple[str, int], set[int]] = {}
        for part, entry in selected:
            if entry.kind not in {"cellIns", "cellDel"} or (
                (entry.kind == "cellIns") != (request.action == "reject")
            ):
                continue
            parents = parents_by_part[part]
            cell = parents[id(parents[id(entry.node)])]
            row = parents[id(cell)]
            removed = removed_by_row.setdefault((part, id(row)), set())
            removed.add(id(cell))
            if len(removed) >= len(row.findall(Q(namespaces[part], "tc"))):
                return _refusal(request.source_sha256, "review_revision_cell_structure_unsupported")
        for part, entry in selected:
            namespace = namespaces[part]
            parents = parents_by_part[part]
            node = entry.node
            if entry.group is not None:
                _apply_group_decision(entry.group, request.action, namespace, parents)
                continue
            if entry.move is not None:
                _apply_move_decision(entry.move, request.action, namespace, parents)
                continue
            if entry.kind in _PROPERTY_CHANGE_TAGS:
                _apply_property_decision(node, entry.kind, request.action, namespace, parents)
                continue
            if entry.kind in _CELL_REVISION_TAGS:
                properties = parents[id(node)]
                cell = parents[id(properties)]
                if entry.kind == "cellMerge":
                    for merge in list(properties.findall(Q(namespace, "vMerge"))):
                        properties.remove(merge)
                    value = node.get(
                        Q(namespace, "vMerge" if request.action == "accept" else "vMergeOrig")
                    )
                    if value is not None:
                        merge = ET.Element(Q(namespace, "vMerge"))
                        merge.set(Q(namespace, "val"), "restart" if value == "rest" else "continue")
                        properties.insert(list(properties).index(node), merge)
                    properties.remove(node)
                elif (entry.kind == "cellIns") == (request.action == "reject"):
                    parents[id(cell)].remove(cell)
                else:
                    properties.remove(node)
                continue
            if entry.kind in _CUSTOM_XML_STARTS:
                root = roots[part]
                end_tag = Q(namespace, entry.kind.removesuffix("Start") + "End")
                end = next(
                    n for n in root.iter(end_tag) if n.get(Q(namespace, "id")) == entry.identifier
                )
                parents[id(end)].remove(end)
                parents[id(node)].remove(node)
                continue
            if entry.kind == "numberingChange":
                parents[id(node)].remove(node)
                continue
            parent = parents.get(id(node))
            if parent is None:
                return _refusal(request.source_sha256, "review_revision_target_malformed")
            index = list(parent).index(node)
            if (request.action == "accept" and entry.kind in {"ins", "conflictIns"}) or (
                request.action == "reject" and entry.kind in {"del", "conflictDel"}
            ):
                for child in list(node):
                    if request.action == "reject" and entry.kind in {
                        "del",
                        "conflictDel",
                    }:
                        for text_node in child.iter(Q(namespace, "delText")):
                            text_node.tag = Q(namespace, "t")
                    parent.insert(index, child)
                    index += 1
            parent.remove(node)
        replacements = {
            part: serialize_xml_with_preserved_qnames(root) for part, root in roots.items()
        }
        return _publish(
            source,
            output,
            request.source_sha256,
            replacements,
            [reference.to_dict() for reference in references],
            [
                {
                    "part": reference.part,
                    "action": request.action,
                    "revision_id": reference.revision_id,
                }
                for reference in references
            ],
            verify=verify,
        )
    except (OSError, PackageError, ET.ParseError):
        return _refusal(request.source_sha256, "review_package_invalid")
