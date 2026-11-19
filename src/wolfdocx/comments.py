"""Source-bound native Word comment discovery and authoring.

The authoring surface is the relationship-reachable comments part together with the
comment markers the package parts actually hold. One comment is one ``w:comment``
element plus the three markers that point at it - ``w:commentRangeStart``,
``w:commentRangeEnd``, and the ``w:commentReference`` run - so a discovered comment
is named by its exact ID and by one fingerprint over its own content and its
resolved anchors. Insert writes both the comments part and the anchors of a range
the caller binds to the source text, update rewrites one comment in place, and
delete removes one comment together with every marker it owns.

Anything that cannot be resolved exactly is refused instead of guessed: an anchor
whose paragraph text this engine does not model, an anchor that is not exactly one
start, one end, and one reference inside a single paragraph, an anchor that lies
outside the parts this engine reaches, a comment body whose text is not the simple
shape this engine rewrites, and a comment ID that is not unique in the package.
Everything else is preserved: a part the request does not change is copied from the
source archive byte for byte, the body text of every other paragraph is untouched,
and existing durable IDs and paragraph IDs retain their lexical values. Reply and
resolve associate modern thread metadata with the last paragraph of each comment.
Duplicate or invalid identity metadata is diagnosed before publication; deletion
keeps orphan metadata rows and unrelated extension data intact. Every reply appends
a durable identity, a commentsExtensible row carrying that durable ID and the UTC
date, and a people entry for its author when none is registered, creating each
reachable part when absent; a malformed or mistyped people or commentsExtensible part
refuses the reply. Existing comments are not normalized or assigned durable
identities by this bootstrap. Discovery reports each comment's durable ID, UTC date,
placeholder flag, reactions, author presence and anchored task IDs.
"""

from __future__ import annotations

import copy
import itertools
import json
import zipfile
from collections import Counter
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from ._comment_extensions import (
    EXTENSIBLE,
    PEOPLE,
    TASKS,
    ExtensionIndex,
    ExtensionPart,
    append_extensible,
    new_root,
    register_person,
    resolve_extension_part,
)
from ._comment_threads import (
    _COMMENTS_EXTENDED_CONTENT_TYPE,
    _COMMENTS_EXTENDED_PART,
    _COMMENTS_EXTENDED_RELATIONSHIP,
    _COMMENTS_IDS_CONTENT_TYPE,
    _COMMENTS_IDS_PART,
    _COMMENTS_IDS_RELATIONSHIP,
    _W14,
    _W15,
    _W16CID,
    CommentReply,
    CommentResolve,
    collect_existing_durable_ids,
    collect_existing_para_ids,
    comment_identity_code,
    generate_collision_safe_durable_id,
    generate_collision_safe_para_id,
    last_comment_paragraph,
    parse_thread_state,
    resolve_metadata_part,
)
from .assembly import _add_document_relationship, _ensure_override, _rels_name
from .authoring import _MAX_TEXT_CHARACTERS
from .canonical import semantic_sha256
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import (
    _DOCUMENT,
    ReviewReport,
    _clone_run,
    _direct_text_run,
    _load_part,
    _mask_text,
    _metadata,
    _paragraph_target_records,
    _ParagraphMap,
    _part_context,
    _positive,
    _ProtectedRange,
    _publish,
    _review_parts,
    _ReviewContextError,
    _ReviewPart,
    _revision_visible_text,
    _sha,
    _sha256,
    _text_node,
    _TextRun,
    _visible_text,
    _word_namespace,
)

_COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
_COMMENTS_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
_STRICT_COMMENTS_RELATIONSHIP = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments"
_COMMENTS_RELATIONSHIPS = frozenset({_COMMENTS_RELATIONSHIP, _STRICT_COMMENTS_RELATIONSHIP})
_COMMENTS_PART = "word/comments.xml"
_COMMENT_REFERENCE_STYLE = "CommentReference"
_RANGE_TAGS = frozenset({"commentRangeStart", "commentRangeEnd"})
_MARKER_TAGS = _RANGE_TAGS | {"commentReference"}
_REFUSED = "refused"
_PUBLISHED = "published"


@dataclass(frozen=True)
class CommentInsert:
    """One new comment anchored to an existing, source-bound paragraph range."""

    part: str
    paragraph: int
    start: int
    end: int
    expected_old_text: str
    structure_fingerprint_sha256: str
    author: str
    initials: str | None
    date: str
    text: str
    end_paragraph: int | None = None
    end_structure_fingerprint_sha256: str | None = None

    def __init__(
        self,
        part: str,
        paragraph: int | None = None,
        start: int = 0,
        end: int = 0,
        expected_old_text: str = "",
        structure_fingerprint_sha256: str = "",
        author: str = "",
        initials: str | None = None,
        date: str = "",
        text: str = "",
        end_paragraph: int | None = None,
        end_structure_fingerprint_sha256: str | None = None,
        *,
        start_paragraph: int | None = None,
        start_offset: int | None = None,
        end_offset: int | None = None,
        start_anchor: Mapping[str, Any] | None = None,
        end_anchor: Mapping[str, Any] | None = None,
        start_structure_fingerprint_sha256: str | None = None,
    ) -> None:
        p = paragraph if paragraph is not None else start_paragraph
        if p is None and start_anchor is not None:
            p = start_anchor.get("paragraph")
        if p is None:
            raise ValueError("comment_paragraph_required")

        s = start if start_offset is None else start_offset
        if start_anchor is not None and "start" in start_anchor:
            s = start_anchor["start"]

        e_p = end_paragraph
        if e_p is None and end_anchor is not None:
            e_p = end_anchor.get("paragraph")

        e = end if end_offset is None else end_offset
        if end_anchor is not None and "end" in end_anchor:
            e = end_anchor["end"]

        fp = structure_fingerprint_sha256
        if not fp and start_structure_fingerprint_sha256 is not None:
            fp = start_structure_fingerprint_sha256
        if not fp and start_anchor is not None:
            fp = start_anchor.get("structure_fingerprint_sha256", "")

        end_fp = end_structure_fingerprint_sha256
        if end_fp is None and end_anchor is not None:
            end_fp = end_anchor.get("structure_fingerprint_sha256")

        object.__setattr__(self, "part", part)
        object.__setattr__(self, "paragraph", p)
        object.__setattr__(self, "start", s)
        object.__setattr__(self, "end", e)
        object.__setattr__(self, "expected_old_text", expected_old_text)
        object.__setattr__(self, "structure_fingerprint_sha256", fp)
        object.__setattr__(self, "author", author)
        object.__setattr__(self, "initials", initials)
        object.__setattr__(self, "date", date)
        object.__setattr__(self, "text", text)
        object.__setattr__(self, "end_paragraph", e_p)
        object.__setattr__(self, "end_structure_fingerprint_sha256", end_fp)

    @property
    def start_paragraph(self) -> int:
        return self.paragraph

    @property
    def effective_end_paragraph(self) -> int:
        return self.end_paragraph if self.end_paragraph is not None else self.paragraph

    def target_dict(self) -> dict[str, Any]:
        d = {
            "action": "insert",
            "part": self.part,
            "paragraph": self.paragraph,
            "start": self.start,
            "end": self.end,
        }
        if self.end_paragraph is not None and self.end_paragraph != self.paragraph:
            d["end_paragraph"] = self.end_paragraph
        return d

    @classmethod
    def from_discovery(
        cls,
        record: Mapping[str, Any],
        *,
        author: str,
        date: str,
        text: str,
        initials: str | None = None,
        range_index: int = 0,
    ) -> CommentInsert:
        """Accept one discovered paragraph target plus comment metadata.

        The discovered target carries editable_ranges; range_index selects which one to use.
        To use a specific range, extract it from editable_ranges and pass it merged with the target,
        or pass the target and specify range_index. The caller supplies the comment
        body text, author, date, and optional initials.
        """
        if "start" in record and "end" in record:
            # Explicit range already provided
            start, end = record["start"], record["end"]
        else:
            # Extract from editable_ranges
            editable = record.get("editable_ranges", [])
            if not editable or range_index >= len(editable):
                raise ValueError("comment_no_editable_range")
            start, end = editable[range_index]["start"], editable[range_index]["end"]

        return cls(
            part=record["part"],
            paragraph=record["paragraph"],
            start=start,
            end=end,
            expected_old_text=record["text"],
            structure_fingerprint_sha256=record["structure_fingerprint_sha256"],
            author=author,
            initials=initials,
            date=date,
            text=text,
        )


@dataclass(frozen=True)
class CommentUpdate:
    """One rewrite of an existing comment, named by its ID and its own fingerprint."""

    comment_id: int
    comment_sha256: str
    author: str
    initials: str | None
    date: str
    text: str

    def target_dict(self) -> dict[str, Any]:
        return {"action": "update", "comment_id": self.comment_id}


@dataclass(frozen=True)
class CommentDelete:
    """One removal of an existing comment, named by its ID and its own fingerprint."""

    comment_id: int
    comment_sha256: str

    def target_dict(self) -> dict[str, Any]:
        return {"action": "delete", "comment_id": self.comment_id}


CommentEdit = CommentInsert | CommentUpdate | CommentDelete | CommentReply | CommentResolve


@dataclass(frozen=True)
class CommentRequest:
    source: Path
    source_sha256: str
    destination: Path
    edits: tuple[CommentEdit, ...]


@dataclass(frozen=True)
class CommentReport:
    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    comments: list[dict[str, Any]]
    changed_targets: list[dict[str, Any]]

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "comments": self.comments,
            "changed_targets": self.changed_targets,
        }


def _comment_part_name(value: Any) -> str:
    """A package part path is validated as a name only; the reachable-part gate decides support."""

    name = _metadata(value, "comment_part")
    if name.startswith("/") or "\\" in name or ".." in name.split("/") or not name.endswith(".xml"):
        raise ValueError("comment_part_invalid")
    return name


def _comment_date(value: Any) -> str:
    text = _metadata(value, "comment_date")
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError as error:
        raise ValueError("comment_date_invalid") from error
    if parsed.tzinfo is None:
        raise ValueError("comment_date_timezone_required")
    return text


def _comment_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value or len(value) > _MAX_TEXT_CHARACTERS:
        raise ValueError(f"{label}_invalid")
    return value


def _initials(value: Mapping[str, Any]) -> str | None:
    """The optional initials attribute; absent means "keep whatever the comment declares"."""

    if "initials" not in value:
        return None
    raw = value["initials"]
    if not isinstance(raw, str) or len(raw) > 256:
        raise ValueError("comment_initials_invalid")
    return raw


def _anchor_object(value: Any, label: str, offset_name: str) -> Mapping[str, Any]:
    """Parse one explicit cross-paragraph endpoint without silently discarding fields."""

    expected = {"paragraph", offset_name, "structure_fingerprint_sha256"}
    if not isinstance(value, Mapping) or set(value) != expected:
        raise ValueError(f"{label}_invalid")
    return value


def _parse_edit(value: Any) -> CommentEdit:
    if not isinstance(value, Mapping) or not value:
        raise ValueError("comment_edit_invalid")
    action = value.get("action")
    if action == "insert":
        allowed_keys = {
            "action",
            "part",
            "paragraph",
            "start_paragraph",
            "end_paragraph",
            "start",
            "start_offset",
            "end",
            "end_offset",
            "expected_old_text",
            "structure_fingerprint_sha256",
            "start_structure_fingerprint_sha256",
            "end_structure_fingerprint_sha256",
            "start_anchor",
            "end_anchor",
            "author",
            "initials",
            "date",
            "text",
        }
        if not set(value) <= allowed_keys:
            raise ValueError("comment_insert_fields_invalid")
        start_anchor = (
            _anchor_object(value["start_anchor"], "comment_start_anchor", "start")
            if "start_anchor" in value
            else None
        )
        end_anchor = (
            _anchor_object(value["end_anchor"], "comment_end_anchor", "end")
            if "end_anchor" in value
            else None
        )

        start_p_raw = value.get("paragraph")
        if start_p_raw is None:
            start_p_raw = value.get("start_paragraph")
        if start_p_raw is None and start_anchor is not None:
            start_p_raw = start_anchor["paragraph"]
        if start_p_raw is None:
            raise ValueError("comment_insert_fields_invalid")
        start_paragraph = _positive(start_p_raw, "comment_paragraph")

        end_p_raw = value.get("end_paragraph")
        if end_p_raw is None and end_anchor is not None:
            end_p_raw = end_anchor["paragraph"]
        end_paragraph = (
            _positive(end_p_raw, "comment_end_paragraph")
            if end_p_raw is not None
            else start_paragraph
        )

        if end_paragraph < start_paragraph:
            raise ValueError("comment_anchor_range_invalid")

        start_raw = value.get("start")
        if start_raw is None:
            start_raw = value.get("start_offset")
        if start_raw is None and start_anchor is not None:
            start_raw = start_anchor["start"]
        if start_raw is None:
            raise ValueError("comment_insert_fields_invalid")
        start = _positive(start_raw, "comment_start", zero=True)

        end_raw = value.get("end")
        if end_raw is None:
            end_raw = value.get("end_offset")
        if end_raw is None and end_anchor is not None:
            end_raw = end_anchor["end"]
        if end_raw is None:
            raise ValueError("comment_insert_fields_invalid")
        end = _positive(end_raw, "comment_end", zero=True)

        # A same-paragraph anchor with ``start == end`` is a point comment (Word's
        # insertion-point comment): it binds no text, so its expected text must be empty.
        point = end_paragraph == start_paragraph and end == start
        if end_paragraph == start_paragraph and end < start:
            raise ValueError("comment_anchor_range_invalid")
        if point and value.get("expected_old_text", "") != "":
            raise ValueError("comment_anchor_range_invalid")

        fp_raw = value.get("structure_fingerprint_sha256")
        if not fp_raw:
            fp_raw = value.get("start_structure_fingerprint_sha256")
        if not fp_raw and start_anchor is not None:
            fp_raw = start_anchor["structure_fingerprint_sha256"]
        if not fp_raw:
            raise ValueError("comment_insert_fields_invalid")
        fp = _sha(fp_raw, "comment_fingerprint")

        end_fp_raw = value.get("end_structure_fingerprint_sha256")
        if end_fp_raw is None and end_anchor is not None:
            end_fp_raw = end_anchor["structure_fingerprint_sha256"]
        end_fp = _sha(end_fp_raw, "comment_fingerprint") if end_fp_raw is not None else None
        if start_anchor is not None and (
            _positive(start_anchor["paragraph"], "comment_paragraph") != start_paragraph
            or _positive(start_anchor["start"], "comment_start", zero=True) != start
            or _sha(start_anchor["structure_fingerprint_sha256"], "comment_fingerprint") != fp
        ):
            raise ValueError("comment_start_anchor_invalid")
        if end_anchor is not None and (
            _positive(end_anchor["paragraph"], "comment_end_paragraph") != end_paragraph
            or _positive(end_anchor["end"], "comment_end", zero=True) != end
            or _sha(end_anchor["structure_fingerprint_sha256"], "comment_fingerprint") != end_fp
        ):
            raise ValueError("comment_end_anchor_invalid")
        if end_paragraph != start_paragraph and end_fp is None:
            raise ValueError("comment_anchor_fingerprints_required")

        if (
            "part" not in value
            or "expected_old_text" not in value
            or "author" not in value
            or "date" not in value
            or "text" not in value
        ):
            raise ValueError("comment_insert_fields_invalid")

        return CommentInsert(
            part=_comment_part_name(value["part"]),
            paragraph=start_paragraph,
            start=start,
            end=end,
            expected_old_text=""
            if point
            else _comment_text(value["expected_old_text"], "comment_expected_old_text"),
            structure_fingerprint_sha256=fp,
            author=_metadata(value["author"], "comment_author"),
            initials=_initials(value),
            date=_comment_date(value["date"]),
            text=_comment_text(value["text"], "comment_text"),
            end_paragraph=end_paragraph if end_paragraph != start_paragraph else None,
            end_structure_fingerprint_sha256=end_fp,
        )
    if action == "update":
        fields = {"action", "comment_id", "comment_sha256", "author", "date", "text"}
        if not fields <= set(value) <= fields | {"initials"}:
            raise ValueError("comment_update_fields_invalid")
        return CommentUpdate(
            _positive(value["comment_id"], "comment_id", zero=True),
            _sha(value["comment_sha256"], "comment_sha256"),
            _metadata(value["author"], "comment_author"),
            _initials(value),
            _comment_date(value["date"]),
            _comment_text(value["text"], "comment_text"),
        )
    if action == "delete":
        if set(value) != {"action", "comment_id", "comment_sha256"}:
            raise ValueError("comment_delete_fields_invalid")
        return CommentDelete(
            _positive(value["comment_id"], "comment_id", zero=True),
            _sha(value["comment_sha256"], "comment_sha256"),
        )
    if action == "reply":
        fields = {"action", "comment_id", "comment_sha256", "author", "date", "text"}
        if not fields <= set(value) <= fields | {"initials"}:
            raise ValueError("comment_reply_fields_invalid")
        return CommentReply(
            _positive(value["comment_id"], "comment_id", zero=True),
            _sha(value["comment_sha256"], "comment_sha256"),
            _metadata(value["author"], "comment_author"),
            _initials(value),
            _comment_date(value["date"]),
            _comment_text(value["text"], "comment_text"),
        )
    if action == "resolve":
        if set(value) != {"action", "comment_id", "comment_sha256", "resolved"}:
            raise ValueError("comment_resolve_fields_invalid")
        if not isinstance(value["resolved"], bool):
            raise ValueError("comment_resolve_status_invalid")
        return CommentResolve(
            _positive(value["comment_id"], "comment_id", zero=True),
            _sha(value["comment_sha256"], "comment_sha256"),
            value["resolved"],
        )
    raise ValueError("comment_edit_action_invalid")


def parse_comment_request(value: Any) -> CommentRequest:
    """Parse the strict JSON request for one atomic comment authoring operation."""

    if not isinstance(value, Mapping) or set(value) != {
        "source",
        "source_sha256",
        "destination",
        "edits",
    }:
        raise ValueError("comment_request_fields_invalid")
    if (
        not isinstance(value["source"], str)
        or not value["source"]
        or not isinstance(value["destination"], str)
        or not value["destination"]
    ):
        raise ValueError("comment_path_invalid")
    edits = value["edits"]
    if not isinstance(edits, list) or not edits:
        raise ValueError("comment_edits_invalid")
    parsed = tuple(_parse_edit(item) for item in edits)
    identifiers = [edit.comment_id for edit in parsed if not isinstance(edit, CommentInsert)]
    if len(set(identifiers)) != len(identifiers):
        raise ValueError("comment_ids_duplicate")
    return CommentRequest(
        Path(value["source"]),
        _sha(value["source_sha256"], "comment_source_sha256"),
        Path(value["destination"]),
        parsed,
    )


def load_comment_request(path: str | Path) -> CommentRequest:
    with Path(path).open(encoding="utf-8") as stream:
        return parse_comment_request(json.load(stream))


def _raw_id(value: Any) -> int | None:
    """The numeric identity of a raw ``w:id``, or None when a request cannot name it exactly."""

    if not isinstance(value, str) or not value.isascii() or not value.isdigit():
        return None
    parsed = int(value)
    return parsed if str(parsed) == value else None


@dataclass(frozen=True)
class _CommentsPart:
    """The one relationship-reachable comments part, with its parsed root."""

    part: str
    data: bytes
    root: ET.Element
    namespace: str


def _resolve_comments_part(
    graph: PackageGraph,
) -> tuple[_CommentsPart | None, str | None]:
    """Resolve the comments part, or the reason this package has none to edit.

    A comments part is eligible only when exactly one non-external ``comments``
    relationship leaves ``word/document.xml``, its content type is the comments
    content type, and its root element is ``w:comments``. An arbitrary user path can
    never be selected, and a package whose comment metadata is ambiguous is refused
    rather than guessed.
    """

    matches = [
        relationship
        for relationship in graph.relationships_for(_DOCUMENT)
        if not relationship.is_external and relationship.type in _COMMENTS_RELATIONSHIPS
    ]
    if not matches:
        return None, None
    if len(matches) > 1:
        return None, "comment_part_ambiguous"
    target = matches[0].resolved_target()
    package_part = graph.part(target)
    if package_part is None or package_part.content_type != _COMMENTS_CONTENT_TYPE:
        return None, "comment_part_unsupported"
    try:
        data = graph.read_part_bytes(target)
        root = parse_xml_with_preserved_qnames(data)
    except (PackageError, ET.ParseError):
        return None, "comment_part_invalid"
    namespace = _word_namespace(root)
    if namespace is None or root.tag != Q(namespace, "comments"):
        return None, "comment_part_unsupported"
    return _CommentsPart(target, data, root, namespace), None


@dataclass(frozen=True)
class _Marker:
    """One comment marker a package part holds, with its schema-relevant location."""

    part: str
    region: str | None
    paragraph: int | None
    kind: str
    comment_id: int | None
    direct_child: bool
    parent_direct_child: bool
    paragraph_child_index: int | None
    parent: ET.Element | None


@dataclass(frozen=True)
class _CommentRecord:
    """One discovered comment: its identity, its anchors, and its bound fingerprint."""

    comment_id: int
    element: ET.Element
    author: str | None
    initials: str | None
    date: str | None
    text: str
    body_code: str | None
    anchors: tuple[dict[str, Any], ...]
    anchor_code: str | None
    fingerprint: str
    parent_comment_id: int | None = None
    resolved: bool = False
    para_id: str | None = None
    extensions: Mapping[str, Any] = field(default_factory=dict)

    def identity(self) -> dict[str, Any]:
        return {
            "author": self.author,
            "initials": self.initials,
            "date": self.date,
            "text": self.text,
            "parent_comment_id": self.parent_comment_id,
            "resolved": self.resolved,
            "anchors": [dict(anchor) for anchor in self.anchors],
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "comment_id": self.comment_id,
            **self.identity(),
            **copy.deepcopy(dict(self.extensions)),
            "comment_sha256": self.fingerprint,
            "updatable": self.body_code is None,
            "update_code": self.body_code,
            "deletable": self.anchor_code is None,
            "delete_code": self.anchor_code,
        }


_ABSENT_PART = ExtensionPart()


@dataclass(frozen=True)
class _CommentSurface:
    """One package's comment index: the comments part, its comments, and the marker scope."""

    part: _CommentsPart | None
    records: tuple[_CommentRecord, ...]
    paragraphs: Mapping[str, tuple[ET.Element, ...]]
    ids: frozenset[int]
    unsupported: tuple[dict[str, Any], ...]
    extended_target: str | None = None
    ids_target: str | None = None
    people: ExtensionPart = _ABSENT_PART
    extensible: ExtensionPart = _ABSENT_PART


def _comment_fingerprint(comment_id: int, identity: Mapping[str, Any]) -> str:
    """The content-and-anchor identity an update or a delete must present back."""

    return semantic_sha256(
        {
            "comment_id": comment_id,
            "author": identity["author"],
            "initials": identity["initials"],
            "date": identity["date"],
            "text": identity["text"],
            "parent_comment_id": identity.get("parent_comment_id"),
            "resolved": bool(identity.get("resolved", False)),
            "anchors": [
                {
                    "part": anchor["part"],
                    "paragraph": anchor["paragraph"],
                    "start": anchor["start"],
                    "end": anchor["end"],
                    "text_sha256": semantic_sha256(anchor["text"]),
                }
                for anchor in identity["anchors"]
            ],
        }
    )


def _annotation_reference_run(node: ET.Element, namespace: str) -> bool:
    """The comment body's own marker run: one ``w:annotationRef`` beside run properties."""

    return (
        node.tag == Q(namespace, "r")
        and any(child.tag == Q(namespace, "annotationRef") for child in node)
        and all(child.tag in {Q(namespace, "rPr"), Q(namespace, "annotationRef")} for child in node)
    )


def _comment_reference_run(node: ET.Element, namespace: str) -> bool:
    """Whether ``node`` is the schema-defined run for one legacy comment reference."""

    return (
        node.tag == Q(namespace, "r")
        and sum(child.tag == Q(namespace, "commentReference") for child in node) == 1
        and all(
            child.tag in {Q(namespace, "rPr"), Q(namespace, "commentReference")} for child in node
        )
    )


def _comment_body(element: ET.Element, namespace: str) -> tuple[str, str | None]:
    """One comment's visible body text and, when it cannot be rewritten, the refusal code.

    The supported body is the shape this engine writes: one paragraph holding the
    optional paragraph properties, the optional ``w:annotationRef`` marker run, and
    simple text runs. A hyperlink, a drawing, a nested field, a second paragraph, or
    markup this engine does not author keeps its text readable but refuses a rewrite,
    so an update never silently discards comment content.
    """

    children = list(element)
    if len(children) != 1 or children[0].tag != Q(namespace, "p"):
        return _visible_text(element, namespace), "comment_body_unsupported"
    paragraph = children[0]
    content: list[ET.Element] = []
    for index, child in enumerate(paragraph):
        if child.tag == Q(namespace, "pPr"):
            if index != 0:
                return _visible_text(element, namespace), "comment_body_unsupported"
            continue
        content.append(child)
    head = 1 if content and _annotation_reference_run(content[0], namespace) else 0
    parts: list[str] = []
    for child in content[head:]:
        if child.tag != Q(namespace, "r"):
            return _visible_text(element, namespace), "comment_body_unsupported"
        texts = [node for node in child if node.tag not in {Q(namespace, "rPr"), Q(namespace, "t")}]
        if texts or len(child.findall(Q(namespace, "t"))) != 1:
            return _visible_text(element, namespace), "comment_body_unsupported"
        parts.append(child.find(Q(namespace, "t")).text or "")
    return "".join(parts), None


def _part_markers(
    part: str, region: str | None, root: ET.Element, namespace: str
) -> tuple[list[_Marker], list[ET.Element]]:
    """Every legacy comment marker one part holds, and that part's paragraphs in order."""

    paragraphs = list(root.iter(Q(namespace, "p")))
    owner: dict[int, int] = {}
    direct: set[int] = set()
    child_indices: dict[int, int] = {}
    parents = {id(child): parent for parent in root.iter() for child in parent}
    for position, paragraph in enumerate(paragraphs, start=1):
        for child_index, child in enumerate(paragraph):
            direct.add(id(child))
            for node in child.iter():
                owner.setdefault(id(node), position)
                child_indices.setdefault(id(node), child_index)
    markers: list[_Marker] = []
    for node in root.iter():
        if not isinstance(node.tag, str) or not node.tag.startswith("{"):
            continue
        node_namespace, local_name = node.tag[1:].split("}", 1)
        if node_namespace not in SUPPORTED_WORD_NAMESPACES or local_name not in _MARKER_TAGS:
            continue
        markers.append(
            _Marker(
                part,
                region,
                owner.get(id(node)),
                local_name,
                _raw_id(node.get(Q(node_namespace, "id"))),
                id(node) in direct,
                id(parents.get(id(node))) in direct,
                child_indices.get(id(node)),
                parents.get(id(node)),
            )
        )
    return markers, paragraphs


def _comment_paragraph_map(paragraph: ET.Element, namespace: str) -> _ParagraphMap:
    text_parts: list[str] = []
    runs: list[_TextRun] = []
    protected: list[_ProtectedRange] = []
    endpoint_positions: set[int] = set()
    annotation_ranges: list[dict[str, Any]] = []
    open_comments: dict[str, int] = {}
    bookmark_stack: list[tuple[str, int]] = []
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
        if local_name == "commentRangeStart":
            identifier = child.get(Q(namespace, "id"))
            if not identifier:
                raise _ReviewContextError("review_annotation_context_unsupported")
            open_comments[identifier] = offset
            endpoint_positions.add(offset)
            continue
        if local_name == "commentRangeEnd":
            identifier = child.get(Q(namespace, "id"))
            if not identifier:
                raise _ReviewContextError("review_annotation_context_unsupported")
            endpoint_positions.add(offset)
            if identifier in open_comments:
                start = open_comments.pop(identifier)
                annotation_ranges.append(
                    {"kind": "comment", "id": identifier, "start": start, "end": offset}
                )
            else:
                annotation_ranges.append(
                    {
                        "kind": "comment",
                        "id": identifier,
                        "start": 0,
                        "end": offset,
                        "open_start": True,
                    }
                )
            continue
        if local_name == "bookmarkStart":
            identifier = child.get(Q(namespace, "id"))
            if not identifier:
                raise _ReviewContextError("review_annotation_context_unsupported")
            bookmark_stack.append((identifier, offset))
            endpoint_positions.add(offset)
            continue
        if local_name == "bookmarkEnd":
            identifier = child.get(Q(namespace, "id"))
            if not identifier or not bookmark_stack or bookmark_stack[-1][0] != identifier:
                raise _ReviewContextError("review_annotation_context_unsupported")
            _, start = bookmark_stack.pop()
            endpoint_positions.add(offset)
            annotation_ranges.append(
                {"kind": "bookmark", "id": identifier, "start": start, "end": offset}
            )
            continue
        if _comment_reference_run(child, namespace):
            endpoint_positions.add(offset)
            continue
        raise _ReviewContextError("review_target_context_unsupported")

    if bookmark_stack:
        raise _ReviewContextError("review_annotation_context_unsupported")

    for identifier, start in open_comments.items():
        annotation_ranges.append(
            {
                "kind": "comment",
                "id": identifier,
                "start": start,
                "end": offset,
                "open_end": True,
            }
        )

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


def _anchor_code(found: tuple[_Marker, ...], allowed: Mapping[str, _ReviewPart]) -> str | None:
    """Why one comment's markers cannot be resolved to a valid comment anchor."""

    if sorted(marker.kind for marker in found) != sorted(_MARKER_TAGS):
        return "comment_anchor_malformed"
    parts = {marker.part for marker in found}
    if len(parts) != 1:
        return "comment_anchor_span_unsupported"
    part = next(iter(parts))
    if part not in allowed:
        return "comment_anchor_scope_unsupported"
    if any(marker.paragraph is None for marker in found):
        return "comment_anchor_context_unsupported"
    start_m = next(m for m in found if m.kind == "commentRangeStart")
    end_m = next(m for m in found if m.kind == "commentRangeEnd")
    reference_m = next(m for m in found if m.kind == "commentReference")
    if (
        not start_m.direct_child
        or not end_m.direct_child
        or not reference_m.parent_direct_child
        or reference_m.parent is None
        or not _comment_reference_run(reference_m.parent, _word_namespace(reference_m.parent) or "")
        or reference_m.parent.tag != Q(_word_namespace(reference_m.parent) or "", "r")
        or reference_m.paragraph_child_index is None
        or end_m.paragraph_child_index is None
        or reference_m.paragraph_child_index != end_m.paragraph_child_index + 1
    ):
        return "comment_anchor_malformed"
    if start_m.paragraph > end_m.paragraph:
        return "comment_anchor_malformed"
    if reference_m.paragraph != end_m.paragraph:
        return "comment_anchor_malformed"
    return None


def _anchor_records(
    groups: Mapping[int, tuple[_Marker, ...]],
    allowed: Mapping[str, _ReviewPart],
    paragraphs: Mapping[str, Sequence[ET.Element]],
    namespaces: Mapping[str, str],
) -> tuple[dict[int, tuple[dict[str, Any], ...]], dict[int, str]]:
    """Resolve every comment's markers to ranges through the review offset mapping."""

    anchors: dict[int, tuple[dict[str, Any], ...]] = {}
    codes: dict[int, str] = {}
    cache: dict[tuple[str, int], _ParagraphMap | None] = {}
    for comment_id, found in groups.items():
        code = _anchor_code(found, allowed)
        if code is None:
            start_m = next(m for m in found if m.kind == "commentRangeStart")
            end_m = next(m for m in found if m.kind == "commentRangeEnd")
            part = start_m.part
            region = start_m.region
            assert start_m.paragraph is not None and end_m.paragraph is not None

            if start_m.paragraph == end_m.paragraph:
                p_idx = start_m.paragraph
                key = (part, p_idx)
                if key not in cache:
                    try:
                        cache[key] = _comment_paragraph_map(
                            paragraphs[part][p_idx - 1],
                            namespaces[part],
                        )
                    except _ReviewContextError:
                        cache[key] = None
                mapping = cache[key]
                if mapping is None:
                    code = "comment_anchor_context_unsupported"
                else:
                    ranges = [
                        item
                        for item in mapping.annotation_ranges
                        if item["kind"] == "comment" and item["id"] == str(comment_id)
                    ]
                    if not ranges:
                        code = "comment_anchor_malformed"
                    else:
                        anchors[comment_id] = (
                            {
                                "part": part,
                                "region": region,
                                "paragraph": p_idx,
                                "start": ranges[0]["start"],
                                "end": ranges[0]["end"],
                                "text": mapping.text[ranges[0]["start"] : ranges[0]["end"]],
                            },
                        )
            else:
                comment_anchors: list[dict[str, Any]] = []
                p_a = start_m.paragraph
                p_b = end_m.paragraph
                for p_idx in range(p_a, p_b + 1):
                    key = (part, p_idx)
                    if key not in cache:
                        try:
                            cache[key] = _comment_paragraph_map(
                                paragraphs[part][p_idx - 1],
                                namespaces[part],
                            )
                        except _ReviewContextError:
                            cache[key] = None
                    mapping = cache[key]
                    if mapping is None:
                        code = "comment_anchor_context_unsupported"
                        break
                    ranges = [
                        item
                        for item in mapping.annotation_ranges
                        if item["kind"] == "comment" and item["id"] == str(comment_id)
                    ]
                    if p_idx == p_a:
                        start_off = ranges[0]["start"] if ranges else 0
                        end_off = len(mapping.text)
                        comment_anchors.append(
                            {
                                "part": part,
                                "region": region,
                                "paragraph": p_idx,
                                "start": start_off,
                                "end": end_off,
                                "text": mapping.text[start_off:end_off],
                            }
                        )
                    elif p_idx == p_b:
                        start_off = 0
                        end_off = ranges[0]["end"] if ranges else len(mapping.text)
                        comment_anchors.append(
                            {
                                "part": part,
                                "region": region,
                                "paragraph": p_idx,
                                "start": start_off,
                                "end": end_off,
                                "text": mapping.text[start_off:end_off],
                            }
                        )
                    else:
                        comment_anchors.append(
                            {
                                "part": part,
                                "region": region,
                                "paragraph": p_idx,
                                "start": 0,
                                "end": len(mapping.text),
                                "text": mapping.text,
                            }
                        )
                if code is None:
                    anchors[comment_id] = tuple(comment_anchors)
        if code is not None:
            codes[comment_id] = code
    return anchors, codes


def _comment_surface(
    parts: Mapping[str, bytes],
    graph: PackageGraph,
    allowed: Mapping[str, _ReviewPart],
    roots: dict[str, ET.Element],
    namespaces: dict[str, str],
) -> tuple[_CommentSurface, str | None]:
    """Index one package's comments part, its comments, and every comment marker it holds.

    The parsed part roots are cached in the caller's ``roots`` and ``namespaces``, so
    the comment index, the requested anchors, and the parts this request rewrites are
    one tree per part: a marker can never be removed from a copy while the original is
    published.
    """
    comments, code = _resolve_comments_part(graph)
    if code is not None:
        return _CommentSurface(None, (), {}, frozenset(), ()), code
    if comments is not None:
        roots[comments.part] = comments.root
        namespaces[comments.part] = comments.namespace
    ext_target, ext_root, ext_code = resolve_metadata_part(
        graph, _COMMENTS_EXTENDED_RELATIONSHIP, _COMMENTS_EXTENDED_CONTENT_TYPE
    )
    if ext_code is not None:
        return _CommentSurface(None, (), {}, frozenset(), ()), ext_code
    if ext_target is not None and ext_root is not None:
        roots[ext_target] = ext_root
        namespaces[ext_target] = _W15
    ids_target, ids_root, ids_code = resolve_metadata_part(
        graph, _COMMENTS_IDS_RELATIONSHIP, _COMMENTS_IDS_CONTENT_TYPE
    )
    if ids_code is not None:
        return _CommentSurface(None, (), {}, frozenset(), ()), ids_code
    if ids_target is not None and ids_root is not None:
        roots[ids_target] = ids_root
        namespaces[ids_target] = _W16CID
    identity_code = comment_identity_code(
        comments.root if comments is not None else None,
        comments.namespace if comments is not None else "",
        ext_root,
        ids_root,
    )
    if identity_code is not None:
        return _CommentSurface(None, (), {}, frozenset(), ()), identity_code
    unsupported: list[dict[str, Any]] = []
    markers: list[_Marker] = []
    paragraphs: dict[str, tuple[ET.Element, ...]] = {}
    people = resolve_extension_part(graph, PEOPLE)
    extensible = resolve_extension_part(graph, EXTENSIBLE)
    tasks = resolve_extension_part(graph, TASKS)
    for kind, part in ((PEOPLE, people), (EXTENSIBLE, extensible), (TASKS, tasks)):
        if part.code is not None:
            unsupported.append({"part": None, "kind": kind.name, "code": part.code})
        elif kind is not TASKS and part.target is not None and part.root is not None:
            roots[part.target] = part.root
            namespaces[part.target] = kind.namespace
    index = ExtensionIndex(people.root, extensible.root, tasks.root)
    durable_by_para = {
        row.get(Q(_W16CID, "paraId"), "").upper(): row.get(Q(_W16CID, "durableId"))
        for row in (ids_root.findall(Q(_W16CID, "commentId")) if ids_root is not None else ())
    }
    for name in sorted(parts):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        if name in roots:
            root, namespace = roots[name], namespaces[name]
        else:
            try:
                root = parse_xml_with_preserved_qnames(parts[name])
            except ET.ParseError:
                unsupported.append({"part": name, "code": "comment_anchor_part_unparsable"})
                continue
            namespace = _word_namespace(root)
            if namespace is None:
                continue
            roots[name] = root
            namespaces[name] = namespace
        region = allowed[name].region if name in allowed else None
        part_markers, part_paragraphs = _part_markers(name, region, root, namespace)
        paragraphs[name] = tuple(part_paragraphs)
        markers.extend(part_markers)
    groups: dict[int, list[_Marker]] = {}
    for marker in markers:
        if marker.comment_id is None:
            unsupported.append(
                {
                    "part": marker.part,
                    "region": marker.region,
                    "kind": marker.kind,
                    "comment_id": None,
                    "code": "comment_anchor_id_unsupported",
                }
            )
            continue
        groups.setdefault(marker.comment_id, []).append(marker)
    anchors, codes = _anchor_records(
        {key: tuple(value) for key, value in groups.items()},
        allowed,
        paragraphs,
        namespaces,
    )
    identifiers: set[int] = set(groups)
    records: list[_CommentRecord] = []
    if comments is not None:
        elements: list[tuple[int, ET.Element]] = []
        duplicates: set[int] = set()
        comment_ids: set[int] = set()
        parents, resolutions, comment_to_para, _ = parse_thread_state(
            comments.root, comments.namespace, ext_root
        )
        for element in comments.root.findall(Q(comments.namespace, "comment")):
            comment_id = _raw_id(element.get(Q(comments.namespace, "id")))
            if comment_id is None:
                unsupported.append(
                    {
                        "part": comments.part,
                        "comment_id": None,
                        "code": "comment_id_unsupported",
                    }
                )
                continue
            if comment_id in comment_ids:
                duplicates.add(comment_id)
            comment_ids.add(comment_id)
            identifiers.add(comment_id)
            elements.append((comment_id, element))
        for comment_id, element in elements:
            text, body_code = _comment_body(element, comments.namespace)
            parent_comment_id = parents.get(comment_id)
            resolved = resolutions.get(comment_id, False)
            para_id = comment_to_para.get(comment_id)
            durable_id = durable_by_para.get(para_id) if para_id is not None else None
            identity = {
                "author": element.get(Q(comments.namespace, "author")),
                "initials": element.get(Q(comments.namespace, "initials")),
                "date": element.get(Q(comments.namespace, "date")),
                "text": text,
                "parent_comment_id": parent_comment_id,
                "resolved": resolved,
                "anchors": list(anchors.get(comment_id, ())),
            }
            ambiguous = "comment_id_ambiguous" if comment_id in duplicates else None
            if parent_comment_id is not None:
                anchor_code = ambiguous or codes.get(comment_id)
            else:
                anchor_code = (
                    ambiguous
                    or codes.get(comment_id)
                    or (None if anchors.get(comment_id) else "comment_anchor_missing")
                )
            records.append(
                _CommentRecord(
                    comment_id,
                    element,
                    identity["author"],
                    identity["initials"],
                    identity["date"],
                    identity["text"],
                    ambiguous or body_code,
                    anchors.get(comment_id, ()),
                    anchor_code,
                    _comment_fingerprint(comment_id, identity),
                    parent_comment_id,
                    resolved,
                    para_id,
                    extensions=index.fields(durable_id, identity["author"]),
                )
            )
    return (
        _CommentSurface(
            comments,
            tuple(records),
            paragraphs,
            frozenset(identifiers),
            tuple(unsupported),
            ext_target,
            ids_target,
            people,
            extensible,
        ),
        None,
    )


def discover_comment_targets(source: str | Path, *, part: str = _DOCUMENT) -> dict[str, Any]:
    """List source-bound comment anchors and the paragraph targets a comment can bind to.

    Read-only: the source package is opened and hashed, never rewritten. ``part``
    selects the body, one referenced header or footer, or one relationship-reachable
    footnotes or endnotes part whose paragraph targets are listed; the comment index
    always covers every package part, because a comment is anchored wherever its
    markers are and is named by its own package-unique ID. Every comment carries its
    author, initials, date, visible body text, its resolved anchors, and the one
    content-and-anchor fingerprint that an update or a delete must present back. Every
    paragraph target carries the same fields the review surface reports, so an
    editable range, its text, and its structure fingerprint bind a new comment exactly
    as they bind a tracked edit. A comment this engine cannot rewrite or remove is
    reported with the code that explains why and is never advertised as editable, and
    marker or comment markup that cannot be named at all is reported under
    ``unsupported``.
    """

    path = Path(source)
    try:
        graph = PackageGraph.from_file(path)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return _discovery_refusal("comment_source_package_invalid")
    try:
        document_root = ET.fromstring(parts[_DOCUMENT])
    except (KeyError, ET.ParseError):
        return _discovery_refusal("comment_source_package_invalid")
    namespace = _word_namespace(document_root)
    if namespace is None:
        return _discovery_refusal("comment_namespace_unsupported")
    allowed = _review_parts(graph, document_root, namespace)
    if part not in allowed:
        return {
            "passed": False,
            "code": "comment_part_unsupported",
            "source_sha256": graph.file_sha256,
            **_part_context(part, allowed),
            "comments": [],
            "targets": [],
            "unsupported": [],
        }
    roots: dict[str, ET.Element] = {}
    namespaces: dict[str, str] = {}
    surface, code = _comment_surface(parts, graph, allowed, roots, namespaces)
    if code is not None:
        return {
            "passed": False,
            "code": code,
            "source_sha256": graph.file_sha256,
            **_part_context(part, allowed),
            "comments": [],
            "targets": [],
            "unsupported": [],
        }
    try:
        root = ET.fromstring(parts[part])
    except (KeyError, ET.ParseError):
        return _discovery_refusal("comment_source_package_invalid")
    part_namespace = _word_namespace(root)
    if part_namespace is None:
        return _discovery_refusal("comment_namespace_unsupported")
    parents = {id(child): parent for parent in root.iter() for child in parent}
    targets = _paragraph_target_records(part, root, part_namespace, allowed[part], parents)
    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        **_part_context(part, allowed),
        "comments": [record.to_dict() for record in surface.records],
        "targets": targets,
        "unsupported": list(surface.unsupported),
    }


def _discovery_refusal(code: str) -> dict[str, Any]:
    return {
        "passed": False,
        "code": code,
        "comments": [],
        "targets": [],
        "unsupported": [],
    }


def _comment_refusal(request: CommentRequest, code: str) -> CommentReport:
    return CommentReport(
        _REFUSED,
        code,
        request.source_sha256,
        None,
        [],
        [edit.target_dict() for edit in request.edits],
    )


def _insert_anchor_code(mapping: _ParagraphMap, edit: CommentInsert) -> str | None:
    """Why one requested anchor cannot become a comment range on this paragraph."""

    if edit.end > len(mapping.text):
        return "comment_target_text_unsupported"
    if mapping.text[edit.start : edit.end] != edit.expected_old_text:
        return "comment_target_text_unsupported"
    if not any(start <= edit.start and edit.end <= end for start, end in mapping.editable_ranges):
        return "comment_target_text_unsupported"
    ranges = [item for item in mapping.annotation_ranges if item["kind"] == "comment"]
    if any(item["start"] < edit.end and edit.start < item["end"] for item in ranges):
        return "comment_anchor_intersection_unsupported"
    if any(edit.start in {item["start"], item["end"]} for item in ranges) or any(
        edit.end in {item["start"], item["end"]} for item in ranges
    ):
        return "comment_anchor_boundary_ambiguous"
    if any(
        edit.start < item.start < edit.end or edit.start < item.end < edit.end
        for item in mapping.protected_ranges
    ):
        return "comment_revision_intersection_unsupported"
    return None


def _range_marker(namespace: str, local_name: str, comment_id: int) -> ET.Element:
    return ET.Element(Q(namespace, local_name), {Q(namespace, "id"): str(comment_id)})


def _reference_run(namespace: str, comment_id: int) -> ET.Element:
    """The body run that points at one comment, exactly as Word writes it."""

    run = ET.Element(Q(namespace, "r"))
    properties = ET.SubElement(run, Q(namespace, "rPr"))
    ET.SubElement(properties, Q(namespace, "rStyle")).set(
        Q(namespace, "val"), _COMMENT_REFERENCE_STYLE
    )
    ET.SubElement(run, Q(namespace, "commentReference")).set(Q(namespace, "id"), str(comment_id))
    return run


def _text_run(namespace: str, text: str) -> ET.Element:
    run = ET.Element(Q(namespace, "r"))
    run.append(_text_node(namespace, text))
    return run


def _comment_element(
    namespace: str,
    comment_id: int,
    edit: CommentInsert | CommentReply,
    para_id: str | None = None,
) -> ET.Element:
    """One new comment body: the marker run Word writes plus the authored text run."""

    attributes = {
        Q(namespace, "id"): str(comment_id),
        Q(namespace, "author"): edit.author,
        Q(namespace, "date"): edit.date,
    }
    if edit.initials is not None:
        attributes[Q(namespace, "initials")] = edit.initials
    element = ET.Element(Q(namespace, "comment"), attributes)
    p_attribs = {f"{{{_W14}}}paraId": para_id} if para_id else {}
    paragraph = ET.SubElement(element, Q(namespace, "p"), p_attribs)
    marker = ET.SubElement(paragraph, Q(namespace, "r"))
    properties = ET.SubElement(marker, Q(namespace, "rPr"))
    ET.SubElement(properties, Q(namespace, "rStyle")).set(
        Q(namespace, "val"), _COMMENT_REFERENCE_STYLE
    )
    ET.SubElement(marker, Q(namespace, "annotationRef"))
    paragraph.append(_text_run(namespace, edit.text))
    return element


def _anchored_children(
    paragraph: ET.Element,
    mapping: _ParagraphMap,
    starts: Mapping[int, Sequence[int]],
    ends: Mapping[int, Sequence[int]],
    namespace: str,
) -> list[ET.Element]:
    """Rewrite one paragraph's children with the markers of new comment anchors.

    A comment marker carries no text, so the paragraph's own text offsets do not move
    and every requested anchor stays valid. Runs the anchors do not touch are copied
    byte for byte, a run that an anchor splits keeps its formatting on both sides, and
    each comment range closes with its reference run before any text that follows it.
    A point comment (``start == end``) is written as Word writes one: its range start,
    its range end, and its reference run, adjacent at its offset, after every range
    that closes there and before every range that opens there.
    """

    runs = {id(run.element): run for run in mapping.runs}
    boundaries = sorted(set(starts.keys()) | set(ends.keys()))
    result: list[ET.Element] = []

    if not mapping.runs:
        result = [copy.deepcopy(child) for child in paragraph]
        head = 1 if result and result[0].tag == Q(namespace, "pPr") else 0
        for comment_id in starts.get(0, ()):
            result.insert(head, _range_marker(namespace, "commentRangeStart", comment_id))
            head += 1
        for comment_id in ends.get(0, ()):
            result.insert(head, _range_marker(namespace, "commentRangeEnd", comment_id))
            head += 1
            result.insert(head, _reference_run(namespace, comment_id))
            head += 1
        return result

    points: dict[int, list[int]] = {}
    for offset, identifiers in starts.items():
        closing = set(ends.get(offset, ()))
        for comment_id in identifiers:
            if comment_id in closing:
                points.setdefault(offset, []).append(comment_id)
    if points:
        point_ids = {comment_id for identifiers in points.values() for comment_id in identifiers}
        starts = {
            offset: [item for item in identifiers if item not in point_ids]
            for offset, identifiers in starts.items()
        }
        ends = {
            offset: [item for item in identifiers if item not in point_ids]
            for offset, identifiers in ends.items()
        }
    emitted: set[int] = set()

    def emit_points(offset: int) -> None:
        if offset in emitted or offset not in points:
            return
        emitted.add(offset)
        for comment_id in points[offset]:
            result.append(_range_marker(namespace, "commentRangeStart", comment_id))
            result.append(_range_marker(namespace, "commentRangeEnd", comment_id))
            result.append(_reference_run(namespace, comment_id))

    for child in paragraph:
        run = runs.get(id(child))
        if run is None or not run.text:
            result.append(copy.deepcopy(child))
            continue
        cuts = [
            run.start,
            *(item for item in boundaries if run.start < item < run.end),
            run.end,
        ]
        for lower, upper in itertools.pairwise(cuts):
            for comment_id in ends.get(lower, ()):
                if lower == 0 and not any(
                    c.tag == Q(namespace, "commentRangeEnd")
                    and c.get(Q(namespace, "id")) == str(comment_id)
                    for c in result
                ):
                    result.append(_range_marker(namespace, "commentRangeEnd", comment_id))
                    result.append(_reference_run(namespace, comment_id))
            emit_points(lower)
            for comment_id in starts.get(lower, ()):
                result.append(_range_marker(namespace, "commentRangeStart", comment_id))
            result.append(
                _clone_run(child, namespace, run.text[lower - run.start : upper - run.start])
            )
            for comment_id in ends.get(upper, ()):
                result.append(_range_marker(namespace, "commentRangeEnd", comment_id))
                result.append(_reference_run(namespace, comment_id))
            emit_points(upper)
    return result


def _marker_count(children: list[ET.Element], namespace: str, comment_id: int) -> int:
    return sum(
        child.tag in {Q(namespace, name) for name in _MARKER_TAGS}
        and child.get(Q(namespace, "id")) == str(comment_id)
        for child in children
    ) + sum(
        1
        for child in children
        for node in child
        if node.tag == Q(namespace, "commentReference")
        and node.get(Q(namespace, "id")) == str(comment_id)
    )


def _remove_anchors(paragraph: ET.Element, namespace: str, comment_id: int) -> None:
    """Remove one comment's three markers from the paragraph that owns them."""

    identifier = str(comment_id)
    for child in list(paragraph):
        if child.tag in {Q(namespace, name) for name in _RANGE_TAGS} and (
            child.get(Q(namespace, "id")) == identifier
        ):
            paragraph.remove(child)
            continue
        if child.tag == Q(namespace, "r") and any(
            node.tag == Q(namespace, "commentReference")
            and node.get(Q(namespace, "id")) == identifier
            for node in child
        ):
            paragraph.remove(child)


def _rewrite_comment(element: ET.Element, edit: CommentUpdate, namespace: str) -> None:
    """Rewrite one comment's author, date, initials, and body text in place.

    The comment element and its body paragraph are kept, so the paragraph identity
    that the modern comment metadata parts are keyed by does not move.
    """

    element.set(Q(namespace, "author"), edit.author)
    element.set(Q(namespace, "date"), edit.date)
    if edit.initials is not None:
        element.set(Q(namespace, "initials"), edit.initials)
    paragraph = next(child for child in element if child.tag == Q(namespace, "p"))
    children = list(paragraph)
    head = 1 if children and children[0].tag == Q(namespace, "pPr") else 0
    if len(children) > head and _annotation_reference_run(children[head], namespace):
        head += 1
    body = children[head:]
    template = next((child for child in body if child.tag == Q(namespace, "r")), None)
    for child in body:
        paragraph.remove(child)
    paragraph.append(
        _clone_run(template, namespace, edit.text)
        if template is not None
        else _text_run(namespace, edit.text)
    )


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is the source part, byte for byte."""

    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        names = set(after.namelist())
        return all(
            info.filename in names and before.read(info.filename) == after.read(info.filename)
            for info in before.infolist()
            if info.filename not in changed
        )


def _comment_index_key(comment_id: int, identity: Mapping[str, Any]) -> tuple[Any, ...]:
    """The comparable identity of one comment: its ID, its content, and its anchors."""

    return (
        comment_id,
        identity["author"],
        identity["initials"],
        identity["date"],
        identity["text"],
        identity.get("parent_comment_id"),
        bool(identity.get("resolved", False)),
        tuple(
            sorted(
                (
                    anchor["part"],
                    anchor["paragraph"],
                    anchor["start"],
                    anchor["end"],
                    anchor["text"],
                )
                for anchor in identity["anchors"]
            )
        ),
    )


def _result_entry(comment_id: int, identity: Mapping[str, Any]) -> dict[str, Any]:
    """One comment of the published index, with the fingerprint a follow-up edit needs."""

    return {
        "comment_id": comment_id,
        **identity,
        "comment_sha256": _comment_fingerprint(comment_id, identity),
    }


def _report(
    published: ReviewReport,
    comments: list[dict[str, Any]],
    targets: list[dict[str, Any]],
) -> CommentReport:
    """The comment report of one publication decision, read off the shared publisher."""

    return CommentReport(
        published.status,
        published.code,
        published.source_sha256,
        published.output_sha256,
        comments,
        targets,
    )


def _verifier(
    source: Path, changed: frozenset[str], expected: Mapping[int, dict[str, Any]]
) -> Callable[[Path], str | None]:
    """The independent postcondition of one comment publication.

    The staged package is re-read, its comments are re-discovered from scratch, and
    the whole resulting index - every comment ID, author, date, text, and anchor -
    must equal the index the request described, with every part the request did not
    change still byte-identical to the source.
    """

    wanted = Counter(_comment_index_key(item, value) for item, value in expected.items())

    def verify(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, changed):
            return "comment_unrelated_part_changed"
        discovery = discover_comment_targets(staged)
        if not discovery["passed"]:
            return "comment_staged_verification_failed"
        found = Counter(
            _comment_index_key(item["comment_id"], item) for item in discovery["comments"]
        )
        return None if found == wanted else "comment_staged_verification_failed"

    return verify


def _free_part_name(base: str, parts: Mapping[str, bytes]) -> str:
    """``base`` when no package part holds it, else the first free ``-N`` sibling name."""

    stem, _, suffix = base.rpartition(".")
    name, serial = base, 2
    while name in parts:
        name = f"{stem}-{serial}.{suffix}"
        serial += 1
    return name


@dataclass(frozen=True)
class _ResolvedInsert:
    """One insert bound to its live paragraph(s), offset mapping(s), and its comment ID."""

    edit: CommentInsert
    comment_id: int
    paragraph: ET.Element
    mapping: _ParagraphMap
    end_paragraph_elem: ET.Element | None = None
    end_mapping: _ParagraphMap | None = None


def author_comment(request: CommentRequest) -> CommentReport:
    """Resolve every edit on the immutable source, verify once, and publish once.

    The whole request is resolved against the source before anything is written: an
    insert must match the current paragraph structure fingerprint and the current text
    of its range, an update or a delete must match the comment's own content-and-anchor
    fingerprint, and every refusal abandons the request with no output file. Comments
    are deleted before they are inserted, so an anchor and a new range can share one
    paragraph; removing a marker moves no text, and every requested offset stays valid.
    A published package is re-read and re-discovered, and its comment index must equal
    the index the request described, with every part the request did not change still
    byte-identical to the source.
    """

    source, destination = request.source.resolve(), request.destination.resolve()
    if source == destination or not source.is_file() or not destination.parent.is_dir():
        return _comment_refusal(request, "comment_destination_invalid")
    if _sha256(source) != request.source_sha256:
        return _comment_refusal(request, "comment_source_hash_mismatch")
    try:
        graph = PackageGraph.from_file(source)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
        if "[Content_Types].xml" not in parts:
            return _comment_refusal(request, "comment_package_invalid")
        document_root = ET.fromstring(parts[_DOCUMENT])
        document_namespace = _word_namespace(document_root)
        if document_namespace is None:
            return _comment_refusal(request, "comment_namespace_unsupported")
        allowed = _review_parts(graph, document_root, document_namespace)
        roots: dict[str, ET.Element] = {}
        namespaces: dict[str, str] = {}
        surface, code = _comment_surface(parts, graph, allowed, roots, namespaces)
        if code is not None:
            return _comment_refusal(request, code)
        inserts: list[_ResolvedInsert] = []
        replies: list[tuple[CommentReply, int, _CommentRecord]] = []
        updates: list[tuple[CommentUpdate, _CommentRecord]] = []
        resolves: list[tuple[CommentResolve, _CommentRecord]] = []
        deletes: list[tuple[CommentDelete, _CommentRecord]] = []
        targets: list[dict[str, Any]] = []
        next_id = max(surface.ids, default=-1) + 1
        for edit in request.edits:
            if isinstance(edit, CommentInsert):
                if edit.part not in allowed:
                    return _comment_refusal(request, "comment_part_unsupported")
                loaded = _load_part(parts, roots, namespaces, edit.part)
                if loaded is None:
                    return _comment_refusal(request, "comment_package_invalid")
                root, namespace = loaded
                part_paragraphs = surface.paragraphs.get(edit.part, ())
                start_p = edit.paragraph
                end_p = edit.effective_end_paragraph
                if (
                    start_p > len(part_paragraphs)
                    or end_p > len(part_paragraphs)
                    or start_p < 1
                    or end_p < start_p
                ):
                    return _comment_refusal(request, "comment_target_missing")

                para_a = part_paragraphs[start_p - 1]
                fp_a = _mask_text(para_a, namespace)
                para_b = part_paragraphs[end_p - 1]
                fp_b = _mask_text(para_b, namespace)

                if start_p == end_p:
                    if fp_a != edit.structure_fingerprint_sha256:
                        return _comment_refusal(request, "comment_target_stale")
                elif (
                    edit.end_structure_fingerprint_sha256 is None
                    or fp_a != edit.structure_fingerprint_sha256
                    or fp_b != edit.end_structure_fingerprint_sha256
                ):
                    return _comment_refusal(request, "comment_target_stale")

                try:
                    mapping_a = _comment_paragraph_map(para_a, namespace)
                except _ReviewContextError as error:
                    return _comment_refusal(request, error.code)

                if start_p == end_p:
                    anchor_code = _insert_anchor_code(mapping_a, edit)
                    if anchor_code is not None:
                        return _comment_refusal(request, anchor_code)
                    mapping_b = None
                    para_b = None
                else:
                    try:
                        mapping_b = _comment_paragraph_map(para_b, namespace)
                    except _ReviewContextError as error:
                        return _comment_refusal(request, error.code)

                    if edit.start > len(mapping_a.text) or edit.end > len(mapping_b.text):
                        return _comment_refusal(request, "comment_target_text_unsupported")

                    seg_a = mapping_a.text[edit.start :]
                    intermediate_maps: list[_ParagraphMap] = []
                    for k in range(start_p + 1, end_p):
                        try:
                            m_k = _comment_paragraph_map(part_paragraphs[k - 1], namespace)
                        except _ReviewContextError as error:
                            return _comment_refusal(request, error.code)
                        if m_k.protected_ranges:
                            return _comment_refusal(
                                request, "comment_revision_intersection_unsupported"
                            )
                        if any(item["kind"] == "comment" for item in m_k.annotation_ranges):
                            return _comment_refusal(
                                request, "comment_anchor_intersection_unsupported"
                            )
                        intermediate_maps.append(m_k)
                    seg_mids = [mapping.text for mapping in intermediate_maps]
                    seg_b = mapping_b.text[: edit.end]

                    all_segs = [seg_a, *seg_mids, seg_b]
                    if edit.expected_old_text != "\n".join(all_segs):
                        return _comment_refusal(request, "comment_target_text_unsupported")

                    if not any(s <= edit.start <= e for s, e in mapping_a.editable_ranges):
                        return _comment_refusal(request, "comment_target_text_unsupported")
                    if edit.end > 0 and not any(
                        s <= edit.end <= e for s, e in mapping_b.editable_ranges
                    ):
                        return _comment_refusal(request, "comment_target_text_unsupported")
                    if any(edit.start < item.end for item in mapping_a.protected_ranges) or any(
                        item.start < edit.end and 0 < item.end
                        for item in mapping_b.protected_ranges
                    ):
                        return _comment_refusal(
                            request, "comment_revision_intersection_unsupported"
                        )

                    ranges_a = [
                        item for item in mapping_a.annotation_ranges if item["kind"] == "comment"
                    ]
                    if any(
                        item["start"] < len(mapping_a.text) and edit.start < item["end"]
                        for item in ranges_a
                    ):
                        return _comment_refusal(request, "comment_anchor_intersection_unsupported")
                    ranges_b = [
                        item for item in mapping_b.annotation_ranges if item["kind"] == "comment"
                    ]
                    if any(item["start"] < edit.end and 0 < item["end"] for item in ranges_b):
                        return _comment_refusal(request, "comment_anchor_intersection_unsupported")

                inserts.append(
                    _ResolvedInsert(
                        edit,
                        next_id,
                        para_a,
                        mapping_a,
                        para_b,
                        mapping_b,
                    )
                )
                targets.append({**edit.target_dict(), "comment_id": next_id})
                next_id += 1
                continue
            record = next(
                (item for item in surface.records if item.comment_id == edit.comment_id),
                None,
            )
            if record is None:
                return _comment_refusal(request, "comment_target_missing")
            if record.fingerprint != edit.comment_sha256:
                return _comment_refusal(request, "comment_target_stale")
            if isinstance(edit, (CommentReply, CommentResolve)):
                if record.body_code == "comment_id_ambiguous":
                    return _comment_refusal(request, "comment_id_ambiguous")
                if last_comment_paragraph(record.element, surface.part.namespace) is None:
                    return _comment_refusal(request, "comment_body_unsupported")
            if isinstance(edit, CommentReply):
                if record.parent_comment_id is not None:
                    return _comment_refusal(request, "comment_thread_reply_to_reply_unsupported")
                replies.append((edit, next_id, record))
                targets.append({**edit.target_dict(), "reply_comment_id": next_id})
                next_id += 1
                continue
            if isinstance(edit, CommentResolve):
                resolves.append((edit, record))
                targets.append(edit.target_dict())
                continue
            refusal = record.body_code if isinstance(edit, CommentUpdate) else record.anchor_code
            if refusal is not None:
                return _comment_refusal(request, refusal)
            if isinstance(edit, CommentUpdate):
                updates.append((edit, record))
            else:
                deletes.append((edit, record))
            targets.append(edit.target_dict())
        deleted_ids = {edit.comment_id for edit, _ in deletes}
        for edit, record in deletes:
            for other in surface.records:
                if (
                    other.parent_comment_id == edit.comment_id
                    and other.comment_id not in deleted_ids
                ):
                    return _comment_refusal(request, "comment_thread_dangling_reply")
        if replies:
            extension_code = surface.people.code or surface.extensible.code
            if extension_code is not None:
                return _comment_refusal(request, extension_code)
        for index, item in enumerate(inserts):
            for previous in inserts[:index]:
                if previous.edit.part != item.edit.part:
                    continue
                p_start_a, p_end_a = (
                    item.edit.paragraph,
                    item.edit.effective_end_paragraph,
                )
                p_start_b, p_end_b = (
                    previous.edit.paragraph,
                    previous.edit.effective_end_paragraph,
                )
                if max(p_start_a, p_start_b) > min(p_end_a, p_end_b):
                    continue
                if p_start_a == p_end_a == p_start_b == p_end_b:
                    if item.edit.start < previous.edit.end and previous.edit.start < item.edit.end:
                        return _comment_refusal(request, "comment_anchor_overlap")
                else:
                    return _comment_refusal(request, "comment_anchor_overlap")

        para_starts_by_offset: dict[tuple[str, int], dict[int, list[int]]] = {}
        para_ends_by_offset: dict[tuple[str, int], dict[int, list[int]]] = {}
        for item in inserts:
            part = item.edit.part
            start_p = item.edit.paragraph
            end_p = item.edit.effective_end_paragraph
            para_starts_by_offset.setdefault((part, start_p), {}).setdefault(
                item.edit.start, []
            ).append(item.comment_id)
            para_ends_by_offset.setdefault((part, end_p), {}).setdefault(item.edit.end, []).append(
                item.comment_id
            )

        all_affected_paras = sorted(set(para_starts_by_offset) | set(para_ends_by_offset))
        comments = surface.part
        if comments is None and (inserts or replies) and _COMMENTS_PART in parts:
            return _comment_refusal(request, "comment_part_conflict")
        parents = {
            part: {id(child): parent for parent in roots[part].iter() for child in parent}
            for part in {item.edit.part for item in inserts}
        }
        changed: set[str] = set()
        result: dict[int, dict[str, Any]] = {
            record.comment_id: record.identity() for record in surface.records
        }
        occupied_para_ids = collect_existing_para_ids(roots)
        occupied_durable_ids = collect_existing_durable_ids(roots)

        ext_target = surface.extended_target
        if ext_target is not None:
            ext_root = roots[ext_target]
        elif replies or resolves:
            ext_target = _COMMENTS_EXTENDED_PART
            serial = 2
            while ext_target in parts:
                ext_target = f"word/commentsExtended-{serial}.xml"
                serial += 1
            ext_root = ET.Element(f"{{{_W15}}}commentsEx")
            roots[ext_target] = ext_root
            namespaces[ext_target] = _W15
        else:
            ext_root = None

        ids_target = surface.ids_target
        if ids_target is not None:
            ids_root = roots[ids_target]
        elif replies:
            ids_target = _COMMENTS_IDS_PART
            serial = 2
            while ids_target in parts:
                ids_target = f"word/commentsIds-{serial}.xml"
                serial += 1
            ids_root = ET.Element(f"{{{_W16CID}}}commentsIds")
            roots[ids_target] = ids_root
            namespaces[ids_target] = _W16CID
        else:
            ids_root = None

        people_target, people_root = surface.people.target, surface.people.root
        extensible_target, extensible_root = surface.extensible.target, surface.extensible.root
        if replies and people_root is None:
            people_target = _free_part_name(PEOPLE.default_part, parts)
            people_root = new_root(PEOPLE)
            roots[people_target] = people_root
            namespaces[people_target] = PEOPLE.namespace
        if replies and extensible_root is None:
            extensible_target = _free_part_name(EXTENSIBLE.default_part, parts)
            extensible_root = new_root(EXTENSIBLE)
            roots[extensible_target] = extensible_root
            namespaces[extensible_target] = EXTENSIBLE.namespace

        for edit, record in deletes:
            for anchor in record.anchors:
                _remove_anchors(
                    surface.paragraphs[anchor["part"]][anchor["paragraph"] - 1],
                    namespaces[anchor["part"]],
                    edit.comment_id,
                )
                changed.add(anchor["part"])
            comments.root.remove(record.element)
            result.pop(edit.comment_id, None)
        for edit, record in updates:
            _rewrite_comment(record.element, edit, comments.namespace)
            current = record.identity()
            result[edit.comment_id] = {
                "author": edit.author,
                "initials": edit.initials if edit.initials is not None else current["initials"],
                "date": edit.date,
                "text": edit.text,
                "parent_comment_id": current["parent_comment_id"],
                "resolved": current["resolved"],
                "anchors": current["anchors"],
            }
        for edit, record in resolves:
            para_id = record.para_id
            if para_id is None:
                para_id = generate_collision_safe_para_id(
                    occupied_para_ids, f"res_{edit.comment_id}"
                )
                p = last_comment_paragraph(record.element, comments.namespace)
                if p is not None:
                    p.attrib[f"{{{_W14}}}paraId"] = para_id
                    changed.add(comments.part)
            if ext_root is not None:
                matched_ex = None
                for child in ext_root:
                    pid = child.get(f"{{{_W15}}}paraId")
                    if pid and pid.upper() == para_id.upper():
                        matched_ex = child
                        break
                if matched_ex is not None:
                    matched_ex.set(f"{{{_W15}}}done", "1" if edit.resolved else "0")
                else:
                    new_ex = ET.Element(
                        f"{{{_W15}}}commentEx",
                        {
                            f"{{{_W15}}}paraId": para_id,
                            f"{{{_W15}}}done": "1" if edit.resolved else "0",
                        },
                    )
                    ext_root.append(new_ex)
                changed.add(ext_target)
            result[edit.comment_id]["resolved"] = edit.resolved
        if comments is not None and (deletes or updates or inserts or replies or resolves):
            changed.add(comments.part)
        insert_para_ids: dict[int, str] = {}
        total_starts: Counter[int] = Counter()
        total_ends: Counter[int] = Counter()
        total_refs: Counter[int] = Counter()

        for part, p_idx in all_affected_paras:
            paragraph = surface.paragraphs[part][p_idx - 1]
            mapping = _comment_paragraph_map(paragraph, namespaces[part])
            starts_map = para_starts_by_offset.get((part, p_idx), {})
            ends_map = para_ends_by_offset.get((part, p_idx), {})
            children = _anchored_children(
                paragraph, mapping, starts_map, ends_map, namespaces[part]
            )
            for child in children:
                if child.tag == Q(namespaces[part], "commentRangeStart"):
                    cid = _raw_id(child.get(Q(namespaces[part], "id")))
                    if cid is not None:
                        total_starts[cid] += 1
                elif child.tag == Q(namespaces[part], "commentRangeEnd"):
                    cid = _raw_id(child.get(Q(namespaces[part], "id")))
                    if cid is not None:
                        total_ends[cid] += 1
                elif child.tag == Q(namespaces[part], "r"):
                    for node in child:
                        if node.tag == Q(namespaces[part], "commentReference"):
                            cid = _raw_id(node.get(Q(namespaces[part], "id")))
                            if cid is not None:
                                total_refs[cid] += 1

            parent = parents[part][id(paragraph)]
            index = list(parent).index(paragraph)
            replacement = copy.copy(paragraph)
            replacement[:] = children
            replacement.text = replacement.tail = None
            parent.remove(paragraph)
            parent.insert(index, replacement)
            changed.add(part)

        for item in inserts:
            cid = item.comment_id
            if total_starts[cid] != 1 or total_ends[cid] != 1 or total_refs[cid] != 1:
                return _comment_refusal(request, "comment_anchor_unsupported")

            part = item.edit.part
            start_p = item.edit.paragraph
            end_p = item.edit.effective_end_paragraph
            if start_p == end_p:
                anchors_list = [
                    {
                        "part": part,
                        "region": allowed[part].region,
                        "paragraph": start_p,
                        "start": item.edit.start,
                        "end": item.edit.end,
                        "text": item.edit.expected_old_text,
                    }
                ]
            else:
                anchors_list = [
                    {
                        "part": part,
                        "region": allowed[part].region,
                        "paragraph": start_p,
                        "start": item.edit.start,
                        "end": len(item.mapping.text),
                        "text": item.mapping.text[item.edit.start :],
                    }
                ]
                for mid_idx in range(start_p + 1, end_p):
                    mid_p = surface.paragraphs[part][mid_idx - 1]
                    mid_map = _comment_paragraph_map(mid_p, namespaces[part])
                    anchors_list.append(
                        {
                            "part": part,
                            "region": allowed[part].region,
                            "paragraph": mid_idx,
                            "start": 0,
                            "end": len(mid_map.text),
                            "text": mid_map.text,
                        }
                    )
                anchors_list.append(
                    {
                        "part": part,
                        "region": allowed[part].region,
                        "paragraph": end_p,
                        "start": 0,
                        "end": item.edit.end,
                        "text": item.end_mapping.text[: item.edit.end]
                        if item.end_mapping is not None
                        else "",
                    }
                )

            result[cid] = {
                "author": item.edit.author,
                "initials": item.edit.initials,
                "date": item.edit.date,
                "text": item.edit.text,
                "parent_comment_id": None,
                "resolved": False,
                "anchors": anchors_list,
            }
        if comments is not None:
            for item in inserts:
                comments.root.append(
                    _comment_element(
                        comments.namespace,
                        item.comment_id,
                        item.edit,
                        insert_para_ids.get(item.comment_id),
                    )
                )
        for edit, reply_id, parent_record in replies:
            parent_para_id = parent_record.para_id
            if parent_para_id is None:
                parent_para_id = generate_collision_safe_para_id(
                    occupied_para_ids, f"parent_{parent_record.comment_id}"
                )
                p = last_comment_paragraph(parent_record.element, comments.namespace)
                if p is not None:
                    p.attrib[f"{{{_W14}}}paraId"] = parent_para_id
                    changed.add(comments.part)
                if ext_root is not None:
                    ext_root.append(
                        ET.Element(
                            f"{{{_W15}}}commentEx",
                            {
                                f"{{{_W15}}}paraId": parent_para_id,
                                f"{{{_W15}}}done": "0",
                            },
                        )
                    )
                    changed.add(ext_target)
            reply_para_id = generate_collision_safe_para_id(occupied_para_ids, f"rep_{reply_id}")
            reply_elem = _comment_element(
                comments.namespace if comments is not None else document_namespace,
                reply_id,
                edit,
                reply_para_id,
            )
            if comments is not None:
                comments.root.append(reply_elem)
                changed.add(comments.part)
            if ext_root is not None:
                reply_ex = ET.Element(
                    f"{{{_W15}}}commentEx",
                    {
                        f"{{{_W15}}}paraId": reply_para_id,
                        f"{{{_W15}}}paraIdParent": parent_para_id,
                        f"{{{_W15}}}done": "0",
                    },
                )
                ext_root.append(reply_ex)
                changed.add(ext_target)
            if ids_root is not None:
                durable_id = generate_collision_safe_durable_id(
                    occupied_durable_ids, f"rep_{reply_id}"
                )
                cid_elem = ET.Element(
                    f"{{{_W16CID}}}commentId",
                    {
                        f"{{{_W16CID}}}paraId": reply_para_id,
                        f"{{{_W16CID}}}durableId": durable_id,
                    },
                )
                ids_root.append(cid_elem)
                changed.add(ids_target)
                if extensible_root is not None and extensible_target is not None:
                    append_extensible(extensible_root, durable_id, edit.date)
                    changed.add(extensible_target)
            if (
                people_root is not None
                and people_target is not None
                and register_person(people_root, edit.author)
            ):
                changed.add(people_target)
            result[reply_id] = {
                "author": edit.author,
                "initials": edit.initials,
                "date": edit.date,
                "text": edit.text,
                "parent_comment_id": parent_record.comment_id,
                "resolved": False,
                "anchors": [],
            }
        replacements: dict[str, bytes] = {
            part: serialize_xml_with_preserved_qnames(roots[part])
            for part in sorted(changed)
            if part in graph.parts
        }
        additions: dict[str, bytes] = {}
        if comments is None and (inserts or replies):
            root = ET.Element(Q(document_namespace, "comments"))
            for item in inserts:
                root.append(
                    _comment_element(
                        document_namespace,
                        item.comment_id,
                        item.edit,
                        insert_para_ids.get(item.comment_id),
                    )
                )
            for edit, reply_id, _ in replies:
                root.append(
                    _comment_element(
                        document_namespace,
                        reply_id,
                        edit,
                        generate_collision_safe_para_id(occupied_para_ids, f"rep_{reply_id}"),
                    )
                )
            additions[_COMMENTS_PART] = serialize_xml_with_preserved_qnames(root)
            _ensure_override(parts, _COMMENTS_PART, _COMMENTS_CONTENT_TYPE)
            _add_document_relationship(parts, _COMMENTS_RELATIONSHIP, _COMMENTS_PART)
            for name in ("[Content_Types].xml", _rels_name(_DOCUMENT)):
                if name in graph.parts:
                    replacements[name] = parts[name]
                else:
                    additions[name] = parts[name]
        for target, root, content_type, relationship in (
            (
                ext_target,
                ext_root,
                _COMMENTS_EXTENDED_CONTENT_TYPE,
                _COMMENTS_EXTENDED_RELATIONSHIP,
            ),
            (ids_target, ids_root, _COMMENTS_IDS_CONTENT_TYPE, _COMMENTS_IDS_RELATIONSHIP),
            (people_target, people_root, PEOPLE.content_type, PEOPLE.relationship),
            (extensible_target, extensible_root, EXTENSIBLE.content_type, EXTENSIBLE.relationship),
        ):
            if target is None or root is None or target not in changed or target in graph.parts:
                continue
            additions[target] = serialize_xml_with_preserved_qnames(root)
            _ensure_override(parts, target, content_type)
            _add_document_relationship(parts, relationship, target)
            for name in ("[Content_Types].xml", _rels_name(_DOCUMENT)):
                if name in graph.parts:
                    replacements[name] = parts[name]
                else:
                    additions[name] = parts[name]
        entries = [_result_entry(item, value) for item, value in sorted(result.items())]
        published = _publish(
            source,
            destination,
            request.source_sha256,
            replacements,
            entries,
            targets,
            additions=additions,
            verify=_verifier(source, frozenset(replacements) | frozenset(additions), result),
            failure_code="comment_staged_verification_failed",
            exists_code="comment_destination_invalid",
            unavailable_code="comment_publication_unavailable",
        )
        return _report(published, entries, targets)
    except (OSError, PackageError, ET.ParseError):
        return _comment_refusal(request, "comment_package_invalid")
