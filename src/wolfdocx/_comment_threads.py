"""Internal thread and extended metadata helpers for Word comments.

Provides collision-safe paragraph/durable ID generation, parent-child thread
linkage, resolution tracking, and XML manipulation for word/commentsExtended.xml,
word/commentsIds.xml, and word/people.xml.
"""

from __future__ import annotations

import hashlib
import re
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any
from xml.etree import ElementTree as ET

from .inspect import Q, parse_xml_with_preserved_qnames
from .package import PackageError, PackageGraph

_DOCUMENT = "word/document.xml"

_COMMENTS_EXTENDED_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
_COMMENTS_EXTENDED_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
_COMMENTS_EXTENDED_PART = "word/commentsExtended.xml"

_COMMENTS_IDS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
)
_COMMENTS_IDS_RELATIONSHIP = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
_COMMENTS_IDS_PART = "word/commentsIds.xml"


_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"


@dataclass(frozen=True)
class CommentReply:
    """One reply appended to an existing parent comment thread."""

    comment_id: int
    comment_sha256: str
    author: str
    initials: str | None
    date: str
    text: str

    def target_dict(self) -> dict[str, Any]:
        return {"action": "reply", "comment_id": self.comment_id}


@dataclass(frozen=True)
class CommentResolve:
    """One resolution status change on an existing comment thread."""

    comment_id: int
    comment_sha256: str
    resolved: bool

    def target_dict(self) -> dict[str, Any]:
        return {"action": "resolve", "comment_id": self.comment_id}


def collect_existing_para_ids(roots: Mapping[str, ET.Element]) -> set[str]:
    """Collect all 8-hex paragraph IDs across all parsed parts."""
    para_ids: set[str] = set()
    for root in roots.values():
        for elem in root.iter():
            for key, val in elem.attrib.items():
                if key.endswith(("paraId", "paraIdParent")):
                    para_ids.add(val.upper())
    return para_ids


def generate_collision_safe_para_id(occupied: set[str], seed: str | int) -> str:
    """Generate an 8-character uppercase hex paragraph ID that avoids collision."""
    h = int(hashlib.sha256(f"para_{seed}".encode()).hexdigest()[:8], 16)
    candidate_int = (h & 0x7FFFFFFF) or 1
    candidate = f"{candidate_int:08X}"
    while candidate in occupied or candidate == "00000000":
        candidate_int = ((candidate_int + 1) & 0x7FFFFFFF) or 1
        candidate = f"{candidate_int:08X}"
    occupied.add(candidate)
    return candidate


def collect_existing_durable_ids(roots: Mapping[str, ET.Element]) -> set[str]:
    """Collect all durable IDs across commentsIds.xml."""
    durable_ids: set[str] = set()
    for root in roots.values():
        for elem in root.iter():
            for key, val in elem.attrib.items():
                if key.endswith("durableId"):
                    durable_ids.add(val.upper())
    return durable_ids


def generate_collision_safe_durable_id(occupied: set[str], seed: str | int) -> str:
    """Generate an 8-hex durable ID in the exclusive range (0, 0x7FFFFFFF)."""
    h = int(hashlib.sha256(f"durable_{seed}".encode()).hexdigest()[:8], 16)
    candidate_int = (h & 0x7FFFFFFF) or 1
    if candidate_int == 0x7FFFFFFF:
        candidate_int = 1
    candidate = f"{candidate_int:08X}"
    while candidate in occupied or candidate == "00000000":
        candidate_int = candidate_int + 1 if candidate_int < 0x7FFFFFFE else 1
        candidate = f"{candidate_int:08X}"
    occupied.add(candidate)
    return candidate


def resolve_metadata_part(
    graph: PackageGraph,
    rel_type: str,
    content_type: str,
) -> tuple[str | None, ET.Element | None, str | None]:
    """Resolve one relationship-reachable modern metadata part from document.xml."""
    matches = [
        r for r in graph.relationships_for(_DOCUMENT) if not r.is_external and r.type == rel_type
    ]
    if not matches:
        return None, None, None
    if len(matches) > 1:
        return None, None, "comment_metadata_part_ambiguous"
    target = matches[0].resolved_target()
    pkg_part = graph.part(target)
    if pkg_part is None or pkg_part.content_type != content_type:
        return None, None, "comment_metadata_part_unsupported"
    try:
        data = graph.read_part_bytes(target)
        root = parse_xml_with_preserved_qnames(data)
    except (PackageError, ET.ParseError):
        return None, None, "comment_metadata_part_invalid"
    return target, root, None


def last_comment_paragraph(comment: ET.Element, namespace: str) -> ET.Element | None:
    """The paragraph used by both commentEx and commentId identity associations."""
    last = None
    for paragraph in comment.iter(Q(namespace, "p")):
        last = paragraph
    return last


def comment_identity_code(
    comments_root: ET.Element | None,
    comments_ns: str,
    extended_root: ET.Element | None,
    ids_root: ET.Element | None,
) -> str | None:
    """Diagnose identity ambiguity without rewriting or discarding orphan rows.

    Deleted comments may leave metadata behind. An unbound row is preserved, but
    a row naming a live nonterminal paragraph cannot identify that comment.
    """
    paragraph_ids: set[str] = set()
    terminal_ids: set[str] = set()
    if comments_root is not None:
        for comment in comments_root.findall(Q(comments_ns, "comment")):
            last = last_comment_paragraph(comment, comments_ns)
            for paragraph in comment.iter(Q(comments_ns, "p")):
                value = paragraph.get(Q(_W14, "paraId"))
                if value is None:
                    continue
                if not re.fullmatch(r"[0-9A-Fa-f]{8}", value) or not (
                    0 < int(value, 16) < 0x80000000
                ):
                    return "comment_paragraph_id_invalid"
                normalized = value.upper()
                if normalized in paragraph_ids:
                    return "comment_paragraph_id_ambiguous"
                paragraph_ids.add(normalized)
                if paragraph is last:
                    terminal_ids.add(normalized)

    for root, namespace, root_name, row_name in (
        (extended_root, _W15, "commentsEx", "commentEx"),
        (ids_root, _W16CID, "commentsIds", "commentId"),
    ):
        if root is None:
            continue
        if root.tag != Q(namespace, root_name):
            return "comment_metadata_part_unsupported"
        associations: set[str] = set()
        durable_ids: set[str] = set()
        for row in root.findall(Q(namespace, row_name)):
            value = row.get(Q(namespace, "paraId"), "")
            if not re.fullmatch(r"[0-9A-Fa-f]{8}", value) or not (0 < int(value, 16) < 0x80000000):
                return "comment_paragraph_id_invalid"
            normalized = value.upper()
            if normalized in associations:
                return "comment_identity_association_ambiguous"
            associations.add(normalized)
            if normalized in paragraph_ids and normalized not in terminal_ids:
                return "comment_identity_association_invalid"
            if namespace == _W16CID:
                durable = row.get(Q(namespace, "durableId"), "")
                if not re.fullmatch(r"[0-9A-Fa-f]{8}", durable) or not (
                    0 < int(durable, 16) < 0x7FFFFFFF
                ):
                    return "comment_durable_id_invalid"
                if durable.upper() in durable_ids:
                    return "comment_durable_id_ambiguous"
                durable_ids.add(durable.upper())
            else:
                parent = row.get(Q(namespace, "paraIdParent"))
                if parent is not None:
                    if not re.fullmatch(r"[0-9A-Fa-f]{8}", parent) or not (
                        0 < int(parent, 16) < 0x80000000
                    ):
                        return "comment_paragraph_id_invalid"
                    parent_id = parent.upper()
                    if parent_id in paragraph_ids and parent_id not in terminal_ids:
                        return "comment_identity_association_invalid"
    return None


def parse_thread_state(
    comments_root: ET.Element,
    comments_ns: str,
    extended_root: ET.Element | None,
) -> tuple[dict[int, int | None], dict[int, bool], dict[int, str], dict[str, int]]:
    """Index parent_comment_id, resolved status, and para_ids for all comments.

    Returns:
      parents: comment_id -> parent_comment_id (or None)
      resolutions: comment_id -> bool
      comment_to_para: comment_id -> para_id (str)
      para_to_comment: para_id -> comment_id
    """
    comment_to_para: dict[int, str] = {}
    para_to_comment: dict[str, int] = {}

    for comment_elem in comments_root.findall(Q(comments_ns, "comment")):
        raw_id = comment_elem.get(Q(comments_ns, "id"))
        if raw_id is None or not raw_id.isdigit():
            continue
        cid = int(raw_id)
        p = last_comment_paragraph(comment_elem, comments_ns)
        if p is not None:
            para_id = p.get(f"{{{_W14}}}paraId")
            if para_id:
                normalized = para_id.upper()
                comment_to_para[cid] = normalized
                para_to_comment[normalized] = cid

    parents: dict[int, int | None] = {}
    resolutions: dict[int, bool] = {}

    if extended_root is not None:
        for child in extended_root:
            if child.tag == f"{{{_W15}}}commentEx":
                pid = child.get(f"{{{_W15}}}paraId")
                if pid:
                    pid_norm = pid.upper()
                    done = child.get(f"{{{_W15}}}done") == "1"
                    parent_pid = child.get(f"{{{_W15}}}paraIdParent")
                    parent_norm = parent_pid.upper() if parent_pid else None
                    cid = para_to_comment.get(pid_norm)
                    if cid is not None:
                        resolutions[cid] = done
                        if parent_norm:
                            parents[cid] = para_to_comment.get(parent_norm)
                        else:
                            parents[cid] = None

    for cid in comment_to_para:
        if cid not in parents:
            parents[cid] = None
        if cid not in resolutions:
            resolutions[cid] = False

    return parents, resolutions, comment_to_para, para_to_comment
