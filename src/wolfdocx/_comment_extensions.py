"""Modern comment extension parts: people, commentsExtensible, and document tasks.

``word/people.xml`` registers the authors of a document, ``word/commentsExtensible.xml``
holds one row per comment durable ID with its UTC date, placeholder flag and reactions,
and the document tasks part anchors tasks to comments by durable ID. This module
resolves those parts, appends the rows an authored reply owes them, and reads them back
as per-comment discovery fields. It never rewrites or discards a row it does not own.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any
from xml.etree import ElementTree as ET

from ._comment_threads import _W15, resolve_metadata_part
from .inspect import Q
from .package import PackageGraph

_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
_CR = "http://schemas.microsoft.com/office/comments/2020/reactions"
_T = "http://schemas.microsoft.com/office/tasks/2019/documenttasks"


@dataclass(frozen=True)
class ExtensionKind:
    """One relationship-reachable extension part this engine reads or appends to."""

    name: str
    relationship: str
    content_type: str
    default_part: str
    namespace: str
    root: str


PEOPLE = ExtensionKind(
    "people",
    "http://schemas.microsoft.com/office/2011/relationships/people",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
    "word/people.xml",
    _W15,
    "people",
)
EXTENSIBLE = ExtensionKind(
    "commentsExtensible",
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
    "word/commentsExtensible.xml",
    _W16CEX,
    "commentsExtensible",
)
TASKS = ExtensionKind(
    "documentTasks",
    "http://schemas.microsoft.com/office/2019/05/relationships/documenttasks",
    "application/vnd.ms-office.documenttasks+xml",
    "word/documenttasks/documenttasks1.xml",
    _T,
    "Tasks",
)


@dataclass(frozen=True)
class ExtensionPart:
    """One resolved extension part: absent (no target), parsed, or refused with a code."""

    target: str | None = None
    root: ET.Element | None = None
    code: str | None = None


def resolve_extension_part(graph: PackageGraph, kind: ExtensionKind) -> ExtensionPart:
    """Resolve one extension part, refusing an ambiguous, mistyped, or unparsable one."""

    target, root, code = resolve_metadata_part(graph, kind.relationship, kind.content_type)
    if code is None and root is not None and root.tag != Q(kind.namespace, kind.root):
        return ExtensionPart(code="comment_metadata_part_unsupported")
    return ExtensionPart(target, root, code)


def new_root(kind: ExtensionKind) -> ET.Element:
    return ET.Element(Q(kind.namespace, kind.root))


def utc_timestamp(date: str) -> str:
    """The ``w16cex:dateUtc`` form of one comment date; the date carries its own offset."""

    return datetime.fromisoformat(date).astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def register_person(root: ET.Element, author: str) -> bool:
    """Append the author as a ``w15:person`` unless one is registered; True when appended."""

    if any(person.get(Q(_W15, "author")) == author for person in root.findall(Q(_W15, "person"))):
        return False
    person = ET.SubElement(root, Q(_W15, "person"), {Q(_W15, "author"): author})
    ET.SubElement(
        person,
        Q(_W15, "presenceInfo"),
        {Q(_W15, "providerId"): "None", Q(_W15, "userId"): author},
    )
    return True


def append_extensible(root: ET.Element, durable_id: str, date: str) -> None:
    """Append the ``w16cex:commentExtensible`` row of one authored comment."""

    ET.SubElement(
        root,
        Q(_W16CEX, "commentExtensible"),
        {Q(_W16CEX, "durableId"): durable_id, Q(_W16CEX, "dateUtc"): utc_timestamp(date)},
    )


def _flag(value: str | None) -> bool | None:
    return {"1": True, "true": True, "0": False, "false": False}.get(value or "")


def _reactions(row: ET.Element) -> list[dict[str, str | None]]:
    return [
        {
            "reaction_type": reaction.get("reactionType"),
            "date_utc": info.get("dateUtc"),
            "user_id": user.get("userId"),
            "user_provider": user.get("userProvider"),
            "user_name": user.get("userName"),
        }
        for reaction in row.iter(Q(_CR, "reaction"))
        for info in reaction.findall(Q(_CR, "reactionInfo"))
        for user in info.findall(Q(_CR, "user"))
    ]


class ExtensionIndex:
    """Per-comment extension fields, indexed once from the people, extensible and tasks parts."""

    def __init__(
        self, people: ET.Element | None, extensible: ET.Element | None, tasks: ET.Element | None
    ) -> None:
        self._presence: dict[str, dict[str, str | None]] = {}
        for person in people.findall(Q(_W15, "person")) if people is not None else ():
            info = person.find(Q(_W15, "presenceInfo"))
            author = person.get(Q(_W15, "author"))
            if author is not None and info is not None:
                self._presence.setdefault(
                    author,
                    {
                        "provider_id": info.get(Q(_W15, "providerId")),
                        "user_id": info.get(Q(_W15, "userId")),
                    },
                )
        self._extensible: dict[str, ET.Element] = {}
        for row in (
            extensible.findall(Q(_W16CEX, "commentExtensible")) if extensible is not None else ()
        ):
            durable = row.get(Q(_W16CEX, "durableId"))
            if durable is not None:
                self._extensible.setdefault(durable.upper(), row)
        self._tasks: dict[str, list[str]] = {}
        for task in tasks.findall(Q(_T, "Task")) if tasks is not None else ():
            task_id = task.get("id")
            anchor = task.find(Q(_T, "Anchor"))
            if task_id is None or anchor is None:
                continue
            for comment in anchor.findall(Q(_T, "Comment")):
                comment_id = comment.get("id")
                if comment_id is not None:
                    self._tasks.setdefault(comment_id.upper(), []).append(task_id)

    def fields(self, durable_id: str | None, author: str | None) -> dict[str, Any]:
        """The extension fields of one comment, named by its durable ID and its author."""

        row = self._extensible.get(durable_id.upper()) if durable_id is not None else None
        presence = self._presence.get(author) if author is not None else None
        return {
            "durable_id": durable_id,
            "date_utc": row.get(Q(_W16CEX, "dateUtc")) if row is not None else None,
            "intelligent_placeholder": (
                _flag(row.get(Q(_W16CEX, "intelligentPlaceholder"))) if row is not None else None
            ),
            "reactions": _reactions(row) if row is not None else [],
            "author_presence": dict(presence) if presence is not None else None,
            "task_ids": list(self._tasks.get(durable_id.upper(), ())) if durable_id else [],
        }
