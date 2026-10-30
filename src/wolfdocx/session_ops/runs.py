"""Source-bound insertion and deletion of text runs, in a session or file to file.

A run is addressed as a 1-based direct ``w:r`` child of a paragraph; the paragraph
is 1-based over every ``w:p`` of a body, header, or footer part in document order,
as :class:`~wolfdocx.ParagraphRequest` addresses it, and is bound by the fingerprint
:func:`~wolfdocx.paragraph_authoring.paragraph_authoring_structure_fingerprint`
returns. An inserted run takes the anchor run's ``w:rPr`` (without its tracked
property change), so it reads as the text next to it; ``\\t`` and ``\\n`` become
``w:tab`` and ``w:br``. Only runs of plain text and simple controls are deleted;
a run holding a field, note or comment reference, drawing, or object is refused.

:func:`author_run` publishes through :meth:`~wolfdocx.Document.save`, and the session
methods plan with the same :func:`_plan_run_edit`, so both rewrite the paragraph
identically. Only the target paragraph's bytes change within its part.
"""

from __future__ import annotations

import copy
import re
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from xml.parsers import expat

from ..inspect import Q
from ..package import PackageError, PackageGraph

__all__ = ["RunReport", "RunRequest", "author_run"]

_PUBLISHED = "published"
_DOCUMENT = "word/document.xml"
_REFUSED = "refused"
_ACTIONS = ("insert", "delete")
_POSITIONS = ("before", "after")
# Run content a deletion may remove without orphaning a field, reference, or part.
_DELETABLE_RUN_CHILDREN = frozenset(
    {
        "rPr",
        "t",
        "tab",
        "br",
        "cr",
        "noBreakHyphen",
        "softHyphen",
        "lastRenderedPageBreak",
    }
)
_CONTROL_SPLIT = re.compile(r"(\t|\n)")

Loader = Callable[[str], tuple[bytes, ET.Element]]
RunPlan = tuple[bytes, ET.Element] | str


@dataclass(frozen=True)
class RunRequest:
    """One source-bound request to insert or delete a text run.

    ``run`` names the anchor run for an insertion (omitted: append after the last
    run) and the run to delete for a deletion.
    """

    source: Path
    destination: Path
    source_sha256: str
    paragraph: int
    structure_fingerprint_sha256: str
    action: str = "insert"
    run: int | None = None
    position: str = "after"
    text: str | None = None
    part: str = _DOCUMENT

    def __post_init__(self) -> None:
        from ..paragraph_authoring import _require_sha256

        for name in ("source", "destination"):
            value = getattr(self, name)
            if not isinstance(value, (str, Path)) or not str(value):
                raise ValueError(f"run_{name}_invalid")
            object.__setattr__(self, name, Path(value))
        _require_sha256(self.source_sha256, "run_source_sha256")
        _require_sha256(self.structure_fingerprint_sha256, "run_structure_fingerprint_sha256")
        _check_run_fields(
            self.paragraph, self.action, self.run, self.position, self.text, self.part
        )

    def target_dict(self) -> dict[str, Any]:
        return _target(self.part, self.paragraph, self.run, self.action, self.position)


@dataclass(frozen=True)
class RunReport:
    """Content-safe result of one :func:`author_run`."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    target: dict[str, Any]
    action: str

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "run",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "action": self.action,
        }


def _check_run_fields(
    paragraph: Any, action: Any, run: Any, position: Any, text: Any, part: Any
) -> None:
    """Raise the request's validation error; shared by requests and session calls."""
    from ..authoring import _MAX_TEXT_CHARACTERS
    from ..paragraph_authoring import _is_valid_xml_text

    if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 1:
        raise ValueError("run_paragraph_invalid")
    if action not in _ACTIONS:
        raise ValueError("run_action_invalid")
    if run is not None and (isinstance(run, bool) or not isinstance(run, int) or run < 1):
        raise ValueError("run_index_invalid")
    if position not in _POSITIONS:
        raise ValueError("run_position_invalid")
    if not isinstance(part, str) or not part:
        raise ValueError("run_part_invalid")
    if action == "delete":
        if run is None:
            raise ValueError("run_index_missing")
        if text is not None or position != "after":
            raise ValueError("run_request_fields_invalid")
        return
    if run is None and position != "after":
        raise ValueError("run_position_invalid")
    if not isinstance(text, str):
        raise TypeError("run_text_invalid")
    if not text:
        raise ValueError("run_text_empty")
    if len(text) > _MAX_TEXT_CHARACTERS:
        raise ValueError("run_text_limit_exceeded")
    if not _is_valid_xml_text(text):
        raise ValueError("run_text_invalid")


def _target(
    part: str, paragraph: int, run: int | None, action: str, position: str
) -> dict[str, Any]:
    target: dict[str, Any] = {"part": part, "paragraph": paragraph}
    if run is not None:
        target["run"] = run
    if action == "insert":
        target["position"] = position
    return target


def _new_run(text: str, anchor: ET.Element | None, ns: str) -> ET.Element:
    from ..authoring import _XML_SPACE_ATTRIBUTE

    run = ET.Element(Q(ns, "r"))
    properties = anchor.find(Q(ns, "rPr")) if anchor is not None else None
    if properties is not None:
        properties = copy.deepcopy(properties)
        for change in properties.findall(Q(ns, "rPrChange")):
            properties.remove(change)
        run.append(properties)
    for piece in _CONTROL_SPLIT.split(text):
        if piece == "\t":
            ET.SubElement(run, Q(ns, "tab"))
        elif piece == "\n":
            ET.SubElement(run, Q(ns, "br"))
        elif piece:
            node = ET.SubElement(run, Q(ns, "t"))
            node.text = piece
            if piece[:1].isspace() or piece[-1:].isspace():
                node.set(_XML_SPACE_ATTRIBUTE, "preserve")
    return run


def _plan_run_edit(
    graph: PackageGraph,
    load: Loader,
    *,
    paragraph: int,
    action: str,
    run: int | None,
    position: str,
    text: str | None,
    part: str,
    fingerprint: str | None,
) -> RunPlan:
    """Rewrite the target paragraph in the tree ``load`` returns and splice it into
    the part's bytes. Returns the new part bytes and the tree they must parse to, or a
    refusal code. ``fingerprint=None`` binds to the current paragraph."""
    from ..authoring import _cell_structure_fingerprint
    from ..paragraph_authoring import (
        _is_admitted_paragraph,
        _paragraph_xml_fragment,
        _paragraph_xml_span,
        _resolve_target_story,
    )

    root, container, ns, part_bytes, err = _resolve_target_story(graph, part, load)
    if err is not None:
        return "run_" + err.removeprefix("paragraph_")
    assert root is not None and container is not None and ns is not None
    assert part_bytes is not None
    paragraphs = list(root.iter(Q(ns, "p")))
    if paragraph > len(paragraphs):
        return "run_paragraph_missing"
    target = paragraphs[paragraph - 1]
    if fingerprint is not None and _cell_structure_fingerprint(target) != fingerprint:
        return "run_structure_fingerprint_mismatch"
    parents = {child: parent for parent in root.iter() for child in parent}
    if not _is_admitted_paragraph(target, container, parents, ns):
        return "run_context_unsupported"
    runs = [child for child in target if child.tag == Q(ns, "r")]
    if run is not None and run > len(runs):
        return "run_target_missing"
    if action == "delete":
        doomed = runs[run - 1]  # type: ignore[operator]
        deletable = {Q(ns, name) for name in _DELETABLE_RUN_CHILDREN}
        if any(child.tag not in deletable for child in doomed):
            return "run_content_unsupported"
        target.remove(doomed)
    else:
        assert text is not None
        anchor = runs[run - 1] if run is not None else (runs[-1] if runs else None)
        new_run = _new_run(text, anchor, ns)
        if anchor is None:
            target.append(new_run)
        else:
            index = list(target).index(anchor)
            target.insert(index + 1 if position == "after" else index, new_run)
    try:
        start, end, _after, encoding = _paragraph_xml_span(part_bytes, ns, paragraph)
        fragment = _paragraph_xml_fragment(target, encoding)
    except (expat.ExpatError, ValueError, LookupError):
        return "run_staged_verification_failed"
    return part_bytes[:start] + fragment + part_bytes[end:], root


_SAVE_CODES = {
    "document_destination_exists": "run_destination_occupied",
    "document_publication_unavailable": "run_publication_unavailable",
}


def author_run(request: RunRequest) -> RunReport:
    """Insert or delete one run of ``request.source`` and publish ``request.destination``.

    The source must still hash to ``request.source_sha256`` and the paragraph must
    still match its structure fingerprint; an existing destination is never replaced.
    """
    from ..document import Document

    def refusal(code: str) -> RunReport:
        return RunReport(
            _REFUSED,
            code,
            request.source_sha256,
            None,
            request.target_dict(),
            request.action,
        )

    if not isinstance(request, RunRequest):
        raise TypeError("run_request_invalid")
    if not request.source.is_file():
        return refusal("run_source_missing")
    if request.destination.exists() or request.destination.is_symlink():
        return refusal("run_destination_occupied")
    try:
        document = Document(request.source)
    except (PackageError, OSError):
        return refusal("run_package_invalid")
    with document:
        if document.source_sha256 != request.source_sha256:
            return refusal("run_source_hash_mismatch")
        edited = document._run_edit(
            request.paragraph,
            request.action,
            request.run,
            request.position,
            request.text,
            request.part,
            request.structure_fingerprint_sha256,
        )
        if not edited.passed:
            return refusal(edited.code or "run_staged_verification_failed")
        saved = document.save(request.destination)
    if not saved.passed:
        return refusal(_SAVE_CODES.get(saved.code or "", "run_staged_verification_failed"))
    return RunReport(
        _PUBLISHED,
        None,
        request.source_sha256,
        saved.destination_sha256,
        request.target_dict(),
        request.action,
    )
