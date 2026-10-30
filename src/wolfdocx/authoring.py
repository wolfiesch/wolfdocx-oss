"""Conservative, source-bound authoring for Word cells and split-run text."""

from __future__ import annotations

import copy
import hashlib
import json
import os
import tempfile
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from xml.parsers import expat

from .content import (
    ContentScope,
    _evaluate_target_changes,
    _matching_elements,
)
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    _inspect_docx,
    _require_verify_mode,
    _verify_changed_parts,
    inspect_docx,
)
from .package import PackageError, PackageGraph

_SHA256_LENGTH = 64
_MAX_TEXT_CHARACTERS = 32_768
_REFUSED = "refused"
_PUBLISHED = "published"
# A session edit held in memory until ``Document.save`` publishes it.
_APPLIED = "applied"
_XML_SPACE_ATTRIBUTE = "{http://www.w3.org/XML/1998/namespace}space"
_MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_STRICT_MATH_NAMESPACE = "http://purl.oclc.org/ooxml/officeDocument/math"
_MATHML_NAMESPACE = "http://www.w3.org/1998/Math/MathML"
_MATH_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "structure_fingerprint_sha256",
        "target",
        "omml_xml",
        "mathml_xml",
    }
)


_CONTEXT_FEATURES = frozenset(
    {
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
        "customXml",
        "del",
        "fldChar",
        "fldSimple",
        "hyperlink",
        "ins",
        "moveFrom",
        "moveTo",
        "sdt",
        "sdtContent",
    }
)


class CellRequestLimitError(ValueError):
    """A content-safe refusal for one bounded caller-supplied text field."""

    def __init__(self, field: str, requested_characters: int) -> None:
        self.field = field
        self.requested_characters = requested_characters
        super().__init__(
            f"{field} has {requested_characters} characters; maximum is "
            f"{_MAX_TEXT_CHARACTERS}. Shorten it or choose a smaller span."
        )

    @property
    def code(self) -> str:
        return f"cell_{self.field}_character_limit"

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "cell",
            "passed": False,
            "status": _REFUSED,
            "code": self.code,
            "limit": {
                "field": self.field,
                "maximum_characters": _MAX_TEXT_CHARACTERS,
                "requested_characters": self.requested_characters,
                "recovery": "Shorten the value or choose a smaller span.",
            },
        }


class SpanRequestLimitError(CellRequestLimitError):
    """A content-safe refusal for one bounded span request text field."""

    @property
    def code(self) -> str:
        return f"span_{self.field}_character_limit"

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "span",
            "passed": False,
            "status": _REFUSED,
            "code": self.code,
            "limit": {
                "field": self.field,
                "maximum_characters": _MAX_TEXT_CHARACTERS,
                "requested_characters": self.requested_characters,
                "recovery": "Shorten the value or choose a smaller span.",
            },
        }


@dataclass(frozen=True)
class CellTextRequest:
    """One source-bound replacement for a simple table-cell text node."""

    source: Path
    destination: Path
    source_sha256: str
    expected_old_text: str
    structure_fingerprint_sha256: str
    table: int
    row: int
    cell: int
    text: str

    @property
    def target_scope(self) -> ContentScope:
        return ContentScope(kind="table_cell", table=self.table, row=self.row, cell=self.cell)


@dataclass(frozen=True)
class CellTextAuthoringReport:
    """Content-safe public result for a cell text authoring request."""

    status: str
    code: str | None
    source_sha256: str
    target: dict[str, int]
    structure_fingerprint_sha256: str | None = None
    destination_sha256: str | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status in (_PUBLISHED, _APPLIED)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "cell",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "target": self.target,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "destination_sha256": self.destination_sha256,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _refusal(request: CellTextRequest, code: str) -> CellTextAuthoringReport:
    return CellTextAuthoringReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        target=request.target_scope.to_dict(),
    )


def _require_sha256(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != _SHA256_LENGTH:
        raise ValueError(f"{label}_invalid")
    try:
        int(value, 16)
    except ValueError as error:
        raise ValueError(f"{label}_invalid") from error
    return value.lower()


def _require_positive_coordinate(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError(f"{label}_invalid")
    return value


def parse_cell_text_request(value: Any) -> CellTextRequest:
    """Parse the intentionally small public request schema without coercion."""

    if not isinstance(value, Mapping):
        raise TypeError("cell_request_invalid")
    expected_keys = {
        "source",
        "destination",
        "source_sha256",
        "expected_old_text",
        "structure_fingerprint_sha256",
        "target",
        "text",
    }
    if set(value) != expected_keys:
        raise ValueError("cell_request_fields_invalid")
    source = value["source"]
    destination = value["destination"]
    expected_old_text = value["expected_old_text"]
    text = value["text"]
    target = value["target"]
    if not isinstance(source, str) or not source:
        raise ValueError("cell_source_invalid")
    if not isinstance(destination, str) or not destination:
        raise ValueError("cell_destination_invalid")
    if not isinstance(expected_old_text, str):
        raise TypeError("cell_expected_old_text_invalid")
    if len(expected_old_text) > _MAX_TEXT_CHARACTERS:
        raise CellRequestLimitError("expected_old_text", len(expected_old_text))
    if not isinstance(text, str):
        raise TypeError("cell_text_invalid")
    if len(text) > _MAX_TEXT_CHARACTERS:
        raise CellRequestLimitError("text", len(text))
    if not isinstance(target, Mapping) or set(target) != {"table", "row", "cell"}:
        raise ValueError("cell_target_invalid")
    return CellTextRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=_require_sha256(value["source_sha256"], "cell_source_sha256"),
        expected_old_text=expected_old_text,
        structure_fingerprint_sha256=_require_sha256(
            value["structure_fingerprint_sha256"], "cell_structure_fingerprint_sha256"
        ),
        table=_require_positive_coordinate(target["table"], "cell_target_table"),
        row=_require_positive_coordinate(target["row"], "cell_target_row"),
        cell=_require_positive_coordinate(target["cell"], "cell_target_cell"),
        text=text,
    )


def load_cell_text_request(path: str | Path) -> CellTextRequest:
    """Load one UTF-8 JSON request file with no path data in its result surface."""

    with Path(path).open(encoding="utf-8") as stream:
        return parse_cell_text_request(json.load(stream))


def _word_namespace(root: ET.Element) -> str | None:
    if not isinstance(root.tag, str) or not root.tag.startswith("{"):
        return None
    namespace = root.tag[1:].split("}", 1)[0]
    return namespace if namespace in SUPPORTED_WORD_NAMESPACES else None


def _selected_cell(
    root: ET.Element, request: CellTextRequest, word_namespace: str
) -> ET.Element | None:
    def word(local):
        return Q(word_namespace, local)
    tables = list(root.iter(word("tbl")))
    if request.table > len(tables):
        return None
    rows = tables[request.table - 1].findall(word("tr"))
    if request.row > len(rows):
        return None
    cells = rows[request.row - 1].findall(word("tc"))
    if request.cell > len(cells):
        return None
    return cells[request.cell - 1]


def _cell_structure_fingerprint(cell: ET.Element) -> str:
    """Hash the selected cell's XML structure while deliberately masking text."""

    def project(element: ET.Element) -> list[Any]:
        return [
            element.tag,
            sorted(element.attrib.items()),
            None if element.tag.rsplit("}", 1)[-1] == "t" else element.text,
            element.tail,
            [project(child) for child in element],
        ]

    canonical = json.dumps(project(cell), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def _has_unsupported_feature(cell: ET.Element, word_namespace: str) -> bool:
    unsupported = {
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
        "del",
        "delText",
        "fldChar",
        "fldSimple",
        "hyperlink",
        "ins",
        "instrText",
        "moveFrom",
        "moveFromRangeEnd",
        "moveFromRangeStart",
        "moveTo",
        "moveToRangeEnd",
        "moveToRangeStart",
    }
    tags = {Q(word_namespace, local) for local in unsupported}
    return any(child.tag in tags for child in cell.iter())


class _ContextState:
    """Ancestor and sibling-range admission state, consumed in XML start order."""

    def __init__(self, word_namespace: str) -> None:
        def word(local):
            return Q(word_namespace, local)
        self.context_tags = {word(feature) for feature in _CONTEXT_FEATURES}
        self.range_starts = {
            word("commentRangeStart"): "comment",
            word("moveFromRangeStart"): "moveFrom",
            word("moveToRangeStart"): "moveTo",
        }
        self.range_ends = {
            word("commentRangeEnd"): "comment",
            word("moveFromRangeEnd"): "moveFrom",
            word("moveToRangeEnd"): "moveTo",
        }
        self.state_tags = set(self.range_starts) | set(self.range_ends) | {word("fldChar")}
        self.active_ranges: set[tuple[str, str]] = set()
        self.active_fields = 0
        self.id_attribute = word("id")
        self.field_char = word("fldChar")
        self.field_char_type = word("fldCharType")

    def enter(self, element: ET.Element, ancestor_context: bool) -> tuple[bool, bool]:
        unsupported = bool(self.active_ranges) or self.active_fields > 0 or ancestor_context
        child_context = ancestor_context or element.tag in self.context_tags
        if element.tag not in self.state_tags:
            return unsupported, child_context
        range_kind = self.range_starts.get(element.tag)
        if range_kind is not None:
            self.active_ranges.add((range_kind, element.attrib.get(self.id_attribute, "")))
        elif element.tag == self.field_char:
            field_type = element.attrib.get(self.field_char_type)
            if field_type == "begin":
                self.active_fields += 1
            elif field_type == "end" and self.active_fields:
                self.active_fields -= 1
        else:
            range_kind = self.range_ends.get(element.tag)
            if range_kind is not None:
                self.active_ranges.discard((range_kind, element.attrib.get(self.id_attribute, "")))
        return unsupported, child_context


class _ContextReached(Exception):
    """Ends a context walk at the one element whose context was asked for."""


def _unsupported_context_by_element(
    root: ET.Element, word_namespace: str, *, stop_at: ET.Element | None = None
) -> dict[int, bool]:
    """Index the existing ancestor/range refusal semantics in one document walk.

    An element's context depends only on what precedes it in document order, so with
    ``stop_at`` the walk ends once that element's context is recorded.
    """

    state = _ContextState(word_namespace)
    unsupported: dict[int, bool] = {}

    def visit(element: ET.Element, ancestor_context: bool) -> None:
        context, child_context = state.enter(element, ancestor_context)
        unsupported[id(element)] = context
        if element is stop_at:
            raise _ContextReached
        for child in element:
            visit(child, child_context)

    try:
        visit(root, False)
    except _ContextReached:
        pass
    return unsupported


def _has_unsupported_ancestor(root: ET.Element, selected: ET.Element, word_namespace: str) -> bool:
    """Refuse wrapper contexts and active sibling ranges around a target."""

    return _unsupported_context_by_element(root, word_namespace, stop_at=selected).get(
        id(selected), True
    )


def _start_tag_end(document_xml: bytes, start: int) -> int | None:
    quote: int | None = None
    for index in range(start, len(document_xml)):
        current = document_xml[index]
        if quote is not None:
            if current == quote:
                quote = None
        elif current in {ord('"'), ord("'")}:
            quote = current
        elif current == ord(">"):
            return index + 1
    return None


def _escape_xml_text(text: str) -> bytes:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").encode("utf-8")


class _TargetResolved(Exception):
    """Stops the patch scan once the target text node's byte range is decided."""


def _patch_xml_text(
    document_xml: bytes, *, word_namespace: str, occurrence: int, text: str
) -> bytes | None:
    """Replace exactly one parsed Word text-node content range without reserializing XML.

    ``document_xml`` is already known to parse (the caller's tree came from exactly
    these bytes), so the scan stops at the target's end tag, or at its start tag when
    the target cannot be patched; the patched bytes are re-parsed by the caller.
    """

    parser = expat.ParserCreate(namespace_separator="}")
    element_stack: list[str] = []
    # Local names of the open elements, so each start tag checks the p/r/t suffix
    # without re-splitting its ancestors' names.
    local_stack: list[str] = []
    found = 0
    target_depth: int | None = None
    content_start: int | None = None
    content_end: int | None = None
    patchable = True
    target_active = False
    target_name = f"{word_namespace}" + "}t"

    def start_element(name: str, attributes: dict[str, str]) -> None:
        nonlocal content_start, found, patchable, target_active, target_depth
        del attributes
        element_stack.append(name)
        local_stack.append(name.rsplit("}", 1)[-1])
        if (
            local_stack[-1] == "t"
            and len(local_stack) >= 3
            and local_stack[-2] == "r"
            and local_stack[-3] == "p"
        ):
            found += 1
            if found == occurrence:
                start = parser.CurrentByteIndex
                content_start = _start_tag_end(document_xml, start)
                target_depth = len(element_stack)
                target_active = True
                if (
                    name != target_name
                    or content_start is None
                    or document_xml[start : content_start - 1].rstrip().endswith(b"/")
                ):
                    patchable = False
                    raise _TargetResolved

    def end_element(name: str) -> None:
        nonlocal content_end, target_active
        if target_active and target_depth == len(element_stack) and name == target_name:
            content_end = parser.CurrentByteIndex
            target_active = False
            raise _TargetResolved
        element_stack.pop()
        local_stack.pop()

    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    try:
        parser.Parse(document_xml, True)
    except _TargetResolved:
        pass
    except expat.ExpatError:
        return None
    if not patchable or content_start is None or content_end is None or content_start > content_end:
        return None
    return document_xml[:content_start] + _escape_xml_text(text) + document_xml[content_end:]


def _simple_text_node(cell: ET.Element, word_namespace: str) -> tuple[ET.Element, str] | None:
    """Return the only supported ``w:t`` and its current cell text.

    Policy: the selected cell must be unmerged and contain exactly one direct
    paragraph, exactly one direct run, and exactly one direct text node. Its
    ``w:pPr`` and ``w:rPr`` children are retained verbatim. Whitespace policy is
    checked against the text node's effective ``xml:space`` before patching.
    """

    def word(local):
        return Q(word_namespace, local)
    if _has_unsupported_feature(cell, word_namespace):
        return None
    tc_pr = cell.find(word("tcPr"))
    if tc_pr is not None and (
        tc_pr.find(word("vMerge")) is not None
        or tc_pr.find(word("hMerge")) is not None
        or tc_pr.find(word("gridSpan")) is not None
    ):
        return None
    if any(child.tag not in {word("tcPr"), word("p")} for child in cell):
        return None
    paragraphs = cell.findall(word("p"))
    if len(paragraphs) != 1:
        return None
    paragraph = paragraphs[0]
    if any(child.tag not in {word("pPr"), word("r")} for child in paragraph):
        return None
    runs = paragraph.findall(word("r"))
    if len(runs) != 1:
        return None
    run = runs[0]
    if any(child.tag not in {word("rPr"), word("t")} for child in run):
        return None
    text_nodes = run.findall(word("t"))
    if len(text_nodes) != 1:
        return None
    text = text_nodes[0].text or ""
    return text_nodes[0], text


def _xml_text_occurrence(root: ET.Element, target: ET.Element) -> int:
    """1-based pre-order rank of ``target`` among elements on a ``p/r/t`` local path."""
    count = 0

    def visit(element: ET.Element, grandparent: str | None, parent: str | None) -> bool:
        nonlocal count
        name = element.tag.rsplit("}", 1)[-1]
        if name == "t" and parent == "r" and grandparent == "p":
            count += 1
            if element is target:
                return True
        for child in element:
            if visit(child, parent, name):
                return True
        return False

    if visit(root, None, None):
        return count
    raise ValueError("cell_target_text_missing")


def _verification_xml_text_occurrence(root: ET.Element, target: ET.Element) -> int:
    """Return the target occurrence under the generic verifier's path grammar."""

    scope = ContentScope(
        kind="xml_text",
        part="word/document.xml",
        element_path=("p", "r", "t"),
        occurrence=1,
    )
    for occurrence, element in enumerate(_matching_elements(root, scope), start=1):
        if element is target:
            return occurrence
    raise ValueError("cell_target_text_missing")


def _non_target_part_bytes_preserved(source: Path, candidate: Path) -> bool:
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(candidate) as rewritten:
        baseline_names = set(baseline.namelist())
        candidate_names = set(rewritten.namelist())
        if baseline_names != candidate_names:
            return False
        return all(
            baseline.read(name) == rewritten.read(name)
            for name in baseline_names
            if name != "word/document.xml"
        )


def _write_staged_package(source: Path, staged: Path, document_xml: bytes) -> None:
    with (
        zipfile.ZipFile(source) as baseline,
        zipfile.ZipFile(staged, "w", compression=zipfile.ZIP_DEFLATED) as rewritten,
    ):
        for info in baseline.infolist():
            data = (
                document_xml
                if info.filename == "word/document.xml"
                else baseline.read(info.filename)
            )
            rewritten.writestr(copy.copy(info), data)


def _cell_text_from_inspection(report: Any, scope: ContentScope) -> str | None:
    table = scope.table
    row = scope.row
    cell = scope.cell
    assert table is not None and row is not None and cell is not None
    tables = report.fingerprint.tables
    if table > len(tables):
        return None
    rows = tables[table - 1]["rows"]
    if row > len(rows):
        return None
    cells = rows[row - 1]["cells"]
    if cell > len(cells):
        return None
    return cells[cell - 1]["text"]


@dataclass(frozen=True)
class _CellPlan:
    """One validated simple-cell replacement, patched into the document part bytes."""

    document_xml: bytes
    word_namespace: str
    target_change_scope: ContentScope


def _plan_cell_text(
    request: CellTextRequest, root: ET.Element, document_xml: bytes
) -> _CellPlan | str:
    """Validate one cell request against ``root`` and patch ``document_xml`` in place.

    ``root`` is only read; it must be parsed from exactly ``document_xml``. Returns a
    refusal code or the patched part bytes with the scope the edit may change.
    """
    word_namespace = _word_namespace(root)
    if word_namespace is None:
        return "source_word_namespace_unsupported"
    selected_cell = _selected_cell(root, request, word_namespace)
    if selected_cell is None:
        return "cell_coordinate_missing"
    if _cell_structure_fingerprint(selected_cell) != request.structure_fingerprint_sha256:
        return "cell_structure_mismatch"
    if _has_unsupported_ancestor(root, selected_cell, word_namespace):
        return "cell_context_unsupported"
    simple_text = _simple_text_node(selected_cell, word_namespace)
    if simple_text is None:
        return "cell_structure_unsupported"
    text_node, old_text = simple_text
    source_has_padding = old_text[:1].isspace() or old_text[-1:].isspace()
    if source_has_padding and not _text_replacements_preserve_whitespace(
        root, {id(text_node): old_text}
    ):
        return "cell_structure_unsupported"
    if old_text != request.expected_old_text:
        return "expected_old_text_mismatch"
    if request.text == old_text:
        return "replacement_unchanged"
    if (
        not source_has_padding
        and (request.text[:1].isspace() or request.text[-1:].isspace())
        and not _text_replacements_preserve_whitespace(root, {id(text_node): request.text})
    ):
        return "replacement_whitespace_unsupported"

    patch_occurrence = _xml_text_occurrence(root, text_node)
    verification_occurrence = _verification_xml_text_occurrence(root, text_node)
    target_change_scope = ContentScope(
        kind="xml_text",
        part="word/document.xml",
        element_path=("p", "r", "t"),
        occurrence=verification_occurrence,
    )
    patched_document_xml = _patch_xml_text(
        document_xml,
        word_namespace=word_namespace,
        occurrence=patch_occurrence,
        text=request.text,
    )
    if patched_document_xml is None:
        return "cell_text_patch_unsupported"
    return _CellPlan(patched_document_xml, word_namespace, target_change_scope)


def _cell_text_landed(
    candidate_root: ET.Element, request: CellTextRequest, word_namespace: str
) -> bool:
    """True when a re-parsed document holds the requested text in an unchanged cell."""
    candidate_cell = _selected_cell(candidate_root, request, word_namespace)
    candidate_text = (
        _simple_text_node(candidate_cell, word_namespace) if candidate_cell is not None else None
    )
    return (
        candidate_cell is not None
        and candidate_text is not None
        and candidate_text[1] == request.text
        and _cell_structure_fingerprint(candidate_cell) == request.structure_fingerprint_sha256
    )


def author_cell_text(request: CellTextRequest, *, verify: str = "full") -> CellTextAuthoringReport:
    """Publish one verified simple-cell replacement only after all checks pass.

    ``verify="full"`` (the default) re-inspects the whole staged package, including
    its document fingerprint, before publishing. ``verify="changed-parts"`` keeps every
    request check and the source-hash re-check, re-parses ``word/document.xml`` under
    the package limits, proves every other part (relationships included) is
    byte-identical to the validated source and re-checks the target cell, but skips
    re-inspecting and fingerprinting the unchanged parts of the staged package.
    """

    _require_verify_mode(verify)
    source = request.source.resolve()
    destination = request.destination.resolve()
    if source == destination:
        return _refusal(request, "source_destination_same")
    if not source.is_file():
        return _refusal(request, "source_missing")
    if not destination.parent.is_dir():
        return _refusal(request, "destination_parent_missing")
    if _sha256_file(source) != request.source_sha256:
        return _refusal(request, "source_hash_mismatch")

    baseline = _inspect_docx(source)
    if not baseline.valid or baseline.file_sha256 != request.source_sha256:
        return _refusal(request, "source_inspection_invalid")
    try:
        source_graph = PackageGraph.from_file(source)
        if source_graph.file_sha256 != request.source_sha256:
            return _refusal(request, "source_hash_changed")
        document_xml = source_graph.read_part_bytes("word/document.xml")
    except PackageError:
        return _refusal(request, "source_document_missing")
    # The baseline tree was parsed from these exact bytes (same package digest); it is
    # shared, so this function only reads it.
    root = baseline.parsed_xml.get("word/document.xml")
    if root is None:
        return _refusal(request, "source_document_invalid")
    plan = _plan_cell_text(request, root, document_xml)
    if isinstance(plan, str):
        return _refusal(request, plan)
    word_namespace = plan.word_namespace
    patched_document_xml = plan.document_xml
    target_change_scope = plan.target_change_scope
    staged_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as staged:
            staged_path = Path(staged.name)
        _write_staged_package(source, staged_path, patched_document_xml)
        if verify == "full":
            candidate = _inspect_docx(staged_path, reuse=baseline)
            target_content_ok = (
                _cell_text_from_inspection(candidate.report, request.target_scope) == request.text
            )
            candidate_valid = candidate.valid
            destination_sha256 = candidate.file_sha256
            target_change_report = _evaluate_target_changes(
                [target_change_scope],
                baseline.part_sha256,
                candidate.part_sha256,
                load_baseline_root=baseline.mutable_root,
                load_candidate_root=candidate.mutable_root,
            )
            non_target_parts_preserved = _non_target_part_bytes_preserved(source, staged_path)
        else:
            verified = _verify_changed_parts(
                baseline, staged_path, {"word/document.xml": patched_document_xml}
            )
            if verified is None:
                return _refusal(request, "staged_verification_failed")
            destination_sha256, candidate_roots = verified
            candidate_valid = True
            target_content_ok = _cell_text_landed(
                candidate_roots["word/document.xml"], request, word_namespace
            )
            target_change_report = _evaluate_target_changes(
                [target_change_scope],
                baseline.part_sha256,
                {
                    **baseline.part_sha256,
                    "word/document.xml": hashlib.sha256(patched_document_xml).hexdigest(),
                },
                load_baseline_root=baseline.mutable_root,
                load_candidate_root=candidate_roots.get,
            )
            non_target_parts_preserved = True
        source_still_bound = _sha256_file(source) == request.source_sha256
        if not (
            candidate_valid
            and target_change_report.passed
            and target_content_ok
            and non_target_parts_preserved
            and source_still_bound
        ):
            return _refusal(
                request,
                "staged_verification_failed" if source_still_bound else "source_hash_changed",
            )
        os.replace(staged_path, destination)
        staged_path = None
        return CellTextAuthoringReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            target=request.target_scope.to_dict(),
            structure_fingerprint_sha256=request.structure_fingerprint_sha256,
            destination_sha256=destination_sha256,
            target_change_report=target_change_report.to_dict(),
            non_target_parts_preserved=True,
        )
    finally:
        if staged_path is not None:
            staged_path.unlink(missing_ok=True)


def cell_structure_fingerprint(
    source: str | Path, *, table: int, row: int, cell: int
) -> str | None:
    """Return the bounded selected-cell structural fingerprint for request construction."""

    document = Path(source)
    inspection = _inspect_docx(document)
    if not inspection.valid:
        return None
    root = inspection.parsed_xml.get("word/document.xml")
    if root is None:
        return None
    namespace = _word_namespace(root)
    if namespace is None:
        return None
    request = CellTextRequest(
        source=document,
        destination=document.with_suffix(".unused.docx"),
        source_sha256=inspection.file_sha256,
        expected_old_text="",
        structure_fingerprint_sha256="0" * _SHA256_LENGTH,
        table=table,
        row=row,
        cell=cell,
        text="",
    )
    selected = _selected_cell(root, request, namespace)
    if selected is None:
        return None
    return _cell_structure_fingerprint(selected)


@dataclass(frozen=True)
class SpanTextRequest:
    source: Path
    destination: Path
    source_sha256: str
    expected_old_text: str
    structure_fingerprint_sha256: str
    paragraph: int
    text: str


def parse_span_text_request(value: Any) -> SpanTextRequest:
    if not isinstance(value, Mapping):
        raise TypeError("span_request_invalid")
    if set(value) != {
        "source",
        "destination",
        "source_sha256",
        "expected_old_text",
        "structure_fingerprint_sha256",
        "paragraph",
        "text",
    }:
        raise ValueError("span_request_fields_invalid")
    expected_old_text = value["expected_old_text"]
    text = value["text"]
    if not isinstance(expected_old_text, str) or not isinstance(text, str):
        raise TypeError("span_text_invalid")
    if not expected_old_text:
        raise ValueError("span_expected_old_text_empty")
    if len(expected_old_text) > _MAX_TEXT_CHARACTERS:
        raise SpanRequestLimitError("expected_old_text", len(expected_old_text))
    if len(text) > _MAX_TEXT_CHARACTERS:
        raise SpanRequestLimitError("text", len(text))
    source, destination = value["source"], value["destination"]
    if (
        not isinstance(source, str)
        or not source
        or not isinstance(destination, str)
        or not destination
    ):
        raise ValueError("span_path_invalid")
    return SpanTextRequest(
        Path(source),
        Path(destination),
        _require_sha256(value["source_sha256"], "span_source_sha256"),
        expected_old_text,
        _require_sha256(value["structure_fingerprint_sha256"], "span_structure_fingerprint_sha256"),
        _require_positive_coordinate(value["paragraph"], "span_paragraph"),
        text,
    )


def load_span_text_request(path: str | Path) -> SpanTextRequest:
    with Path(path).open(encoding="utf-8") as stream:
        return parse_span_text_request(json.load(stream))


def _paragraph_structure_fingerprint(paragraph: ET.Element) -> str:
    return _cell_structure_fingerprint(paragraph)


def paragraph_structure_fingerprint(source: str | Path, *, paragraph: int) -> str | None:
    inspection = inspect_docx(source)
    if not inspection.valid:
        return None
    try:
        root = ET.fromstring(PackageGraph.from_file(source).read_part_bytes("word/document.xml"))
    except (ET.ParseError, PackageError):
        return None
    namespace = _word_namespace(root)
    if namespace is None:
        return None
    paragraphs = list(root.iter(Q(namespace, "p")))
    if paragraph <= 0 or paragraph > len(paragraphs):
        return None
    return _paragraph_structure_fingerprint(paragraphs[paragraph - 1])


@dataclass(frozen=True)
class SpanTextAuthoringReport:
    """Content-safe public result for a paragraph-fragment authoring request."""

    status: str
    code: str | None
    source_sha256: str
    target: dict[str, int]
    structure_fingerprint_sha256: str | None = None
    destination_sha256: str | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_document_bytes_preserved: bool | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "span",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "target": self.target,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "destination_sha256": self.destination_sha256,
            "target_change_report": self.target_change_report,
            "non_target_document_bytes_preserved": (self.non_target_document_bytes_preserved),
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


@dataclass(frozen=True)
class _SpanTextNode:
    node: ET.Element
    text: str
    start: int
    end: int


def _span_refusal(request: SpanTextRequest, code: str) -> SpanTextAuthoringReport:
    return SpanTextAuthoringReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        target={"paragraph": request.paragraph},
    )


def _selected_paragraph(
    root: ET.Element, request: SpanTextRequest, word_namespace: str
) -> ET.Element | None:
    paragraphs = list(root.iter(Q(word_namespace, "p")))
    if request.paragraph > len(paragraphs):
        return None
    return paragraphs[request.paragraph - 1]


def _span_text_nodes(paragraph: ET.Element, word_namespace: str) -> list[_SpanTextNode] | None:
    """Return direct simple-run text nodes whose lexical runs can be patched.

    This deliberately accepts only direct ``w:r`` children with one ``w:t``.
    Any revision, field, comment, annotation, control, or unsupported inline
    content in the paragraph is refused rather than crossed or normalized.
    """

    def word(local):
        return Q(word_namespace, local)
    if _has_unsupported_feature(paragraph, word_namespace):
        return None
    if any(child.tag not in {word("pPr"), word("r")} for child in paragraph):
        return None
    nodes: list[_SpanTextNode] = []
    offset = 0
    for run in paragraph.findall(word("r")):
        if any(child.tag not in {word("rPr"), word("t")} for child in run):
            return None
        text_nodes = run.findall(word("t"))
        if len(text_nodes) != 1 or len(text_nodes[0]) != 0:
            return None
        text = text_nodes[0].text or ""
        nodes.append(
            _SpanTextNode(
                node=text_nodes[0],
                text=text,
                start=offset,
                end=offset + len(text),
            )
        )
        offset += len(text)
    return nodes or None


def _span_match(nodes: list[_SpanTextNode], *, start: int, end: int) -> list[_SpanTextNode] | None:
    selected = [node for node in nodes if node.start < end and start < node.end]
    return selected if len(selected) >= 2 else None


def _span_replacements(
    selected: list[_SpanTextNode], *, start: int, end: int, text: str
) -> dict[int, str]:
    replacements: dict[int, str] = {}
    for index, node in enumerate(selected):
        local_start = max(start, node.start) - node.start
        local_end = min(end, node.end) - node.start
        prefix = node.text[:local_start] if index == 0 else ""
        suffix = node.text[local_end:] if index == len(selected) - 1 else ""
        replacements[id(node.node)] = prefix + text + suffix if index == 0 else suffix
    return replacements


def _effective_xml_space_by_element(root: ET.Element) -> dict[int, str | None]:
    """Index effective ``xml:space`` inheritance without changing XML markup."""

    effective_spaces: dict[int, str | None] = {}

    def visit(element: ET.Element, inherited_space: str | None) -> None:
        effective_space = element.attrib.get(_XML_SPACE_ATTRIBUTE, inherited_space)
        effective_spaces[id(element)] = effective_space
        for child in element:
            visit(child, effective_space)

    visit(root, None)
    return effective_spaces


def _text_replacements_preserve_whitespace(
    root: ET.Element,
    replacements: Mapping[int, str],
    *,
    effective_spaces: Mapping[int, str | None] | None = None,
) -> bool:
    """Require existing effective ``xml:space="preserve"`` for edge whitespace."""

    whitespace_targets = {
        node_id
        for node_id, text in replacements.items()
        if text[:1].isspace() or text[-1:].isspace()
    }
    if not whitespace_targets:
        return True
    spaces = _effective_xml_space_by_element(root) if effective_spaces is None else effective_spaces
    return all(spaces.get(node_id) == "preserve" for node_id in whitespace_targets)


class _PatchTargetsClosed(Exception):
    """Ends a patch scan once every target text node has closed."""


def _patch_xml_texts(
    document_xml: bytes,
    *,
    word_namespace: str,
    replacements: Mapping[int, str],
) -> bytes | None:
    """Patch selected direct ``w:p/w:r/w:t`` contents without reserializing XML.

    Callers pass bytes they have already parsed whole, so the scan stops when the last
    target closes: the unchanged remainder cannot alter which ranges are patched.
    """

    parser = expat.ParserCreate(namespace_separator="}")
    local_stack: list[str] = []
    occurrence = 0
    active: tuple[int, int, int] | None = None
    ranges: list[tuple[int, int, str]] = []
    target_name = f"{word_namespace}" + "}t"
    run_text_path = ["p", "r", "t"]
    patchable = True
    found: set[int] = set()

    def start_element(name: str, attributes: dict[str, str]) -> None:
        nonlocal active, occurrence, patchable
        del attributes
        local_stack.append(name.rsplit("}", 1)[-1])
        if local_stack[-3:] != run_text_path:
            return
        occurrence += 1
        if occurrence not in replacements:
            return
        start = parser.CurrentByteIndex
        content_start = _start_tag_end(document_xml, start)
        if (
            name != target_name
            or content_start is None
            or document_xml[start : content_start - 1].rstrip().endswith(b"/")
        ):
            patchable = False
            return
        active = (occurrence, content_start, len(local_stack))

    def end_element(name: str) -> None:
        nonlocal active, patchable
        if active is not None and len(local_stack) == active[2]:
            target_occurrence, content_start, _ = active
            if name != target_name:
                patchable = False
            else:
                ranges.append(
                    (
                        content_start,
                        parser.CurrentByteIndex,
                        replacements[target_occurrence],
                    )
                )
                found.add(target_occurrence)
            active = None
        local_stack.pop()
        if len(found) == len(replacements):
            raise _PatchTargetsClosed

    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    try:
        parser.Parse(document_xml, True)
    except _PatchTargetsClosed:
        pass
    except expat.ExpatError:
        return None
    if (
        not patchable
        or active is not None
        or found != set(replacements)
        or len(ranges) != len(replacements)
    ):
        return None
    patched = document_xml
    for content_start, content_end, text in reversed(ranges):
        if content_start > content_end:
            return None
        patched = patched[:content_start] + _escape_xml_text(text) + patched[content_end:]
    return patched


def _span_text_occurrences(root: ET.Element, selected: list[_SpanTextNode]) -> dict[int, int]:
    targets = {id(node.node) for node in selected}
    occurrences: dict[int, int] = {}
    occurrence = 0

    def visit(element: ET.Element, parent: str | None, grandparent: str | None) -> None:
        nonlocal occurrence
        local = element.tag.rsplit("}", 1)[-1]
        if grandparent == "p" and parent == "r" and local == "t":
            occurrence += 1
            if id(element) in targets:
                occurrences[id(element)] = occurrence
        for child in element:
            visit(child, local, parent)

    visit(root, None, None)
    if set(occurrences) != targets:
        raise ValueError("span_target_text_missing")
    return occurrences


def _span_text_from_root(
    root: ET.Element, request: SpanTextRequest, word_namespace: str
) -> str | None:
    paragraph = _selected_paragraph(root, request, word_namespace)
    if paragraph is None:
        return None
    nodes = _span_text_nodes(paragraph, word_namespace)
    return "".join(node.text for node in nodes) if nodes is not None else None


def _span_target_change_report(
    request: SpanTextRequest,
    *,
    source_text: str,
    start: int,
    end: int,
    candidate_text: str | None,
    candidate_fingerprint: str | None,
) -> dict[str, Any]:
    expected_text = source_text[:start] + request.text + source_text[end:]
    target_found = candidate_text is not None
    target_changed = target_found and candidate_text != source_text
    structure_preserved = candidate_fingerprint == request.structure_fingerprint_sha256
    return {
        "scope": {"kind": "paragraph", "index": request.paragraph},
        "target_found": target_found,
        "target_changed": target_changed,
        "expected_text_sha256": hashlib.sha256(expected_text.encode("utf-8")).hexdigest(),
        "candidate_text_sha256": (
            hashlib.sha256(candidate_text.encode("utf-8")).hexdigest()
            if candidate_text is not None
            else None
        ),
        "paragraph_structure_preserved": structure_preserved,
        "passed": (
            target_found
            and target_changed
            and candidate_text == expected_text
            and structure_preserved
        ),
    }


def author_span_text(request: SpanTextRequest) -> SpanTextAuthoringReport:
    """Publish one source-bound replacement spanning at least two simple runs.

    The replacement is inserted into the first selected run, so that run's
    ``w:rPr`` supplies the replacement formatting. The selected final run
    retains only its suffix; all other runs and paragraph properties are left
    byte-for-byte intact.
    """

    source = request.source.resolve()
    destination = request.destination.resolve()
    if source == destination:
        return _span_refusal(request, "source_destination_same")
    if not source.is_file():
        return _span_refusal(request, "source_missing")
    if not destination.parent.is_dir():
        return _span_refusal(request, "destination_parent_missing")
    if _sha256_file(source) != request.source_sha256:
        return _span_refusal(request, "source_hash_mismatch")
    if not request.text:
        return _span_refusal(request, "replacement_empty_unsupported")
    if request.text == request.expected_old_text:
        return _span_refusal(request, "replacement_unchanged")
    if request.text[:1].isspace() or request.text[-1:].isspace():
        return _span_refusal(request, "replacement_whitespace_unsupported")

    baseline = inspect_docx(source)
    if not baseline.valid or baseline.file_sha256 != request.source_sha256:
        return _span_refusal(request, "source_inspection_invalid")
    try:
        source_graph = PackageGraph.from_file(source)
        if source_graph.file_sha256 != request.source_sha256:
            return _span_refusal(request, "source_hash_changed")
        document_xml = source_graph.read_part_bytes("word/document.xml")
        root = ET.fromstring(document_xml)
    except PackageError:
        return _span_refusal(request, "source_document_missing")
    except ET.ParseError:
        return _span_refusal(request, "source_document_invalid")
    word_namespace = _word_namespace(root)
    if word_namespace is None:
        return _span_refusal(request, "source_word_namespace_unsupported")
    paragraph = _selected_paragraph(root, request, word_namespace)
    if paragraph is None:
        return _span_refusal(request, "paragraph_coordinate_missing")
    if _paragraph_structure_fingerprint(paragraph) != request.structure_fingerprint_sha256:
        return _span_refusal(request, "paragraph_structure_mismatch")
    if _has_unsupported_ancestor(root, paragraph, word_namespace):
        return _span_refusal(request, "paragraph_context_unsupported")
    nodes = _span_text_nodes(paragraph, word_namespace)
    if nodes is None:
        return _span_refusal(request, "span_structure_unsupported")
    source_text = "".join(node.text for node in nodes)
    start = source_text.find(request.expected_old_text)
    if start < 0:
        return _span_refusal(request, "expected_old_text_mismatch")
    if source_text.find(request.expected_old_text, start + 1) >= 0:
        return _span_refusal(request, "span_occurrence_ambiguous")
    end = start + len(request.expected_old_text)
    selected = _span_match(nodes, start=start, end=end)
    if selected is None:
        return _span_refusal(request, "span_not_cross_run")
    replacements_by_node = _span_replacements(selected, start=start, end=end, text=request.text)
    if not _text_replacements_preserve_whitespace(root, replacements_by_node):
        return _span_refusal(request, "span_whitespace_unsupported")
    try:
        occurrences = _span_text_occurrences(root, selected)
    except ValueError:
        return _span_refusal(request, "span_target_text_missing")
    patched_document_xml = _patch_xml_texts(
        document_xml,
        word_namespace=word_namespace,
        replacements={occurrences[node_id]: text for node_id, text in replacements_by_node.items()},
    )
    if patched_document_xml is None:
        return _span_refusal(request, "span_text_patch_unsupported")

    staged_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as staged:
            staged_path = Path(staged.name)
        _write_staged_package(source, staged_path, patched_document_xml)
        candidate = inspect_docx(staged_path)
        candidate_document_xml: bytes | None = None
        candidate_text: str | None = None
        candidate_fingerprint: str | None = None
        if candidate.valid:
            try:
                candidate_document_xml = PackageGraph.from_file(staged_path).read_part_bytes(
                    "word/document.xml"
                )
                candidate_root = ET.fromstring(candidate_document_xml)
                candidate_namespace = _word_namespace(candidate_root)
                if candidate_namespace is not None:
                    candidate_text = _span_text_from_root(
                        candidate_root, request, candidate_namespace
                    )
                    candidate_paragraph = _selected_paragraph(
                        candidate_root, request, candidate_namespace
                    )
                    if candidate_paragraph is not None:
                        candidate_fingerprint = _paragraph_structure_fingerprint(
                            candidate_paragraph
                        )
            except (ET.ParseError, PackageError):
                pass
        target_change_report = _span_target_change_report(
            request,
            source_text="".join(node.text for node in nodes),
            start=start,
            end=end,
            candidate_text=candidate_text,
            candidate_fingerprint=candidate_fingerprint,
        )
        non_target_document_bytes_preserved = candidate_document_xml == patched_document_xml
        non_target_parts_preserved = _non_target_part_bytes_preserved(source, staged_path)
        source_still_bound = _sha256_file(source) == request.source_sha256
        if not (
            candidate.valid
            and target_change_report["passed"]
            and non_target_document_bytes_preserved
            and non_target_parts_preserved
            and source_still_bound
        ):
            return _span_refusal(
                request,
                "staged_verification_failed" if source_still_bound else "source_hash_changed",
            )
        os.replace(staged_path, destination)
        staged_path = None
        return SpanTextAuthoringReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            target={"paragraph": request.paragraph},
            structure_fingerprint_sha256=request.structure_fingerprint_sha256,
            destination_sha256=candidate.file_sha256,
            target_change_report=target_change_report,
            non_target_document_bytes_preserved=True,
            non_target_parts_preserved=True,
        )
    finally:
        if staged_path is not None:
            staged_path.unlink(missing_ok=True)


@dataclass(frozen=True)
class MathRequest:
    """Source-bound insert, edit or delete of math in a body paragraph.

    Edit/delete select the unique direct math container in that paragraph, or
    the immediately following body-level display equation. Ambiguous targets
    are refused. Delete accepts an optional input but never inserts it.
    """

    source: Path
    destination: Path
    source_sha256: str
    structure_fingerprint_sha256: str
    paragraph: int
    omml_xml: str | None
    mathml_xml: str | None
    latex: str | None = None
    operation: str = "insert"


@dataclass(frozen=True)
class MathAuthoringReport:
    """Content-safe result for one atomically published math mutation."""

    status: str
    code: str | None
    source_sha256: str
    target: dict[str, int]
    input_format: str | None = None
    inserted_container: str | None = None
    structure_fingerprint_sha256: str | None = None
    destination_sha256: str | None = None
    non_target_document_bytes_preserved: bool | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "math",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "target": self.target,
            "input_format": self.input_format,
            "inserted_container": self.inserted_container,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "destination_sha256": self.destination_sha256,
            "non_target_document_bytes_preserved": self.non_target_document_bytes_preserved,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _math_refusal(request: MathRequest, code: str) -> MathAuthoringReport:
    return MathAuthoringReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        target={"paragraph": request.paragraph},
    )


def parse_math_request(value: Any) -> MathRequest:
    """Parse a source-bound math mutation; insertion remains the default."""

    if not isinstance(value, Mapping):
        raise TypeError("math_request_invalid")
    fields = set(value)
    if fields - {"latex", "operation"} != _MATH_REQUEST_FIELDS:
        raise ValueError("math_request_fields_invalid")
    operation = value.get("operation", "insert")
    if operation not in ("insert", "edit", "delete"):
        raise ValueError("math_operation_invalid")
    source = value["source"]
    destination = value["destination"]
    target = value["target"]
    omml_xml = value["omml_xml"]
    mathml_xml = value["mathml_xml"]
    latex = value.get("latex")
    if not isinstance(source, str) or not source:
        raise ValueError("math_source_invalid")
    if not isinstance(destination, str) or not destination:
        raise ValueError("math_destination_invalid")
    if not isinstance(target, Mapping) or set(target) != {"paragraph"}:
        raise ValueError("math_target_invalid")
    if not isinstance(omml_xml, str | type(None)):
        raise TypeError("math_omml_xml_invalid")
    if not isinstance(mathml_xml, str | type(None)):
        raise TypeError("math_mathml_xml_invalid")
    if not isinstance(latex, str | type(None)):
        raise TypeError("math_latex_invalid")
    inputs = [item for item in (omml_xml, mathml_xml, latex) if item is not None]
    if len(inputs) != 1 and not (operation == "delete" and not inputs):
        raise ValueError("math_input_ambiguous")
    if inputs:
        payload = inputs[0]
        if not payload:
            raise ValueError("math_input_empty")
        if len(payload) > _MAX_TEXT_CHARACTERS:
            raise ValueError("math_input_character_limit")
    return MathRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=_require_sha256(value["source_sha256"], "math_source_sha256"),
        structure_fingerprint_sha256=_require_sha256(
            value["structure_fingerprint_sha256"], "math_structure_fingerprint_sha256"
        ),
        paragraph=_require_positive_coordinate(target["paragraph"], "math_target_paragraph"),
        omml_xml=omml_xml,
        mathml_xml=mathml_xml,
        latex=latex,
        operation=operation,
    )


def load_math_request(path: str | Path) -> MathRequest:
    """Load one UTF-8 JSON math mutation request."""

    with Path(path).open(encoding="utf-8") as stream:
        return parse_math_request(json.load(stream))


def _math_element(name: str, text: str | None = None) -> ET.Element:
    element = ET.Element(Q(_MATH_NAMESPACE, name))
    if text is not None:
        element.text = text
    return element


def _math_value(name: str, value: str) -> ET.Element:
    element = _math_element(name)
    element.set(Q(_MATH_NAMESPACE, "val"), value)
    return element


def _math_container(name: str, items: list[ET.Element]) -> ET.Element:
    container = _math_element(name)
    container.extend(items)
    return container


def _math_run(text: str, *, style: str | None = None, script: str | None = None) -> ET.Element:
    """One OMML run; ``style`` is ``m:sty`` (``p`` upright), ``nor`` marks normal text."""

    run = _math_element("r")
    if style is not None or script is not None:
        properties = _math_element("rPr")
        if style == "nor":
            properties.append(_math_element("nor"))
        else:
            if script is not None:
                properties.append(_math_value("scr", script))
            if style is not None:
                properties.append(_math_value("sty", style))
        run.append(properties)
    text_element = _math_element("t", text)
    if text != text.strip(" \t\r\n"):
        text_element.set(_XML_SPACE_ATTRIBUTE, "preserve")
    run.append(text_element)
    return run


def _math_fraction(numerator: list[ET.Element], denominator: list[ET.Element]) -> ET.Element:
    fraction = _math_element("f")
    fraction.extend((_math_container("num", numerator), _math_container("den", denominator)))
    return fraction


def _math_scripts(
    base: list[ET.Element], sub: list[ET.Element] | None, sup: list[ET.Element] | None
) -> ET.Element:
    if sub is not None and sup is not None:
        name, parts = "sSubSup", (("e", base), ("sub", sub), ("sup", sup))
    elif sub is not None:
        name, parts = "sSub", (("e", base), ("sub", sub))
    elif sup is not None:
        name, parts = "sSup", (("e", base), ("sup", sup))
    else:
        raise ValueError("math_script_missing")
    scripted = _math_element(name)
    scripted.extend(_math_container(role, items) for role, items in parts)
    return scripted


def _math_radical(radicand: list[ET.Element], degree: list[ET.Element] | None) -> ET.Element:
    radical = _math_element("rad")
    if degree is None:
        properties = _math_element("radPr")
        properties.append(_math_value("degHide", "1"))
        radical.append(properties)
    radical.extend((_math_container("deg", degree or []), _math_container("e", radicand)))
    return radical


def _math_delimiter(begin: str, end: str, items: list[ET.Element]) -> ET.Element:
    delimiter = _math_element("d")
    if begin != "(" or end != ")":
        properties = _math_element("dPr")
        properties.extend((_math_value("begChr", begin), _math_value("endChr", end)))
        delimiter.append(properties)
    delimiter.append(_math_container("e", items))
    return delimiter


def _math_matrix(rows: list[list[list[ET.Element]]]) -> ET.Element:
    if not rows or any(not row for row in rows):
        raise ValueError("math_matrix_empty")
    matrix = _math_element("m")
    for row in rows:
        matrix_row = _math_element("mr")
        matrix_row.extend(_math_container("e", cell) for cell in row)
        matrix.append(matrix_row)
    return matrix


def _math_accent(character: str, base: list[ET.Element]) -> ET.Element:
    accent = _math_element("acc")
    properties = _math_element("accPr")
    properties.append(_math_value("chr", character))
    accent.extend((properties, _math_container("e", base)))
    return accent


def _math_limit(name: str, base: list[ET.Element], limit: list[ET.Element]) -> ET.Element:
    element = _math_element(name)
    element.extend((_math_container("e", base), _math_container("lim", limit)))
    return element


#: Large operators that become ``m:nary``; integrals keep their limits at the side.
_NARY_OPERATORS = frozenset("∫∬∭∮∯∰∑∏∐⋃⋂⋁⋀⨁⨂⨀")
_INTEGRAL_OPERATORS = frozenset("∫∬∭∮∯∰")
#: Relations end the operand of a preceding n-ary operator.
_MATH_RELATIONS = frozenset("=<>≤≥≠≈≡∼≃≅∝→←↔⇒⇐⇔↦∈∉∋⊂⊆⊃⊇≪≫")
#: Spacing accents are written as the combining characters Word stores in ``m:chr``.
_ACCENT_CHARACTERS = {
    "^": "\u0302",
    "ˆ": "\u0302",
    "\u0302": "\u0302",
    "¯": "\u0305",
    "‾": "\u0305",
    "\u0305": "\u0305",
    "→": "\u20d7",
    "\u20d7": "\u20d7",
    "˙": "\u0307",
    "\u0307": "\u0307",
    "¨": "\u0308",
    "\u0308": "\u0308",
    "~": "\u0303",
    "˜": "\u0303",
    "\u0303": "\u0303",
}


@dataclass
class _Nary:
    """An n-ary operator whose operand is the rest of its row up to a relation."""

    character: str
    limits: str
    sub: list[ET.Element] | None
    sup: list[ET.Element] | None

    def build(self, operand: list[ET.Element]) -> ET.Element:
        nary = _math_element("nary")
        properties = _math_element("naryPr")
        properties.append(_math_value("chr", self.character))
        properties.append(_math_value("limLoc", self.limits))
        if self.sub is None:
            properties.append(_math_value("subHide", "1"))
        if self.sup is None:
            properties.append(_math_value("supHide", "1"))
        nary.extend(
            (
                properties,
                _math_container("sub", self.sub or []),
                _math_container("sup", self.sup or []),
                _math_container("e", operand),
            )
        )
        return nary


_MathItem = list[ET.Element] | _Nary


def _is_relation(item: _MathItem) -> bool:
    if isinstance(item, _Nary) or len(item) != 1 or item[0].tag != Q(_MATH_NAMESPACE, "r"):
        return False
    return "".join(item[0].itertext()) in _MATH_RELATIONS


def _math_row(items: list[_MathItem]) -> list[ET.Element]:
    """Flatten one row, giving each n-ary operator its operand up to a relation."""

    out: list[ET.Element] = []
    index = 0
    while index < len(items):
        item = items[index]
        if isinstance(item, _Nary):
            end = index + 1
            while end < len(items) and not _is_relation(items[end]):
                end += 1
            out.append(item.build(_math_row(items[index + 1 : end])))
            index = end
            continue
        out.extend(item)
        index += 1
    return out


def _mathml_children(element: ET.Element) -> list[ET.Element]:
    return [child for child in element if isinstance(child.tag, str)]


def _mathml_local(element: ET.Element) -> str:
    if not isinstance(element.tag, str) or not element.tag.startswith("{"):
        raise ValueError("mathml_namespace_invalid")
    namespace, local = element.tag[1:].split("}", 1)
    if namespace != _MATHML_NAMESPACE:
        raise ValueError("mathml_namespace_invalid")
    return local


def _mathml_operator(element: ET.Element) -> str | None:
    """The text of one ``mo`` element, or ``None`` for any other element."""

    if _mathml_local(element) != "mo":
        return None
    return "".join(element.itertext()).strip()


def _mathml_to_omml(element: ET.Element) -> list[ET.Element]:
    """Convert one MathML element to the OMML items it stands for."""

    return _math_row([_mathml_item(element)])


def _mathml_row(children: list[ET.Element]) -> list[ET.Element]:
    items: list[_MathItem] = []
    index = 0
    while index < len(children):
        if _mathml_operator(children[index]) == "\u2061":
            if not items or index + 1 == len(children) or isinstance(items[-1], _Nary):
                raise ValueError("mathml_function_invalid")
            function = _math_element("func")
            function.extend(
                (
                    _math_container("fName", items.pop()),
                    _math_container("e", _mathml_to_omml(children[index + 1])),
                )
            )
            items.append([function])
            index += 2
        else:
            items.append(_mathml_item(children[index]))
            index += 1
    return _math_row(items)


_MATHML_VARIANTS = {
    "normal": (None, "p"),
    "bold": (None, "b"),
    "italic": (None, "i"),
    "bold-italic": (None, "bi"),
    "double-struck": ("double-struck", "p"),
    "script": ("script", None),
    "fraktur": ("fraktur", None),
}


# The operator dictionary's fences: an ``mrow`` that opens with one and closes with
# another is a delimiter pair unless either operator says ``fence="false"``.
_MATHML_OPENING_FENCES = frozenset("([{|‖⟨⌊⌈")
_MATHML_CLOSING_FENCES = frozenset(")]}|‖⟩⌋⌉")
_MATHML_NAMED_SPACES = {
    "thinmathspace": "0.1667em",
    "mediummathspace": "0.2222em",
    "thickmathspace": "0.2778em",
}
# The Unicode space each positive ``mspace`` width rounds to, matching the LaTeX
# spacing commands: ``\,``, ``\:``/``\;``, ``\quad`` and ``\qquad``.
_MATHML_SPACE_WIDTHS = (
    (0.17, "\u2009"),
    (0.28, "\u2005"),
    (1.0, "\u2003"),
    (2.0, "\u2003\u2003"),
)


def _mathml_fence(element: ET.Element, text: str, fences: frozenset[str]) -> bool:
    default = "true" if text in fences else "false"
    return element.get("fence", default) == "true"


def _mathml_item(element: ET.Element) -> _MathItem:
    local = _mathml_local(element)
    children = _mathml_children(element)
    if local in {"math", "mrow", "mstyle"}:
        first = _mathml_operator(children[0]) if children else None
        last = _mathml_operator(children[-1]) if len(children) > 1 else None
        if (
            local == "mrow"
            and first is not None
            and last is not None
            and _mathml_fence(children[0], first, _MATHML_OPENING_FENCES)
            and _mathml_fence(children[-1], last, _MATHML_CLOSING_FENCES)
        ):
            return [_math_delimiter(first, last, _mathml_row(children[1:-1]))]
        return _mathml_row(children)
    if local == "semantics":
        if not children:
            raise ValueError("mathml_structure_invalid")
        return _mathml_item(children[0])
    if local == "mspace":
        width = element.get("width", "0em").strip()
        width = _MATHML_NAMED_SPACES.get(width, width)
        if not width.endswith("em"):
            raise ValueError("mathml_space_unsupported")
        try:
            ems = float(width[:-2])
        except ValueError:
            raise ValueError("mathml_space_unsupported") from None
        if ems <= 0:
            return []
        space = next((character for limit, character in _MATHML_SPACE_WIDTHS if ems <= limit), None)
        if space is None:
            raise ValueError("mathml_space_unsupported")
        return [_math_run(space)]
    if local in {"mi", "mn", "mo", "mtext"}:
        text = "".join(element.itertext())
        if local != "mtext":
            text = text.strip()
        if not text:
            raise ValueError("mathml_text_empty")
        if local == "mo" and text in _NARY_OPERATORS:
            return _Nary(text, "subSup" if text in _INTEGRAL_OPERATORS else "undOvr", None, None)
        if local == "mtext":
            return [_math_run(text, style="nor")]
        variant = element.get("mathvariant")
        if variant is not None:
            if variant not in _MATHML_VARIANTS:
                raise ValueError("mathml_variant_unsupported")
            script, style = _MATHML_VARIANTS[variant]
            return [_math_run(text, style=style, script=script)]
        if local == "mi" and len(text) > 1:
            return [_math_run(text, style="p")]
        return [_math_run(text)]
    if local == "mfrac":
        if len(children) != 2:
            raise ValueError("mathml_fraction_invalid")
        return [_math_fraction(_mathml_to_omml(children[0]), _mathml_to_omml(children[1]))]
    if local in {"msqrt", "mroot"}:
        if local == "msqrt":
            return [_math_radical(_mathml_row(children), None)]
        if len(children) != 2:
            raise ValueError("mathml_root_invalid")
        return [_math_radical(_mathml_to_omml(children[0]), _mathml_to_omml(children[1]))]
    scripts = {
        "msub": ("sub",),
        "msup": ("sup",),
        "msubsup": ("sub", "sup"),
        "munder": ("sub",),
        "mover": ("sup",),
        "munderover": ("sub", "sup"),
    }
    if local in scripts:
        roles = scripts[local]
        if len(children) != len(roles) + 1:
            raise ValueError("mathml_script_invalid")
        operator = _mathml_operator(children[0])
        parts = {
            role: _mathml_to_omml(child) for role, child in zip(roles, children[1:], strict=True)
        }
        if operator in _NARY_OPERATORS:
            limits = "undOvr" if local.startswith("mu") or local == "mover" else "subSup"
            return _Nary(operator, limits, parts.get("sub"), parts.get("sup"))
        base = _mathml_to_omml(children[0])
        if local in {"msub", "msup", "msubsup"}:
            return [_math_scripts(base, parts.get("sub"), parts.get("sup"))]
        if local in {"munder", "mover"}:
            character = _mathml_operator(children[1])
            if character in {"\u23df", "\u23de"}:
                group = _math_element("groupChr")
                properties = _math_element("groupChrPr")
                properties.extend(
                    (
                        _math_value("chr", character),
                        _math_value("pos", "bot" if local == "munder" else "top"),
                        _math_value("vertJc", "top" if local == "munder" else "bot"),
                    )
                )
                group.extend((properties, _math_container("e", base)))
                return [group]
        if local == "mover" and element.get("accent") == "true":
            accent = _mathml_operator(children[1]) or "".join(children[1].itertext()).strip()
            if accent not in _ACCENT_CHARACTERS:
                raise ValueError("mathml_accent_unsupported")
            return [_math_accent(_ACCENT_CHARACTERS[accent], base)]
        if local == "munder":
            return [_math_limit("limLow", base, parts["sub"])]
        if local == "mover":
            return [_math_limit("limUpp", base, parts["sup"])]
        return [_math_limit("limUpp", [_math_limit("limLow", base, parts["sub"])], parts["sup"])]
    if local == "mtable":
        rows: list[list[list[ET.Element]]] = []
        for row in children:
            if _mathml_local(row) != "mtr":
                raise ValueError("mathml_table_invalid")
            cells = _mathml_children(row)
            if any(_mathml_local(cell) != "mtd" for cell in cells):
                raise ValueError("mathml_table_invalid")
            rows.append([_mathml_row(_mathml_children(cell)) for cell in cells])
        return [_math_matrix(rows)]
    if local == "mfenced":
        begin = element.get("open", "(")
        end = element.get("close", ")")
        separators = element.get("separators", ",").strip()
        items: list[_MathItem] = []
        for position, child in enumerate(children):
            if position and separators:
                items.append([_math_run(separators[min(position - 1, len(separators) - 1)])])
            items.append(_mathml_item(child))
        return [_math_delimiter(begin, end, _math_row(items))]
    raise ValueError("mathml_element_unsupported")


# -- LaTeX --------------------------------------------------------------------

_LATEX_IDENTIFIERS = {
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ϵ",
    "varepsilon": "ε",
    "zeta": "ζ",
    "eta": "η",
    "theta": "θ",
    "vartheta": "ϑ",
    "iota": "ι",
    "kappa": "κ",
    "lambda": "λ",
    "mu": "μ",
    "nu": "ν",
    "xi": "ξ",
    "pi": "π",
    "varpi": "ϖ",
    "rho": "ρ",
    "varrho": "ϱ",
    "sigma": "σ",
    "varsigma": "ς",
    "tau": "τ",
    "upsilon": "υ",
    "phi": "ϕ",
    "varphi": "φ",
    "chi": "χ",
    "psi": "ψ",
    "omega": "ω",
    "Gamma": "Γ",
    "Delta": "Δ",
    "Theta": "Θ",
    "Lambda": "Λ",
    "Xi": "Ξ",
    "Pi": "Π",
    "Sigma": "Σ",
    "Upsilon": "Υ",
    "Phi": "Φ",
    "Psi": "Ψ",
    "Omega": "Ω",
    "infty": "∞",
    "partial": "∂",
    "nabla": "∇",
    "ell": "ℓ",
    "hbar": "ℏ",
    "emptyset": "∅",
    "aleph": "ℵ",
}
_LATEX_OPERATORS = {
    "pm": "±",
    "mp": "∓",
    "times": "×",
    "cdot": "⋅",
    "div": "÷",
    "ast": "∗",
    "star": "⋆",
    "circ": "∘",
    "bullet": "∙",
    "oplus": "⊕",
    "otimes": "⊗",
    "leq": "≤",
    "le": "≤",
    "geq": "≥",
    "ge": "≥",
    "neq": "≠",
    "ne": "≠",
    "approx": "≈",
    "equiv": "≡",
    "sim": "∼",
    "simeq": "≃",
    "cong": "≅",
    "propto": "∝",
    "ll": "≪",
    "gg": "≫",
    "to": "→",
    "rightarrow": "→",
    "leftarrow": "←",
    "leftrightarrow": "↔",
    "Rightarrow": "⇒",
    "Leftarrow": "⇐",
    "Leftrightarrow": "⇔",
    "implies": "⇒",
    "iff": "⇔",
    "mapsto": "↦",
    "in": "∈",
    "notin": "∉",
    "ni": "∋",
    "subset": "⊂",
    "subseteq": "⊆",
    "supset": "⊃",
    "supseteq": "⊇",
    "cup": "∪",
    "cap": "∩",
    "setminus": "∖",
    "wedge": "∧",
    "land": "∧",
    "vee": "∨",
    "lor": "∨",
    "neg": "¬",
    "forall": "∀",
    "exists": "∃",
    "perp": "⊥",
    "parallel": "∥",
    "mid": "∣",
    "ldots": "…",
    "dots": "…",
    "cdots": "⋯",
    "vdots": "⋮",
    "ddots": "⋱",
    "prime": "′",
    "langle": "⟨",
    "rangle": "⟩",
    "lfloor": "⌊",
    "rfloor": "⌋",
    "lceil": "⌈",
    "rceil": "⌉",
    "{": "{",
    "}": "}",
    "|": "‖",
    "%": "%",
    "#": "#",
    "&": "&",
    "_": "_",
}
_LATEX_NARY = {
    "int": "∫",
    "iint": "∬",
    "iiint": "∭",
    "oint": "∮",
    "sum": "∑",
    "prod": "∏",
    "coprod": "∐",
    "bigcup": "⋃",
    "bigcap": "⋂",
    "bigvee": "⋁",
    "bigwedge": "⋀",
    "bigoplus": "⨁",
    "bigotimes": "⨂",
}
_LATEX_FUNCTIONS = frozenset(
    {
        "sin",
        "cos",
        "tan",
        "cot",
        "sec",
        "csc",
        "arcsin",
        "arccos",
        "arctan",
        "sinh",
        "cosh",
        "tanh",
        "coth",
        "log",
        "ln",
        "lg",
        "exp",
        "det",
        "dim",
        "ker",
        "deg",
        "gcd",
        "arg",
        "max",
        "min",
        "sup",
        "inf",
        "Pr",
        "hom",
    }
)
_LATEX_LIMIT_FUNCTIONS = frozenset({"lim", "limsup", "liminf"})
_LATEX_ACCENTS = {
    "hat": "\u0302",
    "widehat": "\u0302",
    "bar": "\u0305",
    "vec": "\u20d7",
    "dot": "\u0307",
    "ddot": "\u0308",
    "tilde": "\u0303",
    "widetilde": "\u0303",
}
_LATEX_SPACES = {
    ",": "\u2009",
    ":": "\u2005",
    ";": "\u2005",
    "quad": "\u2003",
    "qquad": "\u2003\u2003",
}
_LATEX_IGNORED = frozenset({"!", "displaystyle", "textstyle", "limits", "nolimits"})
_LATEX_FONTS = {
    "mathrm": (None, "p"),
    "operatorname": (None, "p"),
    "mathbf": (None, "b"),
    "mathit": (None, "i"),
    "mathbb": ("double-struck", "p"),
    "mathcal": ("script", None),
    "mathfrak": ("fraktur", None),
}
_LATEX_MATRICES = {
    "matrix": None,
    "pmatrix": ("(", ")"),
    "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"),
    "vmatrix": ("|", "|"),
    "Vmatrix": ("‖", "‖"),
}
_LATEX_CHARACTER_OPERATORS = {"-": "−", "*": "∗", "'": "′"}
_LATEX_DELIMITERS = {
    "(": "(",
    ")": ")",
    "[": "[",
    "]": "]",
    "|": "|",
    ".": "",
    "\\{": "{",
    "\\}": "}",
    "\\|": "‖",
    "\\langle": "⟨",
    "\\rangle": "⟩",
    "\\lfloor": "⌊",
    "\\rfloor": "⌋",
    "\\lceil": "⌈",
    "\\rceil": "⌉",
}
_MAX_LATEX_DEPTH = 64


def _latex_tokens(source: str) -> list[tuple[str, int, int]]:
    """Split LaTeX math into commands (``\\name``), numbers and single characters."""

    tokens: list[tuple[str, int, int]] = []
    index = 0
    while index < len(source):
        character = source[index]
        if character.isspace():
            index += 1
            continue
        end = index + 1
        if character == "\\":
            while end < len(source) and source[end].isascii() and source[end].isalpha():
                end += 1
            if end == index + 1:
                if end >= len(source):
                    raise ValueError("latex_syntax_invalid")
                end += 1
        elif character.isdigit():
            while end < len(source) and (
                source[end].isdigit()
                or (source[end] == "." and end + 1 < len(source) and source[end + 1].isdigit())
            ):
                end += 1
        tokens.append((source[index:end], index, end))
        index = end
    return tokens


_LATEX_TEXT_ESCAPES = {
    "\\{": "{",
    "\\}": "}",
    "\\%": "%",
    "\\&": "&",
    "\\#": "#",
    "\\_": "_",
    "\\ ": " ",
}


class _LatexParser:
    """A recursive-descent reader for the LaTeX math subset Word's equation editor imports."""

    def __init__(self, source: str) -> None:
        self.source = source
        self.spans = _latex_tokens(source)
        self.tokens = [token for token, _start, _end in self.spans]
        self.position = 0
        self.depth = 0

    def peek(self) -> str | None:
        return self.tokens[self.position] if self.position < len(self.tokens) else None

    def take(self) -> str:
        token = self.peek()
        if token is None:
            raise ValueError("latex_syntax_invalid")
        self.position += 1
        return token

    def expect(self, token: str) -> None:
        if self.take() != token:
            raise ValueError("latex_syntax_invalid")

    def parse(self) -> list[ET.Element]:
        items = self.row(frozenset())
        if self.peek() is not None:
            raise ValueError("latex_group_unbalanced")
        return items

    def row(self, stops: frozenset[str]) -> list[ET.Element]:
        """Read items until one of ``stops`` (left unread) or the end of input."""

        self.depth += 1
        if self.depth > _MAX_LATEX_DEPTH:
            raise ValueError("latex_nesting_limit")
        items: list[_MathItem] = []
        while (token := self.peek()) is not None and token not in stops:
            if token == "}":
                raise ValueError("latex_group_unbalanced")
            item = self.atom()
            if item is None:
                continue
            items.append(self.scripts(item))
        self.depth -= 1
        return _math_row(items)

    def group(self) -> list[ET.Element]:
        """One required argument: a braced group or a single atom."""

        if self.peek() == "{":
            self.take()
            items = self.row(frozenset({"}"}))
            self.expect("}")
            return items
        item = self.atom()
        if item is None or isinstance(item, _Nary):
            raise ValueError("latex_argument_invalid")
        return item

    def raw_group(self) -> str:
        """The verbatim source text of one braced argument, for ``\\text`` and names."""

        self.expect("{")
        start = self.spans[self.position - 1][2]
        depth = 1
        while depth:
            token = self.take()
            depth += {"{": 1, "}": -1}.get(token, 0)
        text = self.source[start : self.spans[self.position - 1][1]]
        for escape, character in _LATEX_TEXT_ESCAPES.items():
            text = text.replace(escape, character)
        return text

    def scripts(self, item: _MathItem) -> _MathItem:
        sub: list[ET.Element] | None = None
        explicit: list[ET.Element] | None = None
        primes: list[ET.Element] = []
        limits: str | None = None
        while (token := self.peek()) in {"_", "^", "'", "\\limits", "\\nolimits"}:
            self.take()
            if token == "\\limits":
                limits = "undOvr"
            elif token == "\\nolimits":
                limits = "subSup"
            elif token == "'":
                primes.append(_math_run("′"))
            elif token == "_":
                if sub is not None:
                    raise ValueError("latex_double_subscript")
                sub = self.group()
            else:
                if explicit is not None:
                    raise ValueError("latex_double_superscript")
                explicit = self.group()
        sup = [*primes, *(explicit or [])] if primes or explicit is not None else None
        if isinstance(item, _Nary):
            item.sub, item.sup = sub, sup
            if limits is not None:
                item.limits = limits
            return item
        if sub is None and sup is None:
            return item
        if not item:
            raise ValueError("latex_script_base_missing")
        if len(item) == 1 and item[0].tag == Q(_MATH_NAMESPACE, "r"):
            text = "".join(item[0].itertext())
            if text in {"lim", "lim sup", "lim inf"} and sup is None and sub is not None:
                return [_math_limit("limLow", item, sub)]
        return [_math_scripts(item, sub, sup)]

    def atom(self) -> _MathItem | None:
        token = self.take()
        if token == "{":
            items = self.row(frozenset({"}"}))
            self.expect("}")
            return items
        if token in {"_", "^"}:
            raise ValueError("latex_script_base_missing")
        if token in {"&", "}"} or token == "\\\\":
            raise ValueError("latex_syntax_invalid")
        if not token.startswith("\\"):
            if token[0].isdigit():
                return [_math_run(token)]
            if token.isalpha():
                return [_math_run(token)]
            if token == "~":
                return [_math_run("\u00a0")]
            return [_math_run(_LATEX_CHARACTER_OPERATORS.get(token, token))]
        name = token[1:]
        if name in _LATEX_IGNORED:
            return None
        if name in _LATEX_SPACES:
            return [_math_run(_LATEX_SPACES[name])]
        if name in _LATEX_IDENTIFIERS:
            return [_math_run(_LATEX_IDENTIFIERS[name])]
        if name in _LATEX_OPERATORS:
            return [_math_run(_LATEX_OPERATORS[name])]
        if name in _LATEX_NARY:
            character = _LATEX_NARY[name]
            return _Nary(
                character,
                "subSup" if character in _INTEGRAL_OPERATORS else "undOvr",
                None,
                None,
            )
        if name in _LATEX_FUNCTIONS or name in _LATEX_LIMIT_FUNCTIONS:
            return [
                _math_run(
                    name if name not in {"limsup", "liminf"} else name[:3] + " " + name[3:],
                    style="p",
                )
            ]
        if name in {"frac", "dfrac", "tfrac"}:
            return [_math_fraction(self.group(), self.group())]
        if name == "sqrt":
            degree = None
            if self.peek() == "[":
                self.take()
                degree = self.row(frozenset({"]"}))
                self.expect("]")
            return [_math_radical(self.group(), degree)]
        if name in _LATEX_ACCENTS:
            return [_math_accent(_LATEX_ACCENTS[name], self.group())]
        if name in _LATEX_FONTS:
            script, style = _LATEX_FONTS[name]
            if self.peek() != "{":
                raise ValueError("latex_argument_invalid")
            text = self.raw_group()
            if not text:
                raise ValueError("latex_argument_invalid")
            return [_math_run(text, style=style, script=script)]
        if name in {"text", "textrm", "mbox"}:
            text = self.raw_group()
            if not text:
                raise ValueError("latex_argument_invalid")
            return [_math_run(text, style="nor")]
        if name == "left":
            begin = self.delimiter()
            items = self.row(frozenset({"\\right"}))
            self.expect("\\right")
            return [_math_delimiter(begin, self.delimiter(), items)]
        if name == "begin":
            return [self.environment()]
        raise ValueError("latex_command_unsupported")

    def delimiter(self) -> str:
        token = self.take()
        if token not in _LATEX_DELIMITERS:
            raise ValueError("latex_delimiter_unsupported")
        return _LATEX_DELIMITERS[token]

    def environment(self) -> ET.Element:
        name = self.raw_group()
        if name not in _LATEX_MATRICES:
            raise ValueError("latex_environment_unsupported")
        rows: list[list[list[ET.Element]]] = [[]]
        stops = frozenset({"&", "\\\\", "\\end"})
        while True:
            rows[-1].append(self.row(stops))
            token = self.take()
            if token == "\\\\":
                if self.peek() == "\\end":
                    continue
                rows.append([])
            elif token == "\\end":
                break
        if self.raw_group() != name:
            raise ValueError("latex_environment_unbalanced")
        if len({len(row) for row in rows}) != 1:
            raise ValueError("latex_matrix_ragged")
        matrix = _math_matrix(rows)
        fences = _LATEX_MATRICES[name]
        if fences is None:
            return matrix
        return _math_delimiter(fences[0], fences[1], [matrix])


def latex_to_omml(latex: str) -> ET.Element:
    """Convert one LaTeX math expression to an ``m:oMath`` element.

    The admitted subset is what equation editors exchange: fractions, radicals,
    sub- and superscripts and primes, integrals, sums and other n-ary operators
    (``\\limits``/``\\nolimits``), ``\\left``/``\\right`` delimiters, the matrix
    environments (``matrix``, ``pmatrix``, ``bmatrix``, ``Bmatrix``, ``vmatrix``,
    ``Vmatrix``), accents, Greek letters, named functions, ``\\text`` and the
    ``\\mathrm``/``\\mathbf``/``\\mathbb``/``\\mathcal`` alphabets. An n-ary
    operator's operand is the rest of its group up to the next relation. Anything
    else is refused with a ``latex_*`` code rather than written approximately.
    """

    if not latex.strip():
        raise ValueError("math_input_empty")
    items = _LatexParser(latex).parse()
    if not items:
        raise ValueError("math_input_empty")
    return _math_container("oMath", items)


def _math_payload(request: MathRequest, word_namespace: str) -> tuple[bytes, str, str]:
    """Return schema-positioned OMML bytes, input format, and root container."""

    inputs = [
        item for item in (request.omml_xml, request.mathml_xml, request.latex) if item is not None
    ]
    if len(inputs) != 1:
        raise ValueError("math_input_ambiguous")
    try:
        if request.omml_xml is not None:
            root = ET.fromstring(request.omml_xml)
            input_format = "omml"
        elif request.latex is not None:
            root = latex_to_omml(request.latex)
            input_format = "latex"
        else:
            mathml = ET.fromstring(request.mathml_xml or "")
            root = _math_container("oMath", _mathml_to_omml(mathml))
            if _mathml_local(mathml) == "math" and mathml.get("display") == "block":
                root = _math_container("oMathPara", [root])
            input_format = "mathml"
    except ET.ParseError as error:
        raise ValueError("math_xml_invalid") from error
    if not isinstance(root.tag, str) or not root.tag.startswith("{"):
        raise ValueError("omml_namespace_invalid")
    namespace, container = root.tag[1:].split("}", 1)
    if namespace not in {_MATH_NAMESPACE, _STRICT_MATH_NAMESPACE} or container not in {
        "oMath",
        "oMathPara",
    }:
        raise ValueError("omml_container_invalid")
    target_math = (
        _STRICT_MATH_NAMESPACE
        if word_namespace == "http://purl.oclc.org/ooxml/wordprocessingml/main"
        else _MATH_NAMESPACE
    )
    namespace_map = {
        _MATH_NAMESPACE: target_math,
        _STRICT_MATH_NAMESPACE: target_math,
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main": word_namespace,
        "http://purl.oclc.org/ooxml/wordprocessingml/main": word_namespace,
    }
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        for old, new in namespace_map.items():
            if element.tag.startswith("{" + old + "}"):
                element.tag = Q(new, element.tag.split("}", 1)[1])
                break
        for name, value in list(element.attrib.items()):
            for old, new in namespace_map.items():
                if name.startswith("{" + old + "}"):
                    del element.attrib[name]
                    element.set(Q(new, name.split("}", 1)[1]), value)
                    break
    ET.register_namespace("m", target_math)
    ET.register_namespace("w", word_namespace)
    return ET.tostring(root, encoding="utf-8"), input_format, container


def _paragraph_parent(root: ET.Element, paragraph: ET.Element) -> ET.Element | None:
    return next((parent for parent in root.iter() if paragraph in parent), None)


def _patch_paragraph_math(
    document_xml: bytes,
    *,
    word_namespace: str,
    paragraph: int,
    payload: bytes,
    container: str | None,
    operation: str = "insert",
) -> bytes | None:
    """Patch only lexical target bytes; refuse missing or ambiguous math targets."""

    parser = expat.ParserCreate(namespace_separator="}")
    stack: list[str] = []
    found = 0
    target_depth: int | None = None
    insertion: int | None = None
    target_start: int | None = None
    target_end: int | None = None
    math_spans: list[tuple[int, int, str]] = []
    math_start: int | None = None
    math_depth: int | None = None
    display_adjacent = False
    math_names = {
        f"{ns}" + "}" + name
        for ns in (_MATH_NAMESPACE, _STRICT_MATH_NAMESPACE)
        for name in ("oMath", "oMathPara")
    }
    target_name = f"{word_namespace}" + "}p"

    def start_element(name: str, attributes: dict[str, str]) -> None:
        nonlocal found, target_depth, target_start, math_start, math_depth, display_adjacent
        del attributes
        stack.append(name)
        if name == target_name:
            found += 1
            if found == paragraph:
                target_depth = len(stack)
                target_start = parser.CurrentByteIndex
        direct_math = target_depth is not None and len(stack) == target_depth + 1
        adjacent_math = display_adjacent and name in {
            f"{ns}" + "}oMathPara" for ns in (_MATH_NAMESPACE, _STRICT_MATH_NAMESPACE)
        }
        if name in math_names and (direct_math or adjacent_math) and math_depth is None:
            math_start = parser.CurrentByteIndex
            math_depth = len(stack)
        if display_adjacent:
            display_adjacent = False

    def end_element(name: str) -> None:
        nonlocal insertion, target_depth, target_end, math_start, math_depth, display_adjacent
        if math_depth == len(stack):
            assert math_start is not None
            end = parser.CurrentByteIndex
            if document_xml[end : end + 2] == b"</":
                end = document_xml.find(b">", end) + 1
            math_spans.append((math_start, end, name.split("}", 1)[1]))
            math_start = math_depth = None
        if name == target_name and target_depth == len(stack):
            insertion = parser.CurrentByteIndex
            if document_xml[insertion : insertion + 2] == b"</":
                target_end = document_xml.find(b">", insertion) + 1
            else:
                target_end = insertion
            display_adjacent = True
            target_depth = None
        stack.pop()

    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    try:
        parser.Parse(document_xml, True)
    except expat.ExpatError:
        return None
    if insertion is None:
        return None
    if operation != "insert":
        if not math_spans:
            raise ValueError("math_target_missing")
        if len(math_spans) != 1:
            raise ValueError("math_target_ambiguous")
        start, end, original_container = math_spans[0]
        if operation == "edit" and container != original_container:
            raise ValueError("math_container_position_mismatch")
        return (
            document_xml[:start] + (b"" if operation == "delete" else payload) + document_xml[end:]
        )
    if target_start is not None and target_end == insertion:
        opening = document_xml[target_start:insertion]
        if not opening.rstrip().endswith(b"/>"):
            return None
        name = opening[1:].split(None, 1)[0].rstrip(b"/>")
        expanded = opening.rstrip()[:-2] + b">"
        closing = b"</" + name + b">"
        addition = payload + closing if container == "oMath" else closing + payload
        return document_xml[:target_start] + expanded + addition + document_xml[insertion:]
    if container == "oMath":
        return document_xml[:insertion] + payload + document_xml[insertion:]
    end = document_xml.find(b">", insertion)
    if end < insertion:
        return None
    return document_xml[: end + 1] + payload + document_xml[end + 1 :]


def author_paragraph_math(request: MathRequest) -> MathAuthoringReport:
    """Mutate one equation atomically, preserving every non-target package byte."""

    source = request.source.resolve()
    destination = request.destination.resolve()
    if source == destination:
        return _math_refusal(request, "source_destination_same")
    if not source.is_file():
        return _math_refusal(request, "source_missing")
    if not destination.parent.is_dir():
        return _math_refusal(request, "destination_parent_missing")
    if _sha256_file(source) != request.source_sha256:
        return _math_refusal(request, "source_hash_mismatch")
    if request.operation not in ("insert", "edit", "delete"):
        return _math_refusal(request, "math_operation_invalid")
    baseline = inspect_docx(source)
    if not baseline.valid or baseline.file_sha256 != request.source_sha256:
        return _math_refusal(request, "source_inspection_invalid")
    try:
        graph = PackageGraph.from_file(source)
        document_xml = graph.read_part_bytes("word/document.xml")
        root = ET.fromstring(document_xml)
    except PackageError:
        return _math_refusal(request, "source_document_missing")
    except ET.ParseError:
        return _math_refusal(request, "source_document_invalid")
    word_namespace = _word_namespace(root)
    if word_namespace is None:
        return _math_refusal(request, "source_word_namespace_unsupported")
    try:
        if request.operation == "delete":
            payload, input_format, container = b"", None, None
        else:
            payload, input_format, container = _math_payload(request, word_namespace)
    except ValueError as error:
        return _math_refusal(request, str(error))
    paragraph = _selected_paragraph(root, request, word_namespace)
    if paragraph is None:
        return _math_refusal(request, "paragraph_coordinate_missing")
    if _paragraph_parent(root, paragraph) is not root.find(Q(word_namespace, "body")):
        return _math_refusal(request, "paragraph_not_body")
    if _paragraph_structure_fingerprint(paragraph) != request.structure_fingerprint_sha256:
        return _math_refusal(request, "paragraph_structure_mismatch")
    if _has_unsupported_ancestor(root, paragraph, word_namespace):
        return _math_refusal(request, "paragraph_context_unsupported")
    try:
        patched = _patch_paragraph_math(
            document_xml,
            word_namespace=word_namespace,
            paragraph=request.paragraph,
            payload=payload,
            container=container,
            operation=request.operation,
        )
    except ValueError as error:
        return _math_refusal(request, str(error))
    if patched is None:
        return _math_refusal(request, "paragraph_math_patch_unsupported")
    staged_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as staged:
            staged_path = Path(staged.name)
        _write_staged_package(source, staged_path, patched)
        candidate = inspect_docx(staged_path)
        candidate_xml: bytes | None = None
        inserted = request.operation == "delete"
        if candidate.valid:
            try:
                candidate_xml = PackageGraph.from_file(staged_path).read_part_bytes(
                    "word/document.xml"
                )
                candidate_root = ET.fromstring(candidate_xml)
                candidate_paragraph = _selected_paragraph(candidate_root, request, word_namespace)
                math_namespace = (
                    _STRICT_MATH_NAMESPACE
                    if word_namespace == "http://purl.oclc.org/ooxml/wordprocessingml/main"
                    else _MATH_NAMESPACE
                )
                if container == "oMath":
                    inserted = bool(
                        candidate_paragraph is not None
                        and any(
                            child.tag == Q(math_namespace, "oMath") for child in candidate_paragraph
                        )
                    )
                elif container == "oMathPara":
                    body = candidate_root.find(Q(word_namespace, "body"))
                    inserted = bool(
                        body is not None
                        and any(child.tag == Q(math_namespace, "oMathPara") for child in body)
                    )
            except (ET.ParseError, PackageError):
                pass
        document_preserved = candidate_xml == patched
        parts_preserved = _non_target_part_bytes_preserved(source, staged_path)
        source_still_bound = _sha256_file(source) == request.source_sha256
        if not (
            candidate.valid
            and inserted
            and document_preserved
            and parts_preserved
            and source_still_bound
        ):
            return _math_refusal(
                request,
                "staged_verification_failed" if source_still_bound else "source_hash_changed",
            )
        os.replace(staged_path, destination)
        staged_path = None
        return MathAuthoringReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            target={"paragraph": request.paragraph},
            input_format=input_format,
            inserted_container=container,
            structure_fingerprint_sha256=request.structure_fingerprint_sha256,
            destination_sha256=candidate.file_sha256,
            non_target_document_bytes_preserved=True,
            non_target_parts_preserved=True,
        )
    finally:
        if staged_path is not None:
            staged_path.unlink(missing_ok=True)


def insert_equation(
    source: str | Path,
    destination: str | Path,
    *,
    paragraph: int = 1,
    latex: str | None = None,
    mathml: str | None = None,
    omml: str | None = None,
) -> MathAuthoringReport:
    """Insert one equation, given as LaTeX, MathML or OMML, into a body paragraph.

    Exactly one of ``latex``, ``mathml`` and ``omml`` is accepted. The request is
    bound to the source's current digest and the paragraph's structure fingerprint,
    then published through :func:`author_paragraph_math`, so every other package
    byte is preserved and nothing is written on refusal.
    """

    source_path = Path(source)
    request = MathRequest(
        source=source_path,
        destination=Path(destination),
        source_sha256=_sha256_file(source_path) if source_path.is_file() else "",
        structure_fingerprint_sha256="",
        paragraph=paragraph,
        omml_xml=omml,
        mathml_xml=mathml,
        latex=latex,
    )
    if sum(item is not None for item in (latex, mathml, omml)) != 1:
        return _math_refusal(request, "math_input_ambiguous")
    if not source_path.is_file():
        return _math_refusal(request, "source_missing")
    if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 1:
        return _math_refusal(request, "math_target_paragraph_invalid")
    fingerprint = paragraph_structure_fingerprint(source_path, paragraph=paragraph)
    if fingerprint is None:
        return _math_refusal(request, "paragraph_coordinate_missing")
    return author_paragraph_math(replace(request, structure_fingerprint_sha256=fingerprint))
