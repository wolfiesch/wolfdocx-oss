"""Atomic, source-bound report editing across DOCX body, headers, and footers."""

from __future__ import annotations

import copy
import hashlib
import io
import json
import os
import tempfile
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .authoring import (
    _MAX_TEXT_CHARACTERS,
    _PUBLISHED,
    _REFUSED,
    _cell_structure_fingerprint,
    _effective_xml_space_by_element,
    _has_unsupported_ancestor,
    _patch_xml_texts,
    _sha256_file,
    _simple_text_node,
    _span_text_nodes,
    _span_text_occurrences,
    _SpanTextNode,
    _text_replacements_preserve_whitespace,
    _unsupported_context_by_element,
    _word_namespace,
)
from .canonical import sha256_bytes
from .inspect import PackageLimits, Q, _inspect_bytes, _Inspection
from .package import PackageError, PackageGraph

_DOCUMENT_PART = "word/document.xml"
_HEADER_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)
_FOOTER_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
)


@dataclass(frozen=True)
class ReportEdit:
    """One text replacement resolved only against the original source package."""

    kind: str
    part: str
    expected_old_text: str
    structure_fingerprint_sha256: str
    text: str
    table: int | None = None
    row: int | None = None
    cell: int | None = None
    paragraph: int | None = None

    def target_dict(self) -> dict[str, Any]:
        if self.kind == "cell":
            return {
                "kind": self.kind,
                "part": self.part,
                "target": {"table": self.table, "row": self.row, "cell": self.cell},
            }
        return {"kind": self.kind, "part": self.part, "paragraph": self.paragraph}


@dataclass(frozen=True)
class ReportRequest:
    source: Path
    destination: Path
    source_sha256: str
    edits: tuple[ReportEdit, ...]


@dataclass(frozen=True)
class ReportAuthoringReport:
    status: str
    code: str | None
    source_sha256: str
    edit_targets: list[dict[str, Any]]
    destination_sha256: str | None = None
    target_change_reports: list[dict[str, Any]] | None = None
    unselected_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "report",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "edit_targets": self.edit_targets,
            "destination_sha256": self.destination_sha256,
            "target_change_reports": self.target_change_reports,
            "unselected_parts_preserved": self.unselected_parts_preserved,
        }


@dataclass(frozen=True)
class _TextRegion:
    node: ET.Element
    start: int
    end: int
    text: str


@dataclass(frozen=True)
class _ResolvedEdit:
    edit: ReportEdit
    element: ET.Element
    source_text: str
    start: int
    end: int
    regions: tuple[_TextRegion, ...]


def _require_sha256(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64:
        raise ValueError(f"{label}_invalid")
    try:
        int(value, 16)
    except ValueError as error:
        raise ValueError(f"{label}_invalid") from error
    return value.lower()


def _coordinate(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError(f"{label}_invalid")
    return value


def _parse_edit(value: Any) -> ReportEdit:
    if not isinstance(value, Mapping):
        raise TypeError("report_edit_invalid")
    kind = value.get("kind")
    if kind not in {"cell", "text"}:
        raise ValueError("report_edit_kind_invalid")
    common = {"kind", "expected_old_text", "structure_fingerprint_sha256", "text"}
    allowed = common | ({"target", "part"} if kind == "cell" else {"paragraph", "part"})
    if set(value) - allowed or not common <= set(value):
        raise ValueError("report_edit_fields_invalid")
    part = value.get("part", _DOCUMENT_PART)
    if not isinstance(part, str) or not part:
        raise ValueError("report_part_invalid")
    expected_old_text = value["expected_old_text"]
    text = value["text"]
    if not isinstance(expected_old_text, str) or not isinstance(text, str):
        raise TypeError("report_text_invalid")
    if not expected_old_text:
        raise ValueError("report_expected_old_text_empty")
    if len(expected_old_text) > _MAX_TEXT_CHARACTERS:
        raise ValueError("report_expected_old_text_character_limit")
    if len(text) > _MAX_TEXT_CHARACTERS:
        raise ValueError("report_text_character_limit")
    fingerprint = _require_sha256(
        value["structure_fingerprint_sha256"], "report_structure_fingerprint_sha256"
    )
    if kind == "cell":
        target = value.get("target")
        if not isinstance(target, Mapping) or set(target) != {"table", "row", "cell"}:
            raise ValueError("report_cell_target_invalid")
        return ReportEdit(
            kind=kind,
            part=part,
            expected_old_text=expected_old_text,
            structure_fingerprint_sha256=fingerprint,
            text=text,
            table=_coordinate(target["table"], "report_cell_table"),
            row=_coordinate(target["row"], "report_cell_row"),
            cell=_coordinate(target["cell"], "report_cell_cell"),
        )
    if "paragraph" not in value:
        raise ValueError("report_text_paragraph_missing")
    return ReportEdit(
        kind=kind,
        part=part,
        expected_old_text=expected_old_text,
        structure_fingerprint_sha256=fingerprint,
        text=text,
        paragraph=_coordinate(value["paragraph"], "report_text_paragraph"),
    )


def parse_report_request(value: Any) -> ReportRequest:
    """Parse the strict JSON request for one atomic, report-wide authoring operation."""

    if not isinstance(value, Mapping) or set(value) != {
        "source",
        "destination",
        "source_sha256",
        "edits",
    }:
        raise ValueError("report_request_fields_invalid")
    source, destination, edits = value["source"], value["destination"], value["edits"]
    if (
        not isinstance(source, str)
        or not source
        or not isinstance(destination, str)
        or not destination
    ):
        raise ValueError("report_path_invalid")
    if not isinstance(edits, list) or not edits:
        raise ValueError("report_edits_invalid")
    return ReportRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=_require_sha256(value["source_sha256"], "report_source_sha256"),
        edits=tuple(_parse_edit(item) for item in edits),
    )


def load_report_request(path: str | Path) -> ReportRequest:
    with Path(path).open(encoding="utf-8") as stream:
        return parse_report_request(json.load(stream))


def _report_refusal(request: ReportRequest, code: str) -> ReportAuthoringReport:
    return ReportAuthoringReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        edit_targets=[edit.target_dict() for edit in request.edits],
    )


def _select_cell(
    root: ET.Element, edit: ReportEdit, namespace: str
) -> ET.Element | None:
    assert edit.table is not None and edit.row is not None and edit.cell is not None
    def word(local):
        return Q(namespace, local)
    tables = list(root.iter(word("tbl")))
    if edit.table > len(tables):
        return None
    rows = tables[edit.table - 1].findall(word("tr"))
    if edit.row > len(rows):
        return None
    cells = rows[edit.row - 1].findall(word("tc"))
    return cells[edit.cell - 1] if edit.cell <= len(cells) else None


def _select_paragraph(
    root: ET.Element, edit: ReportEdit, namespace: str
) -> ET.Element | None:
    assert edit.paragraph is not None
    paragraphs = list(root.iter(Q(namespace, "p")))
    return paragraphs[edit.paragraph - 1] if edit.paragraph <= len(paragraphs) else None


def _supported_part(graph: PackageGraph, part: str) -> bool:
    package_part = graph.part(part)
    if package_part is None:
        return False
    if part == _DOCUMENT_PART:
        return True
    if part.startswith("word/header") and part.endswith(".xml"):
        expected_type, relationship_kind, reference_kind = (
            _HEADER_CONTENT_TYPE,
            "header",
            "headerReference",
        )
    elif part.startswith("word/footer") and part.endswith(".xml"):
        expected_type, relationship_kind, reference_kind = (
            _FOOTER_CONTENT_TYPE,
            "footer",
            "footerReference",
        )
    else:
        return False
    if package_part.content_type != expected_type:
        return False
    relationship_ids = {
        relationship.id
        for relationship in graph.relationships_for(_DOCUMENT_PART)
        if not relationship.is_external
        and relationship.resolved_target() == part
        and relationship.type.rsplit("/", 1)[-1] == relationship_kind
    }
    if not relationship_ids:
        return False
    try:
        document_root = ET.fromstring(graph.read_part_bytes(_DOCUMENT_PART))
    except (ET.ParseError, PackageError):
        return False
    return any(
        element.tag.rsplit("}", 1)[-1] == reference_kind
        and any(
            attribute.rsplit("}", 1)[-1] == "id" and value in relationship_ids
            for attribute, value in element.attrib.items()
        )
        for element in document_root.iter()
    )


def _regions_for_text(
    nodes: list[Any], *, start: int, end: int, replacement: str
) -> tuple[_TextRegion, ...]:
    regions: list[_TextRegion] = []
    for node in nodes:
        if node.start >= end or start >= node.end:
            continue
        local_start = max(start, node.start) - node.start
        local_end = min(end, node.end) - node.start
        regions.append(
            _TextRegion(
                node=node.node,
                start=local_start,
                end=local_end,
                text=replacement if node.start <= start < node.end else "",
            )
        )
    return tuple(regions)


def _resolve_edit(
    root: ET.Element, edit: ReportEdit, namespace: str
) -> tuple[_ResolvedEdit | None, str | None]:
    if edit.kind == "cell":
        cell = _select_cell(root, edit, namespace)
        if cell is None:
            return None, "cell_coordinate_missing"
        if _cell_structure_fingerprint(cell) != edit.structure_fingerprint_sha256:
            return None, "cell_structure_mismatch"
        if _has_unsupported_ancestor(root, cell, namespace):
            return None, "cell_context_unsupported"
        text_node = _simple_text_node(cell, namespace)
        if text_node is None:
            return None, "cell_structure_unsupported"
        node, old_text = text_node
        if old_text != edit.expected_old_text:
            return None, "expected_old_text_mismatch"
        if old_text == edit.text:
            return None, "replacement_unchanged"
        return _ResolvedEdit(
            edit,
            cell,
            old_text,
            0,
            len(old_text),
            (_TextRegion(node, 0, len(old_text), edit.text),),
        ), None

    paragraph = _select_paragraph(root, edit, namespace)
    if paragraph is None:
        return None, "paragraph_coordinate_missing"
    if _cell_structure_fingerprint(paragraph) != edit.structure_fingerprint_sha256:
        return None, "paragraph_structure_mismatch"
    if _has_unsupported_ancestor(root, paragraph, namespace):
        return None, "paragraph_context_unsupported"
    nodes = _span_text_nodes(paragraph, namespace)
    if nodes is None:
        return None, "text_structure_unsupported"
    source_text = "".join(node.text for node in nodes)
    start = source_text.find(edit.expected_old_text)
    if start < 0:
        return None, "expected_old_text_mismatch"
    if source_text.find(edit.expected_old_text, start + 1) >= 0:
        return None, "text_occurrence_ambiguous"
    end = start + len(edit.expected_old_text)
    if edit.expected_old_text == edit.text:
        return None, "replacement_unchanged"
    return _ResolvedEdit(
        edit,
        paragraph,
        source_text,
        start,
        end,
        _regions_for_text(nodes, start=start, end=end, replacement=edit.text),
    ), None


def _parents(root: ET.Element) -> dict[int, ET.Element]:
    return {id(child): parent for parent in root.iter() for child in parent}


def _is_ancestor(
    ancestor: ET.Element, element: ET.Element, parents: Mapping[int, ET.Element]
) -> bool:
    current: ET.Element | None = element
    while current is not None:
        if current is ancestor:
            return True
        current = parents.get(id(current))
    return False


def _has_overlapping_targets(
    resolved: list[_ResolvedEdit], roots: Mapping[str, ET.Element]
) -> bool:
    # Parent maps cost one whole-tree walk, so build them only for parts holding a pair.
    parents: dict[str, dict[int, ET.Element]] = {}
    for index, current in enumerate(resolved):
        for previous in resolved[:index]:
            if current.edit.part != previous.edit.part:
                continue
            if (
                current.edit.kind == previous.edit.kind == "text"
                and current.element is previous.element
            ):
                if current.start >= previous.end or previous.start >= current.end:
                    continue
                return True
            part = current.edit.part
            if part not in parents:
                parents[part] = _parents(roots[part])
            part_parents = parents[part]
            if _is_ancestor(
                current.element, previous.element, part_parents
            ) or _is_ancestor(previous.element, current.element, part_parents):
                return True
    return False


def _replacement_texts(resolved: list[_ResolvedEdit]) -> dict[int, str] | None:
    by_node: dict[int, tuple[ET.Element, list[_TextRegion]]] = {}
    for item in resolved:
        for region in item.regions:
            node, regions = by_node.setdefault(id(region.node), (region.node, []))
            assert node is region.node
            regions.append(region)
    replacements: dict[int, str] = {}
    for node_id, (node, regions) in by_node.items():
        cursor = 0
        pieces: list[str] = []
        for region in sorted(regions, key=lambda item: (item.start, item.end)):
            if region.start < cursor or region.end < region.start:
                return None
            pieces.extend((node.text or "")[cursor : region.start])
            pieces.append(region.text)
            cursor = region.end
        pieces.append((node.text or "")[cursor:])
        replacements[node_id] = "".join(pieces)
    return replacements


def _patch_part(
    root: ET.Element,
    part_bytes: bytes,
    namespace: str,
    replacements: Mapping[int, str],
    nodes: Mapping[int, ET.Element],
) -> bytes | None:
    selected: list[_SpanTextNode] = []
    for node_id in replacements:
        node = nodes.get(node_id)
        if node is None:
            return None
        selected.append(_SpanTextNode(node=node, text="", start=0, end=0))
    # The occurrence walk refuses any selected node that is not inside ``root``.
    try:
        occurrences = _span_text_occurrences(root, selected)
    except ValueError:
        return None
    return _patch_xml_texts(
        part_bytes,
        word_namespace=namespace,
        replacements={
            occurrences[node_id]: text for node_id, text in replacements.items()
        },
    )


def _write_staged_package(
    source: Path, staged: Path, replacements: Mapping[str, bytes]
) -> None:
    with (
        zipfile.ZipFile(source) as baseline,
        zipfile.ZipFile(staged, "w", compression=zipfile.ZIP_DEFLATED) as rewritten,
    ):
        for info in baseline.infolist():
            rewritten.writestr(
                copy.copy(info),
                replacements.get(info.filename, baseline.read(info.filename)),
            )


def _unselected_part_bytes_preserved(
    source: Path,
    candidate: Path,
    selected_parts: set[str],
    *,
    allowed_additions: frozenset[str] = frozenset(),
) -> bool:
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(candidate) as rewritten:
        baseline_names = set(baseline.namelist())
        candidate_names = set(rewritten.namelist())
        if baseline_names - candidate_names:
            return False
        if candidate_names - baseline_names - allowed_additions:
            return False
        return all(
            baseline.read(name) == rewritten.read(name)
            for name in baseline_names
            if name not in selected_parts
        )


def _report_unselected_preserved(
    source_bytes: bytes,
    baseline: _Inspection,
    staged_bytes: bytes,
    candidate: _Inspection,
    selected_parts: set[str],
) -> bool:
    """Prove every unselected entry kept its bytes, from digests both inspections hold.

    Each inspection hashed every package part from the bytes it validated, so equal
    SHA-256 digests prove equal bytes without decompressing either archive again.
    Entries the inspections did not hash (non-part members) are compared directly.
    """
    with (
        zipfile.ZipFile(io.BytesIO(source_bytes)) as source_zip,
        zipfile.ZipFile(io.BytesIO(staged_bytes)) as staged_zip,
    ):
        names = set(source_zip.namelist())
        if names != set(staged_zip.namelist()):
            return False
        before, after = baseline.part_sha256, candidate.part_sha256
        for name in names - selected_parts:
            source_digest, staged_digest = before.get(name), after.get(name)
            if source_digest is not None and staged_digest is not None:
                if source_digest != staged_digest:
                    return False
            elif source_zip.read(name) != staged_zip.read(name):
                return False
    return True


def _bound_source_bytes(source: Path, source_sha256: str) -> bytes | str:
    """Read the source once, or the refusal code for an oversized or unbound source."""
    if source.stat().st_size > PackageLimits().max_archive_size:
        # Never materialize an oversized package; its inspection refuses it unread.
        if _sha256_file(source) != source_sha256:
            return "source_hash_mismatch"
        return "source_inspection_invalid"
    data = source.read_bytes()
    return data if sha256_bytes(data) == source_sha256 else "source_hash_mismatch"


def _candidate_target_reports(
    candidate_roots: Mapping[str, ET.Element],
    resolved: list[_ResolvedEdit],
    expected_texts: Mapping[int, str],
) -> list[dict[str, Any]] | None:
    """Re-read every target from the staged inspection's read-only trees."""
    reports: list[dict[str, Any]] = []
    for item in resolved:
        root = candidate_roots.get(item.edit.part)
        if root is None:
            return None
        namespace = _word_namespace(root)
        if namespace is None:
            return None
        if item.edit.kind == "cell":
            element = _select_cell(root, item.edit, namespace)
            text_node = (
                _simple_text_node(element, namespace) if element is not None else None
            )
            actual = text_node[1] if text_node is not None else None
        else:
            element = _select_paragraph(root, item.edit, namespace)
            nodes = _span_text_nodes(element, namespace) if element is not None else None
            actual = "".join(node.text for node in nodes) if nodes is not None else None
        expected = expected_texts[id(item)]
        structure_ok = (
            element is not None
            and _cell_structure_fingerprint(element)
            == item.edit.structure_fingerprint_sha256
        )
        reports.append(
            {
                "target": item.edit.target_dict(),
                "expected_text_sha256": hashlib.sha256(
                    expected.encode("utf-8")
                ).hexdigest(),
                "candidate_text_sha256": hashlib.sha256(
                    actual.encode("utf-8")
                ).hexdigest()
                if actual is not None
                else None,
                "structure_preserved": structure_ok,
                "passed": actual == expected and structure_ok,
            }
        )
    return reports


def report_target_structure_fingerprint(
    source: str | Path,
    *,
    kind: str,
    target: Mapping[str, Any] | int,
    part: str = _DOCUMENT_PART,
) -> str | None:
    """Return a text-masked bound fingerprint for a supported report target."""

    document = Path(source)
    if kind not in {"cell", "text"} or not isinstance(part, str):
        return None
    try:
        graph = PackageGraph.from_file(document)
        if not _supported_part(graph, part):
            return None
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, PackageError):
        return None
    namespace = _word_namespace(root)
    if namespace is None:
        return None
    if kind == "cell":
        if not isinstance(target, Mapping) or set(target) != {"table", "row", "cell"}:
            return None
        try:
            edit = ReportEdit(
                "cell",
                part,
                "x",
                "0" * 64,
                "y",
                **{key: _coordinate(target[key], key) for key in target},
            )
        except (TypeError, ValueError):
            return None
        element = _select_cell(root, edit, namespace)
    else:
        if isinstance(target, bool) or not isinstance(target, int) or target <= 0:
            return None
        element = _select_paragraph(
            root,
            ReportEdit("text", part, "x", "0" * 64, "y", paragraph=target),
            namespace,
        )
    return _cell_structure_fingerprint(element) if element is not None else None


def discover_report_targets(
    source: str | Path, *, part: str = _DOCUMENT_PART
) -> dict[str, Any]:
    """List source-bound targets supported by the read-only report authoring surface."""

    document = Path(source)
    if not document.is_file():
        return {
            "schema_version": 1,
            "operation": "report-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "source_missing",
        }
    try:
        graph = PackageGraph.from_file(document)
        source_sha256 = graph.file_sha256
        if not _supported_part(graph, part):
            return {
                "schema_version": 1,
                "operation": "report-targets",
                "passed": False,
                "status": _REFUSED,
                "code": "report_part_unsupported",
                "source_sha256": source_sha256,
                "part": part,
            }
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, PackageError):
        return {
            "schema_version": 1,
            "operation": "report-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "source_part_invalid",
            "part": part,
        }
    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "report-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "source_word_namespace_unsupported",
            "source_sha256": source_sha256,
            "part": part,
        }

    targets: list[dict[str, Any]] = []
    effective_spaces = _effective_xml_space_by_element(root)
    unsupported_contexts = _unsupported_context_by_element(root, namespace)
    def word(local):
        return Q(namespace, local)
    for table_index, table in enumerate(root.iter(word("tbl")), start=1):
        for row_index, row in enumerate(table.findall(word("tr")), start=1):
            for cell_index, cell in enumerate(row.findall(word("tc")), start=1):
                if unsupported_contexts.get(id(cell), True):
                    continue
                simple = _simple_text_node(cell, namespace)
                if simple is None:
                    continue
                node, old_text = simple
                if not old_text or not _text_replacements_preserve_whitespace(
                    root, {id(node): old_text}, effective_spaces=effective_spaces
                ):
                    continue
                targets.append(
                    {
                        "kind": "cell",
                        "part": part,
                        "target": {
                            "table": table_index,
                            "row": row_index,
                            "cell": cell_index,
                        },
                        "expected_old_text": old_text,
                        "structure_fingerprint_sha256": _cell_structure_fingerprint(
                            cell
                        ),
                    }
                )
    for paragraph_index, paragraph in enumerate(root.iter(word("p")), start=1):
        if unsupported_contexts.get(id(paragraph), True):
            continue
        nodes = _span_text_nodes(paragraph, namespace)
        if nodes is None:
            continue
        old_text = "".join(node.text for node in nodes)
        node_texts = {id(node.node): node.text for node in nodes}
        if not old_text or not _text_replacements_preserve_whitespace(
            root, node_texts, effective_spaces=effective_spaces
        ):
            continue
        targets.append(
            {
                "kind": "text",
                "part": part,
                "paragraph": paragraph_index,
                "expected_old_text": old_text,
                "structure_fingerprint_sha256": _cell_structure_fingerprint(paragraph),
            }
        )
    return {
        "schema_version": 1,
        "operation": "report-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_report(request: ReportRequest) -> ReportAuthoringReport:
    """Resolve every edit on the immutable source, verify once, then publish once."""

    source = request.source.resolve()
    destination = request.destination.resolve()
    if source == destination:
        return _report_refusal(request, "source_destination_same")
    if not source.is_file():
        return _report_refusal(request, "source_missing")
    if not destination.parent.is_dir():
        return _report_refusal(request, "destination_parent_missing")
    source_bytes = _bound_source_bytes(source, request.source_sha256)
    if isinstance(source_bytes, str):
        return _report_refusal(request, source_bytes)
    # Inspection, graph, and part reads all bind to these exact hashed bytes.
    baseline = _inspect_bytes(
        source_bytes, display_path=str(source), file_sha256=request.source_sha256
    )
    if not baseline.valid:
        return _report_refusal(request, "source_inspection_invalid")
    try:
        graph = PackageGraph.from_bytes(
            source_bytes, path=source, file_sha256=request.source_sha256
        )
    except PackageError:
        return _report_refusal(request, "source_package_invalid")

    source_parts: dict[str, bytes] = {}
    roots: dict[str, ET.Element] = {}
    namespaces: dict[str, str] = {}
    resolved: list[_ResolvedEdit] = []
    for edit in request.edits:
        if not _supported_part(graph, edit.part):
            return _report_refusal(request, "report_part_unsupported")
        if edit.part not in roots:
            try:
                source_parts[edit.part] = graph.read_part_bytes(edit.part)
                # The inspection above parsed these same bound bytes; its trees are
                # shared, so this function only reads them.
                root = baseline.parsed_xml.get(edit.part)
                roots[edit.part] = (
                    root if root is not None else ET.fromstring(source_parts[edit.part])
                )
            except PackageError:
                return _report_refusal(request, "report_part_missing")
            except ET.ParseError:
                return _report_refusal(request, "report_part_invalid")
            namespace = _word_namespace(roots[edit.part])
            if namespace is None:
                return _report_refusal(request, "source_word_namespace_unsupported")
            namespaces[edit.part] = namespace
        item, code = _resolve_edit(roots[edit.part], edit, namespaces[edit.part])
        if item is None:
            return _report_refusal(request, code or "report_target_invalid")
        resolved.append(item)
    if _has_overlapping_targets(resolved, roots):
        return _report_refusal(request, "report_edit_overlap")

    expected_texts: dict[int, str] = {}
    for item in resolved:
        expected_texts[id(item)] = (
            item.source_text[: item.start]
            + item.edit.text
            + item.source_text[item.end :]
        )
    for part, root in roots.items():
        text_items = [
            item
            for item in resolved
            if item.edit.part == part and item.edit.kind == "text"
        ]
        by_paragraph: dict[int, list[_ResolvedEdit]] = {}
        for item in text_items:
            by_paragraph.setdefault(id(item.element), []).append(item)
        for items in by_paragraph.values():
            source_text = items[0].source_text
            updated = source_text
            for item in sorted(items, key=lambda value: value.start, reverse=True):
                updated = updated[: item.start] + item.edit.text + updated[item.end :]
            for item in items:
                expected_texts[id(item)] = updated
    replacements = _replacement_texts(resolved)
    if replacements is None:
        return _report_refusal(request, "report_edit_overlap")
    for part, root in roots.items():
        part_replacements = {
            node_id: text
            for node_id, text in replacements.items()
            if any(
                id(region.node) == node_id
                for item in resolved
                if item.edit.part == part
                for region in item.regions
            )
        }
        if not _text_replacements_preserve_whitespace(root, part_replacements):
            return _report_refusal(request, "replacement_whitespace_unsupported")

    patched_parts: dict[str, bytes] = {}
    for part, root in roots.items():
        part_replacements = {
            node_id: text
            for node_id, text in replacements.items()
            if any(
                id(region.node) == node_id
                for item in resolved
                if item.edit.part == part
                for region in item.regions
            )
        }
        part_nodes = {
            id(region.node): region.node
            for item in resolved
            if item.edit.part == part
            for region in item.regions
        }
        patched = _patch_part(
            root, source_parts[part], namespaces[part], part_replacements, part_nodes
        )
        if patched is None:
            return _report_refusal(request, "report_text_patch_unsupported")
        patched_parts[part] = patched

    staged_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as staged:
            staged_path = Path(staged.name)
        _write_staged_package(source, staged_path, patched_parts)
        staged_bytes = staged_path.read_bytes()
        candidate = _inspect_bytes(staged_bytes, display_path=str(staged_path), reuse=baseline)
        target_change_reports = (
            _candidate_target_reports(candidate.parsed_xml, resolved, expected_texts)
            if candidate.valid
            else None
        )
        unselected_parts_preserved = _report_unselected_preserved(
            source_bytes, baseline, staged_bytes, candidate, set(patched_parts)
        )
        try:
            candidate_graph = PackageGraph.from_bytes(
                staged_bytes, path=staged_path, file_sha256=candidate.file_sha256
            )
            selected_parts_exact = all(
                candidate_graph.read_part_bytes(part) == patched
                for part, patched in patched_parts.items()
            )
        except PackageError:
            selected_parts_exact = False
        source_still_bound = _sha256_file(source) == request.source_sha256
        if not (
            candidate.valid
            and target_change_reports is not None
            and all(item["passed"] for item in target_change_reports)
            and unselected_parts_preserved
            and selected_parts_exact
            and source_still_bound
        ):
            return _report_refusal(
                request,
                "staged_verification_failed"
                if source_still_bound
                else "source_hash_changed",
            )
        os.replace(staged_path, destination)
        staged_path = None
        return ReportAuthoringReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            edit_targets=[edit.target_dict() for edit in request.edits],
            destination_sha256=candidate.file_sha256,
            target_change_reports=target_change_reports,
            unselected_parts_preserved=True,
        )
    finally:
        if staged_path is not None:
            staged_path.unlink(missing_ok=True)
