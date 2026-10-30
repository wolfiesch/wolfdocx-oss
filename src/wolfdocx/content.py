from __future__ import annotations

import hashlib
import re
import zipfile
from collections.abc import Callable, Mapping
from dataclasses import asdict, dataclass, field
from itertools import zip_longest
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast
from xml.etree import ElementTree as ET

from .inspect import (
    InspectionReport,
    PackageLimits,
    _parse_bounded_xml,
    _XmlSafetyError,
)

if TYPE_CHECKING:
    from .package import PackageGraph

ASSERTION_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
CONTENT_OPERATORS = {"contains", "not_contains", "equals", "count"}
SCOPE_KINDS = {"document", "paragraph", "table_cell"}
PART_TEXT_PATTERN = re.compile(
    r"^word/(?:header[1-9][0-9]*\.xml|footer[1-9][0-9]*\.xml|comments\.xml)$"
)
XML_ATTRIBUTE_PART_PATTERN = re.compile(
    r"^word/(?:document|styles|numbering|header[1-9][0-9]*|"
    r"footer[1-9][0-9]*|comments)\.xml$"
)
XML_LOCAL_NAME_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]{0,127}$")
SCOPE_KINDS.update({"part_text", "xml_attribute", "xml_text"})
MAX_CONTENT_ASSERTIONS = 100
MAX_TARGET_CHANGES = 100
MAX_ASSERTION_VALUE_CHARACTERS = 32_768
PROPOSABLE_TEXT_ELEMENT_NAMES = {"t", "instrText", "delText"}


class ContentPolicyError(ValueError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


@dataclass(frozen=True)
class ContentScope:
    kind: str
    index: int | None = None
    table: int | None = None
    row: int | None = None
    cell: int | None = None
    part: str | None = None
    element_path: tuple[str, ...] | None = None
    attribute: str | None = None
    occurrence: int | None = None

    @property
    def location(self) -> str:
        if self.kind == "paragraph":
            return f"word/document.xml#paragraph[{self.index}]"
        if self.kind == "table_cell":
            return f"word/document.xml#table[{self.table}]/row[{self.row}]/cell[{self.cell}]"
        if self.kind == "part_text":
            assert self.part is not None
            return self.part
        if self.kind in {"xml_attribute", "xml_text"}:
            assert self.part is not None
            assert self.element_path is not None
            location = f"{self.part}#{self.kind}[{'/'.join(self.element_path)}][{self.occurrence}]"
            return f"{location}/@{self.attribute}" if self.kind == "xml_attribute" else location
        return "word/document.xml"

    def to_dict(self) -> dict[str, Any]:
        payload = {key: value for key, value in asdict(self).items() if value is not None}
        if self.element_path is not None:
            payload["element_path"] = list(self.element_path)
        return payload


@dataclass(frozen=True)
class ContentAssertion:
    assertion_id: str
    scope: ContentScope
    operator: str
    value: str
    expected_count: int | None = None


@dataclass(frozen=True)
class ContentAssertionResult:
    assertion_id: str
    passed: bool
    scope: dict[str, Any]
    location: str
    operator: str
    value_sha256: str
    value_characters: int
    scope_found: bool
    scope_text_sha256: str | None
    scope_characters: int | None
    observed_count: int | None
    expected_count: int | None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ContentAssertionReport:
    results: list[ContentAssertionResult] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return all(result.passed for result in self.results)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "passed": self.passed,
            "assertion_count": len(self.results),
            "results": [result.to_dict() for result in self.results],
        }


@dataclass(frozen=True)
class TargetChangeResult:
    scope: dict[str, Any]
    location: str
    target_found: bool
    target_changed: bool
    baseline_value_sha256: str | None
    candidate_value_sha256: str | None
    remaining_part_equivalent: bool
    passed: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class TargetChangeReport:
    results: list[TargetChangeResult] = field(default_factory=list)
    uncovered_changed_parts: list[str] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return not self.uncovered_changed_parts and all(result.passed for result in self.results)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "passed": self.passed,
            "target_count": len(self.results),
            "uncovered_changed_parts": self.uncovered_changed_parts,
            "results": [result.to_dict() for result in self.results],
        }


@dataclass
class TargetChangeProposalReport:
    status: str
    complete: bool
    selectors: list[dict[str, Any]] = field(default_factory=list)
    changed_word_parts: list[str] = field(default_factory=list)
    changed_word_xml_parts: list[str] = field(default_factory=list)
    covered_changed_parts: list[str] = field(default_factory=list)
    unsupported_changed_parts: list[str] = field(default_factory=list)
    structurally_incompatible_parts: list[str] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)
    validation: TargetChangeReport | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "status": self.status,
            "complete": self.complete,
            "selector_count": len(self.selectors),
            "changed_word_parts": self.changed_word_parts,
            "changed_word_xml_parts": self.changed_word_xml_parts,
            "covered_changed_parts": self.covered_changed_parts,
            "unsupported_changed_parts": self.unsupported_changed_parts,
            "structurally_incompatible_parts": self.structurally_incompatible_parts,
            "issues": self.issues,
            "validation": self.validation.to_dict() if self.validation else None,
        }


def _unknown_fields(value: dict[str, Any], allowed: set[str], label: str) -> None:
    unknown = set(value) - allowed
    if unknown:
        raise ContentPolicyError(
            f"{label}_unknown_fields",
            f"Unknown {label} fields: {', '.join(sorted(unknown))}.",
        )


def _positive_index(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ContentPolicyError(
            "content_scope_index_invalid",
            f"Content assertion {label} must be a positive 1-based integer.",
        )
    return value


def _parse_scope(value: Any) -> ContentScope:
    if not isinstance(value, dict):
        raise ContentPolicyError(
            "content_scope_invalid",
            "Content assertion scope must be a JSON object.",
        )
    kind = value.get("kind")
    if not isinstance(kind, str) or kind not in SCOPE_KINDS:
        raise ContentPolicyError(
            "content_scope_kind_invalid",
            "Content assertion scope kind must be document, paragraph, table_cell, "
            "part_text, xml_attribute, or xml_text.",
        )
    if kind == "document":
        _unknown_fields(value, {"kind"}, "content_scope")
        return ContentScope(kind=kind)
    if kind == "paragraph":
        _unknown_fields(value, {"kind", "index"}, "content_scope")
        return ContentScope(kind=kind, index=_positive_index(value.get("index"), "index"))
    if kind == "part_text":
        _unknown_fields(value, {"kind", "part"}, "content_scope")
        part = value.get("part")
        if not isinstance(part, str) or PART_TEXT_PATTERN.fullmatch(part) is None:
            raise ContentPolicyError(
                "content_scope_part_invalid",
                "part_text scope must name a numbered Word header/footer or comments.xml.",
            )
        return ContentScope(kind=kind, part=part)
    if kind in {"xml_attribute", "xml_text"}:
        _unknown_fields(
            value,
            (
                {"kind", "part", "element_path", "attribute", "occurrence"}
                if kind == "xml_attribute"
                else {"kind", "part", "element_path", "occurrence"}
            ),
            "content_scope",
        )
        part = value.get("part")
        element_path = value.get("element_path")
        attribute = value.get("attribute")
        if not isinstance(part, str) or XML_ATTRIBUTE_PART_PATTERN.fullmatch(part) is None:
            raise ContentPolicyError(
                "content_scope_part_invalid",
                f"{kind} scope must name a supported Word XML part.",
            )
        if (
            not isinstance(element_path, list)
            or not 1 <= len(element_path) <= 6
            or any(
                not isinstance(item, str) or XML_LOCAL_NAME_PATTERN.fullmatch(item) is None
                for item in element_path
            )
        ):
            raise ContentPolicyError(
                "content_scope_element_path_invalid",
                f"{kind} element_path must contain 1-6 XML local names.",
            )
        if kind == "xml_attribute" and (
            not isinstance(attribute, str) or XML_LOCAL_NAME_PATTERN.fullmatch(attribute) is None
        ):
            raise ContentPolicyError(
                "content_scope_attribute_invalid",
                "xml_attribute attribute must be an XML local name.",
            )
        return ContentScope(
            kind=kind,
            part=part,
            element_path=tuple(element_path),
            attribute=attribute if kind == "xml_attribute" else None,
            occurrence=_positive_index(value.get("occurrence"), "occurrence"),
        )
    _unknown_fields(value, {"kind", "table", "row", "cell"}, "content_scope")
    return ContentScope(
        kind=kind,
        table=_positive_index(value.get("table"), "table"),
        row=_positive_index(value.get("row"), "row"),
        cell=_positive_index(value.get("cell"), "cell"),
    )


def parse_content_assertions(value: Any) -> list[ContentAssertion]:
    if not isinstance(value, list):
        raise ContentPolicyError(
            "content_assertions_invalid",
            "content_assertions must be an array.",
        )
    if len(value) > MAX_CONTENT_ASSERTIONS:
        raise ContentPolicyError(
            "content_assertion_limit",
            f"content_assertions cannot exceed {MAX_CONTENT_ASSERTIONS} entries.",
        )
    assertions: list[ContentAssertion] = []
    identifiers: set[str] = set()
    for position, item in enumerate(value, start=1):
        if not isinstance(item, dict):
            raise ContentPolicyError(
                "content_assertion_invalid",
                f"Content assertion {position} must be a JSON object.",
            )
        _unknown_fields(
            item,
            {"id", "scope", "operator", "value", "expected_count"},
            "content_assertion",
        )
        assertion_id = item.get("id")
        if (
            not isinstance(assertion_id, str)
            or ASSERTION_ID_PATTERN.fullmatch(assertion_id) is None
        ):
            raise ContentPolicyError(
                "content_assertion_id_invalid",
                "Content assertion id must be 1-64 portable identifier characters.",
            )
        if assertion_id in identifiers:
            raise ContentPolicyError(
                "content_assertion_id_duplicate",
                f"Content assertion id is duplicated: {assertion_id}.",
            )
        identifiers.add(assertion_id)
        operator = item.get("operator")
        if not isinstance(operator, str) or operator not in CONTENT_OPERATORS:
            raise ContentPolicyError(
                "content_assertion_operator_invalid",
                "Content assertion operator must be contains, not_contains, equals, or count.",
            )
        assertion_value = item.get("value")
        if not isinstance(assertion_value, str):
            raise ContentPolicyError(
                "content_assertion_value_invalid",
                "Content assertion value must be a string.",
            )
        if len(assertion_value) > MAX_ASSERTION_VALUE_CHARACTERS:
            raise ContentPolicyError(
                "content_assertion_value_limit",
                "Content assertion value exceeds the character limit.",
            )
        if operator != "equals" and not assertion_value:
            raise ContentPolicyError(
                "content_assertion_value_empty",
                "Only the equals operator can use an empty assertion value.",
            )
        expected_count = item.get("expected_count")
        if operator == "count":
            if (
                isinstance(expected_count, bool)
                or not isinstance(expected_count, int)
                or expected_count < 0
            ):
                raise ContentPolicyError(
                    "content_assertion_count_invalid",
                    "The count operator requires a non-negative expected_count.",
                )
        elif "expected_count" in item:
            raise ContentPolicyError(
                "content_assertion_count_unexpected",
                "expected_count is only valid with the count operator.",
            )
        assertions.append(
            ContentAssertion(
                assertion_id=assertion_id,
                scope=_parse_scope(item.get("scope")),
                operator=operator,
                value=assertion_value,
                expected_count=expected_count,
            )
        )
    return assertions


def parse_target_changes(value: Any) -> list[ContentScope]:
    if not isinstance(value, list):
        raise ContentPolicyError(
            "allowed_target_changes_invalid",
            "allowed_target_changes must be an array.",
        )
    if len(value) > MAX_TARGET_CHANGES:
        raise ContentPolicyError(
            "allowed_target_change_limit",
            f"allowed_target_changes cannot exceed {MAX_TARGET_CHANGES} entries.",
        )
    scopes: list[ContentScope] = []
    locations: set[str] = set()
    for position, item in enumerate(value, start=1):
        scope = _parse_scope(item)
        if scope.kind not in {"xml_attribute", "xml_text"}:
            raise ContentPolicyError(
                "allowed_target_change_scope_invalid",
                f"Allowed target change {position} must use xml_attribute or xml_text.",
            )
        if scope.location in locations:
            raise ContentPolicyError(
                "allowed_target_change_duplicate",
                f"Allowed target change is duplicated: {scope.location}.",
            )
        locations.add(scope.location)
        scopes.append(scope)
    return scopes


def _local_name(value: str) -> str:
    return value.rsplit("}", 1)[-1]


def _current_file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_xml_root(report: InspectionReport, part: str) -> ET.Element | None:
    path = Path(report.path)
    expected_part = next((item for item in report.parts if item.name == part), None)
    if expected_part is None or not path.is_file():
        return None
    try:
        if _current_file_sha256(path) != report.file_sha256:
            return None
        with zipfile.ZipFile(path) as archive:
            data = archive.read(part)
        if hashlib.sha256(data).hexdigest() != expected_part.sha256:
            return None
        limits = PackageLimits(**report.limits)
        root = _parse_bounded_xml(data, limits=limits, total_elements=[0], preserve_qnames=False)
        if _current_file_sha256(path) != report.file_sha256:
            return None
    except (
        OSError,
        KeyError,
        RuntimeError,
        ValueError,
        zipfile.BadZipFile,
        ET.ParseError,
        _XmlSafetyError,
    ):
        return None
    return root


def _load_xml_root_from_graph(graph: PackageGraph, part: str) -> ET.Element | None:
    from .package import PackageError

    package_part = graph.part(part)
    if package_part is None:
        return None
    try:
        return cast(ET.Element, package_part.parsed_xml())
    except PackageError:
        return None


def _matches_element_path(current: tuple[str, ...], element_path: tuple[str, ...]) -> bool:
    if current[-1] != element_path[-1]:
        return False
    position = 0
    for name in current:
        if name == element_path[position]:
            position += 1
            if position == len(element_path):
                return True
    return False


def _load_mutable_xml_root_from_graph(graph: PackageGraph, part: str) -> ET.Element | None:
    from .package import PackageError

    package_part = graph.part(part)
    if package_part is None:
        return None
    try:
        return package_part.parsed_xml().to_mutable_element()
    except PackageError:
        return None


def _matching_elements(root: ET.Element, scope: ContentScope) -> list[ET.Element]:
    """Elements whose root-to-self local-name path matches ``scope.element_path``.

    Equivalent to :func:`_matches_element_path` on each element's full path: the
    greedy subsequence position is carried down the walk instead of rebuilding and
    rescanning the path tuple at every element, and each distinct tag's local name
    is computed once.
    """
    assert scope.element_path is not None
    element_path = scope.element_path
    length = len(element_path)
    last = element_path[-1]
    matches: list[ET.Element] = []
    local_names: dict[object, str] = {}

    def visit(element: ET.Element, position: int) -> None:
        tag = element.tag
        name = local_names.get(tag)
        if name is None:
            name = local_names[tag] = _local_name(tag)
        if position < length and name == element_path[position]:
            position += 1
        if position == length and name == last:
            matches.append(element)
        for child in element:
            visit(child, position)

    visit(root, 0)
    return matches


def _selected_xml_value(root: ET.Element, scope: ContentScope) -> str | None:
    assert scope.occurrence is not None
    matches = _matching_elements(root, scope)
    if scope.occurrence > len(matches):
        return None
    element = matches[scope.occurrence - 1]
    if scope.kind == "xml_text":
        return element.text or ""
    assert scope.attribute is not None
    for name, value in element.attrib.items():
        if _local_name(name) == scope.attribute:
            return value
    return None


_MaskUndo = tuple[ET.Element, str | None, str | None]


def _mask_xml_scope(
    root: ET.Element, scope: ContentScope, undo: list[_MaskUndo]
) -> tuple[bool, str | None]:
    """Mask one scope's value in place, recording ``(element, attribute, original)``."""
    assert scope.occurrence is not None
    matches = _matching_elements(root, scope)
    if scope.occurrence > len(matches):
        return False, None
    element = matches[scope.occurrence - 1]
    if scope.kind == "xml_text":
        undo.append((element, None, element.text))
        value = element.text or ""
        element.text = "[wolfdocx allowed target]"
        return True, value
    assert scope.attribute is not None
    for name, value in element.attrib.items():
        if _local_name(name) == scope.attribute:
            undo.append((element, name, value))
            element.attrib[name] = "[wolfdocx allowed target]"
            return True, value
    return False, None


def _xml_trees_equal(left: ET.Element, right: ET.Element) -> bool:
    """Compare two trees node by node without materializing either.

    Equal tags, attribute mappings, text, tails and child counts along both pre-order
    walks prove the trees are identical, because a pre-order walk annotated with child
    counts determines the tree shape.
    """
    for left_node, right_node in zip_longest(left.iter(), right.iter()):
        if (
            left_node is None
            or right_node is None
            or left_node.tag != right_node.tag
            or left_node.text != right_node.text
            or left_node.tail != right_node.tail
            or len(left_node) != len(right_node)
            or left_node.attrib != right_node.attrib
        ):
            return False
    return True


def _value_sha256(value: str | None) -> str | None:
    return hashlib.sha256(value.encode("utf-8")).hexdigest() if value is not None else None


def _part_hashes(report: InspectionReport) -> dict[str, str]:
    return {part.name: part.sha256 for part in report.parts}


def _evaluate_target_changes(
    scopes: list[ContentScope],
    baseline_parts: Mapping[str, str],
    candidate_parts: Mapping[str, str],
    *,
    load_baseline_root: Callable[[str], ET.Element | None],
    load_candidate_root: Callable[[str], ET.Element | None],
) -> TargetChangeReport:
    """Verify each scope changed and nothing else in its part did.

    Scopes are masked in place for the comparison and restored before returning, so
    the loaders may return shared trees that nothing else mutates concurrently.
    """
    changed_word_xml_parts = {
        name
        for name in baseline_parts.keys() | candidate_parts.keys()
        if name.startswith("word/")
        and name.endswith(".xml")
        and baseline_parts.get(name) != candidate_parts.get(name)
    }
    declared_parts = {scope.part for scope in scopes if scope.part is not None}
    uncovered = sorted(changed_word_xml_parts - declared_parts)
    results: list[TargetChangeResult] = []
    for part in sorted(declared_parts):
        assert part is not None
        part_scopes = [scope for scope in scopes if scope.part == part]
        baseline_root = load_baseline_root(part)
        candidate_root = load_candidate_root(part)
        part_values: list[tuple[ContentScope, bool, str | None, bool, str | None]] = []
        if baseline_root is not None and candidate_root is not None:
            undo: list[_MaskUndo] = []
            try:
                for scope in part_scopes:
                    baseline_found, baseline_value = _mask_xml_scope(baseline_root, scope, undo)
                    candidate_found, candidate_value = _mask_xml_scope(candidate_root, scope, undo)
                    part_values.append(
                        (
                            scope,
                            baseline_found,
                            baseline_value,
                            candidate_found,
                            candidate_value,
                        )
                    )
                remaining_equivalent = _xml_trees_equal(baseline_root, candidate_root)
            finally:
                # Undo in reverse so a value masked twice regains its original.
                for element, attribute, original in reversed(undo):
                    if attribute is None:
                        element.text = original
                    else:
                        element.attrib[attribute] = original  # type: ignore[assignment]
        else:
            remaining_equivalent = False
            part_values = [(scope, False, None, False, None) for scope in part_scopes]
        for (
            scope,
            baseline_found,
            baseline_value,
            candidate_found,
            candidate_value,
        ) in part_values:
            found = baseline_found and candidate_found
            changed = found and baseline_value != candidate_value
            results.append(
                TargetChangeResult(
                    scope=scope.to_dict(),
                    location=scope.location,
                    target_found=found,
                    target_changed=changed,
                    baseline_value_sha256=_value_sha256(baseline_value),
                    candidate_value_sha256=_value_sha256(candidate_value),
                    remaining_part_equivalent=remaining_equivalent,
                    passed=found and changed and remaining_equivalent,
                )
            )
    return TargetChangeReport(results=results, uncovered_changed_parts=uncovered)


def evaluate_target_changes(
    scopes: list[ContentScope],
    baseline: InspectionReport,
    candidate: InspectionReport,
) -> TargetChangeReport:
    return _evaluate_target_changes(
        scopes,
        _part_hashes(baseline),
        _part_hashes(candidate),
        load_baseline_root=lambda part: _load_xml_root(baseline, part),
        load_candidate_root=lambda part: _load_xml_root(candidate, part),
    )


def _evaluate_target_changes_from_graphs(
    scopes: list[ContentScope],
    baseline: InspectionReport,
    candidate: InspectionReport,
    baseline_graph: PackageGraph,
    candidate_graph: PackageGraph,
) -> TargetChangeReport:
    return _evaluate_target_changes(
        scopes,
        _part_hashes(baseline),
        _part_hashes(candidate),
        load_baseline_root=lambda part: _load_mutable_xml_root_from_graph(baseline_graph, part),
        load_candidate_root=lambda part: _load_mutable_xml_root_from_graph(candidate_graph, part),
    )


def _flatten_xml(
    root: ET.Element,
) -> list[tuple[ET.Element, tuple[str, ...], tuple[str, ...]]]:
    elements: list[tuple[ET.Element, tuple[str, ...], tuple[str, ...]]] = []

    def visit(
        element: ET.Element,
        qualified_ancestors: tuple[str, ...],
        local_ancestors: tuple[str, ...],
    ) -> None:
        qualified_path = (*qualified_ancestors, element.tag)
        local_path = (*local_ancestors, _local_name(element.tag))
        elements.append((element, qualified_path, local_path))
        for child in element:
            visit(child, qualified_path, local_path)

    visit(root, (), ())
    return elements


def _proposal_occurrence(
    elements: list[tuple[ET.Element, tuple[str, ...], tuple[str, ...]]],
    position: int,
    element_path: tuple[str, ...],
) -> int:
    return sum(
        _matches_element_path(local_path, element_path)
        for _, _, local_path in elements[: position + 1]
    )


def _part_target_proposals(
    baseline_root: ET.Element,
    approved_root: ET.Element,
    *,
    part: str,
) -> list[ContentScope] | None:
    baseline_elements = _flatten_xml(baseline_root)
    approved_elements = _flatten_xml(approved_root)
    if len(baseline_elements) != len(approved_elements):
        return None
    proposals: list[ContentScope] = []
    for position, (baseline_item, approved_item) in enumerate(
        zip(baseline_elements, approved_elements, strict=True)
    ):
        baseline_element, baseline_qualified_path, baseline_local_path = baseline_item
        approved_element, approved_qualified_path, _ = approved_item
        if (
            baseline_qualified_path != approved_qualified_path
            or baseline_element.tail != approved_element.tail
            or len(baseline_element) != len(approved_element)
            or set(baseline_element.attrib) != set(approved_element.attrib)
        ):
            return None
        text_changed = baseline_element.text != approved_element.text
        changed_attributes = [
            attribute
            for attribute in sorted(baseline_element.attrib)
            if baseline_element.attrib[attribute] != approved_element.attrib[attribute]
        ]
        if not text_changed and not changed_attributes:
            continue
        element_path = baseline_local_path[-6:]
        occurrence = _proposal_occurrence(baseline_elements, position, element_path)
        if text_changed:
            if (baseline_element.text or "") == (approved_element.text or ""):
                return None
            if (
                len(baseline_element) != 0
                or _local_name(baseline_element.tag) not in PROPOSABLE_TEXT_ELEMENT_NAMES
            ):
                return None
            proposals.append(
                ContentScope(
                    kind="xml_text",
                    part=part,
                    element_path=element_path,
                    occurrence=occurrence,
                )
            )
        local_attributes = [_local_name(name) for name in baseline_element.attrib]
        if changed_attributes and len(local_attributes) != len(set(local_attributes)):
            return None
        for attribute in changed_attributes:
            proposals.append(
                ContentScope(
                    kind="xml_attribute",
                    part=part,
                    element_path=element_path,
                    attribute=_local_name(attribute),
                    occurrence=occurrence,
                )
            )
    return proposals


def propose_target_changes(
    baseline: InspectionReport,
    approved: InspectionReport,
) -> TargetChangeProposalReport:
    baseline_parts = {part.name: part.sha256 for part in baseline.parts}
    approved_parts = {part.name: part.sha256 for part in approved.parts}
    changed_word_parts = sorted(
        name
        for name in baseline_parts.keys() | approved_parts.keys()
        if name.startswith("word/") and baseline_parts.get(name) != approved_parts.get(name)
    )
    changed_word_xml_parts = [name for name in changed_word_parts if name.endswith(".xml")]
    unsupported = [
        part for part in changed_word_parts if XML_ATTRIBUTE_PART_PATTERN.fullmatch(part) is None
    ]
    structural: list[str] = []
    scopes: list[ContentScope] = []
    for part in changed_word_xml_parts:
        if part in unsupported:
            continue
        baseline_root = _load_xml_root(baseline, part)
        approved_root = _load_xml_root(approved, part)
        if baseline_root is None or approved_root is None:
            structural.append(part)
            continue
        part_proposals = _part_target_proposals(
            baseline_root,
            approved_root,
            part=part,
        )
        if part_proposals is None or not part_proposals:
            structural.append(part)
            continue
        scopes.extend(part_proposals)
    issues: list[str] = []
    if unsupported:
        issues.append("unsupported_changed_word_parts")
    if structural:
        issues.append("structurally_incompatible_word_xml_parts")
    if len(scopes) > MAX_TARGET_CHANGES:
        issues.append("allowed_target_change_limit")
        scopes = []
    validation = evaluate_target_changes(scopes, baseline, approved) if scopes else None
    complete = bool(scopes) and not issues and bool(validation and validation.passed)
    if validation is not None and not validation.passed:
        issues.append("target_proposal_validation_failed")
        complete = False
    if not changed_word_parts:
        status = "not_applicable"
    elif complete:
        status = "complete_review_required"
    else:
        status = "incomplete"
    selectors = [scope.to_dict() for scope in scopes]
    covered = sorted({scope.part for scope in scopes if scope.part is not None})
    return TargetChangeProposalReport(
        status=status,
        complete=complete,
        selectors=selectors,
        changed_word_parts=changed_word_parts,
        changed_word_xml_parts=changed_word_xml_parts,
        covered_changed_parts=covered,
        unsupported_changed_parts=unsupported,
        structurally_incompatible_parts=sorted(structural),
        issues=issues,
        validation=validation,
    )


def _scope_text(
    report: InspectionReport,
    scope: ContentScope,
    *,
    load_xml_root: Callable[[str], ET.Element | None],
) -> str | None:
    fingerprint = report.fingerprint
    if scope.kind == "document":
        return "\n".join(paragraph["text"] for paragraph in fingerprint.paragraphs)
    if scope.kind == "paragraph":
        assert scope.index is not None
        if scope.index > len(fingerprint.paragraphs):
            return None
        return fingerprint.paragraphs[scope.index - 1]["text"]
    if scope.kind == "part_text":
        assert scope.part is not None
        return report.part_texts.get(scope.part)
    if scope.kind in {"xml_attribute", "xml_text"}:
        assert scope.part is not None
        root = load_xml_root(scope.part)
        return _selected_xml_value(root, scope) if root is not None else None
    assert scope.table is not None
    assert scope.row is not None
    assert scope.cell is not None
    if scope.table > len(fingerprint.tables):
        return None
    table = fingerprint.tables[scope.table - 1]
    if scope.row > len(table["rows"]):
        return None
    row = table["rows"][scope.row - 1]
    if scope.cell > len(row["cells"]):
        return None
    return row["cells"][scope.cell - 1]["text"]


def _evaluate_content_assertions(
    assertions: list[ContentAssertion],
    candidate: InspectionReport,
    *,
    load_xml_root: Callable[[str], ET.Element | None],
) -> ContentAssertionReport:
    results: list[ContentAssertionResult] = []
    for assertion in assertions:
        text = _scope_text(
            candidate,
            assertion.scope,
            load_xml_root=load_xml_root,
        )
        observed_count: int | None = None
        passed = False
        if text is not None:
            if assertion.operator == "equals":
                observed_count = 1 if text == assertion.value else 0
                passed = text == assertion.value
            else:
                observed_count = text.count(assertion.value)
                if assertion.operator == "contains":
                    passed = observed_count > 0
                elif assertion.operator == "not_contains":
                    passed = observed_count == 0
                else:
                    passed = observed_count == assertion.expected_count
        results.append(
            ContentAssertionResult(
                assertion_id=assertion.assertion_id,
                passed=passed,
                scope=assertion.scope.to_dict(),
                location=assertion.scope.location,
                operator=assertion.operator,
                value_sha256=hashlib.sha256(assertion.value.encode("utf-8")).hexdigest(),
                value_characters=len(assertion.value),
                scope_found=text is not None,
                scope_text_sha256=(
                    hashlib.sha256(text.encode("utf-8")).hexdigest() if text is not None else None
                ),
                scope_characters=len(text) if text is not None else None,
                observed_count=observed_count,
                expected_count=assertion.expected_count,
            )
        )
    return ContentAssertionReport(results=results)


def evaluate_content_assertions(
    assertions: list[ContentAssertion],
    candidate: InspectionReport,
) -> ContentAssertionReport:
    return _evaluate_content_assertions(
        assertions,
        candidate,
        load_xml_root=lambda part: _load_xml_root(candidate, part),
    )


def _evaluate_content_assertions_from_graph(
    assertions: list[ContentAssertion],
    candidate: InspectionReport,
    graph: PackageGraph,
) -> ContentAssertionReport:
    return _evaluate_content_assertions(
        assertions,
        candidate,
        load_xml_root=lambda part: _load_xml_root_from_graph(graph, part),
    )
