"""Local add-in metadata and task panes, edited atomically in a Document session.

Import the value records from this module and use ``document.web_add_ins`` to
inspect or edit them. Collections are live, session-bound sequences; use add,
remove and clear to mutate them. Select a task pane by index with ``set`` when
several panes belong to an add-in. Save through the session's ordinary save API.
No add-in is activated and no store, reference or binding is fetched. Unknown
extension XML is retained; malformed identities and ambiguous unselected panes
are refused rather than guessed.
"""

from __future__ import annotations  # noqa: I001, RUF100

import math
import posixpath
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Generic, ParamSpec, Protocol, TypeVar, overload, runtime_checkable
from xml.etree import ElementTree as ET

from .inspect import _XmlSafetyError, serialize_xml_with_preserved_qnames
from .package import PackageError, PackageGraph, PackageRelationship
from .paragraph_authoring import _is_valid_xml_text

if TYPE_CHECKING:
    from .document import Document

_WE = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
_TP = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
_PANE_REL = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
_ADDIN_REL = "http://schemas.microsoft.com/office/2011/relationships/webextension"
_ADDIN_CT = "application/vnd.ms-office.webextension+xml"
_PANE_CT = "application/vnd.ms-office.webextensiontaskpanes+xml"
_DEFAULT_PANES = "word/webextensions/taskpanes.xml"


class WebAddInError(ValueError):
    """A refused add-in operation; the session remains unchanged."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


@dataclass(frozen=True, slots=True)
class WebAddInReference:
    id: str
    version: str
    store: str
    store_type: str


@dataclass(frozen=True, slots=True)
class WebAddInBinding:
    id: str
    type: str
    application_reference: str


@dataclass(frozen=True, slots=True)
class WebTaskPane:
    dock_state: str = "right"
    visible: bool = True
    width: float = 350.0
    row: int = 0
    locked: bool = False


_T = TypeVar("_T")
_P = ParamSpec("_P")


@runtime_checkable
class WebAddInCollection(Protocol[_T, _P]):
    """Common sequence/mutation contract, with each collection's typed add arguments."""

    def __len__(self) -> int: ...

    @overload
    def __getitem__(self, index: int) -> _T: ...

    @overload
    def __getitem__(self, index: slice) -> tuple[_T, ...]: ...

    def __iter__(self) -> Iterator[_T]: ...
    def add(self, *args: _P.args, **kwargs: _P.kwargs) -> _T: ...
    def remove(self, value: int | _T) -> None: ...
    def clear(self) -> None: ...


class _LiveCollection(Sequence[_T]):
    def _values(self) -> tuple[_T, ...]:
        raise NotImplementedError

    def __len__(self) -> int:
        return len(self._values())

    @overload
    def __getitem__(self, index: int) -> _T: ...

    @overload
    def __getitem__(self, index: slice) -> tuple[_T, ...]: ...

    def __getitem__(self, index: int | slice) -> _T | tuple[_T, ...]:
        return self._values()[index]

    def __iter__(self) -> Iterator[_T]:
        return iter(self._values())

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Sequence):
            return NotImplemented
        return self._values() == tuple(other)

    def _index(self, value: int | _T) -> int:
        values = self._values()
        if type(value) is int:
            if -len(values) <= value < len(values):
                return value % len(values)
        else:
            try:
                return values.index(value)
            except ValueError:
                pass
        raise WebAddInError("web_add_in_metadata_missing")


@dataclass(frozen=True, slots=True)
class WebAddIn:
    """Scalar identity record with live bound metadata collections."""

    part: str
    id: str
    reference: WebAddInReference
    properties: WebAddInProperties
    bindings: WebAddInBindings
    alternate_references: WebAddInReferences
    task_panes: WebTaskPanes


@dataclass(frozen=True, slots=True)
class _WebAddInValues:
    part: str
    id: str
    reference: WebAddInReference
    properties: tuple[tuple[str, str], ...]
    bindings: tuple[WebAddInBinding, ...]
    alternate_references: tuple[WebAddInReference, ...]
    task_panes: tuple[WebTaskPane, ...]


class _MetadataCollection(_LiveCollection[_T], Generic[_T]):
    def __init__(self, manager: WebAddIns, part: str, field: str) -> None:
        self._manager = manager
        self._part = part
        self._field = field

    def _values(self) -> tuple[_T, ...]:
        for record in self._manager._records():
            if record.part == self._part:
                return getattr(record, self._field)
        raise WebAddInError("web_add_in_missing")

    def remove(self, value: int | _T) -> None:
        self._manager._edit_metadata(self._part, self._field, index=self._index(value))

    def clear(self) -> None:
        self._manager._edit_metadata(self._part, self._field)


class WebAddInProperties(_MetadataCollection[tuple[str, str]]):
    def add(self, name: str, value: str) -> tuple[str, str]:
        self._manager.set_property(self._part, name, value)
        return name, value


class WebAddInBindings(_MetadataCollection[WebAddInBinding]):
    def add(self, value: WebAddInBinding) -> WebAddInBinding:
        self._manager.set_binding(self._part, value)
        return value


class WebAddInReferences(_MetadataCollection[WebAddInReference]):
    def add(self, value: WebAddInReference) -> WebAddInReference:
        self._manager._edit_metadata(self._part, self._field, append=_reference(value))
        return value


class WebTaskPanes(_MetadataCollection[WebTaskPane]):
    def add(self, value: WebTaskPane) -> WebTaskPane:
        self._manager._mutate_pane(self._part, value, append=True)
        return value

    def set(self, index: int, value: WebTaskPane) -> None:
        self._manager._mutate_pane(self._part, value, index=self._index(index))

    def remove(self, value: int | WebTaskPane) -> None:
        self._manager._remove_panes(self._part, index=self._index(value))

    def clear(self) -> None:
        self._manager._remove_panes(self._part)


def _text(value: str, *, empty: bool = False) -> str:
    if not isinstance(value, str) or (not value and not empty) or not _is_valid_xml_text(value):
        raise WebAddInError("web_add_in_value_invalid")
    return value


def _reference(value: WebAddInReference) -> ET.Element:
    if not isinstance(value, WebAddInReference):
        raise WebAddInError("web_add_in_reference_invalid")
    return ET.Element(
        f"{{{_WE}}}reference",
        {
            "id": _text(value.id),
            "version": _text(value.version),
            "store": _text(value.store),
            "storeType": _text(value.store_type),
        },
    )


def _binding(value: WebAddInBinding) -> ET.Element:
    if not isinstance(value, WebAddInBinding):
        raise WebAddInError("web_add_in_binding_invalid")
    return ET.Element(
        f"{{{_WE}}}binding",
        {
            "id": _text(value.id),
            "type": _text(value.type),
            "appref": _text(value.application_reference),
        },
    )


def _pane(value: WebTaskPane, relationship_id: str) -> ET.Element:
    if (
        not isinstance(value, WebTaskPane)
        or not isinstance(value.dock_state, str)
        or value.dock_state not in {"left", "right", "top", "bottom", "floating"}
        or type(value.visible) is not bool
        or type(value.locked) is not bool
        or type(value.row) is not int
        or value.row < 0
        or isinstance(value.width, bool)
        or not isinstance(value.width, (int, float))
        or not math.isfinite(value.width)
        or value.width <= 0
    ):
        raise WebAddInError("web_add_in_task_pane_invalid")
    root = ET.Element(
        f"{{{_TP}}}taskpane",
        {
            "dockstate": value.dock_state,
            "visibility": "1" if value.visible else "0",
            "width": format(value.width, ".17g"),
            "row": str(value.row),
            "locked": "1" if value.locked else "0",
        },
    )
    ET.SubElement(root, f"{{{_TP}}}webextensionref", {f"{{{_R}}}id": relationship_id})
    return root


def _preserved_parts(graph: PackageGraph) -> set[str]:
    """Retain opaque add-in graphs, including internal cached dependencies."""
    pending = [
        name for name, part in graph.parts.items() if part.content_type in {_ADDIN_CT, _PANE_CT}
    ]
    result: set[str] = set()
    while pending:
        name = pending.pop()
        if name in result:
            continue
        result.add(name)
        relationships = PackageGraph._rels_part_name(name)
        if relationships in graph.parts:
            result.add(relationships)
        pending.extend(
            target
            for edge in graph.relationships_for(name)
            if not edge.is_external and (target := edge.resolved_target()) in graph.parts
        )
    return result


def _fresh_id(graph: PackageGraph, source: str) -> str:
    occupied = {edge.id for edge in graph.relationships_for(source)}
    index = 1
    while f"rIdWebAddIn{index}" in occupied:
        index += 1
    return f"rIdWebAddIn{index}"


class WebAddIns(_LiveCollection[WebAddIn]):
    """Session-bound add-in metadata. All mutations are planned before adoption."""

    def __init__(self, document: Document) -> None:
        self._document = document

    def _values(self) -> tuple[WebAddIn, ...]:
        return self.items()

    def items(self) -> tuple[WebAddIn, ...]:
        return tuple(
            WebAddIn(
                record.part,
                record.id,
                record.reference,
                WebAddInProperties(self, record.part, "properties"),
                WebAddInBindings(self, record.part, "bindings"),
                WebAddInReferences(self, record.part, "alternate_references"),
                WebTaskPanes(self, record.part, "task_panes"),
            )
            for record in self._records()
        )

    def _graph(self, data: bytes | None = None) -> PackageGraph:
        try:
            return PackageGraph.from_bytes(
                self._document._materialize() if data is None else data,
                path=self._document.path,
                limits=self._document._graph.limits,
            )
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc

    def _root(self, graph: PackageGraph, part: str, tag: str) -> ET.Element:
        try:
            root = self._document._parse(graph.read_part_bytes(part), editable=True)
        except (PackageError, ET.ParseError, _XmlSafetyError, ValueError, KeyError) as exc:
            raise WebAddInError("web_add_in_xml_invalid") from exc
        if root.tag != tag:
            raise WebAddInError("web_add_in_xml_unsupported")
        return root

    def _collection(self, graph: PackageGraph) -> tuple[str, ET.Element] | None:
        edges = [edge for edge in graph.relationships if edge.type == _PANE_REL]
        if len(edges) > 1 or any(edge.is_external for edge in edges):
            raise WebAddInError("web_add_in_task_panes_unsupported")
        if not edges:
            if any(part.content_type == _PANE_CT for part in graph.parts.values()):
                raise WebAddInError("web_add_in_task_panes_unlinked")
            return None
        name = edges[0].resolved_target()
        if name not in graph.parts or graph.parts[name].content_type != _PANE_CT:
            raise WebAddInError("web_add_in_task_panes_invalid")
        return name, self._root(graph, name, f"{{{_TP}}}taskpanes")

    def _pane_records(self, graph: PackageGraph) -> dict[str, list[WebTaskPane]]:
        collection = self._collection(graph)
        result: dict[str, list[WebTaskPane]] = {}
        if collection is None:
            return result
        name, root = collection
        edges = {edge.id: edge for edge in graph.relationships_for(name)}
        for child in root:
            if child.tag != f"{{{_TP}}}taskpane":
                continue
            references = child.findall(f"{{{_TP}}}webextensionref")
            if len(references) != 1:
                raise WebAddInError("web_add_in_task_pane_invalid")
            identifier = references[0].get(f"{{{_R}}}id") or references[0].get(f"{{{_STRICT_R}}}id")
            edge = edges.get(identifier or "")
            if edge is None or edge.is_external or edge.type != _ADDIN_REL:
                raise WebAddInError("web_add_in_task_pane_relationship_invalid")
            target = edge.resolved_target()
            if target not in graph.parts or graph.parts[target].content_type != _ADDIN_CT:
                raise WebAddInError("web_add_in_task_pane_relationship_invalid")
            try:

                def boolean(attribute: str, default: str, node: ET.Element = child) -> bool:
                    value = node.get(attribute, default)
                    if value not in {"0", "1", "false", "true"}:
                        raise WebAddInError("web_add_in_task_pane_invalid")
                    return value in {"1", "true"}

                value = WebTaskPane(
                    child.get("dockstate", ""),
                    boolean("visibility", "0"),
                    float(child.get("width", "nan")),
                    int(child.get("row", "0")),
                    boolean("locked", "0"),
                )
                _pane(value, identifier or "")
            except (ValueError, OverflowError) as exc:
                raise WebAddInError("web_add_in_task_pane_invalid") from exc
            result.setdefault(target, []).append(value)
        return result

    @staticmethod
    def _one(root: ET.Element, name: str, *, create: bool = False) -> ET.Element:
        children = root.findall(f"{{{_WE}}}{name}")
        if len(children) == 1:
            return children[0]
        if not children and create:
            child = ET.Element(f"{{{_WE}}}{name}")
            if name == "alternateReferences":
                root.insert(list(root).index(WebAddIns._one(root, "reference")) + 1, child)
            else:
                root.append(child)
            return child
        raise WebAddInError("web_add_in_metadata_invalid")

    @staticmethod
    def _read_reference(node: ET.Element) -> WebAddInReference:
        value = WebAddInReference(
            *(node.get(name, "") for name in ("id", "version", "store", "storeType"))
        )
        _reference(value)
        return value

    def _records(self) -> tuple[_WebAddInValues, ...]:
        graph = self._graph()
        panes = self._pane_records(graph)
        records = []
        identifiers: set[str] = set()
        for name, part in graph.parts.items():
            if part.content_type != _ADDIN_CT:
                continue
            root = self._root(graph, name, f"{{{_WE}}}webextension")
            identifier = _text(root.get("id", ""))
            if identifier in identifiers:
                raise WebAddInError("web_add_in_id_duplicate")
            identifiers.add(identifier)
            reference = self._read_reference(self._one(root, "reference"))
            properties = []
            bindings = []
            for node in self._one(root, "properties"):
                if node.tag == f"{{{_WE}}}property":
                    properties.append(
                        (_text(node.get("name", "")), _text(node.get("value", ""), empty=True))
                    )
            for node in self._one(root, "bindings"):
                if node.tag == f"{{{_WE}}}binding":
                    value = WebAddInBinding(
                        *(node.get(key, "") for key in ("id", "type", "appref"))
                    )
                    _binding(value)
                    bindings.append(value)
            if len({key for key, _ in properties}) != len(properties) or len(
                {value.id for value in bindings}
            ) != len(bindings):
                raise WebAddInError("web_add_in_metadata_duplicate")
            alternates = root.findall(f"{{{_WE}}}alternateReferences")
            if len(alternates) > 1:
                raise WebAddInError("web_add_in_metadata_duplicate")
            records.append(
                _WebAddInValues(
                    name,
                    identifier,
                    reference,
                    tuple(properties),
                    tuple(bindings),
                    tuple(
                        self._read_reference(node)
                        for node in alternates[0]
                        if node.tag == f"{{{_WE}}}reference"
                    )
                    if alternates
                    else (),
                    tuple(panes.get(name, ())),
                )
            )
        return tuple(records)

    def _selected(self, graph: PackageGraph, part: str) -> ET.Element:
        if not any(value.part == part for value in self._records()):
            raise WebAddInError("web_add_in_missing")
        return self._root(graph, part, f"{{{_WE}}}webextension")

    def _adopt(self, data: bytes) -> None:
        try:
            self._graph(data).validate_relationship_graph()
            self._document._adopt("web-add-in", data, {})
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc

    def _replace(self, graph: PackageGraph, part: str, root: ET.Element) -> None:
        try:
            self._adopt(graph.replace_custom_part(part, serialize_xml_with_preserved_qnames(root)))
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc

    def set_reference(
        self,
        part: str,
        reference: WebAddInReference,
        *,
        alternate_references: Sequence[WebAddInReference] | None = None,
    ) -> None:
        replacement = _reference(reference)
        alternates = (
            None
            if alternate_references is None
            else [_reference(value) for value in alternate_references]
        )
        graph = self._graph()
        root = self._selected(graph, part)
        old = self._one(root, "reference")
        # Keep unknown attributes and children on the primary reference.
        old.attrib.update(replacement.attrib)
        if alternates is not None:
            container = self._one(root, "alternateReferences", create=True)
            for child in list(container):
                if child.tag == f"{{{_WE}}}reference":
                    container.remove(child)
            container.extend(alternates)
        self._replace(graph, part, root)

    def set_property(self, part: str, name: str, value: str) -> None:
        _text(name)
        _text(value, empty=True)
        graph = self._graph()
        root = self._selected(graph, part)
        container = self._one(root, "properties")
        old = next(
            (
                node
                for node in container
                if node.tag == f"{{{_WE}}}property" and node.get("name") == name
            ),
            None,
        )
        if old is None:
            old = ET.SubElement(container, f"{{{_WE}}}property", {"name": name})
        old.set("value", value)
        self._replace(graph, part, root)

    def remove_property(self, part: str, name: str) -> None:
        self._remove_metadata(part, "properties", "property", "name", name)

    def set_binding(self, part: str, binding: WebAddInBinding) -> None:
        replacement = _binding(binding)
        graph = self._graph()
        root = self._selected(graph, part)
        container = self._one(root, "bindings")
        old = next(
            (
                node
                for node in container
                if node.tag == replacement.tag and node.get("id") == binding.id
            ),
            None,
        )
        if old is None:
            container.append(replacement)
        else:
            old.attrib.update(replacement.attrib)
        self._replace(graph, part, root)

    def remove_binding(self, part: str, id: str) -> None:
        self._remove_metadata(part, "bindings", "binding", "id", id)

    def _remove_metadata(
        self, part: str, container_name: str, tag: str, key: str, value: str
    ) -> None:
        _text(value)
        graph = self._graph()
        root = self._selected(graph, part)
        container = self._one(root, container_name)
        old = next(
            (
                node
                for node in container
                if node.tag == f"{{{_WE}}}{tag}" and node.get(key) == value
            ),
            None,
        )
        if old is None:
            raise WebAddInError("web_add_in_metadata_missing")
        container.remove(old)
        self._replace(graph, part, root)

    def _edit_metadata(
        self, part: str, field: str, *, index: int | None = None, append: ET.Element | None = None
    ) -> None:
        container_name, tag = {
            "properties": ("properties", "property"),
            "bindings": ("bindings", "binding"),
            "alternate_references": ("alternateReferences", "reference"),
        }[field]
        graph = self._graph()
        root = self._selected(graph, part)
        container = self._one(root, container_name, create=True)
        nodes = container.findall(f"{{{_WE}}}{tag}")
        if append is not None:
            container.append(append)
        elif index is not None:
            container.remove(nodes[index])
        else:
            for node in nodes:
                container.remove(node)
        self._replace(graph, part, root)

    def _with_pane(
        self,
        graph: PackageGraph,
        part: str,
        value: WebTaskPane,
        *,
        index: int | None = None,
        append: bool = False,
    ) -> bytes:
        _pane(value, "check")
        collection = self._collection(graph)
        if collection is None:
            name = _DEFAULT_PANES
            index = 1
            while name in graph.parts:
                name = f"word/webextensions/taskpanes{index}.xml"
                index += 1
            data = graph.add_custom_part(
                name,
                serialize_xml_with_preserved_qnames(ET.Element(f"{{{_TP}}}taskpanes")),
                _PANE_CT,
                relationships=(PackageRelationship(_fresh_id(graph, ""), _PANE_REL, name),),
            )
            graph = self._graph(data)
            collection = self._collection(graph)
            assert collection is not None
        name, root = collection
        edges = list(graph.relationships_for(name))
        associated = [
            edge
            for edge in edges
            if not edge.is_external and edge.type == _ADDIN_REL and edge.resolved_target() == part
        ]
        panes = []
        for child in root.findall(f"{{{_TP}}}taskpane"):
            reference = child.find(f"{{{_TP}}}webextensionref")
            identifier = (
                None
                if reference is None
                else reference.get(f"{{{_R}}}id") or reference.get(f"{{{_STRICT_R}}}id")
            )
            if identifier in {edge.id for edge in associated}:
                panes.append(child)
        if not append and index is None and len(panes) > 1:
            raise WebAddInError("web_add_in_task_pane_ambiguous")
        old = None if append or not panes else panes[0 if index is None else index]
        if old is None:
            identifier = associated[0].id if associated else _fresh_id(graph, name)
        else:
            reference = old.find(f"{{{_TP}}}webextensionref")
            assert reference is not None
            identifier = reference.get(f"{{{_R}}}id") or reference.get(f"{{{_STRICT_R}}}id") or ""
        replacement = _pane(value, identifier)
        if old is not None:
            old.attrib.update(replacement.attrib)
        else:
            extension = root.find(f"{{{_TP}}}extLst")
            root.insert(
                list(root).index(extension) if extension is not None else len(root), replacement
            )
        if not associated:
            edges.append(
                PackageRelationship(
                    identifier,
                    _ADDIN_REL,
                    posixpath.relpath(part, posixpath.dirname(name)),
                    source_part=name,
                )
            )
        return graph.replace_custom_part(
            name, serialize_xml_with_preserved_qnames(root), relationships=edges
        )

    def set_task_pane(self, part: str, pane: WebTaskPane) -> None:
        self._mutate_pane(part, pane)

    def _mutate_pane(
        self, part: str, pane: WebTaskPane, *, index: int | None = None, append: bool = False
    ) -> None:
        graph = self._graph()
        self._selected(graph, part)
        try:
            self._adopt(self._with_pane(graph, part, pane, index=index, append=append))
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc

    def _remove_panes(self, part: str, *, index: int | None = None) -> None:
        graph = self._graph()
        self._selected(graph, part)
        collection = self._collection(graph)
        if collection is None:
            return
        name, root = collection
        edges = list(graph.relationships_for(name))
        associated = {
            edge.id
            for edge in edges
            if not edge.is_external and edge.type == _ADDIN_REL and edge.resolved_target() == part
        }
        nodes = []
        for node in root.findall(f"{{{_TP}}}taskpane"):
            reference = node.find(f"{{{_TP}}}webextensionref")
            assert reference is not None
            identifier = reference.get(f"{{{_R}}}id") or reference.get(f"{{{_STRICT_R}}}id")
            if identifier in associated:
                nodes.append(node)
        for node in nodes if index is None else [nodes[index]]:
            root.remove(node)
        used = {
            node.get(f"{{{_R}}}id") or node.get(f"{{{_STRICT_R}}}id")
            for node in root.iter(f"{{{_TP}}}webextensionref")
        }
        edges = [edge for edge in edges if edge.id not in associated or edge.id in used]
        try:
            if (
                not any(isinstance(child.tag, str) for child in root)
                and not root.attrib
                and not (root.text or "").strip()
                and not edges
            ):
                data = graph.remove_custom_part(name)
            else:
                data = graph.replace_custom_part(
                    name, serialize_xml_with_preserved_qnames(root), relationships=edges
                )
            self._adopt(data)
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc

    def add(
        self,
        id: str,
        reference: WebAddInReference,
        *,
        properties: Mapping[str, str] | None = None,
        bindings: Sequence[WebAddInBinding] = (),
        alternate_references: Sequence[WebAddInReference] = (),
        task_pane: WebTaskPane | None = None,
    ) -> WebAddIn:
        _text(id)
        if properties is not None and not isinstance(properties, Mapping):
            raise WebAddInError("web_add_in_value_invalid")
        if not isinstance(bindings, Sequence) or not isinstance(alternate_references, Sequence):
            raise WebAddInError("web_add_in_value_invalid")
        root = ET.Element(f"{{{_WE}}}webextension", {"id": id})
        root.append(_reference(reference))
        alternates = ET.SubElement(root, f"{{{_WE}}}alternateReferences")
        alternates.extend(_reference(value) for value in alternate_references)
        bag = ET.SubElement(root, f"{{{_WE}}}properties")
        for name, value in (properties or {}).items():
            ET.SubElement(
                bag, f"{{{_WE}}}property", {"name": _text(name), "value": _text(value, empty=True)}
            )
        binding_nodes = [_binding(value) for value in bindings]
        if len({node.get("id") for node in binding_nodes}) != len(binding_nodes):
            raise WebAddInError("web_add_in_metadata_duplicate")
        ET.SubElement(root, f"{{{_WE}}}bindings").extend(binding_nodes)
        if task_pane is not None:
            _pane(task_pane, "check")
        existing = self.items()
        if any(value.id == id for value in existing):
            raise WebAddInError("web_add_in_id_duplicate")
        graph = self._graph()
        index = 1
        part = f"word/webextensions/webextension{index}.xml"
        while part in graph.parts:
            index += 1
            part = f"word/webextensions/webextension{index}.xml"
        try:
            data = graph.add_custom_part(part, serialize_xml_with_preserved_qnames(root), _ADDIN_CT)
            if task_pane is not None:
                data = self._with_pane(self._graph(data), part, task_pane)
            self._adopt(data)
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc
        return next(value for value in self.items() if value.part == part)

    def remove(self, part: str | int | WebAddIn) -> None:
        if isinstance(part, WebAddIn):
            part = part.part
        elif type(part) is int:
            part = self[self._index(part)].part
        elif not isinstance(part, str):
            raise WebAddInError("web_add_in_value_invalid")
        graph = self._graph()
        self._selected(graph, part)
        self._check_linked_controls(graph)
        self._adopt(self._without_addin(graph, part))

    def clear(self) -> None:
        records = self._records()
        if not records:
            return
        graph = self._graph()
        self._check_linked_controls(graph)
        for record in records:
            data = self._without_addin(graph, record.part)
            graph = self._graph(data)
        self._adopt(data)

    def _check_linked_controls(self, graph: PackageGraph) -> None:
        # Linked controls carry identities outside this metadata model. Never leave
        # them dangling or silently rewrite their body content.
        for name, descriptor in graph.parts.items():
            if descriptor.content_type.startswith(
                "application/vnd.openxmlformats-officedocument.wordprocessingml."
            ) and name.endswith(".xml"):
                try:
                    root = self._document._parse(graph.read_part_bytes(name), editable=False)
                except (ET.ParseError, _XmlSafetyError) as exc:
                    raise WebAddInError("web_add_in_xml_invalid") from exc
                if any(
                    node.tag.rpartition("}")[2] in {"webExtensionLinked", "webExtensionCreated"}
                    for node in root.iter()
                ):
                    raise WebAddInError("web_add_in_linked_control_unsupported")

    def _without_addin(self, graph: PackageGraph, part: str) -> bytes:
        collection = self._collection(graph)
        try:
            if collection is not None:
                name, root = collection
                associated = {
                    edge.id
                    for edge in graph.relationships_for(name)
                    if not edge.is_external
                    and edge.type == _ADDIN_REL
                    and edge.resolved_target() == part
                }
                for node in list(root):
                    reference = node.find(f"{{{_TP}}}webextensionref")
                    rid = (
                        None
                        if reference is None
                        else reference.get(f"{{{_R}}}id") or reference.get(f"{{{_STRICT_R}}}id")
                    )
                    if node.tag == f"{{{_TP}}}taskpane" and rid in associated:
                        root.remove(node)
                graph = self._graph(
                    graph.replace_custom_part(name, serialize_xml_with_preserved_qnames(root))
                )
            data = graph.remove_custom_part(part)
            if (
                collection is not None
                and not any(isinstance(child.tag, str) for child in root)
                and not root.attrib
                and not (root.text or "").strip()
            ):
                data = self._graph(data).remove_custom_part(name)
            return data
        except PackageError as exc:
            raise WebAddInError("web_add_in_package_invalid") from exc
