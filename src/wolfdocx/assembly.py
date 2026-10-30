"""Package editing helpers shared by authoring and session edits.

Relationship naming and allocation, content-type overrides and main-document
relationships keep package edits consistent with the Open Packaging Conventions.
Identity remapping keeps bookmarks, content controls, paragraph IDs, drawing
nonvisual IDs, VML shape names and Word anchor IDs unique within a destination
part when content moves between parts or packages.
"""

from __future__ import annotations

import posixpath
import re
from collections.abc import Mapping
from xml.etree import ElementTree as ET

from ._comment_threads import generate_collision_safe_para_id
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_DOCUMENT = "word/document.xml"
_PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_VML = "urn:schemas-microsoft-com:vml"
_VML_OFFICE = "urn:schemas-microsoft-com:office:office"
_VML_SPID = f"{{{_VML_OFFICE}}}spid"
_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
# Word anchor identities per part: anchor ids are unique across VML (w14) and DrawingML
# (wp14) anchors; DrawingML edit ids form a second domain.
_ANCHOR_ID_DOMAINS = {
    f"{{{_W14}}}anchorId": "anchorId",
    f"{{{_WP14}}}anchorId": "anchorId",
    f"{{{_WP14}}}editId": "editId",
}

# Drawing IDs are unsigned integers, distinct from relationship IDs, paragraph
# IDs, SmartArt modelId values and VML shape names. The outer Word drawing and
# the inner DrawingML objects have separate nonvisual identity domains.
_DRAWING_ID_DOMAINS = {
    **{
        f"{{{namespace}}}docPr": "docPr"
        for namespace in (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
        )
    },
    **{
        f"{{{namespace}}}cNvPr": "cNvPr"
        for namespace in (
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "http://purl.oclc.org/ooxml/drawingml/picture",
            "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
            "http://purl.oclc.org/ooxml/drawingml/chartDrawing",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "http://schemas.microsoft.com/office/drawing/2008/diagram",
        )
    },
}
_DRAWING_CONNECTIONS = {
    f"{{{namespace}}}{name}"
    for namespace in (
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
    )
    for name in ("stCxn", "endCxn")
}
_WORD_DRAWINGS = {
    f"{{{namespace}}}drawing"
    for namespace in (_W, "http://purl.oclc.org/ooxml/wordprocessingml/main")
}
_MAX_DRAWING_ID = 0xFFFFFFFF
_MAX_SDT_ID = 2_147_483_647


def _q(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _word_ns(root: ET.Element) -> str | None:
    return (
        root.tag[1:].split("}", 1)[0]
        if root.tag.startswith("{") and root.tag[1:].split("}", 1)[0] in SUPPORTED_WORD_NAMESPACES
        else None
    )


def _rels_name(part: str) -> str:
    parent, name = posixpath.split(part)
    return posixpath.join(parent, "_rels", name + ".rels")


def _resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _relative(part: str, target: str) -> str:
    return posixpath.relpath(target, posixpath.dirname(part))


def _read_rels(parts: Mapping[str, bytes], owner: str) -> ET.Element:
    name = _rels_name(owner)
    return (
        parse_xml_with_preserved_qnames(parts[name])
        if name in parts
        else ET.Element(_q(_PR, "Relationships"))
    )


def _ids(root: ET.Element) -> set[str]:
    return {node.get("Id") for node in root.findall(_q(_PR, "Relationship")) if node.get("Id")}


def _new_rel_id(root: ET.Element) -> str:
    used = _ids(root)
    serial = 1
    while f"rId{serial}" in used:
        serial += 1
    return f"rId{serial}"


def _ensure_override(destination: dict[str, bytes], part: str, content_type: str) -> None:
    types = parse_xml_with_preserved_qnames(destination["[Content_Types].xml"])
    if not any(
        node.get("PartName") == "/" + part and node.get("ContentType") == content_type
        for node in types.findall(_q(_CT, "Override"))
    ):
        types.append(
            ET.Element(
                _q(_CT, "Override"),
                {"PartName": "/" + part, "ContentType": content_type},
            )
        )
    destination["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types)


def _add_document_relationship(
    destination: dict[str, bytes], relationship_type: str, target: str
) -> None:
    rels_name = _rels_name(_DOCUMENT)
    rels = _read_rels(destination, _DOCUMENT)
    rels.append(
        ET.Element(
            _q(_PR, "Relationship"),
            {
                "Id": _new_rel_id(rels),
                "Type": relationship_type,
                "Target": _relative(_DOCUMENT, target),
            },
        )
    )
    destination[rels_name] = serialize_xml_with_preserved_qnames(rels)


def _remap_bookmarks(
    destination: Mapping[str, bytes], roots: list[ET.Element], ns: str
) -> dict[str, str]:
    destination_roots: list[ET.Element] = []
    for part, data in destination.items():
        if not part.startswith("word/") or not part.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        if _word_ns(root) == ns:
            destination_roots.append(root)
    used_ids = {
        int(node.get(Q(ns, "id")))
        for root in destination_roots
        for node in root.iter(Q(ns, "bookmarkStart"))
        if (node.get(Q(ns, "id")) or "").isdigit()
    }
    used_names = {
        node.get(Q(ns, "name"))
        for root in destination_roots
        for node in root.iter(Q(ns, "bookmarkStart"))
        if node.get(Q(ns, "name"))
    }
    next_id = max(used_ids, default=0) + 1
    names: dict[str, str] = {}
    ids: dict[str, str] = {}
    for root in roots:
        for node in root.iter(Q(ns, "bookmarkStart")):
            old_id, old_name = node.get(Q(ns, "id")), node.get(Q(ns, "name"))
            if old_id is None or old_name is None or old_id in ids or old_name in names:
                raise ValueError("assembly_bookmark_invalid")
            new_id = str(next_id)
            next_id += 1
            base = old_name[:32] or "Bookmark"
            new_name = base
            serial = 2
            while new_name in used_names:
                new_name = f"{base}_{serial}"
                serial += 1
            used_names.add(new_name)
            ids[old_id] = new_id
            names[old_name] = new_name
            node.set(Q(ns, "id"), new_id)
            node.set(Q(ns, "name"), new_name)
    for root in roots:
        for node in root.iter(Q(ns, "bookmarkEnd")):
            old = node.get(Q(ns, "id"))
            if old not in ids:
                raise ValueError("assembly_bookmark_end_invalid")
            node.set(Q(ns, "id"), ids[old])
        for node in root.iter(Q(ns, "hyperlink")):
            anchor = node.get(Q(ns, "anchor"))
            if anchor in names:
                node.set(Q(ns, "anchor"), names[anchor])
    return names


def _allocate_sdt_id(used_ids: set[int]) -> int | None:
    """Deterministically allocate the lowest unused positive integer SDT ID."""
    candidate = 1
    while candidate in used_ids:
        candidate += 1
    return candidate if candidate <= _MAX_SDT_ID else None


def _remap_sdt_ids(roots: list[ET.Element], ns: str, occupied: set[int]) -> None:
    for root in roots:
        for control in root.iter(Q(ns, "sdt")):
            identifiers = control.findall(f"{Q(ns, 'sdtPr')}/{Q(ns, 'id')}")
            if len(identifiers) != 1:
                raise ValueError("assembly_sdt_identity_unsupported")
            node = identifiers[0]
            value = node.get(Q(ns, "val"), "")
            if not value.isdigit() or int(value) > _MAX_SDT_ID:
                raise ValueError("assembly_sdt_identity_unsupported")
            identifier = int(value)
            if identifier in occupied:
                identifier = _allocate_sdt_id(occupied)
                if identifier is None:
                    raise ValueError("assembly_sdt_identity_exhausted")
                node.set(Q(ns, "val"), str(identifier))
            occupied.add(identifier)


def _remap_paragraph_ids(roots: list[ET.Element], ns: str, occupied: set[str]) -> None:
    """Reserve identities within one destination part, changing only collisions."""
    for root in roots:
        for paragraph in root.iter(Q(ns, "p")):
            identifier = paragraph.get(_PARA_ID)
            if identifier is None:
                continue
            if identifier.upper() in occupied:
                paragraph.set(_PARA_ID, generate_collision_safe_para_id(occupied, identifier))
            else:
                occupied.add(identifier.upper())


def _drawing_identifier(node: ET.Element) -> int | None:
    value = node.get("id")
    if value is None:
        return None
    if not re.fullmatch(r"\+?[0-9]+", value.strip()):
        raise ValueError("assembly_drawing_identity_invalid")
    identifier = int(value)
    if identifier > _MAX_DRAWING_ID:
        raise ValueError("assembly_drawing_identity_invalid")
    return identifier


def _drawing_ids(root: ET.Element) -> dict[str, set[int]]:
    occupied: dict[str, set[int]] = {"docPr": set(), "cNvPr": set()}
    for node in root.iter():
        domain = _DRAWING_ID_DOMAINS.get(node.tag)
        if domain is not None and (identifier := _drawing_identifier(node)) is not None:
            occupied[domain].add(identifier)
    return occupied


def _remap_drawing_ids(roots: list[ET.Element], occupied: dict[str, set[int]]) -> None:
    """Allocate only imported collisions within one destination package part.

    Reserve all incoming values before allocation so a collision cannot steal a
    later noncolliding identity. Resolve connector endpoints in their original
    drawing graph before changing anything; repeated imported canvases may have
    identical source IDs without sharing their endpoint bindings.

    Related parts copied alongside imported content retain their own identity
    scope and bytes. Only content merged into existing parts needs rebasing.
    """
    definitions: list[tuple[ET.Element, ET.Element, str, int]] = []
    connections: list[tuple[ET.Element, ET.Element, int]] = []
    local: dict[tuple[ET.Element, str, int], list[ET.Element]] = {}
    imported: dict[tuple[str, int], list[ET.Element]] = {}
    reserved = {domain: set(values) for domain, values in occupied.items()}
    for root in roots:
        pending = [(root, root)]
        while pending:
            node, scope = pending.pop()
            if node.tag in _WORD_DRAWINGS:
                scope = node
            domain = _DRAWING_ID_DOMAINS.get(node.tag)
            if domain is not None:
                identifier = _drawing_identifier(node)
                if identifier is not None:
                    definitions.append((node, scope, domain, identifier))
                    local.setdefault((scope, domain, identifier), []).append(node)
                    imported.setdefault((domain, identifier), []).append(node)
                    reserved[domain].add(identifier)
            elif node.tag in _DRAWING_CONNECTIONS:
                identifier = _drawing_identifier(node)
                if identifier is not None:
                    connections.append((node, scope, identifier))
            pending.extend((child, scope) for child in reversed(node))

    bindings: list[tuple[ET.Element, ET.Element, int]] = []
    for connection, scope, identifier in connections:
        # Inner canvas objects take precedence over their outer Word wrapper.
        targets = (
            local.get((scope, "cNvPr", identifier))
            or local.get((scope, "docPr", identifier))
            or imported.get(("cNvPr", identifier))
            or imported.get(("docPr", identifier))
        )
        if not targets:
            raise ValueError("assembly_drawing_reference_missing")
        if len(targets) != 1:
            raise ValueError("assembly_drawing_reference_ambiguous")
        bindings.append((connection, targets[0], identifier))

    next_ids = {domain: 1 for domain in occupied}
    for node, _, domain, identifier in definitions:
        if identifier in occupied[domain]:
            candidate = next_ids[domain]
            while candidate in reserved[domain]:
                candidate += 1
            if candidate > _MAX_DRAWING_ID:
                raise ValueError("assembly_drawing_identity_exhausted")
            node.set("id", str(candidate))
            reserved[domain].add(candidate)
            next_ids[domain] = candidate + 1
            occupied[domain].add(candidate)
        else:
            occupied[domain].add(identifier)
    for connection, target, identifier in bindings:
        if int(target.attrib["id"]) != identifier:
            connection.set("id", target.attrib["id"])


def _vml_identity_attributes(node: ET.Element) -> tuple[str, ...]:
    if not node.tag.startswith(f"{{{_VML}}}"):
        return ()
    return tuple(name for name in ("id", _VML_SPID) if node.get(name))


def _vml_ids(root: ET.Element) -> set[str]:
    return {node.attrib[name] for node in root.iter() for name in _vml_identity_attributes(node)}


def _fresh_vml_id(value: str, reserved: set[str]) -> str:
    serial_id = re.fullmatch(r"(_x0000_[si])([0-9]+)", value)
    if serial_id is not None:
        # Word derives shape numbers from these spellings; keep them numeric.
        prefix, serial = serial_id.group(1), int(serial_id.group(2))
        while f"{prefix}{serial}" in reserved:
            serial += 1
        return f"{prefix}{serial}"
    serial = 2
    while f"{value}_{serial}" in reserved:
        serial += 1
    return f"{value}_{serial}"


def _remap_vml_ids(roots: list[ET.Element], ns: str, occupied: set[str]) -> None:
    """Rename imported VML shape identities that collide within one destination part.

    VML ``id`` and ``o:spid`` values share one lexical domain. A renamed value
    changes at every imported definition and reference, so an id equal to its spid,
    shapetype uses, connector rules and OLE or ActiveX shape links stay bound.
    """
    incoming = {
        node.attrib[name]
        for root in roots
        for node in root.iter()
        for name in _vml_identity_attributes(node)
    }
    reserved = occupied | incoming
    renamed: dict[str, str] = {}
    for root in roots:
        for node in root.iter():
            for name in _vml_identity_attributes(node):
                value = node.attrib[name]
                if value in occupied and value not in renamed:
                    renamed[value] = _fresh_vml_id(value, reserved)
                    reserved.add(renamed[value])
    occupied.update(incoming - renamed.keys(), renamed.values())
    if not renamed:
        return
    for root in roots:
        for node in root.iter():
            for name in _vml_identity_attributes(node):
                node.set(name, renamed.get(node.attrib[name], node.attrib[name]))
            if node.tag.startswith(f"{{{_VML}}}"):
                references = ("type",)
            elif node.tag == _q(_VML_OFFICE, "OLEObject"):
                references = ("ShapeID",)
            elif node.tag.startswith(f"{{{_VML_OFFICE}}}"):
                references = ("idref",)
            elif node.tag == Q(ns, "control"):
                references = (Q(ns, "shapeid"),)
            elif node.tag.endswith("}compatExt") and "spid" in node.attrib:
                references = ("spid",)
            else:
                continue
            for name in references:
                value = node.get(name)
                if value is None:
                    continue
                fragment = value.startswith("#")
                if fragment != (name in {"type", "idref"}):
                    continue
                target = renamed.get(value.removeprefix("#"))
                if target is not None:
                    node.set(name, f"#{target}" if fragment else target)


def _anchor_ids(root: ET.Element) -> dict[str, set[str]]:
    """Every Word anchor identity in one part, by domain (anchorId, editId)."""
    occupied: dict[str, set[str]] = {"anchorId": set(), "editId": set()}
    for node in root.iter():
        for name, value in node.attrib.items():
            if (domain := _ANCHOR_ID_DOMAINS.get(name)) is not None:
                occupied[domain].add(value.upper())
    return occupied


def _remap_anchor_ids(roots: list[ET.Element], occupied: dict[str, set[str]]) -> None:
    """Give imported anchors whose w14/wp14 anchorId or wp14 editId collides a fresh id.

    Word keeps these ids unique within a part (VML and DrawingML anchors share one
    anchorId domain), and no other markup refers to them, so only the colliding
    attribute changes. Reserve every incoming id first so a reallocation cannot take
    a later imported value.
    """
    reserved = {domain: set(values) for domain, values in occupied.items()}
    for root in roots:
        for domain, values in _anchor_ids(root).items():
            reserved[domain] |= values
    for root in roots:
        for node in root.iter():
            for name, domain in _ANCHOR_ID_DOMAINS.items():
                if (value := node.get(name)) is None:
                    continue
                value = value.upper()
                if value in occupied[domain]:
                    value = generate_collision_safe_para_id(reserved[domain], value)
                    node.set(name, value)
                occupied[domain].add(value)
