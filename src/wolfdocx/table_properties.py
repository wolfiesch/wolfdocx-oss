"""Source-bound table properties authoring and discovery for WordprocessingML documents."""

from __future__ import annotations

import copy
import enum
import json
import posixpath
import re
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .authoring import (
    _PUBLISHED,
    _REFUSED,
    _cell_structure_fingerprint,
    _require_sha256,
    _sha256_file,
    _unsupported_context_by_element,
    _word_namespace,
)
from .inspect import (
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import _DOCUMENT, ReviewReport, _publish

__all__ = [
    "BorderSpec",
    "TableBordersSpec",
    "TableCellMarginsSpec",
    "TableCellSpacingSpec",
    "TableFloatingPositioningSpec",
    "TableIndentSpec",
    "TableLayoutSpec",
    "TableLookSpec",
    "TablePropertiesReport",
    "TablePropertiesRequest",
    "TablePropertiesSpec",
    "TablePropertiesTarget",
    "TableShadingSpec",
    "TableWidthSpec",
    "author_table_properties",
    "discover_table_properties_targets",
    "load_table_properties_request",
    "parse_table_properties_request",
    "table_properties_structure_fingerprint",
]

_HEX_COLOR_PATTERN = re.compile(r"^[0-9a-fA-F]{6}\Z")
_VALID_LAYOUT_TYPES = frozenset({"fixed", "autofit"})
_LOOK_HEX_PATTERN = re.compile(r"^[0-9a-fA-F]{4}\Z")

_VALID_FLOATING_HORIZONTAL_ANCHORS = frozenset({"margin", "page", "text"})
_VALID_FLOATING_VERTICAL_ANCHORS = frozenset({"margin", "page", "text"})
_VALID_FLOATING_HORIZONTAL_ALIGNMENTS = frozenset({"center", "inside", "left", "outside", "right"})
_VALID_FLOATING_VERTICAL_ALIGNMENTS = frozenset(
    {"bottom", "center", "inline", "inside", "outside", "top"}
)
_FLOATING_XML_ATTR_TO_KEY = {
    "horzAnchor": "horizontal_anchor",
    "vertAnchor": "vertical_anchor",
    "tblpX": "horizontal_offset",
    "tblpY": "vertical_offset",
    "leftFromText": "left_from_text",
    "rightFromText": "right_from_text",
    "topFromText": "top_from_text",
    "bottomFromText": "bottom_from_text",
}
_FLOATING_REQUIRED_KEYS = frozenset(_FLOATING_XML_ATTR_TO_KEY.values())
_FLOATING_XML_ATTR_TO_KEY.update(
    {"tblpXSpec": "horizontal_alignment", "tblpYSpec": "vertical_alignment"}
)
_FLOATING_KEY_TO_XML_ATTR = {key: attr for attr, key in _FLOATING_XML_ATTR_TO_KEY.items()}
_FLOATING_STRING_VALUES = {
    "horizontal_anchor": _VALID_FLOATING_HORIZONTAL_ANCHORS,
    "vertical_anchor": _VALID_FLOATING_VERTICAL_ANCHORS,
    "horizontal_alignment": _VALID_FLOATING_HORIZONTAL_ALIGNMENTS,
    "vertical_alignment": _VALID_FLOATING_VERTICAL_ALIGNMENTS,
}
_FLOATING_DISTANCE_KEYS = frozenset(
    {"left_from_text", "right_from_text", "top_from_text", "bottom_from_text"}
)

_HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
_FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
_HEADER_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
_FOOTER_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"


def _is_valid_part_name(part: str) -> bool:
    return (
        isinstance(part, str)
        and bool(part)
        and not part.startswith("/")
        and "\\" not in part
        and ".." not in part.split("/")
        and posixpath.normpath(part) == part
        and part.endswith(".xml")
    )


def _check_reachable_part(
    graph: PackageGraph,
    target_part: str,
    prefix: str,
) -> tuple[bool, str | None]:
    """Check that target_part is _DOCUMENT or a genuinely referenced header/footer in graph.

    Returns (is_valid, error_code).
    """
    if target_part == _DOCUMENT:
        if _DOCUMENT not in graph.parts:
            return False, f"{prefix}_part_unsupported"
        return True, None

    if not _is_valid_part_name(target_part):
        return False, f"{prefix}_part_invalid"

    package_part = graph.part(target_part)
    if package_part is None:
        return False, f"{prefix}_part_unsupported"

    is_header = target_part.startswith("word/header") and target_part.endswith(".xml")
    is_footer = target_part.startswith("word/footer") and target_part.endswith(".xml")
    if not (is_header or is_footer):
        return False, f"{prefix}_part_unsupported"

    if is_header:
        expected_type = _HEADER_CONTENT_TYPE
        rel_type = _HEADER_RELATIONSHIP
        ref_kind = "headerReference"
    else:
        expected_type = _FOOTER_CONTENT_TYPE
        rel_type = _FOOTER_RELATIONSHIP
        ref_kind = "footerReference"

    if package_part.content_type != expected_type:
        return False, f"{prefix}_part_unsupported"

    matching_rel_ids = {
        rel.id
        for rel in graph.relationships_for(_DOCUMENT)
        if not rel.is_external
        and rel.resolved_target() == target_part
        and (
            rel.type == rel_type
            or rel.type.rsplit("/", 1)[-1] == ("header" if is_header else "footer")
        )
    }
    if not matching_rel_ids:
        return False, f"{prefix}_part_unsupported"

    try:
        doc_bytes = graph.read_part_bytes(_DOCUMENT)
        doc_root = ET.fromstring(doc_bytes)
    except (PackageError, ET.ParseError, KeyError, OSError):
        return False, f"{prefix}_package_invalid"

    doc_ns = _word_namespace(doc_root)
    if doc_ns is None:
        return False, f"{prefix}_namespace_unsupported"

    is_referenced = any(
        element.tag == Q(doc_ns, ref_kind)
        and any(
            attribute.rsplit("}", 1)[-1] == "id" and value in matching_rel_ids
            for attribute, value in element.attrib.items()
        )
        for element in doc_root.iter()
    )
    if not is_referenced:
        return False, f"{prefix}_part_unsupported"

    try:
        part_bytes = graph.read_part_bytes(target_part)
        part_root = ET.fromstring(part_bytes)
    except (PackageError, ET.ParseError, KeyError, OSError):
        return False, f"{prefix}_package_invalid"

    part_ns = _word_namespace(part_root)
    if part_ns is None:
        return False, f"{prefix}_namespace_unsupported"

    expected_tag = Q(part_ns, "hdr" if is_header else "ftr")
    if part_root.tag != expected_tag:
        return False, f"{prefix}_part_unsupported"

    return True, None


_LOOK_BIT_FIRST_ROW = 0x0020
_LOOK_BIT_LAST_ROW = 0x0040
_LOOK_BIT_FIRST_COLUMN = 0x0080
_LOOK_BIT_LAST_COLUMN = 0x0100
_LOOK_BIT_NO_HBAND = 0x0200
_LOOK_BIT_NO_VBAND = 0x0400

_LOOK_FLAG_TO_BIT = {
    "first_row": _LOOK_BIT_FIRST_ROW,
    "last_row": _LOOK_BIT_LAST_ROW,
    "first_column": _LOOK_BIT_FIRST_COLUMN,
    "last_column": _LOOK_BIT_LAST_COLUMN,
    "no_hband": _LOOK_BIT_NO_HBAND,
    "no_vband": _LOOK_BIT_NO_VBAND,
}

_LOOK_FLAG_TO_XML_ATTR = {
    "first_row": "firstRow",
    "last_row": "lastRow",
    "first_column": "firstColumn",
    "last_column": "lastColumn",
    "no_hband": "noHBand",
    "no_vband": "noVBand",
}

_LOOK_XML_ATTR_TO_FLAG = {
    "firstRow": "first_row",
    "lastRow": "last_row",
    "firstColumn": "first_column",
    "lastColumn": "last_column",
    "noHBand": "no_hband",
    "noVBand": "no_vband",
}

_ALLOWED_LOOK_INPUT_KEYS = frozenset(
    {
        "first_row",
        "last_row",
        "first_column",
        "last_column",
        "no_hband",
        "no_vband",
        "val",
        "firstRow",
        "lastRow",
        "firstColumn",
        "lastColumn",
        "noHBand",
        "noVBand",
    }
)

_NORMALIZED_LOOK_KEYS = frozenset(
    {
        "first_row",
        "last_row",
        "first_column",
        "last_column",
        "no_hband",
        "no_vband",
        "val",
    }
)

_REQUIRED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "properties",
    }
)
_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "target",
        "structure_fingerprint_sha256",
        "properties",
    }
)

_REQUIRED_TARGET_FIELDS = frozenset({"table"})
_ALLOWED_TARGET_FIELDS = frozenset({"table", "part"})

_ALLOWED_PROPERTY_KEYS = frozenset(
    {
        "width",
        "alignment",
        "borders",
        "shading",
        "cell_margins",
        "style",
        "indent",
        "cell_spacing",
        "layout",
        "look",
        "floating",
        "autofit",
        "overlap",
        "bidi_visual",
        "row_band_size",
        "column_band_size",
        "caption",
        "description",
    }
)
_STYLES = "word/styles.xml"
_UNSET = object()
_SCALAR_PROPERTY_ELEMENTS = {
    "overlap": "tblOverlap",
    "bidi_visual": "bidiVisual",
    "row_band_size": "tblStyleRowBandSize",
    "column_band_size": "tblStyleColBandSize",
    "caption": "tblCaption",
    "description": "tblDescription",
}
_RESETTABLE_PROPERTY_ELEMENTS = {
    "width": "tblW",
    "alignment": "jc",
    "borders": "tblBorders",
    "shading": "shd",
    "cell_margins": "tblCellMar",
}
_BORDER_XML_SIDE_TO_PUBLIC = {"start": "left", "end": "right"}
_STRICT_WORD_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main"

_VALID_ALIGNMENTS = frozenset({"left", "center", "right", "start", "end"})
# python-docx WD_TABLE_ALIGNMENT integer values mapped to their ST_JcTable literal
_WD_TABLE_ALIGNMENT_VALUES: dict[int, str] = {0: "left", 1: "center", 2: "right"}


def _canonical_table_alignment(value: Any, allowed: frozenset[str]) -> str | None:
    """Resolve a table alignment to its lowercase ST_JcTable literal, or ``None``.

    Accepts the literal in any case, a python-docx ``WD_TABLE_ALIGNMENT`` member (by
    name), or that enumeration's integer value (0/1/2 for left/center/right).
    """
    if isinstance(value, enum.Enum):
        value = value.name
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        alignment = _WD_TABLE_ALIGNMENT_VALUES.get(value)
    elif isinstance(value, str):
        alignment = value.strip().lower()
    else:
        return None
    return alignment if alignment in allowed else None


def _table_properties_alignment(value: Any) -> str:
    """Validate a table ``alignment`` value, bare or wrapped as ``{"val": ...}``."""
    if isinstance(value, Mapping):
        if set(value) != {"val"}:
            raise ValueError("table_properties_fields_invalid")
        value = value["val"]
    elif isinstance(value, bool) or not isinstance(value, (str, int, enum.Enum)):
        raise TypeError("table_properties_fields_invalid")
    alignment = _canonical_table_alignment(value, _VALID_ALIGNMENTS)
    if alignment is None:
        raise ValueError("table_properties_fields_invalid")
    return alignment


_ALLOWED_BORDER_SIDES = frozenset({"top", "left", "bottom", "right", "insideH", "insideV"})
_BORDER_SIDE_ALIASES = {"inside_h": "insideH", "inside_v": "insideV"}
# Explicit "no border" override, as Word writes it; suppresses inherited style borders.
_CLEARED_BORDER = {"val": "none", "sz": 0, "space": 0, "color": "auto"}

_VALID_BORDER_STYLES = frozenset(
    {
        "single",
        "double",
        "dashed",
        "dotted",
        "none",
        "nil",
        "thick",
        "wave",
        "dashDot",
        "dotDash",
        "dashDotDot",
        "dotDotDash",
        "dashSmallGap",
        "dashDotStroked",
        "triple",
        "thinThickSmallGap",
        "thickThinSmallGap",
        "thinThickThinSmallGap",
        "thinThickMediumGap",
        "thickThinMediumGap",
        "thinThickThinMediumGap",
        "thinThickLargeGap",
        "thickThinLargeGap",
        "thinThickThinLargeGap",
        "wavyDouble",
        "wavyHeavy",
        "outset",
        "inset",
        "groove",
        "ridge",
    }
)

_ALLOWED_CELL_MARGIN_SIDES = frozenset({"top", "left", "bottom", "right"})

# ECMA-376 / ISO-IEC 29500-1 §17.4.58 CT_TblPr child element sequence order
_TBLPR_CHILD_ORDER: tuple[str, ...] = (
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
)

_TBL_BORDERS_ORDER: tuple[str, ...] = (
    "top",
    "left",
    "bottom",
    "right",
    "insideH",
    "insideV",
)

_TBL_CELL_MAR_ORDER: tuple[str, ...] = (
    "top",
    "left",
    "bottom",
    "right",
)


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is the source part byte for byte."""
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


def _is_supported_table(table: ET.Element, namespace: str) -> bool:
    """True if table has supported body table structure and at least one row."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("tblPr"), word("tblGrid"), word("tr")} for child in table):
        return False
    rows = table.findall(word("tr"))
    return len(rows) > 0


def _insert_or_replace_tblpr_child(
    tbl_pr: ET.Element, new_child: ET.Element, namespace: str
) -> None:
    """Insert or replace a child element in tblPr while strictly maintaining schema order."""
    local_name = new_child.tag.rsplit("}", 1)[-1]
    def word(local):
        return Q(namespace, local)
    for existing in list(tbl_pr.findall(word(local_name))):
        tbl_pr.remove(existing)
    try:
        new_order = _TBLPR_CHILD_ORDER.index(local_name)
    except ValueError:
        new_order = 999
    insert_idx = len(tbl_pr)
    for idx, child in enumerate(tbl_pr):
        child_local = child.tag.rsplit("}", 1)[-1]
        try:
            child_order = _TBLPR_CHILD_ORDER.index(child_local)
        except ValueError:
            child_order = 999
        if child_order > new_order:
            insert_idx = idx
            break
    tbl_pr.insert(insert_idx, new_child)


@dataclass(frozen=True)
class TableWidthSpec:
    value: int
    type: str = "dxa"

    def to_dict(self) -> dict[str, Any]:
        return {"value": self.value, "type": self.type}


@dataclass(frozen=True)
class BorderSpec:
    val: str = "single"
    sz: int | None = 4
    space: int | None = 0
    color: str | None = "auto"

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"val": self.val}
        if self.sz is not None:
            data["sz"] = self.sz
        if self.space is not None:
            data["space"] = self.space
        if self.color is not None:
            data["color"] = self.color
        return data


@dataclass(frozen=True)
class TableBordersSpec:
    top: BorderSpec | None = None
    left: BorderSpec | None = None
    bottom: BorderSpec | None = None
    right: BorderSpec | None = None
    insideH: BorderSpec | None = None
    insideV: BorderSpec | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for side in _TBL_BORDERS_ORDER:
            val = getattr(self, side)
            if val is not None:
                data[side] = val.to_dict() if isinstance(val, BorderSpec) else val
        return data


@dataclass(frozen=True)
class TableShadingSpec:
    fill: str
    val: str = "clear"
    color: str = "auto"

    def to_dict(self) -> dict[str, Any]:
        return {"fill": self.fill, "val": self.val, "color": self.color}


@dataclass(frozen=True)
class TableCellMarginsSpec:
    top: int | None = None
    left: int | None = None
    bottom: int | None = None
    right: int | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for side in _TBL_CELL_MAR_ORDER:
            val = getattr(self, side)
            if val is not None:
                data[side] = val
        return data


@dataclass(frozen=True)
class TableIndentSpec:
    value: int
    type: str = "dxa"

    def to_dict(self) -> dict[str, Any]:
        return {"value": self.value, "type": self.type}


@dataclass(frozen=True)
class TableCellSpacingSpec:
    value: int
    type: str = "dxa"

    def to_dict(self) -> dict[str, Any]:
        return {"value": self.value, "type": self.type}


@dataclass(frozen=True)
class TableLayoutSpec:
    type: str = "fixed"

    def __post_init__(self) -> None:
        if isinstance(self.type, bool) or not isinstance(self.type, str):
            raise TypeError("table_properties_fields_invalid")
        cleaned = self.type.strip()
        if cleaned not in _VALID_LAYOUT_TYPES:
            raise ValueError("table_properties_fields_invalid")
        object.__setattr__(self, "type", cleaned)

    def to_dict(self) -> dict[str, Any]:
        return {"type": self.type}


@dataclass(frozen=True)
class TableFloatingPositioningSpec:
    """Floating table anchors, numeric or aligned positions, and text clearances in twips."""

    horizontal_anchor: str | None = None
    vertical_anchor: str | None = None
    horizontal_offset: int | None = None
    vertical_offset: int | None = None
    left_from_text: int | None = None
    right_from_text: int | None = None
    top_from_text: int | None = None
    bottom_from_text: int | None = None
    horizontal_alignment: str | None = None
    vertical_alignment: str | None = None

    def __post_init__(self) -> None:
        for field, allowed in _FLOATING_STRING_VALUES.items():
            value = getattr(self, field)
            if value is not None:
                if isinstance(value, bool) or not isinstance(value, str):
                    raise TypeError("table_properties_fields_invalid")
                cleaned = value.strip()
                if cleaned not in allowed:
                    raise ValueError("table_properties_fields_invalid")
                object.__setattr__(self, field, cleaned)
        for field in _FLOATING_KEY_TO_XML_ATTR:
            if field in _FLOATING_STRING_VALUES:
                continue
            value = getattr(self, field)
            if value is None:
                continue
            if isinstance(value, bool) or not isinstance(value, int):
                raise TypeError("table_properties_fields_invalid")
            if field in _FLOATING_DISTANCE_KEYS:
                if value < 0 or value > 316800:
                    raise ValueError("table_properties_fields_invalid")
            elif value < -316800 or value > 316800:
                raise ValueError("table_properties_fields_invalid")

    def to_dict(self) -> dict[str, Any]:
        return {
            field: value
            for field in _FLOATING_KEY_TO_XML_ATTR
            if (value := getattr(self, field)) is not None
        }


def _decode_look_bitmask(bitmask: int) -> dict[str, bool]:
    return {
        "first_row": bool(bitmask & _LOOK_BIT_FIRST_ROW),
        "last_row": bool(bitmask & _LOOK_BIT_LAST_ROW),
        "first_column": bool(bitmask & _LOOK_BIT_FIRST_COLUMN),
        "last_column": bool(bitmask & _LOOK_BIT_LAST_COLUMN),
        "no_hband": bool(bitmask & _LOOK_BIT_NO_HBAND),
        "no_vband": bool(bitmask & _LOOK_BIT_NO_VBAND),
    }


def _encode_look_bitmask(flags: Mapping[str, bool], base_val: int = 0) -> int:
    mask = base_val & ~0x07E0
    for flag_name, bit in _LOOK_FLAG_TO_BIT.items():
        if flags.get(flag_name):
            mask |= bit
    return mask


def _verify_look_coherence(
    val: str,
    flags: Mapping[str, bool | None],
) -> None:
    mask = int(val, 16)
    decoded = _decode_look_bitmask(mask)
    for flag_name, expected in decoded.items():
        actual = flags.get(flag_name)
        if actual is not None and actual != expected:
            raise ValueError("table_properties_fields_invalid")


@dataclass(frozen=True)
class TableLookSpec:
    first_row: bool | None = None
    last_row: bool | None = None
    first_column: bool | None = None
    last_column: bool | None = None
    no_hband: bool | None = None
    no_vband: bool | None = None
    val: str | None = None

    def __post_init__(self) -> None:
        for field in (
            "first_row",
            "last_row",
            "first_column",
            "last_column",
            "no_hband",
            "no_vband",
        ):
            val = getattr(self, field)
            if val is not None and not isinstance(val, bool):
                raise TypeError("table_properties_fields_invalid")
        if self.val is not None:
            if isinstance(self.val, bool) or not isinstance(self.val, str):
                raise TypeError("table_properties_fields_invalid")
            if not _LOOK_HEX_PATTERN.match(self.val):
                raise ValueError("table_properties_fields_invalid")
            _verify_look_coherence(
                self.val,
                {
                    "first_row": self.first_row,
                    "last_row": self.last_row,
                    "first_column": self.first_column,
                    "last_column": self.last_column,
                    "no_hband": self.no_hband,
                    "no_vband": self.no_vband,
                },
            )

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        if self.first_row is not None:
            data["first_row"] = self.first_row
        if self.last_row is not None:
            data["last_row"] = self.last_row
        if self.first_column is not None:
            data["first_column"] = self.first_column
        if self.last_column is not None:
            data["last_column"] = self.last_column
        if self.no_hband is not None:
            data["no_hband"] = self.no_hband
        if self.no_vband is not None:
            data["no_vband"] = self.no_vband
        if self.val is not None:
            data["val"] = self.val.upper()
        return data


@dataclass(frozen=True)
class TablePropertiesSpec:
    """Omitted fields stay unchanged; explicit ``None`` removes nullable direct properties."""

    width: TableWidthSpec | None | object = _UNSET
    alignment: str | None | object = _UNSET
    borders: TableBordersSpec | bool | None | object = _UNSET
    shading: TableShadingSpec | None | object = _UNSET
    cell_margins: TableCellMarginsSpec | None | object = _UNSET
    style: str | None | object = _UNSET
    indent: TableIndentSpec | TableWidthSpec | int | dict[str, Any] | None | object = _UNSET
    cell_spacing: TableCellSpacingSpec | TableWidthSpec | int | dict[str, Any] | None | object = (
        _UNSET
    )
    layout: TableLayoutSpec | str | dict[str, Any] | None | object = _UNSET
    look: TableLookSpec | dict[str, Any] | None | object = _UNSET
    floating: TableFloatingPositioningSpec | dict[str, Any] | None | object = _UNSET
    autofit: bool | object = _UNSET
    overlap: str | None | object = _UNSET
    bidi_visual: bool | None | object = _UNSET
    row_band_size: int | None | object = _UNSET
    column_band_size: int | None | object = _UNSET
    caption: str | None | object = _UNSET
    description: str | None | object = _UNSET

    def __post_init__(self) -> None:
        if self.alignment is not _UNSET and self.alignment is not None:
            object.__setattr__(self, "alignment", _table_properties_alignment(self.alignment))

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        if self.width is not _UNSET:
            data["width"] = (
                self.width.to_dict() if isinstance(self.width, TableWidthSpec) else self.width
            )
        if self.alignment is not _UNSET:
            data["alignment"] = self.alignment
        if self.borders is not _UNSET:
            data["borders"] = (
                self.borders.to_dict()
                if isinstance(self.borders, TableBordersSpec)
                else self.borders
            )
        if self.shading is not _UNSET:
            data["shading"] = (
                self.shading.to_dict()
                if isinstance(self.shading, TableShadingSpec)
                else self.shading
            )
        if self.cell_margins is not _UNSET:
            data["cell_margins"] = (
                self.cell_margins.to_dict()
                if isinstance(self.cell_margins, TableCellMarginsSpec)
                else self.cell_margins
            )
        if self.style is not _UNSET:
            data["style"] = self.style
        if self.indent is not _UNSET:
            if self.indent is None:
                data["indent"] = None
            elif hasattr(self.indent, "to_dict"):
                data["indent"] = self.indent.to_dict()
            else:
                data["indent"] = self.indent
        if self.cell_spacing is not _UNSET:
            if self.cell_spacing is None:
                data["cell_spacing"] = None
            elif hasattr(self.cell_spacing, "to_dict"):
                data["cell_spacing"] = self.cell_spacing.to_dict()
            else:
                data["cell_spacing"] = self.cell_spacing
        if self.layout is not _UNSET:
            if self.layout is None:
                data["layout"] = None
            elif hasattr(self.layout, "to_dict"):
                data["layout"] = self.layout.to_dict()
            else:
                data["layout"] = self.layout
        if self.look is not _UNSET:
            if self.look is None:
                data["look"] = None
            elif hasattr(self.look, "to_dict"):
                data["look"] = self.look.to_dict()
            else:
                data["look"] = self.look
        if self.floating is not _UNSET:
            if self.floating is None:
                data["floating"] = None
            elif hasattr(self.floating, "to_dict"):
                data["floating"] = self.floating.to_dict()
            else:
                data["floating"] = self.floating
        if self.autofit is not _UNSET:
            data["autofit"] = self.autofit
        for field in _SCALAR_PROPERTY_ELEMENTS:
            value = getattr(self, field)
            if value is not _UNSET:
                data[field] = value
        return data


def _validate_table_properties_target_dict(target: Any) -> dict[str, Any]:
    if not isinstance(target, Mapping):
        raise TypeError("table_properties_target_invalid")
    keys = set(target)
    if not keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(keys):
        raise ValueError("table_properties_target_fields_invalid")
    table = target["table"]
    if isinstance(table, bool) or not isinstance(table, int) or table < 1:
        raise ValueError("table_properties_table_invalid")
    part = target.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("table_properties_part_invalid")
    return {"table": table, "part": part}


def _validate_floating_positioning(value: Any) -> dict[str, Any]:
    if isinstance(value, TableFloatingPositioningSpec):
        value = value.to_dict()
    if not isinstance(value, Mapping):
        raise TypeError("table_properties_fields_invalid")
    if not value or not set(value).issubset(_FLOATING_KEY_TO_XML_ATTR):
        raise ValueError("table_properties_fields_invalid")

    normalized: dict[str, Any] = {}
    for key, allowed in _FLOATING_STRING_VALUES.items():
        if key not in value:
            continue
        raw_anchor = value[key]
        if isinstance(raw_anchor, bool) or not isinstance(raw_anchor, str):
            raise TypeError("table_properties_fields_invalid")
        anchor = raw_anchor.strip()
        if anchor not in allowed:
            raise ValueError("table_properties_fields_invalid")
        normalized[key] = anchor

    for key in _FLOATING_REQUIRED_KEYS - _FLOATING_STRING_VALUES.keys():
        if key not in value:
            continue
        raw_measure = value[key]
        if isinstance(raw_measure, bool) or not isinstance(raw_measure, int):
            raise TypeError("table_properties_fields_invalid")
        if key in _FLOATING_DISTANCE_KEYS:
            if raw_measure < 0 or raw_measure > 316800:
                raise ValueError("table_properties_fields_invalid")
        elif raw_measure < -316800 or raw_measure > 316800:
            raise ValueError("table_properties_fields_invalid")
        normalized[key] = raw_measure
    return normalized


def _validate_table_properties_dict(properties: Any) -> dict[str, Any]:
    if isinstance(properties, TablePropertiesSpec):
        properties = properties.to_dict()
    if not isinstance(properties, Mapping):
        raise TypeError("table_properties_fields_invalid")
    keys = set(properties)
    if not keys:
        raise ValueError("table_properties_fields_invalid")
    if not keys.issubset(_ALLOWED_PROPERTY_KEYS):
        raise ValueError("table_properties_fields_invalid")

    result: dict[str, Any] = {
        key: None
        for key in _RESETTABLE_PROPERTY_ELEMENTS
        if key in properties and properties[key] is None
    }
    properties = {key: value for key, value in properties.items() if key not in result}

    for key in _SCALAR_PROPERTY_ELEMENTS:
        if key not in properties:
            continue
        value = properties[key]
        if value is None:
            result[key] = None
        elif key == "overlap":
            if not isinstance(value, str):
                raise TypeError("table_properties_fields_invalid")
            cleaned = value.strip()
            if cleaned not in {"never", "overlap"}:
                raise ValueError("table_properties_fields_invalid")
            result[key] = cleaned
        elif key == "bidi_visual":
            if not isinstance(value, bool):
                raise TypeError("table_properties_fields_invalid")
            result[key] = value
        elif key in {"row_band_size", "column_band_size"}:
            if isinstance(value, bool) or not isinstance(value, int):
                raise TypeError("table_properties_fields_invalid")
            if value < 0:
                raise ValueError("table_properties_fields_invalid")
            result[key] = value
        else:
            if not isinstance(value, str):
                raise TypeError("table_properties_fields_invalid")
            result[key] = value

    if "width" in properties:
        w_val = properties["width"]
        if isinstance(w_val, TableWidthSpec):
            w_val = w_val.to_dict()
        if isinstance(w_val, int) and not isinstance(w_val, bool):
            if w_val < 0 or w_val > 316800:
                raise ValueError("table_properties_fields_invalid")
            result["width"] = {"value": w_val, "type": "dxa"}
        elif isinstance(w_val, Mapping):
            w_keys = set(w_val)
            if not w_keys.issubset({"value", "w", "type"}) or not ({"value", "w"} & w_keys):
                raise ValueError("table_properties_fields_invalid")
            raw_w = w_val.get("value", w_val.get("w"))
            if isinstance(raw_w, bool) or not isinstance(raw_w, int) or raw_w < 0:
                raise ValueError("table_properties_fields_invalid")
            w_type = w_val.get("type", "dxa")
            if not isinstance(w_type, str) or w_type not in (
                "dxa",
                "pct",
                "auto",
                "nil",
            ):
                raise ValueError("table_properties_fields_invalid")
            if w_type == "dxa" and raw_w > 316800:
                raise ValueError("table_properties_fields_invalid")
            if w_type == "pct" and raw_w > 50000:
                raise ValueError("table_properties_fields_invalid")
            result["width"] = {"value": raw_w, "type": w_type}
        else:
            raise TypeError("table_properties_fields_invalid")

    if "alignment" in properties:
        result["alignment"] = _table_properties_alignment(properties["alignment"])

    if "borders" in properties:
        b_val = properties["borders"]
        if isinstance(b_val, TableBordersSpec):
            b_val = b_val.to_dict()
        if b_val is False:
            b_val = dict.fromkeys(_TBL_BORDERS_ORDER, False)
        if not isinstance(b_val, Mapping) or not b_val:
            raise ValueError("table_properties_fields_invalid")
        canonical_sides: dict[str, Any] = {}
        for raw_side, raw_side_val in b_val.items():
            side = _BORDER_SIDE_ALIASES.get(raw_side, raw_side)
            if side not in _ALLOWED_BORDER_SIDES or side in canonical_sides:
                raise ValueError("table_properties_fields_invalid")
            canonical_sides[side] = raw_side_val
        borders_dict: dict[str, Any] = {}
        for side, side_val in canonical_sides.items():
            if side_val is False or (isinstance(side_val, str) and side_val == "none"):
                borders_dict[side] = dict(_CLEARED_BORDER)
                continue
            if isinstance(side_val, BorderSpec):
                side_val = side_val.to_dict()
            if not isinstance(side_val, Mapping):
                raise TypeError("table_properties_fields_invalid")
            side_keys = set(side_val)
            if not side_keys.issubset({"val", "style", "sz", "size", "space", "color"}):
                raise ValueError("table_properties_fields_invalid")
            val_style = side_val.get("val", side_val.get("style", "single"))
            if not isinstance(val_style, str) or val_style not in _VALID_BORDER_STYLES:
                raise ValueError("table_properties_fields_invalid")
            sz = side_val.get("sz", side_val.get("size", 4))
            if sz is not None and (
                isinstance(sz, bool) or not isinstance(sz, int) or sz < 0 or sz > 96
            ):
                raise ValueError("table_properties_fields_invalid")
            space = side_val.get("space", 0)
            if space is not None and (
                isinstance(space, bool) or not isinstance(space, int) or space < 0 or space > 64
            ):
                raise ValueError("table_properties_fields_invalid")
            color = side_val.get("color", "auto")
            if color is not None and (
                not isinstance(color, str)
                or (color != "auto" and not _HEX_COLOR_PATTERN.match(color))
            ):
                raise ValueError("table_properties_fields_invalid")
            borders_dict[side] = {
                "val": val_style,
                "sz": sz,
                "space": space,
                "color": color,
            }
        result["borders"] = borders_dict

    if "shading" in properties:
        s_val = properties["shading"]
        if isinstance(s_val, TableShadingSpec):
            s_val = s_val.to_dict()
        if isinstance(s_val, str):
            if s_val != "auto" and not _HEX_COLOR_PATTERN.match(s_val):
                raise ValueError("table_properties_fields_invalid")
            result["shading"] = {"fill": s_val, "val": "clear", "color": "auto"}
        elif isinstance(s_val, Mapping):
            s_keys = set(s_val)
            if not s_keys.issubset({"fill", "val", "color"}) or "fill" not in s_keys:
                raise ValueError("table_properties_fields_invalid")
            fill = s_val["fill"]
            if not isinstance(fill, str) or (fill != "auto" and not _HEX_COLOR_PATTERN.match(fill)):
                raise ValueError("table_properties_fields_invalid")
            val = s_val.get("val", "clear")
            if not isinstance(val, str) or not val:
                raise ValueError("table_properties_fields_invalid")
            color = s_val.get("color", "auto")
            if not isinstance(color, str) or (
                color != "auto" and not _HEX_COLOR_PATTERN.match(color)
            ):
                raise ValueError("table_properties_fields_invalid")
            result["shading"] = {"fill": fill, "val": val, "color": color}
        else:
            raise TypeError("table_properties_fields_invalid")

    if "cell_margins" in properties:
        m_val = properties["cell_margins"]
        if isinstance(m_val, TableCellMarginsSpec):
            m_val = m_val.to_dict()
        if not isinstance(m_val, Mapping) or not m_val:
            raise ValueError("table_properties_fields_invalid")
        m_keys = set(m_val)
        if not m_keys.issubset(_ALLOWED_CELL_MARGIN_SIDES):
            raise ValueError("table_properties_fields_invalid")
        mar_dict: dict[str, Any] = {}
        for side, side_val in m_val.items():
            if (
                isinstance(side_val, bool)
                or not isinstance(side_val, int)
                or side_val < 0
                or side_val > 28800
            ):
                raise ValueError("table_properties_fields_invalid")
            mar_dict[side] = side_val
        result["cell_margins"] = mar_dict

    if "style" in properties:
        st_val = properties["style"]
        if st_val is None:
            result["style"] = None
        elif isinstance(st_val, str):
            style_clean = st_val.strip()
            if not style_clean:
                raise ValueError("table_properties_fields_invalid")
            result["style"] = style_clean
        else:
            raise TypeError("table_properties_fields_invalid")
    if "indent" in properties:
        ind_val = properties["indent"]
        if ind_val is None:
            result["indent"] = None
        else:
            if isinstance(ind_val, (TableIndentSpec, TableWidthSpec)):
                ind_val = ind_val.to_dict()
            if isinstance(ind_val, int) and not isinstance(ind_val, bool):
                if ind_val < -316800 or ind_val > 316800:
                    raise ValueError("table_properties_fields_invalid")
                result["indent"] = {"value": ind_val, "type": "dxa"}
            elif isinstance(ind_val, Mapping):
                ind_keys = set(ind_val)
                if not ind_keys.issubset({"value", "w", "type"}) or not ({"value", "w"} & ind_keys):
                    raise ValueError("table_properties_fields_invalid")
                raw_ind = ind_val.get("value", ind_val.get("w"))
                if isinstance(raw_ind, bool) or not isinstance(raw_ind, int):
                    raise TypeError("table_properties_fields_invalid")
                ind_type = ind_val.get("type", "dxa")
                if not isinstance(ind_type, str) or ind_type not in (
                    "dxa",
                    "pct",
                    "auto",
                    "nil",
                ):
                    raise ValueError("table_properties_fields_invalid")
                if ind_type == "dxa" and (raw_ind < -316800 or raw_ind > 316800):
                    raise ValueError("table_properties_fields_invalid")
                if ind_type == "pct" and (raw_ind < -50000 or raw_ind > 50000):
                    raise ValueError("table_properties_fields_invalid")
                result["indent"] = {"value": raw_ind, "type": ind_type}
            else:
                raise TypeError("table_properties_fields_invalid")

    if "cell_spacing" in properties:
        cs_val = properties["cell_spacing"]
        if cs_val is None:
            result["cell_spacing"] = None
        else:
            if isinstance(cs_val, (TableCellSpacingSpec, TableWidthSpec)):
                cs_val = cs_val.to_dict()
            if isinstance(cs_val, int) and not isinstance(cs_val, bool):
                if cs_val < 0 or cs_val > 316800:
                    raise ValueError("table_properties_fields_invalid")
                result["cell_spacing"] = {"value": cs_val, "type": "dxa"}
            elif isinstance(cs_val, Mapping):
                cs_keys = set(cs_val)
                if not cs_keys.issubset({"value", "w", "type"}) or not ({"value", "w"} & cs_keys):
                    raise ValueError("table_properties_fields_invalid")
                raw_cs = cs_val.get("value", cs_val.get("w"))
                if isinstance(raw_cs, bool) or not isinstance(raw_cs, int):
                    raise TypeError("table_properties_fields_invalid")
                if raw_cs < 0:
                    raise ValueError("table_properties_fields_invalid")
                cs_type = cs_val.get("type", "dxa")
                if not isinstance(cs_type, str) or cs_type not in (
                    "dxa",
                    "pct",
                    "auto",
                    "nil",
                ):
                    raise ValueError("table_properties_fields_invalid")
                if cs_type == "dxa" and raw_cs > 316800:
                    raise ValueError("table_properties_fields_invalid")
                if cs_type == "pct" and raw_cs > 50000:
                    raise ValueError("table_properties_fields_invalid")
                result["cell_spacing"] = {"value": raw_cs, "type": cs_type}
            else:
                raise TypeError("table_properties_fields_invalid")

    if "layout" in properties:
        lay_val = properties["layout"]
        if lay_val is None:
            result["layout"] = None
        elif isinstance(lay_val, TableLayoutSpec):
            result["layout"] = lay_val.to_dict()
        elif isinstance(lay_val, str):
            clean_lay = lay_val.strip()
            if clean_lay not in _VALID_LAYOUT_TYPES:
                raise ValueError("table_properties_fields_invalid")
            result["layout"] = {"type": clean_lay}
        elif isinstance(lay_val, Mapping):
            lay_keys = set(lay_val)
            if "val" in lay_keys:
                raise ValueError("table_properties_fields_invalid")
            if not lay_keys or not lay_keys.issubset({"type"}):
                raise ValueError("table_properties_fields_invalid")
            raw_type = lay_val.get("type")
            if isinstance(raw_type, bool) or not isinstance(raw_type, str):
                raise TypeError("table_properties_fields_invalid")
            clean_type = raw_type.strip()
            if clean_type not in _VALID_LAYOUT_TYPES:
                raise ValueError("table_properties_fields_invalid")
            result["layout"] = {"type": clean_type}
        else:
            raise TypeError("table_properties_fields_invalid")

    if "look" in properties:
        lk_val = properties["look"]
        if lk_val is None:
            result["look"] = None
        else:
            if isinstance(lk_val, TableLookSpec):
                lk_val = lk_val.to_dict()
            if not isinstance(lk_val, Mapping) or not lk_val:
                raise (
                    TypeError("table_properties_fields_invalid")
                    if not isinstance(lk_val, Mapping)
                    else ValueError("table_properties_fields_invalid")
                )
            lk_keys = set(lk_val)
            if not lk_keys.issubset(_ALLOWED_LOOK_INPUT_KEYS):
                raise ValueError("table_properties_fields_invalid")

            norm_look: dict[str, Any] = {}
            for k, v in lk_val.items():
                canonical_k = _LOOK_XML_ATTR_TO_FLAG.get(k, k)
                if canonical_k == "val":
                    if isinstance(v, bool) or not isinstance(v, str):
                        raise TypeError("table_properties_fields_invalid")
                    if not _LOOK_HEX_PATTERN.match(v):
                        raise ValueError("table_properties_fields_invalid")
                    norm_look["val"] = v.upper()
                else:
                    if not isinstance(v, bool):
                        raise TypeError("table_properties_fields_invalid")
                    norm_look[canonical_k] = v

            if "val" in norm_look:
                _verify_look_coherence(norm_look["val"], norm_look)

            result["look"] = norm_look

    if "floating" in properties:
        floating_val = properties["floating"]
        if floating_val is None:
            result["floating"] = None
        else:
            result["floating"] = _validate_floating_positioning(floating_val)

    if "autofit" in properties:
        autofit_val = properties["autofit"]
        if not isinstance(autofit_val, bool):
            raise TypeError("table_properties_fields_invalid")
        autofit_layout = {"type": "autofit" if autofit_val else "fixed"}
        if "layout" in result and result["layout"] != autofit_layout:
            raise ValueError("table_properties_fields_invalid")
        result["layout"] = autofit_layout
    return result


@dataclass(frozen=True)
class TablePropertiesTarget:
    """One discovered candidate table target for table properties authoring."""

    table: int
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    current_properties: dict[str, Any] | None = None
    style: str | None = None
    rows_count: int = 0
    columns_count: int = 0

    @property
    def target(self) -> dict[str, Any]:
        return {"table": self.table, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "part": self.part,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "current_properties": self.current_properties,
            "style": self.style,
            "rows_count": self.rows_count,
            "columns_count": self.columns_count,
        }


@dataclass(frozen=True)
class TablePropertiesRequest:
    """One source-bound request to author direct formatting table properties on an exact table."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    properties: Mapping[str, Any] | TablePropertiesSpec

    def __post_init__(self) -> None:
        if isinstance(self.source, str):
            object.__setattr__(self, "source", Path(self.source))
        elif not isinstance(self.source, Path):
            raise TypeError("table_properties_source_invalid")
        if not str(self.source):
            raise ValueError("table_properties_source_invalid")

        if isinstance(self.destination, str):
            object.__setattr__(self, "destination", Path(self.destination))
        elif not isinstance(self.destination, Path):
            raise TypeError("table_properties_destination_invalid")
        if not str(self.destination):
            raise ValueError("table_properties_destination_invalid")

        _require_sha256(self.source_sha256, "table_properties_source_sha256")
        _require_sha256(
            self.structure_fingerprint_sha256,
            "table_properties_structure_fingerprint_sha256",
        )

        validated_target = _validate_table_properties_target_dict(self.target)
        object.__setattr__(self, "target", validated_target)

        validated_properties = _validate_table_properties_dict(self.properties)
        object.__setattr__(self, "properties", validated_properties)

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        return {"table": self.table, "part": self.part}

    def properties_dict(self) -> dict[str, Any]:
        return dict(self.properties)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "target": self.target_dict(),
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "properties": self.properties_dict(),
        }


@dataclass(frozen=True)
class TablePropertiesReport:
    """Content-safe public result for a table properties authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    properties: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "table-properties",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "target": self.target,
            "properties": self.properties,
            "target_change_report": self.target_change_report,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def _table_properties_refusal(
    request: TablePropertiesRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TablePropertiesReport:
    return TablePropertiesReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        properties=request.properties_dict(),
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_properties_request(value: Any) -> TablePropertiesRequest:
    """Parse and strictly validate a table properties authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_properties_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_properties_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("table_properties_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("table_properties_destination_invalid")

    source_sha256 = _require_sha256(value["source_sha256"], "table_properties_source_sha256")
    structure_fingerprint = _require_sha256(
        value["structure_fingerprint_sha256"],
        "table_properties_structure_fingerprint_sha256",
    )

    target = _validate_table_properties_target_dict(value["target"])
    properties = _validate_table_properties_dict(value["properties"])

    return TablePropertiesRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        target=target,
        structure_fingerprint_sha256=structure_fingerprint,
        properties=properties,
    )


def load_table_properties_request(path: str | Path) -> TablePropertiesRequest:
    """Load and strictly validate one JSON table-properties request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_properties_request(json.load(stream))


def _table_structure_fingerprint(table: ET.Element) -> str:
    """Hash the table's XML structure while deliberately masking cell text."""
    return _cell_structure_fingerprint(table)


def table_properties_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one target table."""
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        valid, _ = _check_reachable_part(graph, part, "table_properties")
        if not valid:
            return None
        root = ET.fromstring(graph.read_part_bytes(part))
        ns = _word_namespace(root)
        if ns is None:
            return None
        tables = list(root.iter(Q(ns, "tbl")))
        if table < 1 or table > len(tables):
            return None
        target_table = tables[table - 1]
        if not _is_supported_table(target_table, ns):
            return None
        return _table_structure_fingerprint(target_table)
    except (PackageError, OSError, ET.ParseError, KeyError):
        return None


def _floating_positioning_from_element(tbl_pr: ET.Element, namespace: str) -> dict[str, Any]:
    def word(local):
        return Q(namespace, local)
    tblp_pr = tbl_pr.find(word("tblpPr"))
    if tblp_pr is None:
        return {}

    result: dict[str, Any] = {}
    for xml_attr, key in _FLOATING_XML_ATTR_TO_KEY.items():
        raw_value = tblp_pr.get(word(xml_attr))
        if raw_value is None:
            continue
        if key in _FLOATING_STRING_VALUES:
            if raw_value in _FLOATING_STRING_VALUES[key]:
                result[key] = raw_value
        else:
            try:
                value = int(raw_value)
            except ValueError:
                continue
            if key in _FLOATING_DISTANCE_KEYS:
                if 0 <= value <= 316800:
                    result[key] = value
            elif -316800 <= value <= 316800:
                result[key] = value
    return result


def _inspect_table_properties(tbl_pr: ET.Element | None, namespace: str) -> dict[str, Any]:
    if tbl_pr is None:
        return {}
    def word(local):
        return Q(namespace, local)
    result: dict[str, Any] = {}
    for key, local in _SCALAR_PROPERTY_ELEMENTS.items():
        element = tbl_pr.find(word(local))
        if element is None:
            continue
        raw_value = element.get(word("val"))
        if key == "bidi_visual":
            if raw_value is None:
                result[key] = True
            elif raw_value.lower() in {"1", "true", "on", "0", "false", "off"}:
                result[key] = raw_value.lower() in {"1", "true", "on"}
        elif raw_value is not None:
            if key == "overlap":
                if raw_value in {"never", "overlap"}:
                    result[key] = raw_value
            elif key in {"row_band_size", "column_band_size"}:
                if re.fullmatch(r"\+?[0-9]+", raw_value.strip()):
                    try:
                        result[key] = int(raw_value)
                    except ValueError:
                        continue
            else:
                result[key] = raw_value

    tbl_style = tbl_pr.find(word("tblStyle"))
    if tbl_style is not None:
        style_val = tbl_style.get(word("val"))
        if style_val is not None:
            result["style"] = style_val
    tbl_w = tbl_pr.find(word("tblW"))
    floating = _floating_positioning_from_element(tbl_pr, namespace)
    if floating:
        result["floating"] = floating
    if tbl_w is not None:
        raw_w = tbl_w.get(word("w"))
        w_val = int(raw_w) if raw_w is not None and raw_w.isdigit() else 0
        result["width"] = {
            "value": w_val,
            "type": tbl_w.get(word("type")) or "auto",
        }

    jc = tbl_pr.find(word("jc"))
    if jc is not None:
        result["alignment"] = jc.get(word("val")) or "left"

    tbl_borders = tbl_pr.find(word("tblBorders"))
    if tbl_borders is not None:
        borders: dict[str, Any] = {}
        for child in tbl_borders:
            xml_side = child.tag.rsplit("}", 1)[-1]
            local = _BORDER_XML_SIDE_TO_PUBLIC.get(xml_side, xml_side)
            if local in _ALLOWED_BORDER_SIDES and (
                local not in borders or xml_side in _BORDER_XML_SIDE_TO_PUBLIC
            ):
                sz_raw = child.get(word("sz"))
                space_raw = child.get(word("space"))
                borders[local] = {
                    "val": child.get(word("val")) or "single",
                    "sz": int(sz_raw) if sz_raw is not None and sz_raw.isdigit() else None,
                    "space": int(space_raw) if space_raw is not None and space_raw.isdigit() else 0,
                    "color": child.get(word("color")) or "auto",
                }
        if borders:
            result["borders"] = borders

    shd = tbl_pr.find(word("shd"))
    if shd is not None:
        result["shading"] = {
            "fill": shd.get(word("fill")) or "auto",
            "val": shd.get(word("val")) or "clear",
            "color": shd.get(word("color")) or "auto",
        }

    tbl_cell_mar = tbl_pr.find(word("tblCellMar"))
    if tbl_cell_mar is not None:
        margins: dict[str, Any] = {}
        for child in tbl_cell_mar:
            local = child.tag.rsplit("}", 1)[-1]
            if local in _ALLOWED_CELL_MARGIN_SIDES:
                w_raw = child.get(word("w"))
                if w_raw is not None and w_raw.isdigit():
                    margins[local] = int(w_raw)
        if margins:
            result["cell_margins"] = margins

    tbl_ind = tbl_pr.find(word("tblInd"))
    if tbl_ind is not None:
        raw_ind = tbl_ind.get(word("w"))
        ind_val = 0
        if raw_ind is not None:
            try:
                ind_val = int(raw_ind.strip())
            except ValueError:
                ind_val = 0
        result["indent"] = {
            "value": ind_val,
            "type": tbl_ind.get(word("type")) or "dxa",
        }

    tbl_cs = tbl_pr.find(word("tblCellSpacing"))
    if tbl_cs is not None:
        raw_cs = tbl_cs.get(word("w"))
        cs_val = 0
        if raw_cs is not None:
            try:
                cs_val = int(raw_cs.strip())
            except ValueError:
                cs_val = 0
        result["cell_spacing"] = {
            "value": cs_val,
            "type": tbl_cs.get(word("type")) or "dxa",
        }

    tbl_layout = tbl_pr.find(word("tblLayout"))
    if tbl_layout is not None:
        l_type = tbl_layout.get(word("type")) or "autofit"
        result["layout"] = {"type": l_type}

    tbl_look = tbl_pr.find(word("tblLook"))
    if tbl_look is not None:
        look_data: dict[str, Any] = {}
        val_attr = tbl_look.get(word("val"))
        val_mask = (
            int(val_attr, 16)
            if val_attr is not None and _LOOK_HEX_PATTERN.match(val_attr)
            else None
        )
        if val_attr is not None:
            look_data["val"] = val_attr.upper()

        for flag_name, xml_attr in _LOOK_FLAG_TO_XML_ATTR.items():
            raw_attr = tbl_look.get(word(xml_attr))
            if raw_attr is not None:
                look_data[flag_name] = raw_attr.lower() in ("1", "true", "on")
            elif val_mask is not None:
                look_data[flag_name] = bool(val_mask & _LOOK_FLAG_TO_BIT[flag_name])
            else:
                look_data[flag_name] = False
        result["look"] = look_data
    return result


def discover_table_properties_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table properties targets for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_properties_source_missing",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    try:
        source_sha256 = _sha256_file(path)
        graph = PackageGraph.from_file(path)
    except (PackageError, OSError, zipfile.BadZipFile):
        return {
            "schema_version": 1,
            "operation": "table-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_properties_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    valid_part, part_err = _check_reachable_part(graph, part, "table_properties")
    if not valid_part:
        return {
            "schema_version": 1,
            "operation": "table-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": part_err or "table_properties_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "table-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_properties_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_properties_namespace_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    def word(local):
        return Q(namespace, local)
    unsupported_context = _unsupported_context_by_element(root, namespace)
    tables = list(root.iter(word("tbl")))
    targets: list[dict[str, Any]] = []

    for table_idx, table in enumerate(tables, start=1):
        if unsupported_context.get(id(table), False):
            continue
        if not _is_supported_table(table, namespace):
            continue
        if len(table.findall(word("tblPr"))) > 1:
            continue
        fp = _table_structure_fingerprint(table)
        tbl_pr = table.find(word("tblPr"))
        current_props = _inspect_table_properties(tbl_pr, namespace)
        style = (
            tbl_pr.find(word("tblStyle")).get(word("val"))
            if tbl_pr is not None and tbl_pr.find(word("tblStyle")) is not None
            else None
        )
        rows = table.findall(word("tr"))
        tbl_grid = table.find(word("tblGrid"))
        grid_cols = tbl_grid.findall(word("gridCol")) if tbl_grid is not None else []
        cols_count = (
            len(grid_cols)
            if grid_cols
            else max((len(r.findall(word("tc"))) for r in rows), default=0)
        )

        target = TablePropertiesTarget(
            table=table_idx,
            structure_fingerprint_sha256=fp,
            part=part,
            current_properties=current_props,
            style=style,
            rows_count=len(rows),
            columns_count=cols_count,
        )
        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-properties-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def _validate_referenced_table_style(graph: PackageGraph, style_id: str) -> str | None:
    """Validate that referenced table style exists, is table type, and lies in supported base subset."""
    if not graph.has_part(_STYLES):
        return "table_properties_style_missing"
    try:
        styles_bytes = graph.read_part_bytes(_STYLES)
        styles_root = ET.fromstring(styles_bytes)
    except (PackageError, ET.ParseError, KeyError, OSError):
        return "table_properties_package_invalid"

    ns = _word_namespace(styles_root)
    if ns is None:
        return "table_properties_namespace_unsupported"

    def word(local):
        return Q(ns, local)

    def _get_val(elem: ET.Element | None, attr: str) -> str | None:
        if elem is None:
            return None
        return elem.get(word(attr)) or elem.get(attr)

    styles_by_id: dict[str, ET.Element] = {}
    for style_elem in styles_root.findall(word("style")):
        sid = _get_val(style_elem, "styleId")
        if sid:
            styles_by_id[sid] = style_elem

    if style_id not in styles_by_id:
        return "table_properties_style_missing"

    target_elem = styles_by_id[style_id]
    stype = _get_val(target_elem, "type")
    if stype != "table":
        return "table_properties_style_type_incompatible"

    visited: set[str] = set()
    curr_id: str | None = style_id
    depth = 0

    while curr_id is not None:
        if curr_id in visited:
            return "table_properties_style_unsupported"
        visited.add(curr_id)
        depth += 1
        if depth > 32:
            return "table_properties_style_unsupported"

        curr_elem = styles_by_id.get(curr_id)
        if curr_elem is None:
            return "table_properties_style_missing"

        curr_type = _get_val(curr_elem, "type")
        if curr_type != "table":
            return "table_properties_style_type_incompatible"

        if curr_elem.find(word("tblStylePr")) is not None:
            return "table_properties_style_unsupported"

        if curr_elem.find(word("next")) is not None:
            return "table_properties_style_unsupported"
        if curr_elem.find(word("pPr")) is not None:
            return "table_properties_style_unsupported"

        based_on_elem = curr_elem.find(word("basedOn"))
        if based_on_elem is not None:
            curr_id = _get_val(based_on_elem, "val")
            if not curr_id:
                curr_id = None
        else:
            curr_id = None

    return None


def author_table_properties(request: TablePropertiesRequest) -> TablePropertiesReport:
    """Author direct formatting table properties on a source-bound document table atomically."""
    if not isinstance(request, TablePropertiesRequest):
        raise TypeError("table_properties_request_invalid")

    source = request.source
    destination = request.destination

    if not source.is_file():
        return _table_properties_refusal(request, "table_properties_source_missing")

    if destination.exists():
        return _table_properties_refusal(request, "table_properties_destination_occupied")

    try:
        source_sha256 = _sha256_file(source)
    except OSError:
        return _table_properties_refusal(request, "table_properties_source_missing")

    if source_sha256 != request.source_sha256:
        return _table_properties_refusal(request, "table_properties_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_properties_refusal(request, "table_properties_package_invalid")

    valid_part, part_err = _check_reachable_part(graph, request.part, "table_properties")
    if not valid_part:
        return _table_properties_refusal(request, part_err or "table_properties_part_unsupported")

    try:
        doc_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (ET.ParseError, KeyError, OSError):
        return _table_properties_refusal(request, "table_properties_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_properties_refusal(request, "table_properties_namespace_unsupported")

    def word(local):
        return Q(namespace, local)
    tables = list(root.iter(word("tbl")))
    if request.table < 1 or request.table > len(tables):
        return _table_properties_refusal(request, "table_properties_target_missing")

    target_table = tables[request.table - 1]

    if not _is_supported_table(target_table, namespace):
        return _table_properties_refusal(request, "table_properties_context_unsupported")

    unsupported_context = _unsupported_context_by_element(root, namespace)
    if unsupported_context.get(id(target_table), False):
        return _table_properties_refusal(request, "table_properties_context_unsupported")

    # Conflict check: duplicate tblPr at table root
    tbl_prs = target_table.findall(word("tblPr"))
    if len(tbl_prs) > 1:
        return _table_properties_refusal(request, "table_properties_conflict")

    # Conflict check: duplicate same-name children inside tblPr
    tbl_pr = tbl_prs[0] if tbl_prs else None
    if tbl_pr is not None:
        child_local_tags = [c.tag.rsplit("}", 1)[-1] for c in tbl_pr]
        if len(child_local_tags) != len(set(child_local_tags)):
            return _table_properties_refusal(request, "table_properties_conflict")
        existing_borders = tbl_pr.find(word("tblBorders"))
        if existing_borders is not None:
            b_sides = [c.tag.rsplit("}", 1)[-1] for c in existing_borders]
            if len(b_sides) != len(set(b_sides)):
                return _table_properties_refusal(request, "table_properties_conflict")
        existing_cell_mar = tbl_pr.find(word("tblCellMar"))
        if existing_cell_mar is not None:
            m_sides = [c.tag.rsplit("}", 1)[-1] for c in existing_cell_mar]
            if len(m_sides) != len(set(m_sides)):
                return _table_properties_refusal(request, "table_properties_conflict")

    # Fingerprint check
    actual_fp = _table_structure_fingerprint(target_table)
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_properties_refusal(request, "table_properties_fingerprint_mismatch")

    props = request.properties_dict()

    if "style" in props and props["style"] is not None:
        style_refusal = _validate_referenced_table_style(graph, props["style"])
        if style_refusal is not None:
            return _table_properties_refusal(request, style_refusal)

    final_floating: dict[str, Any] | None = None
    if "floating" in props and props["floating"] is not None:
        existing_floating = (
            _floating_positioning_from_element(tbl_pr, namespace) if tbl_pr is not None else {}
        )
        final_floating = {**existing_floating, **props["floating"]}
        for axis in ("horizontal", "vertical"):
            offset_key = f"{axis}_offset"
            alignment_key = f"{axis}_alignment"
            if offset_key in props["floating"] and alignment_key not in props["floating"]:
                final_floating.pop(alignment_key, None)
            if alignment_key in final_floating:
                # OOXML alignment takes precedence over the numeric position.
                final_floating.pop(offset_key, None)
        required_keys = _FLOATING_REQUIRED_KEYS - {"horizontal_offset", "vertical_offset"}
        if not required_keys.issubset(final_floating) or any(
            f"{axis}_offset" not in final_floating and f"{axis}_alignment" not in final_floating
            for axis in ("horizontal", "vertical")
        ):
            return _table_properties_refusal(request, "table_properties_floating_incomplete")
        try:
            final_floating = _validate_floating_positioning(final_floating)
        except (TypeError, ValueError):
            return _table_properties_refusal(request, "table_properties_floating_incomplete")

    properties_before = _inspect_table_properties(tbl_pr, namespace)
    original_floating_attributes = (
        dict(existing_floating.attrib)
        if tbl_pr is not None and (existing_floating := tbl_pr.find(word("tblpPr"))) is not None
        else None
    )

    if tbl_pr is None:
        tbl_pr = ET.Element(word("tblPr"))
        target_table.insert(0, tbl_pr)

    for key, local in _RESETTABLE_PROPERTY_ELEMENTS.items():
        if key in props and props[key] is None:
            for existing in list(tbl_pr.findall(word(local))):
                tbl_pr.remove(existing)
    for key, local in _SCALAR_PROPERTY_ELEMENTS.items():
        if key not in props:
            continue
        value = props[key]
        if value is None:
            for existing in list(tbl_pr.findall(word(local))):
                tbl_pr.remove(existing)
        else:
            element = ET.Element(word(local))
            xml_value = ("1" if value else "0") if key == "bidi_visual" else str(value)
            element.set(word("val"), xml_value)
            _insert_or_replace_tblpr_child(tbl_pr, element, namespace)

    if "floating" in props:
        desired_floating = props["floating"]
        if desired_floating is None:
            existing_floating = tbl_pr.find(word("tblpPr"))
            if existing_floating is not None:
                tbl_pr.remove(existing_floating)
        else:
            existing_floating = tbl_pr.find(word("tblpPr"))
            tblp_pr = ET.Element(word("tblpPr"))
            if existing_floating is not None:
                for attr_name, attr_value in existing_floating.attrib.items():
                    tblp_pr.set(attr_name, attr_value)
            assert final_floating is not None
            for key, xml_attr in _FLOATING_KEY_TO_XML_ATTR.items():
                if key in final_floating:
                    tblp_pr.set(word(xml_attr), str(final_floating[key]))
                else:
                    tblp_pr.attrib.pop(word(xml_attr), None)
            _insert_or_replace_tblpr_child(tbl_pr, tblp_pr, namespace)

    if "style" in props:
        desired_style = props["style"]
        if desired_style is None:
            existing_style = tbl_pr.find(word("tblStyle"))
            if existing_style is not None:
                tbl_pr.remove(existing_style)
        else:
            tbl_style = ET.Element(word("tblStyle"))
            tbl_style.set(word("val"), desired_style)
            _insert_or_replace_tblpr_child(tbl_pr, tbl_style, namespace)
    if "width" in props and props["width"] is not None:
        w_spec = props["width"]
        tbl_w = ET.Element(word("tblW"))
        tbl_w.set(word("w"), str(w_spec["value"]))
        tbl_w.set(word("type"), w_spec["type"])
        _insert_or_replace_tblpr_child(tbl_pr, tbl_w, namespace)

    if "alignment" in props and props["alignment"] is not None:
        align_val = props["alignment"]
        jc = ET.Element(word("jc"))
        jc.set(word("val"), align_val)
        _insert_or_replace_tblpr_child(tbl_pr, jc, namespace)
    if "cell_spacing" in props:
        desired_cs = props["cell_spacing"]
        if desired_cs is None:
            existing_cs = tbl_pr.find(word("tblCellSpacing"))
            if existing_cs is not None:
                tbl_pr.remove(existing_cs)
        else:
            tbl_cs = ET.Element(word("tblCellSpacing"))
            tbl_cs.set(word("w"), str(desired_cs["value"]))
            tbl_cs.set(word("type"), desired_cs["type"])
            _insert_or_replace_tblpr_child(tbl_pr, tbl_cs, namespace)

    if "indent" in props:
        desired_ind = props["indent"]
        if desired_ind is None:
            existing_ind = tbl_pr.find(word("tblInd"))
            if existing_ind is not None:
                tbl_pr.remove(existing_ind)
        else:
            tbl_ind = ET.Element(word("tblInd"))
            tbl_ind.set(word("w"), str(desired_ind["value"]))
            tbl_ind.set(word("type"), desired_ind["type"])
            _insert_or_replace_tblpr_child(tbl_pr, tbl_ind, namespace)

    if "borders" in props and props["borders"] is not None:
        borders_spec = props["borders"]
        existing_borders = tbl_pr.find(word("tblBorders"))
        existing_sides: dict[str, ET.Element] = {}
        if existing_borders is not None:
            for c in existing_borders:
                xml_side = c.tag.rsplit("}", 1)[-1]
                side = _BORDER_XML_SIDE_TO_PUBLIC.get(xml_side, xml_side)
                if side not in existing_sides or xml_side in _BORDER_XML_SIDE_TO_PUBLIC:
                    existing_sides[side] = c

        new_borders = ET.Element(word("tblBorders"))
        for side in _TBL_BORDERS_ORDER:
            xml_side = (
                {"left": "start", "right": "end"}.get(side, side)
                if namespace == _STRICT_WORD_NAMESPACE
                else side
            )
            if side in borders_spec:
                spec = borders_spec[side]
                b_elem = ET.Element(word(xml_side))
                b_elem.set(word("val"), spec["val"])
                if spec.get("sz") is not None:
                    b_elem.set(word("sz"), str(spec["sz"]))
                if spec.get("space") is not None:
                    b_elem.set(word("space"), str(spec["space"]))
                if spec.get("color") is not None:
                    b_elem.set(word("color"), spec["color"])
                new_borders.append(b_elem)
            elif side in existing_sides:
                b_elem = copy.deepcopy(existing_sides[side])
                b_elem.tag = word(xml_side)
                new_borders.append(b_elem)
        for side, elem in existing_sides.items():
            if side not in _TBL_BORDERS_ORDER:
                new_borders.append(copy.deepcopy(elem))
        _insert_or_replace_tblpr_child(tbl_pr, new_borders, namespace)

    if "shading" in props and props["shading"] is not None:
        shd_spec = props["shading"]
        shd = ET.Element(word("shd"))
        shd.set(word("val"), shd_spec.get("val", "clear"))
        shd.set(word("color"), shd_spec.get("color", "auto"))
        shd.set(word("fill"), shd_spec["fill"])
        _insert_or_replace_tblpr_child(tbl_pr, shd, namespace)

    if "cell_margins" in props and props["cell_margins"] is not None:
        mar_spec = props["cell_margins"]
        existing_mar = tbl_pr.find(word("tblCellMar"))
        existing_mar_sides: dict[str, ET.Element] = {}
        if existing_mar is not None:
            for c in existing_mar:
                local_side = c.tag.rsplit("}", 1)[-1]
                existing_mar_sides[local_side] = c

        new_mar = ET.Element(word("tblCellMar"))
        for side in _TBL_CELL_MAR_ORDER:
            if side in mar_spec:
                m_elem = ET.Element(word(side))
                m_elem.set(word("w"), str(mar_spec[side]))
                m_elem.set(word("type"), "dxa")
                new_mar.append(m_elem)
            elif side in existing_mar_sides:
                new_mar.append(copy.deepcopy(existing_mar_sides[side]))
        for side, elem in existing_mar_sides.items():
            if side not in _TBL_CELL_MAR_ORDER:
                new_mar.append(copy.deepcopy(elem))
        _insert_or_replace_tblpr_child(tbl_pr, new_mar, namespace)
    if "layout" in props:
        desired_layout = props["layout"]
        if desired_layout is None:
            existing_layout = tbl_pr.find(word("tblLayout"))
            if existing_layout is not None:
                tbl_pr.remove(existing_layout)
        else:
            tbl_layout = ET.Element(word("tblLayout"))
            tbl_layout.set(word("type"), desired_layout["type"])
            _insert_or_replace_tblpr_child(tbl_pr, tbl_layout, namespace)

    if "look" in props:
        desired_look = props["look"]
        if desired_look is None:
            existing_look = tbl_pr.find(word("tblLook"))
            if existing_look is not None:
                tbl_pr.remove(existing_look)
        else:
            existing_look = tbl_pr.find(word("tblLook"))
            existing_flags: dict[str, bool] = {}
            existing_val_int: int | None = None
            if existing_look is not None:
                v_attr = existing_look.get(word("val"))
                if v_attr is not None and _LOOK_HEX_PATTERN.match(v_attr):
                    existing_val_int = int(v_attr, 16)
                for flag_name, xml_attr in _LOOK_FLAG_TO_XML_ATTR.items():
                    raw_attr = existing_look.get(word(xml_attr))
                    if raw_attr is not None:
                        existing_flags[flag_name] = raw_attr.lower() in (
                            "1",
                            "true",
                            "on",
                        )
                    elif existing_val_int is not None:
                        existing_flags[flag_name] = bool(
                            existing_val_int & _LOOK_FLAG_TO_BIT[flag_name]
                        )
                    else:
                        existing_flags[flag_name] = False

            val_decoded_flags: dict[str, bool] = {}
            if "val" in desired_look:
                val_decoded_flags = _decode_look_bitmask(int(desired_look["val"], 16))

            final_flags: dict[str, bool] = {}
            for flag_name in _LOOK_FLAG_TO_BIT:
                if flag_name in desired_look:
                    final_flags[flag_name] = desired_look[flag_name]
                elif flag_name in val_decoded_flags:
                    final_flags[flag_name] = val_decoded_flags[flag_name]
                elif flag_name in existing_flags:
                    final_flags[flag_name] = existing_flags[flag_name]
                else:
                    final_flags[flag_name] = False

            base_mask = existing_val_int if existing_val_int is not None else 0
            computed_mask = _encode_look_bitmask(final_flags, base_mask)
            final_val_str = f"{computed_mask:04X}"

            new_look = ET.Element(word("tblLook"))
            if existing_look is not None:
                for attr_key, attr_val in existing_look.attrib.items():
                    new_look.set(attr_key, attr_val)

            if existing_look is not None:
                if existing_look.get(word("val")) is not None or "val" in desired_look:
                    new_look.set(word("val"), final_val_str)
            else:
                new_look.set(word("val"), final_val_str)

            for flag_name, xml_attr in _LOOK_FLAG_TO_XML_ATTR.items():
                new_look.set(word(xml_attr), "1" if final_flags[flag_name] else "0")

            _insert_or_replace_tblpr_child(tbl_pr, new_look, namespace)

    properties_after = _inspect_table_properties(tbl_pr, namespace)
    rows = target_table.findall(word("tr"))
    tbl_grid = target_table.find(word("tblGrid"))
    grid_cols = tbl_grid.findall(word("gridCol")) if tbl_grid is not None else []
    cols_count = (
        len(grid_cols) if grid_cols else max((len(r.findall(word("tc"))) for r in rows), default=0)
    )
    style_elem = tbl_pr.find(word("tblStyle"))
    table_style = style_elem.get(word("val")) if style_elem is not None else None

    target_change_report: dict[str, Any] = {
        "table": request.table,
        "part": request.part,
        "properties_authored": sorted(props.keys()),
        "properties_before": properties_before,
        "properties_after": properties_after,
        "style": table_style,
        "rows_count": len(rows),
        "columns_count": cols_count,
        "boundaries": {
            "table": request.table,
            "rows": len(rows),
            "columns": cols_count,
        },
    }

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({request.part})):
            return "table_properties_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_properties_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(request.part))
            s_tables = list(s_root.iter(word("tbl")))
            if len(s_tables) != len(tables):
                return "table_properties_staged_verification_failed"
            s_table = s_tables[request.table - 1]
            s_tbl_pr = s_table.find(word("tblPr"))
            if s_tbl_pr is None:
                return "table_properties_staged_verification_failed"
            seen_ranks: list[int] = []
            for child in s_tbl_pr:
                local = child.tag.rsplit("}", 1)[-1]
                if local in _TBLPR_CHILD_ORDER:
                    seen_ranks.append(_TBLPR_CHILD_ORDER.index(local))
            orig_tbl_pr = target_table.find(word("tblPr"))
            orig_ranks: list[int] = []
            if orig_tbl_pr is not None:
                for child in orig_tbl_pr:
                    local = child.tag.rsplit("}", 1)[-1]
                    if local in _TBLPR_CHILD_ORDER:
                        orig_ranks.append(_TBLPR_CHILD_ORDER.index(local))
            if orig_ranks == sorted(orig_ranks) and seen_ranks != sorted(seen_ranks):
                return "table_properties_staged_verification_failed"
            authored_tags = {
                local for key, local in _SCALAR_PROPERTY_ELEMENTS.items() if key in props
            }
            if "style" in props:
                authored_tags.add("tblStyle")
            if "width" in props:
                authored_tags.add("tblW")
            if "alignment" in props:
                authored_tags.add("jc")
            if "cell_spacing" in props:
                authored_tags.add("tblCellSpacing")
            if "indent" in props:
                authored_tags.add("tblInd")
            if "borders" in props:
                authored_tags.add("tblBorders")
            if "shading" in props:
                authored_tags.add("shd")
            if "cell_margins" in props:
                authored_tags.add("tblCellMar")
            if "layout" in props:
                authored_tags.add("tblLayout")
            if "look" in props:
                authored_tags.add("tblLook")
            if "floating" in props:
                authored_tags.add("tblpPr")
            authored_ranks = [
                _TBLPR_CHILD_ORDER.index(child.tag.rsplit("}", 1)[-1])
                for child in s_tbl_pr
                if child.tag.rsplit("}", 1)[-1] in authored_tags
            ]
            if authored_ranks != sorted(authored_ranks):
                return "table_properties_staged_verification_failed"
            for key, local in _SCALAR_PROPERTY_ELEMENTS.items():
                if key not in props:
                    continue
                staged_element = s_tbl_pr.find(word(local))
                value = props[key]
                if value is None:
                    if staged_element is not None:
                        return "table_properties_staged_verification_failed"
                else:
                    expected = ("1" if value else "0") if key == "bidi_visual" else str(value)
                    if staged_element is None or staged_element.get(word("val")) != expected:
                        return "table_properties_staged_verification_failed"
            for key, local in _RESETTABLE_PROPERTY_ELEMENTS.items():
                if key in props and props[key] is None and s_tbl_pr.find(word(local)) is not None:
                    return "table_properties_staged_verification_failed"
            if "floating" not in props:
                staged_floating = s_tbl_pr.find(word("tblpPr"))
                if original_floating_attributes is None:
                    if staged_floating is not None:
                        return "table_properties_staged_verification_failed"
                elif (
                    staged_floating is None
                    or staged_floating.attrib != original_floating_attributes
                ):
                    return "table_properties_staged_verification_failed"
            if orig_tbl_pr is not None:
                for check_local in ("tblInd", "tblCellSpacing", "tblLayout", "tblLook"):
                    prop_key = {
                        "tblInd": "indent",
                        "tblCellSpacing": "cell_spacing",
                        "tblLayout": "layout",
                        "tblLook": "look",
                    }[check_local]
                    if prop_key not in props:
                        orig_elem = orig_tbl_pr.find(word(check_local))
                        s_elem = s_tbl_pr.find(word(check_local))
                        if (orig_elem is None) != (s_elem is None):
                            return "table_properties_staged_verification_failed"
                        if (
                            orig_elem is not None
                            and s_elem is not None
                            and (orig_elem.attrib != s_elem.attrib or orig_elem.text != s_elem.text)
                        ):
                            return "table_properties_staged_verification_failed"
            if len(s_table.findall(word("tr"))) != len(rows):
                return "table_properties_staged_verification_failed"
            s_grid = s_table.find(word("tblGrid"))
            orig_grid = target_table.find(word("tblGrid"))
            if (orig_grid is None) != (s_grid is None):
                return "table_properties_staged_verification_failed"
            if (
                orig_grid is not None
                and s_grid is not None
                and ET.tostring(orig_grid) != ET.tostring(s_grid)
            ):
                return "table_properties_staged_verification_failed"
            if table_style is not None:
                s_style = s_tbl_pr.find(word("tblStyle"))
                if s_style is None or s_style.get(word("val")) != table_style:
                    return "table_properties_staged_verification_failed"
            elif "style" in props:
                s_style = s_tbl_pr.find(word("tblStyle"))
                if s_style is not None:
                    return "table_properties_staged_verification_failed"
            if "floating" in props:
                staged_floating = s_tbl_pr.find(word("tblpPr"))
                if props["floating"] is None:
                    if staged_floating is not None:
                        return "table_properties_staged_verification_failed"
                else:
                    if staged_floating is None or final_floating is None:
                        return "table_properties_staged_verification_failed"
                    for key, xml_attr in _FLOATING_KEY_TO_XML_ATTR.items():
                        expected = str(final_floating[key]) if key in final_floating else None
                        if staged_floating.get(word(xml_attr)) != expected:
                            return "table_properties_staged_verification_failed"
            if "indent" in props:
                s_ind = s_tbl_pr.find(word("tblInd"))
                if props["indent"] is None:
                    if s_ind is not None:
                        return "table_properties_staged_verification_failed"
                else:
                    if (
                        s_ind is None
                        or s_ind.get(word("w")) != str(props["indent"]["value"])
                        or (s_ind.get(word("type")) or "dxa") != props["indent"]["type"]
                    ):
                        return "table_properties_staged_verification_failed"
            if "cell_spacing" in props:
                s_cs = s_tbl_pr.find(word("tblCellSpacing"))
                if props["cell_spacing"] is None:
                    if s_cs is not None:
                        return "table_properties_staged_verification_failed"
                else:
                    if (
                        s_cs is None
                        or s_cs.get(word("w")) != str(props["cell_spacing"]["value"])
                        or (s_cs.get(word("type")) or "dxa") != props["cell_spacing"]["type"]
                    ):
                        return "table_properties_staged_verification_failed"
            if "layout" in props:
                s_layout = s_tbl_pr.find(word("tblLayout"))
                if props["layout"] is None:
                    if s_layout is not None:
                        return "table_properties_staged_verification_failed"
                else:
                    if s_layout is None or s_layout.get(word("type")) != props["layout"]["type"]:
                        return "table_properties_staged_verification_failed"
            if "look" in props:
                s_look = s_tbl_pr.find(word("tblLook"))
                if props["look"] is None:
                    if s_look is not None:
                        return "table_properties_staged_verification_failed"
                else:
                    if s_look is None:
                        return "table_properties_staged_verification_failed"
                    for flag_name, xml_attr in _LOOK_FLAG_TO_XML_ATTR.items():
                        if flag_name in props["look"]:
                            expected_flag = props["look"][flag_name]
                            actual_attr = s_look.get(word(xml_attr))
                            if (
                                actual_attr is None
                                or (actual_attr.lower() in ("1", "true", "on")) != expected_flag
                            ):
                                return "table_properties_staged_verification_failed"
                    if "val" in props["look"] and s_look.get(word("val")) != props["look"]["val"]:
                        return "table_properties_staged_verification_failed"
        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return "table_properties_staged_verification_failed"
        return None

    serialized = serialize_xml_with_preserved_qnames(root)
    replacements = {request.part: serialized}

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=[target_change_report],
        verify=verify_publication,
        failure_code="table_properties_staged_verification_failed",
        exists_code="table_properties_destination_occupied",
        unavailable_code="table_properties_publication_unavailable",
    )

    if not published.passed:
        return _table_properties_refusal(
            request, published.code or "table_properties_staged_verification_failed"
        )

    return TablePropertiesReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        properties=props,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
