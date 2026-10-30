"""Source-bound cell properties: use dict nulls to remove direct properties or border/margin sides."""

from __future__ import annotations

import copy
import enum
import json
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
from .table_merge import _FORBIDDEN_CELL_ELEMENTS
from .table_properties import (
    _HEX_COLOR_PATTERN,
    _VALID_BORDER_STYLES,
    BorderSpec,
    TableBordersSpec,
    TableCellMarginsSpec,
    TableShadingSpec,
    TableWidthSpec,
    _check_reachable_part,
    _unrelated_parts_preserved,
)

__all__ = [
    "BorderSpec",
    "TableBordersSpec",
    "TableCellMarginsSpec",
    "TableCellPropertiesReport",
    "TableCellPropertiesRequest",
    "TableCellPropertiesSpec",
    "TableCellPropertiesTarget",
    "TableShadingSpec",
    "TableWidthSpec",
    "author_table_cell_properties",
    "discover_table_cell_properties_targets",
    "load_table_cell_properties_request",
    "parse_table_cell_properties_request",
    "table_cell_properties_structure_fingerprint",
]

_REQUIRED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "source_sha256",
        "destination",
        "target",
        "structure_fingerprint_sha256",
        "properties",
    }
)
_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "source_sha256",
        "destination",
        "target",
        "structure_fingerprint_sha256",
        "properties",
        "part",
    }
)

_REQUIRED_TARGET_FIELDS = frozenset({"table", "row", "cell"})
_ALLOWED_TARGET_FIELDS = frozenset({"table", "row", "cell", "part"})

_ALLOWED_PROPERTY_KEYS = frozenset(
    {
        "width",
        "vertical_alignment",
        "vAlign",
        "shading",
        "borders",
        "cell_margins",
        "no_wrap",
        "noWrap",
        "text_direction",
        "textDirection",
        "fit_text",
        "hide_mark",
        "cnf_style",
    }
)

# The style-authoring caller shares validation, but does not support direct-cell edits.
_STYLE_PROPERTY_KEYS = _ALLOWED_PROPERTY_KEYS - {"fit_text", "hide_mark", "cnf_style"}
_DIRECT_CELL_PROPERTY_TAGS = {
    "width": "tcW",
    "vertical_alignment": "vAlign",
    "shading": "shd",
    "borders": "tcBorders",
    "cell_margins": "tcMar",
    "no_wrap": "noWrap",
    "text_direction": "textDirection",
    "fit_text": "tcFitText",
    "hide_mark": "hideMark",
    "cnf_style": "cnfStyle",
}
_CNF_STYLE_FLAGS = frozenset(
    {
        "firstRow",
        "lastRow",
        "firstColumn",
        "lastColumn",
        "oddVBand",
        "evenVBand",
        "oddHBand",
        "evenHBand",
        "firstRowFirstColumn",
        "firstRowLastColumn",
        "lastRowFirstColumn",
        "lastRowLastColumn",
    }
)

_VALID_VERTICAL_ALIGNMENTS = frozenset({"top", "center", "both", "bottom"})

# Case-folded spellings accepted for ST_VerticalJc: the literals themselves, the python-docx
# ``WD_CELL_VERTICAL_ALIGNMENT`` member names, and Word UI's "middle" for ``center``.
_VERTICAL_ALIGNMENT_ALIASES: dict[str, str] = {
    **{value: value for value in _VALID_VERTICAL_ALIGNMENTS},
    "middle": "center",
}

# Properties that may be authored on the origin (``w:vMerge w:val="restart"``) cell of a
# vertical merge. Word draws the merged region from the origin cell's tcPr, so these apply to
# the whole region. ``tcW`` stays refused: continuation cells carry their own widths.
_VERTICAL_MERGE_ORIGIN_PROPERTIES = frozenset(
    {"vertical_alignment", "shading", "borders", "cell_margins", "no_wrap", "text_direction"}
)

# ECMA-376 / ISO-IEC 29500-1 §17.18.93 ST_TextDirection
_VALID_TEXT_DIRECTIONS = frozenset({"lrTb", "tbRl", "btLr", "lrTbV", "tbRlV", "tbLrV"})

_ALLOWED_CELL_BORDER_SIDES = frozenset(
    {"top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl"}
)

_ALLOWED_CELL_MARGIN_SIDES = frozenset({"top", "left", "bottom", "right"})

# ECMA-376 / ISO-IEC 29500-1 §17.4.70 CT_TcPr child element sequence order
_TCPR_CHILD_ORDER: tuple[str, ...] = (
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
)

# ECMA-376 / ISO-IEC 29500-1 §17.4.67 CT_TcBorders child element sequence order
_TC_BORDERS_ORDER: tuple[str, ...] = (
    "top",
    "left",
    "start",
    "bottom",
    "right",
    "end",
    "insideH",
    "insideV",
    "tl2br",
    "tr2bl",
)

# ECMA-376 / ISO-IEC 29500-1 §17.4.70 CT_TcMar child element sequence order
_TC_CELL_MAR_ORDER: tuple[str, ...] = (
    "top",
    "left",
    "bottom",
    "right",
)


def _normalize_vertical_alignment(value: Any) -> str:
    """Resolve an ST_VerticalJc literal, python-docx member, or alias to its canonical value."""
    if isinstance(value, enum.Enum):
        value = value.name
    canonical = _VERTICAL_ALIGNMENT_ALIASES.get(value.lower()) if isinstance(value, str) else None
    if canonical is None:
        raise ValueError("table_cell_properties_vertical_alignment_invalid")
    return canonical


def _vertical_merge_origin(tc_pr: ET.Element | None, namespace: str) -> bool:
    """True when tcPr declares exactly one ``w:vMerge w:val="restart"`` (a merge origin)."""
    if tc_pr is None:
        return False
    v_merges = tc_pr.findall(Q(namespace, "vMerge"))
    return len(v_merges) == 1 and v_merges[0].get(Q(namespace, "val")) == "restart"


def _insert_or_replace_tcpr_child(tc_pr: ET.Element, new_child: ET.Element, namespace: str) -> None:
    """Insert or replace a child element in tcPr while strictly maintaining schema order."""
    local_name = new_child.tag.rsplit("}", 1)[-1]
    def word(local):
        return Q(namespace, local)
    for existing in list(tc_pr.findall(word(local_name))):
        tc_pr.remove(existing)
    try:
        new_order = _TCPR_CHILD_ORDER.index(local_name)
    except ValueError:
        new_order = 999
    insert_idx = len(tc_pr)
    for idx, child in enumerate(tc_pr):
        child_local = child.tag.rsplit("}", 1)[-1]
        try:
            child_order = _TCPR_CHILD_ORDER.index(child_local)
        except ValueError:
            child_order = 999
        if child_order > new_order:
            insert_idx = idx
            break
    tc_pr.insert(insert_idx, new_child)


@dataclass(frozen=True)
class TableCellPropertiesSpec:
    width: TableWidthSpec | None = None
    vertical_alignment: str | None = None
    shading: TableShadingSpec | None = None
    borders: TableBordersSpec | None = None
    cell_margins: TableCellMarginsSpec | None = None
    no_wrap: bool | None = None
    text_direction: str | None = None
    fit_text: bool | None = None
    hide_mark: bool | None = None
    cnf_style: Mapping[str, Any] | None = None

    def __post_init__(self) -> None:
        if self.text_direction is not None and (
            not isinstance(self.text_direction, str)
            or self.text_direction not in _VALID_TEXT_DIRECTIONS
        ):
            raise ValueError("table_cell_properties_text_direction_invalid")

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        if self.width is not None:
            data["width"] = self.width.to_dict()
        if self.vertical_alignment is not None:
            data["vertical_alignment"] = self.vertical_alignment
        if self.shading is not None:
            data["shading"] = self.shading.to_dict()
        if self.borders is not None:
            data["borders"] = self.borders.to_dict()
        if self.cell_margins is not None:
            data["cell_margins"] = self.cell_margins.to_dict()
        if self.no_wrap is not None:
            data["no_wrap"] = self.no_wrap
        if self.text_direction is not None:
            data["text_direction"] = self.text_direction
        if self.fit_text is not None:
            data["fit_text"] = self.fit_text
        if self.hide_mark is not None:
            data["hide_mark"] = self.hide_mark
        if self.cnf_style is not None:
            data["cnf_style"] = dict(self.cnf_style)
        return data


def _validate_table_cell_properties_target_dict(target: Any) -> dict[str, Any]:
    if not isinstance(target, Mapping):
        raise TypeError("table_cell_properties_target_invalid")
    keys = set(target.keys())
    if not _REQUIRED_TARGET_FIELDS.issubset(keys):
        raise ValueError("table_cell_properties_target_missing_fields")
    if not keys.issubset(_ALLOWED_TARGET_FIELDS):
        raise ValueError("table_cell_properties_target_extra_fields")

    for field in ("table", "row", "cell"):
        val = target[field]
        if not isinstance(val, int) or isinstance(val, bool) or val < 1:
            raise ValueError(f"table_cell_properties_target_{field}_invalid")

    part = target.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("table_cell_properties_target_part_invalid")

    return {
        "table": target["table"],
        "row": target["row"],
        "cell": target["cell"],
        "part": part,
    }


def _validate_table_cell_properties_dict(
    properties: Any, *, direct_cell: bool = False
) -> dict[str, Any]:
    if isinstance(properties, TableCellPropertiesSpec):
        properties = properties.to_dict()
    elif not isinstance(properties, Mapping):
        raise TypeError("table_cell_properties_properties_invalid")

    keys = set(properties.keys())
    if not keys:
        raise ValueError("table_cell_properties_properties_empty")
    if not keys.issubset(_ALLOWED_PROPERTY_KEYS if direct_cell else _STYLE_PROPERTY_KEYS):
        raise ValueError("table_cell_properties_properties_extra_fields")

    result: dict[str, Any] = {}

    # Normalize vAlign vs vertical_alignment; both keys must name the same alignment
    v_align_val: Any = None
    if "vertical_alignment" in properties and "vAlign" in properties:
        v_align_val = properties["vertical_alignment"]
        other = properties["vAlign"]
        if (v_align_val is None) != (other is None) or (
            v_align_val is not None
            and _normalize_vertical_alignment(v_align_val) != _normalize_vertical_alignment(other)
        ):
            raise ValueError("table_cell_properties_vertical_alignment_invalid")
    elif "vertical_alignment" in properties:
        v_align_val = properties["vertical_alignment"]
    elif "vAlign" in properties:
        v_align_val = properties["vAlign"]

    if v_align_val is not None:
        result["vertical_alignment"] = _normalize_vertical_alignment(v_align_val)
    elif direct_cell and keys & {"vertical_alignment", "vAlign"}:
        result["vertical_alignment"] = None

    # Normalize no_wrap vs noWrap
    no_wrap_val: Any = None
    if "no_wrap" in properties and "noWrap" in properties:
        if properties["no_wrap"] != properties["noWrap"]:
            raise ValueError("table_cell_properties_no_wrap_invalid")
        no_wrap_val = properties["no_wrap"]
    elif "no_wrap" in properties:
        no_wrap_val = properties["no_wrap"]
    elif "noWrap" in properties:
        no_wrap_val = properties["noWrap"]

    if no_wrap_val is not None:
        if not isinstance(no_wrap_val, bool):
            raise ValueError("table_cell_properties_no_wrap_invalid")
        result["no_wrap"] = no_wrap_val
    elif direct_cell and keys & {"no_wrap", "noWrap"}:
        result["no_wrap"] = None

    # Normalize text_direction vs textDirection
    text_direction_val: Any = None
    if "text_direction" in properties and "textDirection" in properties:
        if properties["text_direction"] != properties["textDirection"]:
            raise ValueError("table_cell_properties_text_direction_invalid")
        text_direction_val = properties["text_direction"]
    elif "text_direction" in properties:
        text_direction_val = properties["text_direction"]
    elif "textDirection" in properties:
        text_direction_val = properties["textDirection"]

    if text_direction_val is not None:
        if (
            not isinstance(text_direction_val, str)
            or text_direction_val not in _VALID_TEXT_DIRECTIONS
        ):
            raise ValueError("table_cell_properties_text_direction_invalid")
        result["text_direction"] = text_direction_val
    elif direct_cell and keys & {"text_direction", "textDirection"}:
        result["text_direction"] = None

    for key, val in properties.items():
        if key in (
            "vertical_alignment",
            "vAlign",
            "no_wrap",
            "noWrap",
            "text_direction",
            "textDirection",
        ):
            continue
        if direct_cell and val is None:
            result[key] = None
            continue
        if key in ("fit_text", "hide_mark"):
            if not isinstance(val, bool):
                raise ValueError(f"table_cell_properties_{key}_invalid")
            result[key] = val
            continue
        if key == "cnf_style":
            if not isinstance(val, Mapping) or not val:
                raise ValueError("table_cell_properties_cnf_style_invalid")
            if not set(val).issubset(_CNF_STYLE_FLAGS | {"val"}):
                raise ValueError("table_cell_properties_cnf_style_extra_fields")
            for attr, attr_val in val.items():
                if attr == "val":
                    if (
                        not isinstance(attr_val, str)
                        or len(attr_val) != 12
                        or set(attr_val) - {"0", "1"}
                    ):
                        raise ValueError("table_cell_properties_cnf_style_val_invalid")
                elif not isinstance(attr_val, bool):
                    raise ValueError(f"table_cell_properties_cnf_style_{attr}_invalid")
            result[key] = dict(val)
            continue
        if key == "width":
            if isinstance(val, TableWidthSpec):
                result["width"] = val.to_dict()
            elif isinstance(val, Mapping):
                w_keys = set(val.keys())
                if "value" not in w_keys:
                    raise ValueError("table_cell_properties_width_missing_value")
                if not w_keys.issubset({"value", "type"}):
                    raise ValueError("table_cell_properties_width_extra_fields")
                v = val["value"]
                if not isinstance(v, int) or isinstance(v, bool) or v < 0:
                    raise ValueError("table_cell_properties_width_value_invalid")
                t = val.get("type", "dxa")
                if t not in {"dxa", "auto", "pct", "nil"}:
                    raise ValueError("table_cell_properties_width_type_invalid")
                result["width"] = {"value": v, "type": t}
            else:
                raise TypeError("table_cell_properties_width_invalid")

        elif key == "borders":
            if isinstance(val, TableBordersSpec):
                result["borders"] = val.to_dict()
            elif isinstance(val, Mapping):
                b_keys = set(val.keys())
                if not b_keys:
                    raise ValueError("table_cell_properties_borders_empty")
                sides = _ALLOWED_CELL_BORDER_SIDES
                if not direct_cell:
                    sides = sides - {"start", "end"}
                if not b_keys.issubset(sides):
                    raise ValueError("table_cell_properties_borders_extra_fields")
                b_dict: dict[str, Any] = {}
                for side, side_val in val.items():
                    if direct_cell and side_val is None:
                        b_dict[side] = None
                        continue
                    if isinstance(side_val, BorderSpec):
                        b_dict[side] = side_val.to_dict()
                    elif isinstance(side_val, Mapping):
                        s_keys = set(side_val.keys())
                        if "val" not in s_keys:
                            raise ValueError("table_cell_properties_border_missing_val")
                        if not s_keys.issubset({"val", "sz", "space", "color"}):
                            raise ValueError("table_cell_properties_border_extra_fields")
                        b_style = side_val["val"]
                        if not isinstance(b_style, str) or b_style not in _VALID_BORDER_STYLES:
                            raise ValueError("table_cell_properties_border_val_invalid")
                        entry: dict[str, Any] = {"val": b_style}
                        if "sz" in side_val:
                            sz = side_val["sz"]
                            if sz is not None and (
                                not isinstance(sz, int) or isinstance(sz, bool) or sz < 0
                            ):
                                raise ValueError("table_cell_properties_border_sz_invalid")
                            if sz is not None:
                                entry["sz"] = sz
                        if "space" in side_val:
                            sp = side_val["space"]
                            if sp is not None and (
                                not isinstance(sp, int) or isinstance(sp, bool) or sp < 0
                            ):
                                raise ValueError("table_cell_properties_border_space_invalid")
                            if sp is not None:
                                entry["space"] = sp
                        if "color" in side_val:
                            c = side_val["color"]
                            if c is not None and (
                                not isinstance(c, str)
                                or (c != "auto" and not _HEX_COLOR_PATTERN.match(c))
                            ):
                                raise ValueError("table_cell_properties_border_color_invalid")
                            if c is not None:
                                entry["color"] = c
                        b_dict[side] = entry
                    else:
                        raise TypeError("table_cell_properties_border_side_invalid")
                result["borders"] = b_dict
            else:
                raise TypeError("table_cell_properties_borders_invalid")

        elif key == "shading":
            if isinstance(val, TableShadingSpec):
                result["shading"] = val.to_dict()
            elif isinstance(val, Mapping):
                s_keys = set(val.keys())
                if "fill" not in s_keys:
                    raise ValueError("table_cell_properties_shading_missing_fill")
                if not s_keys.issubset({"fill", "val", "color"}):
                    raise ValueError("table_cell_properties_shading_extra_fields")
                fill = val["fill"]
                if not isinstance(fill, str) or (
                    fill != "auto" and not _HEX_COLOR_PATTERN.match(fill)
                ):
                    raise ValueError("table_cell_properties_shading_fill_invalid")
                s_val = val.get("val", "clear")
                if not isinstance(s_val, str):
                    raise TypeError("table_cell_properties_shading_val_invalid")
                s_color = val.get("color", "auto")
                if not isinstance(s_color, str):
                    raise TypeError("table_cell_properties_shading_color_invalid")
                result["shading"] = {"fill": fill, "val": s_val, "color": s_color}
            else:
                raise TypeError("table_cell_properties_shading_invalid")

        elif key == "cell_margins":
            if isinstance(val, TableCellMarginsSpec):
                result["cell_margins"] = val.to_dict()
            elif isinstance(val, Mapping):
                m_keys = set(val.keys())
                if not m_keys:
                    raise ValueError("table_cell_properties_cell_margins_empty")
                if not m_keys.issubset(_ALLOWED_CELL_MARGIN_SIDES):
                    raise ValueError("table_cell_properties_cell_margins_extra_fields")
                m_dict: dict[str, int | None] = {}
                for side in m_keys:
                    mv = val[side]
                    if direct_cell and mv is None:
                        m_dict[side] = None
                        continue
                    if not isinstance(mv, int) or isinstance(mv, bool) or mv < 0:
                        raise ValueError(f"table_cell_properties_cell_margin_{side}_invalid")
                    m_dict[side] = mv
                result["cell_margins"] = m_dict
            else:
                raise TypeError("table_cell_properties_cell_margins_invalid")

    if not result:
        raise ValueError("table_cell_properties_properties_empty")

    return result


@dataclass(frozen=True)
class TableCellPropertiesTarget:
    """One discovered candidate table cell target for table cell properties authoring."""

    table: int
    row: int
    cell: int
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    current_properties: dict[str, Any] | None = None

    @property
    def target(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "row": self.row,
            "cell": self.cell,
            "part": self.part,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "row": self.row,
            "cell": self.cell,
            "part": self.part,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "current_properties": self.current_properties,
        }


@dataclass(frozen=True)
class TableCellPropertiesRequest:
    """One source-bound request to author direct formatting properties on an exact table cell."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    properties: Mapping[str, Any] | TableCellPropertiesSpec

    def __post_init__(self) -> None:
        if isinstance(self.source, str):
            object.__setattr__(self, "source", Path(self.source))
        elif not isinstance(self.source, Path):
            raise TypeError("table_cell_properties_source_invalid")
        if not str(self.source):
            raise ValueError("table_cell_properties_source_invalid")

        if isinstance(self.destination, str):
            object.__setattr__(self, "destination", Path(self.destination))
        elif not isinstance(self.destination, Path):
            raise TypeError("table_cell_properties_destination_invalid")
        if not str(self.destination):
            raise ValueError("table_cell_properties_destination_invalid")

        _require_sha256(self.source_sha256, "table_cell_properties_source_sha256")
        _require_sha256(
            self.structure_fingerprint_sha256,
            "table_cell_properties_structure_fingerprint_sha256",
        )

        validated_target = _validate_table_cell_properties_target_dict(self.target)
        object.__setattr__(self, "target", validated_target)

        validated_properties = _validate_table_cell_properties_dict(
            self.properties, direct_cell=True
        )
        object.__setattr__(self, "properties", validated_properties)

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def row(self) -> int:
        return self.target["row"]

    @property
    def cell(self) -> int:
        return self.target["cell"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "row": self.row,
            "cell": self.cell,
            "part": self.part,
        }

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
            "part": self.part,
        }


@dataclass(frozen=True)
class TableCellPropertiesReport:
    """Content-safe public result for a table cell properties authoring request."""

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
            "operation": "table-cell-properties",
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


def _table_cell_properties_refusal(
    request: TableCellPropertiesRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TableCellPropertiesReport:
    return TableCellPropertiesReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        properties=request.properties_dict(),
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_cell_properties_request(value: Any) -> TableCellPropertiesRequest:
    """Parse and strictly validate a table cell properties authoring request."""
    if isinstance(value, TableCellPropertiesRequest):
        return value
    if not isinstance(value, Mapping):
        raise TypeError("table_cell_properties_request_invalid")

    keys = set(value.keys())
    if not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_cell_properties_request_missing_fields")
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS):
        raise ValueError("table_cell_properties_request_extra_fields")

    target_raw = value["target"]
    if not isinstance(target_raw, Mapping):
        raise TypeError("table_cell_properties_target_invalid")

    target_dict = dict(target_raw)
    if "part" in value and "part" not in target_dict:
        target_dict["part"] = value["part"]

    return TableCellPropertiesRequest(
        source=value["source"],
        source_sha256=value["source_sha256"],
        destination=value["destination"],
        target=target_dict,
        structure_fingerprint_sha256=value["structure_fingerprint_sha256"],
        properties=value["properties"],
    )


def load_table_cell_properties_request(path: str | Path) -> TableCellPropertiesRequest:
    """Load and strictly validate one JSON table-cell-properties request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_cell_properties_request(json.load(stream))


def _inspect_table_cell_properties(tc_pr: ET.Element | None, namespace: str) -> dict[str, Any]:
    """Extract direct formatting properties from w:tcPr."""
    if tc_pr is None:
        return {}

    def word(local):
        return Q(namespace, local)
    result: dict[str, Any] = {}

    tc_w = tc_pr.find(word("tcW"))
    if tc_w is not None:
        raw_w = tc_w.get(word("w"))
        w_val = int(raw_w) if raw_w is not None and raw_w.isdigit() else 0
        w_type = tc_w.get(word("type"), "dxa")
        result["width"] = {"value": w_val, "type": w_type}

    v_align = tc_pr.find(word("vAlign"))
    if v_align is not None:
        result["vertical_alignment"] = v_align.get(word("val"), "top")

    shd = tc_pr.find(word("shd"))
    if shd is not None:
        result["shading"] = {
            "fill": shd.get(word("fill"), "auto"),
            "val": shd.get(word("val"), "clear"),
            "color": shd.get(word("color"), "auto"),
        }

    tc_borders = tc_pr.find(word("tcBorders"))
    if tc_borders is not None:
        borders_dict: dict[str, Any] = {}
        for side in _TC_BORDERS_ORDER:
            b_elem = tc_borders.find(word(side))
            if b_elem is not None:
                side_dict: dict[str, Any] = {"val": b_elem.get(word("val"), "single")}
                sz = b_elem.get(word("sz"))
                if sz is not None and sz.isdigit():
                    side_dict["sz"] = int(sz)
                space = b_elem.get(word("space"))
                if space is not None and space.isdigit():
                    side_dict["space"] = int(space)
                color = b_elem.get(word("color"))
                if color is not None:
                    side_dict["color"] = color
                borders_dict[side] = side_dict
        if borders_dict:
            result["borders"] = borders_dict

    tc_mar = tc_pr.find(word("tcMar"))
    if tc_mar is not None:
        mar_dict: dict[str, int] = {}
        for side in _TC_CELL_MAR_ORDER:
            m_elem = tc_mar.find(word(side))
            if m_elem is not None:
                w_str = m_elem.get(word("w"))
                if w_str is not None and w_str.isdigit():
                    mar_dict[side] = int(w_str)
        if mar_dict:
            result["cell_margins"] = mar_dict
    no_wrap_elem = tc_pr.find(word("noWrap"))
    if no_wrap_elem is not None:
        val = no_wrap_elem.get(word("val")) or no_wrap_elem.get("val")
        result["no_wrap"] = val is None or val.strip().lower() not in {
            "0",
            "false",
            "off",
        }

    text_direction = tc_pr.find(word("textDirection"))
    if text_direction is not None:
        result["text_direction"] = text_direction.get(word("val"))

    for key, tag in (("fit_text", "tcFitText"), ("hide_mark", "hideMark")):
        element = tc_pr.find(word(tag))
        if element is not None:
            value = element.get(word("val"))
            result[key] = value is None or value.strip().lower() not in {"0", "false", "off"}
    cnf = tc_pr.find(word("cnfStyle"))
    if cnf is not None:
        attrs: dict[str, Any] = {}
        for attr in _CNF_STYLE_FLAGS | {"val"}:
            value = cnf.get(word(attr))
            if value is not None:
                attrs[attr] = value if attr == "val" else value.lower() not in {"0", "false", "off"}
        result["cnf_style"] = attrs

    return result


def _inspect_cell_context(
    cell: ET.Element,
    row: ET.Element,
    table: ET.Element,
    root: ET.Element,
    namespace: str,
    unsupported_context: dict[int, bool],
    parent_map: dict[ET.Element, ET.Element] | None = None,
) -> str | None:
    """Return refusal code if table cell is in unsupported context, else None."""
    def word(local):
        return Q(namespace, local)

    # 1. Unsupported context (e.g. inside w:sdt, w:ins, w:del, comments, headers/footers)
    if (
        unsupported_context.get(id(table), False)
        or unsupported_context.get(id(row), False)
        or unsupported_context.get(id(cell), False)
    ):
        return "table_cell_properties_context_unsupported"

    # 3. Check for row-level unsupported attributes
    if row.find(word("tblPrEx")) is not None:
        return "table_cell_properties_context_unsupported"
    tr_pr = row.find(word("trPr"))
    if tr_pr is not None and (
        tr_pr.find(word("gridBefore")) is not None or tr_pr.find(word("gridAfter")) is not None
    ):
        return "table_cell_properties_context_unsupported"

    # 4. Check for multiple tblPr on table
    if len(table.findall(word("tblPr"))) > 1:
        return "table_cell_properties_cell_ambiguous"

    # 5. Check for multiple tcPr on cell
    if len(cell.findall(word("tcPr"))) > 1:
        return "table_cell_properties_cell_ambiguous"

    # 7. Check for merged cell. The origin of a vertical merge is editable; continuation
    # cells (``w:vMerge`` without ``restart``) belong to the origin and stay protected.
    tc_pr = cell.find(word("tcPr"))
    if tc_pr is not None:
        v_merges = tc_pr.findall(word("vMerge"))
        if len(v_merges) > 1:
            return "table_cell_properties_cell_ambiguous"
        if v_merges and not _vertical_merge_origin(tc_pr, namespace):
            return "table_cell_properties_cell_merged"
        if tc_pr.find(word("hMerge")) is not None:
            return "table_cell_properties_cell_merged"
        if tc_pr.find(word("gridSpan")) is not None:
            return "table_cell_properties_cell_merged"
        # Duplicate property children inside tcPr
        for prop in _DIRECT_CELL_PROPERTY_TAGS.values():
            if len(tc_pr.findall(word(prop))) > 1:
                return "table_cell_properties_cell_ambiguous"

    # 8. Check forbidden elements inside cell
    for child in cell.iter():
        tag_local = child.tag.rsplit("}", 1)[-1]
        if tag_local in _FORBIDDEN_CELL_ELEMENTS:
            return "table_cell_properties_context_unsupported"

    # 9. Cell must have at least one paragraph
    if not cell.findall(word("p")):
        return "table_cell_properties_context_unsupported"

    return None


def table_cell_properties_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    row: int,
    cell: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one target table cell."""
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        valid, _ = _check_reachable_part(graph, part, "table_cell_properties")
        if not valid:
            return None
        root = ET.fromstring(graph.read_part_bytes(part))
        ns = _word_namespace(root)
        if ns is None:
            return None
        def word(local):
            return Q(ns, local)
        tables = list(root.iter(word("tbl")))
        if table < 1 or table > len(tables):
            return None
        target_table = tables[table - 1]
        rows = target_table.findall(word("tr"))
        if row < 1 or row > len(rows):
            return None
        target_row = rows[row - 1]
        cells = target_row.findall(word("tc"))
        if cell < 1 or cell > len(cells):
            return None
        target_cell = cells[cell - 1]

        parent_map = {c: p for p in root.iter() for c in p}
        unsupported_context = _unsupported_context_by_element(root, ns)
        if (
            _inspect_cell_context(
                target_cell,
                target_row,
                target_table,
                root,
                ns,
                unsupported_context,
                parent_map,
            )
            is not None
        ):
            return None
        return _cell_structure_fingerprint(target_cell)
    except (PackageError, OSError, ET.ParseError, KeyError):
        return None


def discover_table_cell_properties_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table cell targets for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-cell-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_cell_properties_source_missing",
            "source_sha256": None,
            "part": part,
            "targets": [],
            "unsupported": [],
        }

    try:
        source_sha256 = _sha256_file(path)
        graph = PackageGraph.from_file(path)
    except (PackageError, OSError, zipfile.BadZipFile):
        return {
            "schema_version": 1,
            "operation": "table-cell-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_cell_properties_package_invalid",
            "source_sha256": None,
            "part": part,
            "targets": [],
            "unsupported": [],
        }

    valid_part, part_err = _check_reachable_part(graph, part, "table_cell_properties")
    if not valid_part:
        return {
            "schema_version": 1,
            "operation": "table-cell-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": part_err or "table_cell_properties_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
            "unsupported": [],
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "table-cell-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_cell_properties_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
            "unsupported": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-cell-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_cell_properties_namespace_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
            "unsupported": [],
        }

    def word(local):
        return Q(namespace, local)
    unsupported_context = _unsupported_context_by_element(root, namespace)
    parent_map = {c: p for p in root.iter() for c in p}
    tables = list(root.iter(word("tbl")))
    targets: list[dict[str, Any]] = []
    unsupported: list[dict[str, Any]] = []

    for table_idx, table in enumerate(tables, start=1):
        rows = table.findall(word("tr"))
        for row_idx, row in enumerate(rows, start=1):
            cells = row.findall(word("tc"))
            for cell_idx, cell in enumerate(cells, start=1):
                reason = _inspect_cell_context(
                    cell,
                    row,
                    table,
                    root,
                    namespace,
                    unsupported_context,
                    parent_map,
                )
                if reason is not None:
                    unsupported.append(
                        {
                            "table": table_idx,
                            "row": row_idx,
                            "cell": cell_idx,
                            "part": part,
                            "code": reason,
                        }
                    )
                    continue

                fp = _cell_structure_fingerprint(cell)
                tc_pr = cell.find(word("tcPr"))
                current_props = _inspect_table_cell_properties(tc_pr, namespace)

                target = TableCellPropertiesTarget(
                    table=table_idx,
                    row=row_idx,
                    cell=cell_idx,
                    part=part,
                    structure_fingerprint_sha256=fp,
                    current_properties=current_props,
                )
                targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-cell-properties-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
        "unsupported": unsupported,
    }


def author_table_cell_properties(
    request: TableCellPropertiesRequest,
) -> TableCellPropertiesReport:
    """Author direct formatting table cell properties on an exact body cell atomically."""
    if not isinstance(request, TableCellPropertiesRequest):
        raise TypeError("table_cell_properties_request_invalid")

    source = request.source.resolve()
    destination = request.destination.resolve()

    if source == destination:
        return _table_cell_properties_refusal(
            request, "table_cell_properties_source_destination_same"
        )
    if not source.is_file():
        return _table_cell_properties_refusal(request, "table_cell_properties_source_missing")
    if destination.exists():
        return _table_cell_properties_refusal(request, "table_cell_properties_destination_occupied")
    if not destination.parent.is_dir():
        return _table_cell_properties_refusal(
            request, "table_cell_properties_destination_parent_missing"
        )

    try:
        actual_source_sha = _sha256_file(source)
    except OSError:
        return _table_cell_properties_refusal(request, "table_cell_properties_source_missing")

    if actual_source_sha != request.source_sha256:
        return _table_cell_properties_refusal(
            request, "table_cell_properties_source_sha256_mismatch"
        )

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_cell_properties_refusal(request, "table_cell_properties_package_invalid")

    valid_part, part_err = _check_reachable_part(graph, request.part, "table_cell_properties")
    if not valid_part:
        return _table_cell_properties_refusal(
            request, part_err or "table_cell_properties_part_unsupported"
        )

    try:
        doc_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (ET.ParseError, KeyError):
        return _table_cell_properties_refusal(request, "table_cell_properties_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_cell_properties_refusal(
            request, "table_cell_properties_namespace_unsupported"
        )

    def word(local):
        return Q(namespace, local)
    tables = list(root.iter(word("tbl")))

    if request.table < 1 or request.table > len(tables):
        return _table_cell_properties_refusal(request, "table_cell_properties_table_missing")

    target_table = tables[request.table - 1]
    rows = target_table.findall(word("tr"))

    if request.row < 1 or request.row > len(rows):
        return _table_cell_properties_refusal(request, "table_cell_properties_row_missing")

    target_row = rows[request.row - 1]
    cells = target_row.findall(word("tc"))

    if request.cell < 1 or request.cell > len(cells):
        return _table_cell_properties_refusal(request, "table_cell_properties_cell_missing")

    target_cell = cells[request.cell - 1]

    # Context and validity checks
    parent_map = {c: p for p in root.iter() for c in p}
    unsupported_context = _unsupported_context_by_element(root, namespace)
    context_reason = _inspect_cell_context(
        target_cell,
        target_row,
        target_table,
        root,
        namespace,
        unsupported_context,
        parent_map,
    )
    if context_reason is not None:
        return _table_cell_properties_refusal(request, context_reason)

    # Verify structural fingerprint
    actual_fp = _cell_structure_fingerprint(target_cell)
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_cell_properties_refusal(
            request, "table_cell_properties_structure_fingerprint_mismatch"
        )

    tc_pr = target_cell.find(word("tcPr"))
    props = request.properties_dict()
    merge_origin = _vertical_merge_origin(tc_pr, namespace)
    if merge_origin and not set(props).issubset(_VERTICAL_MERGE_ORIGIN_PROPERTIES):
        return _table_cell_properties_refusal(
            request, "table_cell_properties_merged_width_unsupported"
        )
    properties_before = _inspect_table_cell_properties(tc_pr, namespace)

    authored_keys = sorted(props)
    # Remove only direct overrides. False booleans remain explicit overrides.
    for key in tuple(props):
        if props[key] is None:
            if tc_pr is not None:
                for child in tc_pr.findall(word(_DIRECT_CELL_PROPERTY_TAGS[key])):
                    tc_pr.remove(child)
            del props[key]

    # Ensure tcPr exists at position 0 of target_cell
    if tc_pr is None:
        tc_pr = ET.Element(word("tcPr"))
        target_cell.insert(0, tc_pr)

    # 1. Width (tcW)
    if "width" in props:
        w_spec = props["width"]
        tc_w = ET.Element(word("tcW"))
        tc_w.set(word("w"), str(w_spec["value"]))
        tc_w.set(word("type"), w_spec["type"])
        _insert_or_replace_tcpr_child(tc_pr, tc_w, namespace)

    # 2. Borders (tcBorders)
    if "borders" in props:
        borders_spec = props["borders"]
        existing_borders = tc_pr.find(word("tcBorders"))
        existing_sides: dict[str, ET.Element] = {}
        if existing_borders is not None:
            for c in existing_borders:
                local_side = c.tag.rsplit("}", 1)[-1]
                existing_sides[local_side] = c

        new_borders = ET.Element(
            word("tcBorders"), dict(existing_borders.attrib) if existing_borders is not None else {}
        )
        for side in _TC_BORDERS_ORDER:
            if side in borders_spec:
                spec = borders_spec[side]
                if spec is None:
                    continue
                b_elem = ET.Element(word(side))
                b_elem.set(word("val"), spec["val"])
                if spec.get("sz") is not None:
                    b_elem.set(word("sz"), str(spec["sz"]))
                if spec.get("space") is not None:
                    b_elem.set(word("space"), str(spec["space"]))
                if spec.get("color") is not None:
                    b_elem.set(word("color"), spec["color"])
                new_borders.append(b_elem)
            elif side in existing_sides:
                new_borders.append(copy.deepcopy(existing_sides[side]))
        for side, elem in existing_sides.items():
            if side not in _TC_BORDERS_ORDER:
                new_borders.append(copy.deepcopy(elem))
        _insert_or_replace_tcpr_child(tc_pr, new_borders, namespace)

    # 3. Shading (shd)
    if "shading" in props:
        shd_spec = props["shading"]
        shd = ET.Element(word("shd"))
        shd.set(word("val"), shd_spec.get("val", "clear"))
        shd.set(word("color"), shd_spec.get("color", "auto"))
        shd.set(word("fill"), shd_spec["fill"])
        _insert_or_replace_tcpr_child(tc_pr, shd, namespace)

    # 4. Cell Margins (tcMar)
    if "cell_margins" in props:
        mar_spec = props["cell_margins"]
        existing_mar = tc_pr.find(word("tcMar"))
        existing_mar_sides: dict[str, ET.Element] = {}
        if existing_mar is not None:
            for c in existing_mar:
                local_side = c.tag.rsplit("}", 1)[-1]
                existing_mar_sides[local_side] = c

        new_mar = ET.Element(
            word("tcMar"), dict(existing_mar.attrib) if existing_mar is not None else {}
        )
        for side in _TC_CELL_MAR_ORDER:
            if side in mar_spec:
                if mar_spec[side] is None:
                    continue
                m_elem = ET.Element(word(side))
                m_elem.set(word("w"), str(mar_spec[side]))
                m_elem.set(word("type"), "dxa")
                new_mar.append(m_elem)
            elif side in existing_mar_sides:
                new_mar.append(copy.deepcopy(existing_mar_sides[side]))
        for side, elem in existing_mar_sides.items():
            if side not in _TC_CELL_MAR_ORDER:
                new_mar.append(copy.deepcopy(elem))
        _insert_or_replace_tcpr_child(tc_pr, new_mar, namespace)

    # 5. Text Direction (textDirection)
    if "text_direction" in props:
        text_dir = ET.Element(word("textDirection"))
        text_dir.set(word("val"), props["text_direction"])
        _insert_or_replace_tcpr_child(tc_pr, text_dir, namespace)

    # 6. Vertical Alignment (vAlign)
    if "vertical_alignment" in props:
        align_val = props["vertical_alignment"]
        v_align = ET.Element(word("vAlign"))
        v_align.set(word("val"), align_val)
        _insert_or_replace_tcpr_child(tc_pr, v_align, namespace)

    # 7. No Wrap (noWrap)
    if "no_wrap" in props:
        val = props["no_wrap"]
        if val is True:
            elem = ET.Element(word("noWrap"))
            _insert_or_replace_tcpr_child(tc_pr, elem, namespace)
        elif val is False:
            elem = ET.Element(word("noWrap"))
            elem.set(word("val"), "off")
            _insert_or_replace_tcpr_child(tc_pr, elem, namespace)

    for key, tag in (("fit_text", "tcFitText"), ("hide_mark", "hideMark")):
        if key in props:
            element = ET.Element(word(tag))
            if not props[key]:
                element.set(word("val"), "off")
            _insert_or_replace_tcpr_child(tc_pr, element, namespace)
    if "cnf_style" in props:
        existing_cnf = tc_pr.find(word("cnfStyle"))
        cnf = ET.Element(
            word("cnfStyle"), dict(existing_cnf.attrib) if existing_cnf is not None else {}
        )
        for attr, value in props["cnf_style"].items():
            cnf.set(word(attr), value if attr == "val" else ("1" if value else "0"))
        _insert_or_replace_tcpr_child(tc_pr, cnf, namespace)

    properties_after = _inspect_table_cell_properties(tc_pr, namespace)

    target_change_report: dict[str, Any] = {
        "table": request.table,
        "row": request.row,
        "cell": request.cell,
        "part": request.part,
        "properties_authored": authored_keys,
        "properties_before": properties_before,
        "properties_after": properties_after,
        "boundaries": {
            "table": request.table,
            "row": request.row,
            "cell": request.cell,
        },
    }

    # Extract text of the cell before publication to verify it is untouched
    original_cell_texts = [t.text or "" for t in target_cell.iter(word("t"))]
    original_cell_p_count = len(target_cell.findall(word("p")))

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({request.part})):
            return "table_cell_properties_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_cell_properties_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(request.part))
            s_tables = list(s_root.iter(word("tbl")))
            if len(s_tables) != len(tables):
                return "table_cell_properties_staged_verification_failed"
            s_table = s_tables[request.table - 1]
            s_rows = s_table.findall(word("tr"))
            if len(s_rows) != len(rows):
                return "table_cell_properties_staged_verification_failed"
            s_row = s_rows[request.row - 1]
            s_cells = s_row.findall(word("tc"))
            if len(s_cells) != len(cells):
                return "table_cell_properties_staged_verification_failed"
            s_cell = s_cells[request.cell - 1]
            s_tc_pr = s_cell.find(word("tcPr"))
            if s_tc_pr is None:
                return "table_cell_properties_staged_verification_failed"

            # Check schema sequence order of tcPr children
            seen_ranks: list[int] = []
            for child in s_tc_pr:
                local = child.tag.rsplit("}", 1)[-1]
                if local in _TCPR_CHILD_ORDER:
                    seen_ranks.append(_TCPR_CHILD_ORDER.index(local))
            if seen_ranks != sorted(seen_ranks):
                return "table_cell_properties_staged_verification_failed"

            # A merge origin must still own its region after authoring
            if merge_origin and not _vertical_merge_origin(s_tc_pr, namespace):
                return "table_cell_properties_staged_verification_failed"

            # Verify cell text and paragraphs were preserved verbatim
            staged_texts = [t.text or "" for t in s_cell.iter(word("t"))]
            if staged_texts != original_cell_texts:
                return "table_cell_properties_staged_verification_failed"
            if len(s_cell.findall(word("p"))) != original_cell_p_count:
                return "table_cell_properties_staged_verification_failed"

        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return "table_cell_properties_staged_verification_failed"
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
        failure_code="table_cell_properties_staged_verification_failed",
        exists_code="table_cell_properties_destination_occupied",
        unavailable_code="table_cell_properties_publication_unavailable",
    )

    if not published.passed:
        return _table_cell_properties_refusal(
            request,
            published.code or "table_cell_properties_staged_verification_failed",
        )

    try:
        destination_sha256: str | None = _sha256_file(destination)
    except OSError:
        destination_sha256 = None

    return TableCellPropertiesReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=destination_sha256,
        target=request.target_dict(),
        properties=request.properties_dict(),
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
