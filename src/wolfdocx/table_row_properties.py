"""Source-bound table row properties authoring and discovery for WordprocessingML documents."""

from __future__ import annotations

import json
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
from .table_properties import _VALID_ALIGNMENTS, _canonical_table_alignment

__all__ = [
    "RowHeightSpec",
    "TableRowPropertiesReport",
    "TableRowPropertiesRequest",
    "TableRowPropertiesSpec",
    "TableRowPropertiesTarget",
    "author_table_row_properties",
    "discover_table_row_properties_targets",
    "load_table_row_properties_request",
    "parse_table_row_properties_request",
    "table_row_properties_structure_fingerprint",
]

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

_REQUIRED_TARGET_FIELDS = frozenset({"table", "row"})
_ALLOWED_TARGET_FIELDS = frozenset({"table", "row", "part"})

# Request key -> trPr child, grouped by value shape.
_FLAG_PROPERTY_ELEMENTS = {
    "repeat_header": "tblHeader",
    "cant_split": "cantSplit",
    "hidden": "hidden",
}
_WIDTH_PROPERTY_ELEMENTS = {
    "width_before": "wBefore",
    "width_after": "wAfter",
    "cell_spacing": "tblCellSpacing",
}
_GRID_PROPERTY_ELEMENTS = {"grid_before": "gridBefore", "grid_after": "gridAfter"}
_PROPERTY_ELEMENTS = {
    **_FLAG_PROPERTY_ELEMENTS,
    **_WIDTH_PROPERTY_ELEMENTS,
    **_GRID_PROPERTY_ELEMENTS,
    "height": "trHeight",
    "cnf_style": "cnfStyle",
    "alignment": "jc",
}
_ALLOWED_PROPERTY_KEYS = frozenset(_PROPERTY_ELEMENTS)
_WIDTH_TYPES = frozenset({"dxa", "pct", "auto", "nil"})
_MAX_WIDTH_DXA = 316800
_MAX_WIDTH_PCT = 50000
_MAX_GRID_SHIFT = 63
_CNF_STYLE = re.compile(r"[01]{12}")
# Case-insensitive spellings, including python-docx WD_ROW_HEIGHT_RULE member names.
_HEIGHT_RULE_ALIASES = {
    "auto": "auto",
    "atleast": "atLeast",
    "at_least": "atLeast",
    "exact": "exact",
    "exactly": "exact",
}
_MAX_ROW_HEIGHT_TWIPS = 31680

# ECMA-376 / ISO-IEC 29500-1 §17.4.80 CT_TrPr child element sequence order
_TRPR_CHILD_ORDER: tuple[str, ...] = (
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
    "ins",
    "del",
    "trPrChange",
)


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is the source part byte for byte."""
    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        before_names = set(before.namelist())
        after_names = set(after.namelist())
        if before_names != after_names:
            return False
        for name in before_names:
            if name in changed:
                continue
            if before.read(name) != after.read(name):
                return False
    return True


def _is_supported_table(table: ET.Element, namespace: str) -> bool:
    """True if table has supported body table structure and at least one row."""
    def word(local):
        return Q(namespace, local)
    if any(child.tag not in {word("tblPr"), word("tblGrid"), word("tr")} for child in table):
        return False
    rows = table.findall(word("tr"))
    return len(rows) > 0


def _insert_or_replace_trpr_child(tr_pr: ET.Element, new_child: ET.Element, namespace: str) -> None:
    """Insert or replace a child element in trPr while strictly maintaining schema order."""
    local_name = new_child.tag.rsplit("}", 1)[-1]
    def word(local):
        return Q(namespace, local)
    for existing in list(tr_pr.findall(word(local_name))):
        tr_pr.remove(existing)
    try:
        new_order = _TRPR_CHILD_ORDER.index(local_name)
    except ValueError:
        new_order = 999
    insert_idx = len(tr_pr)
    for idx, child in enumerate(tr_pr):
        child_local = child.tag.rsplit("}", 1)[-1]
        try:
            child_order = _TRPR_CHILD_ORDER.index(child_local)
        except ValueError:
            child_order = 999
        if child_order > new_order:
            insert_idx = idx
            break
    tr_pr.insert(insert_idx, new_child)


def _normalize_height_rule(rule: Any) -> str:
    """Map a row height rule spelling or enum member to its ST_HeightRule value."""
    if not isinstance(rule, str):
        # Enum members such as python-docx WD_ROW_HEIGHT_RULE.EXACTLY expose ``.name``.
        rule = getattr(rule, "name", None)
        if not isinstance(rule, str):
            raise ValueError("table_row_properties_fields_invalid")  # noqa: TRY004 - existing callers assert ValueError
    canonical = _HEIGHT_RULE_ALIASES.get(rule.lower())
    if canonical is None:
        raise ValueError("table_row_properties_fields_invalid")
    return canonical


@dataclass(frozen=True)
class RowHeightSpec:
    """Specification for table row height (w:trHeight)."""

    value: int
    rule: str = "atLeast"

    def __post_init__(self) -> None:
        if isinstance(self.value, bool) or not isinstance(self.value, int):
            raise TypeError("table_row_properties_fields_invalid")
        if self.value < 0 or self.value > _MAX_ROW_HEIGHT_TWIPS:
            raise ValueError("table_row_properties_fields_invalid")
        object.__setattr__(self, "rule", _normalize_height_rule(self.rule))

    def to_dict(self) -> dict[str, Any]:
        return {"value": self.value, "rule": self.rule}


@dataclass(frozen=True)
class TableRowPropertiesSpec:
    """Specification of row-level properties to author on an existing table row.

    ``None`` leaves a property untouched. To remove a property, pass a request
    mapping whose value for that key is ``None``.
    """

    repeat_header: bool | None = None
    height: RowHeightSpec | dict[str, Any] | int | None = None
    cant_split: bool | None = None
    hidden: bool | None = None
    cnf_style: str | None = None
    width_before: dict[str, Any] | int | None = None
    width_after: dict[str, Any] | int | None = None
    cell_spacing: dict[str, Any] | int | None = None
    alignment: str | None = None
    grid_before: int | None = None
    grid_after: int | None = None

    def __post_init__(self) -> None:
        if self.repeat_header is not None and not isinstance(self.repeat_header, bool):
            raise TypeError("table_row_properties_fields_invalid")
        if self.cant_split is not None and not isinstance(self.cant_split, bool):
            raise TypeError("table_row_properties_fields_invalid")
        if self.hidden is not None and not isinstance(self.hidden, bool):
            raise TypeError("table_row_properties_fields_invalid")
        if self.height is not None:
            if isinstance(self.height, bool):
                raise TypeError("table_row_properties_fields_invalid")
            if isinstance(self.height, int):
                if self.height < 0 or self.height > _MAX_ROW_HEIGHT_TWIPS:
                    raise ValueError("table_row_properties_fields_invalid")
            elif isinstance(self.height, (RowHeightSpec, Mapping)):
                pass
            else:
                raise TypeError("table_row_properties_fields_invalid")

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        if self.repeat_header is not None:
            data["repeat_header"] = self.repeat_header
        if self.cant_split is not None:
            data["cant_split"] = self.cant_split
        if self.height is not None:
            if isinstance(self.height, RowHeightSpec):
                data["height"] = self.height.to_dict()
            elif isinstance(self.height, int):
                data["height"] = {"value": self.height, "rule": "atLeast"}
            elif isinstance(self.height, Mapping):
                rule = self.height.get("rule", self.height.get("hRule", "atLeast"))
                val = self.height.get("value", self.height.get("val", 0))
                data["height"] = {"value": val, "rule": rule}
        for key in (
            "hidden",
            "cnf_style",
            "width_before",
            "width_after",
            "cell_spacing",
            "alignment",
            "grid_before",
            "grid_after",
        ):
            if getattr(self, key) is not None:
                data[key] = getattr(self, key)
        return data


def _validate_table_row_properties_target_dict(target: Any) -> dict[str, Any]:
    if not isinstance(target, Mapping):
        raise TypeError("table_row_properties_target_invalid")
    keys = set(target)
    if not keys.issubset(_ALLOWED_TARGET_FIELDS) or not _REQUIRED_TARGET_FIELDS.issubset(keys):
        raise ValueError("table_row_properties_target_fields_invalid")
    table = target["table"]
    if isinstance(table, bool) or not isinstance(table, int) or table < 1:
        raise ValueError("table_row_properties_table_invalid")
    row = target["row"]
    if isinstance(row, bool) or not isinstance(row, int) or row < 1:
        raise ValueError("table_row_properties_row_invalid")
    part = target.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("table_row_properties_part_invalid")
    return {"table": table, "row": row, "part": part}


def _validate_flag(value: Any) -> bool | None:
    if value is not None and not isinstance(value, bool):
        raise ValueError("table_row_properties_fields_invalid")
    return value


def _validate_height(value: Any) -> dict[str, Any] | None:
    if isinstance(value, RowHeightSpec):
        value = value.to_dict()
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError("table_row_properties_fields_invalid")  # noqa: TRY004 - one error family for every bad value
    if isinstance(value, int):
        if value < 0 or value > _MAX_ROW_HEIGHT_TWIPS:
            raise ValueError("table_row_properties_fields_invalid")
        return {"value": value, "rule": "atLeast"}
    if not isinstance(value, Mapping):
        raise TypeError("table_row_properties_fields_invalid")
    keys = set(value)
    if not keys.issubset({"value", "val", "rule", "hRule"}) or not ({"value", "val"} & keys):
        raise ValueError("table_row_properties_fields_invalid")
    raw = value.get("value", value.get("val"))
    if isinstance(raw, bool) or not isinstance(raw, int) or raw < 0 or raw > _MAX_ROW_HEIGHT_TWIPS:
        raise ValueError("table_row_properties_fields_invalid")
    rule = _normalize_height_rule(value.get("rule", value.get("hRule", "atLeast")))
    return {"value": raw, "rule": rule}


def _validate_width(value: Any) -> dict[str, Any] | None:
    """A ``w:wBefore``-style measure: an int of twips or ``{"value", "type"}``."""
    if value is None:
        return None
    if isinstance(value, int) and not isinstance(value, bool):
        value = {"value": value}
    if not isinstance(value, Mapping):
        raise ValueError("table_row_properties_fields_invalid")  # noqa: TRY004 - one error family for every bad value
    if not set(value).issubset({"value", "type"}) or "value" not in value:
        raise ValueError("table_row_properties_fields_invalid")
    raw, kind = value["value"], value.get("type", "dxa")
    if isinstance(raw, bool) or not isinstance(raw, int) or raw < 0:
        raise ValueError("table_row_properties_fields_invalid")
    if not isinstance(kind, str) or kind not in _WIDTH_TYPES:
        raise ValueError("table_row_properties_fields_invalid")
    if (kind == "dxa" and raw > _MAX_WIDTH_DXA) or (kind == "pct" and raw > _MAX_WIDTH_PCT):
        raise ValueError("table_row_properties_fields_invalid")
    return {"value": raw, "type": kind}


def _validate_alignment(value: Any) -> str | None:
    if value is None:
        return None
    alignment = _canonical_table_alignment(value, _VALID_ALIGNMENTS)
    if alignment is None:
        raise ValueError("table_row_properties_fields_invalid")
    return alignment


def _validate_cnf_style(value: Any) -> str | None:
    if value is not None and (not isinstance(value, str) or not _CNF_STYLE.fullmatch(value)):
        raise ValueError("table_row_properties_fields_invalid")
    return value


def _validate_grid_shift(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError("table_row_properties_fields_invalid")  # noqa: TRY004 - one error family for every bad value
    if value < 0 or value > _MAX_GRID_SHIFT:
        raise ValueError("table_row_properties_fields_invalid")
    return value


def _validate_table_row_properties_dict(properties: Any) -> dict[str, Any]:
    """Validate a properties mapping; a ``None`` value requests removal of that property."""
    if isinstance(properties, TableRowPropertiesSpec):
        properties = properties.to_dict()
    if not isinstance(properties, Mapping):
        raise TypeError("table_row_properties_fields_invalid")
    if not set(properties).issubset(_ALLOWED_PROPERTY_KEYS):
        raise ValueError("table_row_properties_fields_invalid")

    result: dict[str, Any] = {}
    for key, value in properties.items():
        if key in _FLAG_PROPERTY_ELEMENTS:
            result[key] = _validate_flag(value)
        elif key == "height":
            result[key] = _validate_height(value)
        elif key in _WIDTH_PROPERTY_ELEMENTS:
            result[key] = _validate_width(value)
        elif key == "alignment":
            result[key] = _validate_alignment(value)
        elif key == "cnf_style":
            result[key] = _validate_cnf_style(value)
        else:
            result[key] = _validate_grid_shift(value)
    return result


@dataclass(frozen=True)
class TableRowPropertiesTarget:
    """One discovered candidate table row target for row properties authoring."""

    table: int
    row: int
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    cells_count: int = 0
    current_properties: dict[str, Any] | None = None

    @property
    def target(self) -> dict[str, Any]:
        return {"table": self.table, "row": self.row, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "row": self.row,
            "part": self.part,
            "cells_count": self.cells_count,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "current_properties": self.current_properties,
        }


@dataclass(frozen=True)
class TableRowPropertiesRequest:
    """One source-bound request to author direct formatting table row properties on an exact row."""

    source: Path
    destination: Path
    source_sha256: str
    target: Mapping[str, Any]
    structure_fingerprint_sha256: str
    properties: Mapping[str, Any] | TableRowPropertiesSpec

    def __post_init__(self) -> None:
        if isinstance(self.source, str):
            object.__setattr__(self, "source", Path(self.source))
        elif not isinstance(self.source, Path):
            raise TypeError("table_row_properties_source_invalid")
        if not str(self.source):
            raise ValueError("table_row_properties_source_invalid")

        if isinstance(self.destination, str):
            object.__setattr__(self, "destination", Path(self.destination))
        elif not isinstance(self.destination, Path):
            raise TypeError("table_row_properties_destination_invalid")
        if not str(self.destination):
            raise ValueError("table_row_properties_destination_invalid")

        _require_sha256(self.source_sha256, "table_row_properties_source_sha256")
        _require_sha256(
            self.structure_fingerprint_sha256,
            "table_row_properties_structure_fingerprint_sha256",
        )

        validated_target = _validate_table_row_properties_target_dict(self.target)
        object.__setattr__(self, "target", validated_target)

        validated_properties = _validate_table_row_properties_dict(self.properties)
        object.__setattr__(self, "properties", validated_properties)

    @property
    def table(self) -> int:
        return self.target["table"]

    @property
    def row(self) -> int:
        return self.target["row"]

    @property
    def part(self) -> str:
        return self.target.get("part", _DOCUMENT)

    def target_dict(self) -> dict[str, Any]:
        return {"table": self.table, "row": self.row, "part": self.part}

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
class TableRowPropertiesReport:
    """Content-safe public result for a table row properties authoring request."""

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
            "operation": "table-row-properties",
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


def _table_row_properties_refusal(
    request: TableRowPropertiesRequest,
    code: str,
    target: dict[str, Any] | None = None,
) -> TableRowPropertiesReport:
    return TableRowPropertiesReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=target if target is not None else request.target_dict(),
        properties=request.properties_dict(),
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_table_row_properties_request(value: Any) -> TableRowPropertiesRequest:
    """Parse and strictly validate a table row properties authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("table_row_properties_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("table_row_properties_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("table_row_properties_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("table_row_properties_destination_invalid")

    source_sha256 = _require_sha256(value["source_sha256"], "table_row_properties_source_sha256")
    structure_fingerprint = _require_sha256(
        value["structure_fingerprint_sha256"],
        "table_row_properties_structure_fingerprint_sha256",
    )

    target = _validate_table_row_properties_target_dict(value["target"])
    properties = _validate_table_row_properties_dict(value["properties"])

    return TableRowPropertiesRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        target=target,
        structure_fingerprint_sha256=structure_fingerprint,
        properties=properties,
    )


def load_table_row_properties_request(path: str | Path) -> TableRowPropertiesRequest:
    """Load and strictly validate one JSON table-row-properties request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_table_row_properties_request(json.load(stream))


def _row_structure_fingerprint(row: ET.Element) -> str:
    """Hash the row's XML structure while deliberately masking cell text."""
    return _cell_structure_fingerprint(row)


def table_row_properties_structure_fingerprint(
    source: str | Path,
    *,
    table: int,
    row: int,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one target table row."""
    if part != _DOCUMENT:
        return None
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
        if part not in graph.parts:
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
        if not _is_supported_table(target_table, ns):
            return None
        rows = target_table.findall(word("tr"))
        if row < 1 or row > len(rows):
            return None
        target_row = rows[row - 1]
        return _row_structure_fingerprint(target_row)
    except (PackageError, OSError, ET.ParseError, KeyError, IndexError):
        return None


def _int_or_zero(raw: str | None) -> int:
    if raw is not None and (raw.isdigit() or (raw.startswith("-") and raw[1:].isdigit())):
        return int(raw)
    return 0


def _inspect_row_properties(tr_pr: ET.Element | None, namespace: str) -> dict[str, Any]:
    """Inspect and extract current row properties from w:trPr."""
    if tr_pr is None:
        return {}
    def word(local):
        return Q(namespace, local)
    result: dict[str, Any] = {}

    def on_off(element: ET.Element) -> bool:
        val = element.get(word("val"))
        return val is None or val.lower() not in {"0", "false", "off"}

    cnf_style = tr_pr.find(word("cnfStyle"))
    if cnf_style is not None and cnf_style.get(word("val")) is not None:
        result["cnf_style"] = cnf_style.get(word("val"))

    for key, local in _GRID_PROPERTY_ELEMENTS.items():
        grid = tr_pr.find(word(local))
        if grid is not None:
            result[key] = _int_or_zero(grid.get(word("val")))

    for key in ("width_before", "width_after"):
        width = tr_pr.find(word(_WIDTH_PROPERTY_ELEMENTS[key]))
        if width is not None:
            result[key] = {
                "value": _int_or_zero(width.get(word("w"))),
                "type": width.get(word("type")) or "dxa",
            }

    cant_split = tr_pr.find(word("cantSplit"))
    if cant_split is not None:
        result["cant_split"] = on_off(cant_split)

    tr_height = tr_pr.find(word("trHeight"))
    if tr_height is not None:
        result["height"] = {
            "value": _int_or_zero(tr_height.get(word("val"))),
            "rule": tr_height.get(word("hRule")) or "atLeast",
        }

    tbl_header = tr_pr.find(word("tblHeader"))
    if tbl_header is not None:
        result["repeat_header"] = on_off(tbl_header)

    spacing = tr_pr.find(word("tblCellSpacing"))
    if spacing is not None:
        result["cell_spacing"] = {
            "value": _int_or_zero(spacing.get(word("w"))),
            "type": spacing.get(word("type")) or "dxa",
        }

    jc = tr_pr.find(word("jc"))
    if jc is not None:
        result["alignment"] = jc.get(word("val")) or "left"

    hidden = tr_pr.find(word("hidden"))
    if hidden is not None:
        result["hidden"] = on_off(hidden)

    return result


def _row_covers_grid(
    table: ET.Element, row: ET.Element, tr_pr: ET.Element | None, namespace: str
) -> bool:
    """True when gridBefore, the cell spans and gridAfter exactly fill the table grid."""
    def word(local):
        return Q(namespace, local)
    grid = table.find(word("tblGrid"))
    columns = 0 if grid is None else len(grid.findall(word("gridCol")))

    def shift(local: str) -> int:
        element = None if tr_pr is None else tr_pr.find(word(local))
        return 0 if element is None else _int_or_zero(element.get(word("val")))

    spans = 0
    for cell in row.findall(word("tc")):
        grid_span = cell.find(word("tcPr") + "/" + word("gridSpan"))
        raw = None if grid_span is None else grid_span.get(word("val"))
        spans += 1 if raw is None else _int_or_zero(raw)
    return columns > 0 and shift("gridBefore") + spans + shift("gridAfter") == columns


def _row_property_element(key: str, value: Any, namespace: str) -> ET.Element:
    """The trPr child expressing one validated, non-``None`` request value."""
    def word(local):
        return Q(namespace, local)
    element = ET.Element(word(_PROPERTY_ELEMENTS[key]))
    if key in _FLAG_PROPERTY_ELEMENTS:
        if value is False:
            element.set(word("val"), "off")
    elif key == "height":
        element.set(word("val"), str(value["value"]))
        element.set(word("hRule"), value["rule"])
    elif key in _WIDTH_PROPERTY_ELEMENTS:
        element.set(word("w"), str(value["value"]))
        element.set(word("type"), value["type"])
    else:
        element.set(word("val"), str(value))
    return element


def _apply_row_properties(tr_pr: ET.Element, properties: Mapping[str, Any], namespace: str) -> None:
    """Set each named property in schema order; a ``None`` value removes it."""
    for key, value in properties.items():
        if value is None:
            for existing in tr_pr.findall(Q(namespace, _PROPERTY_ELEMENTS[key])):
                tr_pr.remove(existing)
        else:
            _insert_or_replace_trpr_child(
                tr_pr, _row_property_element(key, value, namespace), namespace
            )


def discover_table_row_properties_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate table row properties targets for authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "table-row-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_properties_source_missing",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    try:
        source_sha256 = _sha256_file(path)
    except OSError:
        return {
            "schema_version": 1,
            "operation": "table-row-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_properties_source_missing",
            "source_sha256": None,
            "part": part,
            "targets": [],
        }

    try:
        graph = PackageGraph.from_file(path)
    except (PackageError, OSError, zipfile.BadZipFile):
        return {
            "schema_version": 1,
            "operation": "table-row-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_properties_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    if part != _DOCUMENT or part not in graph.parts:
        return {
            "schema_version": 1,
            "operation": "table-row-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_properties_part_unsupported",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    try:
        doc_bytes = graph.read_part_bytes(part)
        root = ET.fromstring(doc_bytes)
    except (ET.ParseError, KeyError, OSError):
        return {
            "schema_version": 1,
            "operation": "table-row-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_properties_package_invalid",
            "source_sha256": source_sha256,
            "part": part,
            "targets": [],
        }

    namespace = _word_namespace(root)
    if namespace is None:
        return {
            "schema_version": 1,
            "operation": "table-row-properties-targets",
            "passed": False,
            "status": _REFUSED,
            "code": "table_row_properties_namespace_unsupported",
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

        rows = table.findall(word("tr"))
        for row_idx, row in enumerate(rows, start=1):
            if unsupported_context.get(id(row), False):
                continue
            cells = row.findall(word("tc"))
            if not cells:
                continue

            tr_prs = row.findall(word("trPr"))
            if len(tr_prs) > 1:
                continue
            tr_pr = tr_prs[0] if tr_prs else None
            if tr_pr is not None:
                child_local_tags = [c.tag.rsplit("}", 1)[-1] for c in tr_pr]
                if len(child_local_tags) != len(set(child_local_tags)):
                    continue
                shifted = (
                    tr_pr.find(word("gridBefore")) is not None
                    or tr_pr.find(word("gridAfter")) is not None
                )
                if shifted and not _row_covers_grid(table, row, tr_pr, namespace):
                    continue

            fp = _row_structure_fingerprint(row)
            current_props = _inspect_row_properties(tr_pr, namespace)

            target = TableRowPropertiesTarget(
                table=table_idx,
                row=row_idx,
                structure_fingerprint_sha256=fp,
                part=part,
                cells_count=len(cells),
                current_properties=current_props,
            )
            targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "table-row-properties-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_table_row_properties(
    request: TableRowPropertiesRequest,
) -> TableRowPropertiesReport:
    """Author direct formatting table row properties on a source-bound document row atomically."""
    if not isinstance(request, TableRowPropertiesRequest):
        raise TypeError("table_row_properties_request_invalid")

    source = request.source
    destination = request.destination

    if not source.is_file():
        return _table_row_properties_refusal(request, "table_row_properties_source_missing")

    if source.resolve() == destination.resolve() or source == destination:
        return _table_row_properties_refusal(request, "table_row_properties_destination_invalid")

    if not destination.parent.is_dir():
        return _table_row_properties_refusal(request, "table_row_properties_destination_invalid")

    if destination.exists() or destination.is_symlink():
        return _table_row_properties_refusal(request, "table_row_properties_destination_occupied")

    try:
        source_sha256 = _sha256_file(source)
    except OSError:
        return _table_row_properties_refusal(request, "table_row_properties_source_missing")

    if source_sha256 != request.source_sha256:
        return _table_row_properties_refusal(request, "table_row_properties_source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _table_row_properties_refusal(request, "table_row_properties_package_invalid")

    if request.part != _DOCUMENT or request.part not in graph.parts:
        return _table_row_properties_refusal(request, "table_row_properties_part_unsupported")

    try:
        doc_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (ET.ParseError, KeyError, OSError):
        return _table_row_properties_refusal(request, "table_row_properties_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _table_row_properties_refusal(request, "table_row_properties_namespace_unsupported")

    def word(local):
        return Q(namespace, local)
    tables = list(root.iter(word("tbl")))
    if request.table < 1 or request.table > len(tables):
        return _table_row_properties_refusal(request, "table_row_properties_target_missing")

    target_table = tables[request.table - 1]

    if not _is_supported_table(target_table, namespace):
        return _table_row_properties_refusal(request, "table_row_properties_context_unsupported")

    unsupported_context = _unsupported_context_by_element(root, namespace)
    if unsupported_context.get(id(target_table), False):
        return _table_row_properties_refusal(request, "table_row_properties_context_unsupported")

    rows = target_table.findall(word("tr"))
    if request.row < 1 or request.row > len(rows):
        return _table_row_properties_refusal(request, "table_row_properties_target_missing")

    target_row = rows[request.row - 1]

    if unsupported_context.get(id(target_row), False):
        return _table_row_properties_refusal(request, "table_row_properties_context_unsupported")

    # Conflict check: duplicate trPr on target row
    tr_prs = target_row.findall(word("trPr"))
    if len(tr_prs) > 1:
        return _table_row_properties_refusal(request, "table_row_properties_conflict")

    # Conflict check: duplicate same-name children inside trPr
    tr_pr = tr_prs[0] if tr_prs else None
    if tr_pr is not None:
        child_local_tags = [c.tag.rsplit("}", 1)[-1] for c in tr_pr]
        if len(child_local_tags) != len(set(child_local_tags)):
            return _table_row_properties_refusal(request, "table_row_properties_conflict")
        shifted = (
            tr_pr.find(word("gridBefore")) is not None or tr_pr.find(word("gridAfter")) is not None
        )
        if shifted and not _row_covers_grid(target_table, target_row, tr_pr, namespace):
            return _table_row_properties_refusal(
                request, "table_row_properties_context_unsupported"
            )

        # Check trPr position relative to tc elements
        tr_pr_idx = list(target_row).index(tr_pr)
        tc_indices = [i for i, c in enumerate(target_row) if c.tag == word("tc")]
        if tc_indices and tr_pr_idx > tc_indices[0]:
            return _table_row_properties_refusal(request, "table_row_properties_conflict")

    # Fingerprint check
    actual_fp = _row_structure_fingerprint(target_row)
    if actual_fp != request.structure_fingerprint_sha256:
        return _table_row_properties_refusal(request, "table_row_properties_fingerprint_mismatch")

    properties_before = _inspect_row_properties(tr_pr, namespace)

    props = request.properties_dict()
    working = tr_pr if tr_pr is not None else ET.Element(word("trPr"))
    _apply_row_properties(working, props, namespace)
    if _GRID_PROPERTY_ELEMENTS.keys() & props.keys() and not _row_covers_grid(
        target_table, target_row, working, namespace
    ):
        return _table_row_properties_refusal(request, "table_row_properties_grid_mismatch")

    if tr_pr is None and len(working):
        tr_pr = working
        tbl_pr_ex = target_row.find(word("tblPrEx"))
        target_row.insert(0 if tbl_pr_ex is None else list(target_row).index(tbl_pr_ex) + 1, tr_pr)
    elif tr_pr is not None and not len(tr_pr):
        target_row.remove(tr_pr)
        tr_pr = None

    properties_after = _inspect_row_properties(tr_pr, namespace)

    cells = target_row.findall(word("tc"))
    target_change_report: dict[str, Any] = {
        "table": request.table,
        "row": request.row,
        "part": request.part,
        "properties_authored": sorted(props.keys()),
        "properties_before": properties_before,
        "properties_after": properties_after,
        "cells_count": len(cells),
    }

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset({_DOCUMENT})):
            return "table_row_properties_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "table_row_properties_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
            s_tables = list(s_root.iter(word("tbl")))
            if len(s_tables) != len(tables):
                return "table_row_properties_staged_verification_failed"
            s_table = s_tables[request.table - 1]
            s_rows = s_table.findall(word("tr"))
            if len(s_rows) != len(rows):
                return "table_row_properties_staged_verification_failed"
            s_row = s_rows[request.row - 1]
            s_tr_pr = s_row.find(word("trPr"))
            if _inspect_row_properties(s_tr_pr, namespace) != properties_after:
                return "table_row_properties_staged_verification_failed"
            if s_tr_pr is not None:
                seen_ranks: list[int] = []
                for child in s_tr_pr:
                    local = child.tag.rsplit("}", 1)[-1]
                    if local in _TRPR_CHILD_ORDER:
                        seen_ranks.append(_TRPR_CHILD_ORDER.index(local))
                if seen_ranks != sorted(seen_ranks):
                    return "table_row_properties_staged_verification_failed"
            # Verify cell count and cell text are identical to pre-authoring
            s_cells = s_row.findall(word("tc"))
            orig_cells = target_row.findall(word("tc"))
            if len(s_cells) != len(orig_cells):
                return "table_row_properties_staged_verification_failed"
            for s_c, orig_c in zip(s_cells, orig_cells):
                if [t.text for t in s_c.iter(word("t"))] != [
                    t.text for t in orig_c.iter(word("t"))
                ]:
                    return "table_row_properties_staged_verification_failed"
        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return "table_row_properties_staged_verification_failed"
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
        failure_code="table_row_properties_staged_verification_failed",
        exists_code="table_row_properties_destination_occupied",
        unavailable_code="table_row_properties_publication_unavailable",
    )

    if not published.passed:
        return _table_row_properties_refusal(
            request, published.code or "table_row_properties_staged_verification_failed"
        )

    return TableRowPropertiesReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        properties=props,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
