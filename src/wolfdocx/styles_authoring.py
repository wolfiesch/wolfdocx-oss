"""Source-bound native Word style definition authoring and discovery.

Provides strict, schema-valid creation and updates of paragraph, character, and table
style definitions in ``word/styles.xml``, bound to source package hashes and
structural fingerprints.

Refuses unsupported or invalid operations:
- Malformed requests (unknown keys, missing required fields, invalid types).
- Incompatible style types (e.g. numbering styles).
- Duplicate style IDs on creation; missing target styles on update.
- Missing referenced base or next styles, or incompatible reference types.
- Inheritance and next-style cycles detected via graph semantics.
- Stale source hashes or structure fingerprints.
- Occupied destination paths.
- Unsupported revision or wrapper contexts.

Preserves everything else:
- Unrelated package parts are byte-identical.
- Latent styles and document defaults are preserved intact.
- Unspecified properties and unknown children of updated styles are preserved.
- Schema child order of CT_Style, CT_PPr, CT_RPr, CT_TblPr, CT_TrPr, CT_TcPr, and CT_TblStylePr is strictly enforced.
"""

from __future__ import annotations

import copy
import hashlib
import json
import re
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, fields
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree as ET

from .formatting import (
    _HEX_COLOR_PATTERN,
    _PPR_RANK,
    _RPR_RANK,
    _STYLES,
    _VALID_UNDERLINE_STYLES,
    _VALID_VERTICAL_ALIGNMENTS,
    IndentSpec,
    SpacingSpec,
    _canonical_alignment,
    _canonical_highlight,
    _fold_flag_alias,
    _local_name,
    _validate_indent_spec,
    _validate_spacing_spec,
)
from .inspect import (
    STRICT_WORD_NS,
    SUPPORTED_WORD_NAMESPACES,
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import (
    _DOCUMENT,
    ReviewReport,
    _publish,
    _sha,
    _sha256,
    _word_namespace,
)
from .table_cell_properties import (
    _TC_BORDERS_ORDER,
    _TC_CELL_MAR_ORDER,
    TableCellPropertiesSpec,
    _insert_or_replace_tcpr_child,
    _validate_table_cell_properties_dict,
)
from .table_properties import (
    _TBL_BORDERS_ORDER,
    _TBL_CELL_MAR_ORDER,
    TablePropertiesSpec,
    _insert_or_replace_tblpr_child,
    _validate_table_properties_dict,
)
from .table_row_properties import (
    _TRPR_CHILD_ORDER,
    TableRowPropertiesSpec,
    _insert_or_replace_trpr_child,
    _validate_table_row_properties_dict,
)

_PUBLISHED = "published"
_REFUSED = "refused"

_STYLES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
_STRICT_STYLES_REL_TYPE = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles"
_STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is byte-identical to source."""
    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        after_names = set(after.namelist())
        before_names = {info.filename for info in before.infolist()}
        if not after_names.issubset(before_names | changed):
            return False
        return all(
            info.filename in after_names and before.read(info.filename) == after.read(info.filename)
            for info in before.infolist()
            if info.filename not in changed
        )


_ABSENT_STYLES_FINGERPRINT = hashlib.sha256(b"absent:word/styles.xml").hexdigest()

# ECMA-376 CT_Style child element sequence order
_STYLE_CHILD_ORDER: tuple[str, ...] = (
    "name",
    "aliases",
    "basedOn",
    "next",
    "link",
    "autoRedefine",
    "hidden",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "locked",
    "personal",
    "personalCompose",
    "personalReply",
    "rsid",
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
    "tblStylePr",
)
_STYLE_CHILD_RANK: dict[str, int] = {tag: i for i, tag in enumerate(_STYLE_CHILD_ORDER)}

# w:link pairs a paragraph style with its character companion style.
_LINKED_STYLE_TYPE: dict[str, str] = {"paragraph": "character", "character": "paragraph"}


def _set_style_link(style_elem: ET.Element, value: str, namespace: str) -> None:
    """Set w:link at its CT_Style sequence position, leaving unknown children in place."""
    link_elem = style_elem.find(Q(namespace, "link"))
    if link_elem is None:
        link_rank = _STYLE_CHILD_RANK["link"]
        index = next(
            (
                position
                for position, child in enumerate(style_elem)
                if _STYLE_CHILD_RANK.get(_local_name(child.tag), -1) > link_rank
            ),
            len(style_elem),
        )
        link_elem = ET.Element(Q(namespace, "link"))
        style_elem.insert(index, link_elem)
    link_elem.set(Q(namespace, "val"), value)


_STYLE_TABLE_PROPERTY_KEYS: frozenset[str] = frozenset(
    {"alignment", "borders", "shading", "cell_margins"}
)
_VALID_TBL_STYLE_OVERRIDE_TYPES: frozenset[str] = frozenset(
    {
        "wholeTable",
        "firstRow",
        "lastRow",
        "firstCol",
        "lastCol",
        "band1Horz",
        "band2Horz",
        "band1Vert",
        "band2Vert",
        "neCell",
        "nwCell",
        "seCell",
        "swCell",
    }
)

_TBL_STYLE_OVERRIDE_TYPE_ORDER: tuple[str, ...] = (
    "wholeTable",
    "firstRow",
    "lastRow",
    "firstCol",
    "lastCol",
    "band1Horz",
    "band2Horz",
    "band1Vert",
    "band2Vert",
    "nwCell",
    "neCell",
    "swCell",
    "seCell",
)

_OVERRIDE_TYPE_MAP: dict[str, str] = {
    "wholetable": "wholeTable",
    "whole_table": "wholeTable",
    "firstrow": "firstRow",
    "first_row": "firstRow",
    "lastrow": "lastRow",
    "last_row": "lastRow",
    "firstcol": "firstCol",
    "first_col": "firstCol",
    "firstcolumn": "firstCol",
    "first_column": "firstCol",
    "lastcol": "lastCol",
    "last_col": "lastCol",
    "lastcolumn": "lastCol",
    "last_column": "lastCol",
    "nwcell": "nwCell",
    "nw_cell": "nwCell",
    "northwest_cell": "nwCell",
    "necell": "neCell",
    "ne_cell": "neCell",
    "northeast_cell": "neCell",
    "swcell": "swCell",
    "sw_cell": "swCell",
    "southwest_cell": "swCell",
    "secell": "seCell",
    "se_cell": "seCell",
    "southeast_cell": "seCell",
}

# ECMA-376 CT_TblStylePr child element sequence order: pPr, rPr, tblPr, trPr, tcPr
_TBL_STYLE_PR_CHILD_ORDER: tuple[str, ...] = (
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
)
_TBL_STYLE_PR_CHILD_RANK: dict[str, int] = {
    tag: i for i, tag in enumerate(_TBL_STYLE_PR_CHILD_ORDER)
}


@dataclass(frozen=True)
class TableConditionalOverrideSpec:
    """Specification of one conditional table style override (w:tblStylePr)."""

    type: str
    run_properties: Mapping[str, Any] | None = None
    paragraph_properties: Mapping[str, Any] | None = None
    table_properties: Mapping[str, Any] | TablePropertiesSpec | None = None
    row_properties: Mapping[str, Any] | TableRowPropertiesSpec | None = None
    cell_properties: Mapping[str, Any] | TableCellPropertiesSpec | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.type, str) or not self.type:
            raise ValueError("style_conditional_override_type_invalid")

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"type": self.type}
        if self.run_properties is not None:
            result["run_properties"] = dict(self.run_properties)
        if self.paragraph_properties is not None:
            p_dict: dict[str, Any] = {}
            for k, v in self.paragraph_properties.items():
                if hasattr(v, "to_dict"):
                    p_dict[k] = v.to_dict()
                else:
                    p_dict[k] = v
            result["paragraph_properties"] = p_dict
        if self.table_properties is not None:
            if hasattr(self.table_properties, "to_dict"):
                result["table_properties"] = self.table_properties.to_dict()
            else:
                result["table_properties"] = dict(self.table_properties)
        if self.row_properties is not None:
            if hasattr(self.row_properties, "to_dict"):
                result["row_properties"] = self.row_properties.to_dict()
            else:
                result["row_properties"] = dict(self.row_properties)
        if self.cell_properties is not None:
            if isinstance(self.cell_properties, TableCellPropertiesSpec):
                result["cell_properties"] = _table_cell_properties_spec_mapping(
                    self.cell_properties
                )
            else:
                result["cell_properties"] = dict(self.cell_properties)
        return result


# Public property names follow the formatting census; XML names retain schema spelling.
_SCRIPT_RUN_FLAGS = {
    "rtl": "rtl",
    "complex_script": "cs",
    "bold_complex_script": "bCs",
    "italic_complex_script": "iCs",
}
_SCRIPT_PARAGRAPH_FLAGS = {
    "bidi": "bidi",
    "kinsoku": "kinsoku",
    "word_wrap": "wordWrap",
    "overflow_punct": "overflowPunct",
    "top_line_punct": "topLinePunct",
    "auto_space_de": "autoSpaceDE",
    "auto_space_dn": "autoSpaceDN",
    "snap_to_grid": "snapToGrid",
}
_SCRIPT_FONT_ATTRIBUTES = {
    "font_face_high_ansi": "hAnsi",
    "font_face_east_asia": "eastAsia",
    "font_face_complex_script": "cs",
    "font_theme_ascii": "asciiTheme",
    "font_theme_high_ansi": "hAnsiTheme",
    "font_theme_east_asia": "eastAsiaTheme",
    "font_theme_complex_script": "cstheme",
    "font_hint": "hint",
}
_SCRIPT_LANGUAGE_ATTRIBUTES = {
    "language": "val",
    "language_east_asia": "eastAsia",
    "language_bidi": "bidi",
}
_SCRIPT_RUN_KEYS = {
    *_SCRIPT_RUN_FLAGS,
    *_SCRIPT_FONT_ATTRIBUTES,
    *_SCRIPT_LANGUAGE_ATTRIBUTES,
    "font_size_complex_script",
    "east_asian_layout",
    "fit_text",
    "emphasis_mark",
}
_SCRIPT_PARAGRAPH_KEYS = {*_SCRIPT_PARAGRAPH_FLAGS, "text_alignment", "text_direction"}
_STYLE_RUN_KEYS = {
    "bold",
    "italic",
    "underline",
    "strike",
    "color",
    "font_size",
    "vertical_alignment",
    "font_face",
    "highlight",
    "style",
    *_SCRIPT_RUN_KEYS,
}
_STYLE_PARAGRAPH_KEYS = {
    "alignment",
    "spacing",
    "indent",
    "keep_next",
    "keep_lines",
    "keep_with_next",
    "keep_together",
    "outline_level",
    "style",
    *_SCRIPT_PARAGRAPH_KEYS,
}
_SCRIPT_ENUMS = {
    "font_hint": {"default", "eastAsia", "cs"},
    "emphasis_mark": {"none", "dot", "comma", "circle", "underDot"},
    "text_alignment": {"top", "center", "baseline", "bottom", "auto"},
    "text_direction": {"lrTb", "tbRl", "btLr", "lrTbV", "tbRlV", "tbLrV"},
}
_FONT_THEMES = {
    "majorEastAsia",
    "majorBidi",
    "majorAscii",
    "majorHAnsi",
    "minorEastAsia",
    "minorBidi",
    "minorAscii",
    "minorHAnsi",
}


def _validate_script_properties(value: Mapping[str, Any], *, paragraph: bool) -> dict[str, Any]:
    flags = _SCRIPT_PARAGRAPH_FLAGS if paragraph else _SCRIPT_RUN_FLAGS
    keys = _SCRIPT_PARAGRAPH_KEYS if paragraph else _SCRIPT_RUN_KEYS
    allowed = _STYLE_PARAGRAPH_KEYS if paragraph else _STYLE_RUN_KEYS
    unknown = set(value) - allowed
    if unknown:
        kind = "paragraph" if paragraph else "run"
        raise ValueError(f"style_{kind}_properties_unknown_keys: {sorted(unknown)}")
    validated = {}
    for key in keys:
        item = value.get(key)
        if item is None:
            continue
        valid = True
        if key in flags:
            valid = isinstance(item, bool)
        elif key in _SCRIPT_ENUMS:
            valid = isinstance(item, str) and item in _SCRIPT_ENUMS[key]
        elif key.startswith("font_theme_"):
            valid = isinstance(item, str) and item in _FONT_THEMES
        elif key == "font_size_complex_script":
            valid = isinstance(item, int) and not isinstance(item, bool) and item > 0
        elif key in {"east_asian_layout", "fit_text"}:
            if not isinstance(item, Mapping):
                valid = False
            else:
                item = dict(item)
                attributes = (
                    {"id", "combine", "combineBrackets", "vert", "vertCompress"}
                    if key == "east_asian_layout"
                    else {"id", "val"}
                )
                valid = bool(item) and not (set(item) - attributes)
                if key == "fit_text" and "val" not in item:
                    valid = False
                for attribute, content in item.items():
                    if attribute in {"combine", "vert", "vertCompress"}:
                        valid &= isinstance(content, bool)
                    elif attribute == "combineBrackets":
                        valid &= isinstance(content, str) and content in {
                            "none",
                            "round",
                            "square",
                            "angle",
                            "curly",
                        }
                    else:
                        valid &= (
                            isinstance(content, int)
                            and not isinstance(content, bool)
                            and content >= 0
                            and (attribute != "id" or content <= 2147483647)
                        )
        else:
            valid = isinstance(item, str) and bool(item) and len(item) <= 256
        if not valid:
            raise ValueError(f"style_{key}_invalid")
        validated[key] = item
    return validated


def _write_script_properties(
    parent: ET.Element, properties: Mapping[str, Any], namespace: str, *, paragraph: bool
) -> None:
    def word(local):
        return Q(namespace, local)
    flags = _SCRIPT_PARAGRAPH_FLAGS if paragraph else _SCRIPT_RUN_FLAGS
    for key, tag in flags.items():
        if properties.get(key) is not None:
            for existing in list(parent.findall(word(tag))):
                parent.remove(existing)
            parent.append(ET.Element(word(tag), {word("val"): "1" if properties[key] else "0"}))
    grouped = (
        {}
        if paragraph
        else {
            "rFonts": _SCRIPT_FONT_ATTRIBUTES,
            "lang": _SCRIPT_LANGUAGE_ATTRIBUTES,
        }
    )
    for tag, attributes in grouped.items():
        edits = {
            word(attr): str(properties[key])
            for key, attr in attributes.items()
            if properties.get(key) is not None
        }
        if edits:
            element = parent.find(word(tag))
            if element is None:
                element = ET.SubElement(parent, word(tag))
            element.attrib.update(edits)
    scalars = (
        {"text_alignment": "textAlignment", "text_direction": "textDirection"}
        if paragraph
        else {
            "font_size_complex_script": "szCs",
            "emphasis_mark": "em",
            "east_asian_layout": "eastAsianLayout",
            "fit_text": "fitText",
        }
    )
    for key, tag in scalars.items():
        if properties.get(key) is None:
            continue
        for existing in list(parent.findall(word(tag))):
            parent.remove(existing)
        item = properties[key]
        attributes = item if isinstance(item, Mapping) else {"val": item}
        parent.append(
            ET.Element(
                word(tag),
                {
                    word(attr): ("1" if content else "0")
                    if isinstance(content, bool)
                    else str(content)
                    for attr, content in attributes.items()
                },
            )
        )


def _read_script_properties(
    parent: ET.Element | None, namespace: str, *, paragraph: bool
) -> dict[str, Any]:
    if parent is None:
        return {}
    def word(local):
        return Q(namespace, local)
    result: dict[str, Any] = {}
    flags = _SCRIPT_PARAGRAPH_FLAGS if paragraph else _SCRIPT_RUN_FLAGS
    for key, tag in flags.items():
        element = parent.find(word(tag))
        if element is not None:
            result[key] = element.get(word("val"), "1") not in {"0", "false", "off"}
    grouped = (
        {}
        if paragraph
        else {
            "rFonts": _SCRIPT_FONT_ATTRIBUTES,
            "lang": _SCRIPT_LANGUAGE_ATTRIBUTES,
        }
    )
    for tag, attributes in grouped.items():
        element = parent.find(word(tag))
        if element is not None:
            for key, attr in attributes.items():
                if word(attr) in element.attrib:
                    result[key] = element.get(word(attr))
    scalars = (
        {"text_alignment": "textAlignment", "text_direction": "textDirection"}
        if paragraph
        else {
            "font_size_complex_script": "szCs",
            "emphasis_mark": "em",
            "east_asian_layout": "eastAsianLayout",
            "fit_text": "fitText",
        }
    )
    for key, tag in scalars.items():
        element = parent.find(word(tag))
        if element is None:
            continue
        if key in {"east_asian_layout", "fit_text"}:
            attributes = {}
            for name, content in element.attrib.items():
                if not name.startswith(f"{{{namespace}}}"):
                    continue
                attr = _local_name(name)
                if attr in {"combine", "vert", "vertCompress"}:
                    attributes[attr] = content not in {"0", "false", "off"}
                elif attr == "id" or (key == "fit_text" and attr == "val"):
                    try:
                        attributes[attr] = int(content)
                    except ValueError:
                        attributes[attr] = content
                else:
                    attributes[attr] = content
            result[key] = attributes
        else:
            content = element.get(word("val"))
            if key == "font_size_complex_script" and content is not None:
                try:
                    content = int(content)
                except ValueError:
                    pass
            result[key] = content
    return result


@dataclass(frozen=True)
class StyleRequest:
    """Strict request for creating or updating one named style definition.

    ``clear_based_on`` removes an existing ``w:basedOn`` on update; passing
    ``based_on=False`` is the same request, mirroring python-docx
    ``style.base_style = None``.

    Script-specific keys include ``rtl``, ``complex_script``, ``bold_complex_script``,
    ``italic_complex_script``, ``font_size_complex_script``, ``language[_east_asia|_bidi]``,
    script ``font_face_*`` / ``font_theme_*``, ``font_hint``, ``emphasis_mark``,
    ``east_asian_layout`` and ``fit_text``. The latter two accept mappings with
    schema attribute names (for example ``{"id": 1, "combine": True}`` and
    ``{"val": 1440, "id": 7}``). Paragraph keys include ``bidi``, East Asian
    line-breaking flags, ``text_alignment`` and ``text_direction``. Use None to
    leave script properties unchanged; False writes an explicit off value.
    """

    source: Path
    source_sha256: str
    destination: Path
    structure_fingerprint_sha256: str
    action: str
    style_id: str
    type: str | None = None
    name: str | None = None
    based_on: str | None = None
    next: str | None = None
    run_properties: Mapping[str, Any] | None = None
    paragraph_properties: Mapping[str, Any] | None = None
    table_properties: Mapping[str, Any] | None = None
    cell_properties: Mapping[str, Any] | None = None
    row_properties: Mapping[str, Any] | None = None
    conditional_formatting: Sequence[dict[str, Any]] | None = None
    link: str | None = None
    clear_based_on: bool = False

    def __init__(
        self,
        source: str | Path,
        source_sha256: str,
        destination: str | Path,
        structure_fingerprint_sha256: str,
        action: str,
        style_id: str,
        type: str | None = None,
        name: str | None = None,
        based_on: str | Literal[False] | None = None,
        next: str | None = None,
        run_properties: Mapping[str, Any] | None = None,
        paragraph_properties: Mapping[str, Any] | None = None,
        table_properties: Mapping[str, Any] | None = None,
        cell_properties: Mapping[str, Any] | None = None,
        row_properties: Mapping[str, Any] | None = None,
        conditional_formatting: Mapping[str, Any] | Sequence[Any] | None = None,
        link: str | None = None,
        clear_based_on: bool = False,
    ) -> None:
        if based_on is False:
            based_on = None
            clear_based_on = True
        if not isinstance(clear_based_on, bool):
            raise TypeError("style_clear_based_on_invalid")
        if clear_based_on and based_on is not None:
            raise ValueError("style_based_on_clear_conflict")
        object.__setattr__(self, "source", Path(source))
        object.__setattr__(self, "source_sha256", str(source_sha256).lower())
        object.__setattr__(self, "destination", Path(destination))
        object.__setattr__(
            self,
            "structure_fingerprint_sha256",
            str(structure_fingerprint_sha256).lower(),
        )
        object.__setattr__(self, "action", str(action))
        object.__setattr__(self, "style_id", str(style_id))
        object.__setattr__(self, "type", str(type) if type is not None else None)
        object.__setattr__(self, "name", str(name) if name is not None else None)
        object.__setattr__(self, "based_on", str(based_on) if based_on is not None else None)
        object.__setattr__(self, "next", str(next) if next is not None else None)
        object.__setattr__(self, "link", str(link) if link is not None else None)
        object.__setattr__(self, "clear_based_on", clear_based_on)
        norm_run_properties = None
        if run_properties is not None:
            norm_run_properties = dict(run_properties)
            norm_run_properties.update(
                _validate_script_properties(norm_run_properties, paragraph=False)
            )
            highlight = norm_run_properties.get("highlight")
            if highlight is not None and highlight is not False:
                norm_run_properties["highlight"] = _canonical_highlight(
                    highlight, "style_highlight_invalid"
                )
        object.__setattr__(self, "run_properties", norm_run_properties)
        norm_paragraph_properties = None
        if paragraph_properties is not None:
            norm_paragraph_properties = dict(paragraph_properties)
            norm_paragraph_properties.update(
                _validate_script_properties(norm_paragraph_properties, paragraph=True)
            )
            _fold_paragraph_keep_aliases(norm_paragraph_properties)
            if norm_paragraph_properties.get("alignment") is not None:
                norm_paragraph_properties["alignment"] = _canonical_alignment(
                    norm_paragraph_properties["alignment"], "style_alignment_invalid"
                )
            if (
                "spacing" in norm_paragraph_properties
                and norm_paragraph_properties["spacing"] is not None
            ):
                sp = norm_paragraph_properties["spacing"]
                if isinstance(sp, Mapping) and not isinstance(sp, SpacingSpec):
                    norm_paragraph_properties["spacing"] = _validate_spacing_spec(sp)
            if (
                "indent" in norm_paragraph_properties
                and norm_paragraph_properties["indent"] is not None
            ):
                ind = norm_paragraph_properties["indent"]
                if isinstance(ind, Mapping) and not isinstance(ind, IndentSpec):
                    norm_paragraph_properties["indent"] = _validate_indent_spec(ind)
        object.__setattr__(self, "paragraph_properties", norm_paragraph_properties)
        norm_table_properties = None
        if table_properties is not None:
            if isinstance(table_properties, TablePropertiesSpec):
                norm_table_properties = table_properties.to_dict()
            else:
                norm_table_properties = dict(table_properties)
        object.__setattr__(self, "table_properties", norm_table_properties)

        norm_cell_properties = None
        if cell_properties is not None:
            if isinstance(cell_properties, TableCellPropertiesSpec):
                norm_cell_properties = cell_properties.to_dict()
            else:
                norm_cell_properties = dict(cell_properties)
        object.__setattr__(self, "cell_properties", norm_cell_properties)
        norm_row_properties = None
        if row_properties is not None:
            if isinstance(row_properties, TableRowPropertiesSpec):
                norm_row_properties = row_properties.to_dict()
            else:
                norm_row_properties = dict(row_properties)
        object.__setattr__(self, "row_properties", norm_row_properties)

        norm_conditional = None
        if conditional_formatting is not None:
            norm_conditional = _validate_conditional_formatting(conditional_formatting)
        object.__setattr__(self, "conditional_formatting", norm_conditional)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "source": str(self.source),
            "source_sha256": self.source_sha256,
            "destination": str(self.destination),
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "action": self.action,
            "style_id": self.style_id,
        }
        if self.type is not None:
            result["type"] = self.type
        if self.name is not None:
            result["name"] = self.name
        if self.based_on is not None:
            result["based_on"] = self.based_on
        if self.clear_based_on:
            result["clear_based_on"] = True
        if self.next is not None:
            result["next"] = self.next
        if self.link is not None:
            result["link"] = self.link
        if self.run_properties is not None:
            result["run_properties"] = dict(self.run_properties)
        if self.paragraph_properties is not None:
            p_dict: dict[str, Any] = {}
            for k, v in self.paragraph_properties.items():
                if hasattr(v, "to_dict"):
                    p_dict[k] = v.to_dict()
                else:
                    p_dict[k] = v
            result["paragraph_properties"] = p_dict
        if self.table_properties is not None:
            t_dict: dict[str, Any] = {}
            for k, v in self.table_properties.items():
                if hasattr(v, "to_dict"):
                    t_dict[k] = v.to_dict()
                else:
                    t_dict[k] = v
            result["table_properties"] = t_dict
        if self.cell_properties is not None:
            c_dict: dict[str, Any] = {}
            for k, v in self.cell_properties.items():
                if hasattr(v, "to_dict"):
                    c_dict[k] = v.to_dict()
                else:
                    c_dict[k] = v
            result["cell_properties"] = c_dict
        if self.row_properties is not None:
            r_dict: dict[str, Any] = {}
            for k, v in self.row_properties.items():
                if hasattr(v, "to_dict"):
                    r_dict[k] = v.to_dict()
                else:
                    r_dict[k] = v
            result["row_properties"] = r_dict
        if self.conditional_formatting is not None:
            result["conditional_formatting"] = [dict(ov) for ov in self.conditional_formatting]
        return result


@dataclass(frozen=True)
class StyleReport:
    """Content-safe public receipt for one atomic style authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    style: dict[str, Any] | None = None
    changed_targets: list[dict[str, Any]] = ()

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "schema_version": 1,
            "operation": "style",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
        }
        if self.style is not None:
            result["style"] = self.style
        if self.changed_targets:
            result["changed_targets"] = list(self.changed_targets)
        return result


def _style_refusal(request: StyleRequest, code: str) -> StyleReport:
    return StyleReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        style=None,
        changed_targets=[],
    )


def _validate_run_properties_dict(value: Any) -> dict[str, Any]:
    if not isinstance(value, Mapping):
        raise TypeError("style_run_properties_invalid")

    allowed_keys = _STYLE_RUN_KEYS
    unknown = set(value) - allowed_keys
    if unknown:
        raise ValueError(f"style_run_properties_unknown_keys: {sorted(unknown)}")

    validated: dict[str, Any] = _validate_script_properties(value, paragraph=False)

    bold = value.get("bold")
    if bold is not None:
        if not isinstance(bold, bool):
            raise ValueError("style_bold_invalid")
        validated["bold"] = bold

    italic = value.get("italic")
    if italic is not None:
        if not isinstance(italic, bool):
            raise ValueError("style_italic_invalid")
        validated["italic"] = italic

    underline = value.get("underline")
    if underline is not None:
        if not (
            isinstance(underline, bool)
            or (isinstance(underline, str) and underline in _VALID_UNDERLINE_STYLES)
        ):
            raise ValueError("style_underline_invalid")
        validated["underline"] = underline

    strike = value.get("strike")
    if strike is not None:
        if not isinstance(strike, bool):
            raise ValueError("style_strike_invalid")
        validated["strike"] = strike

    color = value.get("color")
    if color is not None:
        if not isinstance(color, str) or (
            color.lower() != "auto" and not _HEX_COLOR_PATTERN.match(color)
        ):
            raise ValueError("style_color_invalid")
        validated["color"] = color

    font_size = value.get("font_size")
    if font_size is not None:
        if isinstance(font_size, bool) or not isinstance(font_size, int) or font_size <= 0:
            raise ValueError("style_font_size_invalid")
        validated["font_size"] = font_size

    vertical_alignment = value.get("vertical_alignment")
    if vertical_alignment is not None:
        if (
            not isinstance(vertical_alignment, str)
            or vertical_alignment not in _VALID_VERTICAL_ALIGNMENTS
        ):
            raise ValueError("style_vertical_alignment_invalid")
        validated["vertical_alignment"] = vertical_alignment

    font_face = value.get("font_face")
    if font_face is not None:
        if not isinstance(font_face, str) or not font_face or len(font_face) > 256:
            raise ValueError("style_font_face_invalid")
        validated["font_face"] = font_face

    # False removes w:highlight; a colour, python-docx name, or WD_COLOR_INDEX member sets it.
    highlight = value.get("highlight")
    if highlight is not None:
        validated["highlight"] = (
            False
            if highlight is False
            else _canonical_highlight(highlight, "style_highlight_invalid")
        )

    style = value.get("style")
    if style is not None:
        if not isinstance(style, str) or not style or len(style) > 256:
            raise ValueError("style_reference_invalid")
        validated["style"] = style

    return validated


_PARAGRAPH_KEEP_ALIASES: tuple[tuple[str, str], ...] = (
    ("keep_next", "keep_with_next"),
    ("keep_lines", "keep_together"),
)


def _fold_paragraph_keep_aliases(properties: dict[str, Any]) -> None:
    """Rewrite python-docx ``keep_with_next``/``keep_together`` onto canonical keys in place."""
    for canonical, alias in _PARAGRAPH_KEEP_ALIASES:
        if alias in properties:
            flag = _fold_flag_alias(
                properties.get(canonical), properties.pop(alias), f"style_{canonical}"
            )
            if flag is not None:
                properties[canonical] = flag


def _validate_paragraph_properties_dict(value: Any) -> dict[str, Any]:
    if not isinstance(value, Mapping):
        raise TypeError("style_paragraph_properties_invalid")

    allowed_keys = _STYLE_PARAGRAPH_KEYS
    unknown = set(value) - allowed_keys
    if unknown:
        raise ValueError(f"style_paragraph_properties_unknown_keys: {sorted(unknown)}")
    value = dict(value)
    _fold_paragraph_keep_aliases(value)

    validated: dict[str, Any] = _validate_script_properties(value, paragraph=True)

    alignment = value.get("alignment")
    if alignment is not None:
        validated["alignment"] = _canonical_alignment(alignment, "style_alignment_invalid")

    if "spacing" in value and value["spacing"] is not None:
        validated["spacing"] = _validate_spacing_spec(value["spacing"])

    if "indent" in value and value["indent"] is not None:
        validated["indent"] = _validate_indent_spec(value["indent"])

    keep_next = value.get("keep_next")
    if keep_next is not None:
        if not isinstance(keep_next, bool):
            raise ValueError("style_keep_next_invalid")
        validated["keep_next"] = keep_next

    keep_lines = value.get("keep_lines")
    if keep_lines is not None:
        if not isinstance(keep_lines, bool):
            raise ValueError("style_keep_lines_invalid")
        validated["keep_lines"] = keep_lines

    outline_level = value.get("outline_level")
    if outline_level is not None:
        if (
            isinstance(outline_level, bool)
            or not isinstance(outline_level, int)
            or not (0 <= outline_level <= 9)
        ):
            raise ValueError("style_outline_level_invalid")
        validated["outline_level"] = outline_level

    style = value.get("style")
    if style is not None:
        if not isinstance(style, str) or not style or len(style) > 256:
            raise ValueError("style_reference_invalid")
        validated["style"] = style

    return validated


def _table_cell_properties_spec_mapping(
    properties: TableCellPropertiesSpec,
) -> dict[str, Any]:
    """Preserve typed cell-property values for the public validator."""
    return {
        field.name: value
        for field in fields(properties)
        if (value := getattr(properties, field.name)) is not None
    }


def _validate_conditional_override_dict(override: Any) -> dict[str, Any]:
    if isinstance(override, TableConditionalOverrideSpec):
        override = override.to_dict()
    if not isinstance(override, Mapping):
        raise TypeError("style_conditional_override_invalid")

    allowed_keys = {
        "type",
        "paragraph_properties",
        "run_properties",
        "table_properties",
        "row_properties",
        "cell_properties",
    }
    unknown = set(override) - allowed_keys
    if unknown:
        raise ValueError(f"style_conditional_override_unknown_keys: {sorted(unknown)}")

    raw_type = override.get("type")
    if not isinstance(raw_type, str) or not raw_type:
        raise ValueError("style_conditional_override_type_invalid")

    canonical_type = _OVERRIDE_TYPE_MAP.get(raw_type.lower(), raw_type)
    if canonical_type not in _VALID_TBL_STYLE_OVERRIDE_TYPES:
        raise ValueError(f"style_conditional_override_type_invalid: {raw_type}")

    result: dict[str, Any] = {"type": canonical_type}

    if "paragraph_properties" in override and override["paragraph_properties"] is not None:
        result["paragraph_properties"] = _validate_paragraph_properties_dict(
            override["paragraph_properties"]
        )

    if "run_properties" in override and override["run_properties"] is not None:
        result["run_properties"] = _validate_run_properties_dict(override["run_properties"])

    if "table_properties" in override and override["table_properties"] is not None:
        t_props = override["table_properties"]
        if isinstance(t_props, TablePropertiesSpec):
            result["table_properties"] = _validate_table_properties_dict(t_props.to_dict())
        elif isinstance(t_props, Mapping):
            result["table_properties"] = _validate_table_properties_dict(t_props)
        else:
            raise TypeError("style_table_properties_invalid")

    if "row_properties" in override and override["row_properties"] is not None:
        r_props = override["row_properties"]
        if isinstance(r_props, TableRowPropertiesSpec):
            result["row_properties"] = _validate_table_row_properties_dict(r_props.to_dict())
        elif isinstance(r_props, Mapping):
            result["row_properties"] = _validate_table_row_properties_dict(r_props)
        else:
            raise TypeError("style_row_properties_invalid")

    if "cell_properties" in override and override["cell_properties"] is not None:
        c_props = override["cell_properties"]
        if isinstance(c_props, TableCellPropertiesSpec):
            result["cell_properties"] = _validate_table_cell_properties_dict(
                _table_cell_properties_spec_mapping(c_props)
            )
        elif isinstance(c_props, Mapping):
            result["cell_properties"] = _validate_table_cell_properties_dict(c_props)
        else:
            raise TypeError("style_cell_properties_invalid")

    return result


def _validate_conditional_formatting(value: Any) -> list[dict[str, Any]]:
    if value is None:
        return []
    if isinstance(value, Mapping):
        if "type" in value or "override_type" in value:
            return [_validate_conditional_override_dict(value)]
        overrides: list[dict[str, Any]] = []
        for ov_type, ov_props in value.items():
            if isinstance(ov_props, Mapping):
                merged = dict(ov_props)
                merged.setdefault("type", ov_type)
                overrides.append(_validate_conditional_override_dict(merged))
            elif isinstance(ov_props, TableConditionalOverrideSpec):
                d = ov_props.to_dict()
                d.setdefault("type", ov_type)
                overrides.append(_validate_conditional_override_dict(d))
            else:
                raise TypeError("style_conditional_override_invalid")
        return overrides
    elif isinstance(value, (list, tuple)):
        return [_validate_conditional_override_dict(item) for item in value]
    else:
        raise TypeError("style_conditional_formatting_invalid")


# python-docx ``_ParagraphStyle`` property names for the style reference request keys.
_STYLE_REFERENCE_ALIASES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("based_on", ("base_style",)),
    ("next", ("next_paragraph_style", "next_style")),
)


def _fold_style_reference_aliases(value: Mapping[str, Any]) -> dict[str, Any]:
    """Rewrite python-docx style reference names onto canonical keys; disagreeing values refuse."""
    folded = dict(value)
    for canonical, aliases in _STYLE_REFERENCE_ALIASES:
        spelled = [folded.pop(alias) for alias in aliases if alias in folded]
        if canonical in folded:
            spelled.append(folded[canonical])
        given = [item for item in spelled if item is not None]
        if any(type(item) is not type(given[0]) or item != given[0] for item in given):
            raise ValueError(f"style_{canonical}_conflict")
        if given:
            folded[canonical] = given[0]
    return folded


def parse_style_request(value: Any) -> StyleRequest:
    """Parse the strict JSON request for creating or updating one style definition."""
    if not isinstance(value, Mapping):
        raise TypeError("style_request_invalid")

    allowed_keys = {
        "source",
        "source_sha256",
        "destination",
        "structure_fingerprint_sha256",
        "action",
        "style_id",
        "type",
        "name",
        "based_on",
        "next",
        "run_properties",
        "paragraph_properties",
        "table_properties",
        "cell_properties",
        "row_properties",
        "conditional_formatting",
        "link",
        "clear_based_on",
        *(alias for _, aliases in _STYLE_REFERENCE_ALIASES for alias in aliases),
    }
    unknown = set(value) - allowed_keys
    if unknown:
        if "tblStylePr" in unknown or "tbl_style_pr" in unknown:
            raise ValueError("style_tbl_style_pr_unsupported")
        raise ValueError(f"style_request_unknown_keys: {sorted(unknown)}")
    value = _fold_style_reference_aliases(value)

    clear_based_on = value.get("clear_based_on", False)
    if not isinstance(clear_based_on, bool):
        raise TypeError("style_clear_based_on_invalid")
    if value.get("based_on") is False:
        clear_based_on = True
    elif clear_based_on and value.get("based_on") is not None:
        raise ValueError("style_based_on_clear_conflict")

    action = value.get("action")
    if not isinstance(action, str) or action not in {"create", "update", "delete"}:
        raise ValueError("style_action_invalid")

    if action == "delete":
        contradictory_keys = [
            k
            for k in (
                "run_properties",
                "paragraph_properties",
                "table_properties",
                "cell_properties",
                "based_on",
                "link",
                "row_properties",
                "conditional_formatting",
            )
            if k in value and value[k] is not None
        ]
        if contradictory_keys or clear_based_on:
            raise ValueError("style_delete_formatting_unsupported")
        required_keys = {
            "source",
            "source_sha256",
            "destination",
            "structure_fingerprint_sha256",
            "action",
            "style_id",
        }
    else:
        required_keys = {
            "source",
            "source_sha256",
            "destination",
            "structure_fingerprint_sha256",
            "action",
            "style_id",
            "type",
            "name",
        }
    missing = required_keys - set(value)
    if missing:
        raise ValueError(f"style_request_missing_keys: {sorted(missing)}")

    source_raw = value["source"]
    dest_raw = value["destination"]
    if not isinstance(source_raw, str) or not source_raw:
        raise ValueError("style_source_path_invalid")
    if not isinstance(dest_raw, str) or not dest_raw:
        raise ValueError("style_destination_path_invalid")

    source_sha256 = _sha(value["source_sha256"], "style_source_sha256")
    structure_fingerprint_sha256 = _sha(
        value["structure_fingerprint_sha256"], "style_structure_fingerprint_sha256"
    )

    style_id = value["style_id"]
    if not isinstance(style_id, str) or not style_id or re.search(r"\s", style_id):
        raise ValueError("style_id_invalid")

    stype = value.get("type")
    if stype is not None:
        if not isinstance(stype, str) or stype not in {
            "paragraph",
            "character",
            "table",
        }:
            raise ValueError("style_type_unsupported")
    elif action != "delete":
        raise ValueError("style_type_unsupported")

    name = value.get("name")
    if name is not None:
        if not isinstance(name, str) or not name.strip():
            raise ValueError("style_name_invalid")
    elif action != "delete":
        raise ValueError("style_name_invalid")

    based_on = None if clear_based_on else value.get("based_on")
    if based_on is not None and (not isinstance(based_on, str) or not based_on.strip()):
        raise ValueError("style_based_on_invalid")
    if clear_based_on and action != "update":
        raise ValueError("style_based_on_clear_requires_update")

    next_style = value.get("next")
    if next_style is not None:
        if stype == "character":
            raise ValueError("style_next_unsupported_for_character_style")
        if stype == "table":
            raise ValueError("style_next_unsupported_for_table_style")
        if not isinstance(next_style, str) or not next_style.strip():
            raise ValueError("style_next_invalid")

    link_style = value.get("link")
    if link_style is not None:
        if stype == "table":
            raise ValueError("style_link_unsupported_for_table_style")
        if not isinstance(link_style, str) or not link_style or re.search(r"\s", link_style):
            raise ValueError("style_link_invalid")
        if link_style == style_id:
            raise ValueError("style_link_self_reference")

    run_props = None
    if "run_properties" in value and value["run_properties"] is not None:
        run_props = _validate_run_properties_dict(value["run_properties"])

    para_props = None
    if "paragraph_properties" in value and value["paragraph_properties"] is not None:
        if stype == "character":
            raise ValueError("style_paragraph_properties_unsupported_for_character_style")
        if stype == "table":
            raise ValueError("style_paragraph_properties_unsupported_for_table_style")
        para_props = _validate_paragraph_properties_dict(value["paragraph_properties"])

    table_props = None
    if "table_properties" in value and value["table_properties"] is not None:
        if stype != "table":
            raise ValueError(f"style_table_properties_unsupported_for_{stype}_style")
        if not isinstance(value["table_properties"], Mapping):
            raise TypeError("style_table_properties_invalid")
        if "tblStylePr" in value["table_properties"] or "tbl_style_pr" in value["table_properties"]:
            raise ValueError("style_tbl_style_pr_unsupported")
        table_props = _validate_table_properties_dict(value["table_properties"])

    cell_props = None
    if "cell_properties" in value and value["cell_properties"] is not None:
        if stype != "table":
            raise ValueError(f"style_cell_properties_unsupported_for_{stype}_style")
        if not isinstance(value["cell_properties"], Mapping):
            raise TypeError("style_cell_properties_invalid")
        if "cnfStyle" in value["cell_properties"] or "tblStylePr" in value["cell_properties"]:
            raise ValueError("style_tbl_style_pr_unsupported")
        cell_props = _validate_table_cell_properties_dict(value["cell_properties"])
    row_props = None
    if "row_properties" in value and value["row_properties"] is not None:
        if stype != "table":
            raise ValueError(f"style_row_properties_unsupported_for_{stype}_style")
        if isinstance(value["row_properties"], TableRowPropertiesSpec):
            row_props = _validate_table_row_properties_dict(value["row_properties"].to_dict())
        elif isinstance(value["row_properties"], Mapping):
            row_props = _validate_table_row_properties_dict(value["row_properties"])
        else:
            raise TypeError("style_row_properties_invalid")

    cond_raw = value.get("conditional_formatting")
    cond_formatting = None
    if cond_raw is not None:
        if stype != "table":
            raise ValueError(f"style_conditional_formatting_unsupported_for_{stype}_style")
        cond_formatting = _validate_conditional_formatting(cond_raw)

    return StyleRequest(
        source=Path(source_raw),
        source_sha256=source_sha256,
        destination=Path(dest_raw),
        structure_fingerprint_sha256=structure_fingerprint_sha256,
        action=action,
        style_id=style_id,
        type=stype,
        name=name,
        based_on=based_on,
        next=next_style,
        run_properties=run_props,
        paragraph_properties=para_props,
        table_properties=table_props,
        cell_properties=cell_props,
        row_properties=row_props,
        conditional_formatting=cond_formatting,
        link=link_style,
        clear_based_on=clear_based_on,
    )


def load_style_request(path: str | Path) -> StyleRequest:
    """Load and strictly validate a style authoring request from JSON on disk."""
    with Path(path).open(encoding="utf-8") as stream:
        return parse_style_request(json.load(stream))


def _element_structure_fingerprint(element: ET.Element) -> str:
    """Hash the XML element structure deterministically."""

    def project(elem: ET.Element) -> list[Any]:
        return [
            elem.tag,
            sorted(elem.attrib.items()),
            elem.text,
            elem.tail,
            [project(child) for child in elem],
        ]

    canonical = json.dumps(project(element), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def style_structure_fingerprint(
    source: str | Path,
    *,
    style_id: str | None = None,
    part: str = _STYLES,
) -> str | None:
    """Return the structural fingerprint for styles.xml or a specific style definition."""
    path = Path(source)
    if not path.is_file():
        return None

    try:
        graph = PackageGraph.from_file(path)
    except (OSError, PackageError, zipfile.BadZipFile):
        return None

    if part not in graph.parts:
        return _ABSENT_STYLES_FINGERPRINT if style_id is None else None

    try:
        styles_root = ET.fromstring(graph.read_part_bytes(part))
    except ET.ParseError:
        return None

    if style_id is None:
        return _element_structure_fingerprint(styles_root)

    ns = (
        _word_namespace(styles_root)
        or "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    )
    def word(local):
        return Q(ns, local)
    for style_elem in styles_root.findall(word("style")):
        sid = style_elem.get(word("styleId")) or style_elem.get("styleId")
        if sid == style_id:
            return _element_structure_fingerprint(style_elem)

    return None


def _find_style_references(
    parts: Mapping[str, bytes],
    style_id: str,
    styles_root: ET.Element | None,
    styles_ns: str,
) -> list[str]:
    """Inspect all package story references, numbering links, and style graph edges."""
    references: list[str] = []

    # 1. Styles graph edges and properties in styles.xml
    if styles_root is not None:
        style_ref_tags = {
            "basedOn",
            "next",
            "link",
            "rStyle",
            "pStyle",
            "tblStyle",
            "styleLink",
        }
        for elem in styles_root.findall(Q(styles_ns, "style")):
            sid = elem.get(Q(styles_ns, "styleId")) or elem.get("styleId")
            if sid == style_id:
                # Exclude target self-reference
                continue

            for node in elem.iter():
                if node is elem:
                    continue
                local = _local_name(node.tag)
                if local in style_ref_tags:
                    for attr_name, attr_val in node.attrib.items():
                        if _local_name(attr_name) == "val" and attr_val == style_id:
                            references.append(f"{_STYLES}:style:{sid}:{local}")
                            break

        doc_defaults = styles_root.find(Q(styles_ns, "docDefaults"))
        if doc_defaults is None:
            doc_defaults = styles_root.find("docDefaults")
        if doc_defaults is not None:
            for node in doc_defaults.iter():
                if node is doc_defaults:
                    continue
                local = _local_name(node.tag)
                if local in style_ref_tags:
                    for attr_name, attr_val in node.attrib.items():
                        if _local_name(attr_name) == "val" and attr_val == style_id:
                            references.append(f"{_STYLES}:docDefaults:{local}")
                            break
    # 2. Inspect all other XML parts in the package
    target_tags = {
        "pStyle",
        "rStyle",
        "tblStyle",
        "cnfStyle",
        "styleLink",
        "numStyleLink",
    }
    for part_name, raw_bytes in parts.items():
        if part_name == _STYLES:
            continue
        if not part_name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(raw_bytes)
        except ET.ParseError:
            continue

        for node in root.iter():
            local = _local_name(node.tag)
            if local in target_tags:
                for attr_name, attr_val in node.attrib.items():
                    if _local_name(attr_name) == "val" and attr_val == style_id:
                        references.append(f"{part_name}:{local}")
                        break

    return references


def discover_style_targets(source: str | Path, *, part: str = _STYLES) -> dict[str, Any]:
    """List source-bound style targets and structure fingerprints without writing."""
    path = Path(source)
    if not path.is_file():
        return {
            "passed": False,
            "code": "style_source_missing",
            "source_sha256": "",
            "part": part,
            "styles_part_exists": False,
            "structure_fingerprint_sha256": "",
            "styles": [],
            "unsupported": [],
        }

    try:
        graph = PackageGraph.from_file(path)
    except (OSError, PackageError, zipfile.BadZipFile):
        return {
            "passed": False,
            "code": "style_package_invalid",
            "source_sha256": "",
            "part": part,
            "styles_part_exists": False,
            "structure_fingerprint_sha256": "",
            "styles": [],
            "unsupported": [],
        }

    styles_part_exists = part in graph.parts
    if not styles_part_exists:
        return {
            "passed": True,
            "code": None,
            "source_sha256": graph.file_sha256,
            "part": part,
            "styles_part_exists": False,
            "structure_fingerprint_sha256": _ABSENT_STYLES_FINGERPRINT,
            "styles": [],
            "unsupported": [],
        }

    try:
        styles_bytes = graph.read_part_bytes(part)
        styles_root = ET.fromstring(styles_bytes)
    except (ET.ParseError, PackageError):
        return {
            "passed": False,
            "code": "style_part_malformed",
            "source_sha256": graph.file_sha256,
            "part": part,
            "styles_part_exists": True,
            "structure_fingerprint_sha256": "",
            "styles": [],
            "unsupported": [],
        }

    ns = (
        _word_namespace(styles_root)
        or "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    )
    def word(local):
        return Q(ns, local)

    root_fp = _element_structure_fingerprint(styles_root)
    try:
        all_parts: dict[str, bytes] = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError, zipfile.BadZipFile):
        all_parts = {}

    styles_list: list[dict[str, Any]] = []
    unsupported: list[dict[str, Any]] = []

    for style_elem in styles_root.findall(word("style")):
        sid = style_elem.get(word("styleId")) or style_elem.get("styleId")
        stype = style_elem.get(word("type")) or style_elem.get("type") or "paragraph"

        name_elem = style_elem.find(word("name"))
        sname = None
        if name_elem is not None:
            sname = name_elem.get(word("val")) or name_elem.get("val")

        based_on_elem = style_elem.find(word("basedOn"))
        based_on = None
        if based_on_elem is not None:
            based_on = based_on_elem.get(word("val")) or based_on_elem.get("val")

        next_elem = style_elem.find(word("next"))
        next_style = None
        if next_elem is not None:
            next_style = next_elem.get(word("val")) or next_elem.get("val")

        is_default = (style_elem.get(word("default")) or style_elem.get("default")) in {
            "1",
            "true",
            "on",
        }
        is_custom = (style_elem.get(word("customStyle")) or style_elem.get("customStyle")) in {
            "1",
            "true",
            "on",
        }

        entry: dict[str, Any] = {
            "style_id": sid,
            "type": stype,
            "name": sname,
            "based_on": based_on,
            "next": next_style,
            "is_default": is_default,
            "is_custom": is_custom,
            "has_run_properties": style_elem.find(word("rPr")) is not None,
            "has_paragraph_properties": style_elem.find(word("pPr")) is not None,
            "has_table_properties": style_elem.find(word("tblPr")) is not None,
            "has_cell_properties": style_elem.find(word("tcPr")) is not None,
            "has_conditional_formatting": style_elem.find(word("tblStylePr")) is not None,
            "conditional_formatting_types": [
                elem.get(word("type")) or elem.get("type")
                for elem in style_elem.findall(word("tblStylePr"))
                if elem.get(word("type")) or elem.get("type")
            ],
            "structure_fingerprint_sha256": _element_structure_fingerprint(style_elem),
        }

        entry["run_properties"] = _read_script_properties(
            style_elem.find(word("rPr")), ns, paragraph=False
        )
        entry["paragraph_properties"] = _read_script_properties(
            style_elem.find(word("pPr")), ns, paragraph=True
        )
        refs = _find_style_references(all_parts, sid, styles_root, ns) if sid else []
        in_use = len(refs) > 0
        refusal_reason = None
        if is_default:
            refusal_reason = "style_default_cannot_be_deleted"
        elif not is_custom:
            refusal_reason = "style_builtin_cannot_be_deleted"
        elif in_use:
            refusal_reason = "style_in_use"

        can_delete = refusal_reason is None
        entry["can_delete"] = can_delete
        entry["deletable"] = can_delete
        entry["refusal_reason"] = refusal_reason
        entry["in_use"] = in_use
        entry["references_count"] = len(refs)
        entry["references"] = refs

        if stype not in {"paragraph", "character", "table"}:
            unsupported.append(entry)
        else:
            styles_list.append(entry)

    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        "part": part,
        "styles_part_exists": True,
        "structure_fingerprint_sha256": root_fp,
        "styles": styles_list,
        "unsupported": unsupported,
    }


def _detect_style_cycle(
    existing_styles: Mapping[str, dict[str, Any]],
    target_id: str,
    new_based_on: str | None,
) -> tuple[bool, str | None]:
    """Check for inheritance (based_on) cycles using graph traversal."""
    based_on_graph: dict[str, str | None] = {}
    for sid, info in existing_styles.items():
        based_on_graph[sid] = info.get("based_on")
    based_on_graph[target_id] = new_based_on

    for start_node in based_on_graph:
        visited: set[str] = set()
        curr: str | None = start_node
        while curr is not None:
            if curr in visited:
                return True, "style_cycle_detected"
            visited.add(curr)
            curr = based_on_graph.get(curr)

    return False, None


def _apply_style_paragraph_properties(
    style_elem: ET.Element,
    properties: Mapping[str, Any],
    namespace: str,
) -> None:
    def word(local):
        return Q(namespace, local)
    ppr = style_elem.find(word("pPr"))
    if ppr is None:
        ppr = ET.Element(word("pPr"))
        style_elem.append(ppr)

    if "alignment" in properties:
        for existing in list(ppr.findall(word("jc"))):
            ppr.remove(existing)
        ppr.append(ET.Element(word("jc"), {word("val"): properties["alignment"]}))

    if "spacing" in properties:
        raw_spacing = properties["spacing"]
        if isinstance(raw_spacing, Mapping) and not isinstance(raw_spacing, SpacingSpec):
            spacing = _validate_spacing_spec(raw_spacing)
        else:
            spacing = raw_spacing
        for existing in list(ppr.findall(word("spacing"))):
            ppr.remove(existing)
        attrib: dict[str, str] = {}
        if spacing.before is not None:
            attrib[word("before")] = str(spacing.before)
        if spacing.after is not None:
            attrib[word("after")] = str(spacing.after)
        if spacing.line is not None:
            attrib[word("line")] = str(spacing.line)
        if spacing.line_rule is not None:
            rule_map = {"auto": "auto", "exact": "exact", "at_least": "atLeast"}
            attrib[word("lineRule")] = rule_map.get(spacing.line_rule, spacing.line_rule)
        ppr.append(ET.Element(word("spacing"), attrib))
    if "indent" in properties:
        raw_indent = properties["indent"]
        if isinstance(raw_indent, Mapping) and not isinstance(raw_indent, IndentSpec):
            indent = _validate_indent_spec(raw_indent)
        else:
            indent = raw_indent
        for existing in list(ppr.findall(word("ind"))):
            ppr.remove(existing)
        attrib = {}
        if indent.left is not None:
            attrib[word("left")] = str(indent.left)
        if indent.right is not None:
            attrib[word("right")] = str(indent.right)
        if indent.first_line is not None:
            attrib[word("firstLine")] = str(indent.first_line)
        if indent.hanging is not None:
            attrib[word("hanging")] = str(indent.hanging)
        if indent.start is not None:
            attrib[word("start")] = str(indent.start)
        if indent.end is not None:
            attrib[word("end")] = str(indent.end)
        ppr.append(ET.Element(word("ind"), attrib))
    if "keep_next" in properties:
        for existing in list(ppr.findall(word("keepNext"))):
            ppr.remove(existing)
        if properties["keep_next"]:
            ppr.append(ET.Element(word("keepNext")))
        else:
            ppr.append(ET.Element(word("keepNext"), {word("val"): "0"}))

    if "keep_lines" in properties:
        for existing in list(ppr.findall(word("keepLines"))):
            ppr.remove(existing)
        if properties["keep_lines"]:
            ppr.append(ET.Element(word("keepLines")))
        else:
            ppr.append(ET.Element(word("keepLines"), {word("val"): "0"}))

    if "outline_level" in properties:
        for existing in list(ppr.findall(word("outlineLvl"))):
            ppr.remove(existing)
        ppr.append(ET.Element(word("outlineLvl"), {word("val"): str(properties["outline_level"])}))
    _write_script_properties(ppr, properties, namespace, paragraph=True)

    # Sort children of pPr by schema order
    ppr[:] = sorted(ppr, key=lambda c: _PPR_RANK.get(_local_name(c.tag), 999))


def _apply_style_run_properties(
    style_elem: ET.Element,
    properties: Mapping[str, Any],
    namespace: str,
) -> None:
    def word(local):
        return Q(namespace, local)
    rpr = style_elem.find(word("rPr"))
    if rpr is None:
        rpr = ET.Element(word("rPr"))
        style_elem.append(rpr)

    if "bold" in properties:
        for existing in list(rpr.findall(word("b"))):
            rpr.remove(existing)
        if properties["bold"]:
            rpr.append(ET.Element(word("b")))
        else:
            rpr.append(ET.Element(word("b"), {word("val"): "0"}))

    if "italic" in properties:
        for existing in list(rpr.findall(word("i"))):
            rpr.remove(existing)
        if properties["italic"]:
            rpr.append(ET.Element(word("i")))
        else:
            rpr.append(ET.Element(word("i"), {word("val"): "0"}))

    if "strike" in properties:
        for existing in list(rpr.findall(word("strike"))):
            rpr.remove(existing)
        if properties["strike"]:
            rpr.append(ET.Element(word("strike")))
        else:
            rpr.append(ET.Element(word("strike"), {word("val"): "0"}))

    if "underline" in properties:
        for existing in list(rpr.findall(word("u"))):
            rpr.remove(existing)
        u_val = properties["underline"]
        val = "single" if u_val is True else ("none" if u_val is False else str(u_val))
        rpr.append(ET.Element(word("u"), {word("val"): val}))

    if "color" in properties:
        for existing in list(rpr.findall(word("color"))):
            rpr.remove(existing)
        val = properties["color"].strip()
        if val.lower() != "auto":
            val = val.upper()
        rpr.append(ET.Element(word("color"), {word("val"): val}))

    if "font_size" in properties:
        for existing in list(rpr.findall(word("sz"))):
            rpr.remove(existing)
        for existing in list(rpr.findall(word("szCs"))):
            rpr.remove(existing)
        val = str(properties["font_size"])
        rpr.append(ET.Element(word("sz"), {word("val"): val}))
        rpr.append(ET.Element(word("szCs"), {word("val"): val}))

    if "vertical_alignment" in properties:
        for existing in list(rpr.findall(word("vertAlign"))):
            rpr.remove(existing)
        rpr.append(ET.Element(word("vertAlign"), {word("val"): properties["vertical_alignment"]}))

    if "font_face" in properties:
        for existing in list(rpr.findall(word("rFonts"))):
            rpr.remove(existing)
        font_face = properties["font_face"]
        rpr.append(
            ET.Element(
                word("rFonts"),
                {
                    word("ascii"): font_face,
                    word("hAnsi"): font_face,
                    word("cs"): font_face,
                },
            )
        )

    highlight = properties.get("highlight")
    if highlight is not None:
        for existing in list(rpr.findall(word("highlight"))):
            rpr.remove(existing)
        if highlight is not False:
            rpr.append(ET.Element(word("highlight"), {word("val"): highlight}))
    _write_script_properties(rpr, properties, namespace, paragraph=False)

    # Sort children of rPr by schema order
    rpr[:] = sorted(rpr, key=lambda c: _RPR_RANK.get(_local_name(c.tag), 999))


def _style_table_properties_supported(properties: Mapping[str, Any]) -> bool:
    # tblpPr governs one table instance's placement and is not legal table-style
    # formatting; reject it rather than accepting a request we cannot serialize.
    if "floating" in properties:
        return False
    return set(properties).issubset(_STYLE_TABLE_PROPERTY_KEYS)


def _apply_style_table_properties(
    style_elem: ET.Element,
    properties: Mapping[str, Any],
    namespace: str,
) -> None:
    def word(local):
        return Q(namespace, local)
    tbl_pr = style_elem.find(word("tblPr"))
    if tbl_pr is None:
        tbl_pr = ET.Element(word("tblPr"))
        style_elem.append(tbl_pr)

    props = _validate_table_properties_dict(properties)

    if "alignment" in props:
        jc = ET.Element(word("jc"))
        jc.set(word("val"), props["alignment"])
        _insert_or_replace_tblpr_child(tbl_pr, jc, namespace)

    if "borders" in props:
        borders_spec = props["borders"]
        existing_borders = tbl_pr.find(word("tblBorders"))
        existing_sides: dict[str, ET.Element] = {}
        if existing_borders is not None:
            for child in existing_borders:
                existing_sides[_local_name(child.tag)] = child

        new_borders = ET.Element(word("tblBorders"))
        for side in _TBL_BORDERS_ORDER:
            if side in borders_spec:
                spec = borders_spec[side]
                border = ET.Element(word(side))
                border.set(word("val"), spec["val"])
                if spec.get("sz") is not None:
                    border.set(word("sz"), str(spec["sz"]))
                if spec.get("space") is not None:
                    border.set(word("space"), str(spec["space"]))
                if spec.get("color") is not None:
                    border.set(word("color"), spec["color"])
                new_borders.append(border)
            elif side in existing_sides:
                new_borders.append(copy.deepcopy(existing_sides[side]))
        for side, child in existing_sides.items():
            if side not in _TBL_BORDERS_ORDER:
                new_borders.append(copy.deepcopy(child))
        _insert_or_replace_tblpr_child(tbl_pr, new_borders, namespace)

    if "shading" in props:
        shading = props["shading"]
        shd = ET.Element(word("shd"))
        shd.set(word("val"), shading.get("val", "clear"))
        shd.set(word("color"), shading.get("color", "auto"))
        shd.set(word("fill"), shading["fill"])
        _insert_or_replace_tblpr_child(tbl_pr, shd, namespace)

    if "cell_margins" in props:
        margins = props["cell_margins"]
        existing_margins = tbl_pr.find(word("tblCellMar"))
        existing_sides: dict[str, ET.Element] = {}
        if existing_margins is not None:
            for child in existing_margins:
                existing_sides[_local_name(child.tag)] = child

        new_margins = ET.Element(word("tblCellMar"))
        for side in _TBL_CELL_MAR_ORDER:
            if side in margins:
                margin = ET.Element(word(side))
                margin.set(word("w"), str(margins[side]))
                margin.set(word("type"), "dxa")
                new_margins.append(margin)
            elif side in existing_sides:
                new_margins.append(copy.deepcopy(existing_sides[side]))
        for side, child in existing_sides.items():
            if side not in _TBL_CELL_MAR_ORDER:
                new_margins.append(copy.deepcopy(child))
        _insert_or_replace_tblpr_child(tbl_pr, new_margins, namespace)


def _apply_style_cell_properties(
    style_elem: ET.Element,
    properties: Mapping[str, Any],
    namespace: str,
) -> None:
    def word(local):
        return Q(namespace, local)
    tc_pr = style_elem.find(word("tcPr"))
    if tc_pr is None:
        tc_pr = ET.Element(word("tcPr"))
        style_elem.append(tc_pr)

    props = _validate_table_cell_properties_dict(properties)

    if "width" in props:
        w_spec = props["width"]
        tc_w = ET.Element(word("tcW"))
        tc_w.set(word("w"), str(w_spec["value"]))
        tc_w.set(word("type"), w_spec["type"])
        _insert_or_replace_tcpr_child(tc_pr, tc_w, namespace)

    if "borders" in props:
        borders_spec = props["borders"]
        existing_borders = tc_pr.find(word("tcBorders"))
        existing_sides: dict[str, ET.Element] = {}
        if existing_borders is not None:
            for c in existing_borders:
                local_side = c.tag.rsplit("}", 1)[-1]
                existing_sides[local_side] = c

        new_borders = ET.Element(word("tcBorders"))
        for side in _TC_BORDERS_ORDER:
            if side in borders_spec:
                spec = borders_spec[side]
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

    if "shading" in props:
        shd_spec = props["shading"]
        shd = ET.Element(word("shd"))
        shd.set(word("val"), shd_spec.get("val", "clear"))
        shd.set(word("color"), shd_spec.get("color", "auto"))
        shd.set(word("fill"), shd_spec["fill"])
        _insert_or_replace_tcpr_child(tc_pr, shd, namespace)

    if "cell_margins" in props:
        mar_spec = props["cell_margins"]
        existing_mar = tc_pr.find(word("tcMar"))
        existing_mar_sides: dict[str, ET.Element] = {}
        if existing_mar is not None:
            for c in existing_mar:
                local_side = c.tag.rsplit("}", 1)[-1]
                existing_mar_sides[local_side] = c

        new_mar = ET.Element(word("tcMar"))
        for side in _TC_CELL_MAR_ORDER:
            if side in mar_spec:
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

    if "vertical_alignment" in props:
        align_val = props["vertical_alignment"]
        v_align = ET.Element(word("vAlign"))
        v_align.set(word("val"), align_val)
        _insert_or_replace_tcpr_child(tc_pr, v_align, namespace)


def _apply_style_row_properties(
    style_elem: ET.Element,
    properties: Mapping[str, Any],
    namespace: str,
) -> None:
    def word(local):
        return Q(namespace, local)
    tr_pr = style_elem.find(word("trPr"))
    if tr_pr is None:
        tr_pr = ET.Element(word("trPr"))
        style_elem.append(tr_pr)

    props = _validate_table_row_properties_dict(properties)

    if "cant_split" in props:
        val = props["cant_split"]
        if val is True:
            elem = ET.Element(word("cantSplit"))
            _insert_or_replace_trpr_child(tr_pr, elem, namespace)
        elif val is False:
            elem = ET.Element(word("cantSplit"), {word("val"): "off"})
            _insert_or_replace_trpr_child(tr_pr, elem, namespace)

    if "height" in props:
        h_spec = props["height"]
        elem = ET.Element(word("trHeight"))
        elem.set(word("val"), str(h_spec["value"]))
        elem.set(word("hRule"), h_spec["rule"])
        _insert_or_replace_trpr_child(tr_pr, elem, namespace)

    if "repeat_header" in props:
        val = props["repeat_header"]
        if val is True:
            elem = ET.Element(word("tblHeader"))
            _insert_or_replace_trpr_child(tr_pr, elem, namespace)
        elif val is False:
            elem = ET.Element(word("tblHeader"), {word("val"): "off"})
            _insert_or_replace_trpr_child(tr_pr, elem, namespace)


def _apply_style_conditional_formatting(
    style_elem: ET.Element,
    conditional_formatting: Sequence[Mapping[str, Any]],
    namespace: str,
) -> None:
    def word(local):
        return Q(namespace, local)

    existing_by_type: dict[str, ET.Element] = {}
    for elem in list(style_elem.findall(word("tblStylePr"))):
        ov_type = elem.get(word("type")) or elem.get("type")
        if ov_type:
            existing_by_type[ov_type] = elem

    for override in conditional_formatting:
        ov_type = override["type"]
        tbl_style_pr = existing_by_type.get(ov_type)
        if tbl_style_pr is None:
            tbl_style_pr = ET.Element(word("tblStylePr"), {word("type"): ov_type})
            style_elem.append(tbl_style_pr)
            existing_by_type[ov_type] = tbl_style_pr

        if "paragraph_properties" in override and override["paragraph_properties"] is not None:
            _apply_style_paragraph_properties(
                tbl_style_pr, override["paragraph_properties"], namespace
            )

        if "run_properties" in override and override["run_properties"] is not None:
            _apply_style_run_properties(tbl_style_pr, override["run_properties"], namespace)

        if "table_properties" in override and override["table_properties"] is not None:
            _apply_style_table_properties(tbl_style_pr, override["table_properties"], namespace)

        if "row_properties" in override and override["row_properties"] is not None:
            _apply_style_row_properties(tbl_style_pr, override["row_properties"], namespace)

        if "cell_properties" in override and override["cell_properties"] is not None:
            _apply_style_cell_properties(tbl_style_pr, override["cell_properties"], namespace)

        tbl_style_pr[:] = sorted(
            tbl_style_pr,
            key=lambda c: _TBL_STYLE_PR_CHILD_RANK.get(_local_name(c.tag), 999),
        )


def author_style(request: StyleRequest) -> StyleReport:
    """Execute source-bound style definition authoring atomically."""
    source = request.source
    destination = request.destination

    if not source.is_file():
        return _style_refusal(request, "style_source_not_found")

    if _sha256(source) != request.source_sha256:
        return _style_refusal(request, "style_source_sha256_mismatch")

    if destination.exists() or destination.is_symlink():
        return _style_refusal(request, "style_destination_invalid")

    if not destination.parent.is_dir():
        return _style_refusal(request, "style_destination_invalid")

    if request.action not in {"create", "update", "delete"}:
        return _style_refusal(request, "style_action_invalid")

    if request.type is not None and request.type not in {
        "paragraph",
        "character",
        "table",
    }:
        return _style_refusal(request, "style_type_unsupported")

    if request.action == "delete":
        if (
            request.run_properties is not None
            or request.paragraph_properties is not None
            or request.table_properties is not None
            or request.cell_properties is not None
            or request.based_on is not None
            or request.next is not None
            or request.link is not None
            or request.clear_based_on
        ):
            return _style_refusal(request, "style_properties_incompatible")
    elif request.clear_based_on and request.action != "update":
        return _style_refusal(request, "style_based_on_clear_requires_update")
    else:
        if request.type not in {"paragraph", "character", "table"}:
            return _style_refusal(request, "style_type_unsupported")

        if request.link is not None:
            if request.type == "table":
                return _style_refusal(request, "style_link_unsupported")
            if request.link == request.style_id:
                return _style_refusal(request, "style_link_self_reference")

        if request.type == "character":
            if request.next is not None:
                return _style_refusal(request, "style_next_unsupported")
            if request.paragraph_properties is not None:
                return _style_refusal(request, "style_properties_incompatible")
            if request.table_properties is not None or request.cell_properties is not None:
                return _style_refusal(request, "style_properties_incompatible")

        if request.type == "paragraph" and (
            request.table_properties is not None or request.cell_properties is not None
        ):
            return _style_refusal(request, "style_properties_incompatible")

        if request.type == "table":
            if request.next is not None:
                return _style_refusal(request, "style_next_unsupported")
            if request.paragraph_properties is not None:
                return _style_refusal(request, "style_properties_incompatible")
            if request.table_properties is not None and not _style_table_properties_supported(
                request.table_properties
            ):
                return _style_refusal(request, "style_table_properties_unsupported_for_table_style")
            if any(
                "table_properties" in override
                and not _style_table_properties_supported(override["table_properties"])
                for override in request.conditional_formatting or ()
            ):
                return _style_refusal(
                    request,
                    "style_conditional_table_properties_unsupported",
                )
            if request.cell_properties is not None and (
                "cnfStyle" in request.cell_properties or "tblStylePr" in request.cell_properties
            ):
                return _style_refusal(request, "style_tbl_style_pr_unsupported")
    try:
        graph = PackageGraph.from_file(source)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError, zipfile.BadZipFile):
        return _style_refusal(request, "style_package_invalid")

    if _DOCUMENT not in parts or "[Content_Types].xml" not in parts:
        return _style_refusal(request, "style_package_invalid")

    try:
        doc_root = ET.fromstring(parts[_DOCUMENT])
    except ET.ParseError:
        return _style_refusal(request, "style_package_invalid")

    main_ns = _word_namespace(doc_root)
    if main_ns is None or main_ns not in SUPPORTED_WORD_NAMESPACES:
        return _style_refusal(request, "style_namespace_unsupported")

    # Check styles.xml presence and parse existing styles
    styles_part_exists = _STYLES in parts
    styles_bytes: bytes | None = None
    if styles_part_exists:
        try:
            styles_bytes = parts[_STYLES]
            styles_root = parse_xml_with_preserved_qnames(styles_bytes)
        except ET.ParseError:
            return _style_refusal(request, "style_package_invalid")
        styles_ns = _word_namespace(styles_root)
        if styles_ns is None:
            return _style_refusal(request, "style_namespace_unsupported")
    else:
        styles_ns = main_ns
        styles_root = ET.Element(Q(styles_ns, "styles"))

    def styles_word(local):
        return Q(styles_ns, local)

    # Verify structural fingerprint
    current_root_fp = (
        _element_structure_fingerprint(styles_root)
        if styles_part_exists
        else _ABSENT_STYLES_FINGERPRINT
    )

    if request.action == "create":
        if request.structure_fingerprint_sha256 != current_root_fp:
            return _style_refusal(request, "style_structure_fingerprint_mismatch")
    else:  # update
        current_target_fp = style_structure_fingerprint(source, style_id=request.style_id)
        if (
            request.structure_fingerprint_sha256 != current_target_fp
            and request.structure_fingerprint_sha256 != current_root_fp
        ):
            return _style_refusal(request, "style_structure_fingerprint_mismatch")

    # Catalog existing styles
    existing_styles: dict[str, dict[str, Any]] = {}
    for elem in styles_root.findall(styles_word("style")):
        sid = elem.get(styles_word("styleId")) or elem.get("styleId")
        if not sid:
            continue
        stype = elem.get(styles_word("type")) or elem.get("type") or "paragraph"
        name_elem = elem.find(styles_word("name"))
        sname = (
            (name_elem.get(styles_word("val")) or name_elem.get("val"))
            if name_elem is not None
            else None
        )
        based_on_elem = elem.find(styles_word("basedOn"))
        based_on = (
            (based_on_elem.get(styles_word("val")) or based_on_elem.get("val"))
            if based_on_elem is not None
            else None
        )
        next_elem = elem.find(styles_word("next"))
        next_style = (
            (next_elem.get(styles_word("val")) or next_elem.get("val"))
            if next_elem is not None
            else None
        )
        link_elem = elem.find(styles_word("link"))
        link_style = (
            (link_elem.get(styles_word("val")) or link_elem.get("val"))
            if link_elem is not None
            else None
        )
        is_default = (elem.get(styles_word("default")) or elem.get("default")) in {
            "1",
            "true",
            "on",
        }
        is_custom = (elem.get(styles_word("customStyle")) or elem.get("customStyle")) in {
            "1",
            "true",
            "on",
        }
        existing_styles[sid] = {
            "type": stype,
            "name": sname,
            "based_on": based_on,
            "next": next_style,
            "link": link_style,
            "element": elem,
            "has_tbl_style_pr": elem.find(styles_word("tblStylePr")) is not None,
            "is_default": is_default,
            "is_custom": is_custom,
        }

    # Action specific pre-checks
    if request.action == "create":
        if request.style_id in existing_styles:
            return _style_refusal(request, "style_duplicate_id")
    elif request.action == "update":
        if not styles_part_exists or request.style_id not in existing_styles:
            return _style_refusal(request, "style_missing")
        if existing_styles[request.style_id]["type"] != request.type:
            return _style_refusal(request, "style_type_incompatible")
    else:  # delete
        if not styles_part_exists or request.style_id not in existing_styles:
            return _style_refusal(request, "style_missing")
        target_info = existing_styles[request.style_id]
        if request.type is not None and target_info["type"] != request.type:
            return _style_refusal(request, "style_type_incompatible")
        if request.name is not None and target_info["name"] != request.name:
            return _style_refusal(request, "style_name_mismatch")
        if target_info.get("is_default"):
            return _style_refusal(request, "style_default_cannot_be_deleted")
        if not target_info.get("is_custom"):
            return _style_refusal(request, "style_builtin_cannot_be_deleted")
        refs = _find_style_references(parts, request.style_id, styles_root, styles_ns)
        if refs:
            return _style_refusal(request, "style_in_use")
    # A linked pair names each other through w:link; the request's companion gains the
    # reciprocal link, and an update that re-pairs releases the previous companion.
    backlink_companion: str | None = None
    released_companion: str | None = None
    if request.action != "delete":
        # Validate referenced styles
        effective_based_on = request.based_on
        if request.action == "update" and effective_based_on is None and not request.clear_based_on:
            effective_based_on = existing_styles[request.style_id].get("based_on")

        if request.based_on is not None:
            if request.based_on not in existing_styles:
                return _style_refusal(request, "style_based_on_missing")
            if existing_styles[request.based_on]["type"] != request.type:
                return _style_refusal(request, "style_based_on_type_mismatch")
            curr: str | None = request.based_on
            visited_bases: set[str] = set()
            while curr is not None and curr in existing_styles and curr not in visited_bases:
                visited_bases.add(curr)
                base_info = existing_styles[curr]
                base_elem = base_info.get("element")
                if base_elem is not None and request.type == "table":
                    if base_elem.find(styles_word("pPr")) is not None:
                        return _style_refusal(request, "style_properties_incompatible")
                    if base_elem.find(styles_word("next")) is not None:
                        return _style_refusal(request, "style_next_unsupported")
                curr = base_info.get("based_on")

        effective_next = request.next
        if request.action == "update" and effective_next is None:
            effective_next = existing_styles[request.style_id].get("next")

        if request.next is not None and request.next != request.style_id:
            if request.next not in existing_styles:
                return _style_refusal(request, "style_next_missing")
            if existing_styles[request.next]["type"] != "paragraph":
                return _style_refusal(request, "style_next_type_mismatch")

        if request.link is not None:
            companion = existing_styles.get(request.link)
            if companion is None:
                return _style_refusal(request, "style_link_missing")
            if companion["type"] != _LINKED_STYLE_TYPE[request.type]:
                return _style_refusal(request, "style_link_type_mismatch")
            if companion["link"] not in {None, request.style_id}:
                return _style_refusal(request, "style_link_companion_conflict")
            if companion["link"] is None:
                backlink_companion = request.link
            previous = (
                existing_styles[request.style_id]["link"] if request.action == "update" else None
            )
            if (
                previous is not None
                and previous != request.link
                and previous in existing_styles
                and existing_styles[previous]["link"] == request.style_id
            ):
                released_companion = previous

        # Validate style references within properties
        if request.run_properties and "style" in request.run_properties:
            ref_id = request.run_properties["style"]
            if ref_id not in existing_styles:
                return _style_refusal(request, "style_reference_missing")
            if existing_styles[ref_id]["type"] != "character":
                return _style_refusal(request, "style_type_incompatible")

        if request.paragraph_properties and "style" in request.paragraph_properties:
            ref_id = request.paragraph_properties["style"]
            if ref_id not in existing_styles:
                return _style_refusal(request, "style_reference_missing")
            if existing_styles[ref_id]["type"] != "paragraph":
                return _style_refusal(request, "style_type_incompatible")

        # Check for inheritance cycles
        has_cycle, cycle_code = _detect_style_cycle(
            existing_styles,
            request.style_id,
            effective_based_on,
        )
        if has_cycle and cycle_code:
            return _style_refusal(request, cycle_code)
    # Perform mutations
    if request.action == "create":
        style_elem = ET.Element(
            styles_word("style"),
            {
                styles_word("type"): request.type,
                styles_word("styleId"): request.style_id,
                styles_word("customStyle"): "1",
            },
        )
        name_elem = ET.Element(styles_word("name"), {styles_word("val"): request.name})
        style_elem.append(name_elem)

        if request.based_on is not None:
            style_elem.append(
                ET.Element(styles_word("basedOn"), {styles_word("val"): request.based_on})
            )
        if request.next is not None:
            style_elem.append(ET.Element(styles_word("next"), {styles_word("val"): request.next}))
        if request.link is not None:
            _set_style_link(style_elem, request.link, styles_ns)
        if request.paragraph_properties is not None:
            _apply_style_paragraph_properties(style_elem, request.paragraph_properties, styles_ns)
        if request.run_properties is not None:
            _apply_style_run_properties(style_elem, request.run_properties, styles_ns)
        if request.table_properties is not None:
            _apply_style_table_properties(style_elem, request.table_properties, styles_ns)
        if request.cell_properties is not None:
            _apply_style_cell_properties(style_elem, request.cell_properties, styles_ns)
        if request.row_properties is not None:
            _apply_style_row_properties(style_elem, request.row_properties, styles_ns)
        if request.conditional_formatting is not None:
            _apply_style_conditional_formatting(
                style_elem, request.conditional_formatting, styles_ns
            )

        def _style_child_sort_key(child: ET.Element) -> tuple[int, int]:
            local = _local_name(child.tag)
            rank = _STYLE_CHILD_RANK.get(local, 999)
            if local == "tblStylePr":
                ov_type = child.get(styles_word("type")) or child.get("type") or ""
                try:
                    sub_rank = _TBL_STYLE_OVERRIDE_TYPE_ORDER.index(ov_type)
                except ValueError:
                    sub_rank = 999
                return (rank, sub_rank)
            return (rank, 0)

        style_elem[:] = sorted(style_elem, key=_style_child_sort_key)
        styles_root.append(style_elem)
    elif request.action == "update":
        style_elem = existing_styles[request.style_id]["element"]

        name_elem = style_elem.find(styles_word("name"))
        if name_elem is None:
            name_elem = ET.Element(styles_word("name"))
            style_elem.append(name_elem)
        name_elem.set(styles_word("val"), request.name)

        if request.clear_based_on:
            for based_on_elem in style_elem.findall(styles_word("basedOn")):
                style_elem.remove(based_on_elem)
        elif request.based_on is not None:
            based_on_elem = style_elem.find(styles_word("basedOn"))
            if based_on_elem is None:
                based_on_elem = ET.Element(styles_word("basedOn"))
                style_elem.append(based_on_elem)
            based_on_elem.set(styles_word("val"), request.based_on)

        if request.next is not None:
            next_elem = style_elem.find(styles_word("next"))
            if next_elem is None:
                next_elem = ET.Element(styles_word("next"))
                style_elem.append(next_elem)
            next_elem.set(styles_word("val"), request.next)

        if request.link is not None:
            _set_style_link(style_elem, request.link, styles_ns)

        if request.paragraph_properties is not None:
            _apply_style_paragraph_properties(style_elem, request.paragraph_properties, styles_ns)
        if request.run_properties is not None:
            _apply_style_run_properties(style_elem, request.run_properties, styles_ns)
        if request.table_properties is not None:
            _apply_style_table_properties(style_elem, request.table_properties, styles_ns)
        if request.cell_properties is not None:
            _apply_style_cell_properties(style_elem, request.cell_properties, styles_ns)
        if request.row_properties is not None:
            _apply_style_row_properties(style_elem, request.row_properties, styles_ns)
        if request.conditional_formatting is not None:
            _apply_style_conditional_formatting(
                style_elem, request.conditional_formatting, styles_ns
            )

        def _style_child_sort_key_up(child: ET.Element) -> tuple[int, int]:
            local = _local_name(child.tag)
            rank = _STYLE_CHILD_RANK.get(local, 999)
            if local == "tblStylePr":
                ov_type = child.get(styles_word("type")) or child.get("type") or ""
                try:
                    sub_rank = _TBL_STYLE_OVERRIDE_TYPE_ORDER.index(ov_type)
                except ValueError:
                    sub_rank = 999
                return (rank, sub_rank)
            return (rank, 0)

        style_elem[:] = sorted(style_elem, key=_style_child_sort_key_up)
    else:  # delete
        style_elem = existing_styles[request.style_id]["element"]
        styles_root.remove(style_elem)

    if backlink_companion is not None:
        _set_style_link(existing_styles[backlink_companion]["element"], request.style_id, styles_ns)
    if released_companion is not None:
        released_elem = existing_styles[released_companion]["element"]
        released_elem.remove(released_elem.find(styles_word("link")))

    # Package preparation
    replacements: dict[str, bytes] = {}
    additions: dict[str, bytes] = {}

    if styles_part_exists:
        replacements[_STYLES] = serialize_xml_with_preserved_qnames(styles_root)
    else:
        # Create styles.xml and register relationship and content type
        relationship_type = (
            _STRICT_STYLES_REL_TYPE if main_ns == STRICT_WORD_NS else _STYLES_REL_TYPE
        )
        types_root = parse_xml_with_preserved_qnames(parts["[Content_Types].xml"])
        has_override = any(
            node.get("PartName") == f"/{_STYLES}" for node in types_root.findall(Q(_CT, "Override"))
        )
        if not has_override:
            types_root.append(
                ET.Element(
                    Q(_CT, "Override"),
                    {"PartName": f"/{_STYLES}", "ContentType": _STYLES_CONTENT_TYPE},
                )
            )
            replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)

        rels_name = "word/_rels/document.xml.rels"
        if rels_name in parts:
            rels_root = parse_xml_with_preserved_qnames(parts[rels_name])
            has_rel = any(
                rel.get("Type") == relationship_type
                for rel in rels_root.findall(Q(_PR, "Relationship"))
            )
            if not has_rel:
                existing_ids = [
                    rel.get("Id", "") for rel in rels_root.findall(Q(_PR, "Relationship"))
                ]
                rid_nums = [
                    int(rid[3:])
                    for rid in existing_ids
                    if rid.startswith("rId") and rid[3:].isdigit()
                ]
                next_rid = f"rId{(max(rid_nums) if rid_nums else 0) + 1}"
                rels_root.append(
                    ET.Element(
                        Q(_PR, "Relationship"),
                        {
                            "Id": next_rid,
                            "Type": relationship_type,
                            "Target": "styles.xml",
                        },
                    )
                )
            replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)
        else:
            rels_root = ET.Element(Q(_PR, "Relationships"))
            rels_root.append(
                ET.Element(
                    Q(_PR, "Relationship"),
                    {
                        "Id": "rIdStyles",
                        "Type": relationship_type,
                        "Target": "styles.xml",
                    },
                )
            )
            additions[rels_name] = serialize_xml_with_preserved_qnames(rels_root)

        additions[_STYLES] = serialize_xml_with_preserved_qnames(styles_root)

    all_changed = frozenset(replacements.keys() | additions.keys())

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, all_changed):
            return "style_staged_verification_failed"

        try:
            staged_graph = PackageGraph.from_file(staged)
        except (OSError, PackageError, zipfile.BadZipFile):
            return "style_staged_verification_failed"

        if _STYLES not in staged_graph.parts:
            return "style_staged_verification_failed"

        try:
            staged_root = ET.fromstring(staged_graph.read_part_bytes(_STYLES))
        except ET.ParseError:
            return "style_staged_verification_failed"

        s_ns = _word_namespace(staged_root) or styles_ns
        def s_word(local):
            return Q(s_ns, local)

        matching = [
            s
            for s in staged_root.findall(s_word("style"))
            if s.get(s_word("styleId")) == request.style_id or s.get("styleId") == request.style_id
        ]
        if request.action == "delete":
            if matching:
                return "style_staged_verification_failed"
            candidate = inspect_docx(staged)
            if not candidate.valid:
                return "style_staged_verification_failed"
            return None

        if not matching:
            return "style_staged_verification_failed"

        m_style = matching[0]
        actual_type = m_style.get(s_word("type")) or m_style.get("type")
        if actual_type != request.type:
            return "style_staged_verification_failed"

        name_el = m_style.find(s_word("name"))
        if name_el is None:
            return "style_staged_verification_failed"
        actual_name = name_el.get(s_word("val")) or name_el.get("val")
        if actual_name != request.name:
            return "style_staged_verification_failed"

        if request.clear_based_on and m_style.find(s_word("basedOn")) is not None:
            return "style_staged_verification_failed"
        if request.based_on is not None:
            based_el = m_style.find(s_word("basedOn"))
            if based_el is None:
                return "style_staged_verification_failed"
            actual_based = based_el.get(s_word("val")) or based_el.get("val")
            if actual_based != request.based_on:
                return "style_staged_verification_failed"

        if request.next is not None:
            next_el = m_style.find(s_word("next"))
            if next_el is None:
                return "style_staged_verification_failed"
            actual_next = next_el.get(s_word("val")) or next_el.get("val")
            if actual_next != request.next:
                return "style_staged_verification_failed"

        def staged_link(style_id: str) -> str | None:
            for staged_style in staged_root.findall(s_word("style")):
                sid = staged_style.get(s_word("styleId")) or staged_style.get("styleId")
                if sid == style_id:
                    link_el = staged_style.find(s_word("link"))
                    if link_el is None:
                        return None
                    return link_el.get(s_word("val")) or link_el.get("val")
            return None

        if request.link is not None and (
            staged_link(request.style_id) != request.link
            or staged_link(request.link) != request.style_id
        ):
            return "style_staged_verification_failed"
        if released_companion is not None and staged_link(released_companion) is not None:
            return "style_staged_verification_failed"
        if request.table_properties is not None and m_style.find(s_word("tblPr")) is None:
            return "style_staged_verification_failed"
        if request.cell_properties is not None and m_style.find(s_word("tcPr")) is None:
            return "style_staged_verification_failed"
        if request.run_properties is not None and m_style.find(s_word("rPr")) is None:
            return "style_staged_verification_failed"
        if request.paragraph_properties is not None and m_style.find(s_word("pPr")) is None:
            return "style_staged_verification_failed"
        if request.row_properties is not None and m_style.find(s_word("trPr")) is None:
            return "style_staged_verification_failed"

        ranks = [_STYLE_CHILD_RANK.get(_local_name(c.tag), 999) for c in m_style]
        if ranks != sorted(ranks):
            return "style_staged_verification_failed"

        tbl_style_prs = m_style.findall(s_word("tblStylePr"))
        for tsp in tbl_style_prs:
            t_type = tsp.get(s_word("type")) or tsp.get("type")
            if not t_type or t_type not in _VALID_TBL_STYLE_OVERRIDE_TYPES:
                return "style_staged_verification_failed"
            tsp_ranks = [_TBL_STYLE_PR_CHILD_RANK.get(_local_name(c.tag), 999) for c in tsp]
            if tsp_ranks != sorted(tsp_ranks):
                return "style_staged_verification_failed"
            p_pr = tsp.find(s_word("pPr"))
            if p_pr is not None:
                p_ranks = [_PPR_RANK.get(_local_name(c.tag), 999) for c in p_pr]
                if p_ranks != sorted(p_ranks):
                    return "style_staged_verification_failed"
            r_pr = tsp.find(s_word("rPr"))
            if r_pr is not None:
                r_ranks = [_RPR_RANK.get(_local_name(c.tag), 999) for c in r_pr]
                if r_ranks != sorted(r_ranks):
                    return "style_staged_verification_failed"
            tr_pr = tsp.find(s_word("trPr"))
            if tr_pr is not None:
                tr_ranks = [
                    _TRPR_CHILD_ORDER.index(_local_name(c.tag))
                    if _local_name(c.tag) in _TRPR_CHILD_ORDER
                    else 999
                    for c in tr_pr
                ]
                if tr_ranks != sorted(tr_ranks):
                    return "style_staged_verification_failed"

        if request.conditional_formatting is not None:
            for expected_ov in request.conditional_formatting:
                ov_type = expected_ov["type"]
                matching_tsp = next(
                    (
                        t
                        for t in tbl_style_prs
                        if (t.get(s_word("type")) or t.get("type")) == ov_type
                    ),
                    None,
                )
                if matching_tsp is None:
                    return "style_staged_verification_failed"
                if (
                    expected_ov.get("paragraph_properties") is not None
                    and matching_tsp.find(s_word("pPr")) is None
                ):
                    return "style_staged_verification_failed"
                if (
                    expected_ov.get("run_properties") is not None
                    and matching_tsp.find(s_word("rPr")) is None
                ):
                    return "style_staged_verification_failed"
                if (
                    expected_ov.get("table_properties") is not None
                    and matching_tsp.find(s_word("tblPr")) is None
                ):
                    return "style_staged_verification_failed"
                if (
                    expected_ov.get("row_properties") is not None
                    and matching_tsp.find(s_word("trPr")) is None
                ):
                    return "style_staged_verification_failed"
                if (
                    expected_ov.get("cell_properties") is not None
                    and matching_tsp.find(s_word("tcPr")) is None
                ):
                    return "style_staged_verification_failed"

        # Check inspection valid
        candidate = inspect_docx(staged)
        if not candidate.valid:
            return "style_staged_verification_failed"

        return None

    target_type = request.type or (
        existing_styles[request.style_id]["type"] if request.style_id in existing_styles else None
    )
    target_name = request.name or (
        existing_styles[request.style_id]["name"] if request.style_id in existing_styles else None
    )

    changed_target = {
        "style_id": request.style_id,
        "action": request.action,
        "type": target_type,
        "name": target_name,
    }
    if request.based_on is not None or request.clear_based_on:
        changed_target["based_on"] = request.based_on
    if request.next is not None:
        changed_target["next"] = request.next
    if request.link is not None:
        changed_target["link"] = request.link
    changed_targets: list[dict[str, Any]] = [changed_target]
    for companion_id, companion_link in (
        (backlink_companion, request.style_id),
        (released_companion, None),
    ):
        if companion_id is not None:
            changed_targets.append(
                {
                    "style_id": companion_id,
                    "action": "update",
                    "type": existing_styles[companion_id]["type"],
                    "name": existing_styles[companion_id]["name"],
                    "link": companion_link,
                }
            )

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=changed_targets,
        additions=additions,
        verify=verify_publication,
        failure_code="style_staged_verification_failed",
        exists_code="style_destination_invalid",
        unavailable_code="style_publication_unavailable",
    )

    style_dict: dict[str, Any] = {
        "style_id": request.style_id,
        "action": request.action,
        "type": target_type,
        "name": target_name,
    }
    if request.based_on is not None or request.clear_based_on:
        style_dict["based_on"] = request.based_on
    if request.next is not None:
        style_dict["next"] = request.next
    if request.link is not None:
        style_dict["link"] = request.link
    if request.run_properties is not None:
        style_dict["run_properties"] = dict(request.run_properties)
    if request.paragraph_properties is not None:
        style_dict["paragraph_properties"] = dict(request.paragraph_properties)

    if request.table_properties is not None:
        style_dict["table_properties"] = dict(request.table_properties)
    if request.cell_properties is not None:
        style_dict["cell_properties"] = dict(request.cell_properties)
    if request.row_properties is not None:
        style_dict["row_properties"] = dict(request.row_properties)
    if request.conditional_formatting is not None:
        style_dict["conditional_formatting"] = [dict(ov) for ov in request.conditional_formatting]
    return StyleReport(
        status=published.status,
        code=published.code,
        source_sha256=published.source_sha256,
        destination_sha256=published.output_sha256,
        style=style_dict if published.passed else None,
        changed_targets=changed_targets if published.passed else [],
    )
