"""Programmatic construction and cursor-based editing of DOCX packages.

New packages are deterministic; opened packages are published through a Document
session so every untouched member retains its original bytes and metadata.
"""

from __future__ import annotations

import copy
import enum
import hashlib
import os
import posixpath
import re
import tempfile
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from ._optional import feature_module
from .formatting import _HIGHLIGHT_CANONICAL_MAP, _PPR_RANK, _RPR_RANK
from .inspect import (
    CONTENT_TYPES,
    PKG_REL,
    Q,
    W,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .model import Node, NodeEditError
from .output import SaveOutputParameters
from .page_setup import (
    _DEFAULT_COLUMN_SPACING,
    _MAX_COLUMN_COUNT,
    _MAX_PAGE_DIMENSION,
    _MIN_PAGE_DIMENSION,
)
from .properties import (
    CORE_PROPERTIES_CONTENT_TYPE,
    CORE_PROPERTIES_PARTNAME,
    CORE_PROPERTIES_REL_TYPE,
    CP_NS,
    DC_NS,
    DCTERMS_NS,
    XSI_NS,
    _format_datetime,
)
from .protection import _SETTINGS, _SETTINGS_CHILD_ORDER_MAP, _SETTINGS_CONTENT_TYPE
from .section_break import _canonical_break_type, _insert_sect_pr_child
from .table_properties import (
    _LOOK_FLAG_TO_XML_ATTR,
    TableLookSpec,
    _decode_look_bitmask,
    _encode_look_bitmask,
    _insert_or_replace_tblpr_child,
)

__all__ = [
    "DocumentBuilder",
    "DocumentHeaderFooter",
    "DocumentHyperlink",
    "DocumentImage",
    "DocumentList",
    "DocumentParagraph",
    "DocumentRun",
    "DocumentSaveReport",
    "DocumentSection",
    "DocumentShape",
    "DocumentTable",
    "DocumentTableCell",
    "DocumentTableColumn",
    "DocumentTableRow",
]


_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_IMAGE_RELATIONSHIP = f"{_REL}/image"
_HYPERLINK_RELATIONSHIP = f"{_REL}/hyperlink"
# Namespaces declared on serialized part roots. Pictures, shapes, hyperlink
# validation, document settings, image probing and fields are feature modules
# resolved through feature_module at the call site, so the paragraph, run, table
# and section builder stays within the standard library and the shared core.
R = _REL
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
_DOCUMENT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
_STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_NUMBERING_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
)
_HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
_FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
# ST_HdrFtr values a section references a header or footer part under.
_HEADER_FOOTER_KINDS = frozenset({"default", "first", "even"})
_PAGE_ORIENTATIONS = frozenset({"portrait", "landscape"})
# Section page-setup keyword -> w:pgMar attribute.
_PAGE_MARGIN_ATTRIBUTES = (
    ("margin_top", "top"),
    ("margin_right", "right"),
    ("margin_bottom", "bottom"),
    ("margin_left", "left"),
    ("header_distance", "header"),
    ("footer_distance", "footer"),
    ("gutter", "gutter"),
)
_LIST_KINDS = frozenset({"numbered", "bulleted"})
# ilvl 0-8: the nine levels of one w:abstractNum.
_LIST_LEVELS = 9
# Word's default outline-list cycle: 1. a. i. for numbers; for bullets the Symbol
# bullet, a Courier New circle and the Wingdings square, as Word itself writes them.
_NUMBERED_LEVEL_FORMATS = ("decimal", "lowerLetter", "lowerRoman")
_BULLET_LEVEL_GLYPHS = (
    ("\uf0b7", "Symbol"),
    ("o", "Courier New"),
    ("\uf0a7", "Wingdings"),
)
_LIST_INDENT_STEP = 720
_LIST_HANGING_INDENT = 360
# CT_TcPr children this module authors, in ECMA-376 17.4.70 sequence order.
_TC_PR_ORDER = ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_STYLE_ID = re.compile(r"[A-Za-z][A-Za-z0-9_-]{0,63}\Z")
# A Word style name of single-space-separated words, such as ``Subtle Emphasis``.
_STYLE_NAME = re.compile(r"[A-Za-z][A-Za-z0-9_-]*(?: [A-Za-z0-9_-]+)+\Z")
# Word names these built-in styles in lower case; python-docx accepts the title case.
_LOWER_CASE_BUILT_IN_STYLE_NAMES = frozenset(f"heading {level}" for level in range(1, 10))
_MAX_TEXT_CHARACTERS = 32_768
_MAX_EMU = 201_168_000
# Matches the tooltip bound of the source-bound hyperlink request parser.
_MAX_TOOLTIP_CHARACTERS = 1_024
_HEX_COLOR = re.compile(r"[0-9A-Fa-f]{6}\Z")
# Word refuses run sizes above 1638 half-points (819 pt).
_MAX_FONT_HALF_POINTS = 1_638
# Matches the font face bound of the source-bound run formatting parser.
_MAX_FONT_NAME_CHARACTERS = 256
# python-docx ``WD_COLOR_INDEX`` member names -> ``ST_HighlightColor``.  ``GREEN`` is
# Word's dark green there, while the ``"green"`` literal stays Word's bright green.
_WD_COLOR_INDEX_HIGHLIGHTS = {
    "BLACK": "black",
    "BLUE": "blue",
    "TURQUOISE": "cyan",
    "BRIGHT_GREEN": "green",
    "PINK": "magenta",
    "RED": "red",
    "YELLOW": "yellow",
    "WHITE": "white",
    "DARK_BLUE": "darkBlue",
    "TEAL": "darkCyan",
    "GREEN": "darkGreen",
    "VIOLET": "darkMagenta",
    "DARK_RED": "darkRed",
    "DARK_YELLOW": "darkYellow",
    "GRAY_50": "darkGray",
    "GRAY_25": "lightGray",
}
_HYPERLINK_STYLE = "Hyperlink"
# MS-DOCX compatSetting namespace for Word's own compatibility settings.
_WORD_COMPAT_SETTING_URI = "http://schemas.microsoft.com/office/word"
# compatibilityMode values for Word 2007, Word 2010 and Word 2013 or later.
_COMPAT_MODES = frozenset({12, 14, 15})
# Conventional prefixes declared on each serialized part root that uses them.
_PART_PREFIXES = (
    ("w", W),
    ("r", R),
    ("wp", WP),
    ("a", A),
    ("pic", PIC),
    ("wps", _WPS),
    ("asvg", ASVG),
)
# Prefixes for the package-level core-properties part; never registered globally.
_CORE_PART_PREFIXES = (
    ("cp", CP_NS),
    ("dc", DC_NS),
    ("dcterms", DCTERMS_NS),
    ("xsi", XSI_NS),
)
# Builder keyword -> core-properties element, in the order Word writes them.
# CT_CoreProperties is an xsd:all group, so element order carries no meaning.
_CORE_PROPERTY_ELEMENTS = (
    ("title", DC_NS, "title"),
    ("subject", DC_NS, "subject"),
    ("author", DC_NS, "creator"),
    ("keywords", CP_NS, "keywords"),
    ("description", DC_NS, "description"),
    ("last_modified_by", CP_NS, "lastModifiedBy"),
    ("revision", CP_NS, "revision"),
    ("last_printed", CP_NS, "lastPrinted"),
    ("created", DCTERMS_NS, "created"),
    ("modified", DCTERMS_NS, "modified"),
    ("category", CP_NS, "category"),
    ("content_status", CP_NS, "contentStatus"),
    ("content_type", CP_NS, "contentType"),
    ("identifier", DC_NS, "identifier"),
    ("language", DC_NS, "language"),
    ("version", CP_NS, "version"),
)
# The OPC schema types dcterms:created and dcterms:modified as W3CDTF dates.
_W3CDTF_CORE_PROPERTIES = frozenset({"created", "modified"})

for _prefix, _namespace in _PART_PREFIXES:
    ET.register_namespace(_prefix, _namespace)


@dataclass
class DocumentParagraph:
    """A paragraph node owned by one :class:`DocumentBuilder`."""

    element: ET.Element
    _builder: DocumentBuilder

    def append_hyperlink(
        self,
        url: str | None = None,
        text: str | None = None,
        *,
        anchor: str | None = None,
        tooltip: str | None = None,
    ) -> DocumentHyperlink:
        """Append an external URL or internal bookmark hyperlink run to this paragraph."""
        return self._builder.append_hyperlink(
            url, text, anchor=anchor, tooltip=tooltip, paragraph=self
        )

    def add_run(self, text: str = "", style: str | None = None, **formatting: Any) -> DocumentRun:
        """Append a text run to this paragraph, like python-docx ``Paragraph.add_run``.

        ``formatting`` takes the keyword arguments of :meth:`DocumentBuilder.append_run`.
        """
        return self._builder.append_run(text, paragraph=self, style=style, **formatting)

    @property
    def runs(self) -> tuple[DocumentRun, ...]:
        """The paragraph's direct ``w:r`` runs in document order, like python-docx."""
        return tuple(DocumentRun(run, self) for run in self.element.findall(Q(W, "r")))

    @property
    def text(self) -> str:
        """The paragraph's run text, hyperlink runs included, like python-docx."""
        return "".join(
            node.text or ""
            for child in self.element
            if child.tag in {Q(W, "r"), Q(W, "hyperlink"), Q(W, "fldSimple")}
            for node in child.findall(
                Q(W, "t") if child.tag == Q(W, "r") else f"{Q(W, 'r')}/{Q(W, 't')}"
            )
        )

    @property
    def outline_level(self) -> int | None:
        """Direct outline level: 0–8 are headings and 9 is body text."""
        node = self.element.find(f"{Q(W, 'pPr')}/{Q(W, 'outlineLvl')}")
        return None if node is None else int(node.get(Q(W, "val"), "9"))

    @outline_level.setter
    def outline_level(self, level: int | None) -> None:
        self.set_outline_level(level)

    def set_outline_level(self, level: int | None) -> DocumentParagraph:
        """Set or remove a direct outline level without changing the style."""
        if level is not None and (
            isinstance(level, bool) or not isinstance(level, int) or not 0 <= level <= 9
        ):
            raise NodeEditError("paragraph_outline_level_invalid")
        self._builder._resolve_paragraph(self)
        self._set_property("outlineLvl", None if level is None else {Q(W, "val"): str(level)})
        return self

    def set_drop_cap(
        self, kind: str = "drop", *, lines: int = 3, distance: int = 0
    ) -> DocumentParagraph:
        """Author a drop cap frame; distance is horizontal spacing in twips.

        ``kind`` is ``drop``, ``margin`` or ``none``. The paragraph contains the
        initial letter(s); the following paragraph supplies the surrounding text.
        """
        if not isinstance(kind, str) or kind not in {"drop", "margin", "none"}:
            raise NodeEditError("paragraph_drop_cap_kind_invalid")
        if isinstance(lines, bool) or not isinstance(lines, int) or not 1 <= lines <= 10:
            raise NodeEditError("paragraph_drop_cap_lines_invalid")
        if (
            isinstance(distance, bool)
            or not isinstance(distance, int)
            or not 0 <= distance <= 31680
        ):
            raise NodeEditError("paragraph_drop_cap_distance_invalid")
        self._builder._resolve_paragraph(self)
        existing = self.element.find(f"{Q(W, 'pPr')}/{Q(W, 'framePr')}")
        attributes = dict(existing.attrib) if existing is not None else {}
        authored = {
            Q(W, "dropCap"): kind,
            Q(W, "lines"): str(lines),
            Q(W, "hSpace"): str(distance),
            Q(W, "wrap"): "around",
            Q(W, "vAnchor"): "text",
            Q(W, "hAnchor"): "text",
        }
        if kind == "none":
            other = {key: value for key, value in attributes.items() if key not in authored}
            if other:
                attributes[Q(W, "dropCap")] = "none"
                attributes.pop(Q(W, "lines"), None)
            else:
                attributes = None
        else:
            attributes.update(authored)
        self._set_property("framePr", attributes)
        return self

    def _set_property(self, name: str, attributes: dict[str, str] | None) -> None:
        properties = self.element.find(Q(W, "pPr"))
        if properties is None:
            if attributes is None:
                return
            properties = ET.Element(Q(W, "pPr"))
            self.element.insert(0, properties)
        for child in list(properties.findall(Q(W, name))):
            properties.remove(child)
        if attributes is not None:
            properties.append(ET.Element(Q(W, name), attributes))
        properties[:] = sorted(
            properties, key=lambda child: _PPR_RANK.get(child.tag.rsplit("}", 1)[-1], 999)
        )


@dataclass(frozen=True)
class DocumentRun:
    """A run node owned by one :class:`DocumentBuilder`."""

    element: ET.Element
    paragraph: DocumentParagraph

    @property
    def text(self) -> str:
        """The run's ``w:t`` text."""
        return "".join(node.text or "" for node in self.element.findall(Q(W, "t")))


@dataclass(frozen=True)
class DocumentTable:
    """A table node owned by one :class:`DocumentBuilder`."""

    element: ET.Element
    _builder: DocumentBuilder

    @property
    def rows(self) -> tuple[DocumentTableRow, ...]:
        """The table's ``w:tr`` rows in document order, like python-docx ``Table.rows``."""
        return tuple(DocumentTableRow(row, self) for row in self.element.findall(Q(W, "tr")))

    @property
    def columns(self) -> tuple[DocumentTableColumn, ...]:
        """The table's grid columns in order, like python-docx ``Table.columns``."""
        grid_columns = self.element.findall(f"{Q(W, 'tblGrid')}/{Q(W, 'gridCol')}")
        return tuple(DocumentTableColumn(index, self) for index in range(len(grid_columns)))

    def cell(self, row_idx: int, col_idx: int) -> DocumentTableCell:
        """The cell at zero-based ``row_idx`` and grid column ``col_idx``, like python-docx.

        Positions inside a merged region return the region's top-left cell.
        """
        rows = self.element.findall(Q(W, "tr"))
        if not _is_index(row_idx, len(rows)):
            raise IndexError("table_row_index_invalid")
        layout = _table_layout(self.element)
        if not _is_index(col_idx, len(layout[row_idx])):
            raise IndexError("table_column_index_invalid")
        return DocumentTableCell(layout[row_idx][col_idx].origin, self)

    def set_style_context(self, look: TableLookSpec | None) -> None:
        """Patch conditional style flags; ``None`` removes the explicit table look.

        Omitted flags retain their current values. ``no_hband`` and ``no_vband``
        disable horizontal and vertical banding respectively.
        """
        self._builder._target_element(self)
        if look is not None and not isinstance(look, TableLookSpec):
            raise NodeEditError("builder_table_look_invalid")
        if look is not None and not look.to_dict():
            return
        properties = self.element.find(Q(W, "tblPr"))
        if properties is None:
            if look is None:
                return
            properties = ET.Element(Q(W, "tblPr"))
            self.element.insert(0, properties)
        existing = properties.find(Q(W, "tblLook"))
        if look is None:
            if existing is not None:
                properties.remove(existing)
            return
        desired = look.to_dict()
        if not desired:
            return
        raw_mask = existing.get(Q(W, "val"), "0000") if existing is not None else "0000"
        if re.fullmatch(r"[0-9A-Fa-f]{4}", raw_mask) is None:
            raise NodeEditError("builder_table_look_invalid")
        base = int(raw_mask, 16)
        flags = _decode_look_bitmask(base)
        if existing is not None:
            for flag, attribute in _LOOK_FLAG_TO_XML_ATTR.items():
                value = existing.get(Q(W, attribute))
                if value is not None:
                    if value not in {"0", "1", "false", "true", "off", "on"}:
                        raise NodeEditError("builder_table_look_invalid")
                    flags[flag] = value in {"1", "true", "on"}
        if look.val is not None:
            flags.update(_decode_look_bitmask(int(look.val, 16)))
        flags.update((key, value) for key, value in desired.items() if key != "val")
        attributes = dict(existing.attrib) if existing is not None else {}
        attributes[Q(W, "val")] = f"{_encode_look_bitmask(flags, base):04X}"
        attributes.update(
            (Q(W, attribute), "1" if flags[flag] else "0")
            for flag, attribute in _LOOK_FLAG_TO_XML_ATTR.items()
        )
        _insert_or_replace_tblpr_child(properties, ET.Element(Q(W, "tblLook"), attributes), W)


@dataclass(frozen=True)
class DocumentTableRow:
    """A ``w:tr`` row of one builder table."""

    element: ET.Element
    table: DocumentTable

    @property
    def cells(self) -> tuple[DocumentTableCell, ...]:
        """The row's ``w:tc`` cells in order."""
        return tuple(
            DocumentTableCell(cell, self.table) for cell in self.element.findall(Q(W, "tc"))
        )


@dataclass(frozen=True)
class DocumentTableColumn:
    """One zero-based grid column of a builder table."""

    index: int
    table: DocumentTable

    @property
    def cells(self) -> tuple[DocumentTableCell, ...]:
        """The column's cells, top to bottom; merged regions repeat their top-left cell."""
        return tuple(
            self.table.cell(row_idx, self.index) for row_idx in range(len(self.table.rows))
        )


# Not frozen: ``text`` is assignable, like python-docx ``_Cell.text``.
@dataclass
class DocumentTableCell:
    """A ``w:tc`` cell of one builder table."""

    element: ET.Element
    table: DocumentTable

    @property
    def paragraphs(self) -> tuple[DocumentParagraph, ...]:
        """The cell's ``w:p`` paragraphs, usable wherever a builder paragraph is."""
        builder = self.table._builder
        return tuple(DocumentParagraph(p, builder) for p in self.element.findall(Q(W, "p")))

    @property
    def text(self) -> str:
        """The cell's paragraph texts joined by line feeds, like python-docx.

        Assigning replaces every paragraph with one holding the new text and keeps
        the cell properties.
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, value: str) -> None:
        value = _require_text(value, "table_cell_text")
        builder = self.table._builder
        for child in list(self.element):
            if child.tag == Q(W, "tcPr"):
                continue
            self.element.remove(child)
            current = builder._current_paragraph
            if current is not None and current.element is child:
                builder._current_paragraph = None
        _append_text_run(ET.SubElement(self.element, Q(W, "p")), value)

    def merge(self, other: DocumentTableCell) -> DocumentTableCell:
        """Merge the rectangle spanned by this cell and ``other``, like python-docx.

        Grid columns join through ``w:gridSpan`` and rows through ``w:vMerge``.  The
        content of every non-empty merged cell moves, in reading order, into the
        top-left cell, which is returned.  Handles to absorbed cells go stale.  A
        rectangle that would cut through an existing merged region is refused.
        """
        if (
            not isinstance(other, DocumentTableCell)
            or other.table.element is not self.table.element
        ):
            raise ValueError("table_merge_cell_invalid")
        layout = _table_layout(self.table.element)
        first = _slot_of(layout, self.element)
        last = _slot_of(layout, other.element)
        if first is None or last is None:
            raise ValueError("table_merge_cell_invalid")
        top = min(first.top, last.top)
        bottom = max(first.bottom, last.bottom)
        left = min(first.left, last.left)
        right = max(first.right, last.right)
        for row_slots in layout[top : bottom + 1]:
            if len(row_slots) <= right or any(
                slot.top < top or slot.bottom > bottom or slot.left < left or slot.right > right
                for slot in row_slots[left : right + 1]
            ):
                raise ValueError("table_merge_not_rectangular")
        grid_widths = [
            int(column.get(Q(W, "w"), "0"))
            for column in self.table.element.findall(f"{Q(W, 'tblGrid')}/{Q(W, 'gridCol')}")
        ]
        width = str(sum(grid_widths[left : right + 1]))
        rows = self.table.element.findall(Q(W, "tr"))
        origin = layout[top][left].cell
        merged_cells = [
            list(dict.fromkeys(slot.cell for slot in layout[row_idx][left : right + 1]))
            for row_idx in range(top, bottom + 1)
        ]
        content = [
            block
            for row_cells in merged_cells
            for cell in row_cells
            if not _cell_is_empty(cell)
            for block in cell
            if block.tag != Q(W, "tcPr")
        ]
        discarded: list[ET.Element] = []
        for row_idx, row_cells in zip(range(top, bottom + 1), merged_cells, strict=True):
            kept, *absorbed = row_cells
            for cell in row_cells:
                for block in [child for child in cell if child.tag != Q(W, "tcPr")]:
                    cell.remove(block)
                    discarded.append(block)
            for cell in absorbed:
                rows[row_idx].remove(cell)
            properties = kept.find(Q(W, "tcPr"))
            if properties is None:
                properties = ET.Element(Q(W, "tcPr"))
                kept.insert(0, properties)
            _set_cell_property(properties, "tcW", {Q(W, "w"): width, Q(W, "type"): "dxa"})
            span = right - left + 1
            _set_cell_property(
                properties, "gridSpan", {Q(W, "val"): str(span)} if span > 1 else None
            )
            _set_cell_property(properties, "hMerge", None)
            if top == bottom:
                _set_cell_property(properties, "vMerge", None)
            else:
                _set_cell_property(
                    properties,
                    "vMerge",
                    {Q(W, "val"): "restart"} if row_idx == top else {},
                )
            if kept is not origin:
                ET.SubElement(kept, Q(W, "p"))
        origin.extend(content)
        if not content:
            ET.SubElement(origin, Q(W, "p"))
        builder = self.table._builder
        current = builder._current_paragraph
        if current is not None and any(
            current.element is block for block in discarded if block not in content
        ):
            builder._current_paragraph = None
        return DocumentTableCell(origin, self.table)


@dataclass(frozen=True)
class DocumentImage:
    """An inline or anchored picture run owned by one :class:`DocumentBuilder`.

    ``relationship_id``/``part_name`` name the picture's own media part (the SVG for
    an SVG picture); ``fallback_relationship_id``/``fallback_part_name`` name the PNG
    fallback an SVG blip embeds. ``image_format`` is the header-read identity, size,
    and DPI.
    """

    element: ET.Element
    paragraph: DocumentParagraph
    relationship_id: str
    part_name: str
    image_format: Any = None  # wolfdocx.image_formats.ImageFormat
    fallback_relationship_id: str | None = None
    fallback_part_name: str | None = None


@dataclass(frozen=True)
class DocumentShape:
    """An inline DrawingML preset AutoShape run owned by one :class:`DocumentBuilder`."""

    element: ET.Element
    paragraph: DocumentParagraph
    preset: str


@dataclass(frozen=True)
class DocumentHyperlink:
    """A ``w:hyperlink`` owned by one :class:`DocumentBuilder`.

    An external link carries ``relationship_id`` and ``url``; an internal link
    carries only the bookmark ``anchor``.
    """

    element: ET.Element
    paragraph: DocumentParagraph
    relationship_id: str | None
    url: str | None
    anchor: str | None = None


@dataclass(frozen=True)
class DocumentSection:
    """A ``w:sectPr`` section of one :class:`DocumentBuilder` document.

    The handle stays valid when a later :meth:`DocumentBuilder.add_section` moves its
    properties from the body into the paragraph that ends the section.
    """

    element: ET.Element
    _builder: DocumentBuilder

    def set_page_setup(
        self,
        *,
        orientation: str | None = None,
        page_width: int | None = None,
        page_height: int | None = None,
        margin_top: int | None = None,
        margin_right: int | None = None,
        margin_bottom: int | None = None,
        margin_left: int | None = None,
        header_distance: int | None = None,
        footer_distance: int | None = None,
        gutter: int | None = None,
        columns: int | None = None,
        column_spacing: int | None = None,
    ) -> None:
        """Set page size, orientation, margins and equal-width columns, all in twips.

        ``None`` leaves a setting unchanged.  ``orientation`` is ``portrait`` or
        ``landscape``; changing it without an explicit size swaps the current width
        and height, and an explicit size that contradicts it is refused.  Every value
        is validated before any is written, so a refused call changes nothing.
        """
        _apply_page_setup(
            self.element,
            orientation=orientation,
            page_width=page_width,
            page_height=page_height,
            margins={
                "margin_top": margin_top,
                "margin_right": margin_right,
                "margin_bottom": margin_bottom,
                "margin_left": margin_left,
                "header_distance": header_distance,
                "footer_distance": footer_distance,
                "gutter": gutter,
            },
            columns=columns,
            column_spacing=column_spacing,
        )

    def header(self, kind: str = "default") -> DocumentHeaderFooter:
        """This section's ``default``, ``first`` or ``even`` page header, created on demand.

        A ``first`` header turns on the section's distinct first page (``w:titlePg``);
        an ``even`` header turns on Word's document-wide odd and even pages setting.
        """
        return self._builder._header_footer(self, "header", kind)

    def footer(self, kind: str = "default") -> DocumentHeaderFooter:
        """This section's ``default``, ``first`` or ``even`` page footer, created on demand."""
        return self._builder._header_footer(self, "footer", kind)


@dataclass(frozen=True)
class DocumentHeaderFooter:
    """A ``w:hdr`` or ``w:ftr`` part referenced by one builder section.

    Paragraphs here accept runs, character styles, shapes, images and hyperlinks.
    Images and external hyperlinks use this part's own relationship file.
    """

    element: ET.Element
    kind: str
    relationship_id: str
    part_name: str
    _builder: DocumentBuilder

    @property
    def paragraphs(self) -> tuple[DocumentParagraph, ...]:
        """The part's ``w:p`` paragraphs in order."""
        return tuple(DocumentParagraph(p, self._builder) for p in self.element.findall(Q(W, "p")))

    def write_paragraph(self, text: str = "", *, style: str | None = None) -> DocumentParagraph:
        """Append a paragraph to this header or footer without moving the body cursor."""
        text = _require_text(text, "paragraph_text")
        element = self._builder._new_paragraph_element(_require_style(style, "paragraph_style"))
        self.element.append(element)
        _append_text_run(element, text)
        return DocumentParagraph(element, self._builder)


@dataclass(frozen=True)
class DocumentList:
    """A ``numbered`` or ``bulleted`` nine-level list with its own ``w:num`` instance.

    Numbered levels cycle ``1.``, ``a.``, ``i.``; bulleted levels cycle a bullet, a
    circle and a square.  Each level indents a further half inch.
    """

    kind: str
    number_id: int
    _builder: DocumentBuilder

    def add_item(
        self, text: str = "", *, level: int = 0, style: str | None = None
    ) -> DocumentParagraph:
        """Append a body paragraph numbered at zero-based ``level`` (0 to 8) of this list."""
        return self._builder._add_list_item(self, text, level, style)


@dataclass(frozen=True)
class DocumentSaveReport(SaveOutputParameters):
    """Receipt returned after a builder package is atomically published."""

    destination: Path
    sha256: str
    paragraph_count: int
    table_count: int
    image_count: int

    def to_dict(self) -> dict[str, str | int]:
        return {
            "destination": str(self.destination),
            "sha256": self.sha256,
            "paragraph_count": self.paragraph_count,
            "table_count": self.table_count,
            "image_count": self.image_count,
        }


@dataclass(frozen=True)
class _ImagePart:
    relationship_id: str
    part_name: str
    content_type: str
    data: bytes
    source_part: str


def _valid_xml_text(text: str) -> bool:
    for char in text:
        codepoint = ord(char)
        if codepoint in (0x9, 0xA, 0xD) or 0x20 <= codepoint <= 0xD7FF:
            continue
        if 0xE000 <= codepoint <= 0xFFFD or 0x10000 <= codepoint <= 0x10FFFF:
            continue
        return False
    return True


def _require_text(value: object, label: str) -> str:
    if not isinstance(value, str) or len(value) > _MAX_TEXT_CHARACTERS:
        raise ValueError(f"{label}_invalid")
    if not _valid_xml_text(value):
        raise ValueError(f"{label}_invalid")
    return value


def _require_style(value: str | None, label: str) -> tuple[str, str] | None:
    """Return ``(styleId, name)`` for a style ID or a spaced Word style name.

    A spaced name maps to the styleId Word derives for it: each word capitalised,
    spaces removed, so ``Subtle Emphasis`` becomes ``SubtleEmphasis``.
    """
    if value is None:
        return None
    if isinstance(value, str) and _STYLE_ID.fullmatch(value) is not None:
        return value, value
    if not isinstance(value, str) or _STYLE_NAME.fullmatch(value) is None:
        raise ValueError(f"{label}_invalid")
    style_id = "".join(word[:1].upper() + word[1:] for word in value.split(" "))
    if _STYLE_ID.fullmatch(style_id) is None:
        raise ValueError(f"{label}_invalid")
    lowered = value.lower()
    return style_id, lowered if lowered in _LOWER_CASE_BUILT_IN_STYLE_NAMES else value


def _require_emu(value: int | None, label: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= _MAX_EMU:
        raise ValueError(f"{label}_invalid")
    return value


def _image_bytes(image: bytes | bytearray | memoryview | str | Path, code: str) -> bytes:
    if isinstance(image, (str, Path)):
        return Path(image).read_bytes()
    if isinstance(image, (bytes, bytearray, memoryview)):
        return bytes(image)
    raise TypeError(code)


def _read_image_format(data: bytes) -> Any:
    image_formats = feature_module("image_formats")
    try:
        return image_formats.read_image_format(data)
    except image_formats.ImageFormatError as error:
        if error.reason == "invalid":
            raise ValueError("picture_media_invalid") from error
        raise ValueError("image_media_unsupported") from error


def _merge_image_dimension(emu: int | None, alias: int | None, label: str) -> int | None:
    """Fold a python-docx style ``width``/``height`` alias into its ``*_emu`` value."""
    if alias is None:
        return emu
    if emu is not None:
        raise ValueError(f"{label}_conflict")
    # python-docx ``Length`` subclasses ``int``; emit a plain EMU integer.
    if isinstance(alias, int) and not isinstance(alias, bool):
        return int(alias)
    return alias


def _require_font_size(value: int | None) -> int | None:
    if value is None:
        return None
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or not 1 <= value <= _MAX_FONT_HALF_POINTS
    ):
        raise ValueError("run_font_size_invalid")
    return value


def _require_font_name(value: str | None) -> str | None:
    if value is None:
        return None
    if (
        not isinstance(value, str)
        or not value.strip()
        or len(value) > _MAX_FONT_NAME_CHARACTERS
        or not _valid_xml_text(value)
    ):
        raise ValueError("run_font_name_invalid")
    return value


def _require_color(value: str | None) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or _HEX_COLOR.fullmatch(value) is None:
        raise ValueError("run_color_invalid")
    return value.upper()


def _require_highlight(value: str | enum.Enum | None) -> str | None:
    """Resolve an ``ST_HighlightColor`` literal, Word/python-docx alias, or member."""
    if value is None:
        return None
    if isinstance(value, enum.Enum):
        highlight = _WD_COLOR_INDEX_HIGHLIGHTS.get(value.name)
    elif isinstance(value, str):
        highlight = _HIGHLIGHT_CANONICAL_MAP.get(value.strip().lower().replace("_", ""))
    else:
        highlight = None
    if highlight is None:
        raise ValueError("run_highlight_invalid")
    return highlight


def _require_core_date(value: datetime | date | str, label: str) -> str:
    # Naive datetimes are UTC, matching python-docx core properties.
    try:
        return _format_datetime(value)
    except ValueError as exc:
        raise ValueError(f"{label}_invalid") from exc


def _require_revision(value: int | str) -> str:
    # cp:revision is xsd:string; Word reads it as a non-negative decimal count.
    if isinstance(value, int) and not isinstance(value, bool) and value >= 0:
        return str(value)
    if (
        isinstance(value, str)
        and len(value) <= _MAX_TEXT_CHARACTERS
        and value.isascii()
        and value.isdigit()
    ):
        return value
    raise ValueError("core_revision_invalid")


def _needs_space_preservation(text: str) -> bool:
    return text.startswith(" ") or text.endswith(" ") or "  " in text


def _append_text_run(paragraph: ET.Element, text: str) -> None:
    """Append one plain ``w:r``/``w:t`` run to ``paragraph``; empty text adds none."""
    if not text:
        return
    text_element = ET.SubElement(ET.SubElement(paragraph, Q(W, "r")), Q(W, "t"))
    text_element.text = text
    if _needs_space_preservation(text):
        text_element.set(_XML_SPACE, "preserve")


def _is_index(value: int, length: int) -> bool:
    return not isinstance(value, bool) and isinstance(value, int) and 0 <= value < length


@dataclass(frozen=True)
class _GridSlot:
    """One grid position of a table row and the merged region covering it."""

    cell: ET.Element  # the physical w:tc of this row covering the position
    origin: ET.Element  # the region's top-left w:tc
    top: int
    left: int
    bottom: int
    right: int


def _grid_span(cell: ET.Element) -> int:
    span = cell.find(f"{Q(W, 'tcPr')}/{Q(W, 'gridSpan')}")
    return 1 if span is None else int(span.get(Q(W, "val"), "1"))


def _continues_vertical_merge(cell: ET.Element) -> bool:
    merge = cell.find(f"{Q(W, 'tcPr')}/{Q(W, 'vMerge')}")
    return merge is not None and merge.get(Q(W, "val"), "continue") == "continue"


def _table_layout(table: ET.Element) -> list[list[_GridSlot]]:
    """Map every row's grid columns to their cells and merged regions."""
    placements: list[list[tuple[int, int, ET.Element, ET.Element, int]]] = []
    above: dict[int, tuple[int, ET.Element, int]] = {}
    for row_idx, row in enumerate(table.findall(Q(W, "tr"))):
        row_placements: list[tuple[int, int, ET.Element, ET.Element, int]] = []
        current: dict[int, tuple[int, ET.Element, int]] = {}
        left = 0
        for cell in row.findall(Q(W, "tc")):
            right = left + _grid_span(cell) - 1
            previous = above.get(left)
            if _continues_vertical_merge(cell) and previous is not None and previous[0] == right:
                origin, top = previous[1], previous[2]
            else:
                origin, top = cell, row_idx
            current[left] = (right, origin, top)
            row_placements.append((left, right, cell, origin, top))
            left = right + 1
        placements.append(row_placements)
        above = current
    bottoms = {
        id(origin): row_idx
        for row_idx, row_placements in enumerate(placements)
        for _, _, _, origin, _ in row_placements
    }
    layout: list[list[_GridSlot]] = []
    for row_placements in placements:
        slots: list[_GridSlot] = []
        for left, right, cell, origin, top in row_placements:
            slot = _GridSlot(cell, origin, top, left, bottoms[id(origin)], right)
            slots.extend([slot] * (right - left + 1))
        layout.append(slots)
    return layout


def _slot_of(layout: list[list[_GridSlot]], cell: ET.Element) -> _GridSlot | None:
    return next((slot for row in layout for slot in row if slot.cell is cell), None)


def _cell_is_empty(cell: ET.Element) -> bool:
    """True when the cell holds only paragraphs without any content."""
    return all(
        child.tag == Q(W, "tcPr")
        or (child.tag == Q(W, "p") and all(node.tag == Q(W, "pPr") for node in child))
        for child in cell
    )


def _set_cell_property(
    properties: ET.Element, name: str, attributes: dict[str, str] | None
) -> None:
    """Replace or remove one ``w:tcPr`` child, keeping ``CT_TcPr`` sequence order."""
    for existing in properties.findall(Q(W, name)):
        properties.remove(existing)
    if attributes is None:
        return
    rank = _TC_PR_ORDER.index(name)
    position = len(properties)
    for index, child in enumerate(properties):
        local_name = child.tag.rsplit("}", 1)[-1]
        if local_name not in _TC_PR_ORDER[: rank + 1]:
            position = index
            break
    properties.insert(position, ET.Element(Q(W, name), attributes))


def _require_twips(value: object, code: str, minimum: int = 0) -> int:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or not minimum <= value <= _MAX_PAGE_DIMENSION
    ):
        raise ValueError(code)
    return value


def _apply_page_setup(
    section: ET.Element,
    *,
    orientation: str | None,
    page_width: int | None,
    page_height: int | None,
    margins: dict[str, int | None],
    columns: int | None,
    column_spacing: int | None,
) -> None:
    """Validate a page-setup change in full, then write it into ``section``."""
    page_size = section.find(Q(W, "pgSz"))
    page_margins = section.find(Q(W, "pgMar"))
    assert page_size is not None and page_margins is not None
    width = int(page_size.get(Q(W, "w"), "0"))
    height = int(page_size.get(Q(W, "h"), "0"))
    if orientation is not None:
        if not isinstance(orientation, str) or orientation.lower() not in _PAGE_ORIENTATIONS:
            raise ValueError("section_orientation_invalid")
        orientation = orientation.lower()
    if page_width is not None:
        width = _require_twips(page_width, "section_page_width_invalid", _MIN_PAGE_DIMENSION)
    if page_height is not None:
        height = _require_twips(page_height, "section_page_height_invalid", _MIN_PAGE_DIMENSION)
    if orientation is None:
        orientation = page_size.get(Q(W, "orient"), "portrait")
    elif (
        page_width is None
        and page_height is None
        and (width > height) != (orientation == "landscape")
    ):
        width, height = height, width
    if (orientation == "landscape" and width < height) or (
        orientation == "portrait" and width > height
    ):
        raise ValueError("section_orientation_conflict")
    margin_values = {
        attribute: int(page_margins.get(Q(W, attribute), "0"))
        for _, attribute in _PAGE_MARGIN_ATTRIBUTES
    }
    for keyword, attribute in _PAGE_MARGIN_ATTRIBUTES:
        value = margins.get(keyword)
        if value is not None:
            margin_values[attribute] = _require_twips(value, f"section_{keyword}_invalid")
    text_width = width - margin_values["left"] - margin_values["right"] - margin_values["gutter"]
    if text_width < 1 or height - margin_values["top"] - margin_values["bottom"] < 1:
        raise ValueError("section_margins_exceed_page")
    column_settings = section.find(Q(W, "cols"))
    count = 1 if column_settings is None else int(column_settings.get(Q(W, "num"), "1"))
    spacing = (
        _DEFAULT_COLUMN_SPACING
        if column_settings is None
        else int(column_settings.get(Q(W, "space"), str(_DEFAULT_COLUMN_SPACING)))
    )
    if columns is not None:
        if (
            isinstance(columns, bool)
            or not isinstance(columns, int)
            or not 1 <= columns <= _MAX_COLUMN_COUNT
        ):
            raise ValueError("section_columns_invalid")
        count = columns
    if column_spacing is not None:
        spacing = _require_twips(column_spacing, "section_column_spacing_invalid")
    if text_width - spacing * (count - 1) < count:
        raise ValueError("section_columns_exceed_page")

    page_size.set(Q(W, "w"), str(width))
    page_size.set(Q(W, "h"), str(height))
    if orientation == "landscape":
        page_size.set(Q(W, "orient"), "landscape")
    else:
        page_size.attrib.pop(Q(W, "orient"), None)
    for attribute, value in margin_values.items():
        page_margins.set(Q(W, attribute), str(value))
    if columns is None and column_spacing is None:
        return
    if column_settings is None:
        column_settings = ET.Element(Q(W, "cols"))
        _insert_sect_pr_child(section, column_settings)
    column_settings.set(Q(W, "space"), str(spacing))
    if count > 1:
        column_settings.set(Q(W, "num"), str(count))
    else:
        column_settings.attrib.pop(Q(W, "num"), None)


class DocumentBuilder:
    """Build or edit a valid document tree with an explicit insertion cursor.

    ``write_paragraph`` establishes the current insertion paragraph.  ``append_run``,
    ``append_image``, ``append_shape`` and ``append_hyperlink`` add to that paragraph
    unless a paragraph handle is supplied.
    ``move_to`` selects a body, header, footer or cell insertion point. Without a
    selected cursor, paragraph and table insertion appends to the body. ``open``
    retains an existing package and saves edits without rebuilding untouched parts.
    Explicit owned paragraph handles may compose detached subtrees; cursor targets
    must be attached to a live story.
    """

    def __init__(self) -> None:
        self._document = ET.Element(Q(W, "document"))
        self._body = ET.SubElement(self._document, Q(W, "body"))
        self._section_properties = ET.SubElement(self._body, Q(W, "sectPr"))
        ET.SubElement(
            self._section_properties,
            Q(W, "pgSz"),
            {Q(W, "w"): "12240", Q(W, "h"): "15840"},
        )
        ET.SubElement(
            self._section_properties,
            Q(W, "pgMar"),
            {
                Q(W, "top"): "1440",
                Q(W, "right"): "1440",
                Q(W, "bottom"): "1440",
                Q(W, "left"): "1440",
                Q(W, "header"): "720",
                Q(W, "footer"): "720",
                Q(W, "gutter"): "0",
            },
        )
        self._current_paragraph: DocumentParagraph | None = None
        self._paragraphs: list[DocumentParagraph] = []
        self._tables: list[DocumentTable] = []
        self._images: list[_ImagePart] = []
        self._paragraph_styles: set[str] = set()
        self._character_styles: set[str] = set()
        self._table_styles: set[str] = set()
        # Word display names for style IDs authored from spaced style names.
        self._style_names: dict[str, str] = {}
        self._hyperlinks: list[tuple[str, str, str]] = []
        self._drawing_count = 0
        self._compat_mode: int | None = None
        self._settings_values: dict[str, bytes | None] = {}
        self._compatibility_options: dict[str, bool | str | None] = {}
        # Document-default run font and half-point size (w:docDefaults/w:rPrDefault).
        self._default_run_font: str | None = None
        self._default_run_size: int | None = None
        self._core_properties: dict[str, str] = {}
        self._sections: list[DocumentSection] = [DocumentSection(self._section_properties, self)]
        # Header and footer parts keyed by their document relationship ID.
        self._header_footers: dict[str, DocumentHeaderFooter] = {}
        self._even_and_odd_headers = False
        # One list kind per w:num, numbered from numId 1.
        self._lists: list[str] = []
        self._session: Any = None
        self._source_roots: dict[str, ET.Element] = {}
        self._root_snapshots: dict[str, bytes] = {}
        self._source_paths: dict[str, dict[tuple[int, ...], ET.Element]] = {}
        self._source_node_snapshots: dict[str, bytes] = {}
        self._source_part_bytes: dict[str, bytes] = {}
        self._reserved_parts: set[str] = set()
        self._reserved_relationships: set[str] = set()
        self._number_id_base = 0
        self._story_relationships: dict[str, set[str]] = {}
        self._abstract_id_base = 0
        self._block_container = self._body
        self._block_index: int | None = None
        self._inline_container: ET.Element | None = None
        self._inline_index: int | None = None
        self._pending_text_offset: tuple[ET.Element, int] | None = None
        self._inline_snapshot: tuple[ET.Element, ...] | None = None
        self._hyperlink_parent_snapshot: tuple[ET.Element, tuple[ET.Element, ...]] | None = None
        self._block_snapshot: tuple[ET.Element, ...] | None = None
        self._pending_snapshot: bytes | None = None

    @classmethod
    def open(cls, source: Any) -> DocumentBuilder:
        """Open a path, seekable binary stream, bytes, or an existing Document.

        Only the conventional transitional main story is authored; other package
        layouts are refused before any publication.
        """
        from .document import Document, _main_part

        session = (
            source
            if isinstance(source, Document)
            else (
                Document.from_bytes(source, path="document.docx")
                if isinstance(source, (bytes, bytearray, memoryview))
                else Document(source)
            )
        )
        main = _main_part(session._graph)
        root = session._parse(session.part_bytes(main), editable=True)
        body = root.find(Q(W, "body"))
        if main != "word/document.xml" or root.tag != Q(W, "document") or body is None:
            raise NodeEditError("builder_package_layout_unsupported")
        builder = cls()
        builder._session = session
        builder._document = root
        builder._body = body
        builder._block_container = body
        builder._reserved_parts = set(session.part_names)
        builder._source_roots[main] = root
        builder._root_snapshots[main] = serialize_xml_with_preserved_qnames(root)
        sections = list(body.iter(Q(W, "sectPr")))
        if not sections or sections[-1] not in list(body):
            raise NodeEditError("builder_final_section_missing")
        builder._sections = [DocumentSection(element, builder) for element in sections]
        builder._section_properties = sections[-1]
        builder._paragraphs = [DocumentParagraph(p, builder) for p in body.findall(Q(W, "p"))]
        builder._tables = [DocumentTable(t, builder) for t in body.findall(Q(W, "tbl"))]
        relationships = session._graph.parts[main].relationships
        builder._reserved_relationships = {relation.id for relation in relationships}
        by_id = {relation.id: relation for relation in relationships}
        for section in sections:
            for reference in section:
                if reference.tag not in {Q(W, "headerReference"), Q(W, "footerReference")}:
                    continue
                rid = reference.get(Q(R, "id"), "")
                relation = by_id.get(rid)
                if relation is None or relation.is_external:
                    raise NodeEditError("builder_story_reference_invalid")
                part_name = relation.resolved_target()
                story = "header" if reference.tag == Q(W, "headerReference") else "footer"
                if relation.type != f"{_REL}/{story}":
                    raise NodeEditError("builder_story_reference_invalid")
                if part_name not in builder._source_roots:
                    story_root = session._parse(session.part_bytes(part_name), editable=True)
                    builder._source_roots[part_name] = story_root
                    builder._root_snapshots[part_name] = serialize_xml_with_preserved_qnames(
                        story_root
                    )
                if builder._source_roots[part_name].tag != Q(
                    W, "hdr" if story == "header" else "ftr"
                ):
                    raise NodeEditError("builder_story_reference_invalid")
                builder._header_footers[rid] = DocumentHeaderFooter(
                    builder._source_roots[part_name],
                    reference.get(Q(W, "type"), "default"),
                    rid,
                    part_name,
                    builder,
                )
        for part_name, story_root in builder._source_roots.items():
            builder._source_part_bytes[part_name] = session.part_bytes(part_name)
            top = body if part_name == main else story_root
            builder._source_node_snapshots[part_name] = serialize_xml_with_preserved_qnames(top)
            paths: dict[tuple[int, ...], ET.Element] = {}
            pending = [((), top)]
            while pending:
                path, item = pending.pop()
                paths[path] = item
                pending.extend(((*path, index), child) for index, child in enumerate(item))
            builder._source_paths[part_name] = paths
        for story_root in builder._source_roots.values():
            for drawing in story_root.iter(Q(WP, "docPr")):
                identifier = drawing.get("id", "")
                if identifier.isdecimal():
                    builder._drawing_count = max(builder._drawing_count, int(identifier))
        numbering = next(
            (
                relation.resolved_target()
                for relation in relationships
                if relation.type.endswith("/numbering") and not relation.is_external
            ),
            None,
        )
        if numbering is not None:
            numbering_root = session._parse(session.part_bytes(numbering), editable=True)
            for tag, attribute, state in (
                ("num", "numId", "_number_id_base"),
                ("abstractNum", "abstractNumId", "_abstract_id_base"),
            ):
                values = [
                    int(node.get(Q(W, attribute), "0"))
                    for node in numbering_root.findall(Q(W, tag))
                ]
                setattr(builder, state, max(values, default=0) + (state == "_abstract_id_base"))
        return builder

    def move_to(
        self, target: Any, *, position: str = "end", offset: int | None = None
    ) -> DocumentBuilder:
        """Select a node, handle or bookmark name for subsequent insertion.

        Paragraphs and runs accept ``start``/``end`` or a Unicode character
        ``offset`` in their ``w:t`` text. ``before``/``after`` choose paragraph/
        table block boundaries or run boundaries. Cells and header/footer handles
        select their story start/end. Bookmark names and fields choose result
        boundaries, not instruction text. Moving alone never splits a run.
        Interior offsets split mixed-content runs only when content is inserted.
        Nontext children consume no offset; at an interior text-node boundary,
        following children stay on the right of the insertion.
        """
        state = (
            self._current_paragraph,
            self._block_container,
            self._block_index,
            self._inline_container,
            self._inline_index,
            self._pending_text_offset,
            self._inline_snapshot,
            self._hyperlink_parent_snapshot,
            self._block_snapshot,
            self._pending_snapshot,
        )
        try:
            self._move_to_target(target, position=position, offset=offset)
        except NodeEditError:
            (
                self._current_paragraph,
                self._block_container,
                self._block_index,
                self._inline_container,
                self._inline_index,
                self._pending_text_offset,
                self._inline_snapshot,
                self._hyperlink_parent_snapshot,
                self._block_snapshot,
                self._pending_snapshot,
            ) = state
            raise
        self._capture_cursor(block=True)
        return self

    def _move_to_target(self, target: Any, *, position: str, offset: int | None) -> DocumentBuilder:
        if position not in {"start", "end", "before", "after"}:
            raise NodeEditError("builder_cursor_position_invalid")
        if isinstance(target, str):
            if offset is not None:
                raise NodeEditError("builder_cursor_offset_invalid")
            return self.move_to_bookmark(target, position=position)
        element = self._target_element(target)
        if element is self._document:
            element = self._body
        parents = self._parents()
        if element.tag == Q(W, "bookmarkStart"):
            if offset is not None:
                raise NodeEditError("builder_cursor_offset_invalid")
            return self._move_to_marker(element, position, bookmark=True)
        if element.tag in {Q(W, "fldSimple"), Q(W, "fldChar")}:
            if offset is not None:
                raise NodeEditError("builder_cursor_offset_invalid")
            return self._move_to_marker(element, position, bookmark=False)
        if element.tag == Q(W, "r"):
            if offset is not None and position in {"before", "after"}:
                raise NodeEditError("builder_cursor_offset_invalid")
            container = parents.get(element)
            paragraph = container
            while paragraph is not None and paragraph.tag != Q(W, "p"):
                paragraph = parents.get(paragraph)
            if paragraph is None or container is None:
                raise NodeEditError("builder_cursor_container_unsupported")
            if container.tag not in {
                Q(W, tag)
                for tag in (
                    "p",
                    "hyperlink",
                    "fldSimple",
                    "sdtContent",
                    "smartTag",
                    "customXml",
                    "ins",
                    "moveTo",
                    "dir",
                    "bdo",
                )
            }:
                raise NodeEditError("builder_cursor_container_unsupported")
            self.move_to(DocumentParagraph(paragraph, self))
            if offset is None:
                self._inline_container = container
                self._inline_index = list(container).index(element) + (position in {"end", "after"})
            else:
                self._place_text_cursor(element, offset)
            return self
        self._inline_container = None
        self._inline_index = None
        self._pending_text_offset = None
        self._current_paragraph = None
        if element.tag in {Q(W, "body"), Q(W, "hdr"), Q(W, "ftr"), Q(W, "tc")}:
            if position not in {"start", "end"} or offset is not None:
                raise NodeEditError("builder_cursor_position_invalid")
            self._block_container = element
            blocks = [
                index
                for index, child in enumerate(element)
                if child.tag not in {Q(W, "tcPr"), Q(W, "sectPr")}
            ]
            self._block_index = (
                (blocks[0] if blocks else self._block_end(element))
                if position == "start"
                else self._block_end(element)
            )
            boundary = (
                element[blocks[0]]
                if position == "start" and blocks
                else element[blocks[-1]]
                if blocks
                else None
            )
            if boundary is not None and boundary.tag == Q(W, "p"):
                paragraph = boundary
                self._current_paragraph = DocumentParagraph(paragraph, self)
                self._inline_container = paragraph
                self._inline_index = (
                    (1 if len(paragraph) and paragraph[0].tag == Q(W, "pPr") else 0)
                    if position == "start"
                    else len(paragraph)
                )
            return self
        parent = parents.get(element)
        if parent is None or element.tag not in {Q(W, "p"), Q(W, "tbl")}:
            raise NodeEditError("builder_cursor_target_unsupported")
        if parent.tag not in {Q(W, "body"), Q(W, "hdr"), Q(W, "ftr"), Q(W, "tc")}:
            raise NodeEditError("builder_cursor_container_unsupported")
        self._block_container = parent
        index = list(parent).index(element)
        self._block_index = index if position in {"start", "before"} else index + 1
        if element.tag == Q(W, "p") and position in {"start", "end"}:
            self._current_paragraph = DocumentParagraph(element, self)
            if offset is None:
                self._inline_container = element
                self._inline_index = (
                    (1 if len(element) and element[0].tag == Q(W, "pPr") else 0)
                    if position == "start"
                    else len(element)
                )
            else:
                self._place_text_cursor(element, offset)
        elif offset is not None:
            raise NodeEditError("builder_cursor_offset_invalid")
        return self

    def move_to_document_start(self) -> DocumentBuilder:
        """Select the main story's start for inline or block insertion."""
        return self.move_to(self._body, position="start")

    def move_to_document_end(self) -> DocumentBuilder:
        """Select the main story's end, before the final section properties."""
        return self.move_to(self._body, position="end")

    def move_to_bookmark(self, name: str, *, position: str = "end") -> DocumentBuilder:
        """Move inside a named bookmark, or outside it with before/after."""
        matches = [
            item
            for root in self._story_roots()
            for item in root.iter(Q(W, "bookmarkStart"))
            if item.get(Q(W, "name")) == name
        ]
        if len(matches) != 1:
            raise NodeEditError("builder_bookmark_target_invalid")
        return self._move_to_marker(matches[0], position, bookmark=True)

    def move_to_field(self, field: Any = 0, *, position: str = "end") -> DocumentBuilder:
        """Move to a field result by node, zero-based index, or merge-field name."""
        if isinstance(field, str):
            return self.move_to_merge_field(field, position=position)
        if isinstance(field, int) and not isinstance(field, bool):
            fields = [
                item
                for root in self._story_roots()
                for item in root.iter()
                if item.tag == Q(W, "fldSimple")
                or (item.tag == Q(W, "fldChar") and item.get(Q(W, "fldCharType")) == "begin")
            ]
            if not _is_index(field, len(fields)):
                raise NodeEditError("builder_field_target_invalid")
            element = fields[field]
        else:
            element = self._target_element(field)
        return self._move_to_marker(element, position, bookmark=False)

    def move_to_merge_field(
        self, name: str, *, position: str = "end", occurrence: int | None = None
    ) -> DocumentBuilder:
        """Select a MERGEFIELD by case-insensitive name in editable story order.

        Names in quoted instructions are supported. Duplicate names require a
        zero-based ``occurrence``; omission selects only an unambiguous name.
        """
        fields_module = feature_module("fields")

        if not isinstance(name, str) or not name or not _valid_xml_text(name):
            raise NodeEditError("builder_merge_field_name_invalid")
        matches = []
        try:
            for root in self._story_roots():
                for field in fields_module.field_nodes(root, W, allow_instruction_nesting=True):
                    if field.code != "MERGEFIELD":
                        continue
                    tokens = fields_module.tokenize_instruction(field.instruction)
                    if len(tokens) < 2 or (
                        not tokens[1].is_quoted and tokens[1].text.startswith("\\")
                    ):
                        raise NodeEditError("builder_merge_field_instruction_invalid")
                    if tokens[1].text.casefold() == name.casefold():
                        matches.append(field.begin)
        except fields_module.FieldStructureError as error:
            raise NodeEditError(error.code) from None
        if occurrence is None:
            if len(matches) != 1:
                raise NodeEditError("builder_merge_field_target_invalid")
            element = matches[0]
        elif _is_index(occurrence, len(matches)):
            element = matches[occurrence]
        else:
            raise NodeEditError("builder_merge_field_target_invalid")
        return self._move_to_marker(element, position, bookmark=False)

    def insert_text(self, text: str, **formatting: Any) -> DocumentRun:
        """Insert a formatted run at the selected inline cursor and advance it."""
        return self.append_run(text, **formatting)

    @property
    def document(self) -> ET.Element:
        """The root ``w:document`` element for read-only tree inspection."""
        return self._document

    @property
    def paragraphs(self) -> tuple[DocumentParagraph, ...]:
        """Paragraphs inserted at document body level, in insertion order."""
        return tuple(self._paragraphs)

    @property
    def tables(self) -> tuple[DocumentTable, ...]:
        """Tables inserted at document body level, in insertion order."""
        return tuple(self._tables)

    @property
    def sections(self) -> tuple[DocumentSection, ...]:
        """The document's sections in order; the last one receives new body content."""
        return tuple(self._sections)

    def write_paragraph(
        self,
        text: str = "",
        *,
        style: str | None = None,
    ) -> DocumentParagraph:
        """Insert a paragraph at the block cursor, then select its end for runs."""
        text = _require_text(text, "paragraph_text")
        resolved_style = _require_style(style, "paragraph_style")
        return self._insert_body_paragraph(self._new_paragraph_element(resolved_style), text)

    def add_section(
        self, start: str | enum.Enum = "new_page", **page_setup: Any
    ) -> DocumentSection:
        """End the current section and start a new one, like python-docx ``add_section``.

        ``start`` is a python-docx ``WD_SECTION`` member or name (``new_page``,
        ``continuous``, ``even_page``, ``odd_page``, ``new_column``) or its
        ``ST_SectionMark`` value.  The new section copies the current page size,
        orientation, margins and columns, then applies ``page_setup``, which takes
        the keywords of :meth:`DocumentSection.set_page_setup`.  Headers and footers
        are not copied: Word shows the previous section's until the new section
        defines its own.  The ending section's properties move into its last body
        paragraph, or into a new empty paragraph when it ends with a table or
        another section break.
        """
        break_type = _canonical_break_type(start)
        if break_type is None:
            raise ValueError("section_start_invalid")
        properties = ET.Element(Q(W, "sectPr"))
        for name in ("pgSz", "pgMar", "cols"):
            inherited = self._section_properties.find(Q(W, name))
            if inherited is not None:
                properties.append(copy.deepcopy(inherited))
        if break_type != "nextPage":
            _insert_sect_pr_child(properties, ET.Element(Q(W, "type"), {Q(W, "val"): break_type}))
        section = DocumentSection(properties, self)
        if page_setup:
            section.set_page_setup(**page_setup)
        carrier = self._body[-2] if len(self._body) > 1 else None
        if (
            carrier is None
            or carrier.tag != Q(W, "p")
            or carrier.find(f"{Q(W, 'pPr')}/{Q(W, 'sectPr')}") is not None
        ):
            carrier = ET.Element(Q(W, "p"))
            self._body.insert(len(self._body) - 1, carrier)
            self._paragraphs.append(DocumentParagraph(carrier, self))
        carrier_properties = carrier.find(Q(W, "pPr"))
        if carrier_properties is None:
            carrier_properties = ET.Element(Q(W, "pPr"))
            carrier.insert(0, carrier_properties)
        # sectPr is the last CT_PPr child this builder writes, so appending keeps order.
        self._body.remove(self._section_properties)
        carrier_properties.append(self._section_properties)
        self._body.append(properties)
        self._section_properties = properties
        self._sections.append(section)
        self._current_paragraph = None
        self._block_container = self._body
        self._block_index = None
        self._inline_container = None
        self._inline_index = None
        self._pending_text_offset = None
        return section

    def add_list(self, kind: str = "numbered") -> DocumentList:
        """Start a ``numbered`` or ``bulleted`` multi-level list whose numbering restarts."""
        if not isinstance(kind, str) or kind not in _LIST_KINDS:
            raise ValueError("list_kind_invalid")
        self._lists.append(kind)
        return DocumentList(kind, self._number_id_base + len(self._lists), self)

    def append_run(
        self,
        text: str,
        *,
        paragraph: DocumentParagraph | None = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        strike: bool = False,
        style: str | None = None,
        font_name: str | None = None,
        font_size: int | None = None,
        color: str | None = None,
        subscript: bool = False,
        superscript: bool = False,
        highlight: str | enum.Enum | None = None,
    ) -> DocumentRun:
        """Append a formatted text run to a builder paragraph.

        ``font_size`` is in half-points (``24`` is 12 pt), ``color`` is a
        six-digit RGB hex value, and ``font_name`` applies to the ASCII,
        high-ANSI and complex-script font slots.  ``subscript`` and
        ``superscript`` set the run's baseline alignment and are mutually exclusive.
        ``style`` is a character style ID or a spaced Word style name such as
        ``Subtle Emphasis``, published under its Word styleId ``SubtleEmphasis``.
        ``highlight`` is an ``ST_HighlightColor`` literal in any case, a Word or
        python-docx ``WD_COLOR_INDEX`` name such as ``turquoise`` or ``dark_blue``,
        or a ``WD_COLOR_INDEX`` member.
        """
        text = _require_text(text, "run_text")
        flags = (bold, italic, underline, strike, subscript, superscript)
        if not all(isinstance(value, bool) for value in flags) or (subscript and superscript):
            raise ValueError("run_format_invalid")
        resolved_style = _require_style(style, "run_style")
        font_name = _require_font_name(font_name)
        font_size = _require_font_size(font_size)
        color = _require_color(color)
        highlight = _require_highlight(highlight)
        target = self._resolve_paragraph(paragraph)
        self._validate_inline_cursor(target, Q(W, "r"), use_cursor=paragraph is None)
        run = ET.Element(Q(W, "r"))
        # Children follow the CT_RPr sequence of ECMA-376 Part 1, 17.3.2.28.
        properties = ET.Element(Q(W, "rPr"))
        if resolved_style is not None:
            style_id = self._register_style(self._character_styles, resolved_style)
            ET.SubElement(properties, Q(W, "rStyle"), {Q(W, "val"): style_id})
        if font_name is not None:
            ET.SubElement(
                properties,
                Q(W, "rFonts"),
                {
                    Q(W, "ascii"): font_name,
                    Q(W, "hAnsi"): font_name,
                    Q(W, "cs"): font_name,
                },
            )
        if bold:
            ET.SubElement(properties, Q(W, "b"))
        if italic:
            ET.SubElement(properties, Q(W, "i"))
        if strike:
            ET.SubElement(properties, Q(W, "strike"))
        if color is not None:
            ET.SubElement(properties, Q(W, "color"), {Q(W, "val"): color})
        if font_size is not None:
            ET.SubElement(properties, Q(W, "sz"), {Q(W, "val"): str(font_size)})
            ET.SubElement(properties, Q(W, "szCs"), {Q(W, "val"): str(font_size)})
        if highlight is not None:
            ET.SubElement(properties, Q(W, "highlight"), {Q(W, "val"): highlight})
        if underline:
            ET.SubElement(properties, Q(W, "u"), {Q(W, "val"): "single"})
        if subscript or superscript:
            ET.SubElement(
                properties,
                Q(W, "vertAlign"),
                {Q(W, "val"): "subscript" if subscript else "superscript"},
            )
        if len(properties):
            run.append(properties)
        text_element = ET.SubElement(run, Q(W, "t"))
        text_element.text = text
        if _needs_space_preservation(text):
            text_element.set(_XML_SPACE, "preserve")
        self._insert_inline(target, run, use_cursor=paragraph is None)
        self._current_paragraph = target
        return DocumentRun(run, target)

    def add_table(
        self, rows: int, cols: int, style: str | None = None, *, look: TableLookSpec | None = None
    ) -> DocumentTable:
        """Insert an empty ``rows`` by ``cols`` table at the selected block cursor.

        Columns split the section's text width between its margins evenly, as
        python-docx does.
        ``look`` selects conditional table-style flags with :class:`TableLookSpec`.
        """
        for value, code in (
            (rows, "table_rows_invalid"),
            (cols, "table_columns_invalid"),
        ):
            if isinstance(value, bool) or not isinstance(value, int) or value < 1:
                raise ValueError(code)
        page_size = self._section_properties.find(Q(W, "pgSz"))
        margins = self._section_properties.find(Q(W, "pgMar"))
        text_width = (
            (int(page_size.get(Q(W, "w"), "12240")) if page_size is not None else 12240)
            - (int(margins.get(Q(W, "left"), "1440")) if margins is not None else 1440)
            - (int(margins.get(Q(W, "right"), "1440")) if margins is not None else 1440)
        )
        return self.insert_table(
            [("",) * cols] * rows,
            column_widths=(max(1, text_width // cols),) * cols,
            style=style,
            look=look,
        )

    def insert_table(
        self,
        rows: Sequence[Sequence[str]],
        *,
        column_widths: Sequence[int] | None = None,
        style: str | None = None,
        look: TableLookSpec | None = None,
    ) -> DocumentTable:
        """Insert a rectangular table at the selected block cursor.

        ``style`` selects its Word style. ``look`` controls conditional first/last
        row/column styling and banding using the shared :class:`TableLookSpec`.
        Without an explicit cursor move, subsequent inline content stays in the
        current paragraph. At an explicit block cursor it follows the table.
        """
        if isinstance(rows, (str, bytes)) or not isinstance(rows, Sequence) or not rows:
            raise ValueError("table_rows_invalid")
        normalized_rows: list[tuple[str, ...]] = []
        column_count: int | None = None
        for row in rows:
            if isinstance(row, (str, bytes)) or not isinstance(row, Sequence) or not row:
                raise ValueError("table_row_invalid")
            normalized = tuple(_require_text(cell, "table_cell_text") for cell in row)
            if column_count is None:
                column_count = len(normalized)
            elif len(normalized) != column_count:
                raise ValueError("table_rows_rectangular_required")
            normalized_rows.append(normalized)
        assert column_count is not None
        if column_widths is None:
            widths = (1440,) * column_count
        else:
            if isinstance(column_widths, (str, bytes)) or len(column_widths) != column_count:
                raise ValueError("table_column_widths_invalid")
            widths = tuple(_require_emu(width, "table_column_width") for width in column_widths)
            if any(width is None for width in widths):
                raise ValueError("table_column_widths_invalid")
        resolved_style = _require_style(style, "table_style")
        if look is not None and not isinstance(look, TableLookSpec):
            raise NodeEditError("builder_table_look_invalid")

        table = ET.Element(Q(W, "tbl"))
        table_properties = ET.SubElement(table, Q(W, "tblPr"))
        if resolved_style is not None:
            style_id = self._register_style(self._table_styles, resolved_style)
            ET.SubElement(table_properties, Q(W, "tblStyle"), {Q(W, "val"): style_id})
        ET.SubElement(
            table_properties,
            Q(W, "tblW"),
            {Q(W, "w"): str(sum(widths)), Q(W, "type"): "dxa"},
        )
        table_grid = ET.SubElement(table, Q(W, "tblGrid"))
        for width in widths:
            ET.SubElement(table_grid, Q(W, "gridCol"), {Q(W, "w"): str(width)})
        for row in normalized_rows:
            table_row = ET.SubElement(table, Q(W, "tr"))
            for width, cell_text in zip(widths, row, strict=True):
                cell = ET.SubElement(table_row, Q(W, "tc"))
                cell_properties = ET.SubElement(cell, Q(W, "tcPr"))
                ET.SubElement(
                    cell_properties,
                    Q(W, "tcW"),
                    {Q(W, "w"): str(width), Q(W, "type"): "dxa"},
                )
                _append_text_run(ET.SubElement(cell, Q(W, "p")), cell_text)
        self._insert_block(table)
        result = DocumentTable(table, self)
        if look is not None:
            result.set_style_context(look)
        if self._block_container is self._body:
            self._tables = [DocumentTable(t, self) for t in self._body.findall(Q(W, "tbl"))]
        # Legacy append-mode tables do not select a new inline paragraph.
        # A moved block cursor must instead advance inline insertion past the table.
        if self._block_index is not None:
            self._current_paragraph = None
        self._inline_container = None
        self._inline_index = None
        self._pending_text_offset = None
        self._capture_cursor()
        return result

    def append_image(
        self,
        image: bytes | bytearray | memoryview | str | Path,
        *,
        width_emu: int | None = None,
        height_emu: int | None = None,
        width: int | None = None,
        height: int | None = None,
        alt_text: str | None = None,
        paragraph: DocumentParagraph | None = None,
        anchor: Any = None,
        svg_fallback: bytes | bytearray | memoryview | str | Path | None = None,
    ) -> DocumentImage:
        """Append a picture as a standard DrawingML picture run.

        PNG, JPEG, GIF, BMP, TIFF, SVG, EMF, and placeable WMF are accepted and
        identified by signature. An SVG requires ``svg_fallback``, the PNG raster
        embedded beside its Office ``asvg:svgBlip`` extension; other formats refuse
        a fallback.

        ``width`` and ``height`` are python-docx compatible aliases for
        ``width_emu`` and ``height_emu``. Supplying both forms is rejected.
        BMP, GIF, TIFF, SVG, EMF, and WMF use their intrinsic header size and DPI;
        one dimension follows their intrinsic physical aspect ratio. PNG and JPEG
        retain their existing 96-DPI pixel sizing and pixel aspect ratio.
        Without ``anchor`` the picture is inline; with a
        :class:`~wolfdocx.picture_insert.PictureAnchorSpec` it is a page-relative
        ``wp:anchor`` floating picture carried by ``paragraph``.
        """
        picture_insert = feature_module("picture_insert")
        image_formats = feature_module("image_formats")
        if anchor is not None and not isinstance(anchor, picture_insert.PictureAnchorSpec):
            raise TypeError("image_anchor_invalid")
        width_emu = _merge_image_dimension(width_emu, width, "image_width")
        height_emu = _merge_image_dimension(height_emu, height, "image_height")
        data = _image_bytes(image, "image_invalid")
        image_format = _read_image_format(data)
        fallback: bytes | None = None
        if image_format.name == "svg":
            if svg_fallback is None:
                raise ValueError("image_svg_fallback_required")
            fallback = _image_bytes(svg_fallback, "image_svg_fallback_invalid")
            try:
                fallback_format = image_formats.read_image_format(fallback)
            except image_formats.ImageFormatError as error:
                raise ValueError("image_svg_fallback_invalid") from error
            if fallback_format.media_type != "image/png":
                raise ValueError("image_svg_fallback_invalid")
        elif svg_fallback is not None:
            raise ValueError("image_svg_fallback_unexpected")
        width_emu = _require_emu(width_emu, "image_width")
        height_emu = _require_emu(height_emu, "image_height")
        if image_format.name in {"png", "jpeg"}:
            # Preserve the established PNG/JPEG bytes, including unequal header DPI.
            pixel_width, pixel_height = image_format.width_px, image_format.height_px
            if width_emu is None and height_emu is None:
                width_emu, height_emu = pixel_width * 9_525, pixel_height * 9_525
            elif height_emu is None:
                height_emu = max(1, round(width_emu * pixel_height / pixel_width))
            elif width_emu is None:
                width_emu = max(1, round(height_emu * pixel_width / pixel_height))
        else:
            width_emu, height_emu = image_formats.fit_extent(image_format, width_emu, height_emu)
        _require_emu(width_emu, "image_width")
        _require_emu(height_emu, "image_height")
        if alt_text is not None:
            alt_text = _require_text(alt_text, "image_alt_text")
        target = self._resolve_paragraph(paragraph)
        self._validate_inline_cursor(target, Q(W, "r"), use_cursor=paragraph is None)
        source_part = self._paragraph_part(target)
        index = len(self._images) + 1
        relationship_id = self._new_relationship_id(f"rIdImage{index}", part_name=source_part)
        part_name = self._new_part_name(f"word/media/image{index}.{image_format.extension}")
        fallback_relationship_id: str | None = None
        fallback_part_name: str | None = None
        if fallback is not None:
            fallback_relationship_id = self._new_relationship_id(
                f"rIdImage{index + 1}", part_name=source_part
            )
            fallback_part_name = self._new_part_name(f"word/media/image{index + 1}.png")
        run = picture_insert._build_drawing_run(
            doc_pr_id=self._next_drawing_id(),
            name=f"Image {index}",
            rel_id=fallback_relationship_id or relationship_id,
            cx=width_emu,
            cy=height_emu,
            alt_text=alt_text,
            word_namespace=W,
            anchor=anchor,
            svg_rel_id=relationship_id if fallback is not None else None,
        )
        self._insert_inline(target, run, use_cursor=paragraph is None)
        self._images.append(
            _ImagePart(relationship_id, part_name, image_format.media_type, data, source_part)
        )
        if fallback is not None:
            assert fallback_relationship_id is not None and fallback_part_name is not None
            self._images.append(
                _ImagePart(
                    fallback_relationship_id, fallback_part_name, "image/png", fallback, source_part
                )
            )
        self._current_paragraph = target
        return DocumentImage(
            run,
            target,
            relationship_id,
            part_name,
            image_format=image_format,
            fallback_relationship_id=fallback_relationship_id,
            fallback_part_name=fallback_part_name,
        )

    def append_shape(
        self,
        preset: str,
        *,
        width_emu: int,
        height_emu: int,
        paragraph: DocumentParagraph | None = None,
    ) -> DocumentShape:
        """Append an inline DrawingML preset AutoShape (``wps:wsp``) run."""
        shapes = feature_module("shapes")
        if not isinstance(preset, str) or preset not in shapes.PRESET_SHAPE_TYPES:
            raise ValueError("shape_preset_invalid")
        if _require_emu(width_emu, "shape_width") is None:
            raise ValueError("shape_width_invalid")
        if _require_emu(height_emu, "shape_height") is None:
            raise ValueError("shape_height_invalid")
        target = self._resolve_paragraph(paragraph)
        self._validate_inline_cursor(target, Q(W, "r"), use_cursor=paragraph is None)
        doc_pr_id = self._next_drawing_id()
        run = shapes._build_inline_shape_run(
            doc_pr_id=doc_pr_id,
            name=f"Shape {doc_pr_id}",
            preset=preset,
            cx=width_emu,
            cy=height_emu,
            word_namespace=W,
        )
        self._insert_inline(target, run, use_cursor=paragraph is None)
        self._current_paragraph = target
        return DocumentShape(run, target, preset)

    def append_hyperlink(
        self,
        url: str | None = None,
        text: str | None = None,
        *,
        anchor: str | None = None,
        tooltip: str | None = None,
        paragraph: DocumentParagraph | None = None,
    ) -> DocumentHyperlink:
        """Append a ``w:hyperlink`` run styled with ``Hyperlink``.

        Supply exactly one target: ``url`` for an external link published as a
        story-part relationship, or ``anchor`` for an internal link to the named
        bookmark, written as ``w:anchor`` with no relationship.

        At a cursor inside a direct paragraph hyperlink, split that hyperlink
        around the new paragraph sibling rather than nesting links. Both retained
        halves keep the original attributes and children. Hyperlink ancestors
        requiring a split of another wrapper raise :class:`NodeEditError`.
        """
        links = feature_module("links")
        if url is not None:
            if anchor is not None:
                raise ValueError("hyperlink_target_ambiguous")
            target_error = links._validate_url(url)
        elif anchor is not None:
            target_error = links._validate_anchor_name(anchor)
        else:
            raise ValueError("hyperlink_target_required")
        if target_error is not None:
            raise ValueError(target_error)
        text = _require_text(text, "hyperlink_text")
        if not text:
            raise ValueError("hyperlink_text_invalid")
        if tooltip is not None:
            tooltip = _require_text(tooltip, "hyperlink_tooltip")
            if not tooltip or len(tooltip) > _MAX_TOOLTIP_CHARACTERS:
                raise ValueError("hyperlink_tooltip_invalid")
        target = self._resolve_paragraph(paragraph)
        self._validate_inline_cursor(target, Q(W, "hyperlink"), use_cursor=paragraph is None)
        source_part = self._paragraph_part(target)
        relationship_id: str | None = None
        attributes: dict[str, str] = {}
        if url is not None:
            relationship_id = self._new_relationship_id(
                f"rIdHyperlink{len(self._hyperlinks) + 1}", part_name=source_part
            )
            self._hyperlinks.append((relationship_id, url, source_part))
            attributes[Q(R, "id")] = relationship_id
        if anchor is not None:
            attributes[Q(W, "anchor")] = anchor
        attributes[Q(W, "history")] = "1"
        hyperlink = ET.Element(Q(W, "hyperlink"), attributes)
        if tooltip is not None:
            hyperlink.set(Q(W, "tooltip"), tooltip)
        run = ET.SubElement(hyperlink, Q(W, "r"))
        properties = ET.SubElement(run, Q(W, "rPr"))
        ET.SubElement(properties, Q(W, "rStyle"), {Q(W, "val"): _HYPERLINK_STYLE})
        text_element = ET.SubElement(run, Q(W, "t"))
        text_element.text = text
        if _needs_space_preservation(text):
            text_element.set(_XML_SPACE, "preserve")
        self._insert_inline(target, hyperlink, use_cursor=paragraph is None)
        self._character_styles.add(_HYPERLINK_STYLE)
        self._current_paragraph = target
        return DocumentHyperlink(hyperlink, target, relationship_id, url, anchor)

    def set_compat_mode(self, mode: int) -> None:
        """Select the Word version whose layout rules the saved document declares.

        ``15`` opens in current mode in Word 2013 and later; ``14`` and ``12`` request
        Word 2010 and Word 2007 compatibility.  The value is written to
        ``word/settings.xml`` as ``w:compat/w:compatSetting`` named
        ``compatibilityMode``, which is the schema-valid form Word reads; ECMA-376
        defines no ``w:compatMode`` element.
        """
        if isinstance(mode, bool) or not isinstance(mode, int) or mode not in _COMPAT_MODES:
            raise ValueError("compat_mode_invalid")
        self._compat_mode = mode

    def set_view(self, view: str | None) -> None:
        """Set the document view; ``None`` removes the view declaration."""
        self._set_settings_value("view", view)

    def set_zoom(self, percent: int | None, *, mode: str = "none") -> None:
        """Set an explicit zoom percentage and native Word zoom mode."""
        self._set_settings_value(
            "zoom", {"percent": percent, "mode": mode} if percent is not None else None
        )

    def set_write_protection(
        self,
        *,
        recommended: bool = False,
        password: str | None = None,
        salt: bytes | None = None,
        spin_count: int = 100000,
    ) -> None:
        """Declare a write reservation, optionally using a SHA-512 verifier.

        This is not encryption. Caller-supplied salt makes the verifier
        deterministic; otherwise the shared routine generates a random salt.
        """
        value: dict[str, Any] = {"recommended": recommended}
        if password is not None:
            value.update(password=password, salt=salt, spin_count=spin_count)
        elif salt is not None or spin_count != 100000:
            raise ValueError("document_settings_password_required")
        self._set_settings_value("write_protection", value)

    def remove_write_protection(self) -> None:
        """Remove a previously declared write reservation."""
        self._set_settings_value("write_protection", None)

    def set_compatibility_option(self, name: str, value: bool | str | None) -> None:
        """Set a native boolean or named string compatibility option; null removes."""
        # Validate through the same source-bound compatibility writer.
        document_settings = feature_module("document_settings")
        document_settings._compat_patch(self._xml(ET.Element(Q(W, "settings"))), {name: value})
        self._compatibility_options[name] = value

    def _set_settings_value(self, name: str, value: Any) -> None:
        if value is None:
            self._settings_values[name] = None
            return
        element = feature_module("document_settings")._element(name, value)
        self._settings_values[name] = ET.tostring(element, encoding="utf-8")

    def set_default_run_font(self, font_name: str, font_size: int | None = None) -> None:
        """Declare the document-default run font and, optionally, its size.

        The values become ``w:docDefaults/w:rPrDefault`` in ``word/styles.xml``, so
        runs and empty paragraphs without their own ``w:rFonts`` or ``w:sz`` resolve
        to them. ``font_size`` is in half-points, like :meth:`append_run`.
        """
        resolved = _require_font_name(font_name)
        if resolved is None:
            raise ValueError("run_font_name_invalid")
        self._default_run_size = _require_font_size(font_size)
        self._default_run_font = resolved

    def set_core_properties(
        self,
        *,
        title: str | None = None,
        author: str | None = None,
        subject: str | None = None,
        keywords: str | None = None,
        description: str | None = None,
        revision: int | str | None = None,
        comments: str | None = None,
        category: str | None = None,
        content_status: str | None = None,
        content_type: str | None = None,
        identifier: str | None = None,
        language: str | None = None,
        last_modified_by: str | None = None,
        version: str | None = None,
        created: datetime | date | str | None = None,
        modified: datetime | date | str | None = None,
        last_printed: datetime | date | str | None = None,
    ) -> None:
        """Set built-in document properties written to ``docProps/core.xml`` on save.

        Keywords follow python-docx ``CoreProperties`` names.  Supplied values replace
        earlier ones; ``None`` leaves a property unchanged.  ``author`` is stored as
        ``dc:creator``; ``comments`` is python-docx's name for ``description`` and must
        agree with it when both are given.  ``revision`` must be a non-negative integer
        or a string of decimal digits.  ``created``, ``modified`` and ``last_printed``
        take a ``datetime``, ``date`` or ISO 8601 string; naive values are UTC.  Every
        value is validated before any is stored, so a refused call changes nothing.
        """
        supplied = {
            "title": title,
            "author": author,
            "subject": subject,
            "keywords": keywords,
            "description": description,
            "comments": comments,
            "category": category,
            "content_status": content_status,
            "content_type": content_type,
            "identifier": identifier,
            "language": language,
            "last_modified_by": last_modified_by,
            "version": version,
        }
        updates = {
            name: _require_text(value, f"core_{name}")
            for name, value in supplied.items()
            if value is not None
        }
        if "comments" in updates:
            comments = updates.pop("comments")
            if updates.setdefault("description", comments) != comments:
                raise ValueError("core_description_conflict")
        if revision is not None:
            updates["revision"] = _require_revision(revision)
        for name, moment in (
            ("created", created),
            ("modified", modified),
            ("last_printed", last_printed),
        ):
            if moment is not None:
                updates[name] = _require_core_date(moment, f"core_{name}")
        self._core_properties.update(updates)

    def save(self, destination: str | Path) -> DocumentSaveReport:
        """Atomically publish the composed document package at ``destination``."""
        path = Path(destination)
        if not path.name or path.suffix.lower() != ".docx":
            raise ValueError("builder_destination_invalid")
        if self._session is not None:
            self._sync_session()
            report = self._session.save(path)
            if report.status != "published":
                raise NodeEditError(report.code or "builder_save_refused")
            return DocumentSaveReport(
                destination=path,
                sha256=report.destination_sha256,
                paragraph_count=len(self.paragraphs),
                table_count=len(self.tables),
                image_count=len(self._images),
            )
        path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="wb",
            suffix=".docx",
            prefix=f".{path.stem}.",
            dir=path.parent,
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED) as package:
                for name, data in self._authored_parts().items():
                    package.writestr(self._zip_info(name), data)
            os.replace(temporary_path, path)
        except BaseException:
            temporary_path.unlink(missing_ok=True)
            raise
        return DocumentSaveReport(
            destination=path,
            sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
            paragraph_count=len(self._paragraphs),
            table_count=len(self._tables),
            image_count=len(self._images),
        )

    @property
    def _has_styles(self) -> bool:
        return bool(
            self._paragraph_styles
            or self._character_styles
            or self._table_styles
            or self._default_run_font is not None
        )

    @property
    def _has_settings(self) -> bool:
        return bool(
            self._compat_mode is not None
            or self._even_and_odd_headers
            or self._settings_values
            or self._compatibility_options
        )

    def _register_style(self, style_ids: set[str], style: tuple[str, str]) -> str:
        style_id, name = style
        style_ids.add(style_id)
        if name != style_id:
            self._style_names.setdefault(style_id, name)
        return style_id

    def _next_drawing_id(self) -> int:
        # wp:docPr/@id must be unique across every drawing in the document.
        self._drawing_count += 1
        return self._drawing_count

    def _resolve_paragraph(self, paragraph: DocumentParagraph | None) -> DocumentParagraph:
        if paragraph is None:
            if self._current_paragraph is None:
                if self._inline_container is not None or self._pending_text_offset is not None:
                    raise NodeEditError("builder_cursor_target_stale")
                return self.write_paragraph()
            self._target_element(self._current_paragraph)
            return self._resolve_paragraph(self._current_paragraph)
        # Importers may author owned paragraphs before attaching their table.
        # Cursor targets separately require membership in a live story tree.
        if not isinstance(paragraph, DocumentParagraph) or paragraph._builder is not self:
            raise NodeEditError("paragraph_invalid")
        return paragraph

    def _paragraph_part(self, paragraph: DocumentParagraph) -> str:
        """Resolve relationship ownership, retaining detached importer composition."""
        for part in self._header_footers.values():
            if any(node is paragraph.element for node in part.element.iter(Q(W, "p"))):
                return part.part_name
        return "word/document.xml"

    def _new_paragraph_element(self, style: tuple[str, str] | None) -> ET.Element:
        paragraph = ET.Element(Q(W, "p"))
        if style is not None:
            properties = ET.SubElement(paragraph, Q(W, "pPr"))
            style_id = self._register_style(self._paragraph_styles, style)
            ET.SubElement(properties, Q(W, "pStyle"), {Q(W, "val"): style_id})
        return paragraph

    def _insert_body_paragraph(self, element: ET.Element, text: str) -> DocumentParagraph:
        self._insert_block(element)
        paragraph = DocumentParagraph(element, self)
        if self._block_container is self._body:
            self._paragraphs = [DocumentParagraph(p, self) for p in self._body.findall(Q(W, "p"))]
        self._current_paragraph = paragraph
        self._inline_container = None
        self._inline_index = None
        self._pending_text_offset = None
        if text:
            self.append_run(text, paragraph=paragraph)
        self._capture_cursor()
        return paragraph

    def _add_list_item(
        self, document_list: DocumentList, text: str, level: int, style: str | None
    ) -> DocumentParagraph:
        text = _require_text(text, "list_item_text")
        if isinstance(level, bool) or not isinstance(level, int) or not 0 <= level < _LIST_LEVELS:
            raise ValueError("list_level_invalid")
        element = self._new_paragraph_element(_require_style(style, "paragraph_style"))
        properties = element.find(Q(W, "pPr"))
        if properties is None:
            properties = ET.SubElement(element, Q(W, "pPr"))
        # numPr follows pStyle in the CT_PPr sequence.
        numbering = ET.SubElement(properties, Q(W, "numPr"))
        ET.SubElement(numbering, Q(W, "ilvl"), {Q(W, "val"): str(level)})
        ET.SubElement(numbering, Q(W, "numId"), {Q(W, "val"): str(document_list.number_id)})
        return self._insert_body_paragraph(element, text)

    def _header_footer(
        self, section: DocumentSection, story: str, kind: str
    ) -> DocumentHeaderFooter:
        self._target_element(section)
        if not isinstance(kind, str) or kind not in _HEADER_FOOTER_KINDS:
            raise ValueError("header_footer_kind_invalid")
        reference_tag = Q(W, f"{story}Reference")
        for reference in section.element.findall(reference_tag):
            if reference.get(Q(W, "type")) == kind:
                return self._header_footers[reference.get(Q(R, "id"), "")]
        root_tag = Q(W, "hdr" if story == "header" else "ftr")
        index = 1 + sum(part.element.tag == root_tag for part in self._header_footers.values())
        part = DocumentHeaderFooter(
            ET.Element(root_tag),
            kind,
            self._new_relationship_id(f"rId{story.title()}{index}"),
            self._new_part_name(f"word/{story}{index}.xml"),
            self,
        )
        self._header_footers[part.relationship_id] = part
        _insert_sect_pr_child(
            section.element,
            ET.Element(reference_tag, {Q(W, "type"): kind, Q(R, "id"): part.relationship_id}),
        )
        if kind == "first" and section.element.find(Q(W, "titlePg")) is None:
            _insert_sect_pr_child(section.element, ET.Element(Q(W, "titlePg")))
        if kind == "even":
            self._even_and_odd_headers = True
        return part

    def _story_roots(self) -> tuple[ET.Element, ...]:
        roots = [self._body]
        for part in self._header_footers.values():
            if all(root is not part.element for root in roots):
                roots.append(part.element)
        return tuple(roots)

    def _parents(self) -> dict[ET.Element, ET.Element]:
        return {
            child: parent
            for root in self._story_roots()
            for parent in root.iter()
            for child in parent
        }

    def _target_element(self, target: Any) -> ET.Element:
        if isinstance(target, Node):
            if target._element is None:
                raise NodeEditError("builder_cursor_node_invalid")
            path: list[int] = []
            current = target
            while current.parent is not None and current.parent._element is not None:
                parent = current.parent
                try:
                    path.append(list(parent._element).index(current._element))
                except ValueError:
                    raise NodeEditError("builder_cursor_node_stale") from None
                current = parent
            if current._element is None or serialize_xml_with_preserved_qnames(
                current._element
            ) != self._source_node_snapshots.get(target.part or ""):
                raise NodeEditError("builder_cursor_node_stale")
            element = self._source_paths.get(target.part or "", {}).get(tuple(reversed(path)))
            if element is None or target._element is None or element.tag != target._element.tag:
                raise NodeEditError("builder_cursor_node_invalid")
        elif isinstance(target, ET.Element):
            element = target
        else:
            owner = getattr(target, "_builder", None)
            if isinstance(target, DocumentTableCell):
                owner = target.table._builder
            elif isinstance(target, DocumentRun):
                owner = target.paragraph._builder
            if owner is not self:
                raise NodeEditError("builder_cursor_owner_invalid")
            element = getattr(target, "element", None)
        if element is None or (
            element is not self._document
            and not any(item is element for root in self._story_roots() for item in root.iter())
        ):
            raise NodeEditError("builder_cursor_target_stale")
        return element

    @staticmethod
    def _block_end(container: ET.Element) -> int:
        section = container.find(Q(W, "sectPr"))
        return list(container).index(section) if section is not None else len(container)

    def _capture_cursor(self, *, block: bool = False) -> None:
        if block:
            self._block_snapshot = tuple(self._block_container)
        self._inline_snapshot = (
            tuple(self._inline_container) if self._inline_container is not None else None
        )
        self._hyperlink_parent_snapshot = None
        if self._inline_container is not None and self._inline_container.tag == Q(W, "hyperlink"):
            parent = self._parents().get(self._inline_container)
            if parent is not None and parent.tag == Q(W, "p"):
                self._hyperlink_parent_snapshot = (parent, tuple(parent))
        self._pending_snapshot = (
            serialize_xml_with_preserved_qnames(self._pending_text_offset[0])
            if self._pending_text_offset is not None
            else None
        )

    def _validate_inline_cursor(
        self, paragraph: DocumentParagraph, tag: str, *, use_cursor: bool
    ) -> None:
        if not use_cursor:
            return
        self._target_element(paragraph)
        container = self._inline_container
        if container is not None:
            if not any(item is container for item in paragraph.element.iter()) or (
                self._inline_snapshot is not None and tuple(container) != self._inline_snapshot
            ):
                raise NodeEditError("builder_cursor_target_stale")
            if self._hyperlink_parent_snapshot is not None:
                parent, snapshot = self._hyperlink_parent_snapshot
                if self._parents().get(container) is not parent or tuple(parent) != snapshot:
                    raise NodeEditError("builder_cursor_target_stale")
            if tag == Q(W, "hyperlink"):
                parents = self._parents()
                ancestor = container
                while ancestor is not paragraph.element:
                    parent = parents.get(ancestor)
                    if ancestor.tag == Q(W, "hyperlink") and (
                        ancestor is not container or parent is not paragraph.element
                    ):
                        raise NodeEditError("builder_cursor_hyperlink_nested_unsupported")
                    if parent is None:
                        raise NodeEditError("builder_cursor_target_stale")
                    ancestor = parent
        if self._pending_text_offset is not None:
            target, _ = self._pending_text_offset
            if (
                not any(item is target for item in paragraph.element.iter())
                or serialize_xml_with_preserved_qnames(target) != self._pending_snapshot
            ):
                raise NodeEditError("builder_cursor_target_stale")

    def _insert_block(self, element: ET.Element) -> None:
        container = self._block_container
        if (
            container is not self._body
            and container not in self._parents()
            and not any(container is root for root in self._story_roots())
        ):
            raise NodeEditError("builder_cursor_target_stale")
        if (
            self._block_index is not None
            and self._block_snapshot is not None
            and tuple(container) != self._block_snapshot
        ):
            raise NodeEditError("builder_cursor_target_stale")
        index = self._block_end(container) if self._block_index is None else self._block_index
        container.insert(index, element)
        if self._block_index is not None:
            self._block_index = index + 1
        if container.tag == Q(W, "tc") and container[-1].tag != Q(W, "p"):
            ET.SubElement(container, Q(W, "p"))
        self._block_snapshot = tuple(container)

    def _insert_inline(
        self, paragraph: DocumentParagraph, element: ET.Element, *, use_cursor: bool
    ) -> None:
        self._validate_inline_cursor(paragraph, element.tag, use_cursor=use_cursor)
        if use_cursor and self._pending_text_offset is not None:
            target, offset = self._pending_text_offset
            if target is not paragraph.element and not any(
                item is target for item in paragraph.element.iter()
            ):
                raise NodeEditError("builder_cursor_target_stale")
            self._place_text_cursor(target, offset, split=True)
        if use_cursor and self._inline_container is not None:
            container = self._inline_container
            if container is not paragraph.element and not any(
                item is container for item in paragraph.element.iter()
            ):
                raise NodeEditError("builder_cursor_target_stale")
            index = self._inline_index if self._inline_index is not None else len(container)
            if container.tag == Q(W, "hyperlink") and element.tag == Q(W, "hyperlink"):
                # Validation above permits only a direct paragraph hyperlink.
                # Partition its children without copying any content-bearing node.
                link_index = list(paragraph.element).index(container)
                if index == 0:
                    index = link_index
                elif index == len(container):
                    index = link_index + 1
                else:
                    right = copy.copy(container)
                    right[:] = list(container)[index:]
                    container[:] = list(container)[:index]
                    container.tail = None
                    paragraph.element.insert(link_index + 1, right)
                    index = link_index + 1
                container = paragraph.element
                self._inline_container = container
            container.insert(index, element)
            self._inline_index = index + 1
        else:
            paragraph.element.append(element)
            self._inline_container = None
            self._inline_index = None
            self._pending_text_offset = None
        self._capture_cursor()

    def _place_text_cursor(self, target: ET.Element, offset: int, *, split: bool = False) -> None:
        if isinstance(offset, bool) or not isinstance(offset, int) or offset < 0:
            raise NodeEditError("builder_cursor_offset_invalid")
        self._pending_text_offset = None
        parents = self._parents()
        paragraph = target
        while paragraph.tag != Q(W, "p"):
            parent = parents.get(paragraph)
            if parent is None:
                raise NodeEditError("builder_cursor_target_stale")
            paragraph = parent
        runs = []
        for run in paragraph.iter(Q(W, "r")):
            ancestor = parents.get(run)
            while (
                ancestor is not None
                and ancestor is not paragraph
                and ancestor.tag != Q(W, "txbxContent")
            ):
                ancestor = parents.get(ancestor)
            if ancestor is paragraph:
                runs.append(run)
        if target.tag == Q(W, "r"):
            runs = [target]
            if not any(node.text for node in target.findall(Q(W, "t"))):
                if offset != 0:
                    raise NodeEditError("builder_cursor_offset_invalid")
                self._inline_container = parents[target]
                self._inline_index = list(parents[target]).index(target)
                return
        total = sum(len(node.text or "") for run in runs for node in run.findall(Q(W, "t")))
        if offset > total:
            raise NodeEditError("builder_cursor_offset_invalid")
        remaining = offset
        for run in runs:
            text = "".join(node.text or "" for node in run.findall(Q(W, "t")))
            if remaining > len(text) or not text:
                remaining -= len(text)
                continue
            container = parents[run]
            if container.tag not in {
                Q(W, tag)
                for tag in (
                    "p",
                    "hyperlink",
                    "fldSimple",
                    "sdtContent",
                    "smartTag",
                    "customXml",
                    "ins",
                    "moveTo",
                    "dir",
                    "bdo",
                )
            }:
                raise NodeEditError("builder_cursor_container_unsupported")
            index = list(container).index(run)
            if remaining in {0, len(text)}:
                self._inline_container = container
                self._inline_index = index + (remaining == len(text))
                return
            if not split:
                self._inline_container = container
                self._inline_index = index
                self._pending_text_offset = (target, offset)
                return
            # Copy only the run shell and its properties. Every other child is
            # assigned to exactly one half; only the straddling w:t is split.
            left, right = copy.copy(run), copy.copy(run)
            left[:] = []
            right[:] = []
            left.tail = None
            right.text = None
            text_offset = remaining
            for child in run:
                if child.tag == Q(W, "rPr"):
                    left.append(child)
                    right.append(copy.deepcopy(child))
                elif text_offset == 0:
                    right.append(child)
                elif child.tag != Q(W, "t"):
                    left.append(child)
                else:
                    value = child.text or ""
                    if text_offset >= len(value):
                        left.append(child)
                        text_offset -= len(value)
                    else:
                        before, after = copy.deepcopy(child), copy.deepcopy(child)
                        before.text, after.text = value[:text_offset], value[text_offset:]
                        before.tail = None
                        for node in (before, after):
                            if _needs_space_preservation(node.text or ""):
                                node.set(_XML_SPACE, "preserve")
                        left.append(before)
                        right.append(after)
                        text_offset = 0
            container.remove(run)
            container.insert(index, left)
            container.insert(index + 1, right)
            self._inline_container = container
            self._inline_index = index + 1
            return
        self._inline_container = paragraph
        self._inline_index = len(paragraph)

    def _move_to_marker(
        self, element: ET.Element, position: str, *, bookmark: bool
    ) -> DocumentBuilder:
        if position not in {"start", "end", "before", "after"}:
            raise NodeEditError("builder_cursor_position_invalid")
        parents = self._parents()
        root = next(
            (root for root in self._story_roots() if any(item is element for item in root.iter())),
            None,
        )
        if root is None:
            raise NodeEditError("builder_cursor_target_stale")
        if bookmark:
            identifier = element.get(Q(W, "id"))
            if identifier is None:
                raise NodeEditError("builder_bookmark_range_invalid")
            markers = [
                item
                for item in root.iter(Q(W, "bookmarkEnd"))
                if item.get(Q(W, "id")) == identifier
            ]
            if len(markers) != 1:
                raise NodeEditError("builder_bookmark_range_invalid")
            starts = [
                item
                for item in root.iter(Q(W, "bookmarkStart"))
                if item.get(Q(W, "id")) == identifier
            ]
            items = list(root.iter())
            if len(starts) != 1 or items.index(element) >= items.index(markers[0]):
                raise NodeEditError("builder_bookmark_range_invalid")
            marker = element if position in {"start", "before"} else markers[0]
            after = position in {"start", "after"}
        elif element.tag == Q(W, "fldSimple"):
            container = parents.get(element)
            paragraph = container
            while paragraph is not None and paragraph.tag != Q(W, "p"):
                paragraph = parents.get(paragraph)
            if paragraph is None or container is None:
                raise NodeEditError("builder_field_container_unsupported")
            if container.tag not in {
                Q(W, tag)
                for tag in (
                    "p",
                    "hyperlink",
                    "fldSimple",
                    "sdtContent",
                    "smartTag",
                    "customXml",
                    "ins",
                    "moveTo",
                    "dir",
                    "bdo",
                )
            }:
                raise NodeEditError("builder_field_container_unsupported")
            self.move_to(DocumentParagraph(paragraph, self))
            if position in {"start", "end"}:
                self._inline_container = element
                self._inline_index = 0 if position == "start" else len(element)
            else:
                self._inline_container = container
                self._inline_index = list(container).index(element) + (position == "after")
            self._capture_cursor(block=True)
            return self
        elif element.tag == Q(W, "fldChar") and element.get(Q(W, "fldCharType")) == "begin":
            fields_module = feature_module("fields")

            try:
                fields = fields_module.field_nodes(root, W, allow_instruction_nesting=True)
            except fields_module.FieldStructureError as error:
                raise NodeEditError(error.code) from None
            field = next((field for field in fields if field.begin is element), None)
            if field is None:
                raise NodeEditError("builder_field_target_invalid")
            separator, end = field.separate, field.end
            if end is None or (separator is None and position in {"start", "end"}):
                raise NodeEditError("builder_field_result_missing")
            marker = element if position == "before" else separator if position == "start" else end
            assert marker is not None
            after = position in {"start", "after"}
            marker = parents.get(marker)
            if marker is None or marker.tag != Q(W, "r"):
                raise NodeEditError("builder_field_container_unsupported")
            if len(marker.findall(Q(W, "fldChar"))) != 1 or any(
                child.tag not in {Q(W, "rPr"), Q(W, "fldChar")} for child in marker
            ):
                raise NodeEditError("builder_field_marker_run_unsupported")
        else:
            raise NodeEditError("builder_field_target_invalid")
        container = parents.get(marker)
        paragraph = container
        while paragraph is not None and paragraph.tag != Q(W, "p"):
            paragraph = parents.get(paragraph)
        if paragraph is None or container is None:
            raise NodeEditError("builder_cursor_container_unsupported")
        if container.tag not in {
            Q(W, tag)
            for tag in (
                "p",
                "hyperlink",
                "fldSimple",
                "sdtContent",
                "smartTag",
                "customXml",
                "ins",
                "moveTo",
                "dir",
                "bdo",
            )
        }:
            raise NodeEditError("builder_cursor_container_unsupported")
        self.move_to(DocumentParagraph(paragraph, self))
        self._inline_container = container
        self._inline_index = list(container).index(marker) + after
        self._capture_cursor(block=True)
        return self

    def _new_relationship_id(self, preferred: str, *, part_name: str = "word/document.xml") -> str:
        if part_name == "word/document.xml":
            reserved = self._reserved_relationships
        else:
            if part_name not in self._story_relationships:
                self._story_relationships[part_name] = (
                    {
                        relation.id
                        for relation in self._session._graph.parts[part_name].relationships
                    }
                    if self._session is not None and part_name in self._session._graph.parts
                    else set()
                )
            reserved = self._story_relationships[part_name]
        value = preferred
        index = 1
        while value in reserved:
            value = f"{preferred}_{index}"
            index += 1
        reserved.add(value)
        return value

    def _new_part_name(self, preferred: str) -> str:
        value = preferred
        stem, extension = posixpath.splitext(preferred)
        index = 1
        while value in self._reserved_parts:
            value = f"{stem}_{index}{extension}"
            index += 1
        self._reserved_parts.add(value)
        return value

    @staticmethod
    def _zip_info(name: str) -> zipfile.ZipInfo:
        info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        info.create_system = 3
        info.external_attr = 0o600 << 16
        return info

    def _authored_parts(self) -> dict[str, bytes]:
        parts = {
            "[Content_Types].xml": self._content_types_xml(),
            "_rels/.rels": self._package_relationships_xml(),
            "word/document.xml": self._xml(self._document),
        }
        if self._core_properties:
            parts[CORE_PROPERTIES_PARTNAME] = self._core_properties_xml()
        if self._has_styles:
            parts["word/styles.xml"] = self._styles_xml()
        if self._has_settings:
            parts[_SETTINGS] = self._settings_xml()
        if self._lists:
            parts["word/numbering.xml"] = self._numbering_xml()
        for part in self._header_footers.values():
            parts[part.part_name] = self._header_footer_xml(part)
        for source in dict.fromkeys(
            [image.source_part for image in self._images]
            + [source for _, _, source in self._hyperlinks]
        ):
            if source != "word/document.xml":
                parts[self._relationships_part(source)] = self._story_relationships_xml(source)
        if (
            self._images
            or self._hyperlinks
            or self._has_styles
            or self._has_settings
            or self._lists
            or self._header_footers
        ):
            parts["word/_rels/document.xml.rels"] = self._document_relationships_xml()
        parts.update((image.part_name, image.data) for image in self._images)
        return parts

    def _sync_session(self) -> None:
        """Merge authored ancillary parts, then stage changed story trees only."""
        session = self._session
        for name, original in self._source_part_bytes.items():
            if name not in session.part_names or session.part_bytes(name) != original:
                raise NodeEditError("builder_source_story_stale")
        parts = self._authored_parts()
        routes: dict[str, str] = {}
        for rel_name in tuple(name for name in parts if name.endswith(".rels")):
            generated = parts.pop(rel_name, None)
            if generated is None:
                continue
            original = session.part_bytes(rel_name) if rel_name in session.part_names else None
            root = (
                parse_xml_with_preserved_qnames(original)
                if original is not None
                else ET.Element(Q(PKG_REL, "Relationships"))
            )
            before = serialize_xml_with_preserved_qnames(root)
            base = posixpath.dirname(posixpath.dirname(rel_name))
            for relation in ET.fromstring(generated):
                kind = relation.get("Type", "").rsplit("/", 1)[-1]
                generated_target = relation.get("Target", "")
                external = relation.get("TargetMode") == "External"
                generated_part = (
                    generated_target
                    if external
                    else generated_target.lstrip("/")
                    if generated_target.startswith("/")
                    else posixpath.normpath(posixpath.join(base, generated_target))
                )
                singleton = (
                    kind in {"officeDocument", "styles", "settings", "numbering"}
                    or relation.get("Type") == CORE_PROPERTIES_REL_TYPE
                )
                existing = (
                    next(
                        (
                            item
                            for item in root
                            if item.get("Type") == relation.get("Type")
                            and item.get("TargetMode") != "External"
                        ),
                        None,
                    )
                    if singleton
                    else next((item for item in root if item.get("Id") == relation.get("Id")), None)
                )
                if existing is not None:
                    target = existing.get("Target", "")
                    resolved = (
                        target
                        if external
                        else target.lstrip("/")
                        if target.startswith("/")
                        else posixpath.normpath(posixpath.join(base, target))
                    )
                    if singleton:
                        routes[generated_part] = resolved
                    elif (
                        existing.get("Type") != relation.get("Type")
                        or resolved != generated_part
                        or existing.get("TargetMode") != relation.get("TargetMode")
                    ):
                        raise NodeEditError("builder_relationship_conflict")
                    continue
                identifier = relation.get("Id", "")
                occupied = {item.get("Id") for item in root}
                index = 1
                while identifier in occupied:
                    identifier = f"{relation.get('Id')}_{index}"
                    index += 1
                relation.set("Id", identifier)
                root.append(relation)
            after = serialize_xml_with_preserved_qnames(root)
            if after != before:
                parts[rel_name] = after
        content_types = parts.pop("[Content_Types].xml")
        types = parse_xml_with_preserved_qnames(session.part_bytes("[Content_Types].xml"))
        before_types = serialize_xml_with_preserved_qnames(types)
        for declaration in ET.fromstring(content_types):
            if declaration.tag == Q(CONTENT_TYPES, "Override"):
                name = declaration.get("PartName", "").lstrip("/")
                declaration.set("PartName", "/" + routes.get(name, name))
                key = "PartName"
            else:
                key = "Extension"
            existing = next(
                (
                    item
                    for item in types
                    if item.tag == declaration.tag and item.get(key) == declaration.get(key)
                ),
                None,
            )
            if existing is None:
                types.append(declaration)
            elif existing.get("ContentType") != declaration.get("ContentType"):
                if key == "Extension":
                    for image in self._images:
                        if image.part_name.endswith("." + declaration.get("Extension", "")):
                            types.append(
                                ET.Element(
                                    Q(CONTENT_TYPES, "Override"),
                                    {
                                        "PartName": "/" + image.part_name,
                                        "ContentType": image.content_type,
                                    },
                                )
                            )
                else:
                    raise NodeEditError("builder_content_type_conflict")
        after_types = serialize_xml_with_preserved_qnames(types)
        if after_types != before_types:
            parts["[Content_Types].xml"] = after_types
        staged: dict[str, bytes] = {}
        for name, data in parts.items():
            if name in self._source_roots:
                root = self._source_roots[name]
                current = serialize_xml_with_preserved_qnames(root)
                if current != self._root_snapshots[name]:
                    staged[name] = current
                continue
            destination = routes.get(name, name)
            if name == _SETTINGS:
                if destination not in session.part_names or session.part_bytes(destination) != data:
                    staged[destination] = data
                continue
            if destination in session.part_names and name in {
                "word/styles.xml",
                "word/numbering.xml",
                CORE_PROPERTIES_PARTNAME,
            }:
                original = session.part_bytes(destination)
                root = parse_xml_with_preserved_qnames(original)
                before = serialize_xml_with_preserved_qnames(root)
                self._merge_authored_xml(root, parse_xml_with_preserved_qnames(data), name)
                current = serialize_xml_with_preserved_qnames(root)
                if current != before:
                    staged[destination] = current
            elif destination not in session.part_names or session.part_bytes(destination) != data:
                staged[destination] = data
        for name, data in staged.items():
            if name not in session._entries:
                session._entries[name] = self._zip_info(name)
            session._set_part(name, data)
            if name in self._source_part_bytes:
                self._source_part_bytes[name] = data

    def _merge_authored_xml(self, root: ET.Element, generated: ET.Element, name: str) -> None:
        if root.tag != generated.tag:
            raise NodeEditError("builder_ancillary_part_layout_unsupported")
        for child in generated:
            if name == "word/styles.xml" and child.tag == Q(W, "style"):
                identifier = child.get(Q(W, "styleId"))
                existing = next(
                    (
                        item
                        for item in root.findall(Q(W, "style"))
                        if item.get(Q(W, "styleId")) == identifier
                    ),
                    None,
                )
                if existing is not None:
                    if existing.get(Q(W, "type")) != child.get(Q(W, "type")):
                        raise NodeEditError("builder_style_type_conflict")
                    continue
            elif name == "word/numbering.xml":
                attribute = Q(W, "numId" if child.tag == Q(W, "num") else "abstractNumId")
                if any(
                    item.get(attribute) == child.get(attribute) for item in root.findall(child.tag)
                ):
                    continue
            elif child.tag == Q(W, "docDefaults"):
                existing = root.find(child.tag)
                if existing is not None:
                    replacement = child.find(Q(W, "rPrDefault"))
                    old = existing.find(Q(W, "rPrDefault"))
                    if old is None:
                        if replacement is not None:
                            existing.insert(0, replacement)
                    elif replacement is not None:
                        old_properties = old.find(Q(W, "rPr"))
                        if old_properties is None:
                            old_properties = ET.SubElement(old, Q(W, "rPr"))
                        new_properties = replacement.find(Q(W, "rPr"))
                        if new_properties is not None:
                            for value in new_properties:
                                for previous in list(old_properties.findall(value.tag)):
                                    old_properties.remove(previous)
                                old_properties.append(value)
                            old_properties[:] = sorted(
                                old_properties,
                                key=lambda item: _RPR_RANK.get(item.tag.rsplit("}", 1)[-1], 999),
                            )
                    continue
            else:
                for old in list(root.findall(child.tag)):
                    root.remove(old)
            root.append(child)
        if name == "word/numbering.xml":
            root[:] = sorted(root, key=lambda item: 1 if item.tag == Q(W, "num") else 0)
        elif name == "word/styles.xml":
            root[:] = sorted(
                root,
                key=lambda item: {"docDefaults": 0, "latentStyles": 1, "style": 2}.get(
                    item.tag.rsplit("}", 1)[-1], 999
                ),
            )

    def _content_types_xml(self) -> bytes:
        root = ET.Element(Q(CONTENT_TYPES, "Types"))
        ET.SubElement(
            root,
            Q(CONTENT_TYPES, "Default"),
            {
                "Extension": "rels",
                "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
            },
        )
        ET.SubElement(
            root,
            Q(CONTENT_TYPES, "Default"),
            {"Extension": "xml", "ContentType": "application/xml"},
        )
        media_extensions = (
            feature_module("image_formats").MEDIA_EXTENSIONS.items() if self._images else ()
        )
        for content_type, extension in media_extensions:
            if any(image.content_type == content_type for image in self._images):
                ET.SubElement(
                    root,
                    Q(CONTENT_TYPES, "Default"),
                    {"Extension": extension, "ContentType": content_type},
                )
        ET.SubElement(
            root,
            Q(CONTENT_TYPES, "Override"),
            {"PartName": "/word/document.xml", "ContentType": _DOCUMENT_CONTENT_TYPE},
        )
        if self._has_styles:
            ET.SubElement(
                root,
                Q(CONTENT_TYPES, "Override"),
                {"PartName": "/word/styles.xml", "ContentType": _STYLES_CONTENT_TYPE},
            )
        if self._has_settings:
            ET.SubElement(
                root,
                Q(CONTENT_TYPES, "Override"),
                {"PartName": f"/{_SETTINGS}", "ContentType": _SETTINGS_CONTENT_TYPE},
            )
        if self._lists:
            ET.SubElement(
                root,
                Q(CONTENT_TYPES, "Override"),
                {
                    "PartName": "/word/numbering.xml",
                    "ContentType": _NUMBERING_CONTENT_TYPE,
                },
            )
        for part in self._header_footers.values():
            ET.SubElement(
                root,
                Q(CONTENT_TYPES, "Override"),
                {
                    "PartName": f"/{part.part_name}",
                    "ContentType": _HEADER_CONTENT_TYPE
                    if part.element.tag == Q(W, "hdr")
                    else _FOOTER_CONTENT_TYPE,
                },
            )
        if self._core_properties:
            ET.SubElement(
                root,
                Q(CONTENT_TYPES, "Override"),
                {
                    "PartName": f"/{CORE_PROPERTIES_PARTNAME}",
                    "ContentType": CORE_PROPERTIES_CONTENT_TYPE,
                },
            )
        return self._xml(root)

    def _package_relationships_xml(self) -> bytes:
        root = ET.Element(Q(PKG_REL, "Relationships"))
        ET.SubElement(
            root,
            Q(PKG_REL, "Relationship"),
            {
                "Id": "rId1",
                "Type": f"{_REL}/officeDocument",
                "Target": "word/document.xml",
            },
        )
        if self._core_properties:
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {
                    "Id": "rId2",
                    "Type": CORE_PROPERTIES_REL_TYPE,
                    "Target": CORE_PROPERTIES_PARTNAME,
                },
            )
        return self._xml(root)

    def _document_relationships_xml(self) -> bytes:
        root = ET.Element(Q(PKG_REL, "Relationships"))
        if self._has_styles:
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {"Id": "rIdStyles", "Type": f"{_REL}/styles", "Target": "styles.xml"},
            )
        if self._has_settings:
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {
                    "Id": "rIdSettings",
                    "Type": f"{_REL}/settings",
                    "Target": _SETTINGS.removeprefix("word/"),
                },
            )
        if self._lists:
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {
                    "Id": "rIdNumbering",
                    "Type": f"{_REL}/numbering",
                    "Target": "numbering.xml",
                },
            )
        for part in self._header_footers.values():
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {
                    "Id": part.relationship_id,
                    "Type": f"{_REL}/{'header' if part.element.tag == Q(W, 'hdr') else 'footer'}",
                    "Target": posixpath.relpath(part.part_name, "word"),
                },
            )
        self._append_story_relationships(root, "word/document.xml")
        return self._xml(root)

    @staticmethod
    def _relationships_part(source: str) -> str:
        directory, name = posixpath.split(source)
        return posixpath.join(directory, "_rels", name + ".rels")

    def _story_relationships_xml(self, source: str) -> bytes:
        root = ET.Element(Q(PKG_REL, "Relationships"))
        self._append_story_relationships(root, source)
        return self._xml(root)

    def _append_story_relationships(self, root: ET.Element, source: str) -> None:
        for image in self._images:
            if image.source_part != source:
                continue
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {
                    "Id": image.relationship_id,
                    "Type": _IMAGE_RELATIONSHIP,
                    "Target": posixpath.relpath(image.part_name, posixpath.dirname(source)),
                },
            )
        for relationship_id, url, part in self._hyperlinks:
            if part != source:
                continue
            ET.SubElement(
                root,
                Q(PKG_REL, "Relationship"),
                {
                    "Id": relationship_id,
                    "Type": _HYPERLINK_RELATIONSHIP,
                    "Target": url,
                    "TargetMode": "External",
                },
            )

    def _styles_xml(self) -> bytes:
        root = ET.Element(Q(W, "styles"))
        if self._default_run_font is not None:
            properties = ET.SubElement(
                ET.SubElement(ET.SubElement(root, Q(W, "docDefaults")), Q(W, "rPrDefault")),
                Q(W, "rPr"),
            )
            name = self._default_run_font
            ET.SubElement(
                properties,
                Q(W, "rFonts"),
                {Q(W, "ascii"): name, Q(W, "hAnsi"): name, Q(W, "cs"): name},
            )
            if self._default_run_size is not None:
                size = str(self._default_run_size)
                ET.SubElement(properties, Q(W, "sz"), {Q(W, "val"): size})
                ET.SubElement(properties, Q(W, "szCs"), {Q(W, "val"): size})
        for style_type, style_ids in (
            ("paragraph", self._paragraph_styles),
            ("character", self._character_styles),
            ("table", self._table_styles),
        ):
            for style_id in sorted(style_ids):
                style = ET.SubElement(
                    root,
                    Q(W, "style"),
                    {Q(W, "type"): style_type, Q(W, "styleId"): style_id},
                )
                name = self._style_names.get(style_id, style_id)
                ET.SubElement(style, Q(W, "name"), {Q(W, "val"): name})
                if style_type == "character" and style_id == _HYPERLINK_STYLE:
                    _append_hyperlink_style_properties(style)
        return self._xml(root)

    def _settings_xml(self) -> bytes:
        data = self._xml(ET.Element(Q(W, "settings")))
        if self._session is not None:
            relation = next(
                (
                    item
                    for item in self._session._graph.parts["word/document.xml"].relationships
                    if item.type.endswith("/settings") and not item.is_external
                ),
                None,
            )
            if relation is not None:
                data = self._session.part_bytes(relation.resolved_target())
        document_settings = feature_module("document_settings")
        if self._even_and_odd_headers:
            data = document_settings._patch_child(
                data,
                Q(W, "evenAndOddHeaders"),
                ET.tostring(ET.Element(Q(W, "evenAndOddHeaders")), encoding="utf-8"),
                _SETTINGS_CHILD_ORDER_MAP,
            )
        if self._compat_mode is not None:
            data = document_settings._compat_patch(
                data, {"compatibilityMode": str(self._compat_mode)}
            )
        for name, payload in self._settings_values.items():
            tag = Q(W, "writeProtection" if name == "write_protection" else name)
            data = document_settings._patch_child(data, tag, payload, _SETTINGS_CHILD_ORDER_MAP)
        if self._compatibility_options:
            data = document_settings._compat_patch(data, self._compatibility_options)
        return data

    def _numbering_xml(self) -> bytes:
        # CT_Numbering lists every w:abstractNum before any w:num.
        root = ET.Element(Q(W, "numbering"))
        for abstract_id, kind in enumerate(self._lists):
            abstract = ET.SubElement(
                root,
                Q(W, "abstractNum"),
                {Q(W, "abstractNumId"): str(self._abstract_id_base + abstract_id)},
            )
            ET.SubElement(abstract, Q(W, "multiLevelType"), {Q(W, "val"): "multilevel"})
            for level in range(_LIST_LEVELS):
                if kind == "numbered":
                    number_format = _NUMBERED_LEVEL_FORMATS[level % len(_NUMBERED_LEVEL_FORMATS)]
                    level_text, font = f"%{level + 1}.", None
                else:
                    number_format = "bullet"
                    level_text, font = _BULLET_LEVEL_GLYPHS[level % len(_BULLET_LEVEL_GLYPHS)]
                node = ET.SubElement(abstract, Q(W, "lvl"), {Q(W, "ilvl"): str(level)})
                ET.SubElement(node, Q(W, "start"), {Q(W, "val"): "1"})
                ET.SubElement(node, Q(W, "numFmt"), {Q(W, "val"): number_format})
                ET.SubElement(node, Q(W, "lvlText"), {Q(W, "val"): level_text})
                ET.SubElement(node, Q(W, "lvlJc"), {Q(W, "val"): "left"})
                indent = ET.SubElement(ET.SubElement(node, Q(W, "pPr")), Q(W, "ind"))
                indent.set(Q(W, "left"), str(_LIST_INDENT_STEP * (level + 1)))
                indent.set(Q(W, "hanging"), str(_LIST_HANGING_INDENT))
                if font is not None:
                    ET.SubElement(
                        ET.SubElement(node, Q(W, "rPr")),
                        Q(W, "rFonts"),
                        {
                            Q(W, "ascii"): font,
                            Q(W, "hAnsi"): font,
                            Q(W, "hint"): "default",
                        },
                    )
        for abstract_id in range(len(self._lists)):
            number = ET.SubElement(
                root, Q(W, "num"), {Q(W, "numId"): str(self._number_id_base + abstract_id + 1)}
            )
            ET.SubElement(
                number,
                Q(W, "abstractNumId"),
                {Q(W, "val"): str(self._abstract_id_base + abstract_id)},
            )
        return self._xml(root)

    def _header_footer_xml(self, part: DocumentHeaderFooter) -> bytes:
        # CT_HdrFtr requires at least one block-level element.
        root = ET.Element(part.element.tag, part.element.attrib)
        root.extend(part.element)
        if not len(root):
            ET.SubElement(root, Q(W, "p"))
        return self._xml(root)

    def _core_properties_xml(self) -> bytes:
        root = ET.Element(Q(CP_NS, "coreProperties"))
        for name, namespace, local_name in _CORE_PROPERTY_ELEMENTS:
            value = self._core_properties.get(name)
            if value is not None:
                element = ET.SubElement(root, Q(namespace, local_name))
                element.text = value
                if name in _W3CDTF_CORE_PROPERTIES:
                    element.set(Q(XSI_NS, "type"), "dcterms:W3CDTF")
        return self._xml(root)

    @staticmethod
    def _xml(root: ET.Element) -> bytes:
        return serialize_xml_with_preserved_qnames(_declare_part_prefixes(root))


def _append_hyperlink_style_properties(style: ET.Element) -> None:
    """Give Word's built-in ``Hyperlink`` character style its standard appearance."""
    ET.SubElement(style, Q(W, "uiPriority"), {Q(W, "val"): "99"})
    ET.SubElement(style, Q(W, "unhideWhenUsed"))
    properties = ET.SubElement(style, Q(W, "rPr"))
    ET.SubElement(properties, Q(W, "color"), {Q(W, "val"): "0563C1"})
    ET.SubElement(properties, Q(W, "u"), {Q(W, "val"): "single"})


def _declare_part_prefixes(root: ET.Element) -> ET.Element:
    """Return a shallow root copy declaring conventional prefixes for used namespaces.

    The serializer never consults the global ElementTree registry, so without these
    declarations new names would be written under generated ``ns0``-style prefixes.
    The public builder tree itself is left unmodified.
    """
    used: set[str] = set()
    for element in root.iter():
        for name in (element.tag, *element.keys()):
            if isinstance(name, str) and name.startswith("{"):
                used.add(name[1:].split("}", 1)[0])
    declarations = {
        f"xmlns:{prefix}": namespace
        for prefix, namespace in (*_PART_PREFIXES, *_CORE_PART_PREFIXES)
        if namespace in used
    }
    if not declarations:
        return root
    declared = ET.Element(root.tag, {**declarations, **root.attrib})
    declared.extend(root)
    return declared
