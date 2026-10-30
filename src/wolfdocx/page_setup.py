"""Source-bound DOCX section page-setup authoring and discovery.

Provides verified, source-bound authoring for selected-section page dimensions
(width, height), portrait/landscape orientation, page margins (top, bottom,
left, right, header, footer, gutter), equal or explicit-width columns with optional
separators, and section document grids (type, linePitch, charSpace).

Columns and document grid:
    Page setup authoring supports bounded equal and explicit-width columns
    (up to 45 columns). ``ColumnsSpec.widths`` contains one positive twip width
    per column; optional ``spaces`` contains one gap per column, with the last
    ignored for usable-width accounting. ``separator=None`` preserves w:sep.
    Column configurations whose occupied width exceeds the usable page width
    are refused as ``page_setup_columns_exceed_usable_width``.
    ``document_grid`` accepts type, line_pitch (twips), and char_space (signed
    1/4096-character units); omitted attributes are preserved.

Odd and even page headers:
    ``different_odd_and_even_pages`` maps to the document-wide
    ``w:evenAndOddHeaders`` toggle in the settings part related from the main
    document, so it applies to every section. Enabling it inserts
    ``<w:evenAndOddHeaders/>`` in settings schema order, creating the settings
    part when the package has none; disabling it removes the toggle. Every other
    settings child is preserved in place.

Orientation:
    Orientation names are matched case-insensitively and authored in the
    lowercase ``ST_PageOrientation`` form. Changing orientation without an
    explicit ``page_size`` swaps the inherited ``w:w`` and ``w:h`` so the
    emitted ``w:pgSz`` describes the rotated page (landscape: width >= height;
    portrait: height >= width). An explicit ``page_size`` that contradicts the
    requested orientation is refused rather than silently rotated.

Footnote numbering:
    ``footnote_properties`` authors the section's ``w:footnotePr``
    (ECMA-376 17.11.11) in ``CT_SectPr`` schema order, directly after the
    header and footer references. Only the requested ``w:numFmt``,
    ``w:numStart``, and ``w:numRestart`` children change; ``w:pos`` and any
    other child are preserved in place. ``clear`` removes the section's own
    ``w:footnotePr`` so the section falls back to the document-wide settings.
"""

from __future__ import annotations

import json
import re
import zipfile
from collections.abc import Mapping
from dataclasses import InitVar, dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .assembly import _new_rel_id, _rels_name
from .authoring import (
    _PUBLISHED,
    _REFUSED,
    _cell_structure_fingerprint,
    _require_sha256,
    _sha256_file,
    _unsupported_context_by_element,
)
from .inspect import (
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .protection import (
    _CT,
    _PR,
    _R,
    _SETTINGS,
    _SETTINGS_CHILD_ORDER_MAP,
    _SETTINGS_CONTENT_TYPE,
    _W,
    _insert_settings_child,
)
from .review import _DOCUMENT, ReviewReport, _publish, _word_namespace

__all__ = [
    "ColumnsSpec",
    "FootnotePropertiesSpec",
    "LineNumberingSpec",
    "MarginsSpec",
    "PageBorderSpec",
    "PageBordersSpec",
    "PageNumberingSpec",
    "PageSetupReport",
    "PageSetupRequest",
    "PageSetupTarget",
    "PageSizeSpec",
    "author_page_setup",
    "discover_page_setup_targets",
    "load_page_setup_request",
    "page_setup_structure_fingerprint",
    "parse_page_setup_request",
]

# Document-wide multiple-page layouts are CT_Settings flags, not section geometry.
# ECMA-376 Part 1 §17.15.1 (mirrorMargins, printTwoOnOne, bookFoldPrinting,
# bookFoldRevPrinting); the settings author owns edits to this mutually-exclusive group.
_MULTIPLE_PAGES = {
    "normal": None,
    "mirror_margins": "mirrorMargins",
    "two_pages_per_sheet": "printTwoOnOne",
    "book_fold": "bookFoldPrinting",
    "reverse_book_fold": "bookFoldRevPrinting",
}


def _multiple_pages_layout(settings_root: ET.Element, namespace: str) -> str:
    enabled = [
        layout
        for layout, tag in _MULTIPLE_PAGES.items()
        if tag is not None and _is_on_off_enabled(settings_root.find(Q(namespace, tag)), namespace)
    ]
    # Do not silently choose one layout when a source carries conflicting active flags.
    return enabled[0] if len(enabled) == 1 else ("normal" if not enabled else "conflicting")


# Meaningful coordinate bounds in twips (WordprocessingML ST_TwipsMeasure)
_MIN_PAGE_DIMENSION = 144  # 0.1 inch
_MAX_PAGE_DIMENSION = 31_680  # 22 inches (Word standard maximum page dimension)
_MAX_COLUMN_COUNT = 45  # Word standard maximum column count
_DEFAULT_COLUMN_SPACING = 720  # 0.5 inch standard column gap

# OOXML ECMA-376 17.6.17 CT_SectPr child schema ordering
_SECT_PR_ORDER: dict[str, int] = {
    "headerReference": 0,
    "footerReference": 0,
    "footnotePr": 1,
    "endnotePr": 2,
    "type": 3,
    "pgSz": 4,
    "pgMar": 5,
    "paperSrc": 6,
    "pgBorders": 7,
    "lnNumType": 8,
    "pgNumType": 9,
    "cols": 10,
    "formProt": 11,
    "vAlign": 12,
    "noEndnote": 13,
    "titlePg": 14,
    "textDirection": 15,
    "bidi": 16,
    "rtlGutter": 17,
    "docGrid": 18,
    "printerSettings": 19,
    "sectPrChange": 20,
}

_REQUIRED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "structure_fingerprint_sha256",
        "section",
    }
)
_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "structure_fingerprint_sha256",
        "section",
        "part",
        "page_size",
        "orientation",
        "margins",
        "columns",
        "page_numbering",
        "line_numbering",
        "different_first_page",
        "different_first_page_header_footer",
        "different_odd_and_even_pages",
        "vertical_alignment",
        "page_borders",
        "footnote_properties",
        "document_grid",
    }
)
_ALLOWED_PAGE_SIZE_FIELDS = frozenset({"width", "height", "code"})
_REQUIRED_PAGE_SIZE_FIELDS = frozenset({"width", "height"})
# python-docx Section.header_distance/footer_distance name the w:pgMar header/footer.
_MARGIN_ALIASES = (("header_distance", "header"), ("footer_distance", "footer"))
_ALLOWED_MARGINS_FIELDS = frozenset(
    {"top", "bottom", "left", "right", "header", "footer", "gutter"}
    | {alias for alias, _ in _MARGIN_ALIASES}
)
_ALLOWED_COLUMNS_FIELDS = frozenset({"count", "spacing", "separator", "widths", "spaces"})
_GRID_FIELDS = {"type": "type", "line_pitch": "linePitch", "char_space": "charSpace"}
_GRID_TYPES = frozenset({"default", "lines", "linesAndChars", "snapToChars"})
_ALLOWED_ORIENTATIONS = frozenset({"portrait", "landscape"})
_ALLOWED_PAGE_NUMBER_FORMATS = frozenset(
    {"decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"}
)
_ALLOWED_PAGE_NUMBERING_FIELDS = frozenset(
    {"start", "format", "fmt", "clear_start", "clear_format", "clear"}
)
_MAX_PAGE_NUMBER = 32_767
_ALLOWED_LINE_NUMBERING_FIELDS = frozenset(
    {"start", "interval", "count_by", "distance", "restart", "clear"}
)
_ALLOWED_LINE_NUMBER_RESTARTS = frozenset({"newPage", "newSection", "continuous"})
_MAX_LINE_NUMBER = 32_767
_ALLOWED_FOOTNOTE_PROPERTIES_FIELDS = frozenset({"format", "start", "restart", "clear"})
# Footnote number formats Word offers for a section (subset of ST_NumberFormat).
_ALLOWED_FOOTNOTE_NUMBER_FORMATS = frozenset(
    {"decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter", "chicago"}
)
_ALLOWED_FOOTNOTE_RESTARTS = frozenset({"continuous", "eachSect", "eachPage"})
_MAX_FOOTNOTE_NUMBER = 32_767
# ECMA-376 17.11.11 CT_FtnProps child order.
_FOOTNOTE_PR_ORDER: dict[str, int] = {"pos": 0, "numFmt": 1, "numStart": 2, "numRestart": 3}
# FootnotePropertiesSpec field -> CT_FtnProps child carrying its w:val.
_FOOTNOTE_PR_FIELDS = (("format", "numFmt"), ("start", "numStart"), ("restart", "numRestart"))

_ALLOWED_VERTICAL_ALIGNMENTS = frozenset({"top", "center", "both", "bottom"})
_PAGE_BORDER_SIDES = ("top", "left", "bottom", "right")
_ALLOWED_PAGE_BORDER_FIELDS = frozenset({"style", "width", "color"})
_VALID_PAGE_BORDER_STYLES = frozenset(
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
_HEX_COLOR_PATTERN = re.compile(r"^[0-9a-fA-F]{6}\Z")


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    """True when every part the request did not change is byte-identical to source."""
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


def _sect_pr_rank(elem: ET.Element) -> int:
    tag = elem.tag.rsplit("}", 1)[-1]
    return _SECT_PR_ORDER.get(tag, 99)


def _insert_sect_pr_child(sect_pr: ET.Element, new_child: ET.Element) -> None:
    """Insert a new child element into sect_pr maintaining OOXML schema order."""
    target_rank = _sect_pr_rank(new_child)
    inserted = False
    for i, child in enumerate(sect_pr):
        child_rank = _sect_pr_rank(child)
        if child_rank > target_rank:
            sect_pr.insert(i, new_child)
            inserted = True
            break
    if not inserted:
        sect_pr.append(new_child)


def _footnote_pr_rank(elem: ET.Element) -> int:
    return _FOOTNOTE_PR_ORDER.get(elem.tag.rsplit("}", 1)[-1], 99)


def _insert_footnote_pr_child(footnote_pr: ET.Element, new_child: ET.Element) -> None:
    """Insert a new w:footnotePr child in CT_FtnProps order; unknown children stay last."""
    target_rank = _footnote_pr_rank(new_child)
    for i, child in enumerate(footnote_pr):
        if _footnote_pr_rank(child) > target_rank:
            footnote_pr.insert(i, new_child)
            return
    footnote_pr.append(new_child)


def _extract_footnote_properties_info(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract the section's own footnote numbering (w:footnotePr) without normalization."""
    elements = sect_pr.findall(Q(ns, "footnotePr"))
    if not elements:
        return {"footnote_properties": None, "has_duplicate_footnote_pr": False}
    element = elements[0]

    def value(name: str) -> str | None:
        child = element.find(Q(ns, name))
        if child is None:
            return None
        return child.get(Q(ns, "val")) or child.get("val")

    raw_start = value("numStart")
    return {
        "footnote_properties": {
            "format": value("numFmt"),
            "start": int(raw_start) if raw_start is not None and raw_start.isdigit() else None,
            "restart": value("numRestart"),
            "position": value("pos"),
        },
        "has_duplicate_footnote_pr": len(elements) > 1
        or any(len(element.findall(Q(ns, name))) > 1 for name in _FOOTNOTE_PR_ORDER),
    }


def _find_sections(root: ET.Element, ns: str) -> list[ET.Element]:
    """Find all section property elements in the document body in order."""
    body = root.find(Q(ns, "body"))
    if body is None:
        return []
    sections: list[ET.Element] = []
    for child in body:
        if child.tag == Q(ns, "p"):
            ppr = child.find(Q(ns, "pPr"))
            if ppr is not None:
                sect_pr = ppr.find(Q(ns, "sectPr"))
                if sect_pr is not None:
                    sections.append(sect_pr)
    final = body.find(Q(ns, "sectPr"))
    if final is not None:
        sections.append(final)
    return sections


def _extract_page_geometry(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract explicit page dimensions, orientation, and margins from a section."""
    pg_sz_elems = sect_pr.findall(Q(ns, "pgSz"))
    pg_mar_elems = sect_pr.findall(Q(ns, "pgMar"))
    has_duplicate_geometry = len(pg_sz_elems) > 1 or len(pg_mar_elems) > 1

    pg_sz = pg_sz_elems[0] if pg_sz_elems else None
    pg_mar = pg_mar_elems[0] if pg_mar_elems else None

    def _int_attr(elem: ET.Element | None, attr: str) -> int | None:
        if elem is None:
            return None
        raw = elem.get(Q(ns, attr)) or elem.get(attr)
        if raw is not None:
            try:
                return int(raw)
            except ValueError:
                return None
        return None

    width = _int_attr(pg_sz, "w")
    height = _int_attr(pg_sz, "h")
    orient = (
        pg_sz.get(Q(ns, "orient"))
        if pg_sz is not None and pg_sz.get(Q(ns, "orient"))
        else "portrait"
    )
    code = pg_sz.get(Q(ns, "code")) if pg_sz is not None and pg_sz.get(Q(ns, "code")) else None

    top = _int_attr(pg_mar, "top")
    bottom = _int_attr(pg_mar, "bottom")
    left = _int_attr(pg_mar, "left")
    right = _int_attr(pg_mar, "right")
    header = _int_attr(pg_mar, "header")
    footer = _int_attr(pg_mar, "footer")
    gutter = _int_attr(pg_mar, "gutter")

    return {
        "width": width,
        "height": height,
        "orientation": orient,
        "code": code,
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right,
        "header_margin": header,
        "footer_margin": footer,
        "gutter_margin": gutter,
        "has_duplicate_geometry": has_duplicate_geometry,
    }


def _extract_columns_info(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract column count, spacing, and equal-width status from a section."""
    cols_elems = sect_pr.findall(Q(ns, "cols"))
    if not cols_elems:
        return {
            "count": 1,
            "spacing": _DEFAULT_COLUMN_SPACING,
            "equal_width": True,
            "separator": False,
            "widths": [],
            "spaces": [],
            "is_valid": True,
            "has_unequal_columns": False,
            "has_duplicate_cols": False,
        }

    has_duplicate_cols = len(cols_elems) > 1
    has_unequal = has_duplicate_cols

    for cols_elem in cols_elems:
        equal_width_attr = cols_elem.get(Q(ns, "equalWidth")) or cols_elem.get("equalWidth")
        if equal_width_attr in {"0", "false"}:
            has_unequal = True
            break
        col_children = cols_elem.findall(Q(ns, "col"))
        if col_children:
            col_widths = [
                w for c in col_children if (w := (c.get(Q(ns, "w")) or c.get("w"))) is not None
            ]
            if len(col_widths) != len(col_children) or len(set(col_widths)) > 1:
                has_unequal = True
                break
            if equal_width_attr not in {"1", "true"}:
                has_unequal = True
                break

    first_cols = cols_elems[0]
    raw_num = first_cols.get(Q(ns, "num")) or first_cols.get("num")
    count = int(raw_num) if raw_num is not None and raw_num.isdigit() else 1

    raw_space = first_cols.get(Q(ns, "space")) or first_cols.get("space")
    spacing = (
        int(raw_space) if raw_space is not None and raw_space.isdigit() else _DEFAULT_COLUMN_SPACING
    )

    children = first_cols.findall(Q(ns, "col"))
    widths = [_signed_int_attr(child, ns, "w") for child in children]
    spaces = [_signed_int_attr(child, ns, "space") for child in children]
    equal_width = (first_cols.get(Q(ns, "equalWidth")) or "1") not in {"0", "false", "off"}
    if not equal_width and raw_num is None:
        count = len(children)
    valid = 1 <= count <= _MAX_COLUMN_COUNT and 0 <= spacing <= _MAX_PAGE_DIMENSION
    if not equal_width:
        valid = (
            valid
            and len(widths) == count
            and all(width is not None and 0 < width <= _MAX_PAGE_DIMENSION for width in widths)
            and all(space is None or 0 <= space <= _MAX_PAGE_DIMENSION for space in spaces)
        )
    elif has_unequal:
        valid = False
    return {
        "count": count,
        "spacing": spacing,
        "equal_width": equal_width,
        "separator": (first_cols.get(Q(ns, "sep")) or "0") not in {"0", "false", "off"},
        "widths": widths,
        "spaces": spaces,
        "is_valid": valid,
        "has_unequal_columns": has_unequal,
        "has_duplicate_cols": has_duplicate_cols,
    }


def _signed_int_attr(elem: ET.Element, ns: str, name: str) -> int | None:
    raw = elem.get(Q(ns, name))
    try:
        return int(raw) if raw is not None else None
    except ValueError:
        return None


def _extract_document_grid(sect_pr: ET.Element, ns: str) -> dict[str, Any] | None:
    grid = sect_pr.find(Q(ns, "docGrid"))
    if grid is None:
        return None
    return {
        "type": grid.get(Q(ns, "type")),
        "line_pitch": _signed_int_attr(grid, ns, "linePitch"),
        "char_space": _signed_int_attr(grid, ns, "charSpace"),
    }


def _validate_document_grid(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    if not isinstance(value, Mapping) or not value or not set(value) <= _GRID_FIELDS.keys():
        raise ValueError("page_setup_document_grid_invalid")
    for name, entry in value.items():
        if name == "type":
            if not isinstance(entry, str) or entry not in _GRID_TYPES:
                raise ValueError("page_setup_document_grid_type_invalid")
        else:
            lower, upper = (
                (0, _MAX_PAGE_DIMENSION) if name == "line_pitch" else (-(2**31), 2**31 - 1)
            )
            if isinstance(entry, bool) or not isinstance(entry, int) or not lower <= entry <= upper:
                raise ValueError(f"page_setup_document_grid_{name}_invalid")
    return dict(value)


def _extract_page_numbering_info(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract explicit page numbering attributes (start, fmt) from a section."""
    pg_num_elems = sect_pr.findall(Q(ns, "pgNumType"))
    has_duplicate = len(pg_num_elems) > 1
    if not pg_num_elems:
        return {
            "has_pg_num_type": False,
            "start": None,
            "format": None,
            "has_duplicate_pg_num": False,
        }
    first = pg_num_elems[0]
    raw_start = first.get(Q(ns, "start")) or first.get("start")
    start: int | None = None
    if raw_start is not None and raw_start.isdigit():
        start = int(raw_start)
    raw_fmt = first.get(Q(ns, "fmt")) or first.get("fmt")
    return {
        "has_pg_num_type": True,
        "start": start,
        "format": raw_fmt,
        "has_duplicate_pg_num": has_duplicate,
    }


def _extract_line_numbering_info(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract explicit section line-numbering settings (w:lnNumType)."""
    elements = sect_pr.findall(Q(ns, "lnNumType"))
    if not elements:
        return {
            "has_ln_num_type": False,
            "start": None,
            "interval": None,
            "distance": None,
            "restart": None,
            "has_duplicate_ln_num": False,
        }
    element = elements[0]

    def integer(name: str) -> int | None:
        raw = element.get(Q(ns, name)) or element.get(name)
        return int(raw) if raw is not None and raw.isdigit() else None

    return {
        "has_ln_num_type": True,
        "start": integer("start"),
        "interval": integer("countBy"),
        "distance": integer("distance"),
        "restart": element.get(Q(ns, "restart")) or element.get("restart"),
        "has_duplicate_ln_num": len(elements) > 1,
    }


def _is_on_off_enabled(elem: ET.Element | None, ns: str) -> bool:
    """Return True if an ST_OnOff toggle element is present and not explicitly off."""
    if elem is None:
        return False
    val = elem.get(Q(ns, "val"))
    if val is None:
        val = elem.get("val")
    if val is None:
        return True
    return str(val).strip().lower() not in {"0", "false", "off"}


def _is_title_pg_enabled(sect_pr: ET.Element | None, ns: str) -> bool:
    """Return True if first-page header/footer behavior (w:titlePg) is effectively enabled."""
    if sect_pr is None:
        return False
    return _is_on_off_enabled(sect_pr.find(Q(ns, "titlePg")), ns)


def _load_document_settings(
    graph: PackageGraph, namespace: str
) -> tuple[str | None, ET.Element | None]:
    """Resolve and parse the settings part related from the main document.

    Returns ``(None, None)`` when the main document declares no settings
    relationship. Raises ``ValueError("page_setup_settings_invalid")`` when the
    declared settings part is missing, ambiguous, unparsable, or not a settings
    root in the document's WordprocessingML namespace.
    """
    try:
        targets = {
            rel.resolved_target()
            for rel in graph.relationships_for(_DOCUMENT)
            if not rel.is_external and rel.type.rsplit("/", 1)[-1] == "settings"
        }
        if not targets:
            return None, None
        if len(targets) != 1:
            raise ValueError("page_setup_settings_invalid")
        (name,) = targets
        root = parse_xml_with_preserved_qnames(graph.read_part_bytes(name))
    except (KeyError, PackageError, ET.ParseError):
        raise ValueError("page_setup_settings_invalid") from None
    if root.tag != Q(namespace, "settings"):
        raise ValueError("page_setup_settings_invalid")
    return name, root


def _is_even_and_odd_headers_enabled(settings_root: ET.Element | None, ns: str) -> bool:
    """Return True if the document-wide w:evenAndOddHeaders toggle is effectively on."""
    if settings_root is None:
        return False
    return _is_on_off_enabled(settings_root.find(Q(ns, "evenAndOddHeaders")), ns)


def _settings_children_except(settings_root: ET.Element, tag: str) -> tuple[bytes, ...]:
    """Serialize every settings child other than ``tag`` in document order."""
    return tuple(ET.tostring(child) for child in settings_root if child.tag != tag)


@dataclass(frozen=True)
class _EvenAndOddHeadersPlan:
    """Package edits that set the document-wide w:evenAndOddHeaders toggle."""

    replacements: dict[str, bytes]
    additions: dict[str, bytes]
    preserved_settings_children: tuple[bytes, ...]
    inserted: bool


def _plan_even_and_odd_headers(
    graph: PackageGraph, namespace: str, enabled: bool
) -> _EvenAndOddHeadersPlan:
    """Plan the minimal settings edit for one w:evenAndOddHeaders request.

    Only the toggle changes: enabling inserts ``<w:evenAndOddHeaders/>`` in
    settings schema order (or drops an explicit off value), disabling removes
    it, and an already-satisfied request plans no edit. A package without a
    settings relationship gets a new settings part only when enabling. Raises
    ``ValueError`` carrying the refusal code.
    """
    toggle_tag = Q(namespace, "evenAndOddHeaders")
    settings_part, settings_root = _load_document_settings(graph, namespace)
    if settings_part is None or settings_root is None:
        if not enabled:
            return _EvenAndOddHeadersPlan({}, {}, (), inserted=False)
        if namespace != _W or _SETTINGS in graph.parts:
            raise ValueError("page_setup_settings_unsupported")
        return _plan_new_settings_part(graph, namespace)

    toggles = settings_root.findall(toggle_tag)
    if len(toggles) > 1:
        raise ValueError("page_setup_settings_unsupported")
    preserved = _settings_children_except(settings_root, toggle_tag)
    toggle = toggles[0] if toggles else None
    inserted = False
    if enabled:
        if toggle is None:
            _insert_settings_child(settings_root, ET.Element(toggle_tag), namespace)
            inserted = True
        elif _is_on_off_enabled(toggle, namespace):
            return _EvenAndOddHeadersPlan({}, {}, preserved, inserted=False)
        else:
            toggle.attrib.pop(Q(namespace, "val"), None)
            toggle.attrib.pop("val", None)
    elif toggle is None:
        return _EvenAndOddHeadersPlan({}, {}, preserved, inserted=False)
    else:
        settings_root.remove(toggle)
    return _EvenAndOddHeadersPlan(
        {settings_part: serialize_xml_with_preserved_qnames(settings_root)},
        {},
        preserved,
        inserted=inserted,
    )


def _plan_new_settings_part(graph: PackageGraph, namespace: str) -> _EvenAndOddHeadersPlan:
    """Create word/settings.xml holding only the toggle, its relationship, and override."""
    settings_root = ET.Element(Q(namespace, "settings"))
    ET.SubElement(settings_root, Q(namespace, "evenAndOddHeaders"))
    additions = {_SETTINGS: serialize_xml_with_preserved_qnames(settings_root)}
    replacements: dict[str, bytes] = {}

    rels_name = _rels_name(_DOCUMENT)
    try:
        if rels_name in graph.parts:
            rels_root = parse_xml_with_preserved_qnames(graph.read_part_bytes(rels_name))
        else:
            rels_root = ET.Element(Q(_PR, "Relationships"))
        types_root = parse_xml_with_preserved_qnames(graph.read_part_bytes("[Content_Types].xml"))
    except (KeyError, PackageError, ET.ParseError):
        raise ValueError("page_setup_package_invalid") from None

    rels_root.append(
        ET.Element(
            Q(_PR, "Relationship"),
            {
                "Id": _new_rel_id(rels_root),
                "Type": f"{_R}/settings",
                "Target": "settings.xml",
            },
        )
    )
    rels_bytes = serialize_xml_with_preserved_qnames(rels_root)
    if rels_name in graph.parts:
        replacements[rels_name] = rels_bytes
    else:
        additions[rels_name] = rels_bytes

    if not any(
        node.get("PartName") == f"/{_SETTINGS}" for node in types_root.findall(Q(_CT, "Override"))
    ):
        types_root.append(
            ET.Element(
                Q(_CT, "Override"),
                {"PartName": f"/{_SETTINGS}", "ContentType": _SETTINGS_CONTENT_TYPE},
            )
        )
    replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)
    return _EvenAndOddHeadersPlan(replacements, additions, (), inserted=True)


def _even_and_odd_headers_published(
    graph: PackageGraph,
    namespace: str,
    plan: _EvenAndOddHeadersPlan,
    enabled: bool,
) -> bool:
    """True when a staged package carries exactly the planned toggle state."""
    toggle_tag = Q(namespace, "evenAndOddHeaders")
    _, settings_root = _load_document_settings(graph, namespace)
    if settings_root is None:
        return not enabled
    toggles = settings_root.findall(toggle_tag)
    if enabled:
        if len(toggles) != 1 or not _is_on_off_enabled(toggles[0], namespace):
            return False
        if plan.inserted:
            rank = _SETTINGS_CHILD_ORDER_MAP["evenAndOddHeaders"]
            children = list(settings_root)
            if any(
                isinstance(child.tag, str)
                and _SETTINGS_CHILD_ORDER_MAP.get(child.tag.rsplit("}", 1)[-1], 999) > rank
                for child in children[: children.index(toggles[0])]
            ):
                return False
    elif toggles:
        return False
    return _settings_children_except(settings_root, toggle_tag) == (
        plan.preserved_settings_children
    )


def _extract_vertical_alignment_info(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract effective section vertical alignment and duplicate state."""
    elements = sect_pr.findall(Q(ns, "vAlign"))
    if not elements:
        return {
            "vertical_alignment": "top",
            "has_v_align": False,
            "has_duplicate_v_align": False,
            "is_valid": True,
        }
    raw = elements[0].get(Q(ns, "val")) or elements[0].get("val") or "top"
    value = raw.strip()
    return {
        "vertical_alignment": value,
        "has_v_align": True,
        "has_duplicate_v_align": len(elements) > 1,
        "is_valid": value in _ALLOWED_VERTICAL_ALIGNMENTS,
    }


def _extract_page_borders_info(sect_pr: ET.Element, ns: str) -> dict[str, Any]:
    """Extract four-sided page borders and duplicate state without normalization."""
    elements = sect_pr.findall(Q(ns, "pgBorders"))
    if not elements:
        return {
            "page_borders": None,
            "has_pg_borders": False,
            "has_duplicate_pg_borders": False,
        }

    borders: dict[str, dict[str, Any]] = {}
    for side in _PAGE_BORDER_SIDES:
        element = elements[0].find(Q(ns, side))
        if element is None:
            continue
        width = element.get(Q(ns, "sz")) or element.get("sz")
        borders[side] = {
            "style": element.get(Q(ns, "val")) or element.get("val"),
            "width": int(width) if width is not None and width.isdigit() else None,
            "color": element.get(Q(ns, "color")) or element.get("color"),
        }
    return {
        "page_borders": borders or None,
        "has_pg_borders": True,
        "has_duplicate_pg_borders": len(elements) > 1,
    }


def _is_unsupported_section(
    sect_pr: ET.Element,
    unsupported_by_id: dict[int, bool],
    ns: str,
) -> bool:
    """True if section or any descendant has unsupported context or revision wrappers."""
    if unsupported_by_id.get(id(sect_pr), False):
        return True
    unsupported_tags = {
        Q(ns, "sectPrChange"),
        Q(ns, "ins"),
        Q(ns, "del"),
        Q(ns, "moveFrom"),
        Q(ns, "moveTo"),
        Q(ns, "sdt"),
        Q(ns, "customXml"),
    }
    for child in sect_pr.iter():
        if child.tag in unsupported_tags or unsupported_by_id.get(id(child), False):
            return True
    return False


@dataclass(frozen=True)
class PageSizeSpec:
    """Explicit page size dimensions in twips."""

    width: int
    height: int
    code: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"width": self.width, "height": self.height}
        if self.code is not None:
            data["code"] = self.code
        return data


@dataclass(frozen=True)
class MarginsSpec:
    """Page margins in twips.

    ``header_distance`` and ``footer_distance`` are constructor-only python-docx
    spellings of ``header`` and ``footer``; a contradictory pair is refused.
    """

    top: int | None = None
    bottom: int | None = None
    left: int | None = None
    right: int | None = None
    header: int | None = None
    footer: int | None = None
    gutter: int | None = None
    header_distance: InitVar[int | None] = None
    footer_distance: InitVar[int | None] = None

    def __post_init__(self, header_distance: int | None, footer_distance: int | None) -> None:
        for (_, canonical), value in zip(
            _MARGIN_ALIASES, (header_distance, footer_distance), strict=True
        ):
            if value is None:
                continue
            current = getattr(self, canonical)
            if current is not None and current != value:
                raise ValueError(f"page_setup_{canonical}_margin_conflict")
            object.__setattr__(self, canonical, value)

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for key in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
            val = getattr(self, key)
            if val is not None:
                data[key] = val
        return data


@dataclass(frozen=True)
class ColumnsSpec:
    """Column configuration; explicit widths select w:equalWidth='0'.

    Validate at the parser and author boundaries so invalid direct specs produce
    the same named, atomic refusal as other direct PageSetupRequest fields.
    """

    count: int = 1
    spacing: int = _DEFAULT_COLUMN_SPACING
    separator: bool | None = None
    widths: tuple[int, ...] | None = None
    spaces: tuple[int | None, ...] | None = None

    def validate(self) -> None:
        for name, value, lower, upper in (
            ("count", self.count, 1, _MAX_COLUMN_COUNT),
            ("spacing", self.spacing, 0, _MAX_PAGE_DIMENSION),
        ):
            if isinstance(value, bool) or not isinstance(value, int) or not lower <= value <= upper:
                raise ValueError(f"page_setup_column_{name}_invalid")
        if self.separator is not None and not isinstance(self.separator, bool):
            raise ValueError("page_setup_column_separator_invalid")
        for name, values, lower in (("widths", self.widths, 1), ("spaces", self.spaces, 0)):
            if values is None:
                continue
            if not isinstance(values, (tuple, list)) or len(values) != self.count:
                raise ValueError(f"page_setup_column_{name}_invalid")
            for value in values:
                if name == "spaces" and value is None:
                    continue
                if (
                    isinstance(value, bool)
                    or not isinstance(value, int)
                    or not lower <= value <= _MAX_PAGE_DIMENSION
                ):
                    raise ValueError(f"page_setup_column_{name}_invalid")
        if self.spaces is not None and self.widths is None:
            raise ValueError("page_setup_column_spaces_invalid")

    def to_dict(self) -> dict[str, Any]:
        data = {"count": self.count, "spacing": self.spacing}
        for name in ("separator", "widths", "spaces"):
            value = getattr(self, name)
            if value is not None:
                data[name] = list(value) if name in {"widths", "spaces"} else value
        return data


@dataclass(frozen=True)
class PageBorderSpec:
    """One page-border side with a WordprocessingML line style, width, and color."""

    style: str
    width: int
    color: str

    def __post_init__(self) -> None:
        if not isinstance(self.style, str) or self.style not in _VALID_PAGE_BORDER_STYLES:
            raise ValueError("page_setup_page_border_style_invalid")
        if (
            isinstance(self.width, bool)
            or not isinstance(self.width, int)
            or not 0 <= self.width <= 96
        ):
            raise ValueError("page_setup_page_border_width_invalid")
        if not isinstance(self.color, str) or (
            self.color != "auto" and not _HEX_COLOR_PATTERN.match(self.color)
        ):
            raise ValueError("page_setup_page_border_color_invalid")

    def to_dict(self) -> dict[str, Any]:
        return {"style": self.style, "width": self.width, "color": self.color}


@dataclass(frozen=True)
class PageBordersSpec:
    """Complete four-sided page-border configuration for one section."""

    top: PageBorderSpec
    left: PageBorderSpec
    bottom: PageBorderSpec
    right: PageBorderSpec

    def __post_init__(self) -> None:
        for side in _PAGE_BORDER_SIDES:
            if not isinstance(getattr(self, side), PageBorderSpec):
                raise TypeError("page_setup_page_borders_invalid")

    def to_dict(self) -> dict[str, Any]:
        return {side: getattr(self, side).to_dict() for side in _PAGE_BORDER_SIDES}


@dataclass(frozen=True)
class PageNumberingSpec:
    """Explicit section page numbering configuration in WordprocessingML (w:pgNumType)."""

    start: int | None = None
    format: str | None = None
    clear_start: bool = False
    clear_format: bool = False
    clear: bool = False

    def __post_init__(self) -> None:
        if not isinstance(self.clear, bool):
            raise TypeError("page_numbering_clear_invalid")
        if not isinstance(self.clear_start, bool):
            raise TypeError("page_numbering_clear_start_invalid")
        if not isinstance(self.clear_format, bool):
            raise TypeError("page_numbering_clear_format_invalid")

        if self.start is not None and (
            isinstance(self.start, bool)
            or not isinstance(self.start, int)
            or not (1 <= self.start <= _MAX_PAGE_NUMBER)
        ):
            raise ValueError("page_numbering_start_invalid")

        if self.format is not None and (
            not isinstance(self.format, str) or self.format not in _ALLOWED_PAGE_NUMBER_FORMATS
        ):
            raise ValueError("page_numbering_format_invalid")

        if self.clear and (
            self.start is not None
            or self.format is not None
            or self.clear_start
            or self.clear_format
        ):
            raise ValueError("page_numbering_clear_conflict")
        if self.clear_start and self.start is not None:
            raise ValueError("page_numbering_clear_conflict")
        if self.clear_format and self.format is not None:
            raise ValueError("page_numbering_clear_conflict")
        if (
            self.start is None
            and self.format is None
            and not self.clear
            and not self.clear_start
            and not self.clear_format
        ):
            raise ValueError("page_numbering_empty")

    @property
    def fmt(self) -> str | None:
        return self.format

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        if self.start is not None:
            data["start"] = self.start
        if self.format is not None:
            data["format"] = self.format
        if self.clear_start:
            data["clear_start"] = True
        if self.clear_format:
            data["clear_format"] = True
        if self.clear:
            data["clear"] = True
        return data


@dataclass(frozen=True)
class LineNumberingSpec:
    """Explicit section line numbering configuration in WordprocessingML."""

    start: int | None = None
    interval: int | None = None
    distance: int | None = None
    restart: str | None = None
    clear: bool = False

    def __post_init__(self) -> None:
        if not isinstance(self.clear, bool):
            raise TypeError("line_numbering_clear_invalid")
        for name, value in (
            ("start", self.start),
            ("interval", self.interval),
            ("distance", self.distance),
        ):
            if value is not None and (
                isinstance(value, bool)
                or not isinstance(value, int)
                or not (0 <= value <= _MAX_LINE_NUMBER)
                or (name != "distance" and value < 1)
            ):
                raise ValueError(f"line_numbering_{name}_invalid")
        if self.restart is not None and self.restart not in _ALLOWED_LINE_NUMBER_RESTARTS:
            raise ValueError("line_numbering_restart_invalid")
        if self.clear and any(
            value is not None for value in (self.start, self.interval, self.distance, self.restart)
        ):
            raise ValueError("line_numbering_clear_conflict")
        if not self.clear and all(
            value is None for value in (self.start, self.interval, self.distance, self.restart)
        ):
            raise ValueError("line_numbering_empty")

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for name in ("start", "interval", "distance", "restart"):
            value = getattr(self, name)
            if value is not None:
                data[name] = value
        if self.clear:
            data["clear"] = True
        return data


@dataclass(frozen=True)
class FootnotePropertiesSpec:
    """Section footnote numbering in WordprocessingML (w:footnotePr, ECMA-376 17.11.11)."""

    format: str | None = None
    start: int | None = None
    restart: str | None = None
    clear: bool = False

    def __post_init__(self) -> None:
        if not isinstance(self.clear, bool):
            raise TypeError("footnote_properties_clear_invalid")
        if self.format is not None and (
            not isinstance(self.format, str) or self.format not in _ALLOWED_FOOTNOTE_NUMBER_FORMATS
        ):
            raise ValueError("footnote_properties_format_invalid")
        if self.start is not None and (
            isinstance(self.start, bool)
            or not isinstance(self.start, int)
            or not (1 <= self.start <= _MAX_FOOTNOTE_NUMBER)
        ):
            raise ValueError("footnote_properties_start_invalid")
        if self.restart is not None and (
            not isinstance(self.restart, str) or self.restart not in _ALLOWED_FOOTNOTE_RESTARTS
        ):
            raise ValueError("footnote_properties_restart_invalid")
        declared = (self.format, self.start, self.restart)
        if self.clear and any(value is not None for value in declared):
            raise ValueError("footnote_properties_clear_conflict")
        if not self.clear and all(value is None for value in declared):
            raise ValueError("footnote_properties_empty")

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for name in ("format", "start", "restart"):
            value = getattr(self, name)
            if value is not None:
                data[name] = value
        if self.clear:
            data["clear"] = True
        return data


def _footnote_properties_spec(value: Any) -> FootnotePropertiesSpec | None:
    """Coerce one footnote_properties request value into a validated spec."""
    if value is None or isinstance(value, FootnotePropertiesSpec):
        return value
    if not isinstance(value, Mapping) or not value:
        raise TypeError("page_setup_footnote_properties_invalid")
    if not set(value).issubset(_ALLOWED_FOOTNOTE_PROPERTIES_FIELDS):
        raise ValueError("page_setup_footnote_properties_invalid")
    return FootnotePropertiesSpec(
        format=value.get("format"),
        start=value.get("start"),
        restart=value.get("restart"),
        clear=value.get("clear", False),
    )


def _resolve_different_first_page(canonical: Any, alias: Any) -> Any:
    """Merge ``different_first_page`` with python-docx ``different_first_page_header_footer``."""
    if alias is None:
        return canonical
    if not isinstance(alias, bool):
        raise ValueError("page_setup_different_first_page_invalid")  # noqa: TRY004 - matches canonical flag
    if canonical is not None and canonical is not alias:
        raise ValueError("page_setup_different_first_page_conflict")
    return alias


@dataclass(frozen=True)
class PageSetupTarget:
    """Discovered section page setup candidate target."""

    section: int
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    page_size: dict[str, Any] | None = None
    orientation: str = "portrait"
    margins: dict[str, Any] | None = None
    columns: dict[str, Any] | None = None
    page_numbering: dict[str, Any] | None = None
    line_numbering: dict[str, Any] | None = None
    different_first_page: bool = False
    different_odd_and_even_pages: bool = False
    vertical_alignment: str = "top"
    page_borders: dict[str, Any] | None = None
    usable_area: dict[str, Any] | None = None
    supported: bool = True
    footnote_properties: dict[str, Any] | None = None
    document_grid: dict[str, Any] | None = None

    @property
    def target(self) -> dict[str, Any]:
        return {"section": self.section, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        return {
            "section": self.section,
            "part": self.part,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "page_size": self.page_size,
            "orientation": self.orientation,
            "margins": self.margins,
            "columns": self.columns,
            "page_numbering": self.page_numbering,
            "line_numbering": self.line_numbering,
            "different_first_page": self.different_first_page,
            "different_odd_and_even_pages": self.different_odd_and_even_pages,
            "vertical_alignment": self.vertical_alignment,
            "page_borders": self.page_borders,
            "footnote_properties": self.footnote_properties,
            "document_grid": self.document_grid,
            "usable_area": self.usable_area,
            "supported": self.supported,
        }


@dataclass(frozen=True)
class PageSetupRequest:
    """Strict request for modifying page setup on one source-bound section."""

    source: Path
    destination: Path
    source_sha256: str
    structure_fingerprint_sha256: str
    section: int = 1
    page_size: PageSizeSpec | None = None
    orientation: str | None = None
    margins: MarginsSpec | None = None
    columns: ColumnsSpec | None = None
    page_numbering: PageNumberingSpec | None = None
    line_numbering: LineNumberingSpec | None = None
    different_first_page: bool | None = None
    different_odd_and_even_pages: bool | None = None
    vertical_alignment: str | None = None
    page_borders: PageBordersSpec | None = None
    part: str = _DOCUMENT
    footnote_properties: FootnotePropertiesSpec | Mapping[str, Any] | None = None
    document_grid: Mapping[str, Any] | None = None
    # Constructor-only python-docx spelling of ``different_first_page``.
    different_first_page_header_footer: InitVar[bool | None] = None

    def __post_init__(self, different_first_page_header_footer: bool | None) -> None:
        object.__setattr__(
            self,
            "different_first_page",
            _resolve_different_first_page(
                self.different_first_page, different_first_page_header_footer
            ),
        )
        object.__setattr__(
            self, "footnote_properties", _footnote_properties_spec(self.footnote_properties)
        )

    @property
    def target(self) -> dict[str, Any]:
        return {"section": self.section, "part": self.part}

    def target_dict(self) -> dict[str, Any]:
        return {"section": self.section, "part": self.part}

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "section": self.section,
            "part": self.part,
        }
        if self.page_size is not None:
            data["page_size"] = self.page_size.to_dict()
        if self.orientation is not None:
            data["orientation"] = self.orientation
        if self.margins is not None:
            data["margins"] = self.margins.to_dict()
        if self.page_numbering is not None:
            data["page_numbering"] = self.page_numbering.to_dict()
        if self.line_numbering is not None:
            data["line_numbering"] = self.line_numbering.to_dict()
        if self.columns is not None:
            data["columns"] = self.columns.to_dict()
        if self.different_first_page is not None:
            data["different_first_page"] = self.different_first_page
        if self.different_odd_and_even_pages is not None:
            data["different_odd_and_even_pages"] = self.different_odd_and_even_pages
        if self.vertical_alignment is not None:
            data["vertical_alignment"] = self.vertical_alignment
        if self.page_borders is not None:
            data["page_borders"] = self.page_borders.to_dict()
        if self.footnote_properties is not None:
            data["footnote_properties"] = self.footnote_properties.to_dict()
        if self.document_grid is not None:
            data["document_grid"] = dict(self.document_grid)
        return data


@dataclass(frozen=True)
class PageSetupReport:
    """Content-safe public report for one page setup authoring operation."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    page_setup: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "schema_version": 1,
            "operation": "page-setup",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
        }
        if self.target is not None:
            result["target"] = self.target
        if self.page_setup is not None:
            result["page_setup"] = self.page_setup
        if self.target_change_report is not None:
            result["target_change_report"] = self.target_change_report
        if self.non_target_parts_preserved is not None:
            result["non_target_parts_preserved"] = self.non_target_parts_preserved
        return result


def _page_setup_refusal(
    request: PageSetupRequest,
    code: str,
) -> PageSetupReport:
    return PageSetupReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target=request.target_dict(),
        page_setup=None,
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_page_setup_request(value: Any) -> PageSetupRequest:
    """Parse and strictly validate a page setup authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("page_setup_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("page_setup_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("page_setup_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("page_setup_destination_invalid")

    source_sha256 = _require_sha256(value["source_sha256"], "page_setup_source_sha256")
    structure_fingerprint = _require_sha256(
        value["structure_fingerprint_sha256"], "page_setup_structure_fingerprint_sha256"
    )

    section = value["section"]
    if isinstance(section, bool) or not isinstance(section, int) or section < 1:
        raise ValueError("page_setup_section_invalid")

    part = value.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise ValueError("page_setup_part_invalid")

    if all(
        value.get(k) is None
        for k in (
            "page_size",
            "orientation",
            "margins",
            "columns",
            "page_numbering",
            "line_numbering",
            "different_first_page",
            "different_first_page_header_footer",
            "different_odd_and_even_pages",
            "vertical_alignment",
            "page_borders",
            "footnote_properties",
            "document_grid",
        )
    ):
        raise ValueError("page_setup_request_empty")

    page_size_spec: PageSizeSpec | None = None
    if "page_size" in value and value["page_size"] is not None:
        raw_ps = value["page_size"]
        if not isinstance(raw_ps, Mapping):
            raise TypeError("page_setup_page_size_invalid")
        ps_keys = set(raw_ps)
        if not ps_keys.issubset(
            _ALLOWED_PAGE_SIZE_FIELDS
        ) or not _REQUIRED_PAGE_SIZE_FIELDS.issubset(ps_keys):
            raise ValueError("page_setup_page_size_invalid")

        width = raw_ps["width"]
        if (
            isinstance(width, bool)
            or not isinstance(width, int)
            or not (_MIN_PAGE_DIMENSION <= width <= _MAX_PAGE_DIMENSION)
        ):
            raise ValueError("page_setup_width_invalid")

        height = raw_ps["height"]
        if (
            isinstance(height, bool)
            or not isinstance(height, int)
            or not (_MIN_PAGE_DIMENSION <= height <= _MAX_PAGE_DIMENSION)
        ):
            raise ValueError("page_setup_height_invalid")

        raw_code = raw_ps.get("code")
        code: str | None = None
        if raw_code is not None:
            if isinstance(raw_code, bool) or not isinstance(raw_code, (int, str)):
                raise ValueError("page_setup_code_invalid")
            code = str(raw_code)

        page_size_spec = PageSizeSpec(width=width, height=height, code=code)

    orientation: str | None = None
    if "orientation" in value and value["orientation"] is not None:
        raw_orient = value["orientation"]
        if not isinstance(raw_orient, str) or raw_orient.lower() not in _ALLOWED_ORIENTATIONS:
            raise ValueError("page_setup_orientation_invalid")
        orientation = raw_orient.lower()

    margins_spec: MarginsSpec | None = None
    if "margins" in value and value["margins"] is not None:
        raw_mar = value["margins"]
        if not isinstance(raw_mar, Mapping) or not raw_mar:
            raise TypeError("page_setup_margins_invalid")
        mar_keys = set(raw_mar)
        if not mar_keys.issubset(_ALLOWED_MARGINS_FIELDS):
            raise ValueError("page_setup_margins_invalid")

        mar_kwargs: dict[str, int] = {}
        for m_key in _ALLOWED_MARGINS_FIELDS:
            if m_key in raw_mar:
                val = raw_mar[m_key]
                if (
                    isinstance(val, bool)
                    or not isinstance(val, int)
                    or not (0 <= val <= _MAX_PAGE_DIMENSION)
                ):
                    raise ValueError(f"page_setup_{m_key}_margin_invalid")
                mar_kwargs[m_key] = val
        margins_spec = MarginsSpec(**mar_kwargs)

    columns_spec: ColumnsSpec | None = None
    if "columns" in value and value["columns"] is not None:
        raw_cols = value["columns"]
        if not isinstance(raw_cols, Mapping) or not raw_cols:
            raise TypeError("page_setup_columns_invalid")
        cols_keys = set(raw_cols)
        if not cols_keys.issubset(_ALLOWED_COLUMNS_FIELDS) or "count" not in cols_keys:
            raise ValueError("page_setup_columns_invalid")

        columns_spec = ColumnsSpec(**raw_cols)
        columns_spec.validate()
    page_numbering_spec: PageNumberingSpec | None = None
    if "page_numbering" in value and value["page_numbering"] is not None:
        raw_num = value["page_numbering"]
        if isinstance(raw_num, PageNumberingSpec):
            page_numbering_spec = raw_num
        elif not isinstance(raw_num, Mapping) or not raw_num:
            raise TypeError("page_setup_page_numbering_invalid")
        else:
            num_keys = set(raw_num)
            if not num_keys.issubset(_ALLOWED_PAGE_NUMBERING_FIELDS):
                raise ValueError("page_setup_page_numbering_invalid")

            clear = raw_num.get("clear", False)
            clear_start = raw_num.get("clear_start", False)
            clear_format = raw_num.get("clear_format", False)
            raw_start = raw_num.get("start")
            raw_fmt = raw_num.get("format", raw_num.get("fmt"))

            page_numbering_spec = PageNumberingSpec(
                start=raw_start,
                format=raw_fmt,
                clear_start=clear_start,
                clear_format=clear_format,
                clear=clear,
            )

    line_numbering_spec: LineNumberingSpec | None = None
    if "line_numbering" in value and value["line_numbering"] is not None:
        raw_line = value["line_numbering"]
        if isinstance(raw_line, LineNumberingSpec):
            line_numbering_spec = raw_line
        elif not isinstance(raw_line, Mapping) or not raw_line:
            raise TypeError("page_setup_line_numbering_invalid")
        else:
            line_keys = set(raw_line)
            if not line_keys.issubset(_ALLOWED_LINE_NUMBERING_FIELDS) or (
                "interval" in line_keys and "count_by" in line_keys
            ):
                raise ValueError("page_setup_line_numbering_invalid")
            line_numbering_spec = LineNumberingSpec(
                start=raw_line.get("start"),
                interval=raw_line.get("interval", raw_line.get("count_by")),
                distance=raw_line.get("distance"),
                restart=raw_line.get("restart"),
                clear=raw_line.get("clear", False),
            )

    different_first_page: bool | None = None
    if "different_first_page" in value and value["different_first_page"] is not None:
        raw_dfp = value["different_first_page"]
        if not isinstance(raw_dfp, bool):
            raise ValueError("page_setup_different_first_page_invalid")
        different_first_page = raw_dfp
    different_first_page = _resolve_different_first_page(
        different_first_page, value.get("different_first_page_header_footer")
    )

    different_odd_and_even_pages: bool | None = None
    if value.get("different_odd_and_even_pages") is not None:
        raw_odd_even = value["different_odd_and_even_pages"]
        if not isinstance(raw_odd_even, bool):
            raise ValueError("page_setup_different_odd_and_even_pages_invalid")
        different_odd_and_even_pages = raw_odd_even

    vertical_alignment: str | None = None
    if "vertical_alignment" in value and value["vertical_alignment"] is not None:
        raw_vertical_alignment = value["vertical_alignment"]
        if (
            not isinstance(raw_vertical_alignment, str)
            or raw_vertical_alignment not in _ALLOWED_VERTICAL_ALIGNMENTS
        ):
            raise ValueError("page_setup_vertical_alignment_invalid")
        vertical_alignment = raw_vertical_alignment

    page_borders_spec: PageBordersSpec | None = None
    if "page_borders" in value and value["page_borders"] is not None:
        raw_borders = value["page_borders"]
        if isinstance(raw_borders, PageBordersSpec):
            page_borders_spec = raw_borders
        elif not isinstance(raw_borders, Mapping):
            raise TypeError("page_setup_page_borders_invalid")
        elif set(raw_borders) != set(_PAGE_BORDER_SIDES):
            raise ValueError("page_setup_page_borders_invalid")
        else:
            sides: dict[str, PageBorderSpec] = {}
            for side in _PAGE_BORDER_SIDES:
                raw_side = raw_borders[side]
                if isinstance(raw_side, PageBorderSpec):
                    sides[side] = raw_side
                    continue
                if (
                    not isinstance(raw_side, Mapping)
                    or set(raw_side) != _ALLOWED_PAGE_BORDER_FIELDS
                ):
                    raise ValueError("page_setup_page_border_invalid")
                sides[side] = PageBorderSpec(
                    style=raw_side["style"],
                    width=raw_side["width"],
                    color=raw_side["color"],
                )
            page_borders_spec = PageBordersSpec(**sides)

    return PageSetupRequest(
        source=Path(source),
        destination=Path(destination),
        source_sha256=source_sha256,
        structure_fingerprint_sha256=structure_fingerprint,
        section=section,
        page_size=page_size_spec,
        orientation=orientation,
        margins=margins_spec,
        columns=columns_spec,
        page_numbering=page_numbering_spec,
        line_numbering=line_numbering_spec,
        different_first_page=different_first_page,
        different_odd_and_even_pages=different_odd_and_even_pages,
        vertical_alignment=vertical_alignment,
        page_borders=page_borders_spec,
        part=part,
        footnote_properties=_footnote_properties_spec(value.get("footnote_properties")),
        document_grid=_validate_document_grid(value.get("document_grid")),
    )


def load_page_setup_request(path: str | Path) -> PageSetupRequest:
    """Load and strictly validate one JSON page-setup request from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_page_setup_request(json.load(stream))


def page_setup_structure_fingerprint(
    source: str | Path,
    *,
    section: int = 1,
    part: str = _DOCUMENT,
) -> str | None:
    """Return the source-bound structural fingerprint for one exact section."""
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
        sections = _find_sections(root, ns)
        if section < 1 or section > len(sections):
            return None
        target_sect_pr = sections[section - 1]
        unsupported = _unsupported_context_by_element(root, ns)
        if _is_unsupported_section(target_sect_pr, unsupported, ns):
            return None
        return _cell_structure_fingerprint(target_sect_pr)
    except (PackageError, OSError, ET.ParseError, KeyError):
        return None


def discover_page_setup_targets(
    source: str | Path,
    *,
    part: str = _DOCUMENT,
) -> dict[str, Any]:
    """Inventory source-bound candidate section targets for page setup authoring."""
    path = Path(source)
    if not path.is_file():
        return {
            "schema_version": 1,
            "operation": "page-setup-targets",
            "passed": False,
            "code": "page_setup_source_package_invalid",
            "source_sha256": "",
            "part": part,
            "targets": [],
        }

    try:
        source_sha256 = _sha256_file(path)
        graph = PackageGraph.from_file(path)
        if part not in graph.parts:
            return {
                "schema_version": 1,
                "operation": "page-setup-targets",
                "passed": False,
                "code": "page_setup_source_package_invalid",
                "source_sha256": source_sha256,
                "part": part,
                "targets": [],
            }
        root = ET.fromstring(graph.read_part_bytes(part))
        ns = _word_namespace(root)
        if ns is None:
            return {
                "schema_version": 1,
                "operation": "page-setup-targets",
                "passed": False,
                "code": "page_setup_namespace_unsupported",
                "source_sha256": source_sha256,
                "part": part,
                "targets": [],
            }
        try:
            _, settings_root = _load_document_settings(graph, ns)
        except ValueError:
            return {
                "schema_version": 1,
                "operation": "page-setup-targets",
                "passed": False,
                "code": "page_setup_settings_invalid",
                "source_sha256": source_sha256,
                "part": part,
                "targets": [],
            }
    except (OSError, PackageError, ET.ParseError, KeyError):
        return {
            "schema_version": 1,
            "operation": "page-setup-targets",
            "passed": False,
            "code": "page_setup_source_package_invalid",
            "source_sha256": "",
            "part": part,
            "targets": [],
        }

    sections = _find_sections(root, ns)
    unsupported_context = _unsupported_context_by_element(root, ns)
    different_odd_and_even_pages = _is_even_and_odd_headers_enabled(settings_root, ns)

    targets: list[dict[str, Any]] = []
    for idx, sect_pr in enumerate(sections, start=1):
        geom = _extract_page_geometry(sect_pr, ns)
        cols_info = _extract_columns_info(sect_pr, ns)
        num_info = _extract_page_numbering_info(sect_pr, ns)
        line_info = _extract_line_numbering_info(sect_pr, ns)
        vertical_info = _extract_vertical_alignment_info(sect_pr, ns)
        border_info = _extract_page_borders_info(sect_pr, ns)
        footnote_info = _extract_footnote_properties_info(sect_pr, ns)
        fp = _cell_structure_fingerprint(sect_pr)

        has_explicit_size = geom["width"] is not None and geom["height"] is not None
        has_explicit_margins = (
            geom["top"] is not None
            and geom["bottom"] is not None
            and geom["left"] is not None
            and geom["right"] is not None
        )
        is_context_supported = not _is_unsupported_section(sect_pr, unsupported_context, ns)
        has_duplicate = (
            geom.get("has_duplicate_geometry", False)
            or cols_info.get("has_duplicate_cols", False)
            or num_info.get("has_duplicate_pg_num", False)
            or line_info.get("has_duplicate_ln_num", False)
            or len(sect_pr.findall(Q(ns, "titlePg"))) > 1
            or vertical_info["has_duplicate_v_align"]
            or border_info["has_duplicate_pg_borders"]
            or footnote_info["has_duplicate_footnote_pr"]
            or len(sect_pr.findall(Q(ns, "docGrid"))) > 1
        )
        is_supported = (
            has_explicit_size
            and has_explicit_margins
            and cols_info["is_valid"]
            and not has_duplicate
            and is_context_supported
            and vertical_info["is_valid"]
        )

        usable_width: int | None = None
        usable_height: int | None = None
        if has_explicit_size and has_explicit_margins:
            uw = geom["width"] - geom["left"] - geom["right"]
            uh = geom["height"] - geom["top"] - geom["bottom"]
            if uw > 0 and uh > 0:
                usable_width = uw
                usable_height = uh

        col_count = cols_info["count"]
        target = PageSetupTarget(
            section=idx,
            structure_fingerprint_sha256=fp,
            part=part,
            page_size={
                "width": geom["width"],
                "height": geom["height"],
                "code": geom["code"],
            },
            orientation=geom["orientation"],
            margins={
                "top": geom["top"],
                "bottom": geom["bottom"],
                "left": geom["left"],
                "right": geom["right"],
                "header": geom["header_margin"],
                "footer": geom["footer_margin"],
                "gutter": geom["gutter_margin"],
            },
            columns={
                "count": col_count,
                "spacing": cols_info["spacing"],
                "equal_width": cols_info["equal_width"],
                "separator": cols_info["separator"],
                "widths": cols_info["widths"],
                "spaces": cols_info["spaces"],
            },
            page_numbering=(
                {"start": num_info["start"], "format": num_info["format"]}
                if num_info["has_pg_num_type"]
                else None
            ),
            line_numbering=(
                {
                    "start": line_info["start"],
                    "interval": line_info["interval"],
                    "distance": line_info["distance"],
                    "restart": line_info["restart"],
                }
                if line_info["has_ln_num_type"]
                else None
            ),
            different_first_page=_is_title_pg_enabled(sect_pr, ns),
            different_odd_and_even_pages=different_odd_and_even_pages,
            vertical_alignment=vertical_info["vertical_alignment"],
            page_borders=border_info["page_borders"],
            footnote_properties=footnote_info["footnote_properties"],
            document_grid=_extract_document_grid(sect_pr, ns),
            usable_area=(
                {"width": usable_width, "height": usable_height}
                if usable_width is not None and usable_height is not None
                else None
            ),
            supported=is_supported,
        )
        targets.append(target.to_dict())

    return {
        "schema_version": 1,
        "operation": "page-setup-targets",
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha256,
        "part": part,
        "targets": targets,
    }


def author_page_setup(request: PageSetupRequest) -> PageSetupReport:
    """Modify selected section page setup geometry atomically and verification-checked."""
    source = request.source
    destination = request.destination

    try:
        if request.columns is not None:
            if not isinstance(request.columns, ColumnsSpec):
                raise ValueError("page_setup_columns_invalid")
            request.columns.validate()
        grid_spec = _validate_document_grid(request.document_grid)
    except ValueError as error:
        return _page_setup_refusal(request, str(error))

    if request.different_first_page is not None and not isinstance(
        request.different_first_page, bool
    ):
        return _page_setup_refusal(request, "page_setup_different_first_page_invalid")
    if request.different_odd_and_even_pages is not None and not isinstance(
        request.different_odd_and_even_pages, bool
    ):
        return _page_setup_refusal(request, "page_setup_different_odd_and_even_pages_invalid")
    if request.vertical_alignment is not None and (
        not isinstance(request.vertical_alignment, str)
        or request.vertical_alignment not in _ALLOWED_VERTICAL_ALIGNMENTS
    ):
        return _page_setup_refusal(request, "page_setup_vertical_alignment_invalid")
    if request.page_borders is not None and not isinstance(request.page_borders, PageBordersSpec):
        return _page_setup_refusal(request, "page_setup_page_borders_invalid")
    if not source.is_file():
        return _page_setup_refusal(request, "page_setup_source_missing")
    if source.resolve() == destination.resolve():
        return _page_setup_refusal(request, "page_setup_destination_invalid")
    if not destination.parent.is_dir():
        return _page_setup_refusal(request, "page_setup_destination_invalid")
    if destination.exists() or destination.is_symlink():
        return _page_setup_refusal(request, "page_setup_destination_occupied")

    try:
        actual_source_sha256 = _sha256_file(source)
    except OSError:
        return _page_setup_refusal(request, "page_setup_source_missing")

    if actual_source_sha256 != request.source_sha256:
        return _page_setup_refusal(request, "page_setup_source_hash_mismatch")

    if request.part != _DOCUMENT:
        return _page_setup_refusal(request, "page_setup_part_unsupported")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _page_setup_refusal(request, "page_setup_package_invalid")

    if _DOCUMENT not in graph.parts:
        return _page_setup_refusal(request, "page_setup_package_invalid")

    try:
        doc_bytes = graph.read_part_bytes(_DOCUMENT)
        root = parse_xml_with_preserved_qnames(doc_bytes)
    except (KeyError, ET.ParseError):
        return _page_setup_refusal(request, "page_setup_package_invalid")

    namespace = _word_namespace(root)
    if namespace is None:
        return _page_setup_refusal(request, "page_setup_namespace_unsupported")

    unsupported_context = _unsupported_context_by_element(root, namespace)
    sections = _find_sections(root, namespace)
    if request.section < 1 or request.section > len(sections):
        return _page_setup_refusal(request, "page_setup_target_missing")

    target_sect_pr = sections[request.section - 1]
    if _is_unsupported_section(target_sect_pr, unsupported_context, namespace):
        return _page_setup_refusal(request, "page_setup_context_unsupported")
    # Check structural fingerprint
    actual_fp = _cell_structure_fingerprint(target_sect_pr)
    if actual_fp != request.structure_fingerprint_sha256:
        return _page_setup_refusal(request, "page_setup_structure_fingerprint_mismatch")

    # Extract existing geometry and column info
    existing_geom = _extract_page_geometry(target_sect_pr, namespace)
    existing_cols = _extract_columns_info(target_sect_pr, namespace)
    existing_num = _extract_page_numbering_info(target_sect_pr, namespace)
    existing_line = _extract_line_numbering_info(target_sect_pr, namespace)
    existing_vertical = _extract_vertical_alignment_info(target_sect_pr, namespace)
    existing_borders = _extract_page_borders_info(target_sect_pr, namespace)
    existing_footnote = _extract_footnote_properties_info(target_sect_pr, namespace)
    if existing_num.get("has_duplicate_pg_num", False):
        return _page_setup_refusal(request, "page_setup_page_numbering_unsupported")
    if existing_line.get("has_duplicate_ln_num", False):
        return _page_setup_refusal(request, "page_setup_line_numbering_unsupported")

    # Refuse invalid duplicate geometry
    if existing_cols.get("has_duplicate_cols", False):
        return _page_setup_refusal(request, "page_setup_columns_unsupported")
    if existing_geom.get("has_duplicate_geometry", False):
        return _page_setup_refusal(request, "page_setup_section_unsupported")
    if len(target_sect_pr.findall(Q(namespace, "titlePg"))) > 1:
        return _page_setup_refusal(request, "page_setup_section_unsupported")
    if request.vertical_alignment is not None and existing_vertical["has_duplicate_v_align"]:
        return _page_setup_refusal(request, "page_setup_vertical_alignment_unsupported")
    if request.page_borders is not None and existing_borders["has_duplicate_pg_borders"]:
        return _page_setup_refusal(request, "page_setup_page_borders_unsupported")
    if request.footnote_properties is not None and existing_footnote["has_duplicate_footnote_pr"]:
        return _page_setup_refusal(request, "page_setup_footnote_properties_unsupported")
    if not existing_cols["is_valid"]:
        return _page_setup_refusal(request, "page_setup_columns_unsupported")
    if grid_spec is not None and len(target_sect_pr.findall(Q(namespace, "docGrid"))) > 1:
        return _page_setup_refusal(request, "page_setup_document_grid_unsupported")
    # Resolve effective width and height
    if request.page_size is not None and request.page_size.width is not None:
        eff_width = request.page_size.width
    elif existing_geom["width"] is not None:
        eff_width = existing_geom["width"]
    else:
        return _page_setup_refusal(request, "page_setup_geometry_missing")

    if request.page_size is not None and request.page_size.height is not None:
        eff_height = request.page_size.height
    elif existing_geom["height"] is not None:
        eff_height = existing_geom["height"]
    else:
        return _page_setup_refusal(request, "page_setup_geometry_missing")

    # Resolve orientation; w:w and w:h always describe the page as oriented.
    if request.orientation is not None:
        if request.orientation not in _ALLOWED_ORIENTATIONS:
            return _page_setup_refusal(request, "page_setup_orientation_invalid")
        eff_orientation = request.orientation
        if (eff_orientation == "landscape" and eff_width < eff_height) or (
            eff_orientation == "portrait" and eff_height < eff_width
        ):
            if request.page_size is not None:
                return _page_setup_refusal(request, "page_setup_orientation_dimensions_mismatch")
            eff_width, eff_height = eff_height, eff_width
    elif existing_geom["orientation"] is not None:
        eff_orientation = existing_geom["orientation"]
    else:
        eff_orientation = "portrait"

    # Resolve effective margins
    if request.margins is not None and request.margins.top is not None:
        eff_top = request.margins.top
    elif existing_geom["top"] is not None:
        eff_top = existing_geom["top"]
    else:
        return _page_setup_refusal(request, "page_setup_geometry_missing")

    if request.margins is not None and request.margins.bottom is not None:
        eff_bottom = request.margins.bottom
    elif existing_geom["bottom"] is not None:
        eff_bottom = existing_geom["bottom"]
    else:
        return _page_setup_refusal(request, "page_setup_geometry_missing")

    if request.margins is not None and request.margins.left is not None:
        eff_left = request.margins.left
    elif existing_geom["left"] is not None:
        eff_left = existing_geom["left"]
    else:
        return _page_setup_refusal(request, "page_setup_geometry_missing")

    if request.margins is not None and request.margins.right is not None:
        eff_right = request.margins.right
    elif existing_geom["right"] is not None:
        eff_right = existing_geom["right"]
    else:
        return _page_setup_refusal(request, "page_setup_geometry_missing")

    # Validate usable page area
    usable_width = eff_width - eff_left - eff_right
    usable_height = eff_height - eff_top - eff_bottom
    if usable_width <= 0 or usable_height <= 0:
        return _page_setup_refusal(request, "page_setup_usable_area_invalid")

    # Resolve effective columns
    if request.columns is not None:
        eff_col_count = request.columns.count
        eff_col_spacing = request.columns.spacing
    else:
        eff_col_count = existing_cols["count"]
        eff_col_spacing = existing_cols["spacing"]

    widths = (
        request.columns.widths
        if request.columns is not None
        else (existing_cols["widths"] if not existing_cols["equal_width"] else None)
    )
    spaces = request.columns.spaces if request.columns is not None else existing_cols["spaces"]
    if widths:
        gaps = spaces if spaces is not None else [None] * eff_col_count
        occupied = sum(widths) + sum(eff_col_spacing if gap is None else gap for gap in gaps[:-1])
        exceeds_width = occupied > usable_width
    else:
        exceeds_width = eff_col_count > 1 and (eff_col_count - 1) * eff_col_spacing >= usable_width
    if exceeds_width:
        return _page_setup_refusal(request, "page_setup_columns_exceed_usable_width")

    # Document-wide odd/even header toggle lives in the settings part.
    odd_even_plan: _EvenAndOddHeadersPlan | None = None
    if request.different_odd_and_even_pages is not None:
        try:
            odd_even_plan = _plan_even_and_odd_headers(
                graph, namespace, request.different_odd_and_even_pages
            )
        except ValueError as error:
            return _page_setup_refusal(request, str(error))

    # Mutate selected section properties
    # 1. w:pgSz (w:orient is an attribute of w:pgSz, never a child element)
    pg_sz = target_sect_pr.find(Q(namespace, "pgSz"))
    if pg_sz is None:
        pg_sz = ET.Element(Q(namespace, "pgSz"))
        _insert_sect_pr_child(target_sect_pr, pg_sz)

    pg_sz.set(Q(namespace, "w"), str(eff_width))
    pg_sz.set(Q(namespace, "h"), str(eff_height))
    if request.orientation is not None:
        pg_sz.set(Q(namespace, "orient"), eff_orientation)
    if request.page_size is not None and request.page_size.code is not None:
        pg_sz.set(Q(namespace, "code"), request.page_size.code)

    # 2. w:pgMar
    pg_mar = target_sect_pr.find(Q(namespace, "pgMar"))
    if pg_mar is None:
        pg_mar = ET.Element(Q(namespace, "pgMar"))
        _insert_sect_pr_child(target_sect_pr, pg_mar)

    pg_mar.set(Q(namespace, "top"), str(eff_top))
    pg_mar.set(Q(namespace, "bottom"), str(eff_bottom))
    pg_mar.set(Q(namespace, "left"), str(eff_left))
    pg_mar.set(Q(namespace, "right"), str(eff_right))

    if request.margins is not None:
        if request.margins.header is not None:
            pg_mar.set(Q(namespace, "header"), str(request.margins.header))
        if request.margins.footer is not None:
            pg_mar.set(Q(namespace, "footer"), str(request.margins.footer))
        if request.margins.gutter is not None:
            pg_mar.set(Q(namespace, "gutter"), str(request.margins.gutter))

    # 3. w:pgBorders
    if request.page_borders is not None:
        pg_borders = target_sect_pr.find(Q(namespace, "pgBorders"))
        if pg_borders is None:
            pg_borders = ET.Element(Q(namespace, "pgBorders"))
            _insert_sect_pr_child(target_sect_pr, pg_borders)
        for side in _PAGE_BORDER_SIDES:
            for existing_side in list(pg_borders.findall(Q(namespace, side))):
                pg_borders.remove(existing_side)
            spec = getattr(request.page_borders, side)
            pg_borders.append(
                ET.Element(
                    Q(namespace, side),
                    {
                        Q(namespace, "val"): spec.style,
                        Q(namespace, "sz"): str(spec.width),
                        Q(namespace, "color"): spec.color,
                    },
                )
            )

    # 3. w:cols
    if request.columns is not None:
        cols_elem = target_sect_pr.find(Q(namespace, "cols"))
        if cols_elem is None:
            cols_elem = ET.Element(Q(namespace, "cols"))
            _insert_sect_pr_child(target_sect_pr, cols_elem)

        cols_elem.set(Q(namespace, "num"), str(request.columns.count))
        cols_elem.set(Q(namespace, "space"), str(request.columns.spacing))
        cols_elem.set(Q(namespace, "equalWidth"), "0" if request.columns.widths else "1")
        if request.columns.separator is not None:
            cols_elem.set(Q(namespace, "sep"), "1" if request.columns.separator else "0")
        for c in list(cols_elem.findall(Q(namespace, "col"))):
            cols_elem.remove(c)
        if request.columns.widths is not None:
            for index, width in enumerate(request.columns.widths):
                column = ET.SubElement(
                    cols_elem, Q(namespace, "col"), {Q(namespace, "w"): str(width)}
                )
                if request.columns.spaces is not None and request.columns.spaces[index] is not None:
                    column.set(Q(namespace, "space"), str(request.columns.spaces[index]))
    if grid_spec is not None:
        grid = target_sect_pr.find(Q(namespace, "docGrid"))
        if grid is None:
            grid = ET.Element(Q(namespace, "docGrid"))
            _insert_sect_pr_child(target_sect_pr, grid)
        for name, value in grid_spec.items():
            grid.set(Q(namespace, _GRID_FIELDS[name]), str(value))
    # 4. w:pgNumType
    if request.page_numbering is not None:
        num_spec = request.page_numbering
        pg_num = target_sect_pr.find(Q(namespace, "pgNumType"))

        if num_spec.clear:
            if pg_num is not None:
                for attr_name in (
                    Q(namespace, "start"),
                    "start",
                    Q(namespace, "fmt"),
                    "fmt",
                ):
                    pg_num.attrib.pop(attr_name, None)
                if not pg_num.attrib and len(pg_num) == 0:
                    target_sect_pr.remove(pg_num)
        else:
            if num_spec.clear_start and pg_num is not None:
                pg_num.attrib.pop(Q(namespace, "start"), None)
                pg_num.attrib.pop("start", None)

            if num_spec.clear_format and pg_num is not None:
                pg_num.attrib.pop(Q(namespace, "fmt"), None)
                pg_num.attrib.pop("fmt", None)

            if num_spec.start is not None:
                if pg_num is None:
                    pg_num = ET.Element(Q(namespace, "pgNumType"))
                    _insert_sect_pr_child(target_sect_pr, pg_num)
                pg_num.set(Q(namespace, "start"), str(num_spec.start))

            if num_spec.format is not None:
                if pg_num is None:
                    pg_num = ET.Element(Q(namespace, "pgNumType"))
                    _insert_sect_pr_child(target_sect_pr, pg_num)
                pg_num.set(Q(namespace, "fmt"), num_spec.format)

            if pg_num is not None and not pg_num.attrib and len(pg_num) == 0:
                target_sect_pr.remove(pg_num)

    # w:footnotePr: only the requested numbering children change.
    if request.footnote_properties is not None:
        footnote_spec = request.footnote_properties
        footnote_pr = target_sect_pr.find(Q(namespace, "footnotePr"))
        if footnote_spec.clear:
            if footnote_pr is not None:
                target_sect_pr.remove(footnote_pr)
        else:
            if footnote_pr is None:
                footnote_pr = ET.Element(Q(namespace, "footnotePr"))
                _insert_sect_pr_child(target_sect_pr, footnote_pr)
            for field, tag in _FOOTNOTE_PR_FIELDS:
                value = getattr(footnote_spec, field)
                if value is None:
                    continue
                child = footnote_pr.find(Q(namespace, tag))
                if child is None:
                    child = ET.Element(Q(namespace, tag))
                    _insert_footnote_pr_child(footnote_pr, child)
                child.attrib.pop("val", None)
                child.set(Q(namespace, "val"), str(value))

    # 5. w:titlePg
    if request.different_first_page is not None:
        for elem in list(target_sect_pr.findall(Q(namespace, "titlePg"))):
            target_sect_pr.remove(elem)
        if request.different_first_page:
            new_title_pg = ET.Element(Q(namespace, "titlePg"))
            _insert_sect_pr_child(target_sect_pr, new_title_pg)

    # 6. w:lnNumType
    if request.line_numbering is not None:
        line_spec = request.line_numbering
        line_num = target_sect_pr.find(Q(namespace, "lnNumType"))
        if line_spec.clear:
            if line_num is not None:
                target_sect_pr.remove(line_num)
        else:
            if line_num is None:
                line_num = ET.Element(Q(namespace, "lnNumType"))
                _insert_sect_pr_child(target_sect_pr, line_num)
            for attr, value in (
                ("start", line_spec.start),
                ("countBy", line_spec.interval),
                ("distance", line_spec.distance),
                ("restart", line_spec.restart),
            ):
                if value is not None:
                    line_num.set(Q(namespace, attr), str(value))

    # 7. w:vAlign
    if request.vertical_alignment is not None:
        v_align = target_sect_pr.find(Q(namespace, "vAlign"))
        if v_align is None:
            v_align = ET.Element(Q(namespace, "vAlign"))
            _insert_sect_pr_child(target_sect_pr, v_align)
        v_align.set(Q(namespace, "val"), request.vertical_alignment)

    def _get_int_attr(elem: ET.Element | None, name: str) -> int | None:
        if elem is None:
            return None
        raw = elem.get(Q(namespace, name)) or elem.get(name)
        return int(raw) if raw is not None and raw.isdigit() else None

    fn = _extract_page_numbering_info(target_sect_pr, namespace)
    ln = _extract_line_numbering_info(target_sect_pr, namespace)
    va = _extract_vertical_alignment_info(target_sect_pr, namespace)
    pb = _extract_page_borders_info(target_sect_pr, namespace)
    fnp = _extract_footnote_properties_info(target_sect_pr, namespace)
    target_change_report: dict[str, Any] = {
        "section": request.section,
        "page_size": {
            "width": eff_width,
            "height": eff_height,
            "code": pg_sz.get(Q(namespace, "code")),
        },
        "orientation": eff_orientation,
        "margins": {
            "top": eff_top,
            "bottom": eff_bottom,
            "left": eff_left,
            "right": eff_right,
            "header": _get_int_attr(pg_mar, "header"),
            "footer": _get_int_attr(pg_mar, "footer"),
            "gutter": _get_int_attr(pg_mar, "gutter"),
        },
        "columns": {
            key: value
            for key, value in _extract_columns_info(target_sect_pr, namespace).items()
            if key in {"count", "spacing", "equal_width", "separator", "widths", "spaces"}
        },
        "document_grid": _extract_document_grid(target_sect_pr, namespace),
        "page_numbering": (
            {
                "start": fn["start"],
                "format": fn["format"],
            }
            if fn["has_pg_num_type"]
            else None
        ),
        "line_numbering": (
            {
                "start": ln["start"],
                "interval": ln["interval"],
                "distance": ln["distance"],
                "restart": ln["restart"],
            }
            if ln["has_ln_num_type"]
            else None
        ),
        "different_first_page": _is_title_pg_enabled(target_sect_pr, namespace),
        "vertical_alignment": va["vertical_alignment"],
        "page_borders": pb["page_borders"],
        "footnote_properties": fnp["footnote_properties"],
        "usable_area": {
            "width": usable_width,
            "height": usable_height,
        },
    }

    def verify_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset(replacements)):
            return "page_setup_staged_verification_failed"
        insp = inspect_docx(staged)
        if not insp.valid:
            return "page_setup_staged_verification_failed"
        try:
            s_graph = PackageGraph.from_file(staged)
            s_root = ET.fromstring(s_graph.read_part_bytes(_DOCUMENT))
            s_ns = _word_namespace(s_root)
            if s_ns != namespace:
                return "page_setup_staged_verification_failed"
            s_sections = _find_sections(s_root, s_ns)
            if len(s_sections) != len(sections):
                return "page_setup_staged_verification_failed"
            # Verify non-target sections are untouched
            for i, (orig_sec, staged_sec) in enumerate(zip(sections, s_sections), start=1):
                if i != request.section and _cell_structure_fingerprint(
                    orig_sec
                ) != _cell_structure_fingerprint(staged_sec):
                    return "page_setup_staged_verification_failed"
            # Verify target section
            target_sec = s_sections[request.section - 1]
            t_geom = _extract_page_geometry(target_sec, s_ns)
            if t_geom["width"] != eff_width or t_geom["height"] != eff_height:
                return "page_setup_staged_verification_failed"
            if request.orientation is not None and t_geom["orientation"] != eff_orientation:
                return "page_setup_staged_verification_failed"
            if (
                t_geom["top"] != eff_top
                or t_geom["bottom"] != eff_bottom
                or t_geom["left"] != eff_left
                or t_geom["right"] != eff_right
            ):
                return "page_setup_staged_verification_failed"
            if _extract_columns_info(target_sec, s_ns) != _extract_columns_info(
                target_sect_pr, namespace
            ):
                return "page_setup_staged_verification_failed"
            if _extract_document_grid(target_sec, s_ns) != _extract_document_grid(
                target_sect_pr, namespace
            ):
                return "page_setup_staged_verification_failed"
            if request.page_numbering is not None:
                t_num = target_sec.find(Q(s_ns, "pgNumType"))
                if request.page_numbering.clear:
                    if t_num is not None and (
                        t_num.get(Q(s_ns, "start")) is not None
                        or t_num.get("start") is not None
                        or t_num.get(Q(s_ns, "fmt")) is not None
                        or t_num.get("fmt") is not None
                    ):
                        return "page_setup_staged_verification_failed"
                else:
                    if request.page_numbering.clear_start:
                        if t_num is not None and (
                            t_num.get(Q(s_ns, "start")) is not None
                            or t_num.get("start") is not None
                        ):
                            return "page_setup_staged_verification_failed"
                    elif request.page_numbering.start is not None:
                        if t_num is None:
                            return "page_setup_staged_verification_failed"
                        act_s = t_num.get(Q(s_ns, "start")) or t_num.get("start")
                        if act_s != str(request.page_numbering.start):
                            return "page_setup_staged_verification_failed"

                    if request.page_numbering.clear_format:
                        if t_num is not None and (
                            t_num.get(Q(s_ns, "fmt")) is not None or t_num.get("fmt") is not None
                        ):
                            return "page_setup_staged_verification_failed"
                    elif request.page_numbering.format is not None:
                        if t_num is None:
                            return "page_setup_staged_verification_failed"
                        act_f = t_num.get(Q(s_ns, "fmt")) or t_num.get("fmt")
                        if act_f != request.page_numbering.format:
                            return "page_setup_staged_verification_failed"
            if request.line_numbering is not None:
                t_line = target_sec.find(Q(s_ns, "lnNumType"))
                if request.line_numbering.clear:
                    if t_line is not None:
                        return "page_setup_staged_verification_failed"
                else:
                    if t_line is None:
                        return "page_setup_staged_verification_failed"
                    for attr, value in (
                        ("start", request.line_numbering.start),
                        ("countBy", request.line_numbering.interval),
                        ("distance", request.line_numbering.distance),
                        ("restart", request.line_numbering.restart),
                    ):
                        if value is not None and (
                            t_line.get(Q(s_ns, attr)) or t_line.get(attr)
                        ) != str(value):
                            return "page_setup_staged_verification_failed"
            if request.different_first_page is not None:
                if _is_title_pg_enabled(target_sec, s_ns) != request.different_first_page:
                    return "page_setup_staged_verification_failed"
                if request.different_first_page:
                    if target_sec.find(Q(s_ns, "titlePg")) is None:
                        return "page_setup_staged_verification_failed"
                else:
                    if target_sec.find(Q(s_ns, "titlePg")) is not None:
                        return "page_setup_staged_verification_failed"
            else:
                if _is_title_pg_enabled(target_sec, s_ns) != _is_title_pg_enabled(
                    target_sect_pr, namespace
                ):
                    return "page_setup_staged_verification_failed"
            t_vertical = _extract_vertical_alignment_info(target_sec, s_ns)
            if request.vertical_alignment is not None:
                if (
                    t_vertical["has_duplicate_v_align"]
                    or t_vertical["vertical_alignment"] != request.vertical_alignment
                ):
                    return "page_setup_staged_verification_failed"
            elif t_vertical != existing_vertical:
                return "page_setup_staged_verification_failed"

            t_borders = _extract_page_borders_info(target_sec, s_ns)
            if request.page_borders is not None:
                if (
                    t_borders["has_duplicate_pg_borders"]
                    or t_borders["page_borders"] != request.page_borders.to_dict()
                ):
                    return "page_setup_staged_verification_failed"
            elif t_borders != existing_borders:
                return "page_setup_staged_verification_failed"

            t_footnote = _extract_footnote_properties_info(target_sec, s_ns)
            if request.footnote_properties is None:
                if t_footnote != existing_footnote:
                    return "page_setup_staged_verification_failed"
            elif request.footnote_properties.clear:
                if t_footnote["footnote_properties"] is not None:
                    return "page_setup_staged_verification_failed"
            else:
                before = existing_footnote["footnote_properties"] or dict.fromkeys(
                    ("format", "start", "restart", "position")
                )
                staged_footnote_pr = target_sec.find(Q(s_ns, "footnotePr"))
                if (
                    t_footnote["has_duplicate_footnote_pr"]
                    or staged_footnote_pr is None
                    or t_footnote["footnote_properties"]
                    != {**before, **request.footnote_properties.to_dict()}
                ):
                    return "page_setup_staged_verification_failed"
                footnote_ranks = [_footnote_pr_rank(child) for child in staged_footnote_pr]
                if footnote_ranks != sorted(footnote_ranks):
                    return "page_setup_staged_verification_failed"

            if (
                odd_even_plan is not None
                and request.different_odd_and_even_pages is not None
                and not _even_and_odd_headers_published(
                    s_graph, s_ns, odd_even_plan, request.different_odd_and_even_pages
                )
            ):
                return "page_setup_staged_verification_failed"
            # Verify children of target_sec are in monotonic OOXML schema order
            ranks = [_sect_pr_rank(child) for child in target_sec]
            if ranks != sorted(ranks):
                return "page_setup_staged_verification_failed"
        except (OSError, PackageError, ET.ParseError, KeyError, IndexError, ValueError):
            return "page_setup_staged_verification_failed"
        return None

    serialized = serialize_xml_with_preserved_qnames(root)
    replacements = {_DOCUMENT: serialized}
    additions: dict[str, bytes] = {}
    if odd_even_plan is not None:
        replacements.update(odd_even_plan.replacements)
        additions.update(odd_even_plan.additions)
        target_change_report["different_odd_and_even_pages"] = request.different_odd_and_even_pages

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        additions=additions,
        revisions=[],
        targets=[target_change_report],
        verify=verify_publication,
        failure_code="page_setup_staged_verification_failed",
        exists_code="page_setup_destination_occupied",
        unavailable_code="page_setup_publication_unavailable",
    )

    if not published.passed:
        return _page_setup_refusal(
            request, published.code or "page_setup_staged_verification_failed"
        )

    return PageSetupReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        target=request.target_dict(),
        page_setup=target_change_report,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )
