"""Source-bound native Word formatting discovery and authoring.

The authoring surface is the main document body (``word/document.xml``) and reachable
package parts. Formatting changes are atomic, source-bound, and strictly schema-valid:
run edits target bounded text offsets within paragraphs, while paragraph edits apply
to entire paragraphs.

Anything that cannot be resolved exactly is refused instead of guessed:
- Malformed requests (unknown keys, missing required fields, empty edits).
- Unknown or missing paragraph and character styles against ``word/styles.xml``.
- Unsafe revision (w:ins, w:del, etc.), comment, or SDT (content control) contexts.
- Overlapping run edits or invalid character ranges (negative, inverted, out-of-bounds).
- Text mismatch between expected_old_text and actual document text at target offsets.
- Stale structure fingerprints.
- Occupied destination paths or modified source packages.

Everything else is preserved: non-target parts are byte-identical to the source archive,
unedited paragraph and run properties are intact, and CT_RPr/CT_PPr element sequences
strictly follow OOXML schema order.

Discovery reports direct w14 text effects using the text_effects decoder, OpenType
values using their native scalar/list types, and scene_3d/props_3d as lossless XML
trees (expanded tag/attribute names, text, and ordered children). Unparsed effects
also retain that XML representation. Stored properties stay visible in legacy
compatibility modes, with ``w14_effects_active=False`` and ``compatibility_mode``
reported separately. Numeral-spacing creation and edits below compatibilityMode 14
are refused before publication; removal is allowed, and unrelated edits preserve
existing suppressed effects.
"""

from __future__ import annotations

import copy
import enum
import itertools
import json
import re
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree as ET
from xml.parsers import expat

from .authoring import _MAX_TEXT_CHARACTERS
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import (
    _DOCUMENT,
    ReviewReport,
    _clone_run,
    _mask_text,
    _paragraph_map,
    _ParagraphMap,
    _publish,
    _review_parts,
    _ReviewContextError,
    _sha,
    _sha256,
    _TextRun,
    _word_namespace,
)
from .table_properties import _VALID_BORDER_STYLES, BorderSpec

_STYLES = "word/styles.xml"
_REFUSED = "refused"
_PUBLISHED = "published"

_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_RUBY_ALIGNMENTS = frozenset(
    {"center", "distributeLetter", "distributeSpace", "left", "right", "rightVertical"}
)
_NUMERAL_SPACING = frozenset({"default", "proportional", "tabular"})
_LEGACY_ANIMATIONS = frozenset(
    {"none", "lights", "blinkBackground", "sparkle", "antsBlack", "antsRed", "shimmer"}
)
_EMPHASIS_MARKS = frozenset({"none", "dot", "comma", "circle", "underDot"})

_HEX_COLOR_PATTERN = re.compile(r"^[0-9a-fA-F]{6}\Z")
_VALID_UNDERLINE_STYLES = frozenset(
    {
        "single",
        "words",
        "double",
        "thick",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashedHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dashDotHeavy",
        "dotDotDash",
        "dashDotDotHeavy",
        "wave",
        "wavyHeavy",
        "wavyDouble",
        "none",
    }
)

_ST_HIGHLIGHT_COLORS: tuple[str, ...] = (
    "black",
    "blue",
    "cyan",
    "darkBlue",
    "darkCyan",
    "darkGray",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "green",
    "lightGray",
    "magenta",
    "none",
    "red",
    "white",
    "yellow",
)
_VALID_HIGHLIGHT_COLORS: frozenset[str] = frozenset(_ST_HIGHLIGHT_COLORS)
# Word UI and python-docx ``WD_COLOR_INDEX`` names resolve to their ``ST_HighlightColor``
# literal. ``ST_HighlightColor`` literals take precedence, so ``green`` stays ``green``
# (Word's "Bright Green") even though ``WD_COLOR_INDEX.GREEN`` writes ``darkGreen``.
_HIGHLIGHT_NAME_ALIASES: dict[str, str] = {
    "turquoise": "cyan",
    "pink": "magenta",
    "teal": "darkCyan",
    "violet": "darkMagenta",
    "brightgreen": "green",
    "bright_green": "green",
    "gray50": "darkGray",
    "gray_50": "darkGray",
    "gray25": "lightGray",
    "gray_25": "lightGray",
}
_HIGHLIGHT_CANONICAL_MAP: dict[str, str] = {
    **_HIGHLIGHT_NAME_ALIASES,
    **{c.lower(): c for c in _ST_HIGHLIGHT_COLORS},
}
# python-docx ``WD_COLOR_INDEX`` member names -> ``ST_HighlightColor``. Members resolve by
# name, so ``WD_COLOR_INDEX.GREEN`` writes ``darkGreen`` while ``"green"`` stays bright green.
_WD_COLOR_INDEX_HIGHLIGHTS: dict[str, str] = {
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
# python-docx ``BabelFish`` UI names of built-in styles whose ``w:name`` Word stores in
# lowercase: ``style="Heading 1"`` names the style whose ``w:name`` is ``heading 1``.
_STYLE_UI_NAMES: dict[str, str] = {
    "caption": "Caption",
    "footer": "Footer",
    "header": "Header",
    **{f"heading {level}": f"Heading {level}" for level in range(1, 10)},
}

_PARAGRAPH_BORDER_SIDES: tuple[str, ...] = (
    "top",
    "left",
    "bottom",
    "right",
    "between",
    "bar",
)

# Paragraph alignments keyed by their spelling folded to lowercase without underscores,
# mapped to the ECMA-376 ST_Jc value: the ST_Jc literals and the python-docx
# WD_PARAGRAPH_ALIGNMENT (WD_ALIGN_PARAGRAPH) names, where ``JUSTIFY`` writes ``both``.
_ALIGNMENT_ALIASES: dict[str, str] = {
    "left": "left",
    "center": "center",
    "right": "right",
    "both": "both",
    "justify": "both",
    "distribute": "distribute",
    "start": "start",
    "end": "end",
    "mediumkashida": "mediumKashida",
    "justifymed": "mediumKashida",
    "highkashida": "highKashida",
    "justifyhi": "highKashida",
    "lowkashida": "lowKashida",
    "justifylow": "lowKashida",
    "thaidistribute": "thaiDistribute",
    "thaijustify": "thaiDistribute",
}
_VALID_VERTICAL_ALIGNMENTS = frozenset({"baseline", "superscript", "subscript"})
# Paragraph line rules keyed by their spelling folded to lowercase without underscores,
# mapped to the canonical ``SpacingSpec.line_rule``: the ECMA-376 ST_LineSpacingRule
# values (``atLeast``) and the python-docx WD_LINE_SPACING names that name a rule alone
# (``AT_LEAST``, ``EXACTLY``, ``MULTIPLE``).
_LINE_RULE_ALIASES: dict[str, str] = {
    "auto": "auto",
    "multiple": "auto",
    "exact": "exact",
    "exactly": "exact",
    "atleast": "at_least",
}
# python-docx WD_LINE_SPACING names that fix the line height as well as the rule: each
# writes ``lineRule="auto"`` with the listed ``w:line`` in 240ths of a line.
_PRESET_LINE_HEIGHTS: dict[str, int] = {
    "single": 240,
    "onepointfive": 360,
    "double": 480,
}
# python-docx ParagraphFormat property names accepted for the spacing keys
_SPACING_KEY_ALIASES: dict[str, str] = {
    "space_before": "before",
    "space_after": "after",
    "line_spacing": "line",
    "line_spacing_rule": "line_rule",
}
# python-docx ParagraphFormat property names accepted for the indent keys
_INDENT_KEY_ALIASES: dict[str, str] = {"first_line_indent": "first_line"}
# ECMA-376 ST_TabJc values accepted for authoring (transitional left/right spellings)
_VALID_TAB_ALIGNMENTS = frozenset({"left", "center", "right", "decimal", "bar", "clear"})
# ECMA-376 ST_TabTlc leader characters keyed by their spelling folded to lowercase without
# underscores, including the python-docx WD_TAB_LEADER names (``DOTS``, ``MIDDLE_DOT``)
_TAB_LEADER_ALIASES: dict[str, str] = {
    "none": "none",
    "spaces": "none",
    "dot": "dot",
    "dots": "dot",
    "hyphen": "hyphen",
    "dashes": "hyphen",
    "underscore": "underscore",
    "lines": "underscore",
    "heavy": "heavy",
    "middledot": "middleDot",
}
# python-docx TabStop property names accepted for the tab stop keys
_TAB_KEY_ALIASES: dict[str, str] = {"alignment": "val", "position": "pos"}

# ECMA-376 CT_RPr sequence order
_RPR_ORDER: tuple[str, ...] = (
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
)
_RPR_RANK: dict[str, int] = {tag: i for i, tag in enumerate(_RPR_ORDER)}

# ECMA-376 CT_Fonts attribute declaration order
_RFONTS_ATTRIBUTE_ORDER: tuple[str, ...] = (
    "hint",
    "ascii",
    "hAnsi",
    "eastAsia",
    "cs",
    "asciiTheme",
    "hAnsiTheme",
    "eastAsiaTheme",
    "cstheme",
)

# ECMA-376 CT_PPrBase / CT_PPr sequence order
_PPR_ORDER: tuple[str, ...] = (
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
)
_PPR_RANK: dict[str, int] = {tag: i for i, tag in enumerate(_PPR_ORDER)}


def _local_name(tag: Any) -> str:
    return tag.split("}", 1)[-1] if isinstance(tag, str) else ""


def _flag(element: ET.Element | None) -> bool | None:
    if element is None:
        return None
    val = element.get(Q(_word_namespace_from_element(element), "val")) or element.get("val")
    if val is None:
        return True
    return val.strip().lower() not in {"0", "false", "off"}


def _word_namespace_from_element(element: ET.Element) -> str:
    if element.tag.startswith("{"):
        return element.tag[1:].split("}", 1)[0]
    return "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _fold_alias(value: str) -> str:
    return value.replace("_", "").lower()


def _conflicting_alias_keys(value: Mapping[str, Any], aliases: Mapping[str, str]) -> list[str]:
    """Keys that name one property twice, once canonically and once through an alias."""
    by_property: dict[str, list[str]] = {}
    for key in value:
        by_property.setdefault(aliases.get(key, key), []).append(key)
    return sorted(key for keys in by_property.values() if len(keys) > 1 for key in keys)


def _member_name(value: Any) -> Any:
    """Name of an Enum member (python-docx enumerations), otherwise the value unchanged."""
    return value.name if isinstance(value, enum.Enum) else value


def _canonical_highlight(value: Any, code: str) -> str:
    """Resolve an ST_HighlightColor literal, Word/python-docx name, or WD_COLOR_INDEX member."""
    if isinstance(value, enum.Enum):
        highlight = _WD_COLOR_INDEX_HIGHLIGHTS.get(value.name)
    elif isinstance(value, str):
        highlight = _HIGHLIGHT_CANONICAL_MAP.get(_fold_alias(value.strip()))
    else:
        highlight = None
    if highlight is None:
        raise ValueError(code)
    return highlight


def _canonical_alignment(value: Any, code: str) -> str:
    """Resolve an ST_Jc literal or python-docx alignment name/member to its ST_Jc value."""
    value = _member_name(value)
    alignment = _ALIGNMENT_ALIASES.get(_fold_alias(value)) if isinstance(value, str) else None
    if alignment is None:
        raise ValueError(code)
    return alignment


def _canonical_line_spacing(line: Any, line_rule: Any) -> tuple[Any, str]:
    """Resolve a line rule spelling to ``(line, canonical rule)``.

    WD_LINE_SPACING ``SINGLE``/``ONE_POINT_FIVE``/``DOUBLE`` resolve to ``auto`` with their
    fixed line height; an explicit ``line`` that disagrees with that height refuses.
    """
    name = _member_name(line_rule)
    folded = _fold_alias(name) if isinstance(name, str) else None
    preset = _PRESET_LINE_HEIGHTS.get(folded)
    if preset is not None:
        if line is not None and line != preset:
            raise ValueError("format_spacing_line_rule_invalid")
        return preset, "auto"
    rule = _LINE_RULE_ALIASES.get(folded)
    if rule is None:
        raise ValueError("format_spacing_line_rule_invalid")
    return line, rule


@dataclass(frozen=True)
class SpacingSpec:
    """Paragraph spacing specification in twips / line-rules.

    ``line_rule`` accepts ECMA-376 and python-docx spellings (strings or WD_LINE_SPACING
    members) and is stored as the canonical ``auto``, ``exact``, or ``at_least``;
    ``single``/``one_point_five``/``double`` also set ``line`` to 240/360/480.
    """

    before: int | None = None
    after: int | None = None
    line: int | None = None
    line_rule: str | None = None

    def __post_init__(self) -> None:
        if self.line_rule is not None:
            line, line_rule = _canonical_line_spacing(self.line, self.line_rule)
            object.__setattr__(self, "line", line)
            object.__setattr__(self, "line_rule", line_rule)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {}
        if self.before is not None:
            result["before"] = self.before
        if self.after is not None:
            result["after"] = self.after
        if self.line is not None:
            result["line"] = self.line
        if self.line_rule is not None:
            result["line_rule"] = self.line_rule
        return result


@dataclass(frozen=True)
class IndentSpec:
    """Paragraph indentation specification in twips.

    A negative ``first_line`` is python-docx's hanging indent: it is stored as
    ``hanging`` with the absolute value, and refuses alongside an explicit ``hanging``.
    """

    left: int | None = None
    right: int | None = None
    first_line: int | None = None
    hanging: int | None = None
    start: int | None = None
    end: int | None = None

    def __post_init__(self) -> None:
        first_line = self.first_line
        if isinstance(first_line, int) and not isinstance(first_line, bool) and first_line < 0:
            if self.hanging is not None:
                raise ValueError("format_indent_first_line_hanging_conflict")
            object.__setattr__(self, "hanging", -first_line)
            object.__setattr__(self, "first_line", None)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {}
        if self.left is not None:
            result["left"] = self.left
        if self.right is not None:
            result["right"] = self.right
        if self.first_line is not None:
            result["first_line"] = self.first_line
        if self.hanging is not None:
            result["hanging"] = self.hanging
        if self.start is not None:
            result["start"] = self.start
        if self.end is not None:
            result["end"] = self.end
        return result


@dataclass(frozen=True)
class ParagraphBordersSpec:
    """Paragraph border specification across supported sides."""

    top: BorderSpec | Mapping[str, Any] | Literal[False] | None = None
    left: BorderSpec | Mapping[str, Any] | Literal[False] | None = None
    bottom: BorderSpec | Mapping[str, Any] | Literal[False] | None = None
    right: BorderSpec | Mapping[str, Any] | Literal[False] | None = None
    between: BorderSpec | Mapping[str, Any] | Literal[False] | None = None
    bar: BorderSpec | Mapping[str, Any] | Literal[False] | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for side in _PARAGRAPH_BORDER_SIDES:
            val = getattr(self, side)
            if val is not None:
                if isinstance(val, BorderSpec):
                    data[side] = val.to_dict()
                elif isinstance(val, Mapping):
                    data[side] = dict(val)
                elif val is False:
                    data[side] = False
        return data


@dataclass(frozen=True)
class TabStopSpec:
    """One custom paragraph tab stop; ``pos`` is measured in twips from the start indent.

    ``val`` and ``leader`` accept python-docx WD_TAB_ALIGNMENT / WD_TAB_LEADER names and
    are stored as their ECMA-376 ``w:tab`` attribute values.
    """

    val: str
    pos: int
    leader: str | None = None

    def __post_init__(self) -> None:
        val = self.val.lower() if isinstance(self.val, str) else None
        if val not in _VALID_TAB_ALIGNMENTS:
            raise ValueError("format_tabs_invalid")
        if isinstance(self.pos, bool) or not isinstance(self.pos, int) or self.pos < 0:
            raise ValueError("format_tabs_invalid")
        leader = self.leader
        if leader is not None:
            leader = (
                _TAB_LEADER_ALIASES.get(_fold_alias(leader)) if isinstance(leader, str) else None
            )
            if leader is None:
                raise ValueError("format_tabs_invalid")
        object.__setattr__(self, "val", val)
        object.__setattr__(self, "leader", leader)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"val": self.val, "pos": self.pos}
        if self.leader is not None:
            result["leader"] = self.leader
        return result


# ECMA-376 17.3.1.11 CT_FramePr attributes keyed by their snake_case request names.
_FRAME_ATTRIBUTES: dict[str, str] = {
    "drop_cap": "dropCap",
    "lines": "lines",
    "w": "w",
    "h": "h",
    "h_space": "hSpace",
    "v_space": "vSpace",
    "wrap": "wrap",
    "h_anchor": "hAnchor",
    "v_anchor": "vAnchor",
    "x": "x",
    "x_align": "xAlign",
    "y": "y",
    "y_align": "yAlign",
    "h_rule": "hRule",
    "anchor_lock": "anchorLock",
}
_FRAME_INTEGER_FIELDS = frozenset({"lines", "w", "h", "h_space", "v_space", "x", "y"})
_FRAME_CHOICES: dict[str, frozenset[str]] = {
    "drop_cap": frozenset({"none", "drop", "margin"}),
    "wrap": frozenset({"auto", "notBeside", "around", "tight", "through", "none"}),
    "h_anchor": frozenset({"text", "margin", "page"}),
    "v_anchor": frozenset({"text", "margin", "page"}),
    "x_align": frozenset({"left", "center", "right", "inside", "outside"}),
    "y_align": frozenset({"inline", "top", "center", "bottom", "inside", "outside"}),
    "h_rule": frozenset({"auto", "exact", "atLeast"}),
}
_FRAME_INT_MAX = 2**31 - 1
# ECMA-376 17.3.3.1 w:br @w:type and @w:clear values (ST_BrType, ST_BrClear)
_BREAK_TYPES = frozenset({"page", "column", "textWrapping"})
_BREAK_CLEARS = frozenset({"none", "left", "right", "all"})


@dataclass(frozen=True)
class FrameSpec:
    """Paragraph frame (w:framePr, ECMA-376 17.3.1.11) in twips unless stated.

    Fields keep the ECMA attribute meaning under snake_case names: ``drop_cap``
    (none, drop, margin), ``lines`` (drop-cap height, 1 to 10), ``w`` and ``h``
    (extent), ``h_space`` and ``v_space`` (text distance), ``wrap`` (auto, notBeside,
    around, tight, through, none), ``h_anchor`` and ``v_anchor`` (text, margin, page),
    ``x`` and ``y`` (signed offsets), ``x_align`` (left, center, right, inside, outside),
    ``y_align`` (inline, top, center, bottom, inside, outside), ``h_rule`` (auto, exact,
    atLeast) and ``anchor_lock``. None leaves an attribute unchanged on an existing
    frame; at least one attribute must be set.
    """

    drop_cap: str | None = None
    lines: int | None = None
    w: int | None = None
    h: int | None = None
    h_space: int | None = None
    v_space: int | None = None
    wrap: str | None = None
    h_anchor: str | None = None
    v_anchor: str | None = None
    x: int | None = None
    x_align: str | None = None
    y: int | None = None
    y_align: str | None = None
    h_rule: str | None = None
    anchor_lock: bool | None = None

    def __post_init__(self) -> None:
        present = 0
        for field in _FRAME_ATTRIBUTES:
            value = getattr(self, field)
            if value is None:
                continue
            present += 1
            if field == "anchor_lock":
                valid = isinstance(value, bool)
            elif field in _FRAME_INTEGER_FIELDS:
                low, high = {
                    "lines": (1, 10),
                    "x": (-_FRAME_INT_MAX - 1, _FRAME_INT_MAX),
                    "y": (-_FRAME_INT_MAX - 1, _FRAME_INT_MAX),
                }.get(field, (0, _FRAME_INT_MAX))
                valid = (
                    isinstance(value, int) and not isinstance(value, bool) and low <= value <= high
                )
            else:
                valid = isinstance(value, str) and value in _FRAME_CHOICES[field]
            if not valid:
                raise ValueError(f"format_frame_{field}_invalid")
        if not present:
            raise ValueError("format_frame_empty_invalid")

    def to_dict(self) -> dict[str, Any]:
        return {
            field: getattr(self, field)
            for field in _FRAME_ATTRIBUTES
            if getattr(self, field) is not None
        }


def _validate_frame_spec(value: Any) -> FrameSpec | Literal[False]:
    if value is False or isinstance(value, FrameSpec):
        return value
    if not isinstance(value, Mapping):
        raise TypeError("format_frame_invalid")
    unknown = set(value) - set(_FRAME_ATTRIBUTES)
    if unknown:
        raise ValueError(f"format_frame_unknown_keys: {sorted(unknown)}")
    return FrameSpec(**value)


def _frame_value(value: bool | int | str) -> str:
    if isinstance(value, bool):
        return "1" if value else "0"
    return str(value)


@dataclass(frozen=True)
class RubySpec:
    """ECMA-376 17.3.3 phonetic guide; all sizes/raise are half-points.

    ``base_text`` must equal the selected text. ``language`` authors the required
    rubyPr/lid; base_font_size authors hpsBaseText, not the surrounding run size.
    Guide text is an annotation, never included in the formatting character range.
    """

    base_text: str
    guide_text: str
    alignment: str = "center"
    raise_: int = 0
    font_size: int = 12
    base_font_size: int = 24
    language: str = "en-US"

    def __post_init__(self) -> None:
        for text in (self.base_text, self.guide_text, self.language):
            if (
                not isinstance(text, str)
                or not text
                or len(text) > _MAX_TEXT_CHARACTERS
                or any(
                    not (
                        c in "\t\n\r"
                        or 0x20 <= ord(c) <= 0xD7FF
                        or 0xE000 <= ord(c) <= 0xFFFD
                        or 0x10000 <= ord(c) <= 0x10FFFF
                    )
                    for c in text
                )
            ):
                raise ValueError("format_ruby_text_invalid")
        if not isinstance(self.alignment, str) or self.alignment not in _RUBY_ALIGNMENTS:
            raise ValueError("format_ruby_alignment_invalid")
        for value, minimum in (
            (self.raise_, 0),
            (self.font_size, 1),
            (self.base_font_size, 1),
        ):
            if (
                isinstance(value, bool)
                or not isinstance(value, int)
                or not minimum <= value <= 32767
            ):
                raise ValueError("format_ruby_size_invalid")

    def to_dict(self) -> dict[str, Any]:
        return {
            "base_text": self.base_text,
            "guide_text": self.guide_text,
            "alignment": self.alignment,
            "raise": self.raise_,
            "font_size": self.font_size,
            "base_font_size": self.base_font_size,
            "language": self.language,
        }


def _ruby_spec(value: Any) -> RubySpec:
    if isinstance(value, RubySpec):
        return value
    if not isinstance(value, Mapping) or set(value) - {
        "base_text",
        "guide_text",
        "alignment",
        "raise",
        "font_size",
        "base_font_size",
        "language",
    }:
        raise ValueError("format_ruby_invalid")
    if not {"base_text", "guide_text"}.issubset(value):
        raise ValueError("format_ruby_invalid")
    return RubySpec(
        base_text=value["base_text"],
        guide_text=value["guide_text"],
        alignment=value.get("alignment", "center"),
        raise_=value.get("raise", 0),
        font_size=value.get("font_size", 12),
        base_font_size=value.get("base_font_size", 24),
        language=value.get("language", "en-US"),
    )


@dataclass(frozen=True)
class RunFormatEdit:
    """Selective formatting anchored to a source-bound base-text character range.

    ``ruby`` wraps one selected run in a phonetic guide; guide sizes and raise
    are half-points (ECMA-376 Part 1, 17.3.3). Existing ruby must be selected
    whole; partial annotations refuse rather than splitting the guide.
    ``numeral_spacing`` authors Office 2010 w14:numSpacing: ``default``,
    ``proportional`` or ``tabular``. ``False`` removes the stored property, including
    in compatibility modes below 14; ``None`` preserves it.
    ``legacy_animation`` stores w:effect: ``none``, ``lights``, ``blinkBackground``,
    ``sparkle``, ``antsBlack``, ``antsRed`` or ``shimmer``.
    ``emphasis_mark`` stores w:em: ``none``, ``dot``, ``comma``, ``circle`` or ``underDot``.
    These legacy properties do not enable animation or emphasis-mark painting.
    None leaves a property unchanged; ``"none"`` explicitly disables it.
    All these properties are returned by discover_format_targets.
    ``break_type`` (``page``, ``column`` or ``textWrapping``) and ``break_clear`` (``none``,
    ``left``, ``right`` or ``all``, textWrapping only) insert one w:br immediately after the
    selected text. A run holding w:br or w:lastRenderedPageBreak is edited whole: a range
    that covers only part of it refuses with ``format_break_partial_range_unsupported``.
    discover_format_targets lists breaks and cached page breaks with their text offsets.
    """

    paragraph: int
    start: int
    end: int
    expected_old_text: str
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    bold: bool | None = None
    italic: bool | None = None
    underline: str | bool | None = None
    strike: bool | None = None
    color: str | None = None
    font_size: int | None = None
    vertical_alignment: str | None = None
    style: str | None = None
    font_face: str | None = None
    font_face_east_asia: str | None = None
    highlight: str | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None
    hidden: bool | None = None
    ruby: RubySpec | None = None
    numeral_spacing: str | Literal[False] | None = None
    legacy_animation: str | None = None
    emphasis_mark: str | None = None
    break_type: str | None = None
    break_clear: str | None = None

    def __post_init__(self) -> None:
        if self.ruby is not None:
            object.__setattr__(self, "ruby", _ruby_spec(self.ruby))
        if (
            self.numeral_spacing is not None
            and self.numeral_spacing is not False
            and (
                not isinstance(self.numeral_spacing, str)
                or self.numeral_spacing not in _NUMERAL_SPACING
            )
        ):
            raise ValueError("format_numeral_spacing_invalid")
        for name, allowed in (
            ("legacy_animation", _LEGACY_ANIMATIONS),
            ("emphasis_mark", _EMPHASIS_MARKS),
        ):
            value = getattr(self, name)
            if value is not None and (not isinstance(value, str) or value not in allowed):
                raise ValueError(f"format_{name}_invalid")
        if self.highlight is not None:
            if isinstance(self.highlight, bool) or not isinstance(self.highlight, str):
                raise ValueError("format_highlight_invalid")
            stripped = self.highlight.strip()
            if stripped.lower() not in _HIGHLIGHT_CANONICAL_MAP:
                raise ValueError("format_highlight_invalid")
            object.__setattr__(self, "highlight", _HIGHLIGHT_CANONICAL_MAP[stripped.lower()])
        if self.small_caps is not None and not isinstance(self.small_caps, bool):
            raise ValueError("format_small_caps_invalid")
        if self.all_caps is not None and not isinstance(self.all_caps, bool):
            raise ValueError("format_all_caps_invalid")
        if self.hidden is not None and not isinstance(self.hidden, bool):
            raise ValueError("format_hidden_invalid")
        if self.font_face_east_asia is not None:
            _validate_font_face_east_asia(self.font_face_east_asia)
        if self.break_clear is not None and self.break_type is None:
            object.__setattr__(self, "break_type", "textWrapping")
        if self.break_type is not None and self.break_type not in _BREAK_TYPES:
            raise ValueError("format_break_type_invalid")
        if self.break_clear is not None and self.break_clear not in _BREAK_CLEARS:
            raise ValueError("format_break_clear_invalid")
        if self.break_clear is not None and self.break_type != "textWrapping":
            raise ValueError("format_break_clear_invalid")
        if self.break_type is not None and self.ruby is not None:
            raise ValueError("format_break_ruby_invalid")

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "part": self.part,
            "paragraph": self.paragraph,
            "start": self.start,
            "end": self.end,
            "expected_old_text": self.expected_old_text,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
        }
        if self.bold is not None:
            result["bold"] = self.bold
        if self.italic is not None:
            result["italic"] = self.italic
        if self.underline is not None:
            result["underline"] = self.underline
        if self.strike is not None:
            result["strike"] = self.strike
        if self.color is not None:
            result["color"] = self.color
        if self.font_size is not None:
            result["font_size"] = self.font_size
        if self.vertical_alignment is not None:
            result["vertical_alignment"] = self.vertical_alignment
        if self.style is not None:
            result["style"] = self.style
        if self.font_face is not None:
            result["font_face"] = self.font_face
        if self.font_face_east_asia is not None:
            result["font_face_east_asia"] = self.font_face_east_asia
        if self.highlight is not None:
            result["highlight"] = self.highlight
        if self.small_caps is not None:
            result["small_caps"] = self.small_caps
        if self.all_caps is not None:
            result["all_caps"] = self.all_caps
        if self.hidden is not None:
            result["hidden"] = self.hidden
        if self.ruby is not None:
            result["ruby"] = self.ruby.to_dict()
        if self.numeral_spacing is not None:
            result["numeral_spacing"] = self.numeral_spacing
        if self.legacy_animation is not None:
            result["legacy_animation"] = self.legacy_animation
        if self.emphasis_mark is not None:
            result["emphasis_mark"] = self.emphasis_mark
        if self.break_type is not None:
            result["break_type"] = self.break_type
        if self.break_clear is not None:
            result["break_clear"] = self.break_clear
        return result


def _fold_flag_alias(canonical: Any, alias: Any, code: str) -> bool | None:
    """Fold a python-docx alias flag into its canonical flag; disagreeing values refuse."""
    for flag in (canonical, alias):
        if flag is not None and not isinstance(flag, bool):
            raise ValueError(f"{code}_invalid")
    if alias is None:
        return canonical
    if canonical is not None and canonical is not alias:
        raise ValueError(f"{code}_conflict")
    return alias


@dataclass(frozen=True)
class ParagraphFormatEdit:
    """Selective paragraph formatting change anchored to a source-bound paragraph.

    ``keep_with_next`` and ``keep_together`` are python-docx names for ``keep_next``
    (``w:keepNext``) and ``keep_lines`` (``w:keepLines``); they fold into the canonical
    fields at construction and read back as ``None``. ``alignment`` accepts ST_Jc values
    and python-docx WD_PARAGRAPH_ALIGNMENT names or members in any case and is stored as
    the ST_Jc value (``JUSTIFY`` becomes ``both``).

    ``suppress_auto_hyphens`` and ``suppress_line_numbers`` author w:suppressAutoHyphens and
    w:suppressLineNumbers; ``frame`` authors w:framePr from a ``FrameSpec`` or its mapping
    (unspecified attributes of an existing frame are kept) and ``False`` removes the frame.
    """

    paragraph: int
    structure_fingerprint_sha256: str
    part: str = _DOCUMENT
    style: str | None = None
    alignment: str | None = None
    spacing: SpacingSpec | None = None
    indent: IndentSpec | None = None
    keep_next: bool | None = None
    keep_lines: bool | None = None
    outline_level: int | None = None
    shading: str | Literal[False] | None = None
    page_break_before: bool | None = None
    borders: ParagraphBordersSpec | Mapping[str, Any] | Literal[False] | None = None
    widow_control: bool | None = None
    tabs: Sequence[TabStopSpec | Mapping[str, Any]] | None = None
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    suppress_auto_hyphens: bool | None = None
    suppress_line_numbers: bool | None = None
    frame: FrameSpec | Mapping[str, Any] | Literal[False] | None = None

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "keep_next",
            _fold_flag_alias(self.keep_next, self.keep_with_next, "format_keep_next"),
        )
        object.__setattr__(
            self,
            "keep_lines",
            _fold_flag_alias(self.keep_lines, self.keep_together, "format_keep_lines"),
        )
        object.__setattr__(self, "keep_with_next", None)
        object.__setattr__(self, "keep_together", None)
        if self.alignment is not None:
            object.__setattr__(
                self,
                "alignment",
                _canonical_alignment(self.alignment, "format_alignment_invalid"),
            )
        if self.page_break_before is not None and not isinstance(self.page_break_before, bool):
            raise ValueError("format_page_break_before_invalid")
        if self.widow_control is not None and not isinstance(self.widow_control, bool):
            raise ValueError("format_widow_control_invalid")
        if self.tabs is not None:
            object.__setattr__(self, "tabs", _validate_tabs(self.tabs))
        if self.shading is not None:
            if isinstance(self.shading, bool):
                if self.shading is not False:
                    raise ValueError("format_shading_invalid")
            elif isinstance(self.shading, str):
                if not _HEX_COLOR_PATTERN.match(self.shading):
                    raise ValueError("format_shading_invalid")
            else:
                raise ValueError("format_shading_invalid")
        if self.borders is not None:
            normalized = _validate_borders_spec(self.borders)
            object.__setattr__(self, "borders", normalized)
        for name in ("suppress_auto_hyphens", "suppress_line_numbers"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, bool):
                raise ValueError(f"format_{name}_invalid")
        if self.frame is not None:
            object.__setattr__(self, "frame", _validate_frame_spec(self.frame))

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "part": self.part,
            "paragraph": self.paragraph,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
        }
        if self.style is not None:
            result["style"] = self.style
        if self.alignment is not None:
            result["alignment"] = self.alignment
        if self.spacing is not None:
            result["spacing"] = self.spacing.to_dict()
        if self.indent is not None:
            result["indent"] = self.indent.to_dict()
        if self.keep_next is not None:
            result["keep_next"] = self.keep_next
        if self.keep_lines is not None:
            result["keep_lines"] = self.keep_lines
        if self.outline_level is not None:
            result["outline_level"] = self.outline_level
        if self.shading is not None:
            result["shading"] = self.shading
        if self.borders is not None:
            result["borders"] = (
                self.borders.to_dict()
                if isinstance(self.borders, ParagraphBordersSpec)
                else self.borders
            )
        if self.page_break_before is not None:
            result["page_break_before"] = self.page_break_before
        if self.widow_control is not None:
            result["widow_control"] = self.widow_control
        if self.tabs is not None:
            result["tabs"] = [_validate_tab_stop(stop).to_dict() for stop in self.tabs]
        if self.suppress_auto_hyphens is not None:
            result["suppress_auto_hyphens"] = self.suppress_auto_hyphens
        if self.suppress_line_numbers is not None:
            result["suppress_line_numbers"] = self.suppress_line_numbers
        if self.frame is not None:
            result["frame"] = self.frame if self.frame is False else self.frame.to_dict()
        return result


@dataclass(frozen=True)
class FormatRequest:
    """Strict request for source-bound paragraph and run formatting changes."""

    source: Path
    source_sha256: str
    destination: Path
    run_edits: tuple[RunFormatEdit, ...] = ()
    paragraph_edits: tuple[ParagraphFormatEdit, ...] = ()


@dataclass(frozen=True)
class FormatReport:
    """Content-safe public receipt for one atomic formatting authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    run_edits: list[dict[str, Any]]
    paragraph_edits: list[dict[str, Any]]
    changed_targets: list[dict[str, Any]]

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "run_edits": self.run_edits,
            "paragraph_edits": self.paragraph_edits,
            "changed_targets": self.changed_targets,
        }


def _validate_font_face_east_asia(value: Any) -> str:
    if not isinstance(value, str) or not value or len(value) > 256:
        raise ValueError("format_font_face_east_asia_invalid")
    return value


def _validate_run_edit_dict(value: Mapping[str, Any]) -> RunFormatEdit:
    allowed_keys = {
        "paragraph",
        "start",
        "end",
        "expected_old_text",
        "structure_fingerprint_sha256",
        "part",
        "bold",
        "italic",
        "underline",
        "strike",
        "color",
        "font_size",
        "vertical_alignment",
        "style",
        "font_face",
        "font_face_east_asia",
        "highlight",
        "small_caps",
        "all_caps",
        "hidden",
        "ruby",
        "numeral_spacing",
        "legacy_animation",
        "emphasis_mark",
        "break_type",
        "break_clear",
    }
    unknown_keys = set(value) - allowed_keys
    if unknown_keys:
        raise ValueError(f"format_run_edit_unknown_keys: {sorted(unknown_keys)}")

    required_keys = {
        "paragraph",
        "start",
        "end",
        "expected_old_text",
        "structure_fingerprint_sha256",
    }
    missing = required_keys - set(value)
    if missing:
        raise ValueError(f"format_run_edit_missing_keys: {sorted(missing)}")

    paragraph = value["paragraph"]
    if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 1:
        raise ValueError("format_paragraph_invalid")

    start = value["start"]
    end = value["end"]
    if isinstance(start, bool) or not isinstance(start, int) or start < 0:
        raise ValueError("format_range_invalid")
    if isinstance(end, bool) or not isinstance(end, int) or end < 0 or end <= start:
        raise ValueError("format_range_invalid")

    expected_old_text = value["expected_old_text"]
    if not isinstance(expected_old_text, str) or len(expected_old_text) > _MAX_TEXT_CHARACTERS:
        raise ValueError("format_expected_old_text_invalid")

    structure_fingerprint = _sha(
        value["structure_fingerprint_sha256"], "format_structure_fingerprint_sha256"
    )

    part = value.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part or not part.endswith(".xml"):
        raise ValueError("format_part_invalid")

    # Format fields
    bold = value.get("bold")
    if bold is not None and not isinstance(bold, bool):
        raise ValueError("format_bold_invalid")

    italic = value.get("italic")
    if italic is not None and not isinstance(italic, bool):
        raise ValueError("format_italic_invalid")

    underline = value.get("underline")
    if underline is not None and not (
        isinstance(underline, bool)
        or (isinstance(underline, str) and underline in _VALID_UNDERLINE_STYLES)
    ):
        raise ValueError("format_underline_invalid")

    strike = value.get("strike")
    if strike is not None and not isinstance(strike, bool):
        raise ValueError("format_strike_invalid")

    color = value.get("color")
    if color is not None and (
        not isinstance(color, str)
        or (color.lower() != "auto" and not _HEX_COLOR_PATTERN.match(color))
    ):
        raise ValueError("format_color_invalid")

    font_size = value.get("font_size")
    if font_size is not None and (
        isinstance(font_size, bool) or not isinstance(font_size, int) or font_size <= 0
    ):
        raise ValueError("format_font_size_invalid")

    vertical_alignment = value.get("vertical_alignment")
    if vertical_alignment is not None and (
        not isinstance(vertical_alignment, str)
        or vertical_alignment not in _VALID_VERTICAL_ALIGNMENTS
    ):
        raise ValueError("format_vertical_alignment_invalid")
    style = value.get("style")
    if style is not None and (not isinstance(style, str) or not style or len(style) > 256):
        raise ValueError("format_style_invalid")

    font_face = value.get("font_face")
    if font_face is not None and (
        not isinstance(font_face, str) or not font_face or len(font_face) > 256
    ):
        raise ValueError("format_font_face_invalid")
    font_face_east_asia = value.get("font_face_east_asia")
    if font_face_east_asia is not None:
        _validate_font_face_east_asia(font_face_east_asia)
    highlight = value.get("highlight")
    if highlight is not None:
        if isinstance(highlight, bool) or not isinstance(highlight, str):
            raise ValueError("format_highlight_invalid")
        stripped = highlight.strip()
        if stripped.lower() not in _HIGHLIGHT_CANONICAL_MAP:
            raise ValueError("format_highlight_invalid")
        highlight = _HIGHLIGHT_CANONICAL_MAP[stripped.lower()]

    small_caps = value.get("small_caps")
    if small_caps is not None and not isinstance(small_caps, bool):
        raise ValueError("format_small_caps_invalid")

    all_caps = value.get("all_caps")
    if all_caps is not None and not isinstance(all_caps, bool):
        raise ValueError("format_all_caps_invalid")

    hidden = value.get("hidden")
    if hidden is not None and not isinstance(hidden, bool):
        raise ValueError("format_hidden_invalid")

    # Verify at least one format property is specified
    format_props = (
        bold,
        italic,
        underline,
        strike,
        color,
        font_size,
        vertical_alignment,
        style,
        font_face,
        font_face_east_asia,
        highlight,
        small_caps,
        all_caps,
        hidden,
        value.get("ruby"),
        value.get("numeral_spacing"),
        value.get("legacy_animation"),
        value.get("emphasis_mark"),
        value.get("break_type"),
        value.get("break_clear"),
    )
    if all(p is None for p in format_props):
        raise ValueError("format_run_edit_noop_invalid")

    return RunFormatEdit(
        paragraph=paragraph,
        start=start,
        end=end,
        expected_old_text=expected_old_text,
        structure_fingerprint_sha256=structure_fingerprint,
        part=part,
        bold=bold,
        italic=italic,
        underline=underline,
        strike=strike,
        color=color,
        font_size=font_size,
        vertical_alignment=vertical_alignment,
        style=style,
        font_face=font_face,
        font_face_east_asia=font_face_east_asia,
        highlight=highlight,
        small_caps=small_caps,
        all_caps=all_caps,
        hidden=hidden,
        ruby=_ruby_spec(value["ruby"]) if value.get("ruby") is not None else None,
        numeral_spacing=value.get("numeral_spacing"),
        legacy_animation=value.get("legacy_animation"),
        emphasis_mark=value.get("emphasis_mark"),
        break_type=value.get("break_type"),
        break_clear=value.get("break_clear"),
    )


def _validate_spacing_spec(value: Any) -> SpacingSpec:
    if isinstance(value, SpacingSpec):
        return value
    if not isinstance(value, Mapping):
        raise TypeError("format_spacing_invalid")
    allowed = {"before", "after", "line", "line_rule", *_SPACING_KEY_ALIASES}
    unknown = set(value) - allowed
    if unknown:
        raise ValueError(f"format_spacing_unknown_keys: {sorted(unknown)}")
    conflicts = _conflicting_alias_keys(value, _SPACING_KEY_ALIASES)
    if conflicts:
        raise ValueError(f"format_spacing_alias_conflict: {conflicts}")
    resolved = {_SPACING_KEY_ALIASES.get(key, key): item for key, item in value.items()}
    before = resolved.get("before")
    after = resolved.get("after")
    line = resolved.get("line")
    line_rule = resolved.get("line_rule")

    for name, val in (("before", before), ("after", after), ("line", line)):
        if val is not None and (isinstance(val, bool) or not isinstance(val, int) or val < 0):
            raise ValueError(f"format_spacing_{name}_invalid")

    if before is None and after is None and line is None and line_rule is None:
        raise ValueError("format_spacing_empty_invalid")

    return SpacingSpec(before=before, after=after, line=line, line_rule=line_rule)


def _validate_indent_spec(value: Any) -> IndentSpec:
    if isinstance(value, IndentSpec):
        return value
    if not isinstance(value, Mapping):
        raise TypeError("format_indent_invalid")
    allowed = {
        "left",
        "right",
        "first_line",
        "hanging",
        "start",
        "end",
        *_INDENT_KEY_ALIASES,
    }
    unknown = set(value) - allowed
    if unknown:
        raise ValueError(f"format_indent_unknown_keys: {sorted(unknown)}")
    conflicts = _conflicting_alias_keys(value, _INDENT_KEY_ALIASES)
    if conflicts:
        raise ValueError(f"format_indent_alias_conflict: {conflicts}")
    resolved = {_INDENT_KEY_ALIASES.get(key, key): item for key, item in value.items()}

    left = resolved.get("left")
    right = resolved.get("right")
    first_line = resolved.get("first_line")
    hanging = resolved.get("hanging")
    start = resolved.get("start")
    end = resolved.get("end")

    for name, val in (
        ("left", left),
        ("right", right),
        ("first_line", first_line),
        ("hanging", hanging),
        ("start", start),
        ("end", end),
    ):
        if val is None:
            continue
        # A negative first_line is a hanging indent; IndentSpec folds it into ``hanging``.
        if isinstance(val, bool) or not isinstance(val, int) or (val < 0 and name != "first_line"):
            raise ValueError(f"format_indent_{name}_invalid")

    if all(v is None for v in (left, right, first_line, hanging, start, end)):
        raise ValueError("format_indent_empty_invalid")

    return IndentSpec(
        left=left,
        right=right,
        first_line=first_line,
        hanging=hanging,
        start=start,
        end=end,
    )


def _validate_border_side_value(val: Any) -> BorderSpec | Literal[False]:
    if val is False:
        return False
    if isinstance(val, BorderSpec):
        val_style = val.val
        if not isinstance(val_style, str) or val_style not in _VALID_BORDER_STYLES:
            raise ValueError("format_borders_invalid")
        if val.sz is not None and (
            isinstance(val.sz, bool) or not isinstance(val.sz, int) or val.sz < 0 or val.sz > 96
        ):
            raise ValueError("format_borders_invalid")
        if val.space is not None and (
            isinstance(val.space, bool)
            or not isinstance(val.space, int)
            or val.space < 0
            or val.space > 64
        ):
            raise ValueError("format_borders_invalid")
        if val.color is not None and (
            not isinstance(val.color, str)
            or (val.color != "auto" and not _HEX_COLOR_PATTERN.match(val.color))
        ):
            raise ValueError("format_borders_invalid")
        return val
    if isinstance(val, Mapping):
        allowed = {"val", "style", "sz", "size", "space", "color"}
        if set(val) - allowed:
            raise ValueError("format_borders_invalid")
        val_style = val.get("val", val.get("style", "single"))
        if not isinstance(val_style, str) or val_style not in _VALID_BORDER_STYLES:
            raise ValueError("format_borders_invalid")
        sz = val.get("sz", val.get("size", 4))
        if sz is not None and (
            isinstance(sz, bool) or not isinstance(sz, int) or sz < 0 or sz > 96
        ):
            raise ValueError("format_borders_invalid")
        space = val.get("space", 0)
        if space is not None and (
            isinstance(space, bool) or not isinstance(space, int) or space < 0 or space > 64
        ):
            raise ValueError("format_borders_invalid")
        color = val.get("color", "auto")
        if color is not None and (
            not isinstance(color, str) or (color != "auto" and not _HEX_COLOR_PATTERN.match(color))
        ):
            raise ValueError("format_borders_invalid")
        return BorderSpec(val=val_style, sz=sz, space=space, color=color)
    raise ValueError("format_borders_invalid")


def _validate_borders_spec(value: Any) -> ParagraphBordersSpec | Literal[False]:
    if value is False:
        return False
    if value is True:
        raise ValueError("format_borders_invalid")
    if isinstance(value, ParagraphBordersSpec):
        sides: dict[str, BorderSpec | Literal[False] | None] = {}
        all_none = True
        for side in _PARAGRAPH_BORDER_SIDES:
            v = getattr(value, side)
            if v is not None:
                all_none = False
                sides[side] = _validate_border_side_value(v)
            else:
                sides[side] = None
        if all_none:
            raise ValueError("format_borders_invalid")
        return ParagraphBordersSpec(**sides)
    if isinstance(value, Mapping):
        if not value:
            raise ValueError("format_borders_invalid")
        unknown = set(value) - set(_PARAGRAPH_BORDER_SIDES)
        if unknown:
            raise ValueError("format_borders_invalid")
        sides = {}
        for side in _PARAGRAPH_BORDER_SIDES:
            if side in value:
                v = value[side]
                if v is None:
                    sides[side] = False
                else:
                    sides[side] = _validate_border_side_value(v)
            else:
                sides[side] = None
        return ParagraphBordersSpec(**sides)
    raise ValueError("format_borders_invalid")


def _validate_tab_stop(value: Any) -> TabStopSpec:
    if isinstance(value, TabStopSpec):
        return value
    if not isinstance(value, Mapping):
        raise TypeError("format_tabs_invalid")
    if set(value) - {
        "val",
        "pos",
        "leader",
        *_TAB_KEY_ALIASES,
    } or _conflicting_alias_keys(value, _TAB_KEY_ALIASES):
        raise ValueError("format_tabs_invalid")
    resolved = {_TAB_KEY_ALIASES.get(key, key): item for key, item in value.items()}
    if "val" not in resolved or "pos" not in resolved:
        raise ValueError("format_tabs_invalid")
    return TabStopSpec(val=resolved["val"], pos=resolved["pos"], leader=resolved.get("leader"))


def _validate_tabs(value: Any) -> tuple[TabStopSpec, ...]:
    """Normalize tab stops into ascending ``pos`` order; empty removes the paragraph's own tabs."""
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        raise TypeError("format_tabs_invalid")
    stops = tuple(sorted((_validate_tab_stop(item) for item in value), key=lambda s: s.pos))
    if len({stop.pos for stop in stops}) != len(stops):
        raise ValueError("format_tabs_invalid")
    return stops


def _validate_paragraph_edit_dict(value: Mapping[str, Any]) -> ParagraphFormatEdit:
    allowed_keys = {
        "paragraph",
        "structure_fingerprint_sha256",
        "part",
        "style",
        "alignment",
        "spacing",
        "indent",
        "keep_next",
        "keep_lines",
        "keep_with_next",
        "keep_together",
        "outline_level",
        "shading",
        "page_break_before",
        "borders",
        "widow_control",
        "tabs",
        "suppress_auto_hyphens",
        "suppress_line_numbers",
        "frame",
    }
    unknown_keys = set(value) - allowed_keys
    if unknown_keys:
        raise ValueError(f"format_paragraph_edit_unknown_keys: {sorted(unknown_keys)}")

    required_keys = {"paragraph", "structure_fingerprint_sha256"}
    missing = required_keys - set(value)
    if missing:
        raise ValueError(f"format_paragraph_edit_missing_keys: {sorted(missing)}")

    paragraph = value["paragraph"]
    if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 1:
        raise ValueError("format_paragraph_invalid")

    structure_fingerprint = _sha(
        value["structure_fingerprint_sha256"], "format_structure_fingerprint_sha256"
    )

    part = value.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part or not part.endswith(".xml"):
        raise ValueError("format_part_invalid")

    style = value.get("style")
    if style is not None and (not isinstance(style, str) or not style or len(style) > 256):
        raise ValueError("format_style_invalid")

    alignment = value.get("alignment")
    if alignment is not None:
        alignment = _canonical_alignment(alignment, "format_alignment_invalid")

    spacing = None
    if "spacing" in value and value["spacing"] is not None:
        spacing = _validate_spacing_spec(value["spacing"])

    indent = None
    if "indent" in value and value["indent"] is not None:
        indent = _validate_indent_spec(value["indent"])

    keep_next = _fold_flag_alias(
        value.get("keep_next"), value.get("keep_with_next"), "format_keep_next"
    )
    keep_lines = _fold_flag_alias(
        value.get("keep_lines"), value.get("keep_together"), "format_keep_lines"
    )

    page_break_before = value.get("page_break_before")
    if page_break_before is not None and not isinstance(page_break_before, bool):
        raise ValueError("format_page_break_before_invalid")

    outline_level = value.get("outline_level")
    if outline_level is not None and (
        isinstance(outline_level, bool)
        or not isinstance(outline_level, int)
        or not (0 <= outline_level <= 9)
    ):
        raise ValueError("format_outline_level_invalid")

    shading = value.get("shading")
    if shading is not None:
        if isinstance(shading, bool):
            if shading is not False:
                raise ValueError("format_shading_invalid")
        elif isinstance(shading, str):
            if not _HEX_COLOR_PATTERN.match(shading):
                raise ValueError("format_shading_invalid")
        else:
            raise ValueError("format_shading_invalid")

    borders = None
    if "borders" in value and value["borders"] is not None:
        borders = _validate_borders_spec(value["borders"])

    widow_control = value.get("widow_control")
    if widow_control is not None and not isinstance(widow_control, bool):
        raise ValueError("format_widow_control_invalid")

    tabs = None
    if "tabs" in value and value["tabs"] is not None:
        tabs = _validate_tabs(value["tabs"])

    suppress_auto_hyphens = value.get("suppress_auto_hyphens")
    if suppress_auto_hyphens is not None and not isinstance(suppress_auto_hyphens, bool):
        raise ValueError("format_suppress_auto_hyphens_invalid")

    suppress_line_numbers = value.get("suppress_line_numbers")
    if suppress_line_numbers is not None and not isinstance(suppress_line_numbers, bool):
        raise ValueError("format_suppress_line_numbers_invalid")

    frame = None
    if "frame" in value and value["frame"] is not None:
        frame = _validate_frame_spec(value["frame"])

    if all(
        p is None
        for p in (
            style,
            alignment,
            spacing,
            indent,
            keep_next,
            keep_lines,
            outline_level,
            shading,
            page_break_before,
            borders,
            widow_control,
            tabs,
            suppress_auto_hyphens,
            suppress_line_numbers,
            frame,
        )
    ):
        raise ValueError("format_paragraph_edit_noop_invalid")

    return ParagraphFormatEdit(
        paragraph=paragraph,
        structure_fingerprint_sha256=structure_fingerprint,
        part=part,
        style=style,
        alignment=alignment,
        spacing=spacing,
        indent=indent,
        keep_next=keep_next,
        keep_lines=keep_lines,
        outline_level=outline_level,
        shading=shading,
        page_break_before=page_break_before,
        borders=borders,
        widow_control=widow_control,
        tabs=tabs,
        suppress_auto_hyphens=suppress_auto_hyphens,
        suppress_line_numbers=suppress_line_numbers,
        frame=frame,
    )


def parse_format_request(value: Any) -> FormatRequest:
    """Parse the strict JSON request for atomic formatting authoring operations."""
    if not isinstance(value, Mapping):
        raise TypeError("format_request_invalid")

    allowed_keys = {
        "source",
        "source_sha256",
        "destination",
        "run_edits",
        "paragraph_edits",
    }
    unknown = set(value) - allowed_keys
    if unknown:
        raise ValueError(f"format_request_unknown_keys: {sorted(unknown)}")

    required = {"source", "source_sha256", "destination"}
    missing = required - set(value)
    if missing:
        raise ValueError(f"format_request_missing_keys: {sorted(missing)}")

    source_raw = value["source"]
    dest_raw = value["destination"]
    if not isinstance(source_raw, str) or not source_raw:
        raise ValueError("format_source_path_invalid")
    if not isinstance(dest_raw, str) or not dest_raw:
        raise ValueError("format_destination_path_invalid")

    source_sha256 = _sha(value["source_sha256"], "format_source_sha256")

    raw_run_edits = value.get("run_edits", ())
    if raw_run_edits is None:
        raw_run_edits = ()
    if not isinstance(raw_run_edits, Sequence) or isinstance(raw_run_edits, (str, bytes)):
        raise TypeError("format_run_edits_invalid")
    run_edits = tuple(_validate_run_edit_dict(item) for item in raw_run_edits)

    raw_para_edits = value.get("paragraph_edits", ())
    if raw_para_edits is None:
        raw_para_edits = ()
    if not isinstance(raw_para_edits, Sequence) or isinstance(raw_para_edits, (str, bytes)):
        raise TypeError("format_paragraph_edits_invalid")
    para_edits = tuple(_validate_paragraph_edit_dict(item) for item in raw_para_edits)

    if not run_edits and not para_edits:
        raise ValueError("format_edits_missing")

    return FormatRequest(
        source=Path(source_raw),
        source_sha256=source_sha256,
        destination=Path(dest_raw),
        run_edits=run_edits,
        paragraph_edits=para_edits,
    )


def load_format_request(path: str | Path) -> FormatRequest:
    """Load a formatting authoring request from a UTF-8 JSON file."""
    with Path(path).open(encoding="utf-8") as stream:
        return parse_format_request(json.load(stream))


def _detect_context_unsupported(paragraph: ET.Element) -> str | None:
    unsupported_revision = {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "rPrChange",
        "pPrChange",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart",
        "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart",
        "customXmlMoveToRangeEnd",
    }
    unsupported_comment = {
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
    }
    unsupported_sdt = {
        "sdt",
        "sdtPr",
        "sdtContent",
        "sdtEndPr",
    }
    unsupported_other = {
        "customXml",
        "fldSimple",
        "fldChar",
    }
    for node in paragraph.iter():
        local = _local_name(node.tag)
        if local in unsupported_revision:
            return "format_revision_context_unsupported"
        if local in unsupported_comment:
            return "format_comment_context_unsupported"
        if local in unsupported_sdt:
            return "format_sdt_context_unsupported"
        if local in unsupported_other:
            return "format_context_unsupported"
    return None


def _is_inside_sdt_or_unsupported(
    paragraph: ET.Element, parents: Mapping[int, ET.Element]
) -> str | None:
    current = parents.get(id(paragraph))
    while current is not None:
        local = _local_name(current.tag)
        if local in {"sdt", "sdtContent"}:
            return "format_sdt_context_unsupported"
        if local in {"ins", "del", "moveFrom", "moveTo"}:
            return "format_revision_context_unsupported"
        current = parents.get(id(current))
    return None


def _catalog_package_styles(
    graph: PackageGraph, ns: str
) -> tuple[dict[str, str], dict[str, str], list[dict[str, Any]]]:
    """Map paragraph and character style references to the ``w:styleId`` they name.

    A reference resolves by style ID first, so every existing ID keeps naming its own
    style, then by ``w:name``, then by python-docx UI name (``Heading 1`` for a style named
    ``heading 1``). Within one tier the first definition wins. Styles without a
    ``w:styleId`` cannot be referenced and are not mapped.
    """
    paragraph_styles: dict[str, str] = {}
    character_styles: dict[str, str] = {}
    catalog: list[dict[str, Any]] = []

    if _STYLES not in graph.parts:
        return paragraph_styles, character_styles, catalog

    try:
        styles_root = ET.fromstring(graph.read_part_bytes(_STYLES))
    except (ET.ParseError, PackageError):
        return paragraph_styles, character_styles, catalog

    # (UI name, w:name, w:styleId) tiers, lowest precedence first.
    tiers: dict[str, tuple[dict[str, str], dict[str, str], dict[str, str]]] = {
        "paragraph": ({}, {}, {}),
        "character": ({}, {}, {}),
    }
    def word(local):
        return Q(ns, local)
    for style_elem in styles_root.findall(word("style")):
        stype = style_elem.get(word("type")) or "paragraph"
        sid = style_elem.get(word("styleId")) or style_elem.get("styleId")
        name_elem = style_elem.find(word("name"))
        sname = None
        if name_elem is not None:
            sname = name_elem.get(word("val")) or name_elem.get("val")

        if not sid:
            continue
        catalog.append({"id": sid, "name": sname, "type": stype})

        kinds = (stype,) if stype in tiers else tuple(tiers)
        for kind in kinds:
            by_ui_name, by_name, by_id = tiers[kind]
            by_id.setdefault(sid, sid)
            if sname:
                by_name.setdefault(sname, sid)
                ui_name = _STYLE_UI_NAMES.get(sname)
                if ui_name is not None:
                    by_ui_name.setdefault(ui_name, sid)

    for kind, resolved in (
        ("paragraph", paragraph_styles),
        ("character", character_styles),
    ):
        by_ui_name, by_name, by_id = tiers[kind]
        resolved.update(by_ui_name)
        resolved.update(by_name)
        resolved.update(by_id)
    return paragraph_styles, character_styles, catalog


def _unrelated_parts_preserved(source: Path, staged: Path, changed_parts: frozenset[str]) -> bool:
    with zipfile.ZipFile(source) as src_zip, zipfile.ZipFile(staged) as dst_zip:
        src_names = set(src_zip.namelist())
        dst_names = set(dst_zip.namelist())
        if src_names != dst_names:
            return False
        for name in src_names:
            if name not in changed_parts and src_zip.read(name) != dst_zip.read(name):
                return False
    return True


def _format_refusal(request: FormatRequest, code: str) -> FormatReport:
    return FormatReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        run_edits=[e.to_dict() for e in request.run_edits],
        paragraph_edits=[e.to_dict() for e in request.paragraph_edits],
        changed_targets=[e.to_dict() for e in request.run_edits]
        + [e.to_dict() for e in request.paragraph_edits],
    )


def _ruby_spans(xml: bytes) -> list[tuple[int, int]]:
    """Locate namespace-qualified ruby byte spans without guessing lexical prefixes."""
    if b"ruby" not in xml:
        return []
    parser = expat.ParserCreate(namespace_separator="}")
    tags = {namespace + "}ruby" for namespace in SUPPORTED_WORD_NAMESPACES}
    starts: list[int] = []
    spans: list[tuple[int, int]] = []

    def start(name: str, attributes: dict[str, str]) -> None:
        if name in tags:
            starts.append(parser.CurrentByteIndex)

    def end(name: str) -> None:
        if name in tags:
            spans.append((starts.pop(), xml.index(b">", parser.CurrentByteIndex) + 1))

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(xml, True)
    return sorted(spans)


def _preserve_unchanged_ruby(source_xml: bytes, output_xml: bytes) -> bytes:
    """Retain the exact lexical bytes of semantically untouched phonetic guides.

    QName-preserving serialization retains their inherited namespace bindings;
    Expat byte offsets additionally preserve Word's empty-tag spelling and attribute
    quoting within ruby. Changed or newly added guides use the normal serializer.
    """
    source_spans = _ruby_spans(source_xml)
    if not source_spans:
        return output_xml
    output_spans = _ruby_spans(output_xml)
    tags = {Q(namespace, "ruby") for namespace in SUPPORTED_WORD_NAMESPACES}

    def keys(xml: bytes) -> list[bytes]:
        result = []
        for element in ET.fromstring(xml).iter():
            if element.tag in tags:
                element.tail = None
                result.append(ET.tostring(element, encoding="utf-8"))
        return result

    originals: dict[bytes, list[bytes]] = {}
    for key, (start, end) in reversed(list(zip(keys(source_xml), source_spans, strict=True))):
        originals.setdefault(key, []).append(source_xml[start:end])
    pieces: list[bytes] = []
    offset = 0
    for key, (start, end) in zip(keys(output_xml), output_spans, strict=True):
        pieces.append(output_xml[offset:start])
        available = originals.get(key)
        pieces.append(available.pop() if available else output_xml[start:end])
        offset = end
    pieces.append(output_xml[offset:])
    return b"".join(pieces)


def _format_paragraph_map(paragraph: ET.Element, namespace: str) -> _ParagraphMap:
    """Map ruby carriers by base text while retaining the original editable run."""
    proxies: dict[int, ET.Element] = {}
    shadow = ET.Element(paragraph.tag)
    for child in paragraph:
        ruby = child.find(Q(namespace, "ruby")) if child.tag == Q(namespace, "r") else None
        if (
            ruby is None
            and child.tag == Q(namespace, "r")
            and _is_break_run(child, namespace)
            and all(
                _local_name(node.tag) in {"rPr", "t", "br", "lastRenderedPageBreak"}
                for node in child
            )
        ):
            proxy = ET.Element(child.tag)
            ET.SubElement(proxy, Q(namespace, "t")).text = "".join(
                str(item) for kind, item in _break_children(child, namespace) if kind == "t"
            )
            shadow.append(proxy)
            proxies[id(proxy)] = child
            continue
        if ruby is None:
            shadow.append(child)
            continue
        if any(_local_name(node.tag) not in {"rPr", "ruby"} for node in child):
            raise _ReviewContextError("format_ruby_context_unsupported")
        base = ruby.find(Q(namespace, "rubyBase"))
        if base is None:
            raise _ReviewContextError("format_ruby_context_unsupported")
        proxy = ET.Element(child.tag)
        ET.SubElement(proxy, Q(namespace, "t")).text = "".join(
            node.text or "" for node in base.iter(Q(namespace, "t"))
        )
        shadow.append(proxy)
        proxies[id(proxy)] = child
    mapping = _paragraph_map(shadow, namespace)
    return replace(
        mapping,
        runs=tuple(
            _TextRun(proxies.get(id(run.element), run.element), run.text, run.start, run.end)
            for run in mapping.runs
        ),
    )


def _read_ruby(run: ET.Element, namespace: str) -> dict[str, Any] | None:
    ruby = run.find(Q(namespace, "ruby"))
    if ruby is None:
        return None
    props = ruby.find(Q(namespace, "rubyPr"))
    result: dict[str, Any] = {}
    for key, local in (("base_text", "rubyBase"), ("guide_text", "rt")):
        container = ruby.find(Q(namespace, local))
        result[key] = (
            ""
            if container is None
            else "".join(node.text or "" for node in container.iter(Q(namespace, "t")))
        )
    for key, local in (
        ("alignment", "rubyAlign"),
        ("raise", "hpsRaise"),
        ("font_size", "hps"),
        ("base_font_size", "hpsBaseText"),
        ("language", "lid"),
    ):
        node = props.find(Q(namespace, local)) if props is not None else None
        value = node.get(Q(namespace, "val")) if node is not None else None
        result[key] = (
            int(value)
            if value is not None
            and key in {"raise", "font_size", "base_font_size"}
            and value.isdecimal()
            else value
        )
    return result


def _apply_run_format(run: ET.Element, edit: RunFormatEdit, namespace: str) -> None:
    def word(local):
        return Q(namespace, local)
    rpr = run.find(word("rPr"))
    if rpr is None:
        rpr = ET.Element(word("rPr"))
        run.insert(0, rpr)

    if edit.bold is not None:
        for existing in list(rpr.findall(word("b"))):
            rpr.remove(existing)
        if edit.bold:
            rpr.append(ET.Element(word("b")))
        else:
            rpr.append(ET.Element(word("b"), {word("val"): "0"}))

    if edit.italic is not None:
        for existing in list(rpr.findall(word("i"))):
            rpr.remove(existing)
        if edit.italic:
            rpr.append(ET.Element(word("i")))
        else:
            rpr.append(ET.Element(word("i"), {word("val"): "0"}))

    for flag, tag in ((edit.all_caps, "caps"), (edit.hidden, "vanish")):
        if flag is None:
            continue
        for existing in list(rpr.findall(word(tag))):
            rpr.remove(existing)
        rpr.append(ET.Element(word(tag)) if flag else ET.Element(word(tag), {word("val"): "0"}))

    if edit.small_caps is not None:
        for existing in list(rpr.findall(word("smallCaps"))):
            rpr.remove(existing)
        if edit.small_caps:
            rpr.append(ET.Element(word("smallCaps")))
        else:
            rpr.append(ET.Element(word("smallCaps"), {word("val"): "0"}))

    if edit.strike is not None:
        for existing in list(rpr.findall(word("strike"))):
            rpr.remove(existing)
        if edit.strike:
            rpr.append(ET.Element(word("strike")))
        else:
            rpr.append(ET.Element(word("strike"), {word("val"): "0"}))

    if edit.underline is not None:
        for existing in list(rpr.findall(word("u"))):
            rpr.remove(existing)
        val = (
            "single"
            if edit.underline is True
            else ("none" if edit.underline is False else str(edit.underline))
        )
        rpr.append(ET.Element(word("u"), {word("val"): val}))

    if edit.color is not None:
        for existing in list(rpr.findall(word("color"))):
            rpr.remove(existing)
        val = edit.color.strip()
        if val.lower() != "auto":
            val = val.upper()
        rpr.append(ET.Element(word("color"), {word("val"): val}))

    if edit.font_size is not None:
        for existing in list(rpr.findall(word("sz"))):
            rpr.remove(existing)
        for existing in list(rpr.findall(word("szCs"))):
            rpr.remove(existing)
        val = str(edit.font_size)
        rpr.append(ET.Element(word("sz"), {word("val"): val}))
        rpr.append(ET.Element(word("szCs"), {word("val"): val}))

    if edit.vertical_alignment is not None:
        for existing in list(rpr.findall(word("vertAlign"))):
            rpr.remove(existing)
        rpr.append(ET.Element(word("vertAlign"), {word("val"): edit.vertical_alignment}))

    if edit.style is not None:
        for existing in list(rpr.findall(word("rStyle"))):
            rpr.remove(existing)
        rpr.append(ET.Element(word("rStyle"), {word("val"): edit.style}))

    if edit.font_face is not None:
        for existing in list(rpr.findall(word("rFonts"))):
            rpr.remove(existing)
        rpr.append(
            ET.Element(
                word("rFonts"),
                {
                    word("ascii"): edit.font_face,
                    word("hAnsi"): edit.font_face,
                    word("cs"): edit.font_face,
                },
            )
        )
    if edit.font_face_east_asia is not None:
        fonts = rpr.find(word("rFonts"))
        if fonts is None:
            fonts = ET.SubElement(rpr, word("rFonts"))
        # ECMA-376 17.3.2.26: w:eastAsiaTheme supersedes w:eastAsia, so drop it
        # to make the explicit face effective; every other attribute is kept.
        fonts.attrib.pop(word("eastAsiaTheme"), None)
        fonts.set(word("eastAsia"), edit.font_face_east_asia)
        ordered = {
            word(name): fonts.attrib[word(name)]
            for name in _RFONTS_ATTRIBUTE_ORDER
            if word(name) in fonts.attrib
        }
        ordered.update(fonts.attrib)
        fonts.attrib.clear()
        fonts.attrib.update(ordered)
    if edit.highlight is not None:
        for existing in list(rpr.findall(word("highlight"))):
            rpr.remove(existing)
        val = _HIGHLIGHT_CANONICAL_MAP.get(
            edit.highlight.strip().lower() if isinstance(edit.highlight, str) else "",
            edit.highlight,
        )
        rpr.append(ET.Element(word("highlight"), {word("val"): val}))
    for value, tag in ((edit.legacy_animation, "effect"), (edit.emphasis_mark, "em")):
        if value is not None:
            for existing in list(rpr.findall(word(tag))):
                rpr.remove(existing)
            ET.SubElement(rpr, word(tag), {word("val"): value})
    if edit.numeral_spacing is not None:
        for existing in list(rpr.findall(Q(_W14, "numSpacing"))):
            rpr.remove(existing)
        if edit.numeral_spacing is not False:
            ET.SubElement(rpr, Q(_W14, "numSpacing"), {Q(_W14, "val"): edit.numeral_spacing})
    if edit.ruby is not None:
        rpr[:] = sorted(rpr, key=lambda child: _RPR_RANK.get(_local_name(child.tag), 999))
        spec = edit.ruby
        for child in list(run):
            if child.tag != word("rPr"):
                run.remove(child)
        ruby = ET.SubElement(run, word("ruby"))
        props = ET.SubElement(ruby, word("rubyPr"))
        for local, value in (
            ("rubyAlign", spec.alignment),
            ("hps", spec.font_size),
            ("hpsRaise", spec.raise_),
            ("hpsBaseText", spec.base_font_size),
            ("lid", spec.language),
        ):
            ET.SubElement(props, word(local), {word("val"): str(value)})
        for local, text in (("rt", spec.guide_text), ("rubyBase", spec.base_text)):
            carrier = ET.SubElement(ET.SubElement(ruby, word(local)), word("r"))
            if local == "rubyBase":
                carrier.append(copy.deepcopy(rpr))
            else:
                guide_props = ET.SubElement(carrier, word("rPr"))
                ET.SubElement(guide_props, word("sz"), {word("val"): str(spec.font_size)})
            ET.SubElement(
                carrier,
                word("t"),
                {"{http://www.w3.org/XML/1998/namespace}space": "preserve"},
            ).text = text

    # Sort rPr children by canonical CT_RPr schema rank
    rpr[:] = sorted(rpr, key=lambda c: _RPR_RANK.get(_local_name(c.tag), 999))


def _apply_paragraph_format(
    paragraph: ET.Element, edit: ParagraphFormatEdit, namespace: str
) -> None:
    def word(local):
        return Q(namespace, local)
    ppr = paragraph.find(word("pPr"))
    if ppr is None:
        ppr = ET.Element(word("pPr"))
        paragraph.insert(0, ppr)

    if edit.style is not None:
        for existing in list(ppr.findall(word("pStyle"))):
            ppr.remove(existing)
        ppr.append(ET.Element(word("pStyle"), {word("val"): edit.style}))

    if edit.alignment is not None:
        for existing in list(ppr.findall(word("jc"))):
            ppr.remove(existing)
        ppr.append(ET.Element(word("jc"), {word("val"): edit.alignment}))

    if edit.spacing is not None:
        for existing in list(ppr.findall(word("spacing"))):
            ppr.remove(existing)
        attrib: dict[str, str] = {}
        if edit.spacing.before is not None:
            attrib[word("before")] = str(edit.spacing.before)
        if edit.spacing.after is not None:
            attrib[word("after")] = str(edit.spacing.after)
        if edit.spacing.line is not None:
            attrib[word("line")] = str(edit.spacing.line)
        if edit.spacing.line_rule is not None:
            rule_map = {"auto": "auto", "exact": "exact", "at_least": "atLeast"}
            attrib[word("lineRule")] = rule_map.get(edit.spacing.line_rule, edit.spacing.line_rule)
        ppr.append(ET.Element(word("spacing"), attrib))

    if edit.indent is not None:
        for existing in list(ppr.findall(word("ind"))):
            ppr.remove(existing)
        attrib = {}
        if edit.indent.left is not None:
            attrib[word("left")] = str(edit.indent.left)
        if edit.indent.right is not None:
            attrib[word("right")] = str(edit.indent.right)
        if edit.indent.first_line is not None:
            attrib[word("firstLine")] = str(edit.indent.first_line)
        if edit.indent.hanging is not None:
            attrib[word("hanging")] = str(edit.indent.hanging)
        if edit.indent.start is not None:
            attrib[word("start")] = str(edit.indent.start)
        if edit.indent.end is not None:
            attrib[word("end")] = str(edit.indent.end)
        ppr.append(ET.Element(word("ind"), attrib))

    if edit.keep_next is not None:
        for existing in list(ppr.findall(word("keepNext"))):
            ppr.remove(existing)
        if edit.keep_next:
            ppr.append(ET.Element(word("keepNext")))
        else:
            ppr.append(ET.Element(word("keepNext"), {word("val"): "0"}))

    if edit.keep_lines is not None:
        for existing in list(ppr.findall(word("keepLines"))):
            ppr.remove(existing)
        if edit.keep_lines:
            ppr.append(ET.Element(word("keepLines")))
        else:
            ppr.append(ET.Element(word("keepLines"), {word("val"): "0"}))

    if edit.page_break_before is not None:
        for existing in list(ppr.findall(word("pageBreakBefore"))):
            ppr.remove(existing)
        if edit.page_break_before:
            ppr.append(ET.Element(word("pageBreakBefore")))
        else:
            ppr.append(ET.Element(word("pageBreakBefore"), {word("val"): "0"}))

    if edit.widow_control is not None:
        for existing in list(ppr.findall(word("widowControl"))):
            ppr.remove(existing)
        if edit.widow_control:
            ppr.append(ET.Element(word("widowControl")))
        else:
            ppr.append(ET.Element(word("widowControl"), {word("val"): "0"}))

    for flag, tag in (
        (edit.suppress_auto_hyphens, "suppressAutoHyphens"),
        (edit.suppress_line_numbers, "suppressLineNumbers"),
    ):
        if flag is None:
            continue
        for existing in list(ppr.findall(word(tag))):
            ppr.remove(existing)
        ppr.append(ET.Element(word(tag)) if flag else ET.Element(word(tag), {word("val"): "0"}))

    if edit.frame is not None:
        frames = ppr.findall(word("framePr"))
        if edit.frame is False:
            for existing in frames:
                ppr.remove(existing)
        else:
            frame = frames[0] if frames else ET.SubElement(ppr, word("framePr"))
            for field, value in edit.frame.to_dict().items():
                frame.set(word(_FRAME_ATTRIBUTES[field]), _frame_value(value))

    if edit.tabs is not None:
        for existing in list(ppr.findall(word("tabs"))):
            ppr.remove(existing)
        # CT_Tabs requires at least one w:tab, so an empty sequence leaves no w:tabs element.
        stops = _validate_tabs(edit.tabs)
        if stops:
            tabs_elem = ET.SubElement(ppr, word("tabs"))
            for stop in stops:
                tab_attrib = {word("val"): stop.val}
                if stop.leader is not None:
                    tab_attrib[word("leader")] = stop.leader
                tab_attrib[word("pos")] = str(stop.pos)
                ET.SubElement(tabs_elem, word("tab"), tab_attrib)

    if edit.outline_level is not None:
        for existing in list(ppr.findall(word("outlineLvl"))):
            ppr.remove(existing)
        ppr.append(ET.Element(word("outlineLvl"), {word("val"): str(edit.outline_level)}))

    if edit.shading is not None:
        for existing in list(ppr.findall(word("shd"))):
            ppr.remove(existing)
        if edit.shading is not False:
            ppr.append(
                ET.Element(
                    word("shd"),
                    {
                        word("val"): "clear",
                        word("color"): "auto",
                        word("fill"): edit.shading,
                    },
                )
            )

    if edit.borders is not None:
        existing_p_bdr = ppr.find(word("pBdr"))
        existing_sides: dict[str, ET.Element] = {}
        if existing_p_bdr is not None:
            for c in existing_p_bdr:
                existing_sides[_local_name(c.tag)] = c
        for existing in list(ppr.findall(word("pBdr"))):
            ppr.remove(existing)

        if edit.borders is not False:
            borders_dict: dict[str, Any] = (
                edit.borders.to_dict()
                if isinstance(edit.borders, ParagraphBordersSpec)
                else dict(edit.borders)
            )
            new_p_bdr = ET.Element(word("pBdr"))
            for side in _PARAGRAPH_BORDER_SIDES:
                if side in borders_dict:
                    spec = borders_dict[side]
                    if spec is not False and spec is not None:
                        b_elem = ET.Element(word(side))
                        val = (
                            spec.val if isinstance(spec, BorderSpec) else spec.get("val", "single")
                        )
                        b_elem.set(word("val"), str(val))
                        sz = spec.sz if isinstance(spec, BorderSpec) else spec.get("sz")
                        if sz is not None:
                            b_elem.set(word("sz"), str(sz))
                        space = spec.space if isinstance(spec, BorderSpec) else spec.get("space")
                        if space is not None:
                            b_elem.set(word("space"), str(space))
                        color = spec.color if isinstance(spec, BorderSpec) else spec.get("color")
                        if color is not None:
                            b_elem.set(word("color"), str(color))
                        new_p_bdr.append(b_elem)
                elif side in existing_sides:
                    new_p_bdr.append(copy.deepcopy(existing_sides[side]))
            for side, elem in existing_sides.items():
                if side not in _PARAGRAPH_BORDER_SIDES:
                    new_p_bdr.append(copy.deepcopy(elem))
            if len(new_p_bdr) > 0:
                ppr.append(new_p_bdr)
    # Sort pPr children by canonical CT_PPr schema rank
    ppr[:] = sorted(ppr, key=lambda c: _PPR_RANK.get(_local_name(c.tag), 999))

    # Enforce pPr is the very first child of w:p
    if paragraph[0] is not ppr:
        paragraph.remove(ppr)
        paragraph.insert(0, ppr)


def _compatibility_mode(parts: Mapping[str, bytes]) -> int | None:
    data = parts.get("word/settings.xml")
    if data is None:
        return None
    try:
        root = ET.fromstring(data)
        ns = _word_namespace(root)
        if ns is None:
            raise ValueError("format_compatibilityMode_invalid")
        settings = [
            elem
            for elem in root.findall(f"{Q(ns, 'compat')}/{Q(ns, 'compatSetting')}")
            if elem.get(Q(ns, "name")) == "compatibilityMode"
            and elem.get(Q(ns, "uri")) == "http://schemas.microsoft.com/office/word"
        ]
        if not settings:
            return None
        if len(settings) != 1:
            raise ValueError("format_compatibilityMode_invalid")
        value = settings[0].get(Q(ns, "val"), "")
        if not value.isdecimal():
            raise ValueError("format_compatibilityMode_invalid")
        return int(value)
    except ET.ParseError as error:
        raise ValueError("format_compatibilityMode_invalid") from error


def _effect_xml(element: ET.Element) -> dict[str, Any]:
    return {
        "tag": element.tag,
        "attributes": dict(element.attrib),
        "text": element.text,
        "children": [_effect_xml(child) for child in element],
    }


def _read_w14_properties(rpr: ET.Element) -> dict[str, Any]:
    # text_effects imports formatting helpers, so defer this import until discovery.
    from .text_effects import read_effect

    properties: dict[str, Any] = {}
    for child in rpr:
        if not child.tag.startswith(f"{{{_W14}}}"):
            continue
        local = _local_name(child.tag)
        if local in {"glow", "shadow", "reflection", "textOutline", "textFill"}:
            parsed = read_effect(child)
            key = {"textOutline": "outline", "textFill": "fill"}.get(local, local)
            properties[key] = parsed[1].to_dict() if parsed is not None else _effect_xml(child)
        elif local in {"scene3d", "props3d"}:
            key = "scene_3d" if local == "scene3d" else "props_3d"
            properties[key] = _effect_xml(child)
        elif local in {"ligatures", "numForm", "numSpacing"}:
            key = {
                "ligatures": "ligatures",
                "numForm": "number_form",
                "numSpacing": "numeral_spacing",
            }[local]
            properties[key] = child.get(Q(_W14, "val"))
        elif local == "cntxtAlts":
            properties["contextual_alternates"] = child.get(Q(_W14, "val"), "1") not in {
                "0",
                "false",
                "off",
            }
        elif local == "stylisticSets":
            properties["stylistic_sets"] = [
                int(style.get(Q(_W14, "id")))
                for style in child.findall(Q(_W14, "styleSet"))
                if style.get(Q(_W14, "id"), "").isdecimal()
                and style.get(Q(_W14, "val"), "1") not in {"0", "false", "off"}
            ]
    return properties


def _is_break_run(run: ET.Element, namespace: str) -> bool:
    """A run carrying w:br or w:lastRenderedPageBreak, which it edits only as a whole."""
    return any(
        child.tag in {Q(namespace, "br"), Q(namespace, "lastRenderedPageBreak")} for child in run
    )


def _break_element(edit: RunFormatEdit, namespace: str) -> ET.Element:
    attributes = {Q(namespace, "type"): str(edit.break_type)}
    if edit.break_clear is not None:
        attributes[Q(namespace, "clear")] = edit.break_clear
    return ET.Element(Q(namespace, "br"), attributes)


def _read_frame(frame: ET.Element, namespace: str) -> dict[str, Any]:
    """Report the attributes of one w:framePr under their snake_case request names."""
    result: dict[str, Any] = {}
    for field, attribute in _FRAME_ATTRIBUTES.items():
        raw = frame.get(Q(namespace, attribute)) or frame.get(attribute)
        if raw is None:
            continue
        if field == "anchor_lock":
            result[field] = raw.strip().lower() not in {"0", "false", "off"}
        elif field in _FRAME_INTEGER_FIELDS:
            result[field] = int(raw) if re.fullmatch(r"-?\d+", raw) else raw
        else:
            result[field] = raw
    return result


def _break_children(run: ET.Element, namespace: str) -> list[tuple[str, ET.Element | str]]:
    """Text and break content of a run in document order (text as plain strings)."""
    content: list[tuple[str, ET.Element | str]] = []
    for child in run:
        local = _local_name(child.tag)
        if child.tag == Q(namespace, "t"):
            content.append(("t", child.text or ""))
        elif local in {"br", "lastRenderedPageBreak"}:
            content.append((local, child))
    return content


def _read_breaks(mapping: _ParagraphMap, namespace: str) -> tuple[list[dict[str, Any]], list[int]]:
    """w:br and w:lastRenderedPageBreak inventory with the paragraph text offset of each."""
    breaks: list[dict[str, Any]] = []
    rendered: list[int] = []
    for run in mapping.runs:
        offset = run.start
        for kind, item in _break_children(run.element, namespace):
            if isinstance(item, str):
                offset += len(item)
            elif kind == "br":
                breaks.append(
                    {
                        "offset": offset,
                        "type": item.get(Q(namespace, "type")) or "textWrapping",
                        "clear": item.get(Q(namespace, "clear")) or "none",
                    }
                )
            else:
                rendered.append(offset)
    return breaks, rendered


def discover_format_targets(source: str | Path, *, part: str = _DOCUMENT) -> dict[str, Any]:
    """List source-bound paragraph targets and available styles without writing."""
    path = Path(source)
    if not path.is_file():
        return {
            "passed": False,
            "code": "format_source_package_invalid",
            "targets": [],
            "styles": [],
        }

    try:
        graph = PackageGraph.from_file(path)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError, zipfile.BadZipFile):
        return {
            "passed": False,
            "code": "format_source_package_invalid",
            "targets": [],
            "styles": [],
        }

    if _DOCUMENT not in parts:
        return {
            "passed": False,
            "code": "format_source_package_invalid",
            "targets": [],
            "styles": [],
        }

    try:
        document_root = ET.fromstring(parts[_DOCUMENT])
    except ET.ParseError:
        return {
            "passed": False,
            "code": "format_source_package_invalid",
            "targets": [],
            "styles": [],
        }

    ns = _word_namespace(document_root)
    if ns is None:
        return {
            "passed": False,
            "code": "format_namespace_unsupported",
            "targets": [],
            "styles": [],
        }

    allowed = _review_parts(graph, document_root, ns)
    if part not in allowed:
        return {
            "passed": False,
            "code": "format_part_unsupported",
            "source_sha256": graph.file_sha256,
            "part": part,
            "targets": [],
            "styles": [],
        }

    try:
        root = ET.fromstring(parts[part])
    except ET.ParseError:
        return {
            "passed": False,
            "code": "format_source_package_invalid",
            "targets": [],
            "styles": [],
        }

    part_ns = _word_namespace(root)
    if part_ns is None:
        return {
            "passed": False,
            "code": "format_namespace_unsupported",
            "targets": [],
            "styles": [],
        }

    parents = {id(child): parent for parent in root.iter() for child in parent}
    def word(local):
        return Q(part_ns, local)

    _, _, styles_catalog = _catalog_package_styles(graph, part_ns)
    try:
        compatibility_mode = _compatibility_mode(parts)
    except ValueError as error:
        return {
            "passed": False,
            "code": str(error),
            "targets": [],
            "styles": styles_catalog,
        }

    # styles_authoring imports this module, so defer the shared script-property reader.
    from .styles_authoring import _read_script_properties

    targets: list[dict[str, Any]] = []
    for p_idx, paragraph in enumerate(root.iter(word("p")), start=1):
        fingerprint = _mask_text(paragraph, part_ns)
        context_code = _detect_context_unsupported(paragraph) or _is_inside_sdt_or_unsupported(
            paragraph, parents
        )

        # Inspect paragraph properties
        ppr = paragraph.find(word("pPr"))
        p_style = None
        p_align = None
        p_keep_next = False
        p_keep_lines = False
        p_page_break_before = False
        # Omitted w:widowControl reports the Word default (on); style inheritance is not resolved.
        p_widow_control = True
        p_tabs: list[dict[str, Any]] | None = None
        p_outline_lvl = None
        p_spacing = None
        p_indent = None
        p_shading = None
        p_suppress_auto_hyphens = False
        p_suppress_line_numbers = False
        p_frame: dict[str, Any] | None = None
        p_script: dict[str, Any] = {}

        if ppr is not None:
            style_elem = ppr.find(word("pStyle"))
            if style_elem is not None:
                p_style = style_elem.get(word("val")) or style_elem.get("val")

            jc_elem = ppr.find(word("jc"))
            if jc_elem is not None:
                p_align = jc_elem.get(word("val")) or jc_elem.get("val")

            kn_elem = ppr.find(word("keepNext"))
            if kn_elem is not None:
                p_keep_next = _flag(kn_elem) is not False

            kl_elem = ppr.find(word("keepLines"))
            if kl_elem is not None:
                p_keep_lines = _flag(kl_elem) is not False

            pbb_elem = ppr.find(word("pageBreakBefore"))
            if pbb_elem is not None:
                p_page_break_before = _flag(pbb_elem) is not False

            wc_elem = ppr.find(word("widowControl"))
            if wc_elem is not None:
                p_widow_control = _flag(wc_elem) is not False

            sah_elem = ppr.find(word("suppressAutoHyphens"))
            if sah_elem is not None:
                p_suppress_auto_hyphens = _flag(sah_elem) is not False

            sln_elem = ppr.find(word("suppressLineNumbers"))
            if sln_elem is not None:
                p_suppress_line_numbers = _flag(sln_elem) is not False

            frame_elem = ppr.find(word("framePr"))
            if frame_elem is not None:
                p_frame = _read_frame(frame_elem, part_ns)

            p_script = _read_script_properties(ppr, part_ns, paragraph=True)

            tabs_elem = ppr.find(word("tabs"))
            if tabs_elem is not None:
                p_tabs = []
                for tab_elem in tabs_elem.findall(word("tab")):
                    pos_raw = tab_elem.get(word("pos")) or tab_elem.get("pos")
                    try:
                        pos_val: int | str | None = int(pos_raw) if pos_raw is not None else None
                    except ValueError:
                        pos_val = pos_raw
                    tab_info: dict[str, Any] = {
                        "val": tab_elem.get(word("val")) or tab_elem.get("val"),
                        "pos": pos_val,
                    }
                    leader = tab_elem.get(word("leader")) or tab_elem.get("leader")
                    if leader is not None:
                        tab_info["leader"] = leader
                    p_tabs.append(tab_info)

            shd_elem = ppr.find(word("shd"))
            if shd_elem is not None:
                fill_val = shd_elem.get(word("fill")) or shd_elem.get("fill")
                if fill_val and fill_val.lower() != "auto":
                    p_shading = fill_val

            ol_elem = ppr.find(word("outlineLvl"))
            ol_val = (
                (ol_elem.get(word("val")) or ol_elem.get("val")) if ol_elem is not None else None
            )
            if ol_val is not None and ol_val.isdigit():
                p_outline_lvl = int(ol_val)

            sp_elem = ppr.find(word("spacing"))
            if sp_elem is not None:
                p_spacing = {}
                for attr_key in ("before", "after", "line", "lineRule"):
                    val = sp_elem.get(word(attr_key)) or sp_elem.get(attr_key)
                    if val is not None:
                        snake = "line_rule" if attr_key == "lineRule" else attr_key
                        # Twips measures are reported as integers so a discovered spacing
                        # can be passed back unchanged as a request's ``spacing``.
                        p_spacing[snake] = int(val) if val.isdecimal() else val

            ind_elem = ppr.find(word("ind"))
            if ind_elem is not None:
                p_indent = {}
                for attr_key in (
                    "left",
                    "right",
                    "firstLine",
                    "hanging",
                    "start",
                    "end",
                ):
                    val = ind_elem.get(word(attr_key)) or ind_elem.get(attr_key)
                    if val is not None:
                        snake = "first_line" if attr_key == "firstLine" else attr_key
                        p_indent[snake] = val

            pbdr_elem = ppr.find(word("pBdr"))
            p_borders = None
            if pbdr_elem is not None:
                b_dict: dict[str, Any] = {}
                for side in _PARAGRAPH_BORDER_SIDES:
                    side_elem = pbdr_elem.find(word(side))
                    if side_elem is not None:
                        val = side_elem.get(word("val")) or side_elem.get("val") or "single"
                        sz_val = side_elem.get(word("sz")) or side_elem.get("sz")
                        sp_val = side_elem.get(word("space")) or side_elem.get("space")
                        col_val = side_elem.get(word("color")) or side_elem.get("color") or "auto"
                        b_dict[side] = {
                            "val": val,
                            "sz": int(sz_val) if sz_val and sz_val.isdigit() else 4,
                            "space": int(sp_val) if sp_val and sp_val.isdigit() else 0,
                            "color": col_val,
                        }
                if b_dict:
                    p_borders = b_dict
        else:
            p_borders = None
        runs_info: list[dict[str, Any]] = []
        editable_ranges: list[list[int]] = []
        p_text = ""
        p_breaks: list[dict[str, Any]] = []
        p_last_rendered: list[int] = []

        if context_code is None:
            try:
                mapping = _format_paragraph_map(paragraph, part_ns)
                p_text = mapping.text
                editable_ranges = [[start, end] for start, end in mapping.editable_ranges]
                p_breaks, p_last_rendered = _read_breaks(mapping, part_ns)
                for r in mapping.runs:
                    r_elem = r.element
                    rpr = r_elem.find(word("rPr"))
                    sc_elem = rpr.find(word("smallCaps")) if rpr is not None else None
                    run_data: dict[str, Any] = {
                        "start": r.start,
                        "end": r.end,
                        "text": r.text,
                        "small_caps": _flag(sc_elem) is True,
                    }
                    ruby_data = _read_ruby(r_elem, part_ns)
                    if ruby_data is not None:
                        run_data["ruby"] = ruby_data
                    if rpr is not None:
                        w14_properties = _read_w14_properties(rpr)
                        run_data.update(w14_properties)
                        if w14_properties:
                            run_data["w14_effects_active"] = (
                                compatibility_mode is None or compatibility_mode >= 14
                            )
                        b_elem = rpr.find(word("b"))
                        if b_elem is not None:
                            run_data["bold"] = _flag(b_elem) is not False
                        i_elem = rpr.find(word("i"))
                        if i_elem is not None:
                            run_data["italic"] = _flag(i_elem) is not False
                        u_elem = rpr.find(word("u"))
                        if u_elem is not None:
                            run_data["underline"] = (
                                u_elem.get(word("val")) or u_elem.get("val") or "single"
                            )
                        s_elem = rpr.find(word("strike"))
                        if s_elem is not None:
                            run_data["strike"] = _flag(s_elem) is not False
                        c_elem = rpr.find(word("color"))
                        if c_elem is not None:
                            run_data["color"] = c_elem.get(word("val")) or c_elem.get("val")
                        sz_elem = rpr.find(word("sz"))
                        sz_val = (
                            (sz_elem.get(word("val")) or sz_elem.get("val"))
                            if sz_elem is not None
                            else None
                        )
                        if sz_val and sz_val.isdigit():
                            run_data["font_size"] = int(sz_val)
                        va_elem = rpr.find(word("vertAlign"))
                        if va_elem is not None:
                            run_data["vertical_alignment"] = va_elem.get(
                                word("val")
                            ) or va_elem.get("val")
                        st_elem = rpr.find(word("rStyle"))
                        if st_elem is not None:
                            run_data["style"] = st_elem.get(word("val")) or st_elem.get("val")
                        rf_elem = rpr.find(word("rFonts"))
                        if rf_elem is not None:
                            run_data["font_face"] = (
                                rf_elem.get(word("ascii"))
                                or rf_elem.get(word("hAnsi"))
                                or rf_elem.get("ascii")
                                or rf_elem.get("hAnsi")
                            )
                            east_asia = rf_elem.get(word("eastAsia")) or rf_elem.get("eastAsia")
                            if east_asia is not None:
                                run_data["font_face_east_asia"] = east_asia
                        hl_elem = rpr.find(word("highlight"))
                        if hl_elem is not None:
                            run_data["highlight"] = hl_elem.get(word("val")) or hl_elem.get("val")
                        for tag, key in (("effect", "legacy_animation"), ("em", "emphasis_mark")):
                            prop = rpr.find(word(tag))
                            if prop is not None:
                                run_data[key] = prop.get(word("val")) or prop.get("val")
                        run_data.update(_read_script_properties(rpr, part_ns, paragraph=False))
                    runs_info.append(run_data)
            except _ReviewContextError as err:
                context_code = err.code or "format_context_unsupported"
                p_text = "".join(t.text or "" for t in paragraph.iter(word("t")))
                editable_ranges = []
        else:
            p_text = "".join(t.text or "" for t in paragraph.iter(word("t")))
            editable_ranges = []

        targets.append(
            {
                "part": part,
                "paragraph": p_idx,
                "structure_fingerprint_sha256": fingerprint,
                "text": p_text,
                "style": p_style,
                "alignment": p_align,
                "keep_next": p_keep_next,
                "keep_lines": p_keep_lines,
                "page_break_before": p_page_break_before,
                "widow_control": p_widow_control,
                "tabs": p_tabs,
                "outline_level": p_outline_lvl,
                "spacing": p_spacing,
                "indent": p_indent,
                "shading": p_shading,
                "borders": p_borders,
                "suppress_auto_hyphens": p_suppress_auto_hyphens,
                "suppress_line_numbers": p_suppress_line_numbers,
                "frame": p_frame,
                **p_script,
                "breaks": p_breaks,
                "last_rendered_page_breaks": p_last_rendered,
                "runs": runs_info,
                "editable_ranges": editable_ranges,
                "context_code": context_code,
            }
        )

    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        "part": part,
        "targets": targets,
        "styles": styles_catalog,
        "compatibility_mode": compatibility_mode,
    }


def author_format(request: FormatRequest) -> FormatReport:
    """Execute source-bound run and paragraph formatting changes atomically."""
    source = request.source
    destination = request.destination

    if not source.is_file():
        return _format_refusal(request, "format_source_not_found")

    if _sha256(source) != request.source_sha256:
        return _format_refusal(request, "format_source_sha256_mismatch")

    if destination.exists() or destination.is_symlink():
        return _format_refusal(request, "format_destination_invalid")

    try:
        graph = PackageGraph.from_file(source)
        parts_bytes = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError, zipfile.BadZipFile):
        return _format_refusal(request, "format_package_invalid")

    if _DOCUMENT not in parts_bytes:
        return _format_refusal(request, "format_package_invalid")

    try:
        doc_root = ET.fromstring(parts_bytes[_DOCUMENT])
    except ET.ParseError:
        return _format_refusal(request, "format_package_invalid")

    main_ns = _word_namespace(doc_root)
    if main_ns is None or main_ns not in SUPPORTED_WORD_NAMESPACES:
        return _format_refusal(request, "format_namespace_unsupported")

    allowed_parts = _review_parts(graph, doc_root, main_ns)

    # Style catalog validation
    para_styles, char_styles, _ = _catalog_package_styles(graph, main_ns)

    for r_edit in request.run_edits:
        if r_edit.color is not None and (
            not isinstance(r_edit.color, str)
            or (r_edit.color.lower() != "auto" and not _HEX_COLOR_PATTERN.match(r_edit.color))
        ):
            return _format_refusal(request, "format_color_invalid")
        if r_edit.highlight is not None and (
            isinstance(r_edit.highlight, bool)
            or not isinstance(r_edit.highlight, str)
            or r_edit.highlight.strip().lower() not in _HIGHLIGHT_CANONICAL_MAP
        ):
            return _format_refusal(request, "format_highlight_invalid")

    # Style references resolve to the w:styleId written into w:rStyle/w:pStyle.
    run_edits: list[RunFormatEdit] = []
    for r_edit in request.run_edits:
        if r_edit.style is not None:
            if _STYLES not in graph.parts or r_edit.style not in char_styles:
                return _format_refusal(request, "format_style_missing")
            r_edit = replace(r_edit, style=char_styles[r_edit.style])
        run_edits.append(r_edit)

    paragraph_edits: list[ParagraphFormatEdit] = []
    for p_edit in request.paragraph_edits:
        if p_edit.style is not None:
            if _STYLES not in graph.parts or p_edit.style not in para_styles:
                return _format_refusal(request, "format_style_missing")
            p_edit = replace(p_edit, style=para_styles[p_edit.style])
        paragraph_edits.append(p_edit)

    # Group edits by part and paragraph
    run_edits_by_part_p: dict[tuple[str, int], list[RunFormatEdit]] = {}
    for r_edit in run_edits:
        if r_edit.part not in allowed_parts:
            return _format_refusal(request, "format_part_unsupported")
        run_edits_by_part_p.setdefault((r_edit.part, r_edit.paragraph), []).append(r_edit)

    para_edits_by_part_p: dict[tuple[str, int], ParagraphFormatEdit] = {}
    for p_edit in paragraph_edits:
        if p_edit.part not in allowed_parts:
            return _format_refusal(request, "format_part_unsupported")
        key = (p_edit.part, p_edit.paragraph)
        if key in para_edits_by_part_p:
            return _format_refusal(request, "format_duplicate_paragraph_edit")
        para_edits_by_part_p[key] = p_edit

    # Check for overlapping run edits within the same paragraph
    for key, p_run_edits in run_edits_by_part_p.items():
        sorted_edits = sorted(p_run_edits, key=lambda e: (e.start, e.end))
        for e1, e2 in itertools.pairwise(sorted_edits):
            if e1.end > e2.start:
                return _format_refusal(request, "format_edits_overlap")

    # Load and parse roots for all touched parts
    touched_part_names = {part for part, _ in run_edits_by_part_p} | {
        part for part, _ in para_edits_by_part_p
    }
    part_roots: dict[str, ET.Element] = {}
    part_namespaces: dict[str, str] = {}
    part_parents: dict[str, dict[int, ET.Element]] = {}

    for part_name in touched_part_names:
        try:
            root = parse_xml_with_preserved_qnames(parts_bytes[part_name])
        except ET.ParseError:
            return _format_refusal(request, "format_package_invalid")
        ns = _word_namespace(root)
        if ns is None:
            return _format_refusal(request, "format_namespace_unsupported")
        part_roots[part_name] = root
        part_namespaces[part_name] = ns
        part_parents[part_name] = {id(child): parent for parent in root.iter() for child in parent}

    # Pre-validation of all edits against source document structure
    for (part_name, p_idx), p_run_edits in run_edits_by_part_p.items():
        root = part_roots[part_name]
        ns = part_namespaces[part_name]
        parents = part_parents[part_name]
        paragraphs = list(root.iter(Q(ns, "p")))
        if p_idx < 1 or p_idx > len(paragraphs):
            return _format_refusal(request, "format_paragraph_missing")

        paragraph = paragraphs[p_idx - 1]
        fp = _mask_text(paragraph, ns)

        for edit in p_run_edits:
            if fp != edit.structure_fingerprint_sha256:
                return _format_refusal(request, "format_structure_fingerprint_mismatch")

        context_unsupported = _detect_context_unsupported(
            paragraph
        ) or _is_inside_sdt_or_unsupported(paragraph, parents)
        if context_unsupported:
            return _format_refusal(request, context_unsupported)

        try:
            mapping = _format_paragraph_map(paragraph, ns)
        except _ReviewContextError:
            return _format_refusal(request, "format_context_unsupported")

        p_text = mapping.text
        for edit in p_run_edits:
            if edit.start < 0 or edit.end > len(p_text) or edit.end <= edit.start:
                return _format_refusal(request, "format_range_invalid")
            if p_text[edit.start : edit.end] != edit.expected_old_text:
                return _format_refusal(request, "format_expected_text_mismatch")
            if edit.ruby is not None and edit.ruby.base_text != edit.expected_old_text:
                return _format_refusal(request, "format_ruby_base_text_mismatch")
            intersected = [
                run for run in mapping.runs if run.start < edit.end and run.end > edit.start
            ]
            if edit.ruby is not None and len(intersected) != 1:
                return _format_refusal(request, "format_ruby_multiple_runs_unsupported")
            for run in intersected:
                if run.element.find(Q(ns, "ruby")) is not None and (
                    edit.start > run.start or edit.end < run.end
                ):
                    return _format_refusal(request, "format_ruby_partial_range_unsupported")
            for run in intersected:
                if edit.break_type is not None and run.element.find(Q(ns, "ruby")) is not None:
                    return _format_refusal(request, "format_break_ruby_unsupported")
                if not _is_break_run(run.element, ns):
                    continue
                if edit.ruby is not None:
                    return _format_refusal(request, "format_break_ruby_unsupported")
                if edit.start > run.start or edit.end < run.end:
                    return _format_refusal(request, "format_break_partial_range_unsupported")

    for (part_name, p_idx), p_edit in para_edits_by_part_p.items():
        root = part_roots[part_name]
        ns = part_namespaces[part_name]
        parents = part_parents[part_name]
        paragraphs = list(root.iter(Q(ns, "p")))
        if p_idx < 1 or p_idx > len(paragraphs):
            return _format_refusal(request, "format_paragraph_missing")

        paragraph = paragraphs[p_idx - 1]
        fp = _mask_text(paragraph, ns)
        if fp != p_edit.structure_fingerprint_sha256:
            return _format_refusal(request, "format_structure_fingerprint_mismatch")

        context_unsupported = _detect_context_unsupported(
            paragraph
        ) or _is_inside_sdt_or_unsupported(paragraph, parents)
        if context_unsupported:
            return _format_refusal(request, context_unsupported)

    if any(edit.numeral_spacing is not None for edit in run_edits):
        try:
            compatibility_mode = _compatibility_mode(parts_bytes)
        except ValueError as error:
            return _format_refusal(request, str(error))
        if (
            compatibility_mode is not None
            and compatibility_mode < 14
            and any(isinstance(edit.numeral_spacing, str) for edit in run_edits)
        ):
            return _format_refusal(request, "format_w14_suppressed_by_compatibilityMode")

    # Perform mutations
    # 1. Apply paragraph format edits
    for (part_name, p_idx), p_edit in para_edits_by_part_p.items():
        root = part_roots[part_name]
        ns = part_namespaces[part_name]
        paragraphs = list(root.iter(Q(ns, "p")))
        paragraph = paragraphs[p_idx - 1]
        _apply_paragraph_format(paragraph, p_edit, ns)

    # 2. Apply run format edits
    for (part_name, p_idx), p_run_edits in run_edits_by_part_p.items():
        root = part_roots[part_name]
        ns = part_namespaces[part_name]
        paragraphs = list(root.iter(Q(ns, "p")))
        paragraph = paragraphs[p_idx - 1]

        mapping = _format_paragraph_map(paragraph, ns)
        runs_by_id = {id(r.element): r for r in mapping.runs}
        boundaries = sorted({e.start for e in p_run_edits} | {e.end for e in p_run_edits})

        child_list = list(paragraph)
        for child in child_list:
            run_data = runs_by_id.get(id(child))
            if run_data is None or not run_data.text:
                continue

            cuts = [
                run_data.start,
                *(b for b in boundaries if run_data.start < b < run_data.end),
                run_data.end,
            ]
            new_runs: list[ET.Element] = []
            for lower, upper in itertools.pairwise(cuts):
                sub_text = run_data.text[lower - run_data.start : upper - run_data.start]
                cloned = (
                    copy.deepcopy(child)
                    if child.find(Q(ns, "ruby")) is not None or _is_break_run(child, ns)
                    else _clone_run(child, ns, sub_text)
                )

                # Check if this slice is within any edit
                matched_edit = next(
                    (e for e in p_run_edits if e.start <= lower and upper <= e.end),
                    None,
                )
                if matched_edit is not None:
                    _apply_run_format(cloned, matched_edit, ns)
                    if matched_edit.break_type is not None and upper == matched_edit.end:
                        cloned.append(_break_element(matched_edit, ns))

                new_runs.append(cloned)

            # Replace child in paragraph with new_runs
            idx = list(paragraph).index(child)
            paragraph.remove(child)
            for offset, new_r in enumerate(new_runs):
                paragraph.insert(idx + offset, new_r)

    # Serialize modified parts
    replacements: dict[str, bytes] = {}
    for part_name, root in part_roots.items():
        if any(
            edit.part == part_name and isinstance(edit.numeral_spacing, str) for edit in run_edits
        ):
            scope = dict(getattr(root, "_namespace_scope", ()))
            prefix = "wdx14"
            while prefix in scope and scope[prefix] != _W14:
                prefix += "x"
            root.set("xmlns:" + prefix, _W14)
            ignorable = root.get(Q(_MC, "Ignorable"), "").split()
            if prefix not in ignorable:
                ignorable.append(prefix)
            root.set(Q(_MC, "Ignorable"), " ".join(ignorable))
        replacements[part_name] = _preserve_unchanged_ruby(
            parts_bytes[part_name], serialize_xml_with_preserved_qnames(root)
        )

    changed_targets = [e.to_dict() for e in run_edits] + [e.to_dict() for e in paragraph_edits]

    def verifier(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset(replacements)):
            return "format_staged_verification_failed"
        return None

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=changed_targets,
        verify=verifier,
        failure_code="format_staged_verification_failed",
        exists_code="format_destination_invalid",
        unavailable_code="format_publication_unavailable",
    )

    return FormatReport(
        status=published.status,
        code=published.code,
        source_sha256=published.source_sha256,
        destination_sha256=published.output_sha256,
        run_edits=[e.to_dict() for e in run_edits],
        paragraph_edits=[e.to_dict() for e in paragraph_edits],
        changed_targets=changed_targets,
    )
