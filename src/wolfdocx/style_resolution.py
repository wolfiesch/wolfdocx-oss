"""Neutral WordprocessingML style and section property resolution.

The read/edit core and the native renderer share these readers: document
defaults, paragraph, character, and table style chains with ``w:basedOn``
inheritance, theme colors and fonts, Markup Compatibility validation of the
parts they read, and section column geometry. Table and page-setup authoring
resolve support through the same rules the renderer applies, without loading
the renderer, its fonts, or its imaging dependencies.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, replace
from typing import Any
from xml.parsers import expat

from .assembly import _tint_or_shade
from .package import PackageError, PackageGraph
from .review import _CONFLICT_REVISION_TAGS

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_WORD_TAG_PREFIXES = (f"{{{_W}}}", f"{{{_STRICT_W}}}")
_MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_WORD2010_WORDML = "http://schemas.microsoft.com/office/word/2010/wordml"
_W14 = f"{{{_WORD2010_WORDML}}}"
#: The Word 2010 run effects that draw glyph paint: fill, outline, shadow, glow and reflection.
_W14_TEXT_EFFECTS = ("textFill", "textOutline", "shadow", "glow", "reflection")
_WORD2012_WORDML = "http://schemas.microsoft.com/office/word/2012/wordml"
_MC_ALTERNATE_CONTENT = f"{{{_MC}}}AlternateContent"
_MC_CHOICE = f"{{{_MC}}}Choice"
_MC_FALLBACK = f"{{{_MC}}}Fallback"

#: Native layout's explicitly declared Markup Compatibility profile.  These are
#: namespace URIs whose WordprocessingML content this renderer implements, never
#: prefix spellings found in a source document.  A differently bound ``w14`` or
#: ``w15`` prefix is therefore not accepted, and an equivalent locally bound
#: prefix is.
_MC_LAYOUT_SUPPORTED_NAMESPACES = frozenset({_W, _MATH, _WORD2010_WORDML, _WORD2012_WORDML})


def _validate_conflict_revision(element: ET.Element) -> None:
    """Refuse coauthor semantics at layout consumption, independently of MC selection."""
    if element.tag in _CONFLICT_REVISION_TAGS:
        raise NativeLayoutError("native_render_conflict_revision_unsupported")


def _mc_layout_metadata(
    data: bytes,
    root: ET.Element,
    *,
    supported_namespaces: frozenset[str] = _MC_LAYOUT_SUPPORTED_NAMESPACES,
    ignored_attributes: dict[int, frozenset[str]] | None = None,
) -> tuple[dict[int, dict[str, str]], dict[int, bool], dict[int, tuple[str | None, ...]]]:
    """Map lexical MC scopes, ignorable wrapper actions, and mandatory URIs.

    :mod:`xml.etree.ElementTree` expands element names but deliberately discards
    lexical namespace declarations.  ``mc:Choice/@Requires``,
    ``mc:ProcessContent``, and ``mc:MustUnderstand`` are lexical, so the raw XML
    is scanned once to bind
    their prefixes at the element where they occur.  The supplied tree is never
    changed: if it is not the graph's exact source tree, no metadata is
    associated with it.

    Ignorable wrapper actions map element IDs to whether their children must be
    processed. QNames are expanded at the declaration, inherited by descendants,
    and restored at scope exit, independently of later prefix rebinding. A
    ``prefix:*`` token admits every element of that namespace (ECMA-376 Part 3).
    """

    choice_scopes: dict[int, dict[str, str]] = {}
    ignored_positions: dict[int, bool] = {}
    required_positions: dict[int, tuple[str | None, ...]] = {}
    namespaces: dict[str, str] = {}
    namespace_frames: list[list[tuple[str, str | None]]] = []
    ignorable = frozenset()
    ignorable_frames: list[frozenset[str]] = []
    process_content: frozenset[tuple[str, str]] = frozenset()
    process_frames: list[frozenset[tuple[str, str]]] = []
    attribute_positions: dict[int, frozenset[str]] = {}
    position = 0

    def split_name(name: str) -> tuple[str, str]:
        prefix, separator, local = name.partition(":")
        return (prefix, local) if separator else ("", prefix)

    def start(name: str, attributes: dict[str, str]) -> None:
        nonlocal ignorable, process_content, position
        changes: list[tuple[str, str | None]] = []
        for attribute, value in attributes.items():
            if attribute == "xmlns":
                prefix = ""
            elif attribute.startswith("xmlns:"):
                prefix = attribute[6:]
            else:
                continue
            changes.append((prefix, namespaces.get(prefix)))
            namespaces[prefix] = value
        namespace_frames.append(changes)

        prefix, local = split_name(name)
        namespace = namespaces.get(prefix)
        additions: set[str] = set()
        process_additions: set[tuple[str, str]] = set()
        requirements: list[str | None] = []
        for attribute, value in attributes.items():
            attribute_prefix, attribute_local = split_name(attribute)
            if not attribute_prefix or namespaces.get(attribute_prefix) != _MC:
                continue
            if attribute_local in {"Ignorable", "MustUnderstand"}:
                prefixes = value.split()
            elif attribute_local in {"ProcessContent", "PreserveElements", "PreserveAttributes"}:
                prefixes = [split_name(token)[0] for token in value.split()]
            else:
                prefixes = []
            if any(token not in namespaces for token in prefixes):
                requirements.append(None)
            if attribute_local == "Ignorable":
                additions.update(
                    namespaces[token] for token in value.split() if token in namespaces
                )
            elif attribute_local == "ProcessContent":
                for token in value.split():
                    token_prefix, token_local = split_name(token)
                    token_namespace = namespaces.get(token_prefix)
                    if token_namespace is not None:
                        process_additions.add((token_namespace, token_local))
            elif attribute_local == "MustUnderstand":
                requirements.extend(namespaces.get(token) for token in value.split())
        ignorable_frames.append(ignorable)
        process_frames.append(process_content)
        if process_additions:
            process_content = process_content | frozenset(process_additions)
        if additions:
            ignorable = ignorable | frozenset(additions)

        if namespace == _MC and local == "Choice":
            choice_scopes[position] = dict(namespaces)
            if any(token not in namespaces for token in attributes.get("Requires", "").split()):
                requirements.append(None)
        if requirements:
            required_positions[position] = tuple(requirements)
        if ignored_attributes is not None:
            ignored_names = frozenset(
                f"{{{namespaces[attribute_prefix]}}}{attribute_local}"
                for attribute in attributes
                for attribute_prefix, attribute_local in [split_name(attribute)]
                if attribute_prefix
                and namespaces.get(attribute_prefix) in ignorable
                and namespaces[attribute_prefix] not in supported_namespaces
            )
            if ignored_names:
                attribute_positions[position] = ignored_names
        if namespace in ignorable and namespace not in supported_namespaces and namespace != _MC:
            named = (namespace, local) in process_content
            ignored_positions[position] = named or (namespace, "*") in process_content
        position += 1

    def end(_name: str) -> None:
        nonlocal ignorable, process_content
        ignorable = ignorable_frames.pop()
        process_content = process_frames.pop()
        for prefix, previous in reversed(namespace_frames.pop()):
            if previous is None:
                del namespaces[prefix]
            else:
                namespaces[prefix] = previous

    parser = expat.ParserCreate()
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except expat.ExpatError as error:
        raise NativeLayoutError("native_render_source_invalid") from error

    indexed_scopes: dict[int, dict[str, str]] = {}
    indexed_ignored: dict[int, bool] = {}
    indexed_required: dict[int, tuple[str | None, ...]] = {}
    actual_count = 0
    for actual_count, element in enumerate(root.iter(), start=1):
        source_position = actual_count - 1
        if ignored_attributes is not None and source_position in attribute_positions:
            ignored_attributes[id(element)] = attribute_positions[source_position]
        scope = choice_scopes.get(source_position)
        if scope is not None and element.tag == _MC_CHOICE:
            indexed_scopes[id(element)] = scope
        if source_position in ignored_positions:
            indexed_ignored[id(element)] = ignored_positions[source_position]
        required = required_positions.get(source_position)
        if required is not None:
            indexed_required[id(element)] = required
    if actual_count != position:
        return ({}, {}, {})
    return (indexed_scopes, indexed_ignored, indexed_required)


def _validate_mc_layout_part(data: bytes, root: ET.Element) -> None:
    """Validate mandatory namespaces and conflict semantics in a consumed part.

    This follows the same source-bound metadata and branch selection as body
    layout, but visits only the selected branches of the one part its caller is
    already about to consume. Ignorable wrappers admitted by ProcessContent are
    traversed; unselected alternatives are not. Other ignored subtrees are not
    consumed, but ``mc:MustUnderstand`` on them or any descendant still refuses,
    as the Open XML SDK does.
    """

    choice_scopes, ignored, required = _mc_layout_metadata(data, root)

    def validate_requirements(element: ET.Element) -> None:
        requirements = required.get(id(element), ())
        if any(namespace not in _MC_LAYOUT_SUPPORTED_NAMESPACES for namespace in requirements):
            raise NativeLayoutError("native_render_markup_compatibility_unsupported")

    def validate(element: ET.Element) -> None:
        _validate_conflict_revision(element)
        validate_requirements(element)

    def selected(alternate: ET.Element) -> ET.Element | None:
        fallback: ET.Element | None = None
        for candidate in alternate:
            if candidate.tag == _MC_CHOICE:
                requires = (candidate.get("Requires") or "").split()
                scopes = choice_scopes.get(id(candidate), {})
                if any(prefix not in scopes for prefix in requires):
                    raise NativeLayoutError("native_render_markup_compatibility_unsupported")
                if requires and all(
                    scopes.get(prefix) in _MC_LAYOUT_SUPPORTED_NAMESPACES for prefix in requires
                ):
                    return candidate
            elif candidate.tag == _MC_FALLBACK:
                fallback = candidate
        return fallback

    def consume(element: ET.Element) -> None:
        validate(element)
        for child in element:
            if ignored.get(id(child)) is False:
                for node in child.iter():
                    validate_requirements(node)
                continue
            if child.tag == _MC_ALTERNATE_CONTENT:
                validate(child)
                branch = selected(child)
                if branch is None:
                    raise NativeLayoutError("native_render_markup_compatibility_unsupported")
                consume(branch)
            else:
                consume(child)

    consume(root)


_SETTINGS = "word/settings.xml"
_STYLES = "word/styles.xml"
_THEME = "word/theme/theme1.xml"

_POINTS_PER_TWIP = 1.0 / 20.0

#: Tab leaders this engine draws, keyed by the lowercased spelling the tab-stop
#: parser reads: OOXML spells ``middleDot`` with a capital D, which lowercasing
#: alone would lose.
_TAB_LEADERS = {
    "none": "none",
    "dot": "dot",
    "hyphen": "hyphen",
    "underscore": "underscore",
    "heavy": "heavy",
    "middledot": "middleDot",
}

_DEFAULT_CELL_MARGIN = 5.75


class NativeLayoutError(Exception):
    """One explicit layout refusal carrying a bounded code."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


@dataclass(frozen=True, slots=True)
class _Border:
    color: str
    width: float
    kind: str = "single"
    space: float = 0.0


def _local(tag: Any) -> str:
    return tag.split("}", 1)[-1] if isinstance(tag, str) else ""


def _child(element: ET.Element | None, name: str) -> ET.Element | None:
    if element is None:
        return None
    for child in element:
        if _local(child.tag) == name:
            return child
    return None


def _word_child(element: ET.Element | None, name: str) -> ET.Element | None:
    """A direct child by local name, never the Word 2010 ``w14`` effect of the same name."""
    if element is None:
        return None
    for child in element:
        if _local(child.tag) == name and not child.tag.startswith(_W14):
            return child
    return None


def _children(element: ET.Element, name: str) -> list[ET.Element]:
    return [child for child in element if _local(child.tag) == name]


def _nested(element: ET.Element | None, *names: str) -> ET.Element | None:
    current = element
    for name in names:
        current = _child(current, name)
    return current


def _value(element: ET.Element | None, name: str) -> str | None:
    if element is None:
        return None
    for key, value in element.attrib.items():
        if _local(key) == name:
            return value
    return None


def _flag(element: ET.Element | None) -> bool | None:
    if element is None:
        return None
    value = _value(element, "val")
    if value is None:
        return True
    return value.strip().lower() not in {"0", "false", "off"}


def _attribute_flag(element: ET.Element, name: str) -> bool:
    """One on/off attribute declared directly on ``element``.

    Automatic paragraph spacing is declared by attributes of ``w:spacing`` rather
    than by child elements, so :func:`_flag` cannot read it. An absent attribute,
    or one written as ``0``/``false``/``off``, means the property is off.
    """

    value = _value(element, name)
    return value is not None and value.strip().lower() not in {"0", "false", "off"}


def _integer(text: str | None) -> int | None:
    if text is None:
        return None
    try:
        return int(text.strip())
    except ValueError:
        return None


def _measure(element: ET.Element | None, name: str, scale: float) -> float | None:
    number = _integer(_value(element, name))
    if number is None:
        return None
    return number * scale


def _number(element: ET.Element | None, name: str, scale: float) -> float | None:
    """Read one numeric attribute that may be fractional, scaled by ``scale``."""

    text = _value(element, name)
    if text is None:
        return None
    try:
        return float(text.strip()) * scale
    except ValueError:
        return None


def _hex_color(value: str | None, default: str = "000000") -> str:
    if value is None:
        return default
    text = value.strip().lstrip("#").lower()
    if text in {"", "auto"}:
        return default
    if len(text) == 3 and all(character in "0123456789abcdef" for character in text):
        return "".join(character * 2 for character in text)
    if len(text) == 6 and all(character in "0123456789abcdef" for character in text):
        return text
    return default


def _theme_color_transform(
    value: str, tint: str | None, shade: str | None, *, hsl: bool = True
) -> str | None:
    # Native rendering permits surrounding whitespace and falls back to the
    # explicit color on malformed transforms; assembly deliberately refuses.
    try:
        return _tint_or_shade(
            value,
            tint.strip() if tint is not None else None,
            shade.strip() if shade is not None else None,
            hsl=hsl,
        ).lower()
    except ValueError:
        return None


_HIGHLIGHTS = {
    "black": "000000",
    "blue": "0000FF",
    "cyan": "00FFFF",
    "darkblue": "000080",
    "darkcyan": "008080",
    "darkgray": "808080",
    "darkgreen": "008000",
    "darkmagenta": "800080",
    "darkred": "800000",
    "darkyellow": "808000",
    "green": "00FF00",
    "lightgray": "C0C0C0",
    "magenta": "FF00FF",
    "red": "FF0000",
    "white": "FFFFFF",
    "yellow": "FFFF00",
}


@dataclass(frozen=True, slots=True)
class _SupplementalFont:
    """A theme face awaiting the consuming run's inherited script language."""

    prefix: str
    role: str
    language_key: str
    fallback: str | _SupplementalFont | None = None


def _merge_run_properties(inherited: dict[str, Any], own: dict[str, Any]) -> None:
    """Cascade declarations without discarding an unresolved theme's fallback."""
    for key, value in own.items():
        if isinstance(value, _SupplementalFont):
            fallback = inherited.get(key)
            if fallback is not None:
                chain = []
                current = value
                while isinstance(current, _SupplementalFont):
                    chain.append(current)
                    current = current.fallback
                if current is None:
                    for declaration in reversed(chain):
                        fallback = replace(declaration, fallback=fallback)
                    value = fallback
        inherited[key] = value


class _Styles:
    """Resolved document defaults, style chains, and theme font names."""

    def __init__(self, graph: PackageGraph) -> None:
        self.default_run: dict[str, Any] = {}
        self.default_paragraph: dict[str, Any] = {}
        #: PDF language comes from the document default run language. OOXML may omit
        #: it; retain the established English default rather than infer it from text.
        self.language = "en-US"
        self.major_font: str | None = None
        self.minor_font: str | None = None
        self._theme_colors: dict[str, str] = {}
        self.chart_line_width = 0.75
        self.chart_line_node: ET.Element | None = None
        self.chart_fill_styles: tuple[ET.Element, ...] = ()
        self.chart_effect_styles: tuple[ET.Element, ...] = ()
        self._theme_fonts: dict[tuple[str, str], str] = {}
        self._paragraph_styles: dict[str, tuple[str | None, dict[str, Any], dict[str, Any]]] = {}
        self._character_styles: dict[str, tuple[str | None, dict[str, Any]]] = {}
        self._table_styles: dict[str, tuple[str | None, dict[str, Any], dict[str, Any]]] = {}
        self.conditional_table_styles: set[str] = set()
        self._paragraph_cache: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {}
        self._character_cache: dict[str, dict[str, Any]] = {}
        self._table_cache: dict[str, dict[str, Any]] = {}
        self._load_theme(graph)
        self._load_styles(graph)

    def _load_theme(self, graph: PackageGraph) -> None:
        theme = graph.part(_THEME)
        if theme is None:
            return
        try:
            # Graph adapters may serialize a normalized drawing view, losing
            # lexical MC prefixes. Validate the original part bytes instead.
            theme_data = theme.read_bytes()
            root = ET.fromstring(theme_data)
            settings_part = graph.part(_SETTINGS)
            settings_data = settings_part.read_bytes() if settings_part is not None else None
            settings = ET.fromstring(settings_data) if settings_data is not None else None
        except (ET.ParseError, PackageError):
            raise NativeLayoutError("native_render_style_unsupported") from None
        _validate_mc_layout_part(theme_data, root)
        if settings is not None and settings_data is not None:
            _validate_mc_layout_part(settings_data, settings)
        color_scheme = next((node for node in root.iter() if _local(node.tag) == "clrScheme"), None)
        if color_scheme is not None:
            for entry in color_scheme:
                color = next(
                    (
                        _value(child, "val")
                        if _local(child.tag) == "srgbClr"
                        else _value(child, "lastClr")
                        if _local(child.tag) == "sysClr"
                        else None
                        for child in entry
                    ),
                    None,
                )
                normalized = _hex_color(color, "")
                if normalized:
                    self._theme_colors[_local(entry.tag).casefold()] = normalized
            mapping = {
                "bg1": "lt1",
                "tx1": "dk1",
                "bg2": "lt2",
                "tx2": "dk2",
                "dark1": "dk1",
                "light1": "lt1",
                "dark2": "dk2",
                "light2": "lt2",
                "hyperlink": "hlink",
                "followedhyperlink": "folhlink",
            }
            scheme_mapping = (
                next(
                    (node for node in settings.iter() if _local(node.tag) == "clrSchemeMapping"),
                    None,
                )
                if settings is not None
                else None
            )
            if scheme_mapping is not None:
                aliases = {
                    "t1": "tx1",
                    "t2": "tx2",
                    "light1": "lt1",
                    "dark1": "dk1",
                    "light2": "lt2",
                    "dark2": "dk2",
                }
                mapping.update(
                    {
                        aliases.get(_local(name).casefold(), _local(name).casefold()): aliases.get(
                            value.casefold(), value.casefold()
                        )
                        for name, value in scheme_mapping.attrib.items()
                    }
                )
            for name, target in mapping.items():
                color = self._theme_colors.get(target)
                if color is not None:
                    self._theme_colors[name] = color
            for name, target in {
                "background1": "bg1",
                "text1": "tx1",
                "background2": "bg2",
                "text2": "tx2",
            }.items():
                color = self._theme_colors.get(target)
                if color is not None:
                    self._theme_colors[name] = color
        line_styles = next((node for node in root.iter() if _local(node.tag) == "lnStyleLst"), None)
        subtle_line = _child(line_styles, "ln") if line_styles is not None else None
        line_width = _integer(_value(subtle_line, "w"))
        if line_width is not None and line_width > 0:
            self.chart_line_width = line_width / 12700.0
        self.chart_line_node = subtle_line
        for name, attribute in (
            ("fillStyleLst", "chart_fill_styles"),
            ("effectStyleLst", "chart_effect_styles"),
        ):
            declarations = next((node for node in root.iter() if _local(node.tag) == name), None)
            if declarations is not None:
                setattr(self, attribute, tuple(declarations))
        scheme = next((node for node in root.iter() if _local(node.tag) == "fontScheme"), None)
        if scheme is None:
            return
        for family in scheme:
            prefix = {"majorFont": "major", "minorFont": "minor"}.get(_local(family.tag))
            if prefix is None:
                continue
            for child in family:
                typeface = (_value(child, "typeface") or "").strip()
                if not typeface:
                    continue
                local = _local(child.tag)
                role = {
                    "latin": "latin",
                    "ea": "east_asian",
                    "cs": "complex",
                }.get(local)
                if role is not None:
                    self._theme_fonts[(prefix, role)] = typeface
                elif local == "font":
                    script = (_value(child, "script") or "").strip().casefold()
                    if script:
                        self._theme_fonts[(prefix, f"east_asian:{script}")] = typeface
                        self._theme_fonts[(prefix, f"complex:{script}")] = typeface
        self.major_font = self._theme_fonts.get(("major", "latin"))
        self.minor_font = self._theme_fonts.get(("minor", "latin"))

    def _load_styles(self, graph: PackageGraph) -> None:
        styles = graph.part(_STYLES)
        if styles is None:
            return
        try:
            data = styles.read_bytes()
            root = ET.fromstring(data)
        except (ET.ParseError, PackageError):
            raise NativeLayoutError("native_render_style_unsupported") from None
        _validate_mc_layout_part(data, root)
        defaults = _child(root, "docDefaults")
        if defaults is not None:
            default_run = _nested(defaults, "rPrDefault", "rPr")
            language = _value(_child(default_run, "lang"), "val")
            if language and language.strip():
                self.language = language.strip()
            self.default_run = self.run_properties(default_run)
            self.default_paragraph = self.paragraph_properties(
                _nested(defaults, "pPrDefault", "pPr")
            )
        for style in _children(root, "style"):
            style_id = _value(style, "styleId")
            if not style_id:
                continue
            kind = _value(style, "type")
            # ``w:basedOn`` is a child element that carries ``w:val``; reading it as
            # an attribute of ``w:style`` would silently drop every style chain.
            based_on = _value(_child(style, "basedOn"), "val")
            if kind == "paragraph":
                self._paragraph_styles[style_id] = (
                    based_on,
                    self.paragraph_properties(_child(style, "pPr")),
                    self.run_properties(_child(style, "rPr")),
                )
            elif kind == "character":
                self._character_styles[style_id] = (
                    based_on,
                    self.run_properties(_child(style, "rPr")),
                )
            elif kind == "table":
                self._table_styles[style_id] = (
                    based_on,
                    self.table_properties(_child(style, "tblPr")),
                    self.cell_properties(_child(style, "tcPr")),
                )
                if _child(style, "tblStylePr") is not None:
                    self.conditional_table_styles.add(style_id)
        if len(self._paragraph_styles) + len(self._character_styles) > 8192:
            raise NativeLayoutError("native_render_limit_exceeded")

    def paragraph_style(self, style_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        return self._resolve_paragraph(style_id, 0)

    def _resolve_paragraph(
        self, style_id: str, depth: int
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        cached = self._paragraph_cache.get(style_id)
        if cached is not None:
            return cached
        if depth > 32:
            raise NativeLayoutError("native_render_style_unsupported")
        entry = self._paragraph_styles.get(style_id)
        self._paragraph_cache[style_id] = ({}, {})
        if entry is None:
            return ({}, {})
        paragraph: dict[str, Any] = {}
        run: dict[str, Any] = {}
        if entry[0]:
            parent_paragraph, parent_run = self._resolve_paragraph(entry[0], depth + 1)
            paragraph.update(parent_paragraph)
            run.update(parent_run)
        if "borders" in entry[1]:
            combined_borders = dict(paragraph.get("borders", {}))
            combined_borders.update(entry[1]["borders"])
            paragraph.update(entry[1])
            paragraph["borders"] = combined_borders
        else:
            paragraph.update(entry[1])
        _merge_run_properties(run, entry[2])
        self._paragraph_cache[style_id] = (paragraph, run)
        return (paragraph, run)

    def character_style(self, style_id: str) -> dict[str, Any]:
        return self._resolve_character(style_id, 0)

    def _resolve_character(self, style_id: str, depth: int) -> dict[str, Any]:
        cached = self._character_cache.get(style_id)
        if cached is not None:
            return cached
        if depth > 32:
            raise NativeLayoutError("native_render_style_unsupported")
        entry = self._character_styles.get(style_id)
        self._character_cache[style_id] = {}
        if entry is None:
            return {}
        resolved: dict[str, Any] = {}
        if entry[0]:
            resolved.update(self._resolve_character(entry[0], depth + 1))
        _merge_run_properties(resolved, entry[1])
        self._character_cache[style_id] = resolved
        return resolved

    def table_style(self, style_id: str) -> dict[str, Any]:
        return self._resolve_table(style_id, 0)

    def _resolve_table(self, style_id: str, depth: int) -> dict[str, Any]:
        cached = self._table_cache.get(style_id)
        if cached is not None:
            return cached
        if depth > 32:
            raise NativeLayoutError("native_render_style_unsupported")
        entry = self._table_styles.get(style_id)
        self._table_cache[style_id] = {}
        if entry is None:
            return {}
        resolved: dict[str, Any] = {}
        if entry[0]:
            resolved.update(self._resolve_table(entry[0], depth + 1))
        resolved.update(entry[1])
        resolved.update(entry[2])
        if style_id in self.conditional_table_styles:
            resolved["conditional_formatting"] = True
        self._table_cache[style_id] = resolved
        return resolved

    # -- property readers -------------------------------------------------
    def run_properties(self, element: ET.Element | None) -> dict[str, Any]:
        properties: dict[str, Any] = {}
        if element is None:
            return properties
        language = _child(element, "lang")
        for attribute, key in (
            ("val", "language"),
            ("eastAsia", "east_asian_language"),
            ("bidi", "complex_language"),
        ):
            value = _value(language, attribute)
            if value and value.strip():
                properties[key] = value.strip()
        fonts = _child(element, "rFonts")
        if fonts is not None:
            family = self._font_family(fonts, "latin")
            if family:
                properties["family"] = family
            east_asian_family = self._font_family(fonts, "east_asian")
            if east_asian_family:
                properties["east_asian_family"] = east_asian_family
            complex_family = self._font_family(fonts, "complex")
            if complex_family:
                properties["complex_family"] = complex_family
        size = _measure(_child(element, "sz"), "val", 0.5)
        if size is not None and size > 0:
            properties["size"] = size
        bold = _child(element, "b")
        if bold is not None:
            properties["bold"] = _flag(bold) is not False
        italic = _child(element, "i")
        if italic is not None:
            properties["italic"] = _flag(italic) is not False
        underline = _child(element, "u")
        if underline is not None:
            properties["underline"] = (_value(underline, "val") or "single").lower() != "none"
        strike = _child(element, "strike")
        if strike is not None:
            properties["strike"] = _flag(strike) is not False
        for name in ("vanish", "outline", "shadow", "emboss", "imprint"):
            effect = _word_child(element, name)
            if effect is not None:
                # An explicit ``w:val="0"`` participates in the style cascade:
                # without retaining it here, a direct-off property could not undo
                # the same property inherited through a character style.
                properties[name] = _flag(effect) is not False
        for name in _W14_TEXT_EFFECTS:
            # A Word 2010 text effect cascades per effect, each with its own element.
            effect = next((child for child in element if child.tag == _W14 + name), None)
            if effect is not None:
                properties["w14_" + name] = effect
        if _child(element, "bdr") is not None:
            properties["run_border"] = True
        kerning = _child(element, "kern")
        if kerning is not None and (_integer(_value(kerning, "val")) or 0) > 0:
            properties["kern"] = True
        rtl = _child(element, "rtl")
        if rtl is not None:
            properties["rtl"] = _flag(rtl) is not False
        color = _child(element, "color")
        if color is not None:
            properties["color"] = _hex_color(self._color_value(color))
        highlight = _child(element, "highlight")
        if highlight is not None:
            properties["highlight"] = _HIGHLIGHTS.get(
                (_value(highlight, "val") or "").strip().lower()
            )
        vertical = _child(element, "vertAlign")
        if vertical is not None:
            properties["vertical"] = (_value(vertical, "val") or "").strip() or None
        style = _child(element, "rStyle")
        if style is not None and _value(style, "val"):
            properties["run_style"] = _value(style, "val")
        return properties

    def _color_value(self, element: ET.Element, *, direct: str = "val") -> str | None:
        prefix = next(
            (prefix for prefix in _WORD_TAG_PREFIXES if element.tag.startswith(prefix)),
            None,
        )
        if prefix is None:
            return None
        fallback = element.get(f"{prefix}{direct}")
        if direct == "fill":
            theme_name, tint_name, shade_name = (
                "themeFill",
                "themeFillTint",
                "themeFillShade",
            )
        else:
            theme_name, tint_name, shade_name = "themeColor", "themeTint", "themeShade"
        theme = element.get(f"{prefix}{theme_name}")
        if not theme:
            return fallback
        value = self._theme_colors.get(theme.strip().casefold())
        if value is None:
            return fallback
        tint = element.get(f"{prefix}{tint_name}")
        # Tint wins for foreground (17.3.2.6/17.3.4/17.3.5) and, as observed
        # in Word 16.113.2, for background themeFillTint/themeFillShade too.
        shade = None if tint is not None else element.get(f"{prefix}{shade_name}")
        transformed = _theme_color_transform(value, tint, shade, hsl=_local(element.tag) != "color")
        return transformed if transformed is not None else fallback

    @staticmethod
    def _script(locale: str | None) -> str | None:
        if not locale:
            return None
        language, _, region = locale.casefold().partition("-")
        if language == "zh":
            return "hant" if region in {"hk", "mo", "tw"} else "hans"
        return {
            "ar": "arab",
            "fa": "arab",
            "ur": "arab",
            "he": "hebr",
            "yi": "hebr",
            "th": "thai",
            "hi": "deva",
            "mr": "deva",
            "ne": "deva",
            "ja": "jpan",
            "ko": "hang",
        }.get(language)

    def _font_family(self, fonts: ET.Element, role: str) -> str | _SupplementalFont | None:
        theme_name, direct_name, language_key = {
            "latin": ("hAnsiTheme", "hAnsi", "language"),
            "east_asian": ("eastAsiaTheme", "eastAsia", "east_asian_language"),
            "complex": ("cstheme", "cs", "complex_language"),
        }[role]
        theme = _value(fonts, theme_name)
        if role == "latin":
            theme = theme or _value(fonts, "asciiTheme")
            direct = _value(fonts, direct_name) or _value(fonts, "ascii")
        else:
            direct = _value(fonts, direct_name)
        if not theme:
            return direct
        name = theme.strip().casefold()
        prefix = (
            "major" if name.startswith("major") else "minor" if name.startswith("minor") else None
        )
        if prefix is None:
            return None
        themed_role = (
            "east_asian"
            if name.endswith("eastasia")
            else "complex"
            if name.endswith("bidi")
            else role
        )
        family = self._theme_fonts.get((prefix, themed_role))
        if family is not None:
            return family
        return _SupplementalFont(prefix, themed_role, language_key)

    def font_family(self, properties: dict[str, Any], key: str) -> str | None:
        family = properties.get(key)
        while isinstance(family, _SupplementalFont):
            script = self._script(properties.get(family.language_key))
            resolved = (
                self._theme_fonts.get((family.prefix, f"{family.role}:{script}"))
                if script is not None
                else None
            )
            family = resolved if resolved is not None else family.fallback
        return family if isinstance(family, str) else None

    def paragraph_properties(self, element: ET.Element | None) -> dict[str, Any]:
        properties: dict[str, Any] = {}
        if element is None:
            return properties
        style = _child(element, "pStyle")
        if style is not None and _value(style, "val"):
            properties["pstyle"] = _value(style, "val")
        outline_level = _integer(_value(_child(element, "outlineLvl"), "val"))
        if outline_level is not None:
            properties["outline_level"] = outline_level
        align = _child(element, "jc")
        if align is not None:
            properties["align"] = (_value(align, "val") or "left").strip()
        bidi = _child(element, "bidi")
        if bidi is not None:
            properties["bidi"] = _flag(bidi) is not False
        spacing = _child(element, "spacing")
        if spacing is not None:
            before = _measure(spacing, "before", _POINTS_PER_TWIP)
            after = _measure(spacing, "after", _POINTS_PER_TWIP)
            line = _number(spacing, "line", 1.0)
            rule = _value(spacing, "lineRule")
            if before is not None:
                properties["before"] = before
            if after is not None:
                properties["after"] = after
            if line is not None:
                properties["line"] = line
            if rule:
                properties["line_rule"] = rule.strip()
            # Line-relative spacing is a separate unit from twips: ``w:beforeLines``
            # and ``w:afterLines`` are hundredths of a line, so they resolve against
            # the paragraph's own line height once the lines are broken.
            before_lines = _number(spacing, "beforeLines", 0.01)
            after_lines = _number(spacing, "afterLines", 0.01)
            if before_lines is not None:
                properties["before_lines"] = before_lines
            if after_lines is not None:
                properties["after_lines"] = after_lines
            if _attribute_flag(spacing, "beforeAutospacing") or _attribute_flag(
                spacing, "afterAutospacing"
            ):
                properties["automatic_spacing"] = True
        indent = _child(element, "ind")
        if indent is not None:
            left = _measure(indent, "left", _POINTS_PER_TWIP)
            if left is None:
                left = _measure(indent, "start", _POINTS_PER_TWIP)
            right = _measure(indent, "right", _POINTS_PER_TWIP)
            if right is None:
                right = _measure(indent, "end", _POINTS_PER_TWIP)
            hanging = _measure(indent, "hanging", _POINTS_PER_TWIP)
            first = _measure(indent, "firstLine", _POINTS_PER_TWIP)
            if left is not None:
                properties["left"] = left
            if right is not None:
                properties["right"] = right
            if hanging is not None:
                properties["first_line"] = -hanging
            elif first is not None:
                properties["first_line"] = first
        tabs = _child(element, "tabs")
        if tabs is not None:
            properties["tabs"] = self.tab_stops(tabs)
        for name, key in (
            ("keepNext", "keep_next"),
            ("keepLines", "keep_lines"),
            ("widowControl", "widow_control"),
            ("pageBreakBefore", "page_break_before"),
            ("contextualSpacing", "contextual"),
            ("doNotHyphenate", "do_not_hyphenate"),
        ):
            container = _child(element, name)
            if container is not None:
                properties[key] = _flag(container) is not False
        for name in (
            "autoSpaceDE",
            "autoSpaceDN",
            "kinsoku",
            "wordWrap",
            "overflowPunct",
            "topLinePunct",
            "adjustRightInd",
        ):
            container = _child(element, name)
            if container is not None and _flag(container) is not False:
                properties["east_asian_boundaries"] = True
        alignment = _value(_child(element, "textAlignment"), "val")
        if alignment and alignment.strip().lower() in {"top", "center", "bottom"}:
            properties["line_alignment"] = alignment.strip().lower()
        shading = _child(element, "shd")
        if shading is not None:
            fill = self._color_value(shading, direct="fill")
            if fill and fill.strip().lower() != "auto":
                properties["shading"] = _hex_color(fill)
        numbering = _child(element, "numPr")
        if numbering is not None:
            properties["num_pr"] = True
            instance = _integer(_value(_child(numbering, "numId"), "val"))
            if instance is not None:
                if instance == 0:
                    # ``w:numId`` zero removes the numbering a style would apply.
                    properties["num_off"] = True
                else:
                    properties["num_id"] = instance
            level = _integer(_value(_child(numbering, "ilvl"), "val"))
            if level is not None:
                properties["num_ilvl"] = level
        pbdr = _child(element, "pBdr")
        if pbdr is not None:
            properties["pbrd"] = True
            if _child(pbdr, "between") is not None or _child(pbdr, "bar") is not None:
                properties["border_unsupported"] = True
            b_map = self.borders(pbdr)
            properties["borders"] = b_map
            for s_name in ("top", "bottom", "left", "right"):
                s_el = _child(pbdr, s_name)
                if s_el is not None:
                    if _value(s_el, "shadow") in {"1", "true"} or _value(s_el, "frame") in {
                        "1",
                        "true",
                    }:
                        properties["border_unsupported"] = True
                    k = (_value(s_el, "val") or "none").strip()
                    if k not in {"none", "nil", "single", "thick"}:
                        properties["border_unsupported"] = True
        if _child(element, "framePr") is not None:
            properties["frame"] = True
        return properties

    def tab_stops(self, element: ET.Element) -> tuple[tuple[float, str, str], ...]:
        """Every declared tab stop with the alignment and leader it asks for.

        A declared stop is only consulted when a paragraph actually uses a tab, so
        it is read here without refusing the document: Word writes centre and
        right stops into its default header and footer styles, and documents carry
        stops they never tab to. Nothing about an unused stop can change a
        rendered line, so the alignment and leader of a stop are checked where a
        tab resolves to it, and a declaration without a usable position is not a
        stop at all.
        """

        stops: dict[float, tuple[str, str]] = {}
        for tab in _children(element, "tab"):
            declared = (_value(tab, "val") or "left").strip().lower()
            kind = {"start": "left", "end": "right"}.get(declared, declared)
            declared_leader = (_value(tab, "leader") or "none").strip().lower()
            leader = _TAB_LEADERS.get(declared_leader, declared_leader)
            position = _measure(tab, "pos", _POINTS_PER_TWIP)
            if position is None or position < 0:
                continue
            if kind == "clear":
                stops.pop(position, None)
                continue
            stops[position] = (kind, leader)
        return tuple((position, kind, leader) for position, (kind, leader) in sorted(stops.items()))

    def table_properties(self, element: ET.Element | None) -> dict[str, Any]:
        properties: dict[str, Any] = {}
        if element is None:
            return properties
        style = _child(element, "tblStyle")
        if style is not None and _value(style, "val"):
            properties["tstyle"] = _value(style, "val")
        width = _child(element, "tblW")
        if width is not None:
            # The declared width is kept in its own unit and validated where the
            # table is used: a style this document never applies must not refuse it.
            kind = (_value(width, "type") or "auto").strip()
            value = _integer(_value(width, "w"))
            if value is not None:
                properties["width"] = (kind, value)
        layout = _value(_child(element, "tblLayout"), "type")
        if layout is not None:
            properties["layout"] = layout.strip()
        indent = _measure(_child(element, "tblInd"), "w", _POINTS_PER_TWIP)
        if indent is not None:
            properties["indent"] = indent
        spacing_elem = _child(element, "tblCellSpacing")
        if spacing_elem is not None:
            spacing_type = (_value(spacing_elem, "type") or "dxa").strip().lower()
            spacing_val = _integer(_value(spacing_elem, "w"))
            if spacing_val is not None:
                properties["cell_spacing"] = (spacing_type, spacing_val)
            elif spacing_type in {"nil", "auto"}:
                properties["cell_spacing"] = (spacing_type, 0)
        align = _child(element, "jc")
        if align is not None:
            properties["align"] = (_value(align, "val") or "left").strip()
        borders = self.borders(_child(element, "tblBorders"))
        if borders:
            properties["borders"] = borders
        margin = _child(element, "tblCellMar")
        if margin is not None:
            properties["cell_margin"] = self.margins(margin)
        return properties

    def cell_properties(self, element: ET.Element | None) -> dict[str, Any]:
        """Cell defaults a table carries, including the ones a style inherits."""

        properties: dict[str, Any] = {}
        if element is None:
            return properties
        vertical = _child(element, "vAlign")
        if vertical is not None:
            properties["vertical_align"] = (_value(vertical, "val") or "top").strip().lower()
        no_wrap = _child(element, "noWrap")
        if no_wrap is not None:
            properties["no_wrap"] = _flag(no_wrap) is not False
        return properties

    def margins(self, element: ET.Element) -> tuple[float, float, float, float]:
        values: list[float] = []
        for name in ("top", "left", "bottom", "right"):
            value = _measure(_child(element, name), "w", _POINTS_PER_TWIP)
            values.append(_DEFAULT_CELL_MARGIN if value is None else value)
        return (values[0], values[1], values[2], values[3])

    def borders(self, element: ET.Element | None) -> dict[str, _Border | None]:
        borders: dict[str, _Border | None] = {}
        if element is None:
            return borders
        for name in ("top", "bottom", "left", "right"):
            side = _child(element, name)
            if side is None:
                continue
            kind = (_value(side, "val") or "none").strip()
            if kind in {"none", "nil"}:
                borders[name] = None
                continue
            size = _measure(side, "sz", 1.0 / 8.0)
            space = _measure(side, "space", 1.0) or 0.0
            borders[name] = _Border(
                _hex_color(self._color_value(side, direct="color")),
                size if size else 0.5,
                kind,
                space,
            )
        return borders


def _parse_section_columns(section: ET.Element) -> tuple[int, float, bool]:
    """Parse and validate section columns against native equal-width admission rules.

    Returns (column_count, column_space_in_points, column_separator).
    Raises NativeLayoutError("native_render_section_unsupported") if the column
    configuration is not admitted (e.g. multiple w:cols, unequalWidth, unequal child
    w:col widths, invalid or negative count/space, or invalid child widths).
    """
    cols_elements = _children(section, "cols")
    if len(cols_elements) > 1:
        raise NativeLayoutError("native_render_section_unsupported")
    columns = cols_elements[0] if cols_elements else None
    column_count = 1
    column_space = 36.0
    column_separator = False
    if columns is not None:
        equal_width = _value(columns, "equalWidth")
        if equal_width is not None and equal_width.strip().lower() in {
            "0",
            "false",
            "off",
        }:
            raise NativeLayoutError("native_render_section_unsupported")
        col_children = list(_children(columns, "col"))
        if col_children:
            if equal_width is None or equal_width.strip().lower() in {
                "0",
                "false",
                "off",
            }:
                raise NativeLayoutError("native_render_section_unsupported")
            col_widths = [_measure(c, "w", _POINTS_PER_TWIP) for c in col_children]
            if any(w is None or w <= 0 for w in col_widths):
                raise NativeLayoutError("native_render_section_unsupported")
            if any(abs(w - col_widths[0]) > 0.01 for w in col_widths):
                raise NativeLayoutError("native_render_section_unsupported")
            column_count = len(col_children)
        num_val = _value(columns, "num")
        if num_val is not None:
            num = _integer(num_val)
            if num is None or num < 1:
                raise NativeLayoutError("native_render_section_unsupported")
            column_count = num
        space_val = _measure(columns, "space", _POINTS_PER_TWIP)
        if space_val is not None:
            if space_val < 0:
                raise NativeLayoutError("native_render_section_unsupported")
            column_space = space_val
        if _attribute_flag(columns, "sep"):
            column_separator = True
    return (column_count, column_space, column_separator)


def is_section_column_layout_supported(
    section: ET.Element,
    width: float | None = None,
    left: float | None = None,
    right: float | None = None,
) -> bool:
    """True if native layout renderer admits this section's column configuration.

    Evaluates the section's actual XML element directly against native admission rules:
    equal-width, valid column count >= 1, non-negative spacing, valid child w:col widths,
    and gutters that do not exceed the page text area.
    Width and margins may be supplied in points or twips (if max > 1000 they are treated as twips).
    If omitted, dimensions are measured directly from the section's pgSz and pgMar;
    absent elements use the renderer's named native Letter/one-inch defaults.
    """
    try:
        count, space, _sep = _parse_section_columns(section)
        if count <= 1:
            return True
        if width is None or left is None or right is None:
            pg_sz = _child(section, "pgSz")
            pg_mar = _child(section, "pgMar")
            resolved_width = _measure(pg_sz, "w", _POINTS_PER_TWIP) if pg_sz is not None else 612.0
            resolved_left = (
                _measure(pg_mar, "left", _POINTS_PER_TWIP) if pg_mar is not None else 72.0
            )
            resolved_right = (
                _measure(pg_mar, "right", _POINTS_PER_TWIP) if pg_mar is not None else 72.0
            )
        else:
            scale = _POINTS_PER_TWIP if max(width, left, right) > 1000 else 1.0
            resolved_width = width * scale
            resolved_left = left * scale
            resolved_right = right * scale
        if resolved_width is None or resolved_left is None or resolved_right is None:
            return False
        total_text_width = resolved_width - resolved_left - resolved_right
        gutters = (count - 1) * space
        return (total_text_width - gutters) > 0
    except (NativeLayoutError, ValueError, TypeError, KeyError, AttributeError):
        return False
