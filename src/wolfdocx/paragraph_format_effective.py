"""Effective (style-resolved) paragraph formatting for format discovery.

:func:`wolfdocx.formatting.discover_format_targets` reports each paragraph's direct
``w:pPr`` values. This module adds the value Word applies for the core paragraph
format (alignment, spacing before and after, line spacing, and left, right and
first-line indentation) by layering, lowest precedence first:

1. the WordprocessingML default when nothing sets a value (``source`` ``default``),
2. ``w:docDefaults/w:pPrDefault/w:pPr`` in ``word/styles.xml`` (``doc_defaults``),
3. the paragraph style's ``w:basedOn`` chain, base style first (``style:<id>``); a
   paragraph without ``w:pStyle``, or naming no defined paragraph style, uses the
   default paragraph style,
4. the paragraph's own ``w:pPr`` (``direct``).

Each value is resolved per attribute, so a style that sets only ``w:spacing/@w:after``
leaves ``before`` to the layers below it; ``w:line`` and ``w:lineRule`` inherit
separately (``rule_source`` names the rule's layer), and ``w:hanging`` wins over
``w:firstLine`` within one ``w:ind``. Measures are twentieths of a point (twips)
with ``points`` beside them; an ``auto`` line spacing is in 240ths of a line with
``lines`` beside it. A first-line indent is negative when it is a hanging indent.
A value that is not a plain integer (a strict universal measure such as ``12pt``) is
reported as written, with ``points`` or ``lines`` ``None``. Contextual spacing
between same-style paragraphs is a layout effect and is not applied here.

Contributions this reader does not model are never guessed into a value; they are
listed in ``unresolved`` with the fields they may change:

- ``numbering``: an active numbering instance (direct or style ``w:numPr``) whose
  level indentation Word layers onto the paragraph,
- ``table_style``: paragraph properties in the enclosing table's style,
- ``character_units`` / ``line_units`` / ``autospacing``: ``*Chars``, ``*Lines`` and
  ``*Autospacing`` attributes, which Word converts with font and layout metrics,
- ``missing_style`` / ``missing_based_on`` / ``style_cycle``: a broken style chain;
  the resolved values ignore the missing or repeated link.

``word/styles.xml`` is parsed under the package's XML safety limits: a part with a
DTD or entity declaration, or over a size or depth limit, yields no effective
format and a refusal code instead of being expanded.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Any
from xml.etree import ElementTree as ET

from .inspect import SUPPORTED_WORD_NAMESPACES, Q
from .package import PackageError, PackageGraph

_STYLES = "word/styles.xml"
_INTEGER = re.compile(r"-?\d+\Z")
_OFF = frozenset({"0", "false", "off"})

#: Every reported field, in output order.
FIELDS = (
    "alignment",
    "spacing_before",
    "spacing_after",
    "line_spacing",
    "indent_left",
    "indent_right",
    "indent_first_line",
)
_INDENT_FIELDS = ("indent_left", "indent_right", "indent_first_line")

#: ``(element, attribute)`` pairs that can change each field.
_FIELD_ATTRIBUTES: Mapping[str, tuple[tuple[str, str], ...]] = {
    "alignment": (("jc", "val"),),
    "spacing_before": (
        ("spacing", "before"),
        ("spacing", "beforeLines"),
        ("spacing", "beforeAutospacing"),
    ),
    "spacing_after": (
        ("spacing", "after"),
        ("spacing", "afterLines"),
        ("spacing", "afterAutospacing"),
    ),
    "line_spacing": (("spacing", "line"), ("spacing", "lineRule")),
    "indent_left": (
        ("ind", "start"),
        ("ind", "left"),
        ("ind", "startChars"),
        ("ind", "leftChars"),
    ),
    "indent_right": (
        ("ind", "end"),
        ("ind", "right"),
        ("ind", "endChars"),
        ("ind", "rightChars"),
    ),
    "indent_first_line": (
        ("ind", "firstLine"),
        ("ind", "hanging"),
        ("ind", "firstLineChars"),
        ("ind", "hangingChars"),
    ),
}

_DEFAULTS: Mapping[str, Any] = {
    "alignment": "left",
    "spacing_before": 0,
    "spacing_after": 0,
    "line_spacing": (240, "auto"),
    "indent_left": 0,
    "indent_right": 0,
    "indent_first_line": 0,
}


def _attr(element: ET.Element | None, namespace: str, name: str) -> str | None:
    if element is None:
        return None
    value = element.get(Q(namespace, name))
    return value if value is not None else element.get(name)


def _on(value: str | None) -> bool:
    return value is not None and value.strip().lower() not in _OFF


def _nonzero(value: str | None) -> bool:
    return value is not None and value.strip() not in {"", "0"}


def _measure(value: str) -> int | str:
    """Integer twips when the attribute is a plain integer, else the raw value."""
    stripped = value.strip()
    return int(stripped) if _INTEGER.match(stripped) else value


@dataclass(frozen=True)
class _Layer:
    source: str
    ppr: ET.Element
    namespace: str

    def child(self, local: str) -> ET.Element | None:
        return self.ppr.find(Q(self.namespace, local))

    def get(self, local: str, attribute: str) -> str | None:
        return _attr(self.child(local), self.namespace, attribute)


def _alignment(layer: _Layer) -> str | None:
    return layer.get("jc", "val")


def _spacing(side: str) -> Callable[[_Layer], str | None]:
    return lambda layer: layer.get("spacing", side)


def _indent(primary: str, legacy: str) -> Callable[[_Layer], str | None]:
    def read(layer: _Layer) -> str | None:
        value = layer.get("ind", primary)
        return value if value is not None else layer.get("ind", legacy)

    return read


def _first_line(layer: _Layer) -> str | None:
    hanging = layer.get("ind", "hanging")
    if hanging is not None:
        measure = _measure(hanging)
        return str(-measure) if isinstance(measure, int) else f"-{hanging}"
    return layer.get("ind", "firstLine")


_READERS: Mapping[str, Callable[[_Layer], str | None]] = {
    "alignment": _alignment,
    "spacing_before": _spacing("before"),
    "spacing_after": _spacing("after"),
    "indent_left": _indent("start", "left"),
    "indent_right": _indent("end", "right"),
    "indent_first_line": _first_line,
}

#: Attributes that override a field's twips value with a unit this reader cannot convert.
_UNIT_OVERRIDES: Mapping[str, tuple[tuple[str, str, str, Callable[[str | None], bool]], ...]] = {
    "spacing_before": (
        ("spacing", "beforeAutospacing", "autospacing", _on),
        ("spacing", "beforeLines", "line_units", _nonzero),
    ),
    "spacing_after": (
        ("spacing", "afterAutospacing", "autospacing", _on),
        ("spacing", "afterLines", "line_units", _nonzero),
    ),
    "indent_left": (
        ("ind", "startChars", "character_units", _nonzero),
        ("ind", "leftChars", "character_units", _nonzero),
    ),
    "indent_right": (
        ("ind", "endChars", "character_units", _nonzero),
        ("ind", "rightChars", "character_units", _nonzero),
    ),
    "indent_first_line": (
        ("ind", "firstLineChars", "character_units", _nonzero),
        ("ind", "hangingChars", "character_units", _nonzero),
    ),
}


def _fields_set_by(ppr: ET.Element, namespace: str) -> list[str]:
    layer = _Layer("", ppr, namespace)
    return [
        field
        for field in FIELDS
        if any(
            layer.get(local, attribute) is not None for local, attribute in _FIELD_ATTRIBUTES[field]
        )
    ]


def _with_units(field: str, value: int | str) -> dict[str, Any]:
    record: dict[str, Any] = {"value": value}
    if field != "alignment":
        record["points"] = value / 20 if isinstance(value, int) else None
    return record


class EffectiveParagraphFormat:
    """Resolve effective paragraph formatting against one package's styles part."""

    def __init__(self, styles_root: ET.Element | None) -> None:
        self._namespace = ""
        self._doc_defaults: ET.Element | None = None
        self._styles: dict[tuple[str, str], ET.Element] = {}
        self._default_ids: dict[str, str] = {}
        self._chains: dict[tuple[str, str], tuple[list[tuple[str, ET.Element]], list[dict]]] = {}
        if styles_root is None or not styles_root.tag.startswith("{"):
            return
        namespace = styles_root.tag[1:].split("}", 1)[0]
        if namespace not in SUPPORTED_WORD_NAMESPACES:
            return
        self._namespace = namespace
        word = lambda local: Q(namespace, local)
        self._doc_defaults = styles_root.find(
            f"{word('docDefaults')}/{word('pPrDefault')}/{word('pPr')}"
        )
        for style in styles_root.findall(word("style")):
            style_id = _attr(style, namespace, "styleId")
            style_type = _attr(style, namespace, "type") or "paragraph"
            if not style_id:
                continue
            self._styles.setdefault((style_type, style_id), style)
            if _on(_attr(style, namespace, "default")):
                self._default_ids.setdefault(style_type, style_id)

    @classmethod
    def from_graph(cls, graph: PackageGraph) -> tuple[EffectiveParagraphFormat | None, str | None]:
        """Build a resolver, or return the refusal code of an unsafe or malformed styles part."""
        if _STYLES not in graph.parts:
            return cls(None), None
        try:
            root = graph._parse_member_xml(_STYLES, graph.read_part_bytes(_STYLES))
        except PackageError as error:
            return None, error.code or "format_styles_invalid"
        return cls(root), None

    def _chain(
        self, style_type: str, style_id: str
    ) -> tuple[list[tuple[str, ET.Element]], list[dict]]:
        """``(style id, style element)`` base first, plus broken-link findings."""
        key = (style_type, style_id)
        cached = self._chains.get(key)
        if cached is not None:
            return cached
        chain: list[tuple[str, ET.Element]] = []
        findings: list[dict[str, Any]] = []
        seen: set[str] = set()
        current: str | None = style_id
        while current is not None:
            if current in seen:
                findings.append({"contributor": "style_cycle", "style": current})
                break
            seen.add(current)
            style = self._styles.get((style_type, current))
            if style is None:
                if chain:
                    findings.append({"contributor": "missing_based_on", "style": current})
                break
            chain.append((current, style))
            current = _attr(style.find(Q(self._namespace, "basedOn")), self._namespace, "val")
        chain.reverse()
        self._chains[key] = (chain, findings)
        return chain, findings

    def _style_ppr_layers(self, style_type: str, style_id: str) -> tuple[list[_Layer], list[dict]]:
        chain, findings = self._chain(style_type, style_id)
        layers = []
        for current, style in chain:
            ppr = style.find(Q(self._namespace, "pPr"))
            if ppr is not None:
                layers.append(_Layer(f"style:{current}", ppr, self._namespace))
        return layers, findings

    def _numbering_active(self, layers: list[_Layer]) -> bool:
        """Whether the highest layer naming ``w:numPr/w:numId`` names a real instance."""
        for layer in reversed(layers):
            num_pr = layer.child("numPr")
            num_id = _attr(
                num_pr.find(Q(layer.namespace, "numId")) if num_pr is not None else None,
                layer.namespace,
                "val",
            )
            if num_id is not None:
                return _nonzero(num_id)
        return False

    def _table_style_fields(
        self, table: ET.Element, namespace: str
    ) -> tuple[str | None, list[str]]:
        properties = table.find(Q(namespace, "tblPr"))
        declared = _attr(
            properties.find(Q(namespace, "tblStyle")) if properties is not None else None,
            namespace,
            "val",
        )
        style_id = declared or self._default_ids.get("table")
        if style_id is None or not self._namespace:
            return None, []
        chain, _ = self._chain("table", style_id)
        fields: set[str] = set()
        for _, style in chain:
            for ppr in style.iter(Q(self._namespace, "pPr")):
                fields.update(_fields_set_by(ppr, self._namespace))
        return style_id, [field for field in FIELDS if field in fields]

    def resolve(
        self,
        paragraph: ET.Element,
        namespace: str,
        table: ET.Element | None = None,
    ) -> dict[str, Any]:
        """The effective format record of one ``w:p`` in a part using ``namespace``."""
        direct = paragraph.find(Q(namespace, "pPr"))
        declared = _attr(
            direct.find(Q(namespace, "pStyle")) if direct is not None else None,
            namespace,
            "val",
        )
        unresolved: list[dict[str, Any]] = []
        style_id = declared
        if declared is not None and ("paragraph", declared) not in self._styles:
            unresolved.append(
                {"contributor": "missing_style", "style": declared, "fields": list(FIELDS)}
            )
            style_id = None
        if style_id is None:
            style_id = self._default_ids.get("paragraph")

        layers: list[_Layer] = []
        if self._doc_defaults is not None:
            layers.append(_Layer("doc_defaults", self._doc_defaults, self._namespace))
        style_chain: list[str] = []
        if style_id is not None:
            style_layers, findings = self._style_ppr_layers("paragraph", style_id)
            layers.extend(style_layers)
            style_chain = [current for current, _ in self._chain("paragraph", style_id)[0]]
            style_chain.reverse()
            unresolved.extend({**finding, "fields": list(FIELDS)} for finding in findings)
        if direct is not None:
            layers.append(_Layer("direct", direct, namespace))

        record: dict[str, Any] = {
            "declared_style": declared,
            "style": style_id,
            "style_chain": style_chain,
        }
        for field in FIELDS:
            record[field] = self._resolve_field(field, layers, unresolved)

        if self._numbering_active(layers):
            unresolved.append({"contributor": "numbering", "fields": list(_INDENT_FIELDS)})
        if table is not None:
            table_style, fields = self._table_style_fields(table, namespace)
            if fields:
                unresolved.append(
                    {"contributor": "table_style", "style": table_style, "fields": fields}
                )
        record["unresolved"] = unresolved
        return record

    def _resolve_field(
        self, field: str, layers: list[_Layer], unresolved: list[dict[str, Any]]
    ) -> dict[str, Any]:
        if field == "line_spacing":
            return self._resolve_line(layers)
        reader = _READERS[field]
        value: int | str = _DEFAULTS[field]
        source = "default"
        for layer in layers:
            raw = reader(layer)
            if raw is not None:
                source = layer.source
                value = raw if field == "alignment" else _measure(raw)
        record = {**_with_units(field, value), "source": source}
        # Word lets a *Chars, *Lines or *Autospacing value replace the twips value,
        # including one inherited from a lower layer, so every layer is checked.
        for layer in reversed(layers):
            for local, attribute, contributor, active in _UNIT_OVERRIDES.get(field, ()):
                if active(layer.get(local, attribute)):
                    unresolved.append(
                        {"contributor": contributor, "source": layer.source, "fields": [field]}
                    )
                    return record
        return record

    @staticmethod
    def _resolve_line(layers: list[_Layer]) -> dict[str, Any]:
        """``w:line`` and ``w:lineRule`` inherit separately; an unset rule is ``auto``."""
        value: int | str
        value, rule = _DEFAULTS["line_spacing"]
        source = rule_source = "default"
        for layer in layers:
            raw = layer.get("spacing", "line")
            if raw is not None:
                value, source = _measure(raw), layer.source
            raw_rule = layer.get("spacing", "lineRule")
            if raw_rule is not None:
                rule, rule_source = raw_rule, layer.source
        record: dict[str, Any] = {"value": value, "rule": rule}
        if rule == "auto":
            record["lines"] = value / 240 if isinstance(value, int) else None
        else:
            record["points"] = value / 20 if isinstance(value, int) else None
        record["source"] = source
        record["rule_source"] = rule_source
        return record
