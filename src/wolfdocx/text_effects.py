"""Source-bound Word 2010 (w14) run text effects: authoring and inspection.

Five effects are authored as Word writes them, as ``w14`` children of ``w:rPr``:

- ``glow``: ``w14:glow`` (radius and color).
- ``shadow``: ``w14:shadow``, an outer shadow (blur, distance, direction, scale, skew).
- ``reflection``: ``w14:reflection`` (blur, distance, opacity ramp, size).
- ``outline``: ``w14:textOutline`` (width, color, dash, join).
- ``fill``: ``w14:textFill`` as a linear or path gradient, a solid color, or no fill.

Lengths are points (stored as EMU, 12700 per point), angles are degrees (stored in
60000ths), and percentages are stored in thousandths. Values are quantized to those
units when a request is parsed, so what :func:`inspect_text_effects` reports for a run is
exactly what was authored, and every inspected effect is itself a valid authoring spec.

In ``w14`` color transforms ``w14:alpha`` is a transparency, reported as
``transparency``. Reflection ``stA``/``endA`` are opacities, reported as
``start_opacity``/``end_opacity``.

Edits target a bounded text range in one paragraph exactly as :mod:`wolfdocx.formatting`
run edits do (paragraph index, offsets, expected text, structure fingerprint from
:func:`wolfdocx.formatting.discover_format_targets`). Runs are split at the range
boundaries. Per effect, a spec sets it, ``False`` removes it, and an absent key leaves it.
Anything that cannot be applied exactly is refused with a ``text_effect_*`` code; the
document root gains the ``w14`` namespace and an ``mc:Ignorable`` entry when it lacks them.
"""

from __future__ import annotations

import itertools
import json
import math
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree as ET

from .formatting import (
    _HEX_COLOR_PATTERN,
    _detect_context_unsupported,
    _is_inside_sdt_or_unsupported,
    _local_name,
    _unrelated_parts_preserved,
)
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    _QNameElement,
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
    _publish,
    _review_parts,
    _ReviewContextError,
    _sha,
    _sha256,
    _word_namespace,
)

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_STYLES = "word/styles.xml"
_REFUSED = "refused"
_PUBLISHED = "published"

EMU_PER_POINT = 12700
_ANGLE_UNIT = 60000
_PERCENT_UNIT = 1000

EFFECT_KEYS: tuple[str, ...] = ("glow", "shadow", "reflection", "outline", "fill")

# Order of the w14 run property elements Word writes after the w: elements of CT_RPr.
_W14_RPR_ORDER: tuple[str, ...] = (
    "glow",
    "shadow",
    "reflection",
    "textOutline",
    "textFill",
    "scene3d",
    "props3d",
    "ligatures",
    "numForm",
    "numSpacing",
    "stylisticSets",
    "cntxtAlts",
)
_W14_RANK = {local: index for index, local in enumerate(_W14_RPR_ORDER)}
_ELEMENT_FOR_KEY = {
    "glow": "glow",
    "shadow": "shadow",
    "reflection": "reflection",
    "outline": "textOutline",
    "fill": "textFill",
}
_KEY_FOR_ELEMENT = {element: key for key, element in _ELEMENT_FOR_KEY.items()}

# w14 ST_SchemeColorVal.
SCHEME_COLORS = frozenset(
    {
        "bg1",
        "tx1",
        "bg2",
        "tx2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "phClr",
    }
)
# EG_ColorTransform sequence order; alpha is carried as ``transparency``.
_COLOR_TRANSFORM_ORDER: tuple[str, ...] = (
    "tint",
    "shade",
    "alpha",
    "hueMod",
    "sat",
    "satOff",
    "satMod",
    "lum",
    "lumOff",
    "lumMod",
)
_COLOR_MODIFIERS = frozenset(_COLOR_TRANSFORM_ORDER) - {"alpha"}
_DASHES = frozenset(
    {
        "solid",
        "dot",
        "sysDash",
        "sysDot",
        "dash",
        "dashDot",
        "lgDash",
        "lgDashDot",
        "lgDashDotDot",
        "sysDashDot",
        "sysDashDotDot",
    }
)
_JOINS = frozenset({"round", "bevel", "miter"})
_CAPS = frozenset({"rnd", "sq", "flat"})
_COMPOUNDS = frozenset({"sng", "dbl", "thickThin", "thinThick", "tri"})
_ALIGNMENTS = frozenset({"tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"})
_PEN_ALIGNMENTS = frozenset({"ctr", "in"})
_PATH_SHADES = frozenset({"circle", "rect", "shape"})

_MAX_RADIUS_PT = 150.0
_MAX_BLUR_PT = 100.0
_MAX_DISTANCE_PT = 200.0
_MAX_OUTLINE_PT = 100.0
_MAX_GRADIENT_STOPS = 10


class TextEffectError(ValueError):
    """A malformed effect spec; ``str(error)`` is the ``text_effect_*`` code."""


def _fail(code: str) -> TextEffectError:
    return TextEffectError(f"text_effect_{code}")


def _number(value: Any, code: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise _fail(code)
    return float(value)


def _points(value: Any, code: str, maximum: float) -> int:
    number = _number(value, code)
    if number < 0 or number > maximum:
        raise _fail(code)
    return round(number * EMU_PER_POINT)


def _percent(value: Any, code: str, low: float = 0.0, high: float = 100.0) -> int:
    number = _number(value, code)
    if number < low or number > high:
        raise _fail(code)
    return round(number * _PERCENT_UNIT)


def _angle(value: Any, code: str) -> int:
    number = _number(value, code)
    if number < 0 or number >= 360:
        raise _fail(code)
    return round(number * _ANGLE_UNIT) % (360 * _ANGLE_UNIT)


def _skew(value: Any, code: str) -> int:
    number = _number(value, code)
    if number <= -90 or number >= 90:
        raise _fail(code)
    return round(number * _ANGLE_UNIT)


def _choice(value: Any, allowed: frozenset[str], code: str) -> str:
    """``value`` when it is one of ``allowed``; ``code`` is the full refusal code."""
    if not isinstance(value, str) or value not in allowed:
        raise TextEffectError(code)
    return value


def _keys(value: Any, allowed: set[str], required: set[str], code: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise _fail(f"{code}_invalid")
    unknown = set(value) - allowed
    if unknown:
        raise _fail(f"{code}_unknown_keys")
    if required - set(value):
        raise _fail(f"{code}_missing_keys")
    return value


def _points_out(emu: int) -> float:
    return round(emu / EMU_PER_POINT, 6)


def _units_out(value: int, unit: int) -> float:
    return round(value / unit, 6)


@dataclass(frozen=True)
class EffectColor:
    """An ``srgbClr`` (``RRGGBB``) or ``schemeClr`` color with its transforms."""

    value: str
    transparency: int = 0
    modifiers: tuple[tuple[str, int], ...] = ()

    @property
    def is_scheme(self) -> bool:
        return self.value in SCHEME_COLORS

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "color": self.value,
            "transparency": _units_out(self.transparency, _PERCENT_UNIT),
        }
        if self.modifiers:
            result["color_modifiers"] = dict(self.modifiers)
        return result


def _color_from(spec: Mapping[str, Any], code: str) -> EffectColor:
    raw = spec.get("color")
    if not isinstance(raw, str):
        raise _fail(f"{code}_color_invalid")
    if raw in SCHEME_COLORS:
        value = raw
    elif _HEX_COLOR_PATTERN.match(raw):
        value = raw.upper()
    else:
        raise _fail(f"{code}_color_invalid")
    transparency = _percent(spec.get("transparency", 0), f"{code}_transparency_invalid")
    raw_modifiers = spec.get("color_modifiers", {})
    if not isinstance(raw_modifiers, Mapping):
        raise _fail(f"{code}_color_modifiers_invalid")
    modifiers: list[tuple[str, int]] = []
    for name in _COLOR_TRANSFORM_ORDER:
        if name not in raw_modifiers:
            continue
        amount = raw_modifiers[name]
        if isinstance(amount, bool) or not isinstance(amount, int) or abs(amount) > 10_000_000:
            raise _fail(f"{code}_color_modifiers_invalid")
        modifiers.append((name, amount))
    if set(raw_modifiers) - _COLOR_MODIFIERS:
        raise _fail(f"{code}_color_modifiers_invalid")
    return EffectColor(value, transparency, tuple(modifiers))


_COLOR_KEYS = {"color", "transparency", "color_modifiers"}


@dataclass(frozen=True)
class Glow:
    radius: int
    color: EffectColor

    @classmethod
    def from_dict(cls, value: Any) -> Glow:
        spec = _keys(value, {"radius"} | _COLOR_KEYS, {"radius", "color"}, "glow")
        return cls(
            radius=_points(spec["radius"], "glow_radius_invalid", _MAX_RADIUS_PT),
            color=_color_from(spec, "glow"),
        )

    def to_dict(self) -> dict[str, Any]:
        return {"radius": _points_out(self.radius), **self.color.to_dict()}


@dataclass(frozen=True)
class OuterShadow:
    color: EffectColor
    blur: int = round(4 * EMU_PER_POINT)
    distance: int = round(3 * EMU_PER_POINT)
    direction: int = 45 * _ANGLE_UNIT
    scale_x: int = 100 * _PERCENT_UNIT
    scale_y: int = 100 * _PERCENT_UNIT
    skew_x: int = 0
    skew_y: int = 0
    alignment: str = "tl"

    @classmethod
    def from_dict(cls, value: Any) -> OuterShadow:
        allowed = {
            "blur",
            "distance",
            "direction",
            "scale_x",
            "scale_y",
            "skew_x",
            "skew_y",
            "alignment",
        } | _COLOR_KEYS
        spec = _keys(value, allowed, {"color"}, "shadow")
        defaults = cls(color=EffectColor("000000"))
        return cls(
            color=_color_from(spec, "shadow"),
            blur=_points(spec["blur"], "shadow_blur_invalid", _MAX_BLUR_PT)
            if "blur" in spec
            else defaults.blur,
            distance=_points(spec["distance"], "shadow_distance_invalid", _MAX_DISTANCE_PT)
            if "distance" in spec
            else defaults.distance,
            direction=_angle(spec["direction"], "shadow_direction_invalid")
            if "direction" in spec
            else defaults.direction,
            scale_x=_percent(spec["scale_x"], "shadow_scale_invalid", -200, 200)
            if "scale_x" in spec
            else defaults.scale_x,
            scale_y=_percent(spec["scale_y"], "shadow_scale_invalid", -200, 200)
            if "scale_y" in spec
            else defaults.scale_y,
            skew_x=_skew(spec["skew_x"], "shadow_skew_invalid")
            if "skew_x" in spec
            else defaults.skew_x,
            skew_y=_skew(spec["skew_y"], "shadow_skew_invalid")
            if "skew_y" in spec
            else defaults.skew_y,
            alignment=_choice(
                spec["alignment"], _ALIGNMENTS, "text_effect_shadow_alignment_invalid"
            )
            if "alignment" in spec
            else defaults.alignment,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "blur": _points_out(self.blur),
            "distance": _points_out(self.distance),
            "direction": _units_out(self.direction, _ANGLE_UNIT),
            "scale_x": _units_out(self.scale_x, _PERCENT_UNIT),
            "scale_y": _units_out(self.scale_y, _PERCENT_UNIT),
            "skew_x": _units_out(self.skew_x, _ANGLE_UNIT),
            "skew_y": _units_out(self.skew_y, _ANGLE_UNIT),
            "alignment": self.alignment,
            **self.color.to_dict(),
        }


@dataclass(frozen=True)
class Reflection:
    blur: int = round(0.5 * EMU_PER_POINT)
    distance: int = 0
    start_opacity: int = 50 * _PERCENT_UNIT
    start_position: int = 0
    end_opacity: int = 300
    end_position: int = 55 * _PERCENT_UNIT
    direction: int = 90 * _ANGLE_UNIT
    fade_direction: int = 90 * _ANGLE_UNIT
    scale_x: int = 100 * _PERCENT_UNIT
    scale_y: int = -100 * _PERCENT_UNIT
    skew_x: int = 0
    skew_y: int = 0
    alignment: str = "bl"

    @classmethod
    def from_dict(cls, value: Any) -> Reflection:
        allowed = {
            "blur",
            "distance",
            "start_opacity",
            "start_position",
            "end_opacity",
            "size",
            "direction",
            "fade_direction",
            "scale_x",
            "scale_y",
            "skew_x",
            "skew_y",
            "alignment",
        }
        spec = _keys(value, allowed, set(), "reflection")
        base = cls()

        def pick(key: str, convert: Any, default: int) -> int:
            return convert(spec[key]) if key in spec else default

        result = cls(
            blur=pick(
                "blur",
                lambda v: _points(v, "reflection_blur_invalid", _MAX_BLUR_PT),
                base.blur,
            ),
            distance=pick(
                "distance",
                lambda v: _points(v, "reflection_distance_invalid", _MAX_DISTANCE_PT),
                base.distance,
            ),
            start_opacity=pick(
                "start_opacity",
                lambda v: _percent(v, "reflection_opacity_invalid"),
                base.start_opacity,
            ),
            start_position=pick(
                "start_position",
                lambda v: _percent(v, "reflection_position_invalid"),
                base.start_position,
            ),
            end_opacity=pick(
                "end_opacity",
                lambda v: _percent(v, "reflection_opacity_invalid"),
                base.end_opacity,
            ),
            end_position=pick(
                "size",
                lambda v: _percent(v, "reflection_position_invalid"),
                base.end_position,
            ),
            direction=pick(
                "direction",
                lambda v: _angle(v, "reflection_direction_invalid"),
                base.direction,
            ),
            fade_direction=pick(
                "fade_direction",
                lambda v: _angle(v, "reflection_direction_invalid"),
                base.fade_direction,
            ),
            scale_x=pick(
                "scale_x",
                lambda v: _percent(v, "reflection_scale_invalid", -200, 200),
                base.scale_x,
            ),
            scale_y=pick(
                "scale_y",
                lambda v: _percent(v, "reflection_scale_invalid", -200, 200),
                base.scale_y,
            ),
            skew_x=pick("skew_x", lambda v: _skew(v, "reflection_skew_invalid"), base.skew_x),
            skew_y=pick("skew_y", lambda v: _skew(v, "reflection_skew_invalid"), base.skew_y),
            alignment=_choice(
                spec["alignment"],
                _ALIGNMENTS,
                "text_effect_reflection_alignment_invalid",
            )
            if "alignment" in spec
            else base.alignment,
        )
        if result.end_position < result.start_position:
            raise _fail("reflection_position_invalid")
        return result

    def to_dict(self) -> dict[str, Any]:
        return {
            "blur": _points_out(self.blur),
            "distance": _points_out(self.distance),
            "start_opacity": _units_out(self.start_opacity, _PERCENT_UNIT),
            "start_position": _units_out(self.start_position, _PERCENT_UNIT),
            "end_opacity": _units_out(self.end_opacity, _PERCENT_UNIT),
            "size": _units_out(self.end_position, _PERCENT_UNIT),
            "direction": _units_out(self.direction, _ANGLE_UNIT),
            "fade_direction": _units_out(self.fade_direction, _ANGLE_UNIT),
            "scale_x": _units_out(self.scale_x, _PERCENT_UNIT),
            "scale_y": _units_out(self.scale_y, _PERCENT_UNIT),
            "skew_x": _units_out(self.skew_x, _ANGLE_UNIT),
            "skew_y": _units_out(self.skew_y, _ANGLE_UNIT),
            "alignment": self.alignment,
        }


@dataclass(frozen=True)
class GradientStop:
    position: int
    color: EffectColor

    def to_dict(self) -> dict[str, Any]:
        return {
            "position": _units_out(self.position, _PERCENT_UNIT),
            **self.color.to_dict(),
        }


@dataclass(frozen=True)
class Fill:
    """``w14:textFill`` (or the fill of ``w14:textOutline``).

    ``kind`` is ``gradient``, ``solid``, or ``none``. A gradient has ``stops`` and either
    a linear ``angle`` or a ``path`` shade (``circle``, ``rect``, ``shape``).
    """

    kind: Literal["gradient", "solid", "none"]
    color: EffectColor | None = None
    stops: tuple[GradientStop, ...] = ()
    angle: int | None = None
    scaled: bool = False
    path: str | None = None

    @classmethod
    def from_dict(cls, value: Any, code: str = "fill") -> Fill:
        if not isinstance(value, Mapping):
            raise _fail(f"{code}_invalid")
        kind = value.get("type")
        if kind == "none":
            _keys(value, {"type"}, {"type"}, code)
            return cls("none")
        if kind == "solid":
            spec = _keys(value, {"type"} | _COLOR_KEYS, {"type", "color"}, code)
            return cls("solid", color=_color_from(spec, code))
        if kind != "gradient":
            raise _fail(f"{code}_type_invalid")
        spec = _keys(value, {"type", "stops", "angle", "scaled", "path"}, {"type", "stops"}, code)
        raw_stops = spec["stops"]
        if (
            not isinstance(raw_stops, Sequence)
            or isinstance(raw_stops, (str, bytes))
            or not 2 <= len(raw_stops) <= _MAX_GRADIENT_STOPS
        ):
            raise _fail(f"{code}_stops_invalid")
        stops: list[GradientStop] = []
        for raw in raw_stops:
            stop = _keys(raw, {"position"} | _COLOR_KEYS, {"position", "color"}, f"{code}_stop")
            stops.append(
                GradientStop(
                    position=_percent(stop["position"], f"{code}_stop_position_invalid"),
                    color=_color_from(stop, f"{code}_stop"),
                )
            )
        if any(a.position > b.position for a, b in itertools.pairwise(stops)):
            raise _fail(f"{code}_stops_unordered")
        if "path" in spec and "angle" in spec:
            raise _fail(f"{code}_shade_conflict")
        path = (
            _choice(spec["path"], _PATH_SHADES, f"text_effect_{code}_path_invalid")
            if "path" in spec
            else None
        )
        scaled = spec.get("scaled", False)
        if not isinstance(scaled, bool):
            raise _fail(f"{code}_scaled_invalid")
        angle = None if path else _angle(spec.get("angle", 0), f"{code}_angle_invalid")
        return cls("gradient", stops=tuple(stops), angle=angle, scaled=scaled, path=path)

    def to_dict(self) -> dict[str, Any]:
        if self.kind == "none":
            return {"type": "none"}
        if self.kind == "solid":
            assert self.color is not None
            return {"type": "solid", **self.color.to_dict()}
        result: dict[str, Any] = {
            "type": "gradient",
            "stops": [stop.to_dict() for stop in self.stops],
        }
        if self.path is not None:
            result["path"] = self.path
        else:
            assert self.angle is not None
            result["angle"] = _units_out(self.angle, _ANGLE_UNIT)
            result["scaled"] = self.scaled
        return result


@dataclass(frozen=True)
class Outline:
    width: int
    fill: Fill
    dash: str = "solid"
    join: str = "round"
    cap: str = "flat"
    compound: str = "sng"
    alignment: str = "ctr"

    @classmethod
    def from_dict(cls, value: Any) -> Outline:
        allowed = {
            "width",
            "fill",
            "dash",
            "join",
            "cap",
            "compound",
            "alignment",
        } | _COLOR_KEYS
        spec = _keys(value, allowed, {"width"}, "outline")
        if ("fill" in spec) == ("color" in spec):
            raise _fail("outline_fill_invalid")
        if "fill" in spec:
            if set(spec) & {"transparency", "color_modifiers"}:
                raise _fail("outline_fill_invalid")
            fill = Fill.from_dict(spec["fill"], "outline_fill")
        else:
            fill = Fill("solid", color=_color_from(spec, "outline"))
        return cls(
            width=_points(spec["width"], "outline_width_invalid", _MAX_OUTLINE_PT),
            fill=fill,
            dash=_choice(spec.get("dash", "solid"), _DASHES, "text_effect_outline_dash_invalid"),
            join=_choice(spec.get("join", "round"), _JOINS, "text_effect_outline_join_invalid"),
            cap=_choice(spec.get("cap", "flat"), _CAPS, "text_effect_outline_cap_invalid"),
            compound=_choice(
                spec.get("compound", "sng"),
                _COMPOUNDS,
                "text_effect_outline_compound_invalid",
            ),
            alignment=_choice(
                spec.get("alignment", "ctr"),
                _PEN_ALIGNMENTS,
                "text_effect_outline_alignment_invalid",
            ),
        )

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "width": _points_out(self.width),
            "dash": self.dash,
            "join": self.join,
            "cap": self.cap,
            "compound": self.compound,
            "alignment": self.alignment,
        }
        if self.fill.kind == "solid":
            assert self.fill.color is not None
            result.update(self.fill.color.to_dict())
        else:
            result["fill"] = self.fill.to_dict()
        return result


_EFFECT_PARSERS: dict[str, Any] = {
    "glow": Glow.from_dict,
    "shadow": OuterShadow.from_dict,
    "reflection": Reflection.from_dict,
    "outline": Outline.from_dict,
    "fill": Fill.from_dict,
}

Effect = Glow | OuterShadow | Reflection | Outline | Fill


@dataclass(frozen=True)
class TextEffects:
    """Per effect: a spec sets it, ``False`` removes it, ``None`` leaves it unchanged."""

    glow: Glow | Literal[False] | None = None
    shadow: OuterShadow | Literal[False] | None = None
    reflection: Reflection | Literal[False] | None = None
    outline: Outline | Literal[False] | None = None
    fill: Fill | Literal[False] | None = None

    @classmethod
    def from_dict(cls, value: Any) -> TextEffects:
        spec = _keys(value, set(EFFECT_KEYS), set(), "effects")
        parsed: dict[str, Any] = {}
        for key, raw in spec.items():
            if raw is False:
                parsed[key] = False
            elif raw is None:
                continue
            else:
                parsed[key] = _EFFECT_PARSERS[key](raw)
        result = cls(**parsed)
        if result.is_empty():
            raise _fail("effects_missing")
        return result

    def items(self) -> list[tuple[str, Effect | Literal[False]]]:
        return [(key, getattr(self, key)) for key in EFFECT_KEYS if getattr(self, key) is not None]

    def is_empty(self) -> bool:
        return not self.items()

    def to_dict(self) -> dict[str, Any]:
        return {key: (False if value is False else value.to_dict()) for key, value in self.items()}


@dataclass(frozen=True)
class TextEffectEdit:
    """Text effects applied to ``expected_old_text`` at ``start:end`` of one paragraph."""

    paragraph: int
    start: int
    end: int
    expected_old_text: str
    structure_fingerprint_sha256: str
    effects: TextEffects
    part: str = _DOCUMENT

    def to_dict(self) -> dict[str, Any]:
        return {
            "part": self.part,
            "paragraph": self.paragraph,
            "start": self.start,
            "end": self.end,
            "expected_old_text": self.expected_old_text,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "effects": self.effects.to_dict(),
        }


@dataclass(frozen=True)
class TextEffectRequest:
    source: Path
    source_sha256: str
    destination: Path
    edits: tuple[TextEffectEdit, ...] = ()


@dataclass(frozen=True)
class TextEffectReport:
    """Content-safe receipt for one atomic text-effect authoring request."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    edits: list[dict[str, Any]] = field(default_factory=list)

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
            "edits": self.edits,
        }


def _parse_edit(value: Any) -> TextEffectEdit:
    spec = _keys(
        value,
        {
            "part",
            "paragraph",
            "start",
            "end",
            "expected_old_text",
            "structure_fingerprint_sha256",
            "effects",
        },
        {
            "paragraph",
            "start",
            "end",
            "expected_old_text",
            "structure_fingerprint_sha256",
            "effects",
        },
        "edit",
    )
    paragraph = spec["paragraph"]
    if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 1:
        raise _fail("paragraph_invalid")
    start, end = spec["start"], spec["end"]
    if (
        isinstance(start, bool)
        or not isinstance(start, int)
        or isinstance(end, bool)
        or not isinstance(end, int)
        or start < 0
        or end <= start
    ):
        raise _fail("range_invalid")
    expected = spec["expected_old_text"]
    if not isinstance(expected, str) or len(expected) != end - start:
        raise _fail("expected_text_invalid")
    part = spec.get("part", _DOCUMENT)
    if not isinstance(part, str) or not part:
        raise _fail("part_invalid")
    try:
        fingerprint = _sha(
            spec["structure_fingerprint_sha256"], "text_effect_structure_fingerprint"
        )
    except ValueError as error:
        raise TextEffectError(str(error)) from None
    return TextEffectEdit(
        paragraph=paragraph,
        start=start,
        end=end,
        expected_old_text=expected,
        structure_fingerprint_sha256=fingerprint,
        effects=TextEffects.from_dict(spec["effects"]),
        part=part,
    )


def parse_text_effect_request(value: Any) -> TextEffectRequest:
    """Parse the strict JSON request; malformed input raises :class:`TextEffectError`."""
    spec = _keys(
        value,
        {"source", "source_sha256", "destination", "edits"},
        {"source", "source_sha256", "destination", "edits"},
        "request",
    )
    source, destination = spec["source"], spec["destination"]
    if not isinstance(source, str) or not source:
        raise _fail("source_path_invalid")
    if not isinstance(destination, str) or not destination:
        raise _fail("destination_path_invalid")
    try:
        source_sha256 = _sha(spec["source_sha256"], "text_effect_source_sha256")
    except ValueError as error:
        raise TextEffectError(str(error)) from None
    raw_edits = spec["edits"]
    if not isinstance(raw_edits, Sequence) or isinstance(raw_edits, (str, bytes)) or not raw_edits:
        raise _fail("edits_missing")
    return TextEffectRequest(
        source=Path(source),
        source_sha256=source_sha256,
        destination=Path(destination),
        edits=tuple(_parse_edit(item) for item in raw_edits),
    )


def load_text_effect_request(path: str | Path) -> TextEffectRequest:
    with Path(path).open(encoding="utf-8") as stream:
        return parse_text_effect_request(json.load(stream))


# --- XML writing ---------------------------------------------------------------------


def _w14(local: str) -> str:
    return Q(W14, local)


def _color_element(color: EffectColor) -> ET.Element:
    element = ET.Element(
        _w14("schemeClr" if color.is_scheme else "srgbClr"), {_w14("val"): color.value}
    )
    transforms = dict(color.modifiers)
    if color.transparency:
        transforms["alpha"] = color.transparency
    for name in _COLOR_TRANSFORM_ORDER:
        if name in transforms:
            ET.SubElement(element, _w14(name), {_w14("val"): str(transforms[name])})
    return element


def _fill_children(fill: Fill) -> list[ET.Element]:
    if fill.kind == "none":
        return [ET.Element(_w14("noFill"))]
    if fill.kind == "solid":
        assert fill.color is not None
        solid = ET.Element(_w14("solidFill"))
        solid.append(_color_element(fill.color))
        return [solid]
    gradient = ET.Element(_w14("gradFill"))
    stops = ET.SubElement(gradient, _w14("gsLst"))
    for stop in fill.stops:
        node = ET.SubElement(stops, _w14("gs"), {_w14("pos"): str(stop.position)})
        node.append(_color_element(stop.color))
    if fill.path is not None:
        shade = ET.SubElement(gradient, _w14("path"), {_w14("path"): fill.path})
        ET.SubElement(
            shade,
            _w14("fillToRect"),
            {
                _w14("l"): "50000",
                _w14("t"): "50000",
                _w14("r"): "50000",
                _w14("b"): "50000",
            },
        )
    else:
        ET.SubElement(
            gradient,
            _w14("lin"),
            {_w14("ang"): str(fill.angle), _w14("scaled"): "1" if fill.scaled else "0"},
        )
    return [gradient]


def _attrs(pairs: Sequence[tuple[str, Any]]) -> dict[str, str]:
    return {_w14(name): str(value) for name, value in pairs}


def effect_element(key: str, effect: Effect) -> ET.Element:
    """The ``w14`` element Word writes for one effect."""
    if isinstance(effect, Glow):
        element = ET.Element(_w14("glow"), _attrs([("rad", effect.radius)]))
        element.append(_color_element(effect.color))
        return element
    if isinstance(effect, OuterShadow):
        element = ET.Element(
            _w14("shadow"),
            _attrs(
                [
                    ("blurRad", effect.blur),
                    ("dist", effect.distance),
                    ("dir", effect.direction),
                    ("sx", effect.scale_x),
                    ("sy", effect.scale_y),
                    ("kx", effect.skew_x),
                    ("ky", effect.skew_y),
                    ("algn", effect.alignment),
                ]
            ),
        )
        element.append(_color_element(effect.color))
        return element
    if isinstance(effect, Reflection):
        return ET.Element(
            _w14("reflection"),
            _attrs(
                [
                    ("blurRad", effect.blur),
                    ("stA", effect.start_opacity),
                    ("stPos", effect.start_position),
                    ("endA", effect.end_opacity),
                    ("endPos", effect.end_position),
                    ("dist", effect.distance),
                    ("dir", effect.direction),
                    ("fadeDir", effect.fade_direction),
                    ("sx", effect.scale_x),
                    ("sy", effect.scale_y),
                    ("kx", effect.skew_x),
                    ("ky", effect.skew_y),
                    ("algn", effect.alignment),
                ]
            ),
        )
    if isinstance(effect, Outline):
        element = ET.Element(
            _w14("textOutline"),
            _attrs(
                [
                    ("w", effect.width),
                    ("cap", effect.cap),
                    ("cmpd", effect.compound),
                    ("algn", effect.alignment),
                ]
            ),
        )
        element.extend(_fill_children(effect.fill))
        ET.SubElement(element, _w14("prstDash"), {_w14("val"): effect.dash})
        if effect.join == "miter":
            ET.SubElement(element, _w14("miter"), {_w14("lim"): "800000"})
        else:
            ET.SubElement(element, _w14(effect.join))
        return element
    assert isinstance(effect, Fill) and key == "fill"
    element = ET.Element(_w14("textFill"))
    element.extend(_fill_children(effect))
    return element


def _apply_effects(run: ET.Element, effects: TextEffects, namespace: str) -> None:
    rpr = run.find(Q(namespace, "rPr"))
    if rpr is None:
        rpr = ET.Element(Q(namespace, "rPr"))
        run.insert(0, rpr)
    for key, value in effects.items():
        for existing in rpr.findall(_w14(_ELEMENT_FOR_KEY[key])):
            rpr.remove(existing)
        if value is not False:
            rpr.append(effect_element(key, value))
    word = [child for child in rpr if not child.tag.startswith(f"{{{W14}}}")]
    trailing = [child for child in word if _local_name(child.tag) == "rPrChange"]
    leading = [child for child in word if _local_name(child.tag) != "rPrChange"]
    w14 = sorted(
        (child for child in rpr if child.tag.startswith(f"{{{W14}}}")),
        key=lambda child: _W14_RANK.get(_local_name(child.tag), len(_W14_RANK)),
    )
    rpr[:] = leading + w14 + trailing
    if not len(rpr):
        run.remove(rpr)


def _declare_w14(root: ET.Element) -> None:
    """Bind ``w14`` on the root and list it in ``mc:Ignorable`` (Word 2007 readers)."""
    if not isinstance(root, _QNameElement):
        return
    scope = dict(root._namespace_scope)
    prefix = next((p for p, uri in scope.items() if p and uri == W14), None)
    if prefix is None:
        prefix = "w14"
        index = 0
        while prefix in scope:
            index += 1
            prefix = f"w14x{index}"
        scope[prefix] = W14
    if not any(p and uri == MC for p, uri in scope.items()):
        mc_prefix = "mc"
        index = 0
        while mc_prefix in scope:
            index += 1
            mc_prefix = f"mc{index}"
        scope[mc_prefix] = MC
    # The attribute's inferred type is its two-binding default; any binding tuple is valid.
    root._namespace_scope = tuple(scope.items())  # type: ignore[assignment]
    ignorable = (root.get(Q(MC, "Ignorable")) or "").split()
    if prefix not in ignorable:
        ignorable.append(prefix)
        root.set(Q(MC, "Ignorable"), " ".join(ignorable))


# --- Reading -------------------------------------------------------------------------


def _int_attr(element: ET.Element, local: str, default: int) -> int:
    raw = element.get(_w14(local))
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _read_color(parent: ET.Element | None) -> EffectColor | None:
    if parent is None:
        return None
    for child in parent:
        local = _local_name(child.tag)
        if not child.tag.startswith(f"{{{W14}}}") or local not in {
            "srgbClr",
            "schemeClr",
        }:
            continue
        value = child.get(_w14("val")) or ""
        if local == "srgbClr":
            value = value.upper()
        transparency = 0
        modifiers: list[tuple[str, int]] = []
        for transform in child:
            name = _local_name(transform.tag)
            amount = _int_attr(transform, "val", 0)
            if name == "alpha":
                transparency = amount
            elif name in _COLOR_MODIFIERS:
                modifiers.append((name, amount))
        modifiers.sort(key=lambda item: _COLOR_TRANSFORM_ORDER.index(item[0]))
        return EffectColor(value, transparency, tuple(modifiers))
    return None


def _read_fill(parent: ET.Element) -> Fill | None:
    for child in parent:
        local = _local_name(child.tag)
        if local == "noFill":
            return Fill("none")
        if local == "solidFill":
            color = _read_color(child)
            return Fill("solid", color=color) if color else None
        if local == "gradFill":
            stops: list[GradientStop] = []
            for stop in child.iter(_w14("gs")):
                color = _read_color(stop)
                if color is not None:
                    stops.append(GradientStop(_int_attr(stop, "pos", 0), color))
            stops.sort(key=lambda stop: stop.position)
            linear = child.find(_w14("lin"))
            shade = child.find(_w14("path"))
            if shade is not None:
                return Fill(
                    "gradient",
                    stops=tuple(stops),
                    path=shade.get(_w14("path"), "circle"),
                )
            angle = _int_attr(linear, "ang", 0) if linear is not None else 0
            scaled = linear is not None and linear.get(_w14("scaled")) in {
                "1",
                "true",
                "on",
            }
            return Fill("gradient", stops=tuple(stops), angle=angle, scaled=scaled)
    return None


def read_effect(element: ET.Element) -> tuple[str, Effect] | None:
    """Parse one ``w14`` run property element, or ``None`` when it is not one of the five."""
    local = _local_name(element.tag)
    if not element.tag.startswith(f"{{{W14}}}") or local not in _KEY_FOR_ELEMENT:
        return None
    key = _KEY_FOR_ELEMENT[local]
    if local == "glow":
        color = _read_color(element)
        return (key, Glow(_int_attr(element, "rad", 0), color)) if color else None
    if local == "shadow":
        color = _read_color(element)
        if color is None:
            return None
        return key, OuterShadow(
            color=color,
            blur=_int_attr(element, "blurRad", 0),
            distance=_int_attr(element, "dist", 0),
            direction=_int_attr(element, "dir", 0),
            scale_x=_int_attr(element, "sx", 100_000),
            scale_y=_int_attr(element, "sy", 100_000),
            skew_x=_int_attr(element, "kx", 0),
            skew_y=_int_attr(element, "ky", 0),
            alignment=element.get(_w14("algn"), "b"),
        )
    if local == "reflection":
        return key, Reflection(
            blur=_int_attr(element, "blurRad", 0),
            distance=_int_attr(element, "dist", 0),
            start_opacity=_int_attr(element, "stA", 100_000),
            start_position=_int_attr(element, "stPos", 0),
            end_opacity=_int_attr(element, "endA", 0),
            end_position=_int_attr(element, "endPos", 100_000),
            direction=_int_attr(element, "dir", 0),
            fade_direction=_int_attr(element, "fadeDir", 5_400_000),
            scale_x=_int_attr(element, "sx", 100_000),
            scale_y=_int_attr(element, "sy", 100_000),
            skew_x=_int_attr(element, "kx", 0),
            skew_y=_int_attr(element, "ky", 0),
            alignment=element.get(_w14("algn"), "b"),
        )
    if local == "textOutline":
        fill = _read_fill(element) or Fill("none")
        dash = element.find(_w14("prstDash"))
        join = next(
            (_local_name(child.tag) for child in element if _local_name(child.tag) in _JOINS),
            "round",
        )
        return key, Outline(
            width=_int_attr(element, "w", 0),
            fill=fill,
            dash=dash.get(_w14("val"), "solid") if dash is not None else "solid",
            join=join,
            cap=element.get(_w14("cap"), "flat"),
            compound=element.get(_w14("cmpd"), "sng"),
            alignment=element.get(_w14("algn"), "ctr"),
        )
    text_fill = _read_fill(element)
    return (key, text_fill) if text_fill else None


def _effects_of(rpr: ET.Element | None) -> dict[str, Effect]:
    effects: dict[str, Effect] = {}
    if rpr is None:
        return effects
    for child in rpr:
        parsed = read_effect(child)
        if parsed is not None:
            effects[parsed[0]] = parsed[1]
    return effects


def _style_effects(
    styles_root: ET.Element | None, namespace: str
) -> tuple[dict[str, tuple[str | None, dict[str, Effect]]], dict[str, Effect]]:
    if styles_root is None:
        return {}, {}

    def word(local: str) -> str:
        return Q(namespace, local)

    defaults = _effects_of(
        styles_root.find(f"{word('docDefaults')}/{word('rPrDefault')}/{word('rPr')}")
    )
    styles: dict[str, tuple[str | None, dict[str, Effect]]] = {}
    for style in styles_root.findall(word("style")):
        style_id = style.get(word("styleId"))
        if not style_id:
            continue
        based_on = style.find(word("basedOn"))
        styles[style_id] = (
            based_on.get(word("val")) if based_on is not None else None,
            _effects_of(style.find(word("rPr"))),
        )
    return styles, defaults


def _chain_effects(
    style_id: str | None, styles: Mapping[str, tuple[str | None, dict[str, Effect]]]
) -> dict[str, Effect]:
    chain: list[dict[str, Effect]] = []
    seen: set[str] = set()
    while style_id and style_id not in seen and style_id in styles:
        seen.add(style_id)
        based_on, effects = styles[style_id]
        chain.append(effects)
        style_id = based_on
    resolved: dict[str, Effect] = {}
    for effects in reversed(chain):
        resolved.update(effects)
    return resolved


def _as_dict(effects: Mapping[str, Effect]) -> dict[str, Any]:
    return {key: effects[key].to_dict() for key in EFFECT_KEYS if key in effects}


def inspect_text_effects(source: str | Path, *, part: str = _DOCUMENT) -> dict[str, Any]:
    """Report the five text effects of every run in ``part``.

    ``runs`` lists every ``w:r`` with direct or inherited effects, in document order:
    1-based ``paragraph`` (as in :func:`wolfdocx.formatting.discover_format_targets`),
    0-based ``run`` within the paragraph, the run's ``start``/``end`` text offsets,
    ``effects`` (direct ``w:rPr``) and ``resolved_effects`` (document defaults, the
    paragraph style chain, the character style chain, then direct, each effect replacing
    the one before). Each effect dict is a valid authoring spec.
    """
    path = Path(source)
    failure = {"passed": False, "part": part, "runs": []}
    try:
        graph = PackageGraph.from_file(path)
        document_root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    except (OSError, PackageError, zipfile.BadZipFile, ET.ParseError, KeyError):
        return {**failure, "code": "text_effect_package_invalid"}
    namespace = _word_namespace(document_root)
    if namespace is None:
        return {**failure, "code": "text_effect_namespace_unsupported"}
    if part not in _review_parts(graph, document_root, namespace):
        return {**failure, "code": "text_effect_part_unsupported"}
    try:
        root = ET.fromstring(graph.read_part_bytes(part))
        styles_root = (
            ET.fromstring(graph.read_part_bytes(_STYLES)) if _STYLES in graph.parts else None
        )
    except (OSError, PackageError, ET.ParseError, KeyError):
        return {**failure, "code": "text_effect_package_invalid"}
    styles, defaults = _style_effects(styles_root, namespace)

    def word(local: str) -> str:
        return Q(namespace, local)

    runs: list[dict[str, Any]] = []
    for paragraph_index, paragraph in enumerate(root.iter(word("p")), start=1):
        ppr = paragraph.find(word("pPr"))
        paragraph_style = ppr.find(word("pStyle")) if ppr is not None else None
        paragraph_effects = dict(defaults)
        paragraph_effects.update(
            _chain_effects(
                paragraph_style.get(word("val")) if paragraph_style is not None else None,
                styles,
            )
        )
        offset = 0
        for run_index, run in enumerate(paragraph.iter(word("r"))):
            text = "".join(node.text or "" for node in run.iter(word("t")))
            start, offset = offset, offset + len(text)
            rpr = run.find(word("rPr"))
            character_style = rpr.find(word("rStyle")) if rpr is not None else None
            style_id = character_style.get(word("val")) if character_style is not None else None
            direct = _effects_of(rpr)
            resolved = dict(paragraph_effects)
            resolved.update(_chain_effects(style_id, styles))
            resolved.update(direct)
            if not resolved:
                continue
            runs.append(
                {
                    "paragraph": paragraph_index,
                    "run": run_index,
                    "start": start,
                    "end": offset,
                    "text": text,
                    "character_style": style_id,
                    "effects": _as_dict(direct),
                    "resolved_effects": _as_dict(resolved),
                }
            )
    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        "part": part,
        "runs": runs,
    }


# --- Authoring -----------------------------------------------------------------------


def _refusal(request: TextEffectRequest, code: str) -> TextEffectReport:
    return TextEffectReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        edits=[edit.to_dict() for edit in request.edits],
    )


def _context_code(code: str) -> str:
    return "text_effect_" + code.removeprefix("format_")


def author_text_effects(request: TextEffectRequest) -> TextEffectReport:
    """Apply every edit's text effects atomically, or refuse with one code."""
    source, destination = request.source, request.destination
    if not source.is_file():
        return _refusal(request, "text_effect_source_not_found")
    if _sha256(source) != request.source_sha256:
        return _refusal(request, "text_effect_source_sha256_mismatch")
    if destination.exists() or destination.is_symlink():
        return _refusal(request, "text_effect_destination_invalid")
    try:
        graph = PackageGraph.from_file(source)
        document_root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    except (OSError, PackageError, zipfile.BadZipFile, ET.ParseError, KeyError):
        return _refusal(request, "text_effect_package_invalid")
    namespace = _word_namespace(document_root)
    if namespace is None or namespace not in SUPPORTED_WORD_NAMESPACES:
        return _refusal(request, "text_effect_namespace_unsupported")
    allowed_parts = _review_parts(graph, document_root, namespace)

    by_paragraph: dict[tuple[str, int], list[TextEffectEdit]] = {}
    for edit in request.edits:
        if edit.part not in allowed_parts:
            return _refusal(request, "text_effect_part_unsupported")
        by_paragraph.setdefault((edit.part, edit.paragraph), []).append(edit)
    for edits in by_paragraph.values():
        ordered = sorted(edits, key=lambda edit: (edit.start, edit.end))
        if any(a.end > b.start for a, b in itertools.pairwise(ordered)):
            return _refusal(request, "text_effect_edits_overlap")

    roots: dict[str, ET.Element] = {}
    for part_name in sorted({part for part, _ in by_paragraph}):
        try:
            root = parse_xml_with_preserved_qnames(graph.read_part_bytes(part_name))
        except (OSError, PackageError, ET.ParseError, KeyError):
            return _refusal(request, "text_effect_package_invalid")
        if _word_namespace(root) != namespace:
            return _refusal(request, "text_effect_namespace_unsupported")
        roots[part_name] = root

    targets: dict[tuple[str, int], ET.Element] = {}
    for (part_name, index), edits in by_paragraph.items():
        root = roots[part_name]
        paragraphs = list(root.iter(Q(namespace, "p")))
        if index > len(paragraphs):
            return _refusal(request, "text_effect_paragraph_missing")
        paragraph = paragraphs[index - 1]
        fingerprint = _mask_text(paragraph, namespace)
        if any(edit.structure_fingerprint_sha256 != fingerprint for edit in edits):
            return _refusal(request, "text_effect_structure_fingerprint_mismatch")
        parents = {id(child): parent for parent in root.iter() for child in parent}
        context = _detect_context_unsupported(paragraph) or _is_inside_sdt_or_unsupported(
            paragraph, parents
        )
        if context:
            return _refusal(request, _context_code(context))
        try:
            text = _paragraph_map(paragraph, namespace).text
        except _ReviewContextError:
            return _refusal(request, "text_effect_context_unsupported")
        for edit in edits:
            if edit.end > len(text):
                return _refusal(request, "text_effect_range_invalid")
            if text[edit.start : edit.end] != edit.expected_old_text:
                return _refusal(request, "text_effect_expected_text_mismatch")
        targets[(part_name, index)] = paragraph

    for key, paragraph in targets.items():
        edits = by_paragraph[key]
        mapping = _paragraph_map(paragraph, namespace)
        runs_by_id = {id(run.element): run for run in mapping.runs}
        boundaries = sorted({edit.start for edit in edits} | {edit.end for edit in edits})
        for child in list(paragraph):
            run = runs_by_id.get(id(child))
            if run is None or not run.text:
                continue
            cuts = [
                run.start,
                *(b for b in boundaries if run.start < b < run.end),
                run.end,
            ]
            pieces: list[ET.Element] = []
            for lower, upper in itertools.pairwise(cuts):
                piece = _clone_run(
                    child, namespace, run.text[lower - run.start : upper - run.start]
                )
                covering = next((e for e in edits if e.start <= lower and upper <= e.end), None)
                if covering is not None:
                    _apply_effects(piece, covering.effects, namespace)
                pieces.append(piece)
            position = list(paragraph).index(child)
            paragraph.remove(child)
            for offset, piece in enumerate(pieces):
                paragraph.insert(position + offset, piece)

    for root in roots.values():
        _declare_w14(root)
    replacements = {name: serialize_xml_with_preserved_qnames(root) for name, root in roots.items()}
    changed = [edit.to_dict() for edit in request.edits]

    def verifier(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, frozenset(replacements)):
            return "text_effect_staged_verification_failed"
        return None

    published: ReviewReport = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=changed,
        verify=verifier,
        failure_code="text_effect_staged_verification_failed",
        exists_code="text_effect_destination_invalid",
        unavailable_code="text_effect_publication_unavailable",
    )
    return TextEffectReport(
        status=published.status,
        code=published.code,
        source_sha256=published.source_sha256,
        destination_sha256=published.output_sha256,
        edits=changed,
    )
