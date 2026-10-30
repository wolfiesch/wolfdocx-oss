"""Source-bound native Word section header and footer region discovery and authoring.

The authoring surface is the main document body (``word/document.xml``) section properties
(``w:sectPr``), the package relationship graph, and the OPC content-type directory.
One region is an explicit header or footer part (``word/headerN.xml`` or ``word/footerN.xml``)
holding styled paragraphs with optional direct run formatting (bold, italic, underline,
color, font size, font name) or multiple runs,
bound to one source section and one variant
(``default``, ``first``, or ``even``).

Strict invariants enforced without mutation:
- Exact source digest matching and occupied destination protection.
- Section index must resolve to an existing source-bound section (1-based).
- Overwriting an explicitly occupied variant on the target section is refused.
- Inherited / link-to-previous semantics across multiple sections are respected.
- Paragraph styles must already exist in ``word/styles.xml``; missing styles are refused.
- Schema ordering within ``w:sectPr`` strictly adheres to OOXML (EG_HdrFtrReferences precede pgSz).
- Variant settings consistency: ``first`` variant ensures ``w:titlePg`` is enabled in ``w:sectPr``;
  ``even`` variant ensures ``w:evenAndOddHeaders`` is enabled in ``word/settings.xml``.
  If settings cannot be represented or are corrupt, the operation refuses rather than
  producing an apparently successful region that cannot display.
- Non-target parts are byte-preserved entry-for-entry.
"""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .assembly import _new_rel_id, _rels_name, _resolve
from .authoring import _MAX_TEXT_CHARACTERS
from .formatting import _HEX_COLOR_PATTERN, _VALID_UNDERLINE_STYLES
from .inspect import (
    SUPPORTED_WORD_NAMESPACES,
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .protection import _insert_settings_child
from .review import (
    _DOCUMENT,
    _publish,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_SETTINGS = "word/settings.xml"
_STYLES = "word/styles.xml"
# Word refuses run sizes above 1638 half-points (819 pt); matches the document builder bound.
_MAX_FONT_HALF_POINTS = 1_638
_MAX_FONT_NAME_CHARACTERS = 256
_INVALID_XML_TEXT = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")

_HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
_FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)

_PUBLISHED = "published"
_REFUSED = "refused"

_ALLOWED_KINDS = frozenset({"header", "footer"})
_ALLOWED_VARIANTS = frozenset({"default", "first", "even"})
_ALLOWED_ACTIONS = frozenset({"create", "replace", "remove", "link", "unlink", "relink"})
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


def _region_flag(value: Any, code: str) -> None:
    if value is not None and not isinstance(value, bool):
        raise TypeError(code)


def _region_underline(value: Any) -> None:
    if value is not None and not (
        isinstance(value, bool) or (isinstance(value, str) and value in _VALID_UNDERLINE_STYLES)
    ):
        raise ValueError("region_run_underline_invalid")


def _region_color(value: Any) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or _HEX_COLOR_PATTERN.fullmatch(value) is None:
        raise ValueError("region_run_color_invalid")
    return value.upper()


def _region_font_size(value: Any) -> None:
    if value is not None and (
        isinstance(value, bool)
        or not isinstance(value, int)
        or not 1 <= value <= _MAX_FONT_HALF_POINTS
    ):
        raise ValueError("region_run_font_size_invalid")


def _region_font_name(value: Any) -> None:
    if value is not None and (
        not isinstance(value, str)
        or not value.strip()
        or len(value) > _MAX_FONT_NAME_CHARACTERS
        or _INVALID_XML_TEXT.search(value) is not None
    ):
        raise ValueError("region_run_font_name_invalid")


@dataclass(frozen=True)
class RegionRun:
    """One directly formatted run inside a header or footer paragraph.

    ``None`` leaves the property to the paragraph default and style hierarchy; ``False``
    writes an explicit off toggle (``w:val="0"``, or ``w:u w:val="none"``) that overrides
    an inherited style. ``underline`` is ``True`` (single) or an ``ST_Underline`` value,
    ``color`` a six-digit RGB hex value (normalized to upper case), ``font_size`` is in
    half-points (``24`` is 12 pt), and ``font_name`` applies to the ASCII, high-ANSI and
    complex-script font slots.
    """

    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | str | None = None
    color: str | None = None
    font_size: int | None = None
    font_name: str | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.text, str):
            raise TypeError("region_run_text_invalid")
        _region_flag(self.bold, "region_run_bold_invalid")
        _region_flag(self.italic, "region_run_italic_invalid")
        _region_underline(self.underline)
        object.__setattr__(self, "color", _region_color(self.color))
        _region_font_size(self.font_size)
        _region_font_name(self.font_name)

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> RegionRun:
        keys = set(value)
        if not keys.issubset(_ALLOWED_RUN_FIELDS) or "text" not in keys:
            raise ValueError("region_run_fields_invalid")
        return cls(
            text=value["text"],
            bold=value.get("bold"),
            italic=value.get("italic"),
            underline=value.get("underline"),
            color=value.get("color"),
            font_size=value.get("font_size"),
            font_name=value.get("font_name"),
        )

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"text": self.text}
        for name in ("bold", "italic", "underline", "color", "font_size", "font_name"):
            flag = getattr(self, name)
            if flag is not None:
                result[name] = flag
        return result


@dataclass(frozen=True)
class RegionParagraph:
    """One styled paragraph to be placed in a header or footer part.

    Content is either ``text`` (one run) or ``runs`` (one or more runs), never both.
    Paragraph-level ``bold``/``italic`` apply to the single ``text`` run, and act as the
    default for each entry of ``runs`` that leaves the same property unset.
    Mapping entries in ``runs`` are validated and normalized to :class:`RegionRun`.
    """

    text: str = ""
    style: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    runs: Sequence[RegionRun | Mapping[str, Any]] | None = None

    def __post_init__(self) -> None:
        _region_flag(self.bold, "region_paragraph_bold_invalid")
        _region_flag(self.italic, "region_paragraph_italic_invalid")
        if self.runs is None:
            return
        if isinstance(self.runs, (str, bytes, Mapping)) or not isinstance(self.runs, Sequence):
            raise TypeError("region_paragraph_runs_invalid")
        normalized: list[RegionRun] = []
        for run in self.runs:
            if isinstance(run, RegionRun):
                normalized.append(run)
            elif isinstance(run, Mapping):
                normalized.append(RegionRun.from_mapping(run))
            else:
                raise TypeError("region_run_invalid")
        if not normalized:
            raise ValueError("region_paragraph_runs_invalid")
        if self.text:
            raise ValueError("region_paragraph_fields_invalid")
        object.__setattr__(self, "runs", tuple(normalized))

    def effective_runs(self) -> tuple[RegionRun, ...]:
        """Runs to serialize, with paragraph-level defaults resolved into each run."""
        if self.runs is None:
            return (RegionRun(self.text, self.bold, self.italic),)
        resolved: list[RegionRun] = []
        for run in self.runs:
            item = run if isinstance(run, RegionRun) else RegionRun.from_mapping(run)
            resolved.append(
                replace(
                    item,
                    bold=self.bold if item.bold is None else item.bold,
                    italic=self.italic if item.italic is None else item.italic,
                )
            )
        return tuple(resolved)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {}
        if self.runs is None:
            result["text"] = self.text
        else:
            result["runs"] = [
                (run if isinstance(run, RegionRun) else RegionRun.from_mapping(run)).to_dict()
                for run in self.runs
            ]
        if self.style is not None:
            result["style"] = self.style
        if self.bold is not None:
            result["bold"] = self.bold
        if self.italic is not None:
            result["italic"] = self.italic
        return result


@dataclass(frozen=True)
class RegionRequest:
    """Strict request for creating, replacing, or removing one section-bound header or footer region."""

    source: Path
    source_sha256: str
    destination: Path
    section: int
    kind: str = "header"
    variant: str = "default"
    paragraphs: tuple[RegionParagraph, ...] = ()
    action: str = "create"
    link_to_previous: bool | None = None

    def __init__(
        self,
        source: str | Path,
        source_sha256: str,
        destination: str | Path,
        section: int | None = None,
        kind: str = "header",
        variant: str = "default",
        paragraphs: Sequence[RegionParagraph | Mapping[str, Any]] = (),
        *,
        action: str = "create",
        section_index: int | None = None,
        link_to_previous: bool | None = None,
        is_linked_to_previous: bool | None = None,
    ) -> None:
        object.__setattr__(self, "source", Path(source))
        object.__setattr__(self, "source_sha256", str(source_sha256).lower())
        object.__setattr__(self, "destination", Path(destination))

        sec = section if section is not None else section_index
        object.__setattr__(self, "section", int(sec) if sec is not None else 1)
        object.__setattr__(self, "kind", kind)
        object.__setattr__(self, "variant", variant)

        link_to_previous = _resolve_link_to_previous(link_to_previous, is_linked_to_previous)
        if link_to_previous is not None:
            if link_to_previous is True and action == "create":
                action = "link"
            elif link_to_previous is False and action == "create" and not paragraphs:
                action = "unlink"

        if action in ("link", "relink") and link_to_previous is None:
            link_to_previous = True
        elif action == "unlink" and link_to_previous is None:
            link_to_previous = False

        object.__setattr__(self, "link_to_previous", link_to_previous)

        if not isinstance(action, str) or action not in _ALLOWED_ACTIONS:
            raise ValueError("region_action_invalid")
        object.__setattr__(self, "action", action)

        norm_paragraphs: list[RegionParagraph] = []
        for p in paragraphs:
            if isinstance(p, RegionParagraph):
                norm_paragraphs.append(p)
            elif isinstance(p, Mapping):
                norm_paragraphs.append(
                    RegionParagraph(
                        text=p.get("text", ""),
                        style=p.get("style"),
                        bold=p.get("bold"),
                        italic=p.get("italic"),
                        runs=p.get("runs"),
                    )
                )
        if action in ("remove", "link", "relink") and norm_paragraphs:
            raise ValueError("region_remove_paragraphs_prohibited")
        object.__setattr__(self, "paragraphs", tuple(norm_paragraphs))

    @property
    def section_index(self) -> int:
        return self.section

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "source": str(self.source),
            "source_sha256": self.source_sha256,
            "destination": str(self.destination),
            "section": self.section,
            "kind": self.kind,
            "variant": self.variant,
            "action": self.action,
            "paragraphs": [p.to_dict() for p in self.paragraphs],
        }
        if self.link_to_previous is not None:
            result["link_to_previous"] = self.link_to_previous
        return result


@dataclass(frozen=True)
class RegionReport:
    """Public content-safe result of one region authoring operation."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    region: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None
    retained_unreferenced_parts: tuple[str, ...] | None = None
    link_to_previous: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "schema_version": 1,
            "operation": "region",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
        }
        if self.region is not None:
            result["region"] = self.region
        if self.target_change_report is not None:
            result["target_change_report"] = self.target_change_report
        if self.non_target_parts_preserved is not None:
            result["non_target_parts_preserved"] = self.non_target_parts_preserved
        if self.retained_unreferenced_parts is not None:
            result["retained_unreferenced_parts"] = list(self.retained_unreferenced_parts)
        if self.link_to_previous is not None:
            result["link_to_previous"] = self.link_to_previous
        return result


def _word_ns(root: ET.Element) -> str:
    if root.tag.startswith("{"):
        ns = root.tag[1:].split("}", 1)[0]
        if ns in SUPPORTED_WORD_NAMESPACES:
            return ns
    return _W


def _sha256_file(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


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


def _find_sections(root: ET.Element, ns: str) -> list[ET.Element]:
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
    pg_sz = sect_pr.find(Q(ns, "pgSz"))
    pg_mar = sect_pr.find(Q(ns, "pgMar"))

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

    top = _int_attr(pg_mar, "top")
    bottom = _int_attr(pg_mar, "bottom")
    left = _int_attr(pg_mar, "left")
    right = _int_attr(pg_mar, "right")
    header = _int_attr(pg_mar, "header")
    footer = _int_attr(pg_mar, "footer")

    return {
        "width": width,
        "height": height,
        "orientation": orient,
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right,
        "header_margin": header,
        "footer_margin": footer,
    }


def _is_title_pg_enabled(sect_pr: ET.Element, ns: str) -> bool:
    elem = sect_pr.find(Q(ns, "titlePg"))
    if elem is None:
        return False
    val = elem.get(Q(ns, "val"))
    return val is None or val.lower() not in {"0", "false", "off"}


def _is_even_and_odd_enabled(settings_root: ET.Element | None, ns: str) -> bool:
    if settings_root is None:
        return False
    elem = settings_root.find(Q(ns, "evenAndOddHeaders"))
    if elem is None:
        return False
    val = elem.get(Q(ns, "val"))
    return val is None or val.lower() not in {"0", "false", "off"}


def _extract_explicit_references(
    sect_pr: ET.Element,
    ns: str,
    rel_id_to_part: Mapping[str, str],
) -> list[dict[str, Any]]:
    refs: list[dict[str, Any]] = []
    relationship_ns = _R if ns == _W else _STRICT_R
    for child in sect_pr:
        child_ns = child.tag[1:].split("}", 1)[0] if child.tag.startswith("{") else ""
        if child_ns in SUPPORTED_WORD_NAMESPACES:
            tag = child.tag.rsplit("}", 1)[-1]
            if tag in {"headerReference", "footerReference"}:
                kind = "header" if tag == "headerReference" else "footer"
                variant = child.get(Q(ns, "type")) or "default"
                rel_id = child.get(Q(relationship_ns, "id")) or child.get("id") or ""
                part = rel_id_to_part.get(rel_id, "")
                refs.append(
                    {
                        "kind": kind,
                        "variant": variant,
                        "relationship_id": rel_id,
                        "part": part,
                        "link_to_previous": False,
                    }
                )
    return refs


def _serialize_region_part(
    kind: str,
    paragraphs: Sequence[RegionParagraph],
    ns: str,
) -> bytes:
    tag = "hdr" if kind == "header" else "ftr"
    root = ET.Element(Q(ns, tag))
    if not paragraphs:
        ET.SubElement(root, Q(ns, "p"))
    else:
        for p in paragraphs:
            p_elem = ET.SubElement(root, Q(ns, "p"))
            if p.style:
                ppr = ET.SubElement(p_elem, Q(ns, "pPr"))
                ET.SubElement(ppr, Q(ns, "pStyle"), {Q(ns, "val"): p.style})
            for run in p.effective_runs():
                r_elem = ET.SubElement(p_elem, Q(ns, "r"))
                # ECMA-376 17.3.2.28 CT_RPr sequence:
                # rFonts, b, i, color, sz, szCs, u.
                rpr = ET.Element(Q(ns, "rPr"))
                if run.font_name is not None:
                    ET.SubElement(
                        rpr,
                        Q(ns, "rFonts"),
                        {
                            Q(ns, "ascii"): run.font_name,
                            Q(ns, "hAnsi"): run.font_name,
                            Q(ns, "cs"): run.font_name,
                        },
                    )
                for local, flag in (("b", run.bold), ("i", run.italic)):
                    if flag is True:
                        ET.SubElement(rpr, Q(ns, local))
                    elif flag is False:
                        ET.SubElement(rpr, Q(ns, local), {Q(ns, "val"): "0"})
                if run.color is not None:
                    ET.SubElement(rpr, Q(ns, "color"), {Q(ns, "val"): run.color})
                if run.font_size is not None:
                    size = str(run.font_size)
                    ET.SubElement(rpr, Q(ns, "sz"), {Q(ns, "val"): size})
                    ET.SubElement(rpr, Q(ns, "szCs"), {Q(ns, "val"): size})
                if run.underline is not None:
                    underline = run.underline
                    if isinstance(underline, bool):
                        underline = "single" if underline else "none"
                    ET.SubElement(rpr, Q(ns, "u"), {Q(ns, "val"): underline})
                if len(rpr):
                    r_elem.append(rpr)
                t_elem = ET.SubElement(r_elem, Q(ns, "t"))
                t_elem.text = run.text
                if run.text[:1].isspace() or run.text[-1:].isspace():
                    t_elem.set(_XML_SPACE, "preserve")
    return serialize_xml_with_preserved_qnames(root)


def _resolve_link_to_previous(link_to_previous: Any, is_linked_to_previous: Any) -> bool | None:
    """Merge the canonical flag with its python-docx ``is_linked_to_previous`` alias."""
    for flag in (link_to_previous, is_linked_to_previous):
        if flag is not None and not isinstance(flag, bool):
            raise TypeError("region_link_to_previous_invalid")
    if is_linked_to_previous is None:
        return link_to_previous
    if link_to_previous is not None and link_to_previous is not is_linked_to_previous:
        raise ValueError("region_link_to_previous_conflict")
    return is_linked_to_previous


_ALLOWED_REQUEST_FIELDS = frozenset(
    {
        "source",
        "source_sha256",
        "destination",
        "section",
        "kind",
        "variant",
        "paragraphs",
        "action",
        "link_to_previous",
        "is_linked_to_previous",
    }
)
_REQUIRED_REQUEST_FIELDS = frozenset({"source", "source_sha256", "destination", "section"})
_ALLOWED_PARAGRAPH_FIELDS = frozenset({"text", "style", "bold", "italic", "runs"})
_ALLOWED_RUN_FIELDS = frozenset(
    {"text", "bold", "italic", "underline", "color", "font_size", "font_name"}
)


def parse_region_request(value: Any) -> RegionRequest:
    """Parse and strictly validate a region authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("region_request_invalid")

    keys = set(value)
    if not keys.issubset(_ALLOWED_REQUEST_FIELDS) or not _REQUIRED_REQUEST_FIELDS.issubset(keys):
        raise ValueError("region_request_fields_invalid")

    source = value["source"]
    if not isinstance(source, str) or not source:
        raise ValueError("region_source_invalid")

    destination = value["destination"]
    if not isinstance(destination, str) or not destination:
        raise ValueError("region_destination_invalid")

    source_sha256 = value["source_sha256"]
    if not isinstance(source_sha256, str) or len(source_sha256) != 64:
        raise ValueError("region_source_sha256_invalid")
    try:
        int(source_sha256, 16)
    except ValueError as err:
        raise ValueError("region_source_sha256_invalid") from err
    source_sha256 = source_sha256.lower()

    sec = value["section"]
    if isinstance(sec, bool) or not isinstance(sec, int):
        raise TypeError("region_section_invalid")
    if sec < 1:
        raise ValueError("region_section_invalid")

    action = value.get("action", "create")
    if not isinstance(action, str) or action not in _ALLOWED_ACTIONS:
        raise ValueError("region_action_invalid")

    kind = value.get("kind", "header")
    if not isinstance(kind, str) or kind not in _ALLOWED_KINDS:
        raise ValueError("region_kind_invalid")

    variant = value.get("variant", "default")
    if not isinstance(variant, str) or variant not in _ALLOWED_VARIANTS:
        raise ValueError("region_variant_invalid")

    raw_paragraphs = value.get("paragraphs", [])
    if isinstance(raw_paragraphs, (str, bytes, Mapping)) or not isinstance(
        raw_paragraphs, Sequence
    ):
        raise TypeError("region_paragraphs_invalid")

    parsed_paragraphs: list[RegionParagraph] = []
    for p in raw_paragraphs:
        if isinstance(p, RegionParagraph):
            parsed_paragraphs.append(p)
        elif isinstance(p, Mapping):
            p_keys = set(p)
            if (
                not p_keys.issubset(_ALLOWED_PARAGRAPH_FIELDS)
                or len(p_keys & {"text", "runs"}) != 1
            ):
                raise ValueError("region_paragraph_fields_invalid")
            text = p.get("text", "")
            if not isinstance(text, str):
                raise TypeError("region_paragraph_text_invalid")
            paragraph = RegionParagraph(
                text=text,
                style=p.get("style"),
                bold=p.get("bold"),
                italic=p.get("italic"),
                runs=p.get("runs"),
            )
            if sum(len(run.text) for run in paragraph.effective_runs()) > _MAX_TEXT_CHARACTERS:
                raise ValueError("region_text_limit_exceeded")
            style = paragraph.style
            if style is not None and (not isinstance(style, str) or not style.strip()):
                raise ValueError("region_paragraph_style_invalid")
            parsed_paragraphs.append(paragraph)
        else:
            raise TypeError("region_paragraph_invalid")

    if action in ("remove", "link", "relink") and parsed_paragraphs:
        raise ValueError("region_remove_paragraphs_prohibited")

    raw_link_to_previous = _resolve_link_to_previous(
        value.get("link_to_previous"), value.get("is_linked_to_previous")
    )

    return RegionRequest(
        source=Path(source),
        source_sha256=source_sha256,
        destination=Path(destination),
        section=sec,
        kind=kind,
        variant=variant,
        paragraphs=tuple(parsed_paragraphs),
        action=action,
        link_to_previous=raw_link_to_previous,
    )


def load_region_request(path: str | Path) -> RegionRequest:
    """Load and validate a region authoring request from JSON on disk."""
    with Path(path).open(encoding="utf-8") as stream:
        return parse_region_request(json.load(stream))


def discover_region_targets(source: str | Path) -> dict[str, Any]:
    """Inventory section header and footer references and free variants without mutating."""
    path = Path(source)
    if not path.is_file():
        return {
            "passed": False,
            "code": "source_missing",
            "source_sha256": "",
            "sections": [],
        }

    try:
        source_sha = _sha256_file(path)
        graph = PackageGraph.from_file(path)
        if _DOCUMENT not in graph.parts:
            return {
                "passed": False,
                "code": "source_package_invalid",
                "source_sha256": source_sha,
                "sections": [],
            }
        doc_root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
        doc_tag = doc_root.tag
        doc_ns = doc_tag[1:].split("}", 1)[0] if doc_tag.startswith("{") else ""
        if doc_ns not in SUPPORTED_WORD_NAMESPACES or doc_tag != Q(doc_ns, "document"):
            return {
                "passed": False,
                "code": "source_package_invalid",
                "source_sha256": source_sha,
                "sections": [],
            }
    except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError, PackageError):
        return {
            "passed": False,
            "code": "source_package_invalid",
            "source_sha256": "",
            "sections": [],
        }

    ns = _word_ns(doc_root)
    sections = _find_sections(doc_root, ns)

    # Build relationship id -> part map
    rel_id_to_part: dict[str, str] = {}
    rels_name = _rels_name(_DOCUMENT)
    if rels_name in graph.parts:
        try:
            rels_root = ET.fromstring(graph.read_part_bytes(rels_name))
            for rel in rels_root.findall(Q(_PR, "Relationship")):
                rel_id = rel.get("Id")
                target = rel.get("Target")
                if rel_id and target:
                    rel_id_to_part[rel_id] = _resolve(_DOCUMENT, target)
        except (ET.ParseError, PackageError):
            pass

    # Read settings
    settings_root: ET.Element | None = None
    if _SETTINGS in graph.parts:
        try:
            settings_root = ET.fromstring(graph.read_part_bytes(_SETTINGS))
        except (ET.ParseError, PackageError):
            pass

    even_and_odd_enabled = _is_even_and_odd_enabled(settings_root, ns)

    all_combinations = [
        ("header", "default"),
        ("header", "first"),
        ("header", "even"),
        ("footer", "default"),
        ("footer", "first"),
        ("footer", "even"),
    ]

    active_references: dict[tuple[str, str], dict[str, Any]] = {}
    section_targets: list[dict[str, Any]] = []

    for idx, sect_pr in enumerate(sections, start=1):
        geom = _extract_page_geometry(sect_pr, ns)
        title_page = _is_title_pg_enabled(sect_pr, ns)
        explicit_refs = _extract_explicit_references(sect_pr, ns, rel_id_to_part)
        explicit_keys = {(r["kind"], r["variant"]) for r in explicit_refs}

        inherited_refs: list[dict[str, Any]] = []
        for key, ref_info in sorted(active_references.items()):
            if key not in explicit_keys:
                inherited_refs.append(
                    {
                        "kind": key[0],
                        "variant": key[1],
                        "inherited_from_section": ref_info["section"],
                        "relationship_id": ref_info["relationship_id"],
                        "part": ref_info["part"],
                        "link_to_previous": True,
                    }
                )

        for r in explicit_refs:
            active_references[(r["kind"], r["variant"])] = {
                "section": idx,
                "relationship_id": r["relationship_id"],
                "part": r["part"],
            }

        free_variants = [
            {"kind": k, "variant": v} for k, v in all_combinations if (k, v) not in explicit_keys
        ]
        occupied_variants = [
            {"kind": k, "variant": v} for k, v in all_combinations if (k, v) in explicit_keys
        ]

        section_targets.append(
            {
                "index": idx,
                "page_geometry": geom,
                "page_width": geom["width"],
                "page_height": geom["height"],
                "orientation": geom["orientation"],
                "title_page": title_page,
                "references": explicit_refs,
                "inherited_references": inherited_refs,
                "occupied_variants": occupied_variants,
                "free_variants": free_variants,
            }
        )

    return {
        "passed": True,
        "source_sha256": source_sha,
        "even_and_odd_headers": even_and_odd_enabled,
        "sections_count": len(sections),
        "sections": section_targets,
    }


def _refusal(request: RegionRequest, code: str) -> RegionReport:
    return RegionReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        region=None,
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def author_region(request: RegionRequest) -> RegionReport:
    """Create, replace, or remove one section-bound header or footer region atomically and verify output."""
    source = request.source.resolve()
    destination = request.destination.resolve()

    if source == destination:
        return _refusal(request, "source_destination_same")
    if not source.is_file():
        return _refusal(request, "source_missing")
    if not destination.parent.is_dir():
        return _refusal(request, "destination_parent_missing")
    if destination.exists() or destination.is_symlink():
        return _refusal(request, "region_destination_occupied")

    if _sha256_file(source) != request.source_sha256:
        return _refusal(request, "source_hash_mismatch")

    try:
        graph = PackageGraph.from_file(source)
    except (OSError, zipfile.BadZipFile, PackageError):
        return _refusal(request, "source_package_invalid")

    if _DOCUMENT not in graph.parts:
        return _refusal(request, "source_package_invalid")

    try:
        original_doc_bytes = graph.read_part_bytes(_DOCUMENT)
        doc_root = parse_xml_with_preserved_qnames(original_doc_bytes)
    except (ET.ParseError, PackageError):
        return _refusal(request, "source_package_invalid")

    doc_tag = doc_root.tag
    doc_ns = doc_tag[1:].split("}", 1)[0] if doc_tag.startswith("{") else ""
    if doc_ns not in SUPPORTED_WORD_NAMESPACES or doc_tag != Q(doc_ns, "document"):
        return _refusal(request, "region_document_root_invalid")

    ns = doc_ns
    relationship_ns = _R if ns == _W else _STRICT_R
    sections = _find_sections(doc_root, ns)

    if request.section_index < 1 or request.section_index > len(sections):
        return _refusal(request, "region_section_not_found")

    target_sect_pr = sections[request.section_index - 1]

    # Defense: remove must reject non-empty paragraphs
    if request.action == "remove" and request.paragraphs:
        return _refusal(request, "region_remove_paragraphs_prohibited")

    # Validate for duplicate and foreign namespace references across all sections
    for sec in sections:
        seen_explicit_variants: set[tuple[str, str]] = set()
        for child in sec:
            child_tag = child.tag.rsplit("}", 1)[-1]
            if child_tag in ("headerReference", "footerReference"):
                child_ns = child.tag[1:].split("}", 1)[0] if child.tag.startswith("{") else ""
                if child_ns not in SUPPORTED_WORD_NAMESPACES:
                    return _refusal(request, "region_reference_foreign_namespace")
                if child.tag not in (
                    Q(child_ns, "headerReference"),
                    Q(child_ns, "footerReference"),
                ):
                    return _refusal(request, "region_reference_foreign_namespace")
                kind_val = "header" if child_tag == "headerReference" else "footer"
                variant_val = child.get(Q(child_ns, "type"))
                if not variant_val or variant_val not in _ALLOWED_VARIANTS:
                    return _refusal(request, "region_reference_malformed")
                if not child.get(Q(relationship_ns, "id")):
                    return _refusal(request, "region_relationship_invalid")
                key = (kind_val, variant_val)
                if key in seen_explicit_variants:
                    return _refusal(request, "region_reference_duplicate")
                seen_explicit_variants.add(key)

    # Check target explicit reference in target_sect_pr with exact QName
    ref_tag = "headerReference" if request.kind == "header" else "footerReference"
    expected_ref_qname = Q(ns, ref_tag)
    target_ref: ET.Element | None = None
    for child in target_sect_pr:
        if child.tag == expected_ref_qname:
            variant = child.get(Q(ns, "type"))
            if variant == request.variant:
                target_ref = child
                break

    if request.action == "create" and target_ref is not None:
        return _refusal(request, "region_variant_occupied")
    if request.action in ("replace", "remove") and target_ref is None:
        return _refusal(request, "region_target_missing")
    replacements: dict[str, bytes] = {}
    additions: dict[str, bytes] = {}

    # Load content types
    if "[Content_Types].xml" not in graph.parts:
        return _refusal(request, "source_package_invalid")
    original_types_bytes = graph.read_part_bytes("[Content_Types].xml")
    try:
        types_root = parse_xml_with_preserved_qnames(original_types_bytes)
    except ET.ParseError:
        return _refusal(request, "source_package_invalid")

    if types_root.tag != Q(_CT, "Types"):
        return _refusal(request, "source_package_invalid")

    # Load document relationships
    rels_name = _rels_name(_DOCUMENT)
    if rels_name in graph.parts:
        original_rels_bytes = graph.read_part_bytes(rels_name)
        try:
            rels_root = parse_xml_with_preserved_qnames(original_rels_bytes)
        except ET.ParseError:
            return _refusal(request, "source_package_invalid")
        if rels_root.tag != Q(_PR, "Relationships"):
            return _refusal(request, "source_package_invalid")
    else:
        if request.action in ("replace", "remove"):
            return _refusal(request, "region_relationship_missing")
        original_rels_bytes = b""
        rels_root = ET.Element(Q(_PR, "Relationships"))

    if request.action in ("link", "relink"):
        if request.section_index < 2:
            return _refusal(request, "region_predecessor_section_missing")
        if request.paragraphs:
            return _refusal(request, "region_relink_paragraphs_prohibited")
        if target_ref is None:
            return _refusal(request, "region_already_linked")

        assert target_ref is not None
        target_sect_pr.remove(target_ref)
        replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(doc_root)

        all_changed = frozenset(replacements.keys())

        def verify_link_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, all_changed):
                return "region_unrelated_parts_changed"
            try:
                insp = inspect_docx(staged)
                if not insp.valid:
                    return "region_staged_verification_failed"
                staged_graph = PackageGraph.from_file(staged)
                staged_doc_root = ET.fromstring(staged_graph.read_part_bytes(_DOCUMENT))
                staged_ns = _word_ns(staged_doc_root)
                staged_sections = _find_sections(staged_doc_root, staged_ns)
                staged_target_sec = staged_sections[request.section_index - 1]
                staged_ref_qname = Q(staged_ns, ref_tag)
                for child in staged_target_sec:
                    if child.tag == staged_ref_qname:
                        v = child.get(Q(staged_ns, "type")) or "default"
                        if v == request.variant:
                            return "region_staged_verification_failed"
            except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError, PackageError):
                return "region_staged_verification_failed"
            return None

        region_info = {
            "section": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "link_to_previous": True,
        }
        target_change_report = {
            "action": "relink",
            "section_index": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "link_to_previous": True,
        }

        published = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[region_info],
            verify=verify_link_publication,
            failure_code="region_staged_verification_failed",
            exists_code="region_destination_occupied",
            unavailable_code="region_publication_unavailable",
        )

        if not published.passed:
            return RegionReport(
                status=_REFUSED,
                code=published.code,
                source_sha256=request.source_sha256,
                destination_sha256=None,
                region=None,
                target_change_report=None,
                non_target_parts_preserved=None,
                retained_unreferenced_parts=None,
                link_to_previous=True,
            )

        return RegionReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            region=region_info,
            target_change_report=target_change_report,
            non_target_parts_preserved=True,
            retained_unreferenced_parts=None,
            link_to_previous=True,
        )

    if request.action == "unlink":
        if request.section_index < 2:
            return _refusal(request, "region_predecessor_section_missing")
        if target_ref is not None:
            return _refusal(request, "region_already_unlinked")

        # Resolve predecessor active story
        inherited_elem: ET.Element | None = None
        inherited_sec_idx: int | None = None
        for s_i in range(request.section_index - 2, -1, -1):
            pred_sect = sections[s_i]
            for child in pred_sect:
                if child.tag == expected_ref_qname:
                    v = child.get(Q(ns, "type")) or "default"
                    if v == request.variant:
                        inherited_elem = child
                        inherited_sec_idx = s_i + 1
                        break
            if inherited_elem is not None:
                break

        if inherited_elem is None:
            return _refusal(request, "region_inherited_story_missing")

        pred_rel_id = inherited_elem.get(Q(relationship_ns, "id")) or inherited_elem.get("id") or ""
        if not pred_rel_id:
            return _refusal(request, "region_relationship_invalid")

        pred_rel_elem: ET.Element | None = None
        for r in rels_root.findall(Q(_PR, "Relationship")):
            if r.get("Id") == pred_rel_id:
                pred_rel_elem = r
                break
        if pred_rel_elem is None:
            return _refusal(request, "region_relationship_missing")

        pred_raw_target = pred_rel_elem.get("Target", "").strip()
        pred_target_part = _resolve(_DOCUMENT, pred_raw_target)
        if pred_target_part not in graph.parts:
            return _refusal(request, "region_story_part_missing")

        active_story_bytes = graph.read_part_bytes(pred_target_part)

        if request.paragraphs:
            styles_to_check = {p.style for p in request.paragraphs if p.style}
            if styles_to_check:
                if _STYLES not in graph.parts:
                    return _refusal(request, "region_style_missing")
                try:
                    styles_root = ET.fromstring(graph.read_part_bytes(_STYLES))
                    known_styles: set[str] = set()
                    for style_elem in styles_root.findall(Q(ns, "style")):
                        style_id = style_elem.get(Q(ns, "styleId")) or style_elem.get("styleId")
                        if style_id:
                            known_styles.add(style_id)
                        name_elem = style_elem.find(Q(ns, "name"))
                        if name_elem is not None:
                            name_val = name_elem.get(Q(ns, "val")) or name_elem.get("val")
                            if name_val:
                                known_styles.add(name_val)
                    for style in styles_to_check:
                        if style not in known_styles:
                            return _refusal(request, "region_style_missing")
                except (ET.ParseError, PackageError):
                    return _refusal(request, "region_style_missing")
            new_story_bytes = _serialize_region_part(request.kind, request.paragraphs, ns)
        else:
            new_story_bytes = active_story_bytes

        # Handle first page titlePg setting
        if request.variant == "first":
            title_pg_elem = target_sect_pr.find(Q(ns, "titlePg"))
            if title_pg_elem is None:
                new_title_pg = ET.Element(Q(ns, "titlePg"))
                _insert_sect_pr_child(target_sect_pr, new_title_pg)
            else:
                val = title_pg_elem.get(Q(ns, "val"))
                if val is not None and val.lower() in {"0", "false", "off"}:
                    title_pg_elem.set(Q(ns, "val"), "1")

        # Handle even page settings
        if request.variant == "even":
            if _SETTINGS in graph.parts:
                try:
                    original_settings_bytes = graph.read_part_bytes(_SETTINGS)
                    settings_root = parse_xml_with_preserved_qnames(original_settings_bytes)
                except (ET.ParseError, PackageError):
                    return _refusal(request, "region_settings_invalid")

                settings_ns = _word_ns(settings_root)
                even_elem = settings_root.find(Q(settings_ns, "evenAndOddHeaders"))
                if even_elem is None:
                    _insert_settings_child(
                        settings_root,
                        ET.Element(Q(settings_ns, "evenAndOddHeaders")),
                        settings_ns,
                    )
                else:
                    val = even_elem.get(Q(settings_ns, "val"))
                    if val is not None and val.lower() in {"0", "false", "off"}:
                        even_elem.set(Q(settings_ns, "val"), "1")
                replacements[_SETTINGS] = serialize_xml_with_preserved_qnames(settings_root)
            else:
                settings_root = ET.Element(Q(ns, "settings"))
                _insert_settings_child(settings_root, ET.Element(Q(ns, "evenAndOddHeaders")), ns)
                additions[_SETTINGS] = serialize_xml_with_preserved_qnames(settings_root)
                settings_rel_id = _new_rel_id(rels_root)
                rels_root.append(
                    ET.Element(
                        Q(_PR, "Relationship"),
                        {
                            "Id": settings_rel_id,
                            "Type": f"{relationship_ns}/settings",
                            "Target": "settings.xml",
                        },
                    )
                )
                if not any(
                    node.get("PartName") == f"/{_SETTINGS}"
                    for node in types_root.findall(Q(_CT, "Override"))
                ):
                    types_root.append(
                        ET.Element(
                            Q(_CT, "Override"),
                            {
                                "PartName": f"/{_SETTINGS}",
                                "ContentType": _SETTINGS_CONTENT_TYPE,
                            },
                        )
                    )

        # Allocate new part name
        prefix = f"word/{request.kind}"
        serial = 1
        while f"{prefix}{serial}.xml" in graph.parts or f"{prefix}{serial}.xml" in additions:
            serial += 1
        part_name = f"{prefix}{serial}.xml"
        rel_target = f"{request.kind}{serial}.xml"
        rel_type = f"{relationship_ns}/{request.kind}"
        content_type = _HEADER_CONTENT_TYPE if request.kind == "header" else _FOOTER_CONTENT_TYPE

        additions[part_name] = new_story_bytes

        rel_id = _new_rel_id(rels_root)
        rels_root.append(
            ET.Element(
                Q(_PR, "Relationship"),
                {
                    "Id": rel_id,
                    "Type": rel_type,
                    "Target": rel_target,
                },
            )
        )
        replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)

        types_root.append(
            ET.Element(
                Q(_CT, "Override"),
                {
                    "PartName": f"/{part_name}",
                    "ContentType": content_type,
                },
            )
        )
        replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)

        new_ref = ET.Element(
            expected_ref_qname,
            {
                Q(ns, "type"): request.variant,
                Q(relationship_ns, "id"): rel_id,
            },
        )
        _insert_sect_pr_child(target_sect_pr, new_ref)
        replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(doc_root)

        all_changed = frozenset(replacements.keys() | additions.keys())

        def verify_unlink_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, all_changed):
                return "region_unrelated_parts_changed"
            try:
                insp = inspect_docx(staged)
                if not insp.valid:
                    return "region_staged_verification_failed"
                staged_graph = PackageGraph.from_file(staged)
                if part_name not in staged_graph.parts:
                    return "region_staged_verification_failed"
                staged_doc_root = ET.fromstring(staged_graph.read_part_bytes(_DOCUMENT))
                staged_ns = _word_ns(staged_doc_root)
                staged_sections = _find_sections(staged_doc_root, staged_ns)
                staged_target_sec = staged_sections[request.section_index - 1]
                staged_ref_qname = Q(staged_ns, ref_tag)
                found = False
                for child in staged_target_sec:
                    if child.tag == staged_ref_qname:
                        v = child.get(Q(staged_ns, "type")) or "default"
                        r = child.get(Q(relationship_ns, "id")) or child.get("id")
                        if v == request.variant and r == rel_id:
                            found = True
                            break
                if not found:
                    return "region_staged_verification_failed"
            except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError, PackageError):
                return "region_staged_verification_failed"
            return None

        region_info = {
            "section": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "part": part_name,
            "relationship_id": rel_id,
            "link_to_previous": False,
            "inherited_from_section": inherited_sec_idx,
            "cloned_from_part": pred_target_part,
        }
        target_change_report = {
            "action": "unlink",
            "section_index": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "part": part_name,
            "relationship_id": rel_id,
            "link_to_previous": False,
            "inherited_from_section": inherited_sec_idx,
            "cloned_from_part": pred_target_part,
        }

        published = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[region_info],
            additions=additions,
            verify=verify_unlink_publication,
            failure_code="region_staged_verification_failed",
            exists_code="region_destination_occupied",
            unavailable_code="region_publication_unavailable",
        )

        if not published.passed:
            return RegionReport(
                status=_REFUSED,
                code=published.code,
                source_sha256=request.source_sha256,
                destination_sha256=None,
                region=None,
                target_change_report=None,
                non_target_parts_preserved=None,
                retained_unreferenced_parts=None,
                link_to_previous=False,
            )

        return RegionReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            region=region_info,
            target_change_report=target_change_report,
            non_target_parts_preserved=True,
            retained_unreferenced_parts=None,
            link_to_previous=False,
        )

    # For replace and remove: validate relationship, target part, content type, and story root
    target_part: str = ""
    rel_id: str = ""
    if request.action in ("replace", "remove"):
        assert target_ref is not None
        rel_id = target_ref.get(Q(relationship_ns, "id")) or ""
        if not rel_id:
            return _refusal(request, "region_relationship_invalid")

        rel_elem: ET.Element | None = None
        for r in rels_root.findall(Q(_PR, "Relationship")):
            if r.get("Id") == rel_id:
                rel_elem = r
                break
        if rel_elem is None:
            return _refusal(request, "region_relationship_missing")

        if rel_elem.get("TargetMode") == "External":
            return _refusal(request, "region_relationship_external_unsupported")

        raw_target = rel_elem.get("Target", "").strip()
        if (
            not raw_target
            or raw_target.startswith("/")
            or ".." in raw_target.split("/")
            or "\\" in raw_target
        ):
            return _refusal(request, "region_relationship_unsafe")

        rel_type = rel_elem.get("Type", "")
        expected_rel_type = f"{relationship_ns}/{request.kind}"
        if rel_type != expected_rel_type:
            return _refusal(request, "region_relationship_type_invalid")

        target_part = _resolve(_DOCUMENT, raw_target)
        if target_part not in graph.parts:
            return _refusal(request, "region_story_part_missing")

        expected_content_type = (
            _HEADER_CONTENT_TYPE if request.kind == "header" else _FOOTER_CONTENT_TYPE
        )
        target_overrides = [
            override
            for override in types_root.findall(Q(_CT, "Override"))
            if override.get("PartName") == f"/{target_part}"
        ]
        if not target_overrides:
            return _refusal(request, "region_story_type_invalid")
        if len(target_overrides) > 1:
            return _refusal(request, "region_content_type_duplicate")
        ct_elem = target_overrides[0]
        if ct_elem.get("ContentType") != expected_content_type:
            return _refusal(request, "region_story_type_invalid")
        try:
            story_bytes = graph.read_part_bytes(target_part)
            story_root = ET.fromstring(story_bytes)
        except (ET.ParseError, PackageError):
            return _refusal(request, "region_story_root_invalid")

        story_ns = _word_ns(story_root)
        expected_story_tag = "hdr" if request.kind == "header" else "ftr"
        if (
            story_root.tag != Q(story_ns, expected_story_tag)
            or story_ns not in SUPPORTED_WORD_NAMESPACES
        ):
            return _refusal(request, "region_story_root_invalid")

    # Validate paragraph styles against word/styles.xml (for create and replace)
    if request.action in ("create", "replace"):
        styles_to_check = {p.style for p in request.paragraphs if p.style}
        if styles_to_check:
            if _STYLES not in graph.parts:
                return _refusal(request, "region_style_missing")
            try:
                styles_root = ET.fromstring(graph.read_part_bytes(_STYLES))
                known_styles: set[str] = set()
                for style_elem in styles_root.findall(Q(ns, "style")):
                    style_id = style_elem.get(Q(ns, "styleId")) or style_elem.get("styleId")
                    if style_id:
                        known_styles.add(style_id)
                    name_elem = style_elem.find(Q(ns, "name"))
                    if name_elem is not None:
                        name_val = name_elem.get(Q(ns, "val")) or name_elem.get("val")
                        if name_val:
                            known_styles.add(name_val)
                for style in styles_to_check:
                    if style not in known_styles:
                        return _refusal(request, "region_style_missing")
            except (ET.ParseError, PackageError):
                return _refusal(request, "region_style_missing")

    if request.action == "remove":
        assert target_ref is not None
        target_sect_pr.remove(target_ref)
        replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(doc_root)

        # Build rel_id -> part mapping
        rel_id_to_part: dict[str, str] = {}
        for rel in rels_root.findall(Q(_PR, "Relationship")):
            r_id = rel.get("Id")
            r_tgt = rel.get("Target")
            if r_id and r_tgt:
                rel_id_to_part[r_id] = _resolve(_DOCUMENT, r_tgt)

        # Check if target_part is still referenced by any section in the document
        remaining_refs = 0
        for sec in sections:
            for child in sec:
                child_tag = child.tag.rsplit("}", 1)[-1]
                if child_tag in ("headerReference", "footerReference"):
                    c_rid = child.get(Q(relationship_ns, "id"))
                    if c_rid and rel_id_to_part.get(c_rid) == target_part:
                        remaining_refs += 1

        retained_unreferenced = target_part if remaining_refs == 0 else None
        retained_tuple: tuple[str, ...] = (retained_unreferenced,) if retained_unreferenced else ()

        all_changed = frozenset(replacements.keys() | additions.keys())

        def verify_remove_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, all_changed):
                return "region_unrelated_parts_changed"
            try:
                staged_graph = PackageGraph.from_file(staged)
            except (OSError, PackageError):
                return "region_staged_verification_failed"

            disc = discover_region_targets(staged)
            if not disc.get("passed"):
                return "region_staged_verification_failed"

            sections_list = disc.get("sections", [])
            if len(sections_list) < request.section_index:
                return "region_staged_verification_failed"

            staged_sec = sections_list[request.section_index - 1]
            has_ref = any(
                r.get("kind") == request.kind and r.get("variant") == request.variant
                for r in staged_sec.get("references", [])
            )
            if has_ref:
                return "region_staged_verification_failed"

            if retained_unreferenced and retained_unreferenced not in staged_graph.parts:
                return "region_staged_verification_failed"

            return None

        region_info: dict[str, Any] = {
            "action": "remove",
            "section": request.section,
            "section_index": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "removed_reference": {
                "kind": request.kind,
                "variant": request.variant,
                "relationship_id": rel_id,
                "part": target_part,
            },
            "retained_unreferenced_parts": list(retained_tuple),
        }

        target_change_report: dict[str, Any] = {
            "passed": True,
            "action": "remove",
            "section": request.section,
            "section_index": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "removed_reference": {
                "kind": request.kind,
                "variant": request.variant,
                "relationship_id": rel_id,
                "part": target_part,
            },
            "retained_unreferenced_parts": list(retained_tuple),
        }

        published = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[region_info],
            additions=additions,
            verify=verify_remove_publication,
            failure_code="region_staged_verification_failed",
            exists_code="region_destination_occupied",
            unavailable_code="region_publication_unavailable",
        )

        if not published.passed:
            return RegionReport(
                status=_REFUSED,
                code=published.code,
                source_sha256=request.source_sha256,
                destination_sha256=None,
                region=None,
                target_change_report=None,
                non_target_parts_preserved=None,
                retained_unreferenced_parts=None,
            )

        return RegionReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            region=region_info,
            target_change_report=target_change_report,
            non_target_parts_preserved=True,
            retained_unreferenced_parts=retained_tuple,
        )

    if request.action == "replace":
        assert target_ref is not None

        # Build rel_id -> part mapping
        rel_id_to_part = {}
        for rel in rels_root.findall(Q(_PR, "Relationship")):
            r_id = rel.get("Id")
            r_tgt = rel.get("Target")
            if r_id and r_tgt:
                rel_id_to_part[r_id] = _resolve(_DOCUMENT, r_tgt)

        # Count references to target_part across all sections
        part_ref_count = 0
        for sec in sections:
            for child in sec:
                child_tag = child.tag.rsplit("}", 1)[-1]
                if child_tag in ("headerReference", "footerReference"):
                    c_rid = child.get(Q(relationship_ns, "id"))
                    if c_rid and rel_id_to_part.get(c_rid) == target_part:
                        part_ref_count += 1

        is_shared = part_ref_count > 1
        header_bytes = _serialize_region_part(request.kind, request.paragraphs, ns)

        if is_shared:
            prefix = f"word/{request.kind}"
            serial = 1
            while f"{prefix}{serial}.xml" in graph.parts or f"{prefix}{serial}.xml" in additions:
                serial += 1
            final_part = f"{prefix}{serial}.xml"
            rel_target = f"{request.kind}{serial}.xml"
            rel_type = f"{relationship_ns}/{request.kind}"
            content_type = (
                _HEADER_CONTENT_TYPE if request.kind == "header" else _FOOTER_CONTENT_TYPE
            )

            final_rel_id = _new_rel_id(rels_root)
            rels_root.append(
                ET.Element(
                    Q(_PR, "Relationship"),
                    {
                        "Id": final_rel_id,
                        "Type": rel_type,
                        "Target": rel_target,
                    },
                )
            )

            if not any(
                node.get("PartName") == f"/{final_part}"
                for node in types_root.findall(Q(_CT, "Override"))
            ):
                types_root.append(
                    ET.Element(
                        Q(_CT, "Override"),
                        {
                            "PartName": f"/{final_part}",
                            "ContentType": content_type,
                        },
                    )
                )

            additions[final_part] = header_bytes
            target_ref.set(Q(relationship_ns, "id"), final_rel_id)

            replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(doc_root)
            replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)
            replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)
        else:
            final_part = target_part
            final_rel_id = rel_id
            replacements[final_part] = header_bytes
            replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(doc_root)

        all_changed = frozenset(replacements.keys() | additions.keys())

        def verify_replace_publication(staged: Path) -> str | None:
            if not _unrelated_parts_preserved(source, staged, all_changed):
                return "region_unrelated_parts_changed"
            try:
                staged_graph = PackageGraph.from_file(staged)
            except (OSError, PackageError):
                return "region_staged_verification_failed"

            if final_part not in staged_graph.parts:
                return "region_staged_verification_failed"
            if staged_graph.read_part_bytes(final_part) != header_bytes:
                return "region_staged_verification_failed"

            disc = discover_region_targets(staged)
            if not disc.get("passed"):
                return "region_staged_verification_failed"

            sections_list = disc.get("sections", [])
            if len(sections_list) < request.section_index:
                return "region_staged_verification_failed"

            staged_sec = sections_list[request.section_index - 1]
            has_ref = any(
                r.get("kind") == request.kind
                and r.get("variant") == request.variant
                and r.get("relationship_id") == final_rel_id
                and r.get("part") == final_part
                for r in staged_sec.get("references", [])
            )
            if not has_ref:
                return "region_staged_verification_failed"

            if is_shared:
                if target_part not in staged_graph.parts:
                    return "region_staged_verification_failed"
                if staged_graph.read_part_bytes(target_part) != graph.read_part_bytes(target_part):
                    return "region_staged_verification_failed"

            return None

        region_info = {
            "action": "replace",
            "section": request.section,
            "section_index": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "part": final_part,
            "relationship_id": final_rel_id,
            "paragraphs_count": len(request.paragraphs),
            "isolated": is_shared,
            "previous_part": target_part,
        }

        target_change_report = {
            "passed": True,
            "action": "replace",
            "section": request.section,
            "section_index": request.section_index,
            "kind": request.kind,
            "variant": request.variant,
            "part": final_part,
            "relationship_id": final_rel_id,
            "isolated": is_shared,
        }

        published = _publish(
            source=source,
            output=destination,
            source_sha256=request.source_sha256,
            replacements=replacements,
            revisions=[],
            targets=[region_info],
            additions=additions,
            verify=verify_replace_publication,
            failure_code="region_staged_verification_failed",
            exists_code="region_destination_occupied",
            unavailable_code="region_publication_unavailable",
        )

        if not published.passed:
            return RegionReport(
                status=_REFUSED,
                code=published.code,
                source_sha256=request.source_sha256,
                destination_sha256=None,
                region=None,
                target_change_report=None,
                non_target_parts_preserved=None,
                retained_unreferenced_parts=None,
            )

        return RegionReport(
            status=_PUBLISHED,
            code=None,
            source_sha256=request.source_sha256,
            destination_sha256=published.output_sha256,
            region=region_info,
            target_change_report=target_change_report,
            non_target_parts_preserved=True,
            retained_unreferenced_parts=None,
        )

    # request.action == "create"
    # Handle first page titlePg setting
    if request.variant == "first":
        title_pg_elem = target_sect_pr.find(Q(ns, "titlePg"))
        if title_pg_elem is None:
            new_title_pg = ET.Element(Q(ns, "titlePg"))
            _insert_sect_pr_child(target_sect_pr, new_title_pg)
        else:
            val = title_pg_elem.get(Q(ns, "val"))
            if val is not None and val.lower() in {"0", "false", "off"}:
                title_pg_elem.set(Q(ns, "val"), "1")

    # Handle even page settings
    if request.variant == "even":
        if _SETTINGS in graph.parts:
            try:
                original_settings_bytes = graph.read_part_bytes(_SETTINGS)
                settings_root = parse_xml_with_preserved_qnames(original_settings_bytes)
            except (ET.ParseError, PackageError):
                return _refusal(request, "region_settings_invalid")

            settings_ns = _word_ns(settings_root)
            even_elem = settings_root.find(Q(settings_ns, "evenAndOddHeaders"))
            if even_elem is None:
                _insert_settings_child(
                    settings_root, ET.Element(Q(settings_ns, "evenAndOddHeaders")), settings_ns
                )
            else:
                val = even_elem.get(Q(settings_ns, "val"))
                if val is not None and val.lower() in {"0", "false", "off"}:
                    even_elem.set(Q(settings_ns, "val"), "1")
            replacements[_SETTINGS] = serialize_xml_with_preserved_qnames(settings_root)
        else:
            # Create new word/settings.xml
            settings_root = ET.Element(Q(ns, "settings"))
            _insert_settings_child(settings_root, ET.Element(Q(ns, "evenAndOddHeaders")), ns)
            additions[_SETTINGS] = serialize_xml_with_preserved_qnames(settings_root)

            # Add relationship for settings
            settings_rel_id = _new_rel_id(rels_root)
            rels_root.append(
                ET.Element(
                    Q(_PR, "Relationship"),
                    {
                        "Id": settings_rel_id,
                        "Type": f"{relationship_ns}/settings",
                        "Target": "settings.xml",
                    },
                )
            )

            # Add content type override for settings
            if not any(
                node.get("PartName") == f"/{_SETTINGS}"
                for node in types_root.findall(Q(_CT, "Override"))
            ):
                types_root.append(
                    ET.Element(
                        Q(_CT, "Override"),
                        {
                            "PartName": f"/{_SETTINGS}",
                            "ContentType": _SETTINGS_CONTENT_TYPE,
                        },
                    )
                )

    # Allocate new part name: word/headerN.xml or word/footerN.xml
    prefix = f"word/{request.kind}"
    serial = 1
    while f"{prefix}{serial}.xml" in graph.parts or f"{prefix}{serial}.xml" in additions:
        serial += 1
    part_name = f"{prefix}{serial}.xml"
    rel_target = f"{request.kind}{serial}.xml"
    rel_type = f"{relationship_ns}/{request.kind}"
    content_type = _HEADER_CONTENT_TYPE if request.kind == "header" else _FOOTER_CONTENT_TYPE

    # Allocate relationship in word/_rels/document.xml.rels
    rel_id = _new_rel_id(rels_root)
    rels_root.append(
        ET.Element(
            Q(_PR, "Relationship"),
            {
                "Id": rel_id,
                "Type": rel_type,
                "Target": rel_target,
            },
        )
    )

    # Add content type override in [Content_Types].xml
    if not any(
        node.get("PartName") == f"/{part_name}" for node in types_root.findall(Q(_CT, "Override"))
    ):
        types_root.append(
            ET.Element(
                Q(_CT, "Override"),
                {
                    "PartName": f"/{part_name}",
                    "ContentType": content_type,
                },
            )
        )

    # Build and add header/footer part
    header_bytes = _serialize_region_part(request.kind, request.paragraphs, ns)
    additions[part_name] = header_bytes

    # Insert reference into sectPr in schema order (EG_HdrFtrReferences)
    new_ref = ET.Element(
        Q(ns, ref_tag),
        {
            Q(ns, "type"): request.variant,
            Q(relationship_ns, "id"): rel_id,
        },
    )
    _insert_sect_pr_child(target_sect_pr, new_ref)

    # Serialize modified existing parts
    replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(doc_root)
    replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)
    replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)

    all_changed = frozenset(replacements.keys() | additions.keys())

    def verify_create_publication(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, all_changed):
            return "region_unrelated_parts_changed"
        try:
            staged_graph = PackageGraph.from_file(staged)
        except (OSError, PackageError):
            return "region_staged_verification_failed"

        if part_name not in staged_graph.parts:
            return "region_staged_verification_failed"
        if staged_graph.read_part_bytes(part_name) != header_bytes:
            return "region_staged_verification_failed"

        disc = discover_region_targets(staged)
        if not disc.get("passed"):
            return "region_staged_verification_failed"

        sections_list = disc.get("sections", [])
        if len(sections_list) < request.section_index:
            return "region_staged_verification_failed"

        staged_sec = sections_list[request.section_index - 1]
        has_ref = any(
            r.get("kind") == request.kind
            and r.get("variant") == request.variant
            and r.get("relationship_id") == rel_id
            and r.get("part") == part_name
            for r in staged_sec.get("references", [])
        )
        if not has_ref:
            return "region_staged_verification_failed"

        if request.variant == "first" and not staged_sec.get("title_page"):
            return "region_staged_verification_failed"
        if request.variant == "even" and not disc.get("even_and_odd_headers"):
            return "region_staged_verification_failed"

        return None

    region_info: dict[str, Any] = {
        "action": "create",
        "section": request.section,
        "section_index": request.section_index,
        "kind": request.kind,
        "variant": request.variant,
        "part": part_name,
        "relationship_id": rel_id,
        "paragraphs_count": len(request.paragraphs),
    }

    target_change_report: dict[str, Any] = {
        "passed": True,
        "action": "create",
        "section": request.section,
        "section_index": request.section_index,
        "kind": request.kind,
        "variant": request.variant,
        "part": part_name,
        "relationship_id": rel_id,
    }

    published = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=[region_info],
        additions=additions,
        verify=verify_create_publication,
        failure_code="region_staged_verification_failed",
        exists_code="region_destination_occupied",
        unavailable_code="region_publication_unavailable",
    )

    if not published.passed:
        return RegionReport(
            status=_REFUSED,
            code=published.code,
            source_sha256=request.source_sha256,
            destination_sha256=None,
            region=None,
            target_change_report=None,
            non_target_parts_preserved=None,
            retained_unreferenced_parts=None,
        )

    return RegionReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=request.source_sha256,
        destination_sha256=published.output_sha256,
        region=region_info,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
        retained_unreferenced_parts=None,
    )
