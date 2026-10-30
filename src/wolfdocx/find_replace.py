"""Cross-run and cross-paragraph regex find and replace with formatting preservation.

Provides atomic search and replacement spanning arbitrary ``w:r`` run boundaries
within a paragraph or across consecutive paragraphs, preserving boundary formatting,
paragraph properties, inline wrapper ancestry, and untouched package parts. Discovery
includes tracked revision text and field instructions; edits to those protected
values refuse atomically with ``find_replace_protected_text``.
"""

from __future__ import annotations

import bisect
import copy
import hashlib
import json
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Callable, Iterator, Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from wolfdocx.inspect import (
    _inspect_docx,
    _require_verify_mode,
    _unchanged_parts_identical,
    _verify_changed_parts,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from wolfdocx.package import (
    PackageError,
    PackageFormatError,
    PackageGraph,
    PackageSafetyError,
)
from wolfdocx.review import _link_no_clobber

__all__ = [
    "FindReplaceOptions",
    "FindReplaceReport",
    "FindReplaceRequest",
    "discover_matches",
    "find_replace",
    "load_find_replace_request",
    "parse_find_replace_request",
]
_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DOCUMENT = "word/document.xml"
_SHA256_LENGTH = 64


class _FindReplaceValidationError(ValueError):
    """Invalid find/replace values or protected text, identified by a refusal code."""


def _q(tag: str, ns: str = _WORD_NS) -> str:
    return f"{{{ns}}}{tag}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def _word_namespace(root: ET.Element) -> str:
    if root.tag.startswith("{") and "}" in root.tag:
        return root.tag[1 : root.tag.rindex("}")]
    return _WORD_NS


@dataclass(frozen=True, slots=True)
class FindReplaceOptions:
    """Configuration options for regex find and replace operations.

    The searched text joins consecutive paragraphs with ``"\\n"``, so a match may span
    paragraphs. ``multiline`` (the default) makes ``^`` and ``$`` anchor at every
    paragraph and line-break boundary, the way they anchor on each paragraph's text.
    Without it they anchor only at the edges of a whole run of consecutive sibling
    paragraphs (a table cell's content, or the body text between tables).
    """

    ignore_case: bool = False
    multiline: bool = True
    dotall: bool = True
    paragraph_merge_policy: str = "keep_first"  # "keep_first" | "keep_last" | "merge"
    max_replacements: int | None = None

    def __post_init__(self) -> None:
        for name in ("ignore_case", "multiline", "dotall"):
            if not isinstance(getattr(self, name), bool):
                raise _FindReplaceValidationError(f"find_replace_option_invalid:{name}")
        if self.paragraph_merge_policy not in ("keep_first", "keep_last", "merge"):
            raise _FindReplaceValidationError("find_replace_option_invalid:paragraph_merge_policy")
        budget = self.max_replacements
        if budget is not None and (type(budget) is not int or budget < 0):
            raise _FindReplaceValidationError("find_replace_option_invalid:max_replacements")

    def to_dict(self) -> dict[str, Any]:
        return {
            "ignore_case": self.ignore_case,
            "multiline": self.multiline,
            "dotall": self.dotall,
            "paragraph_merge_policy": self.paragraph_merge_policy,
            "max_replacements": self.max_replacements,
        }


class FindReplaceRequest:
    """One bounded find and replace operation over a DOCX package."""

    source: Path
    source_sha256: str
    destination: Path
    pattern: str
    replacement: str | Callable[[re.Match[str]], str]
    part: str
    options: FindReplaceOptions

    def __init__(
        self,
        source: str | Path,
        destination: str | Path | None = None,
        *,
        output: str | Path | None = None,
        pattern: str = "",
        replacement: str | Callable[[re.Match[str]], str] = "",
        source_sha256: str | None = None,
        part: str = _DOCUMENT,
        options: FindReplaceOptions | Mapping[str, Any] | None = None,
    ) -> None:
        src_path = Path(source)
        object.__setattr__(self, "source", src_path)

        dest = destination if destination is not None else output
        if dest is None:
            raise ValueError("find_replace_destination_required")
        object.__setattr__(self, "destination", Path(dest))

        if source_sha256 is not None:
            sha = source_sha256
        elif src_path.is_file():
            sha = _sha256(src_path)
        else:
            sha = ""
        object.__setattr__(self, "source_sha256", sha)

        object.__setattr__(self, "pattern", str(pattern))
        object.__setattr__(self, "replacement", replacement)
        object.__setattr__(self, "part", str(part))

        if isinstance(options, FindReplaceOptions):
            opts = options
        elif isinstance(options, Mapping):
            opts = FindReplaceOptions(**options)
        else:
            opts = FindReplaceOptions()
        object.__setattr__(self, "options", opts)

    @property
    def output(self) -> Path:
        """Alias for destination."""
        return self.destination

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "pattern": self.pattern,
            "replacement": (
                self.replacement if isinstance(self.replacement, str) else "<callable>"
            ),
            "part": self.part,
            "options": self.options.to_dict(),
        }


@dataclass(frozen=True, slots=True)
class FindReplaceReport:
    """The receipt of one find and replace operation."""

    status: str
    code: str | None
    source_sha256: str
    destination: str | None
    match_count: int
    replacement_count: int
    matches: list[dict[str, Any]]
    unselected_parts_preserved: bool

    @property
    def passed(self) -> bool:
        return self.status in ("published", "applied")

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination": self.destination,
            "match_count": self.match_count,
            "replacement_count": self.replacement_count,
            "matches": self.matches,
            "unselected_parts_preserved": self.unselected_parts_preserved,
        }


def _refusal(
    source_sha256: str,
    code: str,
    destination: Path | None = None,
) -> FindReplaceReport:
    return FindReplaceReport(
        status="refused",
        code=code,
        source_sha256=source_sha256,
        destination=str(destination) if destination is not None else None,
        match_count=0,
        replacement_count=0,
        matches=[],
        unselected_parts_preserved=True,
    )


def parse_find_replace_request(data: Any) -> FindReplaceRequest:
    """Parse and validate a dictionary representation of a FindReplaceRequest."""
    if not isinstance(data, Mapping):
        raise TypeError("find_replace_request_invalid")
    source = data.get("source")
    if not source:
        raise ValueError("find_replace_source_required")
    destination = data.get("destination") or data.get("output")
    if not destination:
        raise ValueError("find_replace_destination_required")
    pattern = data.get("pattern", "")
    replacement = data.get("replacement", "")
    source_sha256 = data.get("source_sha256")
    part = data.get("part", _DOCUMENT)
    raw_opts = data.get("options")
    options = FindReplaceOptions(**raw_opts) if isinstance(raw_opts, Mapping) else None
    return FindReplaceRequest(
        source=source,
        destination=destination,
        pattern=pattern,
        replacement=replacement,
        source_sha256=source_sha256,
        part=part,
        options=options,
    )


def load_find_replace_request(path: str | Path) -> FindReplaceRequest:
    """Load and validate a JSON FindReplaceRequest from disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_find_replace_request(json.load(stream))


# ---------------------------------------------------------------------------
# DOM Inspection and Run Text Helpers
# ---------------------------------------------------------------------------

_TEXT_TAGS = frozenset({"t", "tab", "br", "cr", "noBreakHyphen"})


def _extract_run_text_segments(run: ET.Element, ns: str) -> list[tuple[str, ET.Element | None]]:
    """Extract ``(text_piece, source_element)`` from the text-bearing children of a run.

    ``w:t`` text carries no element; ``w:tab``, ``w:br``, ``w:cr`` and ``w:noBreakHyphen``
    carry their source element so a rebuilt run re-emits them as elements. A page break
    contributes no searchable text but is still carried, as a zero-width segment.
    """
    segments: list[tuple[str, ET.Element | None]] = []
    t_tag = _q("t", ns)
    tab_tag = _q("tab", ns)
    br_tag = _q("br", ns)
    cr_tag = _q("cr", ns)
    hyphen_tag = _q("noBreakHyphen", ns)

    for child in run:
        tag = child.tag
        if tag == t_tag:
            segments.append((child.text or "", None))
        elif tag == tab_tag:
            segments.append(("\t", child))
        elif tag == br_tag:
            br_type = child.get(_q("type", ns))
            segments.append(("" if br_type == "page" else "\n", child))
        elif tag == cr_tag:
            segments.append(("\r", child))
        elif tag == hyphen_tag:
            segments.append(("-", child))
    return segments


def _is_text_run(run: ET.Element, ns: str) -> bool:
    """True if run contains only run properties and text-bearing child elements."""
    rpr_tag = _q("rPr", ns)
    for child in run:
        tag = child.tag
        if tag == rpr_tag:
            continue
        local = _local(tag)
        if local not in _TEXT_TAGS:
            return False
    return True


def _make_text(text: str, ns: str) -> ET.Element:
    t_node = ET.Element(_q("t", ns))
    t_node.text = text
    if text.startswith(" ") or text.endswith(" ") or "  " in text or "\t" in text:
        t_node.set(_XML_SPACE, "preserve")
    return t_node


def _make_run(
    parts: list[str | ET.Element],
    rpr: ET.Element | None,
    ns: str = _WORD_NS,
) -> ET.Element:
    """Create a new w:r with a clone of ``rpr`` holding ``parts`` in order.

    Consecutive characters coalesce into one ``w:t``; elements are cloned in place.
    """
    run = ET.Element(_q("r", ns))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    chars: list[str] = []
    for part in parts:
        if isinstance(part, str):
            chars.append(part)
            continue
        if chars:
            run.append(_make_text("".join(chars), ns))
            chars = []
        run.append(copy.deepcopy(part))
    if chars:
        run.append(_make_text("".join(chars), ns))
    return run


def _replacement_parts(text: str, ns: str) -> list[str | ET.Element]:
    """Split replacement text into characters and ``w:tab`` elements for its tabs."""
    if "\t" not in text:
        return [text]
    parts: list[str | ET.Element] = []
    for index, piece in enumerate(text.split("\t")):
        if index:
            parts.append(ET.Element(_q("tab", ns)))
        if piece:
            parts.append(piece)
    return parts


def _rpr_signature(rpr: ET.Element | None) -> str:
    """Canonical string representation of rPr for coalescing runs."""
    if rpr is None:
        return ""
    return ET.tostring(rpr, encoding="unicode")


# ---------------------------------------------------------------------------
# Paragraph Group Processing: Coordinate Mapping & Splicing
# ---------------------------------------------------------------------------


@dataclass
class _CharCoord:
    char: str
    para_idx: int
    rpr: ET.Element | None
    is_para_break: bool = False
    # The source w:tab/w:br/w:cr/w:noBreakHyphen this character stands for, if any.
    element: ET.Element | None = None
    wrappers: tuple[ET.Element, ...] = ()
    protected: bool = False


@dataclass
class _ParagraphItem:
    """One element of a paragraph: either an opaque element or a text run."""

    is_opaque: bool
    opaque_element: ET.Element | None = None
    rpr: ET.Element | None = None
    text: str = ""
    segments: list[tuple[str, ET.Element | None]] = field(default_factory=list)
    wrappers: tuple[ET.Element, ...] = ()
    protected: bool = False


def _decompose_paragraph(p: ET.Element, ns: str) -> tuple[ET.Element | None, list[_ParagraphItem]]:
    """Flatten editable inline content while retaining wrapper ancestry and barriers."""
    ppr = p.find(_q("pPr", ns))
    items: list[_ParagraphItem] = []
    wrapper_tags = {
        _q(tag, ns)
        for tag in ("hyperlink", "sdt", "sdtContent", "customXml", "smartTag", "fldSimple")
    }

    def visit(child: ET.Element, wrappers: tuple[ET.Element, ...]) -> None:
        if child.tag in {_q("ins", ns), _q("del", ns)}:
            text = "".join(
                node.text or ""
                for node in child.iter()
                if node.tag in {_q("t", ns), _q("delText", ns), _q("instrText", ns)}
            )
            items.append(_ParagraphItem(True, child, text=text, wrappers=wrappers, protected=True))
        elif child.tag in wrapper_tags:
            ancestry = (*wrappers, child)
            # Anchor the shell even when every result character is deleted.
            items.append(_ParagraphItem(True, wrappers=ancestry))
            if child.tag == _q("fldSimple", ns):
                items.append(
                    _ParagraphItem(
                        False,
                        text=child.get(_q("instr", ns), ""),
                        wrappers=ancestry,
                        protected=True,
                    )
                )
            for nested in child:
                visit(nested, ancestry)
        elif child.tag == _q("r", ns) and _is_text_run(child, ns):
            segments = _extract_run_text_segments(child, ns)
            items.append(
                _ParagraphItem(
                    False,
                    rpr=child.find(_q("rPr", ns)),
                    text="".join(seg[0] for seg in segments),
                    segments=segments,
                    wrappers=wrappers,
                )
            )
        else:
            text = "".join(
                node.text or ""
                for node in child.iter()
                if node.tag in {_q("instrText", ns), _q("delText", ns)}
            )
            items.append(
                _ParagraphItem(True, child, text=text, wrappers=wrappers, protected=bool(text))
            )

    for child in p:
        if child.tag != _q("pPr", ns):
            visit(child, ())
    return ppr, items


def _apply_ppr_policy(
    ppr_first: ET.Element | None,
    ppr_last: ET.Element | None,
    policy: str,
) -> ET.Element | None:
    """Determine paragraph properties for a merged boundary paragraph."""
    if policy == "keep_last":
        return copy.deepcopy(ppr_last) if ppr_last is not None else None
    if policy == "merge":
        if ppr_first is None:
            return copy.deepcopy(ppr_last) if ppr_last is not None else None
        if ppr_last is None:
            return copy.deepcopy(ppr_first)
        merged = copy.deepcopy(ppr_first)
        first_tags = {child.tag for child in merged}
        for child in ppr_last:
            if child.tag not in first_tags:
                merged.append(copy.deepcopy(child))
            else:
                for idx, existing in enumerate(list(merged)):
                    if existing.tag == child.tag:
                        merged[idx] = copy.deepcopy(child)
                        break
        return merged
    # Default: "keep_first"
    return copy.deepcopy(ppr_first) if ppr_first is not None else None


def _process_paragraph_group(
    group: list[tuple[int, ET.Element]],
    compiled_pattern: re.Pattern[str],
    replacement: str | Callable[[re.Match[str]], str],
    options: FindReplaceOptions,
    ns: str,
    part_name: str,
    remaining_budget: int | None,
    discovery: bool = False,
) -> tuple[list[tuple[int, int, list[ET.Element]]], list[dict[str, Any]], int]:
    """Search and atomically replace patterns across a group of consecutive paragraphs.

    The group's paragraph texts are searched as one ``"\\n"``-joined string, and only
    the paragraphs a match touches (joined with every paragraph a cross-paragraph
    match spans) are rebuilt; every other paragraph keeps its element untouched.

    Returns:
        (rebuilt_ranges, match_records, num_replacements_performed), where each
        rebuilt range is ``(first_child_index, last_child_index, new_paragraphs)``
        in document order.
    """
    if not group:
        return [], [], 0

    decomposed = [_decompose_paragraph(p, ns) for _, p in group]
    texts = ["".join(item.text for item in items) for _, items in decomposed]
    matches = list(compiled_pattern.finditer("\n".join(texts)))

    if remaining_budget is not None and remaining_budget >= 0:
        matches = matches[:remaining_budget]

    if not matches:
        return [], [], 0

    # starts[k] is paragraph k's first offset; its virtual break sits at
    # starts[k + 1] - 1 and belongs to paragraph k.
    starts: list[int] = []
    offset = 0
    for text in texts:
        starts.append(offset)
        offset += len(text) + 1

    def paragraph_at(position: int) -> int:
        return bisect.bisect_right(starts, position) - 1

    if not discovery:
        protected_ranges: list[tuple[int, int]] = []
        for start, (_, items) in zip(starts, decomposed, strict=True):
            for item in items:
                end = start + len(item.text)
                if item.protected:
                    protected_ranges.append((start, end))
                start = end
        for match in matches:
            s, e = match.span()
            if any(
                s < end and e > start or s == e and start <= s < end
                for start, end in protected_ranges
            ):
                raise _FindReplaceValidationError("find_replace_protected_text")

    match_records: list[dict[str, Any]] = []
    # Each cluster: [first_paragraph, last_paragraph, [(start, end, replacement), ...]]
    clusters: list[list[Any]] = []
    for m in matches:
        s, e = m.span()
        try:
            repl_str = replacement(m) if callable(replacement) else m.expand(replacement)
        except (TypeError, ValueError, IndexError, re.error) as exc:
            raise _FindReplaceValidationError("find_replace_replacement_invalid") from exc
        if not isinstance(repl_str, str):
            raise _FindReplaceValidationError("find_replace_replacement_invalid")
        first, last = paragraph_at(s), paragraph_at(e)
        touched_paras = list(range(first, paragraph_at(e - 1) + 1)) if e > s else []
        match_records.append(
            {
                "pattern": compiled_pattern.pattern,
                "matched_text": m.group(0),
                "replacement_text": repl_str,
                "part": part_name,
                "paragraphs": touched_paras,
                "span": [s, e],
            }
        )
        if clusters and first <= clusters[-1][1]:
            clusters[-1][1] = max(clusters[-1][1], last)
            clusters[-1][2].append((s, e, repl_str))
        else:
            clusters.append([first, last, [(s, e, repl_str)]])
    if discovery:
        return [], match_records, len(matches)

    rebuilt: list[tuple[int, int, list[ET.Element]]] = []
    for first, last, spans in clusters:
        base = starts[first]
        new_paragraphs = _rebuild_paragraphs(
            decomposed[first : last + 1],
            [(s - base, e - base, repl_str) for s, e, repl_str in spans],
            options,
            ns,
            [paragraph for _, paragraph in group[first : last + 1]],
        )
        rebuilt.append((group[first][0], group[last][0], new_paragraphs))
    return rebuilt, match_records, len(matches)


def _rebuild_paragraphs(
    decomposed: list[tuple[ET.Element | None, list[_ParagraphItem]]],
    spans: list[tuple[int, int, str]],
    options: FindReplaceOptions,
    ns: str,
    source_paragraphs: list[ET.Element],
) -> list[ET.Element]:
    """Rebuild consecutive paragraphs with ``spans`` replaced in their joined text.

    Span offsets index the paragraphs' texts joined with one virtual ``"\\n"`` break.
    """
    # Build character coordinate stream
    char_coords: list[_CharCoord] = []
    # Elements anchored at a character offset, in document order:
    # (offset, element, in_run, rpr). Opaque paragraph children have ``in_run`` False;
    # zero-width run children (page breaks) are re-emitted inside a run with ``rpr``.
    opaque_anchors: list[
        tuple[int, ET.Element | None, bool, ET.Element | None, tuple[ET.Element, ...]]
    ] = []

    for p_idx, (_, items) in enumerate(decomposed):
        for item in items:
            if item.is_opaque:
                opaque_anchors.append(
                    (len(char_coords), item.opaque_element, False, None, item.wrappers)
                )
            if item.protected:
                char_coords.extend(
                    _CharCoord(ch, p_idx, item.rpr, wrappers=item.wrappers, protected=True)
                    for ch in item.text
                )
                continue
            if item.is_opaque:
                continue
            for text, element in item.segments:
                if element is not None and not text:
                    opaque_anchors.append(
                        (len(char_coords), element, True, item.rpr, item.wrappers)
                    )
                for ch in text:
                    char_coords.append(
                        _CharCoord(
                            char=ch,
                            para_idx=p_idx,
                            rpr=item.rpr,
                            element=element,
                            wrappers=item.wrappers,
                        )
                    )
        # Virtual paragraph break between consecutive paragraphs
        if p_idx < len(decomposed) - 1:
            char_coords.append(
                _CharCoord(
                    char="\n",
                    para_idx=p_idx,
                    rpr=None,
                    is_para_break=True,
                )
            )

    # Map each match to its replacement string and chosen inherited rPr
    # Each entry: (start, end, repl_text, inherited_rpr, touched_paras)
    replacements: list[tuple[int, int, str, ET.Element | None, list[int]]] = []

    for s, e, repl_str in spans:
        # Inherit rpr from the first character in the matched range that has rpr
        inherited_rpr: ET.Element | None = None
        touched_paras: list[int] = []
        for i in range(s, e):
            cc = char_coords[i]
            if cc.para_idx not in touched_paras:
                touched_paras.append(cc.para_idx)
            if inherited_rpr is None and cc.rpr is not None:
                inherited_rpr = copy.deepcopy(cc.rpr)

        replacements.append((s, e, repl_str, inherited_rpr, touched_paras))

    # Reconstruct paragraphs by streaming through characters and replacements
    new_paragraphs: list[ET.Element] = []

    current_para_children: list[ET.Element] = []
    current_run_parts: list[str | ET.Element] = []
    current_run_rpr: ET.Element | None = None
    current_ppr: ET.Element | None = copy.deepcopy(decomposed[0][0])
    current_p_idx = 0
    current_wrappers: tuple[ET.Element, ...] = ()
    current_shell_idx = 0
    emitted_shells: set[int] = set()
    wrapper_copies: dict[int, ET.Element] = {}

    def append_child(element: ET.Element | None, wrappers: tuple[ET.Element, ...]) -> None:
        parent = None
        for wrapper in wrappers:
            clone = wrapper_copies.get(id(wrapper))
            if clone is None:
                clone = ET.Element(wrapper.tag, dict(wrapper.attrib))
                clone.text, clone.tail = wrapper.text, wrapper.tail
                wrapper_copies[id(wrapper)] = clone
                if parent is None:
                    current_para_children.append(clone)
                else:
                    parent.append(clone)
            parent = clone
        if element is None:
            return
        if parent is None:
            current_para_children.append(element)
        else:
            parent.append(element)

    def flush_run() -> None:
        nonlocal current_run_parts, current_run_rpr
        if current_run_parts:
            run_elem = _make_run(current_run_parts, current_run_rpr, ns)
            append_child(run_elem, current_wrappers)
            current_run_parts = []
            current_run_rpr = None

    def close_paragraph() -> None:
        nonlocal current_para_children, current_ppr, wrapper_copies
        flush_run()
        # Retain paragraph identity and lexical namespace provenance. A split's
        # additional paragraphs are new identities, not copies of the old anchor.
        p = copy.copy(source_paragraphs[current_shell_idx])
        del p[:]
        if current_shell_idx in emitted_shells:
            for name in ("paraId", "textId"):
                p.attrib.pop(_q(name, "http://schemas.microsoft.com/office/word/2010/wordml"), None)
        emitted_shells.add(current_shell_idx)
        if current_ppr is not None:
            p.append(current_ppr)
        for child in current_para_children:
            p.append(child)
        new_paragraphs.append(p)
        current_para_children = []
        current_ppr = None
        wrapper_copies = {}

    signatures: dict[int, str] = {}

    def starts_new_run(rpr: ET.Element | None) -> bool:
        # Runs coalesce only across identical rPr; the source rPr elements outlive this
        # call through ``decomposed``, so their ids are stable signature keys.
        if not current_run_parts or rpr is current_run_rpr:
            return False
        signature = []
        for element in (current_run_rpr, rpr):
            if element is None:
                signature.append("")
                continue
            cached = signatures.get(id(element))
            if cached is None:
                cached = signatures[id(element)] = _rpr_signature(element)
            signature.append(cached)
        return signature[0] != signature[1]

    def append_to_run(
        part: str | ET.Element, rpr: ET.Element | None, wrappers: tuple[ET.Element, ...] = ()
    ) -> None:
        nonlocal current_run_rpr, current_wrappers
        if wrappers != current_wrappers or starts_new_run(rpr):
            flush_run()
        current_run_rpr = rpr
        current_wrappers = wrappers
        current_run_parts.append(part)

    def emit_untouched(cc: _CharCoord) -> None:
        nonlocal current_p_idx, current_ppr, current_shell_idx
        if cc.is_para_break:
            close_paragraph()
            current_p_idx += 1
            current_shell_idx = cc.para_idx + 1
            if current_p_idx < len(decomposed):
                current_ppr = copy.deepcopy(decomposed[current_p_idx][0])
        elif not cc.protected:
            append_to_run(cc.char if cc.element is None else cc.element, cc.rpr, cc.wrappers)

    cursor = 0
    opaque_idx = 0

    def check_opaque_before(pos: int) -> None:
        nonlocal opaque_idx
        while opaque_idx < len(opaque_anchors) and opaque_anchors[opaque_idx][0] <= pos:
            _, element, in_run, rpr, wrappers = opaque_anchors[opaque_idx]
            if in_run and element is not None:
                append_to_run(element, rpr, wrappers)
            else:
                flush_run()
                append_child(copy.deepcopy(element), wrappers)
            opaque_idx += 1

    for s, e, repl_str, repl_rpr, touched_paras in replacements:
        # Emit untouched slice before this match
        while cursor < s:
            check_opaque_before(cursor)
            emit_untouched(char_coords[cursor])
            cursor += 1

        check_opaque_before(s)
        repl_wrappers = (
            char_coords[s].wrappers
            if s < len(char_coords)
            else char_coords[s - 1].wrappers
            if s
            else ()
        )

        # Handle boundary paragraph properties if match crossed paragraphs
        if len(touched_paras) > 1:
            first_p_idx = touched_paras[0]
            last_p_idx = touched_paras[-1]
            ppr_first = decomposed[first_p_idx][0]
            ppr_last = decomposed[last_p_idx][0]
            current_ppr = _apply_ppr_policy(ppr_first, ppr_last, options.paragraph_merge_policy)
            if options.paragraph_merge_policy == "keep_last":
                current_shell_idx = last_p_idx

        # Emit replacement
        if "\n" in repl_str:
            lines = repl_str.split("\n")
            for line_idx, line in enumerate(lines):
                if line_idx > 0:
                    close_paragraph()
                    p_idx = touched_paras[-1] if touched_paras else current_p_idx
                    current_ppr = copy.deepcopy(decomposed[p_idx][0])
                    current_shell_idx = p_idx
                if line:
                    flush_run()
                    append_child(
                        _make_run(_replacement_parts(line, ns), repl_rpr, ns), repl_wrappers
                    )
        else:
            if repl_str:
                flush_run()
                append_child(
                    _make_run(_replacement_parts(repl_str, ns), repl_rpr, ns), repl_wrappers
                )

        # Advance cursor past matched region
        cursor = e

    # Emit trailing untouched characters
    while cursor < len(char_coords):
        check_opaque_before(cursor)
        emit_untouched(char_coords[cursor])
        cursor += 1

    # Remaining opaque elements at the end
    check_opaque_before(len(char_coords))
    close_paragraph()

    return new_paragraphs


def _find_paragraph_containers(root: ET.Element, ns: str) -> Iterator[ET.Element]:
    """Visit live paragraph containers in document order, after ancestor splices."""
    p_tag = _q("p", ns)
    stack = [iter((root,))]
    while stack:
        elem = next(stack[-1], None)
        if elem is None:
            stack.pop()
            continue
        if any(child.tag == p_tag for child in elem):
            yield elem
        # Defer child traversal until the caller has rebuilt this container, so text
        # boxes refer to the published host's drawing rather than a detached clone.
        stack.append(iter(elem))


def _process_story_root(
    root: ET.Element,
    compiled_pattern: re.Pattern[str],
    replacement: str | Callable[[re.Match[str]], str],
    options: FindReplaceOptions,
    ns: str,
    part_name: str,
    max_replacements: int | None,
    discovery: bool = False,
) -> tuple[list[dict[str, Any]], int]:
    """Execute find and replace across all paragraph containers in one story root."""
    containers = _find_paragraph_containers(root, ns)
    p_tag = _q("p", ns)

    total_matches: list[dict[str, Any]] = []
    replacements_done = 0

    for container in containers:
        children = list(container)
        groups: list[list[tuple[int, ET.Element]]] = []
        current_group: list[tuple[int, ET.Element]] = []

        for idx, child in enumerate(children):
            if child.tag == p_tag:
                current_group.append((idx, child))
            else:
                if current_group:
                    groups.append(current_group)
                    current_group = []
        if current_group:
            groups.append(current_group)

        # Process groups in reverse order to preserve container slice indices
        for grp in reversed(groups):
            rem_budget = (
                max_replacements - replacements_done if max_replacements is not None else None
            )
            if rem_budget is not None and rem_budget <= 0:
                break

            rebuilt, records, num_done = _process_paragraph_group(
                grp,
                compiled_pattern,
                replacement,
                options,
                ns,
                part_name,
                rem_budget,
                discovery,
            )
            # Splice from the back so earlier child indices stay valid.
            for first_idx, last_idx, new_paras in reversed(rebuilt):
                container[first_idx : last_idx + 1] = new_paras
            total_matches.extend(records)
            replacements_done += num_done

    return total_matches, replacements_done


def _compile_pattern(pattern: str, options: FindReplaceOptions) -> re.Pattern[str] | None:
    """Compile ``pattern`` under ``options``, or ``None`` when it is not a valid regex."""
    flags = 0
    if options.ignore_case:
        flags |= re.IGNORECASE
    if options.multiline:
        flags |= re.MULTILINE
    if options.dotall:
        flags |= re.DOTALL
    try:
        return re.compile(pattern, flags)
    except re.error:
        return None


def _replace_in_story(
    root: ET.Element, compiled_pattern: re.Pattern[str], request: FindReplaceRequest
) -> tuple[list[dict[str, Any]], int, bytes]:
    """Rewrite ``request.part``'s editable tree in place and serialize it.

    ``root`` must be parsed with preserved qualified names. Returns the match records,
    the replacement count, and the rewritten part bytes; a call that replaces nothing
    leaves ``root`` untouched and returns empty bytes, which callers refuse with
    ``find_replace_no_match`` rather than publish.
    """
    matches, count = _process_story_root(
        root,
        compiled_pattern,
        request.replacement,
        request.options,
        _word_namespace(root),
        request.part,
        request.options.max_replacements,
    )
    if not count:
        return matches, 0, b""
    return matches, count, serialize_xml_with_preserved_qnames(root)


# ---------------------------------------------------------------------------
# Package Writing & Publication
# ---------------------------------------------------------------------------


def _write_package(
    source: Path,
    output: Path,
    replacements: Mapping[str, bytes],
) -> None:
    """Write rewritten package, preserving every untouched part and entry byte-for-byte."""
    with (
        zipfile.ZipFile(source) as src_zip,
        zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip,
    ):
        seen: set[str] = set()
        for info in src_zip.infolist():
            if info.filename in seen:
                continue
            seen.add(info.filename)
            if info.filename in replacements:
                dst_zip.writestr(info, replacements[info.filename])
            else:
                dst_zip.writestr(info, src_zip.read(info.filename))


# ---------------------------------------------------------------------------
# Public Functions
# ---------------------------------------------------------------------------


def discover_matches(
    source: str | Path,
    pattern: str,
    *,
    part: str = _DOCUMENT,
    options: FindReplaceOptions | Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Read regex matches, including inline wrappers, revisions and field instructions."""
    src_path = Path(source).resolve()
    if not src_path.is_file():
        return {"passed": False, "code": "find_replace_source_missing"}

    digest = _sha256(src_path)

    if isinstance(options, FindReplaceOptions):
        opts = options
    elif isinstance(options, Mapping):
        opts = FindReplaceOptions(**options)
    else:
        opts = FindReplaceOptions()

    compiled_pattern = _compile_pattern(pattern, opts)
    if compiled_pattern is None:
        return {
            "passed": False,
            "code": "find_replace_pattern_invalid",
            "source_sha256": digest,
        }

    try:
        graph = PackageGraph.from_file(src_path)
    except (
        PackageError,
        PackageSafetyError,
        PackageFormatError,
        OSError,
        zipfile.BadZipFile,
    ):
        return {
            "passed": False,
            "code": "source_package_invalid",
            "source_sha256": digest,
        }

    if not graph.has_part(part):
        return {
            "passed": False,
            "code": "find_replace_part_unsupported",
            "source_sha256": digest,
        }

    try:
        root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, KeyError):
        return {
            "passed": False,
            "code": "source_package_invalid",
            "source_sha256": digest,
        }

    ns = _word_namespace(root)
    # Perform discovery on a deep clone without persisting
    clone_root = copy.deepcopy(root)
    matches, _ = _process_story_root(
        clone_root,
        compiled_pattern,
        "",
        opts,
        ns,
        part,
        opts.max_replacements,
        discovery=True,
    )

    return {
        "passed": True,
        "source_sha256": digest,
        "pattern": pattern,
        "part": part,
        "match_count": len(matches),
        "matches": matches,
    }


def find_replace(request: FindReplaceRequest, *, verify: str = "full") -> FindReplaceReport:
    """Execute atomic regex find and replace across runs and consecutive paragraphs.

    Preserve hyperlink, SDT, customXml, smartTag and simple-field result wrappers.
    Matches touching revision text or field instructions refuse with
    ``find_replace_protected_text``. Invalid replacement values or group references
    refuse with ``find_replace_replacement_invalid``; neither refusal publishes output.

    A pattern that replaces nothing refuses with ``find_replace_no_match`` and publishes
    nothing, matching ``Document.find_replace``.

    Both verification modes require every staged part other than ``request.part``
    (relationships included) to hash identically to the validated source.
    ``verify="full"`` (the default) additionally inspects the whole staged package;
    ``verify="changed-parts"`` re-parses only ``request.part`` under the package
    limits, relying on the source inspection for the byte-identical parts.
    """
    if not isinstance(request, FindReplaceRequest):
        raise TypeError("find_replace_request_invalid")
    _require_verify_mode(verify)

    source = request.source.resolve()
    destination = request.destination.resolve()

    if not source.is_file():
        return _refusal(request.source_sha256, "find_replace_source_missing", destination)

    current_hash = _sha256(source)
    if request.source_sha256 and current_hash != request.source_sha256:
        return _refusal(request.source_sha256, "source_hash_changed", destination)

    if destination.exists() or destination.is_symlink():
        return _refusal(current_hash, "destination_exists", destination)

    if destination == source or not destination.parent.is_dir():
        return _refusal(current_hash, "destination_invalid", destination)

    compiled_pattern = _compile_pattern(request.pattern, request.options)
    if compiled_pattern is None:
        return _refusal(current_hash, "find_replace_pattern_invalid", destination)

    # Open package graph and inspect
    try:
        graph = PackageGraph.from_file(source)
    except (
        PackageError,
        PackageSafetyError,
        PackageFormatError,
        OSError,
        zipfile.BadZipFile,
    ):
        return _refusal(current_hash, "source_package_invalid", destination)

    baseline = _inspect_docx(source)
    if not baseline.valid:
        return _refusal(current_hash, "source_package_invalid", destination)

    if not graph.has_part(request.part):
        return _refusal(current_hash, "find_replace_part_unsupported", destination)

    try:
        part_bytes = graph.read_part_bytes(request.part)
        root = parse_xml_with_preserved_qnames(part_bytes)
    except (ET.ParseError, KeyError):
        return _refusal(current_hash, "source_package_invalid", destination)

    try:
        matches, count, new_part_bytes = _replace_in_story(root, compiled_pattern, request)
    except _FindReplaceValidationError as exc:
        return _refusal(current_hash, str(exc), destination)
    if not count:
        return _refusal(current_hash, "find_replace_no_match", destination)
    replacements = {request.part: new_part_bytes}

    # Atomic stage and publish
    staged: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as tmp:
            staged = Path(tmp.name)

        _write_package(source, staged, replacements)

        # Staged verification: the whole package, or (changed-parts) only the edited
        # part. Either way every other part must be byte-identical to the validated
        # source, relationships and content types included.
        if baseline.file_sha256 != current_hash:
            staged_valid = False
        elif verify == "full":
            candidate = _inspect_docx(staged, reuse=baseline)
            staged_valid = candidate.valid and _unchanged_parts_identical(
                baseline.part_sha256,
                candidate.part_sha256,
                candidate.part_sha256.__getitem__,
                replacements,
            )
        else:
            staged_valid = _verify_changed_parts(baseline, staged, replacements) is not None
        if not staged_valid:
            return _refusal(current_hash, "find_replace_staged_verification_failed", destination)

        # Source concurrent change verification
        if _sha256(source) != current_hash:
            return _refusal(current_hash, "source_hash_changed", destination)

        # No-clobber publish: a destination created after the pre-check is preserved.
        linked = _link_no_clobber(
            staged,
            destination,
            exists_code="destination_exists",
            unavailable_code="find_replace_publication_unavailable",
        )
        if linked is not None:
            return _refusal(current_hash, linked, destination)

        return FindReplaceReport(
            status="published",
            code=None,
            source_sha256=current_hash,
            destination=str(destination),
            match_count=len(matches),
            replacement_count=count,
            matches=matches,
            unselected_parts_preserved=True,
        )
    finally:
        if staged is not None:
            staged.unlink(missing_ok=True)
