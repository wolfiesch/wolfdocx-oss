"""Guarded direct-body run text editing, without rebuilding sibling runs."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from .authoring import _MAX_TEXT_CHARACTERS
from .paragraph_authoring import _is_valid_xml_text

__all__ = ["RunTextEditError"]

_WORD_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    }
)
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_TEXT_CHILDREN = frozenset({"t", "tab", "br", "cr", "noBreakHyphen"})


class RunTextEditError(ValueError):
    """An addressed edit refused before changing the session, identified by ``code``."""

    def __init__(self, code: str) -> None:
        self.code = code
        super().__init__(code)


def _validate_request(paragraph: int, run: int, text: str, expected_old_text: str | None) -> None:
    for name, index in (("paragraph", paragraph), ("run", run)):
        if type(index) is not int or index < 0:
            raise RunTextEditError(f"run_text_{name}_index_invalid")
    for name, value in (("text", text), ("expected_old_text", expected_old_text)):
        if value is None and name == "expected_old_text":
            continue
        if not isinstance(value, str) or not _is_valid_xml_text(value):
            raise RunTextEditError(f"run_text_{name}_invalid")
        if len(value) > _MAX_TEXT_CHARACTERS:
            raise RunTextEditError(f"run_text_{name}_limit_exceeded")


def _run_text(run: ET.Element, namespace: str) -> str:
    """Visible direct-child text, using the same controls as the body run view."""
    values = []
    prefix = f"{{{namespace}}}"
    for child in run:
        if child.tag == prefix + "t":
            values.append(child.text or "")
        elif child.tag == prefix + "tab":
            values.append("\t")
        elif child.tag == prefix + "cr":
            values.append("\n")
        elif child.tag == prefix + "br":
            if child.get(prefix + "type", "textWrapping") == "textWrapping":
                values.append("\n")
        elif child.tag == prefix + "noBreakHyphen":
            values.append("-")
    return "".join(values)


def _text_children(text: str, namespace: str) -> list[ET.Element]:
    children: list[ET.Element] = []
    start = 0

    def append_text(stop: int) -> None:
        if stop == start:
            return
        value = text[start:stop]
        child = ET.Element(f"{{{namespace}}}t")
        child.text = value
        if value[0].isspace() or value[-1].isspace():
            child.set(_XML_SPACE, "preserve")
        children.append(child)

    controls = {"\t": "tab", "\n": "br", "\r": "cr", "\u2011": "noBreakHyphen"}
    for index, char in enumerate(text):
        tag = controls.get(char)
        if tag is not None:
            append_text(index)
            children.append(ET.Element(f"{{{namespace}}}{tag}"))
            start = index + 1
    append_text(len(text))
    return children


def _replace_run_text(
    root: ET.Element, paragraph: int, run: int, text: str, expected_old_text: str | None
) -> None:
    namespace = root.tag.removeprefix("{").partition("}")[0]
    if namespace not in _WORD_NAMESPACES or root.tag != f"{{{namespace}}}document":
        raise RunTextEditError("run_text_document_invalid")
    body = root.find(f"{{{namespace}}}body")
    paragraphs = [] if body is None else body.findall(f"{{{namespace}}}p")
    if paragraph >= len(paragraphs):
        raise RunTextEditError("run_text_paragraph_out_of_range")
    runs = paragraphs[paragraph].findall(f"{{{namespace}}}r")
    if run >= len(runs):
        raise RunTextEditError("run_text_run_out_of_range")
    selected = runs[run]
    if expected_old_text is not None and _run_text(selected, namespace) != expected_old_text:
        raise RunTextEditError("run_text_guard_mismatch")

    text_tags = {f"{{{namespace}}}{tag}" for tag in _TEXT_CHILDREN}
    replacement = _text_children(text, namespace)
    children: list[ET.Element] = []
    inserted = False
    for child in selected:
        if child.tag in text_tags:
            if not inserted:
                children.extend(replacement)
                inserted = True
        else:
            children.append(child)
    if not inserted:
        children.extend(replacement)
    selected[:] = children
