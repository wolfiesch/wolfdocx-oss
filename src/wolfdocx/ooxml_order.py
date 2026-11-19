"""Schema child order of WordprocessingML property containers.

Writers that insert or merge paragraph and run properties reorder the children
with these ranks so the result follows the ``CT_PPrBase`` and ``CT_RPr``
sequences. Importers, assembly, and session edits share one definition.
"""

from __future__ import annotations

from collections.abc import Mapping
from xml.etree import ElementTree as ET

# CT_PPrBase and CT_RPr child sequences (ECMA-376 Part 1, 17.3.1.26 and 17.3.2.28).
_PPR_ORDER = {
    name: index
    for index, name in enumerate(
        (
            "pStyle",
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "framePr",
            "widowControl",
        )
        + ("numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens")
        + ("kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE")
        + ("autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind")
        + (
            "contextualSpacing",
            "mirrorIndents",
            "suppressOverlap",
            "jc",
            "textDirection",
        )
        + (
            "textAlignment",
            "textboxTightWrap",
            "outlineLvl",
            "divId",
            "cnfStyle",
            "rPr",
        )
    )
}
_RPR_ORDER = {
    name: index
    for index, name in enumerate(
        ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike")
        + ("dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid")
        + ("vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz")
        + ("szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign")
        + ("rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")
    )
}


def _sort_children(element: ET.Element, order: Mapping[str, int]) -> None:
    children = sorted(
        element, key=lambda child: order.get(child.tag.rpartition("}")[2], len(order))
    )
    element[:] = children
