"""Standard-library helpers that build small DOCX packages for the tests.

Every fixture is generated from inline XML so the suite needs no binary fixture files
and no third-party packages beyond pytest.
"""

from __future__ import annotations

import re
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CONTENT_TYPES_DEFAULTS = (
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Default Extension="jpeg" ContentType="image/jpeg"/>'
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    f'Type="{R}/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)
PART_OVERRIDES = {
    "word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "word/settings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    "word/header1.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "word/footer1.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
}
STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles xmlns:w="{W}">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:eastAsia="{family}" w:cs="{family}"/>'
    '<w:sz w:val="24"/><w:szCs w:val="24"/>'
    "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/>'
    '<w:jc w:val="center"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/>'
    "<w:rPr><w:i/></w:rPr></w:style>"
    "</w:styles>"
)

#: Content types for well-known untyped fixture parts; anything else is opaque bytes.
KNOWN_PART_CONTENT_TYPES = {"word/vbaProject.bin": "application/vnd.ms-office.vbaProject"}


def styles_xml(family: str = "Calibri") -> str:
    return STYLES.replace("{family}", family)


def settings_xml(**flags: bool) -> str:
    extra = "".join(
        f"<w:{name}/>" if value else f'<w:{name} w:val="false"/>' for name, value in flags.items()
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W}">{extra}</w:settings>'


def sect(
    *,
    width: int = 12240,
    height: int = 15840,
    top: int = 1440,
    right: int = 1440,
    bottom: int = 1440,
    left: int = 1440,
    header: int = 720,
    footer: int = 720,
    title_page: bool = False,
    columns: int = 1,
    references: str = "",
    extra: str = "",
) -> str:
    """One section with explicit page geometry and optional region references."""

    title = "<w:titlePg/>" if title_page else ""
    column_properties = f'<w:cols w:num="{columns}"/>' if columns != 1 else "<w:cols/>"
    return (
        "<w:sectPr>"
        f"{references}"
        f'<w:pgSz w:w="{width}" w:h="{height}"/>'
        f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" '
        f'w:left="{left}" w:header="{header}" w:footer="{footer}"/>'
        f"{column_properties}{title}{extra}"
        "</w:sectPr>"
    )


def header_reference(identifier: str, variant: str = "default") -> str:
    return f'<w:headerReference w:type="{variant}" r:id="{identifier}" xmlns:r="{R}"/>'


def footer_reference(identifier: str, variant: str = "default") -> str:
    return f'<w:footerReference w:type="{variant}" r:id="{identifier}" xmlns:r="{R}"/>'


def region_relationship(identifier: str, kind: str, target: str) -> str:
    return f'<Relationship Id="{identifier}" Type="{R}/{kind}" Target="{target}"/>'


def paragraph(
    text: str,
    *,
    properties: str = "",
    run_properties: str = "",
    style: str | None = None,
) -> str:
    """One paragraph, or one empty paragraph when ``text`` is empty."""

    style_element = f'<w:pStyle w:val="{style}"/>' if style else ""
    ppr = f"<w:pPr>{style_element}{properties}</w:pPr>" if (style_element or properties) else ""
    rpr = f"<w:rPr>{run_properties}</w:rPr>" if run_properties else ""
    run = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{run}</w:p>"


def table(rows: str, *, properties: str = "", grid: str = "") -> str:
    return f"<w:tbl><w:tblPr>{properties}</w:tblPr>{grid}{rows}</w:tbl>"


def grid(columns: Sequence[int]) -> str:
    cells = "".join(f'<w:gridCol w:w="{value}"/>' for value in columns)
    return f"<w:tblGrid>{cells}</w:tblGrid>"


def row(cells: str, *, properties: str = "") -> str:
    return f"<w:tr><w:trPr>{properties}</w:trPr>{cells}</w:tr>"


def cell(content: str, *, properties: str = "", span: int = 1) -> str:
    span_element = f'<w:gridSpan w:val="{span}"/>' if span != 1 else ""
    return f"<w:tc><w:tcPr>{span_element}{properties}</w:tcPr>{content}</w:tc>"


def _without_repeats(fragment: str, builtin: str) -> str:
    """Drop declarations that repeat a built-in Default or Override exactly."""

    if not fragment:
        return fragment
    for declaration in re.findall(r"<(?:Default|Override) [^>]*/>", builtin):
        fragment = fragment.replace(declaration, "")
    return fragment


def declare_untyped_parts(types: str, names: Sequence[str]) -> str:
    """Add an Override for each part that no Default or Override covers.

    OPC recognizes only parts that have a content type (ECMA-376 Part 2, M1.13).
    """

    extensions = {value.lower() for value in re.findall(r'Extension="([^"]*)"', types)}
    declared = {value.lstrip("/").lower() for value in re.findall(r'PartName="([^"]*)"', types)}
    missing = [
        name
        for name in dict.fromkeys(names)
        if name != "[Content_Types].xml"
        and not name.endswith("/")
        and name.lower() not in declared
        and name.rpartition(".")[2].lower() not in extensions
    ]
    additions = "".join(
        f'<Override PartName="/{name}" ContentType="'
        f'{KNOWN_PART_CONTENT_TYPES.get(name, "application/octet-stream")}"/>'
        for name in missing
    )
    return types.replace("</Types>", f"{additions}</Types>")


def write_package(
    path: Path,
    body: str,
    *,
    sect_properties: str | None = None,
    parts: Mapping[str, bytes | str] | None = None,
    relationships: str = "",
    overrides: str = "",
    family: str = "Calibri",
    settings: str | None = None,
    content_types_extra: str = "",
) -> Path:
    """Write one DOCX package with an explicit main document part.

    Every caller part without a declared Default or Override receives an Override
    (see :func:`declare_untyped_parts`), because OPC recognizes only parts that
    have a content type. Caller declarations that repeat a built-in one are dropped.
    A caller that supplies its own ``[Content_Types].xml`` part keeps it unchanged.
    """

    section = sect_properties if sect_properties is not None else sect()
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}{section}</w:body></w:document>'
    )
    builtin = (
        f"{CONTENT_TYPES_DEFAULTS}"
        '<Override PartName="/word/document.xml" ContentType="'
        f'{PART_OVERRIDES["word/document.xml"]}"/>'
        '<Override PartName="/word/styles.xml" ContentType="'
        f'{PART_OVERRIDES["word/styles.xml"]}"/>'
        '<Override PartName="/word/settings.xml" ContentType="'
        f'{PART_OVERRIDES["word/settings.xml"]}"/>'
    )
    types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f"{builtin}{_without_repeats(overrides, builtin)}"
        f"{_without_repeats(content_types_extra, builtin)}</Types>"
    )
    document_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rIdStyles" Type="{R}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rIdSettings" Type="{R}/settings" Target="settings.xml"/>'
        f"{relationships}</Relationships>"
    )
    payload: dict[str, bytes | str] = {
        "[Content_Types].xml": types,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": document_rels,
        "word/document.xml": document,
        "word/styles.xml": styles_xml(family),
        "word/settings.xml": settings if settings is not None else settings_xml(),
    }
    extra = dict(parts or {})
    if "[Content_Types].xml" not in extra:
        payload["[Content_Types].xml"] = declare_untyped_parts(types, [*payload, *extra])
    payload.update(extra)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, value in payload.items():
            metadata = zipfile.ZipInfo(name, date_time=(2020, 1, 1, 0, 0, 0))
            metadata.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(metadata, value)
    return path


# -- Styled table package ------------------------------------------------------

_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_CNF_ATTRIBUTES = (
    "firstRow",
    "lastRow",
    "firstColumn",
    "lastColumn",
    "oddVBand",
    "evenVBand",
    "oddHBand",
    "evenHBand",
    "firstRowFirstColumn",
    "firstRowLastColumn",
    "lastRowFirstColumn",
    "lastRowLastColumn",
)
#: Cell texts shared by all three tables: four rows of two cells.
STYLED_TABLE_TEXT = (
    ("foobar", "barfoo"),
    ("foo", "bar"),
    ("bar", "foo"),
    ("foo to you", "and a hearty foo to you too sir!"),
)


def _cnf(value: str) -> str:
    flags = "".join(
        f' w:{name}="{value[index]}"' for index, name in enumerate(_CNF_ATTRIBUTES)
    )
    return f'<w:cnfStyle w:val="{value}"{flags}/>'


def _styled_cell_paragraph(
    text: str, identifier: int, *, spell: bool, properties: str = "", prefix: str = ""
) -> str:
    run = f"<w:r><w:t>{text}</w:t></w:r>"
    if spell:
        run = f'<w:proofErr w:type="spellStart"/>{run}<w:proofErr w:type="spellEnd"/>'
    ppr = f"<w:pPr>{properties}</w:pPr>" if properties else ""
    return (
        f'<w:p w14:paraId="{identifier:08X}" w14:textId="77777777" '
        'w:rsidR="000D62E2" w:rsidRDefault="000D62E2" w:rsidP="0004702B">'
        f"{ppr}{prefix}{run}</w:p>"
    )


def _styled_table(index: int) -> str:
    """One two-column table; the third carries conditional formatting and a bookmark."""

    style = {1: "", 2: '<w:tblStyle w:val="TableGrid"/>'}.get(
        index, '<w:tblStyle w:val="LightShading-Accent1"/>'
    )
    conditional = index == 3
    row_cnf = {1: "100000000000", 2: "000000100000", 4: "000000100000"}
    cell_cnf = {1: "100000000000", 2: "000000100000", 3: "000000000000", 4: "000000100000"}
    rows = []
    for row_number, texts in enumerate(STYLED_TABLE_TEXT, start=1):
        trpr = (
            f"<w:trPr>{_cnf(row_cnf[row_number])}</w:trPr>"
            if conditional and row_number in row_cnf
            else ""
        )
        cells = []
        for cell_number, text in enumerate(texts, start=1):
            first_column = conditional and cell_number == 1
            tcpr = (_cnf("001000000000") if first_column else "") + (
                '<w:tcW w:w="4428" w:type="dxa"/>'
            )
            ppr = _cnf(cell_cnf[row_number]) if conditional and cell_number == 2 else ""
            bookmark = (
                '<w:bookmarkStart w:id="0" w:name="_GoBack"/>'
                if conditional and row_number == 1 and cell_number == 1
                else ""
            )
            identifier = 0x10000000 * index + 0x100 * row_number + cell_number
            cells.append(
                f"<w:tc><w:tcPr>{tcpr}</w:tcPr>"
                + _styled_cell_paragraph(
                    text,
                    identifier,
                    spell=row_number == 1 and cell_number == 2,
                    properties=ppr,
                    prefix=bookmark,
                )
                + "</w:tc>"
            )
        rows.append(
            f'<w:tr w:rsidR="000D62E2" w14:paraId="{0x0A000000 * index + row_number:08X}" '
            'w14:textId="77777777" w:rsidTr="0004702B">' + trpr + "".join(cells) + "</w:tr>"
        )
    bookmark_end = '<w:bookmarkEnd w:id="0"/>' if conditional else ""
    return (
        f'<w:tbl><w:tblPr>{style}<w:tblW w:w="0" w:type="auto"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
        'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4428"/><w:gridCol w:w="4428"/></w:tblGrid>'
        + "".join(rows)
        + bookmark_end
        + "</w:tbl>"
    )


def _empty_body_paragraph(identifier: int) -> str:
    return (
        f'<w:p w14:paraId="{identifier:08X}" w14:textId="77777777" '
        'w:rsidR="000D62E2" w:rsidRDefault="000D62E2" w:rsidP="000D62E2"/>'
    )


def styled_table_document_xml() -> str:
    """Three styled two-column tables separated by empty paragraphs.

    The main part declares the ``w14``/``wp14`` extension namespaces and lists them in
    ``mc:Ignorable``, as word processors write them. The first ``<w:t>bar</w:t>`` is
    table 1, row 2, cell 2 and the first ``<w:r><w:t>foobar</w:t></w:r>`` is the
    first paragraph in document order.
    """

    body = "".join(
        _styled_table(index) + _empty_body_paragraph(0x7F000000 + index) for index in (1, 2, 3)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:mc="{_MC}" xmlns:r="{R}" xmlns:wp14="{_WP14}" '
        f'xmlns:wp="{_WP}" xmlns:w="{W}" xmlns:w14="{_W14}" mc:Ignorable="w14 wp14">'
        f"<w:body>{body}"
        '<w:sectPr w:rsidR="000D62E2" w:rsidSect="00034616">'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/>'
        "</w:sectPr></w:body></w:document>"
    )


def _styled_table_styles_xml() -> str:
    table_style = (
        '<w:style w:type="table" w:styleId="{id}"><w:name w:val="{name}"/>'
        '<w:basedOn w:val="TableNormal"/><w:uiPriority w:val="{priority}"/>'
        "<w:tblPr><w:tblBorders>{borders}</w:tblBorders></w:tblPr></w:style>"
    )
    grid_borders = "".join(
        f'<w:{edge} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    shading_borders = "".join(
        f'<w:{edge} w:val="single" w:sz="8" w:space="0" w:color="4F81BD"/>'
        for edge in ("top", "bottom")
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:mc="{_MC}" xmlns:r="{R}" xmlns:w="{W}" xmlns:w14="{_W14}" '
        'mc:Ignorable="w14">'
        "<w:docDefaults><w:rPrDefault><w:rPr>"
        '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/>'
        '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>'
        "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/>'
        "<w:semiHidden/><w:unhideWhenUsed/></w:style>"
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
        '<w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/>'
        "<w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/>"
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
        "</w:tblPr></w:style>"
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList">'
        '<w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/>'
        "<w:unhideWhenUsed/></w:style>"
        + table_style.format(id="TableGrid", name="Table Grid", priority=59, borders=grid_borders)
        + table_style.format(
            id="LightShading-Accent1",
            name="Light Shading Accent 1",
            priority=60,
            borders=shading_borders,
        )
        + "</w:styles>"
    )


def write_styled_table_package(path: Path) -> Path:
    """Write the three-table package used by the cell, span and row authoring tests."""

    office = "application/vnd.openxmlformats-officedocument"
    types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Override PartName="/word/document.xml" ContentType="{office}.wordprocessingml.document.main+xml"/>'
        f'<Override PartName="/word/styles.xml" ContentType="{office}.wordprocessingml.styles+xml"/>'
        f'<Override PartName="/word/settings.xml" ContentType="{office}.wordprocessingml.settings+xml"/>'
        f'<Override PartName="/word/webSettings.xml" ContentType="{office}.wordprocessingml.webSettings+xml"/>'
        f'<Override PartName="/word/fontTable.xml" ContentType="{office}.wordprocessingml.fontTable+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        f'<Override PartName="/docProps/app.xml" ContentType="{office}.extended-properties+xml"/>'
        "</Types>"
    )
    package_rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{package_rels}">'
        f'<Relationship Id="rId3" Type="{package_rels}/metadata/core-properties" Target="docProps/core.xml"/>'
        f'<Relationship Id="rId4" Type="{R}/extended-properties" Target="docProps/app.xml"/>'
        f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    document_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{package_rels}">'
        f'<Relationship Id="rId3" Type="{R}/settings" Target="settings.xml"/>'
        f'<Relationship Id="rId4" Type="{R}/webSettings" Target="webSettings.xml"/>'
        f'<Relationship Id="rId5" Type="{R}/fontTable" Target="fontTable.xml"/>'
        f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    settings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:mc="{_MC}" xmlns:r="{R}" xmlns:w="{W}" xmlns:w14="{_W14}" '
        'mc:Ignorable="w14"><w:zoom w:percent="150"/><w:proofState w:spelling="clean"/>'
        '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>'
        '<w:themeFontLang w:val="en-US"/></w:settings>'
    )
    web_settings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:webSettings xmlns:mc="{_MC}" xmlns:r="{R}" xmlns:w="{W}" xmlns:w14="{_W14}" '
        'mc:Ignorable="w14"><w:allowPNG/><w:doNotSaveAsSingleFile/></w:webSettings>'
    )
    font_table = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts xmlns:r="{R}" xmlns:w="{W}">'
        '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/>'
        '<w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
        "</w:fonts>"
    )
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title></dc:title><dc:creator>Fixture</dc:creator><cp:revision>4</cp:revision>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2013-12-23T18:29:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2015-02-08T22:05:00Z</dcterms:modified>'
        "</cp:coreProperties>"
    )
    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Template>Normal.dotm</Template><TotalTime>5</TotalTime><Pages>1</Pages>"
        "<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>"
        "<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged></Properties>"
    )
    members = (
        ("[Content_Types].xml", types),
        ("_rels/.rels", root_rels),
        ("word/_rels/document.xml.rels", document_rels),
        ("word/document.xml", styled_table_document_xml()),
        ("word/settings.xml", settings),
        ("word/webSettings.xml", web_settings),
        ("docProps/core.xml", core),
        ("word/styles.xml", _styled_table_styles_xml()),
        ("word/fontTable.xml", font_table),
        ("docProps/app.xml", app),
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, value in members:
            metadata = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            metadata.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(metadata, value)
    return path


# -- Independent reader --------------------------------------------------------


@dataclass(frozen=True)
class RunView:
    """The visible text and direct bold flag of one run, read without the library."""

    text: str
    bold: bool | None


def _run_text(run: ET.Element) -> str:
    """Run text as WordprocessingML defines its inline content.

    ``w:t`` contributes its characters, ``w:tab`` and ``w:ptab`` a tab, ``w:cr`` and a
    text-wrapping ``w:br`` a line feed, ``w:noBreakHyphen`` a hyphen. Page and column
    breaks contribute nothing.
    """

    pieces = []
    for child in run:
        local = child.tag.rpartition("}")[2]
        if local == "t":
            pieces.append(child.text or "")
        elif local in {"tab", "ptab"}:
            pieces.append("\t")
        elif local == "cr":
            pieces.append("\n")
        elif local == "br":
            if child.get(f"{{{W}}}type", "textWrapping") == "textWrapping":
                pieces.append("\n")
        elif local == "noBreakHyphen":
            pieces.append("-")
    return "".join(pieces)


def _run_bold(run: ET.Element) -> bool | None:
    properties = run.find(f"{{{W}}}rPr")
    bold = None if properties is None else properties.find(f"{{{W}}}b")
    if bold is None:
        return None
    return bold.get(f"{{{W}}}val", "true") not in {"false", "0", "off"}


def body_runs(source: Path | BinaryIO) -> list[list[RunView]]:
    """Direct body paragraphs and their direct runs, read with the standard library.

    Tables, content controls, hyperlinks and revision wrappers are skipped, so only
    ``w:body/w:p/w:r`` contributes.
    """

    with zipfile.ZipFile(source) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(f"{{{W}}}body")
    assert body is not None
    return [
        [RunView(_run_text(run), _run_bold(run)) for run in paragraph.findall(f"{{{W}}}r")]
        for paragraph in body.findall(f"{{{W}}}p")
    ]
