"""Source-bound native Word document property discovery and authoring.

Implements inspection and mutation of core, typed custom, and extended document
properties, plus document variables in the settings part related from the main
document. Variable names match case-insensitively using ``str.lower`` (as field
evaluation does); string values upsert and ``None`` removes. Existing name
spelling is retained. Variables never automatically refresh fields.

Strict and Transitional property parts and Custom XML stores are read and written
in their source namespace profile. Custom property names are casefolded identities;
updates retain source spelling, pid, fmtid, linkTarget, and untouched value XML.
Use ``link_target`` and ``fmtid`` on custom edits to author those attributes; an
empty link_target removes the attribute, while None keeps the existing value.
Exact VT scalar names are supported alongside text/number/bool/date aliases.
The legacy date alias writes filetime; ``vt_date`` writes the VT date element.
Decimal and currency values use decimal strings, binary values use base64, and
empty/null use None. Vstream values carry ``value`` (base64) and ``version`` (GUID
or None). Structured vectors carry ``base_type`` and ``values``; arrays also carry
``lower_bounds`` and ``upper_bounds`` lists. Variant members carry ``type`` and
``value``. Discovery retains legacy labels and adds the exact ``vt_type``.
Extended heading_pairs, titles_of_parts, and h_links use that vector model;
dig_sig uses base64. All other extended fields use strings, integers, or booleans.

The UI measurement unit is a Word application preference, not a DOCX setting
(Microsoft Word ``Options.MeasurementUnit``:
https://learn.microsoft.com/en-us/office/vba/api/word.options.measurementunit).
The request accepts ``mm``, ``cm``, ``inches``, ``points``, or ``picas`` for strict
validation, but any non-``None`` unit is refused with
``properties_measurement_unit_unsupported`` rather than inventing OOXML.
Discovery and reports return ``measurement_unit=None``.

Relationships are resolved dynamically rather than assuming canonical part
names. Missing parts, relationships, and content types are created when needed.
Unknown XML and QName-valued attributes are preserved semantically, and all
non-target package parts remain byte-for-byte unchanged.
"""

from __future__ import annotations

import base64
import binascii
import copy
import hashlib
import json
import math
import os
import posixpath
import re
import tempfile
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from datetime import UTC, date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .inspect import (
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .protection import _insert_settings_child
from .review import ReviewReport, _publish, _sha, _sha256, _word_namespace

CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
DCMITYPE_NS = "http://purl.org/dc/dcmitype/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

CUSTPROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VTYPES_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CORE_PROPERTIES_REL_TYPE = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)
CUSTOM_PROPERTIES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
)
EXTENDED_PROPERTIES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
)

CORE_PROPERTIES_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml"
CUSTOM_PROPERTIES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.custom-properties+xml"
)
EXTENDED_PROPERTIES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.extended-properties+xml"
)

CORE_PROPERTIES_PARTNAME = "docProps/core.xml"
CUSTOM_PROPERTIES_PARTNAME = "docProps/custom.xml"
EXTENDED_PROPERTIES_PARTNAME = "docProps/app.xml"

CUSTOM_PROPERTY_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

OFFICE_DOCUMENT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
CUSTOM_XML_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
)
CUSTOM_XML_PROPERTIES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
)
CUSTOM_XML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
CUSTOM_XML_DATA_CONTENT_TYPE = "application/xml"
CUSTOM_XML_PROPERTIES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
)

_STRICT_PREFIX = "http://purl.oclc.org/ooxml/"
_STRICT_NAMESPACES = {
    CUSTPROPS_NS: _STRICT_PREFIX + "officeDocument/customProperties",
    VTYPES_NS: _STRICT_PREFIX + "officeDocument/docPropsVTypes",
    EP_NS: _STRICT_PREFIX + "officeDocument/extendedProperties",
    CUSTOM_XML_NS: _STRICT_PREFIX + "officeDocument/customXml",
}


def _strict_relationship(rel_type: str) -> str:
    return (
        rel_type.replace(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
            _STRICT_PREFIX + "officeDocument/relationships/",
        )
        .replace("/custom-properties", "/customProperties")
        .replace("/extended-properties", "/extendedProperties")
    )


def _part_namespace(root: ET.Element, transitional: str, local: str) -> str:
    for namespace in (transitional, _STRICT_NAMESPACES[transitional]):
        if root.tag == f"{{{namespace}}}{local}":
            return namespace
    raise ValueError("properties_part_invalid")


def _package_strict(graph: PackageGraph) -> bool:
    document = _custom_xml_document_part(graph)
    root = parse_xml_with_preserved_qnames(graph.read_part_bytes(document))
    namespace = _word_namespace(root)
    if namespace is None or root.tag != f"{{{namespace}}}document":
        raise ValueError("properties_document_part_invalid")
    return namespace == _STRICT_PREFIX + "wordprocessingml/main"


_CUSTOM_XML_ITEM_NAME = re.compile(r"customXml/item([1-9][0-9]*)\.xml\Z")
_CUSTOM_XML_STORE_ITEM_ID = re.compile(
    r"\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
    r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?\Z"
)

INT32_MIN = -2147483648
INT32_MAX = 2147483647
MAX_PROPERTY_NAME_LENGTH = 255
_MAX_TEXT_CHARACTERS = 32768
_MAX_VECTOR_ITEMS = 4096
_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)
_MEASUREMENT_UNITS = frozenset({"mm", "cm", "inches", "points", "picas"})

# Legacy vector request aliases; exact scalar base types use their VT validators.
_VECTOR_WRITE_BASE_TYPES: dict[str, str] = {
    "lpwstr": "text",
    "i4": "number",
    "r8": "number",
    "bool": "bool",
    "filetime": "date",
}
# Legacy vector discovery labels. Other bases use the structured typed vector model.
_VECTOR_TEXT_BASE_TYPES = frozenset({"lpwstr", "lpstr", "bstr"})
_VECTOR_INT_BASE_RANGES: dict[str, tuple[int, int]] = {
    "i1": (-(2**7), 2**7 - 1),
    "i2": (-(2**15), 2**15 - 1),
    "i4": (INT32_MIN, INT32_MAX),
    "i8": (-(2**63), 2**63 - 1),
    "ui1": (0, 2**8 - 1),
    "ui2": (0, 2**16 - 1),
    "ui4": (0, 2**32 - 1),
    "ui8": (0, 2**64 - 1),
}
_VECTOR_READ_BASE_TYPES = (
    _VECTOR_TEXT_BASE_TYPES
    | frozenset(_VECTOR_INT_BASE_RANGES)
    | frozenset({"r4", "r8", "bool", "filetime"})
)
_PUBLISHED = "published"
_REFUSED = "refused"


# Standard core property fields: (namespace, local_tag, kind)
CORE_PROPERTY_SPEC: dict[str, tuple[str, str, str]] = {
    "title": (DC_NS, "title", "text"),
    "subject": (DC_NS, "subject", "text"),
    "creator": (DC_NS, "creator", "text"),
    "keywords": (CP_NS, "keywords", "text"),
    "description": (DC_NS, "description", "text"),
    "category": (CP_NS, "category", "text"),
    "contentStatus": (CP_NS, "contentStatus", "text"),
    "lastModifiedBy": (CP_NS, "lastModifiedBy", "text"),
    "revision": (CP_NS, "revision", "revision"),
    "created": (DCTERMS_NS, "created", "date"),
    "modified": (DCTERMS_NS, "modified", "date"),
    "lastPrinted": (CP_NS, "lastPrinted", "date"),
    "contentType": (CP_NS, "contentType", "text"),
    "identifier": (DC_NS, "identifier", "text"),
    "language": (DC_NS, "language", "text"),
    "version": (CP_NS, "version", "text"),
}

# python-docx / Word UI names for core fields: alias -> canonical key
CORE_PROPERTY_ALIASES: dict[str, str] = {
    "author": "creator",
    "comments": "description",
    "last_modified_by": "lastModifiedBy",
    "content_status": "contentStatus",
    "last_printed": "lastPrinted",
    "content_type": "contentType",
}

# Standard extended property fields: (normative_xml_tag, kind)
EXTENDED_PROPERTY_SPEC: dict[str, tuple[str, str]] = {
    "template": ("Template", "string"),
    "manager": ("Manager", "string"),
    "company": ("Company", "string"),
    "application": ("Application", "string"),
    "app_version": ("AppVersion", "string"),
    "hyperlink_base": ("HyperlinkBase", "string"),
    "doc_security": ("DocSecurity", "int"),
    "total_time": ("TotalTime", "int"),
    "pages": ("Pages", "int"),
    "words": ("Words", "int"),
    "characters": ("Characters", "int"),
    "characters_with_spaces": ("CharactersWithSpaces", "int"),
    "lines": ("Lines", "int"),
    "paragraphs": ("Paragraphs", "int"),
    "scale_crop": ("ScaleCrop", "bool"),
    "links_up_to_date": ("LinksUpToDate", "bool"),
    "shared_doc": ("SharedDoc", "bool"),
    "hyperlinks_changed": ("HyperlinksChanged", "bool"),
    "presentation_format": ("PresentationFormat", "string"),
    "slides": ("Slides", "int"),
    "notes": ("Notes", "int"),
    "hidden_slides": ("HiddenSlides", "int"),
    "mm_clips": ("MMClips", "int"),
    "heading_pairs": ("HeadingPairs", "vector"),
    "titles_of_parts": ("TitlesOfParts", "vector"),
    "h_links": ("HLinks", "vector"),
    "dig_sig": ("DigSig", "blob"),
}

EXTENDED_PROPERTY_TAG_TO_KEY: dict[str, str] = {
    xml_tag: key for key, (xml_tag, _) in EXTENDED_PROPERTY_SPEC.items()
}


@dataclass(frozen=True)
class UntouchedVariant:
    """An unmodeled custom property value preserved verbatim."""

    tag: str
    raw_element: ET.Element

    def __repr__(self) -> str:
        return f"<UntouchedVariant tag={self.tag!r}>"


@dataclass
class CustomPropertyRecord:
    """Internal model of one custom property inside custom.xml.

    ``raw_element`` retains the original value XML. Edits operate on the source
    property tree, preserving every unedited property and its attributes.
    """

    name: str
    value: Any
    pid: int
    fmtid: str = CUSTOM_PROPERTY_FMTID
    val_type: str = ""
    raw_element: ET.Element | None = None
    link_target: str | None = None
    vt_type: str = ""

    def copy(self) -> CustomPropertyRecord:
        return CustomPropertyRecord(
            name=self.name,
            value=copy.deepcopy(self.value),
            pid=self.pid,
            fmtid=self.fmtid,
            val_type=self.val_type,
            raw_element=copy.deepcopy(self.raw_element),
            link_target=self.link_target,
            vt_type=self.vt_type,
        )


@dataclass(frozen=True)
class CustomPropertyEdit:
    """One requested modification to custom document properties."""

    name: str
    action: str = "set"
    value: Any = None
    type: str | None = None
    link_target: str | None = None
    fmtid: str | None = None

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"name": self.name, "action": self.action}
        if self.action != "delete":
            result["value"] = self.value
            if self.type is not None:
                result["type"] = self.type
            if self.link_target is not None:
                result["link_target"] = self.link_target
            if self.fmtid is not None:
                result["fmtid"] = self.fmtid
        return result


@dataclass(frozen=True)
class PropertiesRequest:
    """Source-bound property and settings edits; appended fields default to unchanged.

    ``variables`` maps names to strings (upsert) or ``None`` (remove), matching
    existing names with ``str.lower`` and retaining their original spelling.
    ``measurement_unit`` accepts ``mm``, ``cm``, ``inches``, ``points``, or
    ``picas`` but authoring refuses non-``None`` values: the UI preference belongs
    to Word's application-level Options, not this document. ``None`` leaves it
    unchanged. Field refresh is separate.
    """

    source: Path
    source_sha256: str
    destination: Path
    core: dict[str, Any] | None = None
    custom: tuple[CustomPropertyEdit, ...] = ()
    extended: dict[str, Any] | None = None
    variables: dict[str, str | None] | None = None
    measurement_unit: str | None = None


@dataclass(frozen=True)
class PropertiesReport:
    """Publication metadata and variables; application-level measurement units are unknown."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    core: dict[str, Any]
    custom: list[dict[str, Any]]
    changed_targets: list[dict[str, Any]]
    extended: dict[str, Any]
    variables: dict[str, str] = field(default_factory=dict)
    measurement_unit: str | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "core": self.core,
            "custom": self.custom,
            "extended": self.extended,
            "variables": self.variables,
            "measurement_unit": self.measurement_unit,
            "changed_targets": self.changed_targets,
        }


def _properties_refusal(request: PropertiesRequest, code: str) -> PropertiesReport:
    return PropertiesReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        core={},
        custom=[],
        changed_targets=[],
        extended={},
    )


def _naive_utc_datetime(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value.replace(microsecond=0)
    return value.astimezone(UTC).replace(tzinfo=None, microsecond=0)


def _format_datetime(value: datetime | date | str) -> str:
    if isinstance(value, datetime):
        return f"{_naive_utc_datetime(value).isoformat()}Z"
    if isinstance(value, date):
        dt = datetime.combine(value, datetime.min.time())
        return f"{_naive_utc_datetime(dt).isoformat()}Z"
    if isinstance(value, str):
        text = value.strip()
        iso_candidate = text[:-1] + "+00:00" if text.endswith("Z") else text
        try:
            parsed = datetime.fromisoformat(iso_candidate)
        except ValueError as err:
            raise ValueError("properties_core_value_invalid") from err
        return f"{_naive_utc_datetime(parsed).isoformat()}Z"
    raise ValueError("properties_core_value_invalid")


def _validate_core_text(text: Any) -> str:
    if not isinstance(text, str):
        raise TypeError("properties_core_value_invalid")
    if len(text) > _MAX_TEXT_CHARACTERS:
        raise ValueError("properties_core_value_invalid")
    if any(ord(char) < 32 and char not in ("\t", "\n", "\r") for char in text):
        raise ValueError("properties_core_value_invalid")
    return text


def _validate_core_property(name: str, value: Any) -> str | None:
    name = CORE_PROPERTY_ALIASES.get(name, name)
    if name not in CORE_PROPERTY_SPEC:
        raise ValueError("properties_core_property_unsupported")
    if value is None:
        return None
    _, _, kind = CORE_PROPERTY_SPEC[name]
    if kind == "text":
        return _validate_core_text(value)
    if kind == "revision":
        # bool subclasses int; True/False would serialize as "True"/"False".
        if isinstance(value, bool):
            raise ValueError("properties_core_value_invalid")
        if isinstance(value, int):
            if value < 0:
                raise ValueError("properties_core_value_invalid")
            return str(value)
        if isinstance(value, str):
            rev_str = value.strip()
            if not rev_str.isdigit():
                raise ValueError("properties_core_value_invalid")
            return rev_str
        raise ValueError("properties_core_value_invalid")
    if kind == "date":
        return _format_datetime(value)
    raise ValueError("properties_core_property_unsupported")


def _parse_core_dict(core_raw: Mapping[Any, Any]) -> dict[str, str | None]:
    """Validate core updates keyed by canonical name; aliases must agree with canonicals."""
    core_parsed: dict[str, str | None] = {}
    for key, raw in core_raw.items():
        if not isinstance(key, str):
            raise TypeError("properties_core_property_unsupported")
        value = _validate_core_property(key, raw)
        name = CORE_PROPERTY_ALIASES.get(key, key)
        if name in core_parsed and core_parsed[name] != value:
            raise ValueError("properties_core_alias_conflict")
        core_parsed[name] = value
    return core_parsed


def _validate_extended_text(text: Any) -> str:
    if not isinstance(text, str):
        raise TypeError("properties_extended_value_invalid")
    if len(text) > _MAX_TEXT_CHARACTERS:
        raise ValueError("properties_extended_value_invalid")
    if any(ord(char) < 32 and char not in ("\t", "\n", "\r") for char in text):
        raise ValueError("properties_extended_value_invalid")
    return text


def _validate_extended_property(name: str, value: Any) -> Any:
    if not isinstance(name, str) or name not in EXTENDED_PROPERTY_SPEC:
        raise ValueError("properties_extended_property_unsupported")
    if value is None:
        return None
    _, kind = EXTENDED_PROPERTY_SPEC[name]
    if kind in ("vector", "blob"):
        try:
            return _variant_value(_variant_element(kind, value, VTYPES_NS))
        except (ValueError, TypeError) as exc:
            raise ValueError("properties_extended_value_invalid") from exc
    if kind == "string":
        try:
            return _validate_extended_text(value)
        except TypeError as err:
            raise ValueError("properties_extended_value_invalid") from err
    if kind == "int":
        if isinstance(value, bool) or not isinstance(value, int):
            raise ValueError("properties_extended_value_invalid")
        if value < 0 or value > INT32_MAX:
            raise ValueError("properties_extended_value_invalid")
        return value
    if kind == "bool":
        if not isinstance(value, bool):
            raise ValueError("properties_extended_value_invalid")
        return value
    raise ValueError("properties_extended_property_unsupported")


def _validate_custom_name(name: Any) -> str:
    if not isinstance(name, str):
        raise TypeError("properties_custom_name_invalid")
    if not name:
        raise ValueError("properties_custom_name_invalid")
    if len(name) > MAX_PROPERTY_NAME_LENGTH:
        raise ValueError("properties_custom_name_invalid")
    if any(ord(char) < 32 for char in name):
        raise ValueError("properties_custom_name_invalid")
    return name


def _infer_vector_base_type(items: Sequence[Any]) -> str:
    if all(isinstance(item, bool) for item in items):
        return "bool"
    if all(isinstance(item, str) for item in items):
        return "lpwstr"
    if all(isinstance(item, date) for item in items):
        return "filetime"
    if all(isinstance(item, (int, float)) and not isinstance(item, bool) for item in items):
        if all(isinstance(item, int) and INT32_MIN <= item <= INT32_MAX for item in items):
            return "i4"
        return "r8"
    raise ValueError("properties_custom_value_invalid")


def _validate_custom_vector(value: Any, base_type: str | None) -> tuple[tuple[Any, ...], str]:
    if base_type is not None and base_type not in _VT_SCALARS:
        raise ValueError("properties_custom_type_unsupported")
    items = tuple(value) if isinstance(value, (list, tuple)) else None
    if items is None:
        raise ValueError("properties_custom_value_invalid")
    if not items or len(items) > _MAX_VECTOR_ITEMS:
        raise ValueError("properties_custom_vector_size_invalid")
    if base_type is None:
        base_type = _infer_vector_base_type(items)
    scalar_type = _VECTOR_WRITE_BASE_TYPES.get(base_type, base_type)
    if scalar_type == "date":
        scalar_type = "vt_date"
    normalized: list[Any] = []
    for item in items:
        norm_item, item_type = _validate_custom_value(item, scalar_type)
        if base_type == "i4" and item_type != "i4":
            raise ValueError("properties_custom_value_invalid")
        normalized.append(float(norm_item) if base_type == "r8" else norm_item)
    return tuple(normalized), f"vector_{base_type}"


def _validate_custom_value(value: Any, declared_type: str | None = None) -> tuple[Any, str]:
    if declared_type is not None:
        if not isinstance(declared_type, str):
            raise TypeError("properties_custom_type_invalid")
        norm_type = declared_type.strip()
        if norm_type == "vector" and isinstance(value, Mapping):
            return _variant_value(_variant_element("vector", value, VTYPES_NS)), "vector"
        if norm_type == "vt_date":
            return _variant_value(_variant_element("date", value, VTYPES_NS)), "vt_date"
        if norm_type in _VT_SCALARS | {"array", "variant"} and norm_type not in {"bool", "date"}:
            if norm_type == "filetime":
                return _validate_custom_value(value, "date")
            normalized = _variant_value(_variant_element(norm_type, value, VTYPES_NS))
            return normalized, norm_type
        if norm_type == "vector":
            return _validate_custom_vector(value, None)
        if norm_type.startswith("vector_"):
            return _validate_custom_vector(value, norm_type.removeprefix("vector_"))
        if norm_type == "text":
            if not isinstance(value, str):
                raise ValueError("properties_custom_value_invalid")
            if len(value) > _MAX_TEXT_CHARACTERS:
                raise ValueError("properties_custom_value_invalid")
            if any(ord(char) < 32 and char not in ("\t", "\n", "\r") for char in value):
                raise ValueError("properties_custom_value_invalid")
            return value, "lpwstr"
        if norm_type == "number":
            if isinstance(value, bool):
                raise ValueError("properties_custom_value_invalid")
            if isinstance(value, int):
                if INT32_MIN <= value <= INT32_MAX:
                    return int(value), "i4"
                return float(value), "r8"
            if isinstance(value, float):
                if not math.isfinite(value):
                    raise ValueError("properties_custom_value_invalid")
                return float(value), "r8"
            if isinstance(value, str):
                s = value.strip()
                try:
                    iv = int(s)
                    if INT32_MIN <= iv <= INT32_MAX:
                        return iv, "i4"
                    return float(iv), "r8"
                except ValueError:
                    try:
                        fv = float(s)
                        if not math.isfinite(fv):
                            raise ValueError("properties_custom_value_invalid")
                        return fv, "r8"
                    except ValueError as err:
                        raise ValueError("properties_custom_value_invalid") from err
            raise ValueError("properties_custom_value_invalid")
        if norm_type == "bool":
            if isinstance(value, bool):
                return bool(value), "bool"
            if isinstance(value, str):
                s = value.strip().lower()
                if s in ("true", "1"):
                    return True, "bool"
                if s in ("false", "0"):
                    return False, "bool"
            raise ValueError("properties_custom_value_invalid")
        if norm_type == "date":
            if isinstance(value, datetime):
                return _naive_utc_datetime(value), "filetime"
            if isinstance(value, date):
                dt = datetime.combine(value, datetime.min.time())
                return _naive_utc_datetime(dt), "filetime"
            if isinstance(value, str):
                s = value.strip()
                iso_cand = s[:-1] + "+00:00" if s.endswith("Z") else s
                try:
                    parsed = datetime.fromisoformat(iso_cand)
                    return _naive_utc_datetime(parsed), "filetime"
                except ValueError as err:
                    raise ValueError("properties_custom_value_invalid") from err
            raise ValueError("properties_custom_value_invalid")
        raise ValueError("properties_custom_type_unsupported")

    # Inferred type
    if isinstance(value, bool):
        return bool(value), "bool"
    if isinstance(value, int):
        if INT32_MIN <= value <= INT32_MAX:
            return int(value), "i4"
        return float(value), "r8"
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("properties_custom_value_invalid")
        return float(value), "r8"
    if isinstance(value, (datetime, date)):
        if isinstance(value, datetime):
            return _naive_utc_datetime(value), "filetime"
        dt = datetime.combine(value, datetime.min.time())
        return _naive_utc_datetime(dt), "filetime"
    if isinstance(value, str):
        if len(value) > _MAX_TEXT_CHARACTERS:
            raise ValueError("properties_custom_value_invalid")
        if any(ord(char) < 32 and char not in ("\t", "\n", "\r") for char in value):
            raise ValueError("properties_custom_value_invalid")
        return value, "lpwstr"
    if isinstance(value, (list, tuple)):
        return _validate_custom_vector(value, None)
    raise ValueError("properties_custom_type_unsupported")


def _parse_custom_property_edit(item: Any) -> CustomPropertyEdit:
    if not isinstance(item, Mapping) or not item:
        raise ValueError("properties_custom_edit_invalid")
    name = _validate_custom_name(item.get("name"))
    action = item.get("action", "set")
    value = item.get("value")
    declared_type = item.get("type")

    if action == "delete":
        if set(item) - {"name", "action"}:
            raise ValueError("properties_custom_edit_fields_invalid")
        return CustomPropertyEdit(name=name, action="delete", value=None, type=None)

    if action == "set":
        allowed_fields = {"name", "action", "value", "type", "link_target", "fmtid"}
        if not set(item).issubset(allowed_fields):
            raise ValueError("properties_custom_edit_fields_invalid")
        norm_val, val_type = _validate_custom_value(value, declared_type)
        link_target = item.get("link_target")
        if link_target is not None:
            link_target = _validate_custom_name(link_target) if link_target else ""
        fmtid = item.get("fmtid")
        if fmtid is not None:
            _variant_element("clsid", fmtid, VTYPES_NS)
        type_label = (
            "text"
            if val_type == "lpwstr"
            else "date"
            if val_type == "filetime"
            else "number"
            if val_type in ("i4", "r8")
            else val_type
        )
        return CustomPropertyEdit(
            name=name,
            action="set",
            value=norm_val,
            type=type_label,
            link_target=link_target,
            fmtid=fmtid,
        )

    raise ValueError("properties_custom_action_invalid")


def _validate_variable_text(value: Any, *, name: bool = False) -> str:
    """Validate bounded variable text and the complete XML 1.0 character range."""
    code = "properties_variable_name_invalid" if name else "properties_variable_value_invalid"
    if not isinstance(value, str):
        raise TypeError(code)
    limit = MAX_PROPERTY_NAME_LENGTH if name else _MAX_TEXT_CHARACTERS
    if len(value) > limit or (name and not value):
        raise ValueError(code)
    for char in value:
        point = ord(char)
        if name and point < 32:
            raise ValueError(code)
        if not (
            point in (9, 10, 13)
            or 0x20 <= point <= 0xD7FF
            or 0xE000 <= point <= 0xFFFD
            or 0x10000 <= point <= 0x10FFFF
        ):
            raise ValueError(code)
    return value


def _parse_variables(value: Any) -> dict[str, str | None] | None:
    if value is None:
        return None
    if not isinstance(value, Mapping):
        raise TypeError("properties_variables_invalid")
    parsed: dict[str, str | None] = {}
    names: set[str] = set()
    for name, text in value.items():
        name = _validate_variable_text(name, name=True)
        if name.lower() in names:
            raise ValueError("properties_variable_name_duplicate")
        names.add(name.lower())
        parsed[name] = None if text is None else _validate_variable_text(text)
    return parsed


def _validate_measurement_unit(value: Any) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError("properties_measurement_unit_invalid")
    if value not in _MEASUREMENT_UNITS:
        raise ValueError("properties_measurement_unit_invalid")
    return value


def parse_properties_request(value: Any) -> PropertiesRequest:
    """Parse and validate a structured document properties authoring request."""
    if not isinstance(value, Mapping):
        raise TypeError("properties_request_invalid")

    allowed_keys = {
        "source",
        "source_sha256",
        "destination",
        "core",
        "custom",
        "extended",
        "variables",
        "measurement_unit",
    }
    if not set(value).issubset(allowed_keys):
        raise ValueError("properties_request_fields_invalid")
    if not {"source", "source_sha256", "destination"}.issubset(set(value)):
        raise ValueError("properties_request_fields_invalid")

    source_raw = value.get("source")
    dest_raw = value.get("destination")
    if not isinstance(source_raw, (str, Path)) or not str(source_raw):
        raise ValueError("properties_path_invalid")
    if not isinstance(dest_raw, (str, Path)) or not str(dest_raw):
        raise ValueError("properties_path_invalid")

    source_path = Path(source_raw)
    dest_path = Path(dest_raw)
    source_sha256 = _sha(value["source_sha256"], "properties_source_sha256")

    core_parsed: dict[str, Any] | None = None
    if "core" in value and value["core"] is not None:
        core_raw = value["core"]
        if not isinstance(core_raw, Mapping):
            raise TypeError("properties_core_invalid")
        core_parsed = _parse_core_dict(core_raw)

    custom_parsed: list[CustomPropertyEdit] = []
    if "custom" in value and value["custom"] is not None:
        custom_raw = value["custom"]
        if not isinstance(custom_raw, list):
            raise TypeError("properties_custom_invalid")
        for item in custom_raw:
            custom_parsed.append(_parse_custom_property_edit(item))

    names = [e.name.casefold() for e in custom_parsed]
    if len(names) != len(set(names)):
        raise ValueError("properties_custom_name_duplicate")

    extended_parsed: dict[str, Any] | None = None
    if "extended" in value and value["extended"] is not None:
        extended_raw = value["extended"]
        if not isinstance(extended_raw, Mapping):
            raise TypeError("properties_extended_invalid")
        extended_parsed = {}
        for k, v in extended_raw.items():
            if not isinstance(k, str):
                raise TypeError("properties_extended_property_unsupported")
            val = _validate_extended_property(k, v)
            extended_parsed[k] = val

    return PropertiesRequest(
        source=source_path,
        source_sha256=source_sha256,
        destination=dest_path,
        core=core_parsed,
        custom=tuple(custom_parsed),
        extended=extended_parsed,
        variables=_parse_variables(value.get("variables")),
        measurement_unit=_validate_measurement_unit(value.get("measurement_unit")),
    )


def load_properties_request(path: str | Path) -> PropertiesRequest:
    """Load a document properties authoring request from a UTF-8 JSON file."""
    with Path(path).open(encoding="utf-8") as stream:
        return parse_properties_request(json.load(stream))


def _resolve_root_relationship(
    rels_root: ET.Element, rel_type: str
) -> tuple[str | None, str | None]:
    """Resolve a relationship by Type from _rels/.rels.

    Returns (rel_id, target_part_name) or (None, None).
    """
    matching: list[ET.Element] = []
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if rel.attrib.get("Type") in (rel_type, _strict_relationship(rel_type)):
            matching.append(rel)

    if not matching:
        return None, None
    if len(matching) > 1:
        raise ValueError("properties_relationship_ambiguous")

    rel = matching[0]
    rel_id = rel.attrib.get("Id")
    if not rel_id:
        raise ValueError("properties_relationship_invalid")
    if rel.attrib.get("TargetMode") == "External":
        raise ValueError("properties_relationship_external_unsupported")

    raw_target = rel.attrib.get("Target", "").strip()
    if not raw_target:
        raise ValueError("properties_relationship_target_invalid")

    clean_target = posixpath.normpath(raw_target.lstrip("/")).lstrip("/")
    if clean_target.startswith("../") or ".." in clean_target.split("/"):
        raise ValueError("properties_relationship_target_invalid")

    return rel_id, clean_target


def _verify_content_type(ct_root: ET.Element, part_name: str, expected_content_type: str) -> None:
    override_name = "/" + part_name.lstrip("/")
    for override in ct_root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        if override.attrib.get("PartName") == override_name:
            if override.attrib.get("ContentType") != expected_content_type:
                raise ValueError("properties_content_type_invalid")
            return


def _ensure_content_type_override(ct_root: ET.Element, part_name: str, content_type: str) -> None:
    override_name = "/" + part_name.lstrip("/")
    for override in ct_root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        if override.attrib.get("PartName") == override_name:
            override.attrib["ContentType"] = content_type
            return
    ET.SubElement(
        ct_root,
        f"{{{CONTENT_TYPES_NS}}}Override",
        {"PartName": override_name, "ContentType": content_type},
    )


def _new_rel_id(rels_root: ET.Element) -> str:
    used: set[str] = set()
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        if rel_id:
            used.add(rel_id)
    serial = 1
    while f"rId{serial}" in used:
        serial += 1
    return f"rId{serial}"


def _unrelated_parts_preserved(source: Path, candidate: Path, changed: frozenset[str]) -> bool:
    with (
        zipfile.ZipFile(source) as before,
        zipfile.ZipFile(candidate) as after,
    ):
        after_names = set(after.namelist())
        return all(
            info.filename in after_names and before.read(info.filename) == after.read(info.filename)
            for info in before.infolist()
            if info.filename not in changed
        )


_VT_INTS = {**_VECTOR_INT_BASE_RANGES, "int": (INT32_MIN, INT32_MAX), "uint": (0, 2**32 - 1)}
_VT_BINARY = frozenset({"blob", "oblob", "stream", "ostream", "storage", "ostorage"})
_VT_SCALARS = (
    _VECTOR_TEXT_BASE_TYPES
    | frozenset(_VT_INTS)
    | _VT_BINARY
    | frozenset(
        {
            "r4",
            "r8",
            "decimal",
            "cy",
            "bool",
            "date",
            "filetime",
            "error",
            "clsid",
            "vstream",
            "empty",
            "null",
        }
    )
)


def _variant_element(kind: str, value: Any, namespace: str, depth: int = 0) -> ET.Element:
    """Build one validated VT value. Structured variants use JSON-compatible typed values."""
    if depth > 32:
        raise ValueError("properties_custom_value_invalid")
    root = ET.Element(f"{{{namespace}}}{kind}")
    invalid = "properties_custom_value_invalid"
    if kind in ("vector", "array"):
        fields = {"base_type", "values"}
        if kind == "array":
            fields |= {"lower_bounds", "upper_bounds"}
        if not isinstance(value, Mapping) or set(value) != fields:
            raise ValueError(invalid)
        base = value["base_type"]
        if not isinstance(base, str) or base not in _VT_SCALARS | {"variant"}:
            raise ValueError("properties_custom_type_unsupported")
        items = value["values"]
        if not isinstance(items, (list, tuple)) or len(items) > _MAX_VECTOR_ITEMS:
            raise ValueError("properties_custom_vector_size_invalid")
        root.set("baseType", base)
        if kind == "vector":
            root.set("size", str(len(items)))
        else:
            lower, upper = value["lower_bounds"], value["upper_bounds"]
            if (
                not isinstance(lower, (list, tuple))
                or not isinstance(upper, (list, tuple))
                or not 1 <= len(lower) <= 32
                or len(lower) != len(upper)
                or any(
                    isinstance(i, bool) or not isinstance(i, int) or not INT32_MIN <= i <= INT32_MAX
                    for i in (*lower, *upper)
                )
                or any(hi < lo - 1 for lo, hi in zip(lower, upper, strict=True))
                or math.prod(hi - lo + 1 for lo, hi in zip(lower, upper, strict=True)) != len(items)
            ):
                raise ValueError(invalid)
            root.set("lBounds", ",".join(map(str, lower)))
            root.set("uBounds", ",".join(map(str, upper)))
        for item in items:
            root.append(_variant_element(base, item, namespace, depth + 1))
    elif kind == "variant":
        if not isinstance(value, Mapping) or set(value) != {"type", "value"}:
            raise ValueError(invalid)
        nested = value["type"]
        if not isinstance(nested, str):
            raise ValueError(invalid)
        root.append(_variant_element(nested, value["value"], namespace, depth + 1))
    elif kind in _VECTOR_TEXT_BASE_TYPES:
        if not isinstance(value, str) or len(value) > _MAX_TEXT_CHARACTERS:
            raise ValueError(invalid)
        if any(
            not (
                c in "\t\n\r"
                or 32 <= ord(c) <= 0xD7FF
                or 0xE000 <= ord(c) <= 0xFFFD
                or 0x10000 <= ord(c) <= 0x10FFFF
            )
            for c in value
        ):
            raise ValueError(invalid)
        root.text = value
    elif kind in _VT_INTS:
        low, high = _VT_INTS[kind]
        if isinstance(value, bool) or not isinstance(value, int) or not low <= value <= high:
            raise ValueError(invalid)
        root.text = str(value)
    elif kind in ("r4", "r8"):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(invalid)
        try:
            number = float(value)
        except OverflowError as exc:
            raise ValueError(invalid) from exc
        if not math.isfinite(number) or (kind == "r4" and abs(number) > 3.4028234663852886e38):
            raise ValueError(invalid)
        root.text = str(number)
    elif kind in ("decimal", "cy"):
        if not isinstance(value, (str, int, Decimal)) or isinstance(value, bool):
            raise ValueError(invalid)
        if isinstance(value, str) and len(value) > _MAX_TEXT_CHARACTERS:
            raise ValueError(invalid)
        try:
            number = Decimal(value)
        except InvalidOperation as exc:
            raise ValueError(invalid) from exc
        if (
            not number.is_finite()
            or abs(number.adjusted()) > _MAX_TEXT_CHARACTERS
            or abs(number.as_tuple().exponent) > _MAX_TEXT_CHARACTERS
        ):
            raise ValueError(invalid)
        if kind == "cy":
            numerator, denominator = number.as_integer_ratio()
            scaled, remainder = divmod(numerator * 10000, denominator)
            if remainder or not -(2**63) <= scaled <= 2**63 - 1:
                raise ValueError(invalid)
        root.text = format(number, "f")
        if len(root.text) > _MAX_TEXT_CHARACTERS:
            raise ValueError(invalid)
    elif kind == "bool":
        if not isinstance(value, bool):
            raise ValueError(invalid)
        root.text = "true" if value else "false"
    elif kind in ("date", "filetime"):
        try:
            root.text = _format_datetime(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(invalid) from exc
    elif kind in _VT_BINARY or kind == "vstream":
        if kind == "vstream":
            if not isinstance(value, Mapping) or set(value) != {"value", "version"}:
                raise ValueError(invalid)
            version = value["version"]
            if version is not None:
                _variant_element("clsid", version, namespace)
                root.set("version", version)
            value = value["value"]
        if not isinstance(value, str) or len(value) > _MAX_TEXT_CHARACTERS:
            raise ValueError(invalid)
        try:
            root.text = base64.b64encode(base64.b64decode(value, validate=True)).decode("ascii")
        except (ValueError, binascii.Error) as exc:
            raise ValueError(invalid) from exc
    elif kind == "clsid":
        if not isinstance(value, str) or not re.fullmatch(
            r"\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}", value
        ):
            raise ValueError(invalid)
        root.text = value
    elif kind == "error":
        if not isinstance(value, str) or not re.fullmatch(r"0x[0-9A-Fa-f]{8}", value):
            raise ValueError(invalid)
        root.text = value
    elif kind in ("empty", "null"):
        if value is not None:
            raise ValueError(invalid)
    else:
        raise ValueError("properties_custom_type_unsupported")
    return root


def _variant_value(element: ET.Element, depth: int = 0) -> Any:
    """Read and validate VT values without float conversion of decimal or currency."""
    if depth > 32:
        raise ValueError("properties_part_invalid")
    namespace = element.tag.split("}", 1)[0].lstrip("{")
    if namespace not in (VTYPES_NS, _STRICT_NAMESPACES[VTYPES_NS]):
        raise ValueError("properties_part_invalid")
    kind = element.tag.split("}", 1)[1]
    text = element.text or ""
    try:
        if kind in ("array", "vector"):
            base = element.get("baseType")
            items = list(element)
            if len(items) > _MAX_VECTOR_ITEMS:
                raise ValueError("properties_part_invalid")
            if any(item.tag != f"{{{namespace}}}{base}" for item in items):
                raise ValueError("properties_part_invalid")
            value = {"base_type": base, "values": [_variant_value(i, depth + 1) for i in items]}
            if kind == "array":
                value["lower_bounds"] = [int(i) for i in element.get("lBounds", "").split(",")]
                value["upper_bounds"] = [int(i) for i in element.get("uBounds", "").split(",")]
            elif int(element.get("size", "")) != len(items):
                raise ValueError("properties_part_invalid")
        elif kind == "variant":
            if len(element) != 1:
                raise ValueError("properties_part_invalid")
            value = {
                "type": element[0].tag.split("}", 1)[-1],
                "value": _variant_value(element[0], depth + 1),
            }
        else:
            if len(element):
                raise ValueError("properties_part_invalid")
            if kind in _VT_INTS:
                value = int(text.strip())
            elif kind in ("r4", "r8"):
                value = float(text.strip())
            elif kind == "bool":
                if text.strip().lower() not in ("true", "false", "1", "0"):
                    raise ValueError("properties_part_invalid")
                value = text.strip().lower() in ("true", "1")
            elif kind == "vstream":
                value = {"value": "".join(text.split()), "version": element.get("version")}
            elif kind in ("empty", "null"):
                if text.strip():
                    raise ValueError("properties_part_invalid")
                value = None
            elif kind in _VT_BINARY:
                value = "".join(text.split())
            else:
                value = text
        canonical = _variant_element(kind, value, namespace, depth)
        if kind in ("cy", "decimal", "date", "filetime"):
            return canonical.text
        return value
    except (TypeError, ValueError, OverflowError, InvalidOperation) as exc:
        raise ValueError("properties_part_invalid") from exc


def _vt_text(vt_type: str, value: Any) -> str:
    if vt_type == "bool":
        return "true" if value else "false"
    if vt_type == "filetime":
        return f"{_naive_utc_datetime(value).isoformat()}Z"
    return str(value)


def _read_custom_properties_records(xml_bytes: bytes) -> dict[str, CustomPropertyRecord]:
    try:
        root = parse_xml_with_preserved_qnames(xml_bytes)
    except ET.ParseError as exc:
        raise ValueError("properties_part_malformed") from exc
    namespace = _part_namespace(root, CUSTPROPS_NS, "Properties")
    records: dict[str, CustomPropertyRecord] = {}
    seen_names: set[str] = set()
    seen_pids: set[int] = set()
    for child in root:
        if child.tag != f"{{{namespace}}}property" or len(child) != 1:
            raise ValueError("properties_part_invalid")
        name = child.get("name")
        if not name:
            raise ValueError("properties_part_invalid")
        _validate_custom_name(name)
        if name.casefold() in seen_names:
            raise ValueError("properties_custom_name_duplicate")
        seen_names.add(name.casefold())
        try:
            pid = int(child.get("pid", ""))
        except ValueError as exc:
            raise ValueError("properties_part_invalid") from exc
        if not 2 <= pid <= INT32_MAX or pid in seen_pids:
            raise ValueError("properties_part_invalid")
        seen_pids.add(pid)
        fmtid = child.get("fmtid", CUSTOM_PROPERTY_FMTID)
        _variant_element("clsid", fmtid, VTYPES_NS)
        element = child[0]
        vt_type = element.tag.split("}", 1)[-1]
        if vt_type in _VT_SCALARS | {"array", "vector", "variant"}:
            value = _variant_value(element)
            val_type = vt_type
        else:
            value = UntouchedVariant(tag=element.tag, raw_element=copy.deepcopy(element))
            val_type = "variant"
        if vt_type in ("i1", "i2", "int"):
            val_type = "i4"
        elif vt_type == "r4":
            val_type = "r8"
        elif vt_type == "date":
            val_type = "vt_date"
        elif vt_type == "filetime":
            value, _ = _validate_custom_value(value, "date")
        elif vt_type == "vector" and value["base_type"] in _VECTOR_READ_BASE_TYPES:
            val_type = "vector_" + value["base_type"]
            value = tuple(value["values"])
            if val_type == "vector_filetime":
                value = tuple(_validate_custom_value(v, "date")[0] for v in value)
        records[name] = CustomPropertyRecord(
            name=name,
            value=value,
            pid=pid,
            fmtid=fmtid,
            val_type=val_type,
            raw_element=copy.deepcopy(element),
            link_target=child.get("linkTarget"),
            vt_type=vt_type,
        )
    return records


def _apply_core_properties(
    existing_xml: bytes | None,
    updates: Mapping[str, Any],
) -> bytes:
    if existing_xml is not None:
        try:
            root = parse_xml_with_preserved_qnames(existing_xml)
        except ET.ParseError as exc:
            raise ValueError("properties_part_malformed") from exc
        if root.tag != f"{{{CP_NS}}}coreProperties":
            raise ValueError("properties_part_invalid")
    else:
        root = ET.Element(
            f"{{{CP_NS}}}coreProperties",
            {
                f"{{{XSI_NS}}}schemaLocation": f"{CP_NS} http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            },
        )

    for key, value in updates.items():
        key = CORE_PROPERTY_ALIASES.get(key, key)
        if key not in CORE_PROPERTY_SPEC:
            raise ValueError("properties_core_property_unsupported")
        ns, local_name, kind = CORE_PROPERTY_SPEC[key]
        # Directly constructed requests bypass parse_properties_request; validate here
        # so author_properties refuses instead of serializing an invalid value.
        try:
            value = _validate_core_property(key, value)
        except TypeError as exc:
            raise ValueError(str(exc)) from exc
        tag = f"{{{ns}}}{local_name}"
        matching = root.findall(tag)

        if value is None:
            for elem in matching:
                root.remove(elem)
        else:
            if matching:
                elem = matching[0]
                elem.text = str(value)
                if kind == "date" and key in ("created", "modified"):
                    elem.set(f"{{{XSI_NS}}}type", ET.QName(DCTERMS_NS, "W3CDTF"))
            else:
                elem = ET.SubElement(root, tag)
                if kind == "date" and key in ("created", "modified"):
                    elem.set(f"{{{XSI_NS}}}type", ET.QName(DCTERMS_NS, "W3CDTF"))
                elem.text = str(value)

    return serialize_xml_with_preserved_qnames(root)


def _apply_custom_properties(
    existing_xml: bytes | None,
    edits: Sequence[CustomPropertyEdit],
    namespace: str = CUSTPROPS_NS,
) -> tuple[bytes | None, list[tuple[CustomPropertyEdit, int]]]:
    records = _read_custom_properties_records(existing_xml) if existing_xml is not None else {}
    root = (
        parse_xml_with_preserved_qnames(existing_xml)
        if existing_xml is not None
        else ET.Element(f"{{{namespace}}}Properties")
    )
    namespace = _part_namespace(root, CUSTPROPS_NS, "Properties")
    vt_namespace = _STRICT_NAMESPACES[VTYPES_NS] if namespace != CUSTPROPS_NS else VTYPES_NS
    nodes = {node.get("name").casefold(): node for node in root}
    names = [_validate_custom_name(edit.name).casefold() for edit in edits]
    if len(names) != len(set(names)):
        raise ValueError("properties_custom_name_duplicate")
    next_pid = max((r.pid for r in records.values()), default=1) + 1
    changed_targets = []
    for edit in edits:
        node = nodes.get(edit.name.casefold())
        if edit.action == "delete":
            pid = int(node.get("pid")) if node is not None else 0
            if node is not None:
                root.remove(node)
            changed_targets.append((edit, pid))
            continue
        if edit.action != "set":
            raise ValueError("properties_custom_action_invalid")
        value, kind = _validate_custom_value(edit.value, edit.type)
        if kind.startswith("vector_"):
            value = {"base_type": kind.removeprefix("vector_"), "values": list(value)}
            kind = "vector"
        element = _variant_element("date" if kind == "vt_date" else kind, value, vt_namespace)
        if node is None:
            if next_pid > INT32_MAX:
                raise ValueError("properties_custom_value_invalid")
            node = ET.SubElement(
                root,
                f"{{{namespace}}}property",
                {
                    "name": edit.name,
                    "pid": str(next_pid),
                    "fmtid": CUSTOM_PROPERTY_FMTID,
                },
            )
            next_pid += 1
        for child in list(node):
            node.remove(child)
        node.append(element)
        if edit.fmtid is not None:
            _variant_element("clsid", edit.fmtid, vt_namespace)
            node.set("fmtid", edit.fmtid)
        if edit.link_target is not None:
            if edit.link_target:
                node.set("linkTarget", _validate_custom_name(edit.link_target))
            else:
                node.attrib.pop("linkTarget", None)
        changed_targets.append((edit, int(node.get("pid"))))
    if existing_xml is None and not len(root):
        return None, changed_targets
    return serialize_xml_with_preserved_qnames(root), changed_targets


def _apply_extended_properties(
    existing_xml: bytes | None,
    updates: Mapping[str, Any],
    namespace: str = EP_NS,
) -> bytes | None:
    if not isinstance(updates, Mapping):
        raise TypeError("properties_extended_invalid")

    # Validate all extended values before applying any
    validated_updates: dict[str, tuple[Any, str, str]] = {}
    for key, value in updates.items():
        if not isinstance(key, str) or key not in EXTENDED_PROPERTY_SPEC:
            raise ValueError("properties_extended_property_unsupported")
        validated_value = _validate_extended_property(key, value)
        xml_tag, kind = EXTENDED_PROPERTY_SPEC[key]
        validated_updates[key] = (validated_value, xml_tag, kind)

    if existing_xml is not None:
        try:
            root = parse_xml_with_preserved_qnames(existing_xml)
        except ET.ParseError as exc:
            raise ValueError("properties_part_malformed") from exc
        namespace = _part_namespace(root, EP_NS, "Properties")
    else:
        if not any(val is not None for val, _, _ in validated_updates.values()):
            return None
        root = ET.Element(f"{{{namespace}}}Properties")

    # Check for ambiguous duplicates
    for _, xml_tag, _ in validated_updates.values():
        tag = f"{{{namespace}}}{xml_tag}"
        matching = root.findall(tag)
        if len(matching) > 1:
            raise ValueError("properties_part_invalid")

    # Apply all validated updates
    for key, (validated_value, xml_tag, kind) in validated_updates.items():
        tag = f"{{{namespace}}}{xml_tag}"
        matching = root.findall(tag)

        if validated_value is None:
            for elem in matching:
                root.remove(elem)
        else:
            if kind in ("vector", "blob"):
                node = matching[0] if matching else ET.SubElement(root, tag)
                for child in list(node):
                    node.remove(child)
                node.text = None
                vt_namespace = VTYPES_NS if namespace == EP_NS else _STRICT_NAMESPACES[VTYPES_NS]
                node.append(_variant_element(kind, validated_value, vt_namespace))
                continue
            if kind in ("string", "int"):
                text_val = str(validated_value)
            elif kind == "bool":
                text_val = "true" if validated_value else "false"
            else:
                raise ValueError("properties_extended_property_unsupported")

            if matching:
                matching[0].text = text_val
            else:
                elem = ET.SubElement(root, tag)
                elem.text = text_val

    serialized = serialize_xml_with_preserved_qnames(root)
    if existing_xml is None:
        root_start = serialized.find(b"<", serialized.find(b"?>") + 2)
        prefix_match = re.match(rb"<([A-Za-z_][A-Za-z0-9_.-]*):Properties", serialized[root_start:])
        if prefix_match is not None:
            prefix = prefix_match.group(1)
            serialized = re.sub(rb"(</?)" + re.escape(prefix) + rb":", rb"\1", serialized)
            declaration = b" xmlns:" + prefix + b'="' + namespace.encode() + b'"'
            if declaration in serialized:
                serialized = serialized.replace(
                    declaration, b' xmlns="' + namespace.encode() + b'"', 1
                )
    return serialized


def _read_extended_properties(xml_bytes: bytes) -> dict[str, Any]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise ValueError("properties_part_malformed") from exc
    namespace = _part_namespace(root, EP_NS, "Properties")
    values = {}
    for name, (tag, kind) in EXTENDED_PROPERTY_SPEC.items():
        nodes = root.findall(f"{{{namespace}}}{tag}")
        if len(nodes) > 1:
            raise ValueError("properties_part_invalid")
        if not nodes:
            values[name] = None
            continue
        node = nodes[0]
        try:
            if kind in ("vector", "blob"):
                vt_namespace = VTYPES_NS if namespace == EP_NS else _STRICT_NAMESPACES[VTYPES_NS]
                if len(node) != 1 or node[0].tag != f"{{{vt_namespace}}}{kind}":
                    raise ValueError("properties_part_invalid")
                values[name] = _variant_value(node[0])
            else:
                if len(node):
                    raise ValueError("properties_part_invalid")
                text = node.text or ""
                value = int(text.strip()) if kind == "int" else text
                if kind == "bool":
                    if text.strip().lower() not in ("true", "false", "0", "1"):
                        raise ValueError("properties_part_invalid")
                    value = text.strip().lower() in ("true", "1")
                values[name] = _validate_extended_property(name, value)
        except (ValueError, TypeError) as exc:
            raise ValueError("properties_part_invalid") from exc
    return values


def _load_property_settings(
    graph: PackageGraph,
) -> tuple[str, str, str, str | None, ET.Element | None]:
    """Resolve settings through the main document, including Strict OOXML namespaces."""
    documents = [rel for rel in graph.relationships_for("") if rel.type.endswith("/officeDocument")]
    if len(documents) != 1 or documents[0].is_external:
        raise ValueError("properties_document_relationship_invalid")
    document = documents[0].resolved_target()
    try:
        document_root = parse_xml_with_preserved_qnames(graph.read_part_bytes(document))
    except (KeyError, PackageError, ET.ParseError):
        raise ValueError("properties_document_part_invalid") from None
    namespace = _word_namespace(document_root)
    if namespace is None or document_root.tag != f"{{{namespace}}}document":
        raise ValueError("properties_document_part_invalid")
    rel_type = documents[0].type.rsplit("/", 1)[0] + "/settings"
    relationships = [
        rel for rel in graph.relationships_for(document) if rel.type.endswith("/settings")
    ]
    if len(relationships) > 1:
        raise ValueError("properties_settings_relationship_ambiguous")
    if not relationships:
        return document, namespace, rel_type, None, None
    relationship = relationships[0]
    if relationship.is_external:
        raise ValueError("properties_settings_relationship_external_unsupported")
    target = relationship.resolved_target()
    part = graph.part(target)
    if part is None:
        raise ValueError("properties_settings_part_missing")
    if part.content_type != _SETTINGS_CONTENT_TYPE:
        raise ValueError("properties_settings_content_type_invalid")
    try:
        root = parse_xml_with_preserved_qnames(graph.read_part_bytes(target))
    except (PackageError, ET.ParseError):
        raise ValueError("properties_settings_part_malformed") from None
    if root.tag != f"{{{namespace}}}settings":
        raise ValueError("properties_settings_part_invalid")
    return document, namespace, rel_type, target, root


def _read_document_variables(root: ET.Element | None, namespace: str) -> dict[str, str]:
    """List original variable spelling; reject ambiguous lowercased identities."""
    if root is None:
        return {}
    containers = root.findall(f"{{{namespace}}}docVars")
    if len(containers) > 1:
        raise ValueError("properties_docvars_ambiguous")
    variables: dict[str, str] = {}
    seen: set[str] = set()
    if containers:
        for node in containers[0].findall(f"{{{namespace}}}docVar"):
            name = _validate_variable_text(node.get(f"{{{namespace}}}name"), name=True)
            value = _validate_variable_text(node.get(f"{{{namespace}}}val"))
            if name.lower() in seen:
                raise ValueError("properties_variable_name_duplicate")
            seen.add(name.lower())
            variables[name] = value
    return variables


def _document_variable_values(graph: PackageGraph) -> dict[str, str]:
    """Return validated lowercase variable identities for DOCVARIABLE evaluation.

    Resolves the document-owned settings part and rejects malformed or ambiguous
    topology and variable names instead of silently selecting a value. Consumers
    should translate ``ValueError``, ``TypeError``, and ``PackageError`` into their
    own typed refusal surface.
    """
    _, namespace, _, _, root = _load_property_settings(graph)
    return {
        name.lower(): value for name, value in _read_document_variables(root, namespace).items()
    }


def _plan_variable_properties(
    graph: PackageGraph,
    parts: Mapping[str, bytes],
    ct_root: ET.Element,
    updates: dict[str, str | None] | None,
) -> tuple[dict[str, bytes], dict[str, bytes], list[dict[str, Any]], bool]:
    """Plan variable edits without refreshing fields or rewriting unrelated settings."""
    replacements: dict[str, bytes] = {}
    additions: dict[str, bytes] = {}
    changes: list[dict[str, Any]] = []
    if not updates:
        return replacements, additions, changes, False
    document, namespace, rel_type, target, root = _load_property_settings(graph)
    _read_document_variables(root, namespace)
    if root is None:
        root = ET.Element(f"{{{namespace}}}settings")
    container = root.find(f"{{{namespace}}}docVars")
    nodes = (
        {
            node.attrib[f"{{{namespace}}}name"].lower(): node
            for node in container.findall(f"{{{namespace}}}docVar")
        }
        if container is not None
        else {}
    )
    for name, value in updates.items():
        node = nodes.get(name.lower())
        if value is None:
            if node is None or container is None:
                continue
            container.remove(node)
            del nodes[name.lower()]
        else:
            if node is not None and node.get(f"{{{namespace}}}val") == value:
                continue
            if container is None:
                container = ET.Element(f"{{{namespace}}}docVars")
                root.append(container)
            if node is None:
                node = ET.SubElement(
                    container, f"{{{namespace}}}docVar", {f"{{{namespace}}}name": name}
                )
                nodes[name.lower()] = node
            node.set(f"{{{namespace}}}val", value)
        changes.append(
            {
                "target": "variables",
                "name": name,
                "action": "delete" if value is None else "set",
            }
        )
    if not changes:
        return replacements, additions, changes, False
    if container is not None:
        root.remove(container)
        # Reuse the CT_Settings sequence from protection, preserving other children.
        if len(container) or container.attrib or (container.text or "").strip():
            _insert_settings_child(root, container, namespace)
    serialized = serialize_xml_with_preserved_qnames(root)
    if target is not None:
        replacements[target] = serialized
        return replacements, additions, changes, False

    # New settings live beside the actual document, never clobbering an orphan part.
    target = posixpath.join(posixpath.dirname(document), "settings.xml")
    if target in parts:
        raise ValueError("properties_settings_part_collision")
    rels_name = posixpath.join(
        posixpath.dirname(document), "_rels", posixpath.basename(document) + ".rels"
    )
    try:
        rels_root = (
            parse_xml_with_preserved_qnames(parts[rels_name])
            if rels_name in parts
            else ET.Element(f"{{{PKG_REL_NS}}}Relationships")
        )
    except ET.ParseError:
        raise ValueError("properties_settings_relationship_invalid") from None
    if rels_root.tag != f"{{{PKG_REL_NS}}}Relationships":
        raise ValueError("properties_settings_relationship_invalid")
    ET.SubElement(
        rels_root,
        f"{{{PKG_REL_NS}}}Relationship",
        {"Id": _new_rel_id(rels_root), "Type": rel_type, "Target": "settings.xml"},
    )
    rels_xml = serialize_xml_with_preserved_qnames(rels_root)
    if rels_name in parts:
        replacements[rels_name] = rels_xml
    else:
        additions[rels_name] = rels_xml
    additions[target] = serialized
    _verify_content_type(ct_root, target, _SETTINGS_CONTENT_TYPE)
    _ensure_content_type_override(ct_root, target, _SETTINGS_CONTENT_TYPE)
    return replacements, additions, changes, True


def discover_property_targets(source: str | Path) -> dict[str, Any]:
    """Inspect source package document property parts, relationships, and metadata."""
    path = Path(source)
    if not path.is_file():
        return {
            "passed": False,
            "code": "properties_source_missing",
            "source_sha256": "",
            "core": {"exists": False, "part": None, "properties": {}},
            "custom": {"exists": False, "part": None, "properties": []},
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    try:
        graph = PackageGraph.from_file(path)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return {
            "passed": False,
            "code": "properties_source_package_invalid",
            "source_sha256": "",
            "core": {"exists": False, "part": None, "properties": {}},
            "custom": {"exists": False, "part": None, "properties": []},
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    if "_rels/.rels" not in parts:
        return {
            "passed": False,
            "code": "properties_source_package_invalid",
            "source_sha256": graph.file_sha256,
            "core": {"exists": False, "part": None, "properties": {}},
            "custom": {"exists": False, "part": None, "properties": []},
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    try:
        rels_root = ET.fromstring(parts["_rels/.rels"])
    except ET.ParseError:
        return {
            "passed": False,
            "code": "properties_source_package_invalid",
            "source_sha256": graph.file_sha256,
            "core": {"exists": False, "part": None, "properties": {}},
            "custom": {"exists": False, "part": None, "properties": []},
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    # Discover core properties
    try:
        _, core_part = _resolve_root_relationship(rels_root, CORE_PROPERTIES_REL_TYPE)
    except ValueError as exc:
        return {
            "passed": False,
            "code": str(exc),
            "source_sha256": graph.file_sha256,
            "core": {"exists": False, "part": None, "properties": {}},
            "custom": {"exists": False, "part": None, "properties": []},
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    core_dict: dict[str, str | None] = {}
    core_exists = core_part is not None and core_part in parts
    if core_exists and core_part is not None:
        try:
            core_root = ET.fromstring(parts[core_part])
            if core_root.tag != f"{{{CP_NS}}}coreProperties":
                return {
                    "passed": False,
                    "code": "properties_part_invalid",
                    "source_sha256": graph.file_sha256,
                    "core": {"exists": True, "part": core_part, "properties": {}},
                    "custom": {"exists": False, "part": None, "properties": []},
                    "extended": {
                        "exists": False,
                        "part": None,
                        "properties": {},
                        "read_only": False,
                    },
                }
            for prop_name, (ns, local_name, _) in CORE_PROPERTY_SPEC.items():
                node = core_root.find(f"{{{ns}}}{local_name}")
                core_dict[prop_name] = node.text if node is not None else None
        except ET.ParseError:
            return {
                "passed": False,
                "code": "properties_part_malformed",
                "source_sha256": graph.file_sha256,
                "core": {"exists": True, "part": core_part, "properties": {}},
                "custom": {"exists": False, "part": None, "properties": []},
                "extended": {
                    "exists": False,
                    "part": None,
                    "properties": {},
                    "read_only": False,
                },
            }

    # Discover custom properties
    try:
        _, custom_part = _resolve_root_relationship(rels_root, CUSTOM_PROPERTIES_REL_TYPE)
    except ValueError as exc:
        return {
            "passed": False,
            "code": str(exc),
            "source_sha256": graph.file_sha256,
            "core": {"exists": core_exists, "part": core_part, "properties": core_dict},
            "custom": {"exists": False, "part": None, "properties": []},
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    custom_props: list[dict[str, Any]] = []
    custom_exists = custom_part is not None and custom_part in parts
    if custom_exists and custom_part is not None:
        try:
            custom_records = _read_custom_properties_records(parts[custom_part])
            for rec in custom_records.values():
                type_label = (
                    "text"
                    if rec.val_type == "lpwstr"
                    else "date"
                    if rec.val_type == "filetime"
                    else "number"
                    if rec.val_type in ("i4", "r8")
                    else rec.val_type
                )
                val_repr = (
                    f"{_naive_utc_datetime(rec.value).isoformat()}Z"
                    if rec.val_type == "filetime"
                    else [_vt_text("filetime", item) for item in rec.value]
                    if rec.val_type == "vector_filetime"
                    else list(rec.value)
                    if rec.val_type.startswith("vector_")
                    else repr(rec.value)
                    if isinstance(rec.value, UntouchedVariant)
                    else rec.value
                )
                custom_props.append(
                    {
                        "name": rec.name,
                        "value": val_repr,
                        "type": type_label,
                        "pid": rec.pid,
                        "fmtid": rec.fmtid,
                        "link_target": rec.link_target,
                        "vt_type": rec.vt_type,
                    }
                )
        except ValueError as exc:
            return {
                "passed": False,
                "code": str(exc),
                "source_sha256": graph.file_sha256,
                "core": {
                    "exists": core_exists,
                    "part": core_part,
                    "properties": core_dict,
                },
                "custom": {"exists": True, "part": custom_part, "properties": []},
                "extended": {
                    "exists": False,
                    "part": None,
                    "properties": {},
                    "read_only": False,
                },
            }

    # Discover extended properties.
    try:
        _, extended_part = _resolve_root_relationship(rels_root, EXTENDED_PROPERTIES_REL_TYPE)
    except ValueError as exc:
        return {
            "passed": False,
            "code": str(exc),
            "source_sha256": graph.file_sha256,
            "core": {"exists": core_exists, "part": core_part, "properties": core_dict},
            "custom": {
                "exists": custom_exists,
                "part": custom_part,
                "properties": custom_props,
            },
            "extended": {
                "exists": False,
                "part": None,
                "properties": {},
                "read_only": False,
            },
        }

    extended_dict: dict[str, Any] = {}
    extended_exists = extended_part is not None and extended_part in parts
    if extended_exists and extended_part is not None:
        try:
            extended_dict = _read_extended_properties(parts[extended_part])
        except ValueError as exc:
            return {
                "passed": False,
                "code": str(exc),
                "source_sha256": graph.file_sha256,
                "core": {"exists": core_exists, "part": core_part, "properties": core_dict},
                "custom": {
                    "exists": custom_exists,
                    "part": custom_part,
                    "properties": custom_props,
                },
                "extended": {
                    "exists": True,
                    "part": extended_part,
                    "properties": {},
                    "read_only": False,
                },
            }
    try:
        _, namespace, _, settings_part, settings_root = _load_property_settings(graph)
        variables = _read_document_variables(settings_root, namespace)
    except (ValueError, TypeError) as exc:
        return {
            "passed": False,
            "code": str(exc),
            "source_sha256": graph.file_sha256,
            "core": {"exists": core_exists, "part": core_part, "properties": core_dict},
            "custom": {
                "exists": custom_exists,
                "part": custom_part,
                "properties": custom_props,
            },
            "extended": {
                "exists": extended_exists,
                "part": extended_part,
                "properties": extended_dict,
                "read_only": False,
            },
            "variables": {},
            "measurement_unit": None,
        }
    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        "core_exists": core_exists,
        "custom_exists": custom_exists,
        "extended_exists": extended_exists,
        "core_part": core_part,
        "custom_part": custom_part,
        "extended_part": extended_part,
        "settings_part": settings_part,
        "variables": variables,
        "measurement_unit": None,
        "core": {
            "exists": core_exists,
            "part": core_part,
            "properties": core_dict,
        },
        "custom": {
            "exists": custom_exists,
            "part": custom_part,
            "properties": custom_props,
        },
        "extended": {
            "exists": extended_exists,
            "part": extended_part,
            "properties": extended_dict,
            "read_only": False,
        },
    }


discover_properties = discover_property_targets


def author_properties(request: PropertiesRequest) -> PropertiesReport:
    """Apply metadata and case-insensitive document-variable edits without field refresh."""
    if not request.source.is_file():
        return _properties_refusal(request, "properties_source_missing")
    if request.destination.exists() or request.destination.is_symlink():
        return _properties_refusal(request, "properties_destination_invalid")

    try:
        graph = PackageGraph.from_file(request.source)
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
    except (OSError, PackageError):
        return _properties_refusal(request, "properties_source_package_invalid")

    if graph.file_sha256 != request.source_sha256:
        return _properties_refusal(request, "properties_source_stale")

    if "[Content_Types].xml" not in parts or "_rels/.rels" not in parts:
        return _properties_refusal(request, "properties_source_package_invalid")

    try:
        rels_root = parse_xml_with_preserved_qnames(parts["_rels/.rels"])
        ct_root = parse_xml_with_preserved_qnames(parts["[Content_Types].xml"])
    except (ET.ParseError, ValueError, PackageError):
        return _properties_refusal(request, "properties_source_package_invalid")

    # Existing property parts carry their own namespace. Resolve the document
    # profile only when creating a part, or after the owned diagnostics below.
    strict: bool | None = None

    try:
        _, core_part = _resolve_root_relationship(rels_root, CORE_PROPERTIES_REL_TYPE)
    except ValueError as exc:
        return _properties_refusal(request, str(exc))

    if core_part is not None:
        if core_part not in parts:
            return _properties_refusal(request, "properties_part_missing")
        try:
            _verify_content_type(ct_root, core_part, CORE_PROPERTIES_CONTENT_TYPE)
        except ValueError as exc:
            return _properties_refusal(request, str(exc))

    try:
        _, custom_part = _resolve_root_relationship(rels_root, CUSTOM_PROPERTIES_REL_TYPE)
    except ValueError as exc:
        return _properties_refusal(request, str(exc))

    if custom_part is not None:
        if custom_part not in parts:
            return _properties_refusal(request, "properties_part_missing")
        try:
            _verify_content_type(ct_root, custom_part, CUSTOM_PROPERTIES_CONTENT_TYPE)
        except ValueError as exc:
            return _properties_refusal(request, str(exc))

    try:
        _, extended_part = _resolve_root_relationship(rels_root, EXTENDED_PROPERTIES_REL_TYPE)
    except ValueError as exc:
        return _properties_refusal(request, str(exc))

    if extended_part is not None:
        if extended_part not in parts:
            return _properties_refusal(request, "properties_part_missing")
        try:
            _verify_content_type(ct_root, extended_part, EXTENDED_PROPERTIES_CONTENT_TYPE)
        except ValueError as exc:
            return _properties_refusal(request, str(exc))
    try:
        variable_updates = _parse_variables(request.variables)
        _validate_measurement_unit(request.measurement_unit)
        if request.measurement_unit is not None:
            return _properties_refusal(request, "properties_measurement_unit_unsupported")
        replacements, additions, changed_targets, settings_ct_changed = _plan_variable_properties(
            graph, parts, ct_root, variable_updates
        )
    except (ValueError, TypeError) as exc:
        return _properties_refusal(request, str(exc))
    rels_changed = False
    ct_changed = settings_ct_changed

    # Core property authoring
    if request.core is not None:
        target_core_part = core_part or CORE_PROPERTIES_PARTNAME
        existing_core_xml = parts.get(target_core_part)
        try:
            new_core_xml = _apply_core_properties(existing_core_xml, request.core)
        except ValueError as exc:
            return _properties_refusal(request, str(exc))

        if core_part is None:
            additions[target_core_part] = new_core_xml
            new_id = _new_rel_id(rels_root)
            ET.SubElement(
                rels_root,
                f"{{{PKG_REL_NS}}}Relationship",
                {
                    "Id": new_id,
                    "Type": CORE_PROPERTIES_REL_TYPE,
                    "Target": target_core_part,
                },
            )
            rels_changed = True
            _ensure_content_type_override(ct_root, target_core_part, CORE_PROPERTIES_CONTENT_TYPE)
            ct_changed = True
        else:
            replacements[target_core_part] = new_core_xml

        for k, v in request.core.items():
            changed_targets.append(
                {
                    "target": "core",
                    "name": k,
                    "action": "delete" if v is None else "set",
                }
            )

    # Custom property authoring
    if request.custom:
        target_custom_part = custom_part or CUSTOM_PROPERTIES_PARTNAME
        existing_custom_xml = parts.get(target_custom_part)
        if custom_part is None:
            try:
                strict = _package_strict(graph)
            except (ET.ParseError, ValueError, PackageError):
                return _properties_refusal(request, "properties_source_package_invalid")
        try:
            new_custom_xml, custom_changes = _apply_custom_properties(
                existing_custom_xml,
                request.custom,
                _STRICT_NAMESPACES[CUSTPROPS_NS] if strict else CUSTPROPS_NS,
            )
        except (ValueError, TypeError) as exc:
            return _properties_refusal(request, str(exc))

        for edit, pid in custom_changes:
            target_entry: dict[str, Any] = {
                "target": "custom",
                "name": edit.name,
                "action": edit.action,
            }
            if edit.action != "delete":
                target_entry["pid"] = pid
            changed_targets.append(target_entry)

        if custom_part is None:
            if new_custom_xml is not None:
                additions[target_custom_part] = new_custom_xml
                new_id = _new_rel_id(rels_root)
                ET.SubElement(
                    rels_root,
                    f"{{{PKG_REL_NS}}}Relationship",
                    {
                        "Id": new_id,
                        "Type": _strict_relationship(CUSTOM_PROPERTIES_REL_TYPE)
                        if strict
                        else CUSTOM_PROPERTIES_REL_TYPE,
                        "Target": target_custom_part,
                    },
                )
                rels_changed = True
                _ensure_content_type_override(
                    ct_root, target_custom_part, CUSTOM_PROPERTIES_CONTENT_TYPE
                )
                ct_changed = True
        else:
            if new_custom_xml is not None:
                replacements[target_custom_part] = new_custom_xml

    # Extended property authoring
    if request.extended is not None:
        target_extended_part = extended_part or EXTENDED_PROPERTIES_PARTNAME
        existing_extended_xml = parts.get(target_extended_part)
        if extended_part is None and strict is None:
            try:
                strict = _package_strict(graph)
            except (ET.ParseError, ValueError, PackageError):
                return _properties_refusal(request, "properties_source_package_invalid")
        try:
            new_extended_xml = _apply_extended_properties(
                existing_extended_xml,
                request.extended,
                _STRICT_NAMESPACES[EP_NS] if strict else EP_NS,
            )
        except (ValueError, TypeError) as exc:
            return _properties_refusal(request, str(exc))

        if extended_part is None:
            if new_extended_xml is not None:
                additions[target_extended_part] = new_extended_xml
                new_id = _new_rel_id(rels_root)
                ET.SubElement(
                    rels_root,
                    f"{{{PKG_REL_NS}}}Relationship",
                    {
                        "Id": new_id,
                        "Type": _strict_relationship(EXTENDED_PROPERTIES_REL_TYPE)
                        if strict
                        else EXTENDED_PROPERTIES_REL_TYPE,
                        "Target": target_extended_part,
                    },
                )
                rels_changed = True
                _ensure_content_type_override(
                    ct_root, target_extended_part, EXTENDED_PROPERTIES_CONTENT_TYPE
                )
                ct_changed = True
        else:
            if new_extended_xml is not None:
                replacements[target_extended_part] = new_extended_xml

        for k, v in request.extended.items():
            changed_targets.append(
                {
                    "target": "extended",
                    "name": k,
                    "action": "delete" if v is None else "set",
                }
            )
    if strict is None:
        try:
            _package_strict(graph)
        except (ET.ParseError, ValueError, PackageError):
            return _properties_refusal(request, "properties_source_package_invalid")

    if rels_changed:
        replacements["_rels/.rels"] = serialize_xml_with_preserved_qnames(rels_root)
    if ct_changed:
        replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(ct_root)

    all_touched = frozenset(set(replacements) | set(additions))

    def verifier(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(request.source, staged, all_touched):
            return "properties_staged_verification_failed"
        disc = discover_property_targets(staged)
        if not disc["passed"]:
            return "properties_staged_verification_failed"
        return None

    published: ReviewReport = _publish(
        source=request.source,
        output=request.destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        revisions=[],
        targets=changed_targets,
        additions=additions,
        verify=verifier,
        failure_code="properties_staged_verification_failed",
        exists_code="properties_destination_invalid",
        unavailable_code="properties_publication_unavailable",
    )

    if published.status == _PUBLISHED:
        final_disc = discover_property_targets(request.destination)
        final_core = final_disc["core"]["properties"] if final_disc.get("core") else {}
        final_custom = final_disc["custom"]["properties"] if final_disc.get("custom") else []
        final_extended = final_disc["extended"]["properties"] if final_disc.get("extended") else {}
        final_variables = final_disc["variables"]
    else:
        final_core = {}
        final_custom = []
        final_extended = {}
        final_variables = {}
    return PropertiesReport(
        status=published.status,
        code=published.code,
        source_sha256=published.source_sha256,
        destination_sha256=published.output_sha256,
        core=final_core,
        custom=final_custom,
        changed_targets=changed_targets,
        extended=final_extended,
        variables=final_variables,
    )


def _store_item_key(value: str) -> str:
    if not isinstance(value, str) or _CUSTOM_XML_STORE_ITEM_ID.fullmatch(value) is None:
        raise ValueError("custom_xml_store_item_id_invalid")
    return value.strip("{}").casefold()


def _custom_xml_payload(value: bytes | str | None) -> bytes:
    if isinstance(value, str):
        payload = value.encode("utf-8")
    elif isinstance(value, bytes):
        payload = value
    else:
        raise TypeError("custom_xml_payload_invalid")
    if not payload or b"<!DOCTYPE" in payload.upper() or b"<!ENTITY" in payload.upper():
        raise ValueError("custom_xml_payload_invalid")
    try:
        ET.fromstring(payload)
    except ET.ParseError as exc:
        raise ValueError("custom_xml_payload_invalid") from exc
    return payload


def _custom_xml_schema_refs(value: Sequence[str] | None) -> tuple[str, ...] | None:
    if value is None:
        return None
    if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
        raise TypeError("custom_xml_schema_refs_invalid")
    refs = tuple(value)
    if len(refs) > 256 or any(
        not isinstance(ref, str) or not ref or len(ref) > 2048 for ref in refs
    ):
        raise ValueError("custom_xml_schema_refs_invalid")
    if len(set(refs)) != len(refs):
        raise ValueError("custom_xml_schema_refs_invalid")
    return refs


@dataclass(frozen=True)
class CustomXmlStorageEdit:
    """One add, update, replacement, or removal of a document-owned Custom XML item."""

    action: str
    store_item_id: str
    payload: bytes | str | None = None
    schema_refs: Sequence[str] | None = None

    def __post_init__(self) -> None:
        if self.action not in {"add", "replace", "update", "delete"}:
            raise ValueError("custom_xml_action_invalid")
        _store_item_key(self.store_item_id)
        payload = _custom_xml_payload(self.payload) if self.payload is not None else None
        refs = _custom_xml_schema_refs(self.schema_refs)
        if self.action == "add" and payload is None:
            raise ValueError("custom_xml_payload_required")
        if self.action in {"replace", "update"} and payload is None and refs is None:
            raise ValueError("custom_xml_replacement_empty")
        if self.action == "delete" and (payload is not None or refs is not None):
            raise ValueError("custom_xml_delete_fields_invalid")
        object.__setattr__(self, "payload", payload)
        object.__setattr__(self, "schema_refs", refs)


@dataclass(frozen=True)
class CustomXmlStorageRequest:
    """Source-bound lifecycle request for document-owned Custom XML storage."""

    source: Path
    destination: Path
    source_sha256: str
    edits: tuple[CustomXmlStorageEdit, ...]

    def __post_init__(self) -> None:
        if not self.edits:
            raise ValueError("custom_xml_edits_required")
        if len({_store_item_key(edit.store_item_id) for edit in self.edits}) != len(self.edits):
            raise ValueError("custom_xml_store_item_id_duplicate")
        object.__setattr__(self, "source_sha256", _sha(self.source_sha256, "source_sha256"))


@dataclass(frozen=True)
class CustomXmlStoragePart:
    """One read-only Custom XML data item and its document-owned part topology."""

    store_item_id: str
    payload_part: str
    properties_part: str
    relationship_id: str
    schema_refs: tuple[str, ...]
    payload: bytes

    def to_dict(self, *, include_payload: bool = False) -> dict[str, Any]:
        result: dict[str, Any] = {
            "store_item_id": self.store_item_id,
            "payload_part": self.payload_part,
            "properties_part": self.properties_part,
            "relationship_id": self.relationship_id,
            "schema_refs": list(self.schema_refs),
            "payload_sha256": hashlib.sha256(self.payload).hexdigest(),
        }
        if include_payload:
            result["payload"] = self.payload
        return result


@dataclass(frozen=True)
class CustomXmlStorageReport:
    """Result of a Custom XML storage lifecycle operation."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None
    items: tuple[CustomXmlStoragePart, ...]
    changed_targets: tuple[dict[str, str], ...]

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "passed": self.passed,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "items": [item.to_dict() for item in self.items],
            "changed_targets": list(self.changed_targets),
        }


def _custom_xml_document_part(graph: PackageGraph) -> str:
    matches = [
        relationship
        for relationship in graph.relationships_for("")
        if relationship.type
        in (OFFICE_DOCUMENT_REL_TYPE, _strict_relationship(OFFICE_DOCUMENT_REL_TYPE))
        and not relationship.is_external
    ]
    if len(matches) != 1:
        raise ValueError("custom_xml_document_relationship_invalid")
    document_part = matches[0].resolved_target()
    if document_part not in graph.parts:
        raise ValueError("custom_xml_document_part_missing")
    return document_part


def _custom_xml_relationship_part(part_name: str) -> str:
    return f"{posixpath.dirname(part_name)}/_rels/{posixpath.basename(part_name)}.rels"


def _custom_xml_properties_xml(
    store_item_id: str,
    schema_refs: Sequence[str],
    namespace: str = CUSTOM_XML_NS,
) -> bytes:
    ET.register_namespace("ds", namespace)
    root = ET.Element(
        f"{{{namespace}}}datastoreItem",
        {f"{{{namespace}}}itemID": store_item_id},
    )
    if schema_refs:
        refs = ET.SubElement(root, f"{{{namespace}}}schemaRefs")
        for uri in schema_refs:
            ET.SubElement(refs, f"{{{namespace}}}schemaRef", {f"{{{namespace}}}uri": uri})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _custom_xml_item_parts(
    graph: PackageGraph,
) -> tuple[str, tuple[CustomXmlStoragePart, ...]]:
    graph.validate_relationship_graph()
    document_part = _custom_xml_document_part(graph)
    result: list[CustomXmlStoragePart] = []
    seen_ids: set[str] = set()
    for relationship in graph.relationships_for(document_part):
        if (
            relationship.type
            not in (CUSTOM_XML_REL_TYPE, _strict_relationship(CUSTOM_XML_REL_TYPE))
            or relationship.is_external
        ):
            continue
        payload_part = relationship.resolved_target()
        item_match = _CUSTOM_XML_ITEM_NAME.fullmatch(payload_part)
        if item_match is None or payload_part not in graph.parts:
            raise ValueError("custom_xml_payload_part_invalid")
        if graph.part(payload_part).content_type != CUSTOM_XML_DATA_CONTENT_TYPE:
            raise ValueError("custom_xml_payload_content_type_invalid")
        properties_matches = [
            item_relationship
            for item_relationship in graph.relationships_for(payload_part)
            if item_relationship.type
            in (
                CUSTOM_XML_PROPERTIES_REL_TYPE,
                _strict_relationship(CUSTOM_XML_PROPERTIES_REL_TYPE),
            )
            and not item_relationship.is_external
        ]
        if len(properties_matches) != 1:
            raise ValueError("custom_xml_properties_relationship_invalid")
        properties_part = properties_matches[0].resolved_target()
        expected_properties_part = f"customXml/itemProps{item_match.group(1)}.xml"
        if properties_part != expected_properties_part or properties_part not in graph.parts:
            raise ValueError("custom_xml_properties_part_invalid")
        if graph.part(properties_part).content_type != CUSTOM_XML_PROPERTIES_CONTENT_TYPE:
            raise ValueError("custom_xml_properties_content_type_invalid")
        try:
            root = ET.fromstring(graph.read_part_bytes(properties_part))
        except (ET.ParseError, PackageError) as exc:
            raise ValueError("custom_xml_properties_part_invalid") from exc
        try:
            namespace = _part_namespace(root, CUSTOM_XML_NS, "datastoreItem")
        except ValueError as exc:
            raise ValueError("custom_xml_properties_part_invalid") from exc
        store_item_id = root.get(f"{{{namespace}}}itemID")
        item_key = _store_item_key(store_item_id or "")
        if item_key in seen_ids:
            raise ValueError("custom_xml_store_item_id_duplicate")
        seen_ids.add(item_key)
        schema_refs_parent = root.find(f"{{{namespace}}}schemaRefs")
        schema_refs = (
            ()
            if schema_refs_parent is None
            else tuple(
                node.get(f"{{{namespace}}}uri") or ""
                for node in schema_refs_parent.findall(f"{{{namespace}}}schemaRef")
            )
        )
        _custom_xml_schema_refs(schema_refs)
        result.append(
            CustomXmlStoragePart(
                store_item_id=store_item_id,
                payload_part=payload_part,
                properties_part=properties_part,
                relationship_id=relationship.id,
                schema_refs=schema_refs,
                payload=graph.read_part_bytes(payload_part),
            )
        )
    return document_part, tuple(sorted(result, key=lambda item: item.payload_part))


def discover_custom_xml_storage(source: str | Path) -> dict[str, Any]:
    """Inspect Custom XML items owned by the document part, without exposing payloads."""
    path = Path(source)
    if not path.is_file():
        return {"passed": False, "code": "custom_xml_source_missing", "items": []}
    try:
        graph = PackageGraph.from_file(path)
        document_part, items = _custom_xml_item_parts(graph)
    except (OSError, PackageError, ValueError):
        return {
            "passed": False,
            "code": "custom_xml_source_package_invalid",
            "items": [],
        }
    return {
        "passed": True,
        "code": None,
        "source_sha256": graph.file_sha256,
        "document_part": document_part,
        "items": [item.to_dict() for item in items],
    }


def read_custom_xml_part(source: str | Path, store_item_id: str) -> CustomXmlStoragePart | None:
    """Return one exact document-owned Custom XML item by its declared store identity."""
    item_key = _store_item_key(store_item_id)
    try:
        _, items = _custom_xml_item_parts(PackageGraph.from_file(source))
    except (OSError, PackageError, ValueError):
        return None
    return next(
        (item for item in items if _store_item_key(item.store_item_id) == item_key),
        None,
    )


def parse_custom_xml_storage_request(value: Any) -> CustomXmlStorageRequest:
    """Parse a strict JSON-compatible Custom XML storage lifecycle request."""
    if not isinstance(value, Mapping) or set(value) != {
        "source",
        "destination",
        "source_sha256",
        "edits",
    }:
        raise ValueError("custom_xml_request_invalid")
    if not isinstance(value["source"], str) or not value["source"]:
        raise ValueError("custom_xml_source_invalid")
    if not isinstance(value["destination"], str) or not value["destination"]:
        raise ValueError("custom_xml_destination_invalid")
    raw_edits = value["edits"]
    if isinstance(raw_edits, (str, bytes)) or not isinstance(raw_edits, Sequence):
        raise TypeError("custom_xml_edits_invalid")
    edits: list[CustomXmlStorageEdit] = []
    for raw_edit in raw_edits:
        if not isinstance(raw_edit, Mapping):
            raise TypeError("custom_xml_edit_invalid")
        allowed = {"action", "store_item_id", "payload", "schema_refs"}
        if set(raw_edit) - allowed or {"action", "store_item_id"} - set(raw_edit):
            raise ValueError("custom_xml_edit_invalid")
        edits.append(
            CustomXmlStorageEdit(
                action=raw_edit["action"],
                store_item_id=raw_edit["store_item_id"],
                payload=raw_edit.get("payload"),
                schema_refs=raw_edit.get("schema_refs"),
            )
        )
    return CustomXmlStorageRequest(
        source=Path(value["source"]),
        destination=Path(value["destination"]),
        source_sha256=value["source_sha256"],
        edits=tuple(edits),
    )


def load_custom_xml_storage_request(path: str | Path) -> CustomXmlStorageRequest:
    """Load a Custom XML storage lifecycle request from a UTF-8 JSON file."""
    with Path(path).open(encoding="utf-8") as stream:
        return parse_custom_xml_storage_request(json.load(stream))


def _remove_content_type_override(ct_root: ET.Element, part_name: str) -> None:
    override_name = "/" + part_name.lstrip("/")
    for node in list(ct_root.findall(f"{{{CONTENT_TYPES_NS}}}Override")):
        if node.get("PartName") == override_name:
            ct_root.remove(node)


def _new_custom_xml_serial(parts: Mapping[str, bytes]) -> int:
    serials = {
        int(match.group(1))
        for name in parts
        if (match := _CUSTOM_XML_ITEM_NAME.fullmatch(name)) is not None
    }
    serial = 1
    while serial in serials:
        serial += 1
    return serial


def _custom_xml_refusal(request: CustomXmlStorageRequest, code: str) -> CustomXmlStorageReport:
    return CustomXmlStorageReport(_REFUSED, code, request.source_sha256, None, (), ())


def _write_custom_xml_staged(
    source: Path,
    staged: Path,
    replacements: Mapping[str, bytes],
    additions: Mapping[str, bytes],
    removals: frozenset[str],
) -> None:
    with (
        zipfile.ZipFile(source) as archive,
        zipfile.ZipFile(staged, "w", compression=zipfile.ZIP_DEFLATED) as rewritten,
    ):
        seen: set[str] = set()
        for info in archive.infolist():
            if info.filename in seen:
                continue
            seen.add(info.filename)
            if info.filename not in removals:
                rewritten.writestr(
                    info, replacements.get(info.filename, archive.read(info.filename))
                )
        for name, data in additions.items():
            if name in seen:
                raise ValueError("custom_xml_addition_conflict")
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            rewritten.writestr(info, data)


def author_custom_xml_storage(
    request: CustomXmlStorageRequest,
) -> CustomXmlStorageReport:
    """Add, replace, or remove document-owned Custom XML payload and properties parts."""
    if not request.source.is_file():
        return _custom_xml_refusal(request, "custom_xml_source_missing")
    if request.destination.exists() or request.destination.is_symlink():
        return _custom_xml_refusal(request, "custom_xml_destination_invalid")
    try:
        graph = PackageGraph.from_file(request.source)
        if graph.file_sha256 != request.source_sha256:
            return _custom_xml_refusal(request, "custom_xml_source_stale")
        document_part, existing_items = _custom_xml_item_parts(graph)
        strict = _package_strict(graph)
        namespace = _STRICT_NAMESPACES[CUSTOM_XML_NS] if strict else CUSTOM_XML_NS
        xml_rel_type = _strict_relationship(CUSTOM_XML_REL_TYPE) if strict else CUSTOM_XML_REL_TYPE
        props_rel_type = (
            _strict_relationship(CUSTOM_XML_PROPERTIES_REL_TYPE)
            if strict
            else CUSTOM_XML_PROPERTIES_REL_TYPE
        )
        parts = {name: graph.read_part_bytes(name) for name in graph.parts}
        ct_root = parse_xml_with_preserved_qnames(parts["[Content_Types].xml"])
        document_rels_part = _custom_xml_relationship_part(document_part)
        document_rels_original = parts.get(document_rels_part)
        document_rels = (
            parse_xml_with_preserved_qnames(document_rels_original)
            if document_rels_original is not None
            else ET.Element(f"{{{PKG_REL_NS}}}Relationships")
        )
    except (OSError, PackageError, ValueError, ET.ParseError, KeyError):
        return _custom_xml_refusal(request, "custom_xml_source_package_invalid")

    existing_by_id = {_store_item_key(item.store_item_id): item for item in existing_items}
    replacements: dict[str, bytes] = {}
    additions: dict[str, bytes] = {}
    removals: set[str] = set()
    changed_targets: list[dict[str, str]] = []
    document_rels_changed = False
    content_types_changed = False
    for edit in request.edits:
        item_key = _store_item_key(edit.store_item_id)
        existing = existing_by_id.get(item_key)
        if edit.action == "add":
            if existing is not None:
                return _custom_xml_refusal(request, "custom_xml_store_item_id_exists")
            serial = _new_custom_xml_serial(parts | additions)
            payload_part = f"customXml/item{serial}.xml"
            properties_part = f"customXml/itemProps{serial}.xml"
            item_rels_part = _custom_xml_relationship_part(payload_part)
            additions[payload_part] = edit.payload
            additions[properties_part] = _custom_xml_properties_xml(
                edit.store_item_id,
                edit.schema_refs or (),
                namespace,
            )
            additions[item_rels_part] = (
                f'<?xml version="1.0" encoding="UTF-8"?>'
                f'<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rIdProps" '
                f'Type="{props_rel_type}" Target="itemProps{serial}.xml"/>'
                "</Relationships>"
            ).encode()
            target = posixpath.relpath(payload_part, posixpath.dirname(document_part))
            ET.SubElement(
                document_rels,
                f"{{{PKG_REL_NS}}}Relationship",
                {
                    "Id": _new_rel_id(document_rels),
                    "Type": xml_rel_type,
                    "Target": target,
                },
            )
            _ensure_content_type_override(ct_root, payload_part, CUSTOM_XML_DATA_CONTENT_TYPE)
            _ensure_content_type_override(
                ct_root, properties_part, CUSTOM_XML_PROPERTIES_CONTENT_TYPE
            )
            document_rels_changed = True
            content_types_changed = True
        elif existing is None:
            return _custom_xml_refusal(request, "custom_xml_store_item_id_missing")
        elif edit.action in {"replace", "update"}:
            if edit.payload is not None:
                replacements[existing.payload_part] = edit.payload
            if edit.schema_refs is not None:
                replacements[existing.properties_part] = _custom_xml_properties_xml(
                    existing.store_item_id,
                    edit.schema_refs,
                    namespace,
                )
        else:
            removals.update(
                {
                    existing.payload_part,
                    existing.properties_part,
                    _custom_xml_relationship_part(existing.payload_part),
                }
            )
            for relationship in list(document_rels.findall(f"{{{PKG_REL_NS}}}Relationship")):
                if relationship.get("Id") == existing.relationship_id:
                    document_rels.remove(relationship)
                    document_rels_changed = True
            _remove_content_type_override(ct_root, existing.payload_part)
            _remove_content_type_override(ct_root, existing.properties_part)
            content_types_changed = True
        changed_targets.append({"action": edit.action, "store_item_id": edit.store_item_id})

    if document_rels_changed:
        serialized_rels = ET.tostring(document_rels, encoding="utf-8", xml_declaration=True)
        if document_rels_original is None:
            additions[document_rels_part] = serialized_rels
        else:
            replacements[document_rels_part] = serialize_xml_with_preserved_qnames(document_rels)
    if content_types_changed:
        replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(ct_root)
    touched = frozenset(set(replacements) | set(additions) | removals)
    staged: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=request.destination.parent,
            prefix=f".{request.destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as stream:
            staged = Path(stream.name)
        _write_custom_xml_staged(
            request.source, staged, replacements, additions, frozenset(removals)
        )
        if _sha256(request.source) != request.source_sha256:
            return _custom_xml_refusal(request, "custom_xml_staged_verification_failed")
        if not _unrelated_parts_preserved(request.source, staged, touched):
            return _custom_xml_refusal(request, "custom_xml_staged_verification_failed")
        candidate = inspect_docx(staged)
        candidate_graph = PackageGraph.from_file(staged)
        candidate_graph.validate_relationship_graph()
        _, candidate_items = _custom_xml_item_parts(candidate_graph)
        if not candidate.valid or _sha256(request.source) != request.source_sha256:
            return _custom_xml_refusal(request, "custom_xml_staged_verification_failed")
        try:
            os.link(staged, request.destination)
        except FileExistsError:
            return _custom_xml_refusal(request, "custom_xml_destination_invalid")
        except (AttributeError, NotImplementedError, OSError):
            return _custom_xml_refusal(request, "custom_xml_publication_unavailable")
        return CustomXmlStorageReport(
            _PUBLISHED,
            None,
            request.source_sha256,
            candidate.file_sha256,
            candidate_items,
            tuple(changed_targets),
        )
    except (OSError, PackageError, ValueError, ET.ParseError):
        return _custom_xml_refusal(request, "custom_xml_staged_verification_failed")
    finally:
        if staged is not None:
            staged.unlink(missing_ok=True)
