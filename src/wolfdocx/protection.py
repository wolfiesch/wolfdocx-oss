"""Source-bound documentProtection discovery and atomic authoring for WordprocessingML.

This module implements atomic authoring and discovery of document editing restrictions
(documentProtection) in ``word/settings.xml`` per ECMA-376 / ISO/IEC 29500:
- Supported editing restriction modes: ``readOnly``, ``comments``, ``trackedChanges``, ``forms``
- Explicit removal of document protection
- Enforcement flag (``w:enforcement="1"`` or ``"0"``)
- Formatting restrictions flag (``w:formatting="1"`` or ``"0"``)
- Strict OOXML schema ordering within ``w:settings``
- Safe creation of ``word/settings.xml``, document relationships, and content type
  overrides when absent
- Cryptographic / password metadata boundary enforcement: editing restrictions are not encryption;
  existing cryptographic metadata and foreign settings are never silently discarded
- Reusable no-clobber publication via ``wolfdocx.review._publish``
- Unrelated package parts are verified byte-identical
- Editing exceptions (``w:permStart``/``w:permEnd``) over whole body or table-cell
  paragraphs for one named editor (``w:ed``) or one editor group (``w:edGrp``) under
  ``readOnly`` or ``comments`` protection; existing ranges are preserved and inventoried
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import zipfile
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .assembly import _new_rel_id, _rels_name
from .inspect import (
    Q,
    inspect_docx,
    parse_xml_with_preserved_qnames,
    serialize_xml_with_preserved_qnames,
)
from .package import PackageError, PackageGraph
from .review import (
    _DOCUMENT,
    _publish,
    _word_namespace,
)

__all__ = [
    "HyphenationReport",
    "HyphenationRequest",
    "PermissionRange",
    "ProtectionReport",
    "ProtectionRequest",
    "author_hyphenation",
    "author_protection",
    "discover_protection_targets",
    "hyphenation_structure_fingerprint",
    "load_protection_request",
    "parse_protection_request",
    "protection_password_hash",
    "protection_structure_fingerprint",
    "verify_write_protection",
]

_SETTINGS = "word/settings.xml"
_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)

_PUBLISHED = "published"
_REFUSED = "refused"

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_SUPPORTED_EDIT_MODES = ("readOnly", "comments", "trackedChanges", "forms")

# Word honors editing exceptions only for these document-wide restrictions.
_PERMISSION_RANGE_EDIT_MODES = ("readOnly", "comments")
# ST_EdGrp enumeration from wml.xsd (ISO/IEC 29500-1 17.18.18).
_PERMISSION_EDITOR_GROUPS = (
    "none",
    "everyone",
    "administrators",
    "contributors",
    "editors",
    "owners",
    "current",
)
_PERMISSION_RANGE_KEYS = frozenset({"start_paragraph", "end_paragraph", "editor", "editor_group"})
_MAX_PERMISSION_RANGES = 64
_MAX_PERMISSION_EDITOR_CHARACTERS = 256
# Paragraph containers whose child paragraphs one range may span.
_PERMISSION_RANGE_CONTAINERS = frozenset({"body", "tc"})
# Block-level siblings a multi-paragraph range may cross besides paragraphs.
_PERMISSION_RANGE_BLOCK_MARKERS = frozenset(
    {
        "bookmarkStart",
        "bookmarkEnd",
        "permStart",
        "permEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "proofErr",
    }
)

_CRYPT_ATTRIBUTES = frozenset(
    {
        "cryptProviderType",
        "cryptAlgorithmClass",
        "cryptAlgorithmType",
        "cryptAlgorithmSid",
        "cryptSpinCount",
        "cryptProvider",
        "algIdExt",
        "algIdExtSource",
        "cryptProviderTypeExt",
        "cryptProviderTypeExtSource",
        "hash",
        "salt",
    }
)

_FORBIDDEN_REQUEST_CRYPTO_KEYS = frozenset(
    {
        "password",
        "pass",
        "pwd",
        "crypto",
        "crypt",
        "algorithm",
        "hash",
        "salt",
        "cryptProviderType",
        "cryptAlgorithmClass",
        "cryptAlgorithmType",
        "cryptAlgorithmSid",
        "cryptSpinCount",
        "cryptProvider",
    }
)

# OOXML CT_Settings sequence definition from wml.xsd
_SETTINGS_CHILD_ORDER = (
    "writeProtection",
    "view",
    "zoom",
    "removePersonalInformation",
    "removeDateAndTime",
    "doNotDisplayPageBoundaries",
    "displayBackgroundShape",
    "printPostScriptOverText",
    "printFractionalCharacterWidth",
    "printFormsData",
    "embedTrueTypeFonts",
    "embedSystemFonts",
    "saveSubsetFonts",
    "saveFormsData",
    "mirrorMargins",
    "alignBordersAndEdges",
    "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter",
    "gutterAtTop",
    "hideSpellingErrors",
    "hideGrammaticalErrors",
    "activeWritingStyle",
    "proofState",
    "formsDesign",
    "attachedTemplate",
    "linkStyles",
    "stylePaneFormatFilter",
    "stylePaneSortMethod",
    "documentType",
    "mailMerge",
    "revisionView",
    "trackRevisions",
    "doNotTrackMoves",
    "doNotTrackFormatting",
    "documentProtection",
    "autoFormatOverride",
    "styleLockTheme",
    "styleLockQFSet",
    "defaultTabStop",
    "autoHyphenation",
    "consecutiveHyphenLimit",
    "hyphenationZone",
    "doNotHyphenateCaps",
    "showEnvelope",
    "summaryLength",
    "clickAndTypeStyle",
    "defaultTableStyle",
    "evenAndOddHeaders",
    "bookFoldRevPrinting",
    "bookFoldPrinting",
    "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin",
    "doNotShadeFormData",
    "noPunctuationKerning",
    "characterSpacingControl",
    "printTwoOnOne",
    "strictFirstAndLastChars",
    "noLineBreaksAfter",
    "noLineBreaksBefore",
    "savePreviewPicture",
    "doNotValidateAgainstSchema",
    "saveInvalidXml",
    "ignoreMixedContent",
    "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml",
    "saveXmlDataOnly",
    "useXSLTWhenSaving",
    "saveThroughXslt",
    "showXMLTags",
    "alwaysMergeEmptyNamespace",
    "updateFields",
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "docVars",
    "rsids",
    "mathPr",
    "uiCompat97To2003",
    "attachedSchema",
    "themeFontLang",
    "clrSchemeMapping",
    "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures",
    "forceUpgrade",
    "captions",
    "readModeInkLockDown",
    "smartTagType",
    "schemaLibrary",
    "shapeDefaults",
    "doNotEmbedSmartTags",
    "decimalSymbol",
    "listSeparator",
)

_SETTINGS_CHILD_ORDER_MAP = {name: i for i, name in enumerate(_SETTINGS_CHILD_ORDER)}

# ECMA-376 Part 4 §14.8.1, initial code and encryption matrix.
_PASSWORD_INITIAL_CODES = (
    0xE1F0,
    0x1D0F,
    0xCC9C,
    0x84C0,
    0x110C,
    0x0E10,
    0xF1CE,
    0x313E,
    0x1872,
    0xE139,
    0xD40F,
    0x84F9,
    0x280C,
    0xA96A,
    0x4EC3,
)
_PASSWORD_MATRIX = (
    (0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09),
    (0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF),
    (0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0),
    (0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40),
    (0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5),
    (0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A),
    (0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9),
    (0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0),
    (0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC),
    (0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10),
    (0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168),
    (0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C),
    (0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD),
    (0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC),
    (0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4),
)


def protection_password_hash(
    password: str,
    *,
    salt: bytes | None = None,
    spin_count: int = 100000,
    legacy: bool = True,
) -> dict[str, str]:
    """Return SHA-512 password-verifier attributes, never the password.

    Legacy preprocessing follows ECMA-376 Part 4 §14.8.1 and MS-OE376
    §2.1.410(a): reverse the key bytes, then hash its UTF-16LE uppercase
    hexadecimal representation. Agile verifiers hash the UTF-16LE password
    directly. Each spin appends a zero-based little-endian 32-bit counter.
    This is an editing reservation, not package encryption.
    """
    if not isinstance(password, str) or not isinstance(legacy, bool):
        raise TypeError("protection_password_invalid")
    if (
        isinstance(spin_count, bool)
        or not isinstance(spin_count, int)
        or not 0 <= spin_count <= 5000000
    ):
        raise ValueError("protection_spin_count_invalid")
    if salt is None:
        salt = os.urandom(16)
    if not isinstance(salt, bytes) or len(salt) != 16:
        raise ValueError("protection_salt_invalid")
    try:
        encoded = password.removeprefix("\ufeff").encode("utf-16le")
    except UnicodeEncodeError:
        raise ValueError("protection_password_invalid") from None
    if legacy:
        # One byte per UTF-16 code unit, low byte unless it is zero.
        reduced = bytes(encoded[i] or encoded[i + 1] for i in range(0, min(len(encoded), 30), 2))
        high = _PASSWORD_INITIAL_CODES[len(reduced) - 1] if reduced else 0
        for position, value in enumerate(reduced):
            row = _PASSWORD_MATRIX[15 - len(reduced) + position]
            for bit in range(7):
                if value & (1 << bit):
                    high ^= row[bit]
        low = 0
        for value in reversed(reduced):
            low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ value
        if reduced:
            low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ len(reduced) ^ 0xCE4B
        encoded = (
            (low.to_bytes(2, "little") + high.to_bytes(2, "little"))
            .hex()
            .upper()
            .encode("utf-16le")
        )
    digest = hashlib.sha512(salt + encoded).digest()
    for count in range(spin_count):
        digest = hashlib.sha512(digest + count.to_bytes(4, "little")).digest()
    hashed, salted = (base64.b64encode(value).decode("ascii") for value in (digest, salt))
    if legacy:
        return {
            "cryptProviderType": "rsaAES",
            "cryptAlgorithmClass": "hash",
            "cryptAlgorithmType": "typeAny",
            "cryptAlgorithmSid": "14",
            "cryptSpinCount": str(spin_count),
            "hash": hashed,
            "salt": salted,
        }
    return {
        "algorithmName": "SHA-512",
        "spinCount": str(spin_count),
        "hashValue": hashed,
        "saltValue": salted,
    }


def verify_write_protection(source: str | Path, password: str) -> bool:
    """Compare a password with a legacy or agile SHA-512 write-reservation verifier.

    Return False only for a mismatch; missing, malformed, or unsupported verifiers
    raise ValueError with a ``protection_*`` code. A non-string password raises
    TypeError. This neither encrypts the package nor enforces a reservation.
    """
    if not isinstance(password, str):
        raise TypeError("protection_password_invalid")
    path = Path(source)
    if not path.is_file():
        raise ValueError("protection_source_missing")
    try:
        graph = PackageGraph.from_file(path)
    except (OSError, PackageError, zipfile.BadZipFile, KeyError):
        raise ValueError("protection_package_invalid") from None
    if _DOCUMENT not in graph.parts:
        raise ValueError("protection_package_invalid")
    if _SETTINGS not in graph.parts:
        raise ValueError("protection_password_verifier_missing")
    try:
        root = ET.fromstring(graph.read_part_bytes(_SETTINGS))
    except (ET.ParseError, PackageError, OSError, zipfile.BadZipFile):
        raise ValueError("protection_settings_invalid") from None
    namespace = _word_namespace(root)
    if namespace is None or root.tag != Q(namespace, "settings"):
        raise ValueError("protection_settings_invalid")
    reservations = root.findall(Q(namespace, "writeProtection"))
    if not reservations:
        raise ValueError("protection_password_verifier_missing")
    if len(reservations) != 1 or len(reservations[0]):
        raise ValueError("protection_password_verifier_invalid")
    attributes = {
        key.removeprefix(f"{{{namespace}}}"): value
        for key, value in reservations[0].attrib.items()
        if key.startswith(f"{{{namespace}}}")
    }
    agile_keys = {"algorithmName", "spinCount", "hashValue", "saltValue"}
    legacy_keys = {
        "cryptProviderType",
        "cryptAlgorithmClass",
        "cryptAlgorithmType",
        "cryptAlgorithmSid",
        "cryptSpinCount",
        "hash",
        "salt",
    }
    has_agile = bool(agile_keys.intersection(attributes))
    has_legacy = bool(_CRYPT_ATTRIBUTES.intersection(attributes))
    if has_agile and has_legacy:
        raise ValueError("protection_password_verifier_invalid")
    if "password" in attributes:
        raise ValueError("protection_algorithm_unsupported")
    if not has_agile and not has_legacy:
        raise ValueError("protection_password_verifier_missing")
    if (_CRYPT_ATTRIBUTES - legacy_keys).intersection(attributes):
        raise ValueError("protection_provider_unsupported")
    required = agile_keys if has_agile else legacy_keys
    if not required.issubset(attributes):
        raise ValueError("protection_password_verifier_invalid")
    if has_agile:
        if attributes["algorithmName"] != "SHA-512":
            raise ValueError("protection_algorithm_unsupported")
        spin_value = attributes["spinCount"]
        salt_value = attributes["saltValue"]
        hash_value = attributes["hashValue"]
    else:
        if attributes["cryptProviderType"] != "rsaAES":
            raise ValueError("protection_provider_unsupported")
        if (
            attributes["cryptAlgorithmClass"] != "hash"
            or attributes["cryptAlgorithmType"] != "typeAny"
            or attributes["cryptAlgorithmSid"] != "14"
        ):
            raise ValueError("protection_algorithm_unsupported")
        spin_value = attributes["cryptSpinCount"]
        salt_value = attributes["salt"]
        hash_value = attributes["hash"]
    if not spin_value.isascii() or not spin_value.isdecimal():
        raise ValueError("protection_spin_count_invalid")
    try:
        spins = int(spin_value)
    except ValueError:
        raise ValueError("protection_spin_count_invalid") from None
    if not 0 <= spins <= 5000000:
        raise ValueError("protection_spin_count_invalid")
    try:
        salt = base64.b64decode(salt_value, validate=True)
    except ValueError:
        raise ValueError("protection_salt_invalid") from None
    if len(salt) != 16:
        raise ValueError("protection_salt_invalid")
    try:
        expected = base64.b64decode(hash_value, validate=True)
    except ValueError:
        raise ValueError("protection_hash_invalid") from None
    if len(expected) != hashlib.sha512().digest_size:
        raise ValueError("protection_hash_invalid")
    candidate = protection_password_hash(
        password, salt=salt, spin_count=spins, legacy=not has_agile
    )
    actual = base64.b64decode(candidate["hashValue" if has_agile else "hash"], validate=True)
    return hmac.compare_digest(actual, expected)


_ALLOWED_REQUEST_KEYS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "structure_fingerprint_sha256",
        "action",
        "edit",
        "enforcement",
        "formatting",
        "part",
        "permission_ranges",
    }
)

_REQUIRED_REQUEST_KEYS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "structure_fingerprint_sha256",
    }
)

_ABSENT_SETTINGS_FINGERPRINT = hashlib.sha256(b"absent:word/settings.xml").hexdigest()
_ABSENT_PROTECTION_FINGERPRINT = hashlib.sha256(b"absent:documentProtection").hexdigest()


def _sha256_file(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def _element_structure_fingerprint(element: ET.Element) -> str:
    """Hash the XML element structure deterministically without mutating."""

    def project(elem: ET.Element) -> list[Any]:
        return [
            elem.tag,
            sorted(elem.attrib.items()),
            elem.text,
            elem.tail,
            [project(child) for child in elem],
        ]

    canonical = json.dumps(project(element), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def protection_structure_fingerprint(
    source: str | Path,
    *,
    part: str = _SETTINGS,
) -> str | None:
    """Return the structural fingerprint for settings.xml or documentProtection element."""
    path = Path(source)
    if not path.is_file():
        return None
    try:
        graph = PackageGraph.from_file(path)
    except (OSError, PackageError, zipfile.BadZipFile):
        return None

    if part not in graph.parts:
        return _ABSENT_SETTINGS_FINGERPRINT

    try:
        settings_root = ET.fromstring(graph.read_part_bytes(part))
    except (ET.ParseError, PackageError):
        return None

    ns = _word_namespace(settings_root) or _W
    protect_elems = settings_root.findall(Q(ns, "documentProtection"))
    if not protect_elems:
        return _ABSENT_PROTECTION_FINGERPRINT

    if len(protect_elems) == 1:
        return _element_structure_fingerprint(protect_elems[0])

    return _element_structure_fingerprint(settings_root)


def _parse_on_off(val: str | None) -> bool:
    if val is None:
        return False
    return str(val).strip().lower() in {"1", "true", "on"}


def _insert_settings_child(
    settings_root: ET.Element, new_child: ET.Element, namespace: str
) -> None:
    """Insert a new child element into settings_root maintaining OOXML schema order."""
    local_name = new_child.tag.rsplit("}", 1)[-1]
    new_order = _SETTINGS_CHILD_ORDER_MAP.get(local_name, 999)
    insert_idx = len(settings_root)
    for idx, child in enumerate(settings_root):
        child_local = child.tag.rsplit("}", 1)[-1]
        child_order = _SETTINGS_CHILD_ORDER_MAP.get(child_local, 999)
        if child_order > new_order:
            insert_idx = idx
            break
    settings_root.insert(insert_idx, new_child)


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


def _validate_sha256(val: Any, name: str) -> str:
    if (
        not isinstance(val, str)
        or len(val) != 64
        or any(c not in "0123456789abcdefABCDEF" for c in val)
    ):
        raise ValueError(f"{name}_invalid")
    return val.lower()


def _valid_permission_editor(value: Any) -> bool:
    return (
        isinstance(value, str)
        and 0 < len(value) <= _MAX_PERMISSION_EDITOR_CHARACTERS
        and value == value.strip()
        and value.isprintable()
    )


@dataclass(frozen=True)
class PermissionRange:
    """One editing exception (``w:permStart``/``w:permEnd``) over whole paragraphs.

    ``start_paragraph`` and ``end_paragraph`` are inclusive 1-based indexes over every
    ``w:p`` in ``word/document.xml`` in document order, the same addressing used by the
    other paragraph-bound authoring requests. Both paragraphs must be children of the
    same body or table cell. Exactly one of ``editor`` (``w:ed``, one named user such as
    an email address or ``DOMAIN\\user``) or ``editor_group`` (``w:edGrp``, one
    ``ST_EdGrp`` group such as ``everyone``) is set.
    """

    start_paragraph: int
    end_paragraph: int
    editor: str | None = None
    editor_group: str | None = None

    def __post_init__(self) -> None:
        for index in (self.start_paragraph, self.end_paragraph):
            if isinstance(index, bool) or not isinstance(index, int) or index < 1:
                raise ValueError("protection_permission_range_paragraph_invalid")
        if self.end_paragraph < self.start_paragraph:
            raise ValueError("protection_permission_range_paragraph_invalid")
        if (self.editor is None) == (self.editor_group is None):
            raise ValueError("protection_permission_range_editor_invalid")
        if self.editor is not None and not _valid_permission_editor(self.editor):
            raise ValueError("protection_permission_range_editor_invalid")
        if self.editor_group is not None and self.editor_group not in _PERMISSION_EDITOR_GROUPS:
            raise ValueError("protection_permission_range_editor_group_unsupported")

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "start_paragraph": self.start_paragraph,
            "end_paragraph": self.end_paragraph,
        }
        if self.editor is not None:
            result["editor"] = self.editor
        else:
            result["editor_group"] = self.editor_group
        return result


def _parse_permission_ranges(value: Any) -> tuple[PermissionRange, ...]:
    if not isinstance(value, list) or not value or len(value) > _MAX_PERMISSION_RANGES:
        raise ValueError("protection_permission_ranges_invalid")
    ranges: list[PermissionRange] = []
    for item in value:
        if not isinstance(item, Mapping):
            # Preserve the public malformed-request ValueError contract.
            raise ValueError("protection_permission_ranges_invalid")  # noqa: TRY004
        keys = set(item)
        if not {"start_paragraph", "end_paragraph"} <= keys <= _PERMISSION_RANGE_KEYS:
            raise ValueError("protection_permission_ranges_invalid")
        ranges.append(
            PermissionRange(
                start_paragraph=item["start_paragraph"],
                end_paragraph=item["end_paragraph"],
                editor=item.get("editor"),
                editor_group=item.get("editor_group"),
            )
        )
    return tuple(ranges)


@dataclass(frozen=True)
class ProtectionRequest:
    """Strict request for modifying or removing document editing restrictions."""

    source: Path
    destination: Path
    source_sha256: str
    structure_fingerprint_sha256: str
    action: str = "set"
    edit: str | None = None
    enforcement: bool = True
    formatting: bool = False
    part: str = _SETTINGS
    permission_ranges: tuple[PermissionRange, ...] = ()

    def __post_init__(self) -> None:
        if not isinstance(self.source, Path):
            object.__setattr__(self, "source", Path(self.source))
        if not isinstance(self.destination, Path):
            object.__setattr__(self, "destination", Path(self.destination))
        object.__setattr__(
            self,
            "source_sha256",
            _validate_sha256(self.source_sha256, "protection_source_sha256"),
        )
        object.__setattr__(
            self,
            "structure_fingerprint_sha256",
            _validate_sha256(
                self.structure_fingerprint_sha256,
                "protection_structure_fingerprint_sha256",
            ),
        )
        if self.action not in {"set", "remove"}:
            raise ValueError("protection_action_unsupported")
        if (
            self.action == "set" and (self.edit is None or self.edit not in _SUPPORTED_EDIT_MODES)
        ) or (
            self.action == "remove"
            and self.edit is not None
            and self.edit not in _SUPPORTED_EDIT_MODES
            and self.edit != "none"
        ):
            raise ValueError("protection_edit_mode_unsupported")
        if self.part != _SETTINGS:
            raise ValueError("protection_part_unsupported")
        ranges = tuple(self.permission_ranges)
        if len(ranges) > _MAX_PERMISSION_RANGES or not all(
            isinstance(item, PermissionRange) for item in ranges
        ):
            raise ValueError("protection_permission_ranges_invalid")
        if len({tuple(sorted(item.to_dict().items())) for item in ranges}) != len(ranges):
            raise ValueError("protection_permission_ranges_duplicate")
        if ranges and (self.action != "set" or self.edit not in _PERMISSION_RANGE_EDIT_MODES):
            raise ValueError("protection_permission_ranges_mode_unsupported")
        object.__setattr__(self, "permission_ranges", ranges)

    @property
    def target(self) -> dict[str, Any]:
        return {"part": self.part}

    def target_dict(self) -> dict[str, Any]:
        return {"part": self.part}

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "source": str(self.source),
            "destination": str(self.destination),
            "source_sha256": self.source_sha256,
            "structure_fingerprint_sha256": self.structure_fingerprint_sha256,
            "action": self.action,
            "part": self.part,
        }
        if self.edit is not None:
            result["edit"] = self.edit
        if self.enforcement is not None:
            result["enforcement"] = self.enforcement
        if self.formatting is not None:
            result["formatting"] = self.formatting
        if self.permission_ranges:
            result["permission_ranges"] = [item.to_dict() for item in self.permission_ranges]
        return result


@dataclass(frozen=True)
class ProtectionReport:
    """Content-safe public report for one document protection authoring operation."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    target: dict[str, Any] | None = None
    protection: dict[str, Any] | None = None
    target_change_report: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "schema_version": 1,
            "operation": "protection",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
        }
        if self.target is not None:
            result["target"] = self.target
        if self.protection is not None:
            result["protection"] = self.protection
        if self.target_change_report is not None:
            result["target_change_report"] = self.target_change_report
        if self.non_target_parts_preserved is not None:
            result["non_target_parts_preserved"] = self.non_target_parts_preserved
        return result


def _protection_refusal(request: ProtectionRequest, code: str) -> ProtectionReport:
    return ProtectionReport(
        status=_REFUSED,
        code=code,
        source_sha256=request.source_sha256,
        destination_sha256=None,
        target={"part": request.part},
        protection=None,
        target_change_report=None,
        non_target_parts_preserved=None,
    )


def parse_protection_request(value: Any) -> ProtectionRequest:
    """Parse and strictly validate one JSON-compatible protection request."""
    if not isinstance(value, Mapping):
        raise TypeError("protection_request_invalid")

    keys = set(value)

    # Check for forbidden cryptographic or password fields
    for forbidden in _FORBIDDEN_REQUEST_CRYPTO_KEYS:
        if forbidden in keys:
            raise ValueError("protection_cryptographic_unsupported")

    if not _REQUIRED_REQUEST_KEYS.issubset(keys):
        raise ValueError("protection_request_fields_invalid")

    if not keys.issubset(_ALLOWED_REQUEST_KEYS):
        raise ValueError("protection_request_fields_invalid")

    source_raw = value["source"]
    dest_raw = value["destination"]
    source_sha_raw = value["source_sha256"]
    fp_raw = value["structure_fingerprint_sha256"]

    action = value.get("action")
    edit = value.get("edit")
    if action is None:
        action = "set" if edit is not None else "remove"
    elif not isinstance(action, str) or action not in {"set", "remove"}:
        raise ValueError("protection_action_unsupported")

    if action == "set":
        if edit is None or not isinstance(edit, str) or edit not in _SUPPORTED_EDIT_MODES:
            raise ValueError("protection_edit_mode_unsupported")
    else:  # action == "remove"
        if edit is not None and (
            not isinstance(edit, str) or (edit not in _SUPPORTED_EDIT_MODES and edit != "none")
        ):
            raise ValueError("protection_edit_mode_unsupported")

    enf_raw = value.get("enforcement")
    if enf_raw is None:
        enforcement = True
    elif isinstance(enf_raw, bool):
        enforcement = enf_raw
    elif isinstance(enf_raw, (int, str)):
        enforcement = str(enf_raw).strip().lower() in {"1", "true", "on"}
    else:
        raise ValueError("protection_enforcement_invalid")

    fmt_raw = value.get("formatting")
    if fmt_raw is None:
        formatting = False
    elif isinstance(fmt_raw, bool):
        formatting = fmt_raw
    elif isinstance(fmt_raw, (int, str)):
        formatting = str(fmt_raw).strip().lower() in {"1", "true", "on"}
    else:
        raise ValueError("protection_formatting_invalid")

    part = value.get("part", _SETTINGS)
    if not isinstance(part, str) or part != _SETTINGS:
        raise ValueError("protection_part_unsupported")

    permission_ranges = (
        _parse_permission_ranges(value["permission_ranges"]) if "permission_ranges" in value else ()
    )

    return ProtectionRequest(
        source=Path(source_raw),
        destination=Path(dest_raw),
        source_sha256=str(source_sha_raw),
        structure_fingerprint_sha256=str(fp_raw),
        action=action,
        edit=edit,
        enforcement=enforcement,
        formatting=formatting,
        part=part,
        permission_ranges=permission_ranges,
    )


def load_protection_request(path: str | Path) -> ProtectionRequest:
    """Load and strictly validate a protection request from JSON on disk."""
    with Path(path).open("r", encoding="utf-8") as stream:
        return parse_protection_request(json.load(stream))


def _containing_paragraph(
    element: ET.Element, parents: Mapping[ET.Element, ET.Element], paragraph_tag: str
) -> ET.Element | None:
    ancestor = parents.get(element)
    while ancestor is not None and ancestor.tag != paragraph_tag:
        ancestor = parents.get(ancestor)
    return ancestor


def _permission_range_inventory(root: ET.Element, namespace: str) -> list[dict[str, Any]]:
    """Inventory ``w:permStart``/``w:permEnd`` pairs by 1-based paragraph position.

    A marker inside a paragraph reports that paragraph. A block-level start reports the
    next paragraph and a block-level end the previous one: the text each one bounds.
    Ends without a start are reported as unbalanced entries without an editor.
    """
    paragraph_tag = Q(namespace, "p")
    start_tag = Q(namespace, "permStart")
    end_tag = Q(namespace, "permEnd")
    id_attribute = Q(namespace, "id")
    parents = {child: parent for parent in root.iter() for child in parent}
    paragraph_index: dict[ET.Element, int] = {}
    starts: list[tuple[ET.Element, int]] = []
    ends: dict[str | None, list[int]] = {}
    count = 0
    for element in root.iter():
        if element.tag == paragraph_tag:
            count += 1
            paragraph_index[element] = count
        elif element.tag == start_tag or element.tag == end_tag:
            paragraph = _containing_paragraph(element, parents, paragraph_tag)
            if paragraph is not None:
                position = paragraph_index[paragraph]
            else:
                position = count + 1 if element.tag == start_tag else count
            if element.tag == start_tag:
                starts.append((element, position))
            else:
                ends.setdefault(element.get(id_attribute), []).append(position)

    start_ids = [element.get(id_attribute) for element, _ in starts]
    inventory: list[dict[str, Any]] = []
    for element, position in starts:
        range_id = element.get(id_attribute)
        matched = ends.get(range_id, [])
        balanced = (
            range_id is not None
            and start_ids.count(range_id) == 1
            and len(matched) == 1
            and matched[0] >= position
        )
        entry: dict[str, Any] = {
            "id": range_id,
            "editor": element.get(Q(namespace, "ed")),
            "editor_group": element.get(Q(namespace, "edGrp")),
            "start_paragraph": position,
            "end_paragraph": matched[0] if balanced else None,
            "balanced": balanced,
        }
        for attribute, key in (
            ("colFirst", "column_first"),
            ("colLast", "column_last"),
        ):
            column = element.get(Q(namespace, attribute))
            if column is not None:
                entry[key] = column
        inventory.append(entry)
    for range_id, positions in ends.items():
        if range_id in start_ids:
            continue
        inventory.extend(
            {
                "id": range_id,
                "editor": None,
                "editor_group": None,
                "start_paragraph": None,
                "end_paragraph": position,
                "balanced": False,
            }
            for position in positions
        )
    return inventory


def discover_protection_targets(
    source: str | Path,
    *,
    part: str = _SETTINGS,
) -> dict[str, Any]:
    """Inventory document protection settings and editing exceptions without mutating."""
    path = Path(source)
    if not path.is_file():
        return {
            "passed": False,
            "status": _REFUSED,
            "code": "protection_source_missing",
            "source_sha256": "",
            "structure_fingerprint_sha256": "",
            "settings_part_present": False,
            "protected": False,
            "protection": None,
            "targets": [],
            "supported_modes": list(_SUPPORTED_EDIT_MODES),
            "permission_ranges": [],
        }

    try:
        source_sha = _sha256_file(path)
        graph = PackageGraph.from_file(path)
        if _DOCUMENT not in graph.parts:
            return {
                "passed": False,
                "status": _REFUSED,
                "code": "protection_package_invalid",
                "source_sha256": source_sha,
                "structure_fingerprint_sha256": "",
                "settings_part_present": False,
                "protected": False,
                "protection": None,
                "targets": [],
                "supported_modes": list(_SUPPORTED_EDIT_MODES),
                "permission_ranges": [],
            }
        document_root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
    except (OSError, zipfile.BadZipFile, KeyError, PackageError, ET.ParseError):
        return {
            "passed": False,
            "status": _REFUSED,
            "code": "protection_package_invalid",
            "source_sha256": "",
            "structure_fingerprint_sha256": "",
            "settings_part_present": False,
            "protected": False,
            "protection": None,
            "targets": [],
            "supported_modes": list(_SUPPORTED_EDIT_MODES),
            "permission_ranges": [],
        }

    permission_ranges = _permission_range_inventory(
        document_root, _word_namespace(document_root) or _W
    )
    fp = protection_structure_fingerprint(path, part=part) or ""

    if part not in graph.parts:
        target_info = {
            "part": part,
            "protected": False,
            "edit": None,
            "enforcement": None,
            "formatting": None,
            "has_cryptographic_metadata": False,
            "structure_fingerprint_sha256": fp,
        }
        return {
            "passed": True,
            "status": "discovered",
            "code": None,
            "source_sha256": source_sha,
            "structure_fingerprint_sha256": fp,
            "settings_part_present": False,
            "protected": False,
            "protection": None,
            "targets": [target_info],
            "supported_modes": list(_SUPPORTED_EDIT_MODES),
            "permission_ranges": permission_ranges,
        }

    try:
        settings_bytes = graph.read_part_bytes(part)
        settings_root = ET.fromstring(settings_bytes)
    except (ET.ParseError, PackageError):
        return {
            "passed": False,
            "status": _REFUSED,
            "code": "protection_settings_invalid",
            "source_sha256": source_sha,
            "structure_fingerprint_sha256": fp,
            "settings_part_present": True,
            "protected": False,
            "protection": None,
            "targets": [],
            "supported_modes": list(_SUPPORTED_EDIT_MODES),
            "permission_ranges": permission_ranges,
        }

    ns = _word_namespace(settings_root) or _W
    doc_protect_elems = settings_root.findall(Q(ns, "documentProtection"))
    if len(doc_protect_elems) > 1:
        return {
            "passed": False,
            "status": _REFUSED,
            "code": "protection_structure_unsupported",
            "source_sha256": source_sha,
            "structure_fingerprint_sha256": fp,
            "settings_part_present": True,
            "protected": True,
            "protection": None,
            "targets": [],
            "supported_modes": list(_SUPPORTED_EDIT_MODES),
            "permission_ranges": permission_ranges,
        }

    if not doc_protect_elems:
        target_info = {
            "part": part,
            "protected": False,
            "edit": None,
            "enforcement": None,
            "formatting": None,
            "has_cryptographic_metadata": False,
            "structure_fingerprint_sha256": fp,
        }
        return {
            "passed": True,
            "status": "discovered",
            "code": None,
            "source_sha256": source_sha,
            "structure_fingerprint_sha256": fp,
            "settings_part_present": True,
            "protected": False,
            "protection": None,
            "targets": [target_info],
            "supported_modes": list(_SUPPORTED_EDIT_MODES),
            "permission_ranges": permission_ranges,
        }

    dp_elem = doc_protect_elems[0]
    edit_val = dp_elem.get(Q(ns, "edit")) or dp_elem.get("edit")
    enf_raw = dp_elem.get(Q(ns, "enforcement")) or dp_elem.get("enforcement")
    fmt_raw = dp_elem.get(Q(ns, "formatting")) or dp_elem.get("formatting")

    enf_val = _parse_on_off(enf_raw) if enf_raw is not None else True
    fmt_val = _parse_on_off(fmt_raw) if fmt_raw is not None else False

    crypt_attrs: dict[str, str] = {}
    for k, v in dp_elem.attrib.items():
        local = k.rsplit("}", 1)[-1]
        if local in _CRYPT_ATTRIBUTES:
            crypt_attrs[local] = v

    protection_data = {
        "edit": edit_val,
        "enforcement": enf_val,
        "formatting": fmt_val,
        "has_cryptographic_metadata": len(crypt_attrs) > 0,
        "cryptographic_attributes": sorted(crypt_attrs.keys()),
    }
    target_info = {
        "part": part,
        "protected": True,
        "edit": edit_val,
        "enforcement": enf_val,
        "formatting": fmt_val,
        "has_cryptographic_metadata": len(crypt_attrs) > 0,
        "structure_fingerprint_sha256": fp,
    }

    return {
        "passed": True,
        "status": "discovered",
        "code": None,
        "source_sha256": source_sha,
        "structure_fingerprint_sha256": fp,
        "settings_part_present": True,
        "protected": True,
        "protection": protection_data,
        "targets": [target_info],
        "supported_modes": list(_SUPPORTED_EDIT_MODES),
        "permission_ranges": permission_ranges,
    }


def _author_permission_ranges(
    root: ET.Element, ranges: tuple[PermissionRange, ...]
) -> list[dict[str, Any]] | str:
    """Insert one ``w:permStart``/``w:permEnd`` pair per range; return records or a code.

    The start marker follows the start paragraph's properties (and any leading start
    markers); the end marker closes the end paragraph. Identifiers continue after the
    largest existing decimal range identifier and never reuse an existing one.
    """
    namespace = _word_namespace(root)
    if namespace is None:
        return "protection_package_invalid"
    paragraph_tag = Q(namespace, "p")
    start_tag = Q(namespace, "permStart")
    end_tag = Q(namespace, "permEnd")
    id_attribute = Q(namespace, "id")
    paragraphs = list(root.iter(paragraph_tag))
    parents = {child: parent for parent in root.iter() for child in parent}
    used_ids = {
        element.get(id_attribute)
        for element in root.iter()
        if element.tag == start_tag or element.tag == end_tag
    }
    next_id = 1 + max(
        (int(value) for value in used_ids if value is not None and value.isdecimal()),
        default=0,
    )

    selections: list[tuple[PermissionRange, ET.Element, ET.Element]] = []
    for item in ranges:
        if item.end_paragraph > len(paragraphs):
            return "protection_permission_range_paragraph_not_found"
        first = paragraphs[item.start_paragraph - 1]
        last = paragraphs[item.end_paragraph - 1]
        container = parents.get(first)
        if (
            container is None
            or parents.get(last) is not container
            or container.tag not in {Q(namespace, name) for name in _PERMISSION_RANGE_CONTAINERS}
        ):
            return "protection_permission_range_structure_unsupported"
        siblings = list(container)
        spanned = siblings[siblings.index(first) : siblings.index(last) + 1]
        if any(
            child.tag != paragraph_tag
            and child.tag not in {Q(namespace, name) for name in _PERMISSION_RANGE_BLOCK_MARKERS}
            for child in spanned
        ):
            return "protection_permission_range_structure_unsupported"
        selections.append((item, first, last))

    records: list[dict[str, Any]] = []
    for item, first, last in selections:
        while str(next_id) in used_ids:
            next_id += 1
        range_id = str(next_id)
        used_ids.add(range_id)
        next_id += 1

        marker = ET.Element(start_tag)
        marker.set(id_attribute, range_id)
        if item.editor is not None:
            marker.set(Q(namespace, "ed"), item.editor)
        else:
            marker.set(Q(namespace, "edGrp"), item.editor_group)  # type: ignore[arg-type]
        position = 1 if len(first) and first[0].tag == Q(namespace, "pPr") else 0
        while position < len(first) and first[position].tag == start_tag:
            position += 1
        first.insert(position, marker)
        closing = ET.Element(end_tag)
        closing.set(id_attribute, range_id)
        last.append(closing)
        records.append({"id": range_id, **item.to_dict()})
    return records


def _permission_ranges_verified(
    source_document: bytes, staged_document: bytes, records: list[dict[str, Any]]
) -> bool:
    """True when exactly the recorded markers were added and nothing else changed."""
    try:
        before = ET.fromstring(source_document)
        after = ET.fromstring(staged_document)
    except ET.ParseError:
        return False
    namespace = _word_namespace(after)
    if namespace is None or _word_namespace(before) != namespace:
        return False
    paragraphs = list(after.iter(Q(namespace, "p")))
    parents = {child: parent for parent in after.iter() for child in parent}
    id_attribute = Q(namespace, "id")
    for record in records:
        starts = [
            element
            for element in after.iter(Q(namespace, "permStart"))
            if element.get(id_attribute) == record["id"]
        ]
        ends = [
            element
            for element in after.iter(Q(namespace, "permEnd"))
            if element.get(id_attribute) == record["id"]
        ]
        if len(starts) != 1 or len(ends) != 1:
            return False
        expected_start = {id_attribute: record["id"]}
        if "editor" in record:
            expected_start[Q(namespace, "ed")] = record["editor"]
        else:
            expected_start[Q(namespace, "edGrp")] = record["editor_group"]
        if (
            dict(starts[0].attrib) != expected_start
            or dict(ends[0].attrib) != {id_attribute: record["id"]}
            or parents.get(starts[0]) is not paragraphs[record["start_paragraph"] - 1]
            or parents.get(ends[0]) is not paragraphs[record["end_paragraph"] - 1]
        ):
            return False
        for marker in (starts[0], ends[0]):
            parents[marker].remove(marker)
    return _element_structure_fingerprint(after) == _element_structure_fingerprint(before)


def author_protection(
    request: ProtectionRequest | Mapping[str, Any],
) -> ProtectionReport:
    """Execute source-bound document protection edit or removal atomically."""
    if isinstance(request, Mapping):
        try:
            req = parse_protection_request(request)
        except (TypeError, ValueError):
            source_sha = str(request.get("source_sha256", ""))
            return ProtectionReport(
                status=_REFUSED,
                code="protection_request_invalid",
                source_sha256=source_sha,
                destination_sha256=None,
                target={"part": _SETTINGS},
                protection=None,
                target_change_report=None,
                non_target_parts_preserved=None,
            )
    else:
        req = request

    source = req.source
    destination = req.destination

    try:
        if source == destination or source.resolve() == destination.resolve():
            return _protection_refusal(req, "protection_destination_invalid")
    except OSError:
        if source == destination:
            return _protection_refusal(req, "protection_destination_invalid")

    if not source.is_file():
        return _protection_refusal(req, "protection_source_missing")
    if destination.exists() or destination.is_symlink():
        return _protection_refusal(req, "protection_destination_occupied")
    if not destination.parent.is_dir():
        return _protection_refusal(req, "protection_destination_invalid")

    try:
        actual_source_sha = _sha256_file(source)
    except OSError:
        return _protection_refusal(req, "protection_source_missing")

    if actual_source_sha != req.source_sha256:
        return _protection_refusal(req, "protection_source_hash_mismatch")

    baseline = inspect_docx(source)
    if not baseline.valid or baseline.file_sha256 != req.source_sha256:
        return _protection_refusal(req, "protection_package_invalid")

    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _protection_refusal(req, "protection_package_invalid")

    if _DOCUMENT not in graph.parts:
        return _protection_refusal(req, "protection_package_invalid")

    if req.part != _SETTINGS:
        return _protection_refusal(req, "protection_part_unsupported")

    actual_fp = protection_structure_fingerprint(source, part=req.part)
    if actual_fp != req.structure_fingerprint_sha256:
        return _protection_refusal(req, "protection_structure_fingerprint_mismatch")

    replacements: dict[str, bytes] = {}
    additions: dict[str, bytes] = {}

    if _SETTINGS not in graph.parts:
        # Settings part absent
        if req.action == "remove":
            return _protection_refusal(req, "protection_not_found")

        # action == "set": create settings.xml, document relationship, content type override
        try:
            doc_root = ET.fromstring(graph.read_part_bytes(_DOCUMENT))
        except (ET.ParseError, PackageError):
            return _protection_refusal(req, "protection_package_invalid")

        ns = _word_namespace(doc_root) or _W

        # Create new word/settings.xml
        settings_root = ET.Element(Q(ns, "settings"))
        dp_elem = ET.SubElement(settings_root, Q(ns, "documentProtection"))
        dp_elem.set(Q(ns, "edit"), req.edit)  # type: ignore[arg-type]
        dp_elem.set(Q(ns, "enforcement"), "1" if req.enforcement else "0")
        if req.formatting:
            dp_elem.set(Q(ns, "formatting"), "1")

        new_settings_bytes = serialize_xml_with_preserved_qnames(settings_root)
        additions[_SETTINGS] = new_settings_bytes

        # Load document relationships to add relationship for settings
        rels_name = _rels_name(_DOCUMENT)
        if rels_name in graph.parts:
            try:
                original_rels_bytes = graph.read_part_bytes(rels_name)
                rels_root = parse_xml_with_preserved_qnames(original_rels_bytes)
            except (ET.ParseError, PackageError):
                return _protection_refusal(req, "protection_package_invalid")
        else:
            original_rels_bytes = b""
            rels_root = ET.Element(Q(_PR, "Relationships"))

        settings_rel_id = _new_rel_id(rels_root)
        rels_root.append(
            ET.Element(
                Q(_PR, "Relationship"),
                {
                    "Id": settings_rel_id,
                    "Type": f"{_R}/settings",
                    "Target": "settings.xml",
                },
            )
        )
        replacements[rels_name] = serialize_xml_with_preserved_qnames(rels_root)

        # Load [Content_Types].xml to add override for word/settings.xml
        if "[Content_Types].xml" not in graph.parts:
            return _protection_refusal(req, "protection_package_invalid")
        try:
            original_types_bytes = graph.read_part_bytes("[Content_Types].xml")
            types_root = parse_xml_with_preserved_qnames(original_types_bytes)
        except (ET.ParseError, PackageError):
            return _protection_refusal(req, "protection_package_invalid")

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
        replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types_root)

    else:
        # Settings part present
        try:
            original_settings_bytes = graph.read_part_bytes(_SETTINGS)
            settings_root = parse_xml_with_preserved_qnames(original_settings_bytes)
        except (ET.ParseError, PackageError):
            return _protection_refusal(req, "protection_settings_invalid")

        local_root_name = settings_root.tag.rsplit("}", 1)[-1]
        if local_root_name != "settings":
            return _protection_refusal(req, "protection_settings_invalid")

        settings_ns = _word_namespace(settings_root) or _W
        existing_protects = settings_root.findall(Q(settings_ns, "documentProtection"))

        if len(existing_protects) > 1:
            return _protection_refusal(req, "protection_structure_unsupported")

        if existing_protects:
            existing_protect = existing_protects[0]
            # Check for unsupported child elements inside documentProtection
            if len(existing_protect) > 0:
                return _protection_refusal(req, "protection_structure_unsupported")

            # Check for cryptographic / password metadata on existing element
            has_crypto = any(
                k.rsplit("}", 1)[-1] in _CRYPT_ATTRIBUTES for k in existing_protect.attrib
            )
            if has_crypto:
                return _protection_refusal(req, "protection_cryptographic_metadata_present")

        if req.action == "remove":
            if not existing_protects:
                return _protection_refusal(req, "protection_not_found")
            settings_root.remove(existing_protects[0])
            replacements[_SETTINGS] = serialize_xml_with_preserved_qnames(settings_root)

        elif req.action == "set":
            if existing_protects:
                # Update existing element in place
                target_elem = existing_protects[0]
                target_elem.attrib.clear()
                target_elem.set(Q(settings_ns, "edit"), req.edit)  # type: ignore[arg-type]
                target_elem.set(Q(settings_ns, "enforcement"), "1" if req.enforcement else "0")
                if req.formatting:
                    target_elem.set(Q(settings_ns, "formatting"), "1")
            else:
                # Insert new element maintaining strict schema order
                new_dp = ET.Element(Q(settings_ns, "documentProtection"))
                new_dp.set(Q(settings_ns, "edit"), req.edit)  # type: ignore[arg-type]
                new_dp.set(Q(settings_ns, "enforcement"), "1" if req.enforcement else "0")
                if req.formatting:
                    new_dp.set(Q(settings_ns, "formatting"), "1")
                _insert_settings_child(settings_root, new_dp, settings_ns)

            replacements[_SETTINGS] = serialize_xml_with_preserved_qnames(settings_root)

    added_ranges: list[dict[str, Any]] = []
    source_document = b""
    if req.permission_ranges:
        try:
            source_document = graph.read_part_bytes(_DOCUMENT)
            document_root = parse_xml_with_preserved_qnames(source_document)
        except (ET.ParseError, PackageError):
            return _protection_refusal(req, "protection_package_invalid")
        authored = _author_permission_ranges(document_root, req.permission_ranges)
        if isinstance(authored, str):
            return _protection_refusal(req, authored)
        added_ranges = authored
        replacements[_DOCUMENT] = serialize_xml_with_preserved_qnames(document_root)

    changed_parts = frozenset(replacements.keys()) | frozenset(additions.keys())

    def verify_staged(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, changed_parts):
            return "protection_unrelated_parts_altered"

        try:
            staged_graph = PackageGraph.from_file(staged)
        except (PackageError, OSError, zipfile.BadZipFile):
            return "protection_staged_verification_failed"

        if req.action == "set":
            if _SETTINGS not in staged_graph.parts:
                return "protection_staged_verification_failed"
            try:
                staged_settings = ET.fromstring(staged_graph.read_part_bytes(_SETTINGS))
            except (ET.ParseError, PackageError):
                return "protection_staged_verification_failed"

            staged_ns = _word_namespace(staged_settings) or _W
            dp = staged_settings.find(Q(staged_ns, "documentProtection"))
            if dp is None:
                return "protection_staged_verification_failed"

            dp_edit = dp.get(Q(staged_ns, "edit")) or dp.get("edit")
            if dp_edit != req.edit:
                return "protection_staged_verification_failed"

            dp_enf = dp.get(Q(staged_ns, "enforcement")) or dp.get("enforcement")
            expected_enf = "1" if req.enforcement else "0"
            if dp_enf != expected_enf:
                return "protection_staged_verification_failed"

            dp_fmt = dp.get(Q(staged_ns, "formatting")) or dp.get("formatting")
            if req.formatting:
                if dp_fmt != "1":
                    return "protection_staged_verification_failed"
            else:
                if dp_fmt is not None and dp_fmt.lower() in {"1", "true", "on"}:
                    return "protection_staged_verification_failed"

            # Check schema ordering
            children = list(staged_settings)
            dp_idx = children.index(dp)
            dp_order = _SETTINGS_CHILD_ORDER_MAP.get("documentProtection", 34)
            for i, child in enumerate(children):
                ch_local = child.tag.rsplit("}", 1)[-1]
                ch_order = _SETTINGS_CHILD_ORDER_MAP.get(ch_local, 999)
                if i < dp_idx and ch_order > dp_order:
                    return "protection_staged_verification_failed"
                if i > dp_idx and ch_order < dp_order:
                    return "protection_staged_verification_failed"

            if added_ranges:
                try:
                    staged_document = staged_graph.read_part_bytes(_DOCUMENT)
                except PackageError:
                    return "protection_staged_verification_failed"
                if not _permission_ranges_verified(source_document, staged_document, added_ranges):
                    return "protection_staged_verification_failed"

            # If settings was added, verify relationship and override
            if _SETTINGS in additions:
                rels_name = _rels_name(_DOCUMENT)
                if rels_name not in staged_graph.parts:
                    return "protection_staged_verification_failed"
                try:
                    s_rels = ET.fromstring(staged_graph.read_part_bytes(rels_name))
                    has_rel = any(
                        r.get("Target") == "settings.xml"
                        or (r.get("Type") or "").endswith("/settings")
                        for r in s_rels.findall(Q(_PR, "Relationship"))
                    )
                    if not has_rel:
                        return "protection_staged_verification_failed"
                except (ET.ParseError, PackageError):
                    return "protection_staged_verification_failed"

                if "[Content_Types].xml" not in staged_graph.parts:
                    return "protection_staged_verification_failed"
                try:
                    s_types = ET.fromstring(staged_graph.read_part_bytes("[Content_Types].xml"))
                    has_override = any(
                        ov.get("PartName") == f"/{_SETTINGS}"
                        for ov in s_types.findall(Q(_CT, "Override"))
                    )
                    if not has_override:
                        return "protection_staged_verification_failed"
                except (ET.ParseError, PackageError):
                    return "protection_staged_verification_failed"

        elif req.action == "remove":
            if _SETTINGS in staged_graph.parts:
                try:
                    staged_settings = ET.fromstring(staged_graph.read_part_bytes(_SETTINGS))
                    staged_ns = _word_namespace(staged_settings) or _W
                    if staged_settings.find(Q(staged_ns, "documentProtection")) is not None:
                        return "protection_staged_verification_failed"
                except (ET.ParseError, PackageError):
                    return "protection_staged_verification_failed"

        return None

    published_report = _publish(
        source=source,
        output=destination,
        source_sha256=req.source_sha256,
        replacements=replacements,
        additions=additions,
        revisions=[],
        targets=[req.target_dict()],
        verify=verify_staged,
        failure_code="protection_staged_verification_failed",
        exists_code="protection_destination_occupied",
        unavailable_code="protection_publication_unavailable",
    )

    if published_report.status != _PUBLISHED:
        return ProtectionReport(
            status=_REFUSED,
            code=published_report.code,
            source_sha256=req.source_sha256,
            destination_sha256=None,
            target=req.target_dict(),
            protection=None,
            target_change_report=None,
            non_target_parts_preserved=None,
        )

    target_change_report = {
        "action": req.action,
        "part": req.part,
    }
    if req.action == "set":
        target_change_report["edit"] = req.edit
        target_change_report["enforcement"] = req.enforcement
        target_change_report["formatting"] = req.formatting
        protection_dict: dict[str, Any] | None = {
            "edit": req.edit,
            "enforcement": req.enforcement,
            "formatting": req.formatting,
        }
        if added_ranges:
            target_change_report["permission_ranges_added"] = len(added_ranges)
            protection_dict["permission_ranges"] = added_ranges  # type: ignore[index]
    else:
        protection_dict = None

    return ProtectionReport(
        status=_PUBLISHED,
        code=None,
        source_sha256=req.source_sha256,
        destination_sha256=published_report.output_sha256,
        target=req.target_dict(),
        protection=protection_dict,
        target_change_report=target_change_report,
        non_target_parts_preserved=True,
    )


@dataclass(frozen=True)
class HyphenationRequest:
    """Source-bound authoring request for document-wide automatic hyphenation."""

    source: Path
    destination: Path
    source_sha256: str
    structure_fingerprint_sha256: str
    auto_hyphenation: bool
    hyphenation_zone: int
    part: str = _SETTINGS

    def __post_init__(self) -> None:
        object.__setattr__(self, "source", Path(self.source))
        object.__setattr__(self, "destination", Path(self.destination))
        object.__setattr__(
            self,
            "source_sha256",
            _validate_sha256(self.source_sha256, "hyphenation_source_sha256"),
        )
        object.__setattr__(
            self,
            "structure_fingerprint_sha256",
            _validate_sha256(
                self.structure_fingerprint_sha256,
                "hyphenation_structure_fingerprint_sha256",
            ),
        )
        if not isinstance(self.auto_hyphenation, bool):
            raise TypeError("hyphenation_auto_invalid")
        if not isinstance(self.hyphenation_zone, int):
            raise TypeError("hyphenation_zone_invalid")
        if self.hyphenation_zone < 0:
            raise ValueError("hyphenation_zone_invalid")
        if self.part != _SETTINGS:
            raise ValueError("hyphenation_part_unsupported")

    def target_dict(self) -> dict[str, Any]:
        return {"part": self.part, "settings": ["autoHyphenation", "hyphenationZone"]}


@dataclass(frozen=True)
class HyphenationReport:
    """Content-safe result for automatic-hyphenation settings authoring."""

    status: str
    code: str | None
    source_sha256: str
    destination_sha256: str | None = None
    settings: dict[str, Any] | None = None
    non_target_parts_preserved: bool | None = None

    @property
    def passed(self) -> bool:
        return self.status == _PUBLISHED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "hyphenation",
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "source_sha256": self.source_sha256,
            "destination_sha256": self.destination_sha256,
            "settings": self.settings,
            "non_target_parts_preserved": self.non_target_parts_preserved,
        }


def hyphenation_structure_fingerprint(source: str | Path, *, part: str = _SETTINGS) -> str | None:
    """Return the source-bound settings fingerprint used by hyphenation authoring."""

    return protection_structure_fingerprint(source, part=part)


def _hyphenation_refusal(request: HyphenationRequest, code: str) -> HyphenationReport:
    return HyphenationReport(_REFUSED, code, request.source_sha256)


def author_hyphenation(
    request: HyphenationRequest | Mapping[str, Any],
) -> HyphenationReport:
    """Atomically author ``w:autoHyphenation`` and ``w:hyphenationZone``."""

    if isinstance(request, Mapping):
        try:
            request = HyphenationRequest(
                source=Path(request["source"]),
                destination=Path(request["destination"]),
                source_sha256=str(request["source_sha256"]),
                structure_fingerprint_sha256=str(request["structure_fingerprint_sha256"]),
                auto_hyphenation=request["auto_hyphenation"],
                hyphenation_zone=request["hyphenation_zone"],
                part=str(request.get("part", _SETTINGS)),
            )
        except (KeyError, TypeError, ValueError):
            return HyphenationReport(_REFUSED, "hyphenation_request_invalid", "")
    source, destination = request.source, request.destination
    try:
        if source.resolve() == destination.resolve():
            return _hyphenation_refusal(request, "hyphenation_destination_invalid")
    except OSError:
        return _hyphenation_refusal(request, "hyphenation_destination_invalid")
    if not source.is_file():
        return _hyphenation_refusal(request, "hyphenation_source_missing")
    if destination.exists() or destination.is_symlink() or not destination.parent.is_dir():
        return _hyphenation_refusal(request, "hyphenation_destination_invalid")
    if _sha256_file(source) != request.source_sha256:
        return _hyphenation_refusal(request, "hyphenation_source_hash_mismatch")
    if not inspect_docx(source).valid:
        return _hyphenation_refusal(request, "hyphenation_package_invalid")
    try:
        graph = PackageGraph.from_file(source)
    except (PackageError, OSError, zipfile.BadZipFile):
        return _hyphenation_refusal(request, "hyphenation_package_invalid")
    if (
        _DOCUMENT not in graph.parts
        or hyphenation_structure_fingerprint(source, part=request.part)
        != request.structure_fingerprint_sha256
    ):
        return _hyphenation_refusal(request, "hyphenation_structure_fingerprint_mismatch")

    additions: dict[str, bytes] = {}
    replacements: dict[str, bytes] = {}
    original_settings: bytes | None = None
    if _SETTINGS in graph.parts:
        try:
            original_settings = graph.read_part_bytes(_SETTINGS)
            root = parse_xml_with_preserved_qnames(original_settings)
        except (PackageError, ET.ParseError):
            return _hyphenation_refusal(request, "hyphenation_settings_invalid")
        if root.tag.rsplit("}", 1)[-1] != "settings":
            return _hyphenation_refusal(request, "hyphenation_settings_invalid")
    else:
        try:
            root = ET.Element(
                Q(
                    _word_namespace(ET.fromstring(graph.read_part_bytes(_DOCUMENT))) or _W,
                    "settings",
                )
            )
        except (PackageError, ET.ParseError):
            return _hyphenation_refusal(request, "hyphenation_package_invalid")
    namespace = _word_namespace(root) or _W
    values = {
        "autoHyphenation": "1" if request.auto_hyphenation else "0",
        "hyphenationZone": str(request.hyphenation_zone),
    }
    for name, value in values.items():
        matches = root.findall(Q(namespace, name))
        if len(matches) > 1:
            return _hyphenation_refusal(request, "hyphenation_structure_unsupported")
        if matches:
            element = matches[0]
            if len(element):
                return _hyphenation_refusal(request, "hyphenation_structure_unsupported")
            element.attrib.clear()
        else:
            element = ET.Element(Q(namespace, name))
            _insert_settings_child(root, element, namespace)
        element.set(Q(namespace, "val"), value)
    settings_bytes = serialize_xml_with_preserved_qnames(root)
    if original_settings is None:
        additions[_SETTINGS] = settings_bytes
        rels_name = _rels_name(_DOCUMENT)
        try:
            rels_original = graph.read_part_bytes(rels_name) if rels_name in graph.parts else b""
            rels = (
                parse_xml_with_preserved_qnames(rels_original)
                if rels_original
                else ET.Element(Q(_PR, "Relationships"))
            )
            types_original = graph.read_part_bytes("[Content_Types].xml")
            types = parse_xml_with_preserved_qnames(types_original)
        except (PackageError, ET.ParseError):
            return _hyphenation_refusal(request, "hyphenation_package_invalid")
        rels.append(
            ET.Element(
                Q(_PR, "Relationship"),
                {
                    "Id": _new_rel_id(rels),
                    "Type": f"{_R}/settings",
                    "Target": "settings.xml",
                },
            )
        )
        replacements[rels_name] = serialize_xml_with_preserved_qnames(rels)
        if not any(
            node.get("PartName") == f"/{_SETTINGS}" for node in types.findall(Q(_CT, "Override"))
        ):
            types.append(
                ET.Element(
                    Q(_CT, "Override"),
                    {
                        "PartName": f"/{_SETTINGS}",
                        "ContentType": _SETTINGS_CONTENT_TYPE,
                    },
                )
            )
        replacements["[Content_Types].xml"] = serialize_xml_with_preserved_qnames(types)
    else:
        replacements[_SETTINGS] = settings_bytes
    changed = frozenset(additions) | frozenset(replacements)

    def verify(staged: Path) -> str | None:
        if not _unrelated_parts_preserved(source, staged, changed):
            return "hyphenation_unrelated_parts_altered"
        try:
            root = ET.fromstring(PackageGraph.from_file(staged).read_part_bytes(_SETTINGS))
        except (PackageError, OSError, zipfile.BadZipFile, ET.ParseError):
            return "hyphenation_staged_verification_failed"
        namespace = _word_namespace(root) or _W
        for name, value in values.items():
            elements = root.findall(Q(namespace, name))
            if (
                len(elements) != 1
                or (elements[0].get(Q(namespace, "val")) or elements[0].get("val")) != value
            ):
                return "hyphenation_staged_verification_failed"
        return None

    published = _publish(
        source=source,
        output=destination,
        source_sha256=request.source_sha256,
        replacements=replacements,
        additions=additions,
        revisions=[],
        targets=[request.target_dict()],
        verify=verify,
        failure_code="hyphenation_staged_verification_failed",
        exists_code="hyphenation_destination_occupied",
        unavailable_code="hyphenation_publication_unavailable",
    )
    if published.status != _PUBLISHED:
        return _hyphenation_refusal(
            request, published.code or "hyphenation_publication_unavailable"
        )
    return HyphenationReport(
        _PUBLISHED,
        None,
        request.source_sha256,
        published.output_sha256,
        {
            "auto_hyphenation": request.auto_hyphenation,
            "hyphenation_zone": request.hyphenation_zone,
        },
        True,
    )
