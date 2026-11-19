"""One load-options object for every DOCX open and import path.

:class:`LoadOptions` carries the password, the text encoding, the recovery mode,
a progress callback, the editing-language preference, and the plain-text and HTML
import switches.  The same object is accepted by :class:`wolfdocx.Document`,
:meth:`wolfdocx.Document.from_bytes`, :func:`wolfdocx.txt_import.plaintext_to_docx`,
:func:`wolfdocx.txt_import.parse_plaintext`, :func:`wolfdocx.html_import.insert_html`,
:func:`wolfdocx.html_import.html_to_wordprocessingml`, and :func:`load_document`,
which detects the source format and routes every other importer (RTF, ODT, MHTML,
legacy DOC, Word 2003 WordML, PDF, Flat OPC and Markdown) through the same object.
Container-only ``inspect_format`` reports template/macro package families,
security flags and bounded DOC FIB metadata; ``detect_format`` retains the
string-returning contract. ``load_document`` keeps those family names in its
load receipt instead of collapsing them to DOCX/text.

Recovery has two modes.  ``refuse`` admits a package only when it is whole: a
readable ZIP whose members pass their CRC, a resolvable relationship graph, a
WordprocessingML main part, well-formed XML in every XML part, and no relationship
reference that names a missing relationship.  ``best_effort`` repairs what it can
and records every repair as a :class:`RecoveryAction` on the :class:`LoadReport`;
a package it cannot repair into one that passes the same admission is refused
with a named code.  A recovered session carries its actions into every save
report, so a repaired document is never published as if it were whole.

For diagnosis and byte-preserving edits of invalid parts, explicitly request
``defer_part_validation=True``. ZIP/member integrity and package graph syntax
are admitted at open; internal target existence and part XML are checked when
consumed or changed. This opt-in performs no repairs and cannot be combined with
``best_effort``. Default callers retain full refuse-mode admission.

Passwords are used only in memory to decrypt an Agile-encrypted package through
the qualified optional provider; no report, representation, or error repeats one.
"""

from __future__ import annotations

import codecs
import hashlib
import io
import re
import struct
import unicodedata
import zipfile
import zlib
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field, replace
from enum import StrEnum
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal
from urllib.parse import urljoin
from xml.etree import ElementTree as ET
from xml.parsers import expat

from ._optional import feature_module
from .admission import (
    _CFB_SIGNATURE,
    _MAIN_CONTENT_TYPES,
    _admit,
    _Damage,
    _main_part_name,
)
from .canonical import sha256_file
from .inspect import (
    PackageLimits,
    _parse_bounded_xml,
    _safe_member_name,
    _XmlSafetyError,
)
from .package import PackageError, PackageGraph

if TYPE_CHECKING:
    _feature_builder = feature_module("builder")
    DocumentBuilder = _feature_builder.DocumentBuilder
    DocumentSaveReport = _feature_builder.DocumentSaveReport
    from .document import Document

    _feature_font_sources = feature_module("font_sources")
    FontSettings = _feature_font_sources.FontSettings
    FontSource = _feature_font_sources.FontSource
    _feature_legacy_doc = feature_module("legacy_doc")
    _CompoundFile = _feature_legacy_doc._CompoundFile
    _feature_native_layout = feature_module("native_layout")
    LayoutOptions = _feature_native_layout.LayoutOptions
    from .render_warnings import RenderWarning

__all__ = [
    "BlockImportMode",
    "EditingLanguage",
    "FormatDetection",
    "LanguagePreferences",
    "LeadingSpaces",
    "LoadError",
    "LoadOptions",
    "LoadProgress",
    "LoadReport",
    "LoadedPackage",
    "ProcessorContext",
    "RecoveryAction",
    "RecoveryMode",
    "ResourceRequest",
    "TextDirection",
    "TrailingSpaces",
    "detect_format",
    "detect_text_encoding",
    "inspect_format",
    "load_document",
    "load_package",
]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_OFFICE_DOCUMENT = f"{_R}/officeDocument"
_STYLES_RELATIONSHIP = f"{_R}/styles"
_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
_DEFAULT_CONTENT_TYPES = {
    "rels": _RELATIONSHIPS_CONTENT_TYPE,
    "xml": "application/xml",
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "svg": "image/svg+xml",
    "odttf": "application/vnd.openxmlformats-officedocument.obfuscatedFont",
}
# Word parts a rebuilt package can name again: part name, content type, and the
# relationship type that attaches the part to the main document (None for package-
# level parts).
_KNOWN_PARTS: dict[str, tuple[str, str | None]] = {
    "word/document.xml": (f"{_WML}.document.main+xml", None),
    "word/styles.xml": (f"{_WML}.styles+xml", _STYLES_RELATIONSHIP),
    "word/settings.xml": (f"{_WML}.settings+xml", f"{_R}/settings"),
    "word/webSettings.xml": (f"{_WML}.webSettings+xml", f"{_R}/webSettings"),
    "word/fontTable.xml": (f"{_WML}.fontTable+xml", f"{_R}/fontTable"),
    "word/numbering.xml": (f"{_WML}.numbering+xml", f"{_R}/numbering"),
    "word/footnotes.xml": (f"{_WML}.footnotes+xml", f"{_R}/footnotes"),
    "word/endnotes.xml": (f"{_WML}.endnotes+xml", f"{_R}/endnotes"),
    "word/comments.xml": (f"{_WML}.comments+xml", f"{_R}/comments"),
    "word/theme/theme1.xml": (
        "application/vnd.openxmlformats-officedocument.theme+xml",
        f"{_R}/theme",
    ),
    "docProps/core.xml": (
        "application/vnd.openxmlformats-package.core-properties+xml",
        None,
    ),
    "docProps/app.xml": (
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        None,
    ),
}
_PACKAGE_LEVEL_RELATIONSHIPS = {
    "docProps/core.xml": (
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    ),
    "docProps/app.xml": f"{_R}/extended-properties",
}
_HEADER_FOOTER = re.compile(r"word/(header|footer)[0-9]+\.xml")
_LANGUAGE_TAG = re.compile(r"[A-Za-z]{2,3}(?:-[A-Za-z0-9]{1,8}){0,3}")
_ENCRYPTED_PACKAGE_NAME = "EncryptedPackage".encode("utf-16-le")
_LOCAL_HEADER = b"PK\x03\x04"
_LOCAL_HEADER_STRUCT = struct.Struct("<4sHHHHHIIIHH")
_DATA_DESCRIPTOR = b"PK\x07\x08"
_FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
_UNSUPPORTED_REFERENCE_PARENTS = frozenset({"drawing", "pict", "object"})


class RecoveryMode(StrEnum):
    """What an open does with a package that does not pass admission."""

    REFUSE = "refuse"
    BEST_EFFORT = "best_effort"


class LeadingSpaces(StrEnum):
    """How plain-text import treats spaces at the start of a line."""

    PRESERVE = "preserve"
    TRIM = "trim"
    CONVERT_TO_INDENT = "convert_to_indent"
    CONVERT_TO_TABS = "convert_to_tabs"


class TrailingSpaces(StrEnum):
    """How plain-text import treats spaces at the end of a line."""

    PRESERVE = "preserve"
    TRIM = "trim"


class TextDirection(StrEnum):
    """Fallback paragraph direction on load and during plain-text import."""

    LEFT_TO_RIGHT = "ltr"
    RIGHT_TO_LEFT = "rtl"
    AUTO = "auto"


class BlockImportMode(StrEnum):
    """How HTML import represents ``div`` blocks that contain other blocks.

    ``merge`` folds each enclosing block's indentation, border, background, and
    alignment into the paragraphs it contains.  ``preserve`` keeps the blocks as
    native Word HTML divisions (``w:divs`` in the web settings part) and links each
    contained paragraph to its innermost division with ``w:divId``.
    """

    MERGE = "merge"
    PRESERVE = "preserve"


@dataclass(frozen=True)
class LoadProgress:
    """One progress notification: ``stage`` reached, ``completed`` of ``total``."""

    stage: str
    completed: int
    total: int

    @property
    def fraction(self) -> float:
        return self.completed / self.total if self.total else 1.0


@dataclass(frozen=True)
class RecoveryAction:
    """One repair a best-effort load made, named and located but content-free."""

    code: str
    target: str | None = None

    def __str__(self) -> str:
        return self.code if self.target is None else f"{self.code}:{self.target}"


def _enum(kind: type[StrEnum], value: Any, code: str) -> Any:
    try:
        return kind(value)
    except ValueError:
        raise ValueError(code) from None


@dataclass(frozen=True)
class ResourceRequest:
    """Caller-resolved resource request. The engine never dereferences its URI."""

    uri: str
    kind: str


@dataclass(frozen=True)
class EditingLanguage:
    """An editing language and its priority; equal priorities retain caller order."""

    tag: str
    priority: int = 0

    def __post_init__(self) -> None:
        if not isinstance(self.tag, str) or _LANGUAGE_TAG.fullmatch(self.tag) is None:
            raise ValueError("load_editing_language_invalid")
        if type(self.priority) is not int:
            raise ValueError("load_language_priority_invalid")


@dataclass(frozen=True)
class LanguagePreferences:
    """Ordered language candidates and explicit Latin, bidi and East Asian defaults.

    Source declarations and run-local attributes take precedence. The highest
    priority candidate supplies the Latin default unless ``latin`` is specified.
    """

    languages: tuple[EditingLanguage, ...] = ()
    latin: str | None = None
    bidi: str | None = None
    east_asia: str | None = None

    def __post_init__(self) -> None:
        languages = tuple(self.languages)
        if any(not isinstance(language, EditingLanguage) for language in languages):
            raise TypeError("load_language_preferences_invalid")
        if len({language.tag.casefold() for language in languages}) != len(languages):
            raise ValueError("load_language_duplicate")
        for tag in (self.latin, self.bidi, self.east_asia):
            if tag is not None and (
                not isinstance(tag, str) or _LANGUAGE_TAG.fullmatch(tag) is None
            ):
                raise ValueError("load_editing_language_invalid")
        object.__setattr__(self, "languages", languages)

    @property
    def default(self) -> str | None:
        if self.latin is not None:
            return self.latin
        return (
            max(self.languages, key=lambda language: language.priority).tag
            if self.languages
            else None
        )


_WORD_VERSIONS = {
    "word2007": 12,
    "word2010": 14,
    "word2013": 15,
    "word2016": 15,
    "word2019": 15,
    "word2021": 15,
    "word2024": 15,
}


@dataclass(frozen=True)
class LoadOptions:
    """Options every open and import path accepts.

    ``password`` decrypts an Agile-encrypted package in memory; it is ignored for an
    unencrypted source and never appears in a representation or report.
    ``encoding`` names the text encoding of plain-text, HTML and Markdown byte
    sources and the default RTF encoding (source controls take precedence).
    DOCX parts declare their own encoding, so a package open ignores it.
    ``recovery`` is ``"refuse"`` (default) or ``"best_effort"``.
    ``defer_part_validation`` explicitly admits the ZIP and package graph for
    diagnosis or opaque preservation, deferring internal target existence and
    part XML checks until consumption. It cannot be combined with recovery.
    Archive, member integrity and bounded parsing checks still apply.
    ``progress`` receives :class:`LoadProgress` notifications; an exception it raises
    cancels the load, which is then refused with ``load_cancelled``.
    ``editing_language`` is a BCP 47 tag used when the source declares no default
    language: imported and loaded runs without an explicit language are tagged with
    it, including runs created by subsequent session edits.
    ``text_direction`` supplies a fallback direction for loaded paragraphs without
    explicit ``w:bidi``. The default ``ltr`` leaves DOCX formatting untouched;
    ``rtl`` and ``auto`` also apply to subsequent session edits.
    ``leading_spaces`` and ``trailing_spaces`` govern plain-text import;
    ``block_import_mode`` governs HTML import. ``pdf_pages`` is an ordered,
    one-based page selection for PDF imports, with the same boundaries as
    :func:`wolfdocx.pdf_import.pdf_to_docx`.
    ``legacy_encoding`` is an explicit fallback after UTF detection. HTML accepts
    ``html_css`` and ``control_representation``; ``base_uri`` resolves relative
    stylesheet/image references, and only ``resource_loading(ResourceRequest)``
    may supply their bytes. Unsupported CSS has named diagnostics. TXT accepts
    ``detect_lists`` and ``leading_spaces="convert_to_tabs"`` with ``tab_width``.
    Markdown accepts ``markdown_import_underline``,
    ``markdown_preserve_empty_lines`` and ``markdown_soft_line_break``.
    PDF accepts ``pdf_skip_images``, zero-based ``pdf_page_index`` plus optional
    ``pdf_page_count``, or the existing one-based ``pdf_pages`` (never both), and
    ``pdf_use_crop_box``. Import-specific progress accompanies shared stages.
    ``temp_folder`` selects private snapshots and load/render staging without
    changing process-wide tempfile settings. ``font_settings`` supplies the
    session's default render policy; ``warning_callback`` receives immutable load
    and render diagnostics. ``update_dirty_fields`` refreshes dirty unlocked
    fields and refuses unsupported evaluations. ``shape_to_math`` promotes native
    DrawingML equations and refuses binary equations. Disabling
    ``preserve_include_picture`` removes unlocked field instructions but keeps
    cached drawings. ``word_version`` selects the Word compatibility mode.
    ``language_preferences`` supplies prioritized language candidates and Latin,
    bidi and East Asian defaults, preserving explicit source formatting.
    """

    password: str | None = field(default=None, repr=False)
    encoding: str | None = None
    recovery: RecoveryMode = RecoveryMode.REFUSE
    progress: Callable[[LoadProgress], None] | None = field(default=None, repr=False, compare=False)
    editing_language: str | None = None
    leading_spaces: LeadingSpaces = LeadingSpaces.PRESERVE
    trailing_spaces: TrailingSpaces = TrailingSpaces.PRESERVE
    text_direction: TextDirection = TextDirection.LEFT_TO_RIGHT
    block_import_mode: BlockImportMode = BlockImportMode.MERGE
    defer_part_validation: bool = False
    pdf_pages: tuple[int, ...] | range | None = None
    legacy_encoding: str | None = None
    base_uri: str | None = None
    resource_loading: Callable[[ResourceRequest], bytes | None] | None = field(
        default=None, repr=False, compare=False
    )
    html_css: str = ""
    control_representation: str = "structured_tags"
    detect_lists: bool = False
    tab_width: int = 4
    markdown_import_underline: bool = False
    markdown_preserve_empty_lines: bool = False
    markdown_soft_line_break: str | None = None
    pdf_skip_images: bool = False
    pdf_page_index: int = 0
    pdf_page_count: int | None = None
    pdf_use_crop_box: bool = False
    update_dirty_fields: bool = False
    temp_folder: str | Path | None = None
    font_settings: FontSettings | None = field(default=None, repr=False, compare=False)
    warning_callback: Callable[[RenderWarning], None] | None = field(
        default=None, repr=False, compare=False
    )
    shape_to_math: bool = False
    preserve_include_picture: bool = True
    word_version: str | None = None
    language_preferences: LanguagePreferences | None = None

    def __post_init__(self) -> None:
        if self.password is not None and (not isinstance(self.password, str) or not self.password):
            raise ValueError("load_password_invalid")
        if self.encoding is not None and (
            not isinstance(self.encoding, str) or not self.encoding.strip()
        ):
            raise ValueError("load_encoding_invalid")
        if self.progress is not None and not callable(self.progress):
            raise ValueError("load_progress_invalid")
        if self.editing_language is not None and (
            not isinstance(self.editing_language, str)
            or _LANGUAGE_TAG.fullmatch(self.editing_language) is None
        ):
            raise ValueError("load_editing_language_invalid")
        for name, kind in (
            ("recovery", RecoveryMode),
            ("leading_spaces", LeadingSpaces),
            ("trailing_spaces", TrailingSpaces),
            ("text_direction", TextDirection),
            ("block_import_mode", BlockImportMode),
        ):
            object.__setattr__(self, name, _enum(kind, getattr(self, name), f"load_{name}_invalid"))
        if not isinstance(self.defer_part_validation, bool):
            raise TypeError("load_defer_part_validation_invalid")
        if self.defer_part_validation and self.best_effort:
            raise ValueError("load_deferred_recovery_invalid")
        for name in (
            "detect_lists",
            "markdown_import_underline",
            "markdown_preserve_empty_lines",
            "pdf_skip_images",
            "pdf_use_crop_box",
            "update_dirty_fields",
            "shape_to_math",
            "preserve_include_picture",
        ):
            if type(getattr(self, name)) is not bool:
                raise TypeError(f"load_{name}_invalid")
        for name in ("resource_loading", "warning_callback"):
            callback = getattr(self, name)
            if callback is not None and not callable(callback):
                raise TypeError(f"load_{name}_invalid")
        if self.base_uri is not None and (
            not isinstance(self.base_uri, str) or any(ord(char) < 32 for char in self.base_uri)
        ):
            raise ValueError("load_base_uri_invalid")
        if not isinstance(self.html_css, str) or len(self.html_css) > 1024 * 1024:
            raise ValueError("load_html_css_invalid")
        if self.control_representation not in {"structured_tags", "legacy_form_fields"}:
            raise ValueError("load_control_representation_invalid")
        if type(self.tab_width) is not int or not 1 <= self.tab_width <= 32:
            raise ValueError("load_tab_width_invalid")
        if type(self.pdf_page_index) is not int or self.pdf_page_index < 0:
            raise ValueError("load_pdf_page_index_invalid")
        if self.pdf_page_count is not None and (
            type(self.pdf_page_count) is not int or self.pdf_page_count < 1
        ):
            raise ValueError("load_pdf_page_count_invalid")
        if self.markdown_soft_line_break is not None and (
            not isinstance(self.markdown_soft_line_break, str)
            or len(self.markdown_soft_line_break) != 1
            or not (
                self.markdown_soft_line_break in {"\n", "\t"}
                or 0x20 <= ord(self.markdown_soft_line_break) <= 0xD7FF
                or 0xE000 <= ord(self.markdown_soft_line_break) <= 0xFFFD
                or 0x10000 <= ord(self.markdown_soft_line_break) <= 0x10FFFF
            )
        ):
            raise ValueError("load_markdown_soft_line_break_invalid")
        if self.legacy_encoding is not None:
            try:
                codec = codecs.lookup(self.legacy_encoding)
            except (LookupError, TypeError):
                raise ValueError("load_legacy_encoding_invalid") from None
            if not getattr(codec, "_is_text_encoding", True):
                raise ValueError("load_legacy_encoding_invalid")
            object.__setattr__(self, "legacy_encoding", codec.name)
        if self.temp_folder is not None:
            if not isinstance(self.temp_folder, (str, Path)) or not str(self.temp_folder):
                raise ValueError("load_temp_folder_invalid")
            folder = Path(self.temp_folder).resolve()
            if not folder.is_dir():
                raise ValueError("load_temp_folder_missing")
            object.__setattr__(self, "temp_folder", folder)
        if self.font_settings is not None:
            _feature_font_sources = feature_module("font_sources")
            FontSettings = _feature_font_sources.FontSettings

            if not isinstance(self.font_settings, FontSettings):
                raise TypeError("load_font_settings_invalid")
            self.font_settings.render_options()
        if self.word_version is not None and self.word_version not in _WORD_VERSIONS:
            raise ValueError("load_word_version_invalid")
        if self.language_preferences is not None and not isinstance(
            self.language_preferences, LanguagePreferences
        ):
            raise TypeError("load_language_preferences_invalid")

    @property
    def best_effort(self) -> bool:
        return self.recovery is RecoveryMode.BEST_EFFORT

    def notify(self, stage: str, completed: int, total: int) -> None:
        """Report one stage; a callback exception becomes ``load_cancelled``."""
        if self.progress is None:
            return
        try:
            self.progress(LoadProgress(stage, completed, total))
        except Exception as exc:
            raise LoadError("load_cancelled", "The progress callback cancelled the load.") from exc

    @property
    def preferred_language(self) -> str | None:
        return self.editing_language or (
            self.language_preferences.default if self.language_preferences is not None else None
        )

    def load_resource(self, uri: str, kind: str) -> bytes:
        """Resolve through the caller only; reject missing, oversized or invalid bytes."""
        if not isinstance(uri, str) or not uri or any(ord(char) < 32 for char in uri):
            raise LoadError("load_resource_uri_invalid", "Resource URI is invalid.")
        if self.resource_loading is None:
            raise LoadError(
                "load_resource_callback_required", "A caller resource callback is required."
            )
        request = ResourceRequest(urljoin(self.base_uri or "", uri), kind)
        try:
            payload = self.resource_loading(request)
        except Exception as exc:
            raise LoadError(
                "load_resource_callback_failed", "Caller resource loading failed."
            ) from exc
        if payload is None:
            raise LoadError("load_resource_unavailable", "The caller did not supply the resource.")
        if not isinstance(payload, bytes):
            raise LoadError("load_resource_bytes_invalid", "Resource callbacks must return bytes.")
        if len(payload) > PackageLimits().max_entry_size:
            raise LoadError("load_resource_size_limit", "Caller resource exceeds the entry limit.")
        return payload

    def warn(self, code: str, target: str | None = None) -> None:
        """Deliver immutable, content-safe load diagnostics to an explicit caller."""
        if self.warning_callback is None:
            return
        from .render_warnings import RenderWarning, WarningCategory, WarningSource

        self.warning_callback(
            RenderWarning(
                WarningCategory.RECOVERY,
                WarningSource.LOAD,
                code,
                kind=code,
                locator=target,
                reason=code,
            )
        )


@dataclass(frozen=True)
class ProcessorContext:
    """Shared policies for conversion, assembly, splitting and template processing.

    ``load_options`` governs source admission and import. Font discovery is local
    and explicit unless ``use_system_fonts`` is requested; renderers consume the
    same font sources, substitutions and default family. ``layout_options`` supplies
    native layout policy, and ``warning_callback`` receives immutable warnings
    without removing them from operation receipts. ``cleanup`` contains mail-merge
    cleanup names and is combined with request-local cleanup. Reporting
    policies retain literal tags or substitute empty text for missing members and
    expose only safe public members of explicitly supplied ``known_types``.
    Defaults preserve each existing processor's behavior.
    """

    load_options: LoadOptions = field(default_factory=LoadOptions)
    font_sources: tuple[FontSource, ...] = ()
    font_substitutions: Mapping[str, tuple[str, ...]] | None = None
    default_font: str | None = None
    use_system_fonts: bool = False
    cleanup: frozenset[str] = frozenset()
    missing_member: Literal["error", "empty", "keep_tag"] = "error"
    known_types: tuple[type, ...] = ()
    layout_options: LayoutOptions | None = None
    warning_callback: Callable[[RenderWarning], None] | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.load_options, LoadOptions):
            raise TypeError("processor_load_options_invalid")
        if self.layout_options is not None:
            _feature_native_layout = feature_module("native_layout")
            LayoutOptions = _feature_native_layout.LayoutOptions

            if not isinstance(self.layout_options, LayoutOptions):
                raise TypeError("processor_layout_options_invalid")
        if self.warning_callback is not None and not callable(self.warning_callback):
            raise TypeError("processor_warning_callback_invalid")
        if not isinstance(self.use_system_fonts, bool):
            raise TypeError("processor_system_fonts_invalid")
        if self.default_font is not None and (
            not isinstance(self.default_font, str) or not self.default_font.strip()
        ):
            raise ValueError("native_render_default_font_invalid")
        if self.font_sources or self.font_substitutions is not None:
            _feature_font_sources = feature_module("font_sources")
            FontSource = _feature_font_sources.FontSource
            parse_substitutions = _feature_font_sources.parse_substitutions

            if any(not isinstance(source, FontSource) for source in self.font_sources):
                raise TypeError("native_render_font_source_invalid")
            if self.font_substitutions is not None:
                from types import MappingProxyType

                object.__setattr__(
                    self,
                    "font_substitutions",
                    MappingProxyType(parse_substitutions(self.font_substitutions)),
                )
        object.__setattr__(self, "font_sources", tuple(self.font_sources))
        cleanup = frozenset(self.cleanup)
        if cleanup - {
            "unused_fields",
            "unused_regions",
            "unused_region_markers",
            "empty_paragraphs",
            "containing_tables",
            "static_fields",
        }:
            raise ValueError("mail_merge_cleanup_invalid")
        object.__setattr__(self, "cleanup", cleanup)
        if self.missing_member not in {"error", "empty", "keep_tag"}:
            raise ValueError("linq_missing_member_option_invalid")
        if any(not isinstance(kind, type) for kind in self.known_types):
            raise TypeError("linq_known_types_invalid")
        object.__setattr__(self, "known_types", tuple(self.known_types))

    def notify_warnings(self, warnings: Iterable[RenderWarning]) -> None:
        """Deliver typed diagnostics without consuming an operation's receipt.

        Callback exceptions propagate unchanged. Warning records are immutable;
        callers may retain them independently of the operation report.
        """
        from .render_warnings import RenderWarning

        for warning in warnings:
            if not isinstance(warning, RenderWarning):
                raise TypeError("processor_warning_invalid")
            if self.warning_callback is not None:
                self.warning_callback(warning)


@dataclass(frozen=True)
class LoadReport:
    """Content-safe receipt for one open or import under :class:`LoadOptions`."""

    status: str
    code: str | None
    source_format: str
    source_sha256: str | None
    package_sha256: str | None
    encrypted: bool
    recovery_mode: str
    recovery: tuple[RecoveryAction, ...] = ()
    editing_language: str | None = None
    editing_language_source: str | None = None
    target: str | None = None

    @property
    def recovered(self) -> bool:
        return bool(self.recovery)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": "document-load",
            "status": self.status,
            "code": self.code,
            "target": self.target,
            "source_format": self.source_format,
            "source_sha256": self.source_sha256,
            "package_sha256": self.package_sha256,
            "encrypted": self.encrypted,
            "recovery_mode": self.recovery_mode,
            "recovery": [str(action) for action in self.recovery],
            "editing_language": self.editing_language,
            "editing_language_source": self.editing_language_source,
        }


class LoadError(PackageError):
    """A load refused with a named ``code``; ``report`` holds the refusal receipt."""

    def __init__(
        self,
        code: str,
        message: str,
        entry_name: str | None = None,
        *,
        report: LoadReport | None = None,
    ) -> None:
        super().__init__(code, message, entry_name)
        self.report = report


@dataclass(frozen=True)
class LoadedPackage:
    """The admitted package bytes or private disk snapshot, graph, and receipt."""

    data: bytes | Path
    graph: PackageGraph
    report: LoadReport


def load_package(
    data: bytes | Path,
    *,
    options: LoadOptions | None = None,
    limits: PackageLimits | None = None,
    path: str | Path = "document.docx",
) -> LoadedPackage:
    """Admit one DOCX package from bytes or a caller-owned private disk snapshot.

    Raises :class:`LoadError` with a named code when the package is refused.
    """
    options = options or LoadOptions()
    limits = limits or PackageLimits()
    source_sha256 = (
        sha256_file(data) if isinstance(data, Path) else hashlib.sha256(data).hexdigest()
    )
    mode = options.recovery.value

    def refuse(code: str, target: str | None = None, *, encrypted: bool = False) -> LoadError:
        report = LoadReport(
            "refused", code, "docx", source_sha256, None, encrypted, mode, target=target
        )
        return LoadError(code, f"Load refused: {code}.", target, report=report)

    options.notify("read", 1, 5)
    size = data.stat().st_size if isinstance(data, Path) else len(data)
    if size > limits.max_archive_size:
        raise refuse("archive_size_limit")
    if isinstance(data, Path):
        with data.open("rb") as stream:
            signature = stream.read(len(_CFB_SIGNATURE))
    else:
        signature = data[: len(_CFB_SIGNATURE)]
    encrypted = signature == _CFB_SIGNATURE
    if encrypted:
        if isinstance(data, Path):
            data = data.read_bytes()
        if _ENCRYPTED_PACKAGE_NAME not in data:
            raise refuse("load_format_unsupported")
        if options.password is None:
            raise refuse("load_password_required", encrypted=True)
        code = _decrypt(data, options.password, limits)
        if isinstance(code, str):
            raise refuse(code, encrypted=True)
        data = code
    options.notify("decrypt", 2, 5)
    actions: tuple[RecoveryAction, ...] = ()
    try:
        graph = _admit(data, limits, path, defer_parts=options.defer_part_validation)
    except _Damage as damage:
        if not options.best_effort:
            raise refuse(damage.code, damage.target, encrypted=encrypted) from None
        options.notify("recover", 3, 5)
        try:
            data, actions = _recover(data.read_bytes() if isinstance(data, Path) else data, limits)
        except _Damage as unrecoverable:
            raise refuse(unrecoverable.code, unrecoverable.target, encrypted=encrypted) from None
        try:
            graph = _admit(data, limits, path)
        except _Damage as residual:
            raise refuse(
                "recovery_unverified",
                f"{residual.code}:{residual.target}",
                encrypted=encrypted,
            ) from None
    options.notify("validate", 4, 5)
    declared = _declared_language(graph, limits)
    language = declared or options.preferred_language
    report = LoadReport(
        "recovered" if actions else "opened",
        None,
        "encrypted_docx" if encrypted else "docx",
        source_sha256,
        graph.file_sha256,
        encrypted,
        mode,
        actions,
        language,
        "document" if declared else ("preference" if language else None),
    )
    for action in actions:
        options.warn(action.code, action.target)
    options.notify("complete", 5, 5)
    return LoadedPackage(data, graph, report)


def _decrypt(data: bytes, password: str, limits: PackageLimits) -> bytes | str:
    _feature_security = feature_module("security")
    MSOFFCRYPTO_DISTRIBUTION = _feature_security.MSOFFCRYPTO_DISTRIBUTION
    MSOFFCRYPTO_VERSION = _feature_security.MSOFFCRYPTO_VERSION

    try:
        if version(MSOFFCRYPTO_DISTRIBUTION) != MSOFFCRYPTO_VERSION:
            return "load_decryption_provider_unqualified"
        provider = feature_module("_commercial_features")
        exceptions, OOXMLFile = provider.decryption_provider()
    except (PackageNotFoundError, ImportError):
        return "load_decryption_provider_missing"
    try:
        office = OOXMLFile(io.BytesIO(data))
        if not office.is_encrypted():
            return "load_format_unsupported"
        if office.type != "agile":
            return "load_encryption_scheme_unsupported"
        office.load_key(password=password, verify_password=True)
    except exceptions.InvalidKeyError:
        return "load_password_rejected"
    except exceptions.DecryptionError:
        return "load_encryption_scheme_unsupported"
    except Exception:  # noqa: BLE001 - untrusted container bytes are a refusal
        return "load_encrypted_container_invalid"
    plaintext = io.BytesIO()
    try:
        office.decrypt(plaintext, verify_integrity=True)
    except exceptions.InvalidKeyError:
        return "load_encrypted_integrity_failed"
    except Exception:  # noqa: BLE001 - provider messages may quote content
        return "load_encrypted_container_invalid"
    payload = plaintext.getvalue()
    if len(payload) > limits.max_archive_size:
        return "archive_size_limit"
    return payload


def _declared_language(graph: PackageGraph, limits: PackageLimits) -> str | None:
    main = _main_part_name(graph)
    if main is None:
        return None
    for relationship in graph.relationships_for(main):
        if relationship.type != _STYLES_RELATIONSHIP or relationship.is_external:
            continue
        target = relationship.resolved_target()
        if target not in graph.parts:
            return None
        try:
            root = _parse_bounded_xml(
                graph.read_part_bytes(target),
                limits=limits,
                total_elements=[0],
                preserve_qnames=False,
            )
        except (ET.ParseError, PackageError, _XmlSafetyError):
            return None
        lang = root.find(f"{{{_W}}}docDefaults/{{{_W}}}rPrDefault/{{{_W}}}rPr/{{{_W}}}lang")
        value = None if lang is None else lang.get(f"{{{_W}}}val")
        return value or None
    return None


# -- best-effort recovery ------------------------------------------------------------


def _recover(data: bytes, limits: PackageLimits) -> tuple[bytes, tuple[RecoveryAction, ...]]:
    actions: list[RecoveryAction] = []
    members = _salvage_members(data, limits, actions)
    if not members:
        raise _Damage("recovery_no_package_members")
    _repair_xml_parts(members, limits, actions)
    content_types = _repair_content_types(members, limits, actions)
    main = _repair_package_relationships(members, content_types, limits, actions)
    _repair_main_relationships(members, main, actions)
    _prune_relationships(members, limits, actions)
    _strip_dangling_references(members, limits, actions)
    return _rebuild(members, content_types), tuple(actions)


def _salvage_members(
    data: bytes, limits: PackageLimits, actions: list[RecoveryAction]
) -> dict[str, bytes]:
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except (zipfile.BadZipFile, OSError, ValueError, EOFError):
        archive = None
    if archive is None:
        members: dict[str, bytes] = {}
        found = False
        for name, payload in _scan_local_headers(data, limits, actions):
            found = True
            _admit_member(members, name, payload, limits, actions)
        if found:
            actions.insert(0, RecoveryAction("central_directory_rebuilt"))
        return members
    members = {}
    with archive:
        for info in archive.infolist()[: limits.max_entries]:
            if info.filename.endswith("/"):
                actions.append(RecoveryAction("member_dropped_directory", info.filename))
                continue
            if info.file_size > limits.max_entry_size:
                actions.append(RecoveryAction("member_dropped_size_limit", info.filename))
                continue
            try:
                payload = archive.read(info)
            except (
                zipfile.BadZipFile,
                zlib.error,
                EOFError,
                OSError,
                NotImplementedError,
            ):
                actions.append(RecoveryAction("member_dropped_corrupt", info.filename))
                continue
            _admit_member(members, info.filename, payload, limits, actions)
    return members


def _admit_member(
    members: dict[str, bytes],
    name: str,
    payload: bytes,
    limits: PackageLimits,
    actions: list[RecoveryAction],
) -> None:
    if not name or not _safe_member_name(name):
        actions.append(RecoveryAction("member_dropped_unsafe_path", None))
    elif name.endswith("/"):
        actions.append(RecoveryAction("member_dropped_directory", name))
    elif name in members:
        actions.append(RecoveryAction("member_dropped_duplicate", name))
    elif len(members) >= limits.max_entries:
        actions.append(RecoveryAction("member_dropped_entry_limit", name))
    else:
        members[name] = payload


def _scan_local_headers(
    data: bytes, limits: PackageLimits, actions: list[RecoveryAction]
) -> Iterator[tuple[str, bytes]]:
    """Recover members from local file headers when the central directory is gone."""
    offset = data.find(_LOCAL_HEADER)
    while offset != -1 and offset + _LOCAL_HEADER_STRUCT.size <= len(data):
        (
            _,
            _,
            flags,
            method,
            _,
            _,
            crc,
            compressed,
            _size,
            name_length,
            extra_length,
        ) = _LOCAL_HEADER_STRUCT.unpack_from(data, offset)
        start = offset + _LOCAL_HEADER_STRUCT.size + name_length + extra_length
        raw_name = data[offset + _LOCAL_HEADER_STRUCT.size : start - extra_length]
        name = raw_name.decode("utf-8" if flags & 0x800 else "cp437", errors="replace")
        payload, end, code = _inflate_member(data, start, flags, method, crc, compressed, limits)
        if code is not None:
            actions.append(RecoveryAction(code, name))
        else:
            assert payload is not None
            yield name, payload
        offset = data.find(_LOCAL_HEADER, max(end, offset + 4))


def _inflate_member(
    data: bytes,
    start: int,
    flags: int,
    method: int,
    crc: int,
    compressed: int,
    limits: PackageLimits,
) -> tuple[bytes | None, int, str | None]:
    descriptor = bool(flags & 0x08)
    if flags & 0x01:
        return None, start, "member_dropped_encrypted"
    if method == zipfile.ZIP_STORED:
        if descriptor and compressed == 0:
            return None, start, "member_dropped_truncated"
        end = start + compressed
        if end > len(data):
            return None, start, "member_dropped_truncated"
        payload = data[start:end]
    elif method == zipfile.ZIP_DEFLATED:
        inflater = zlib.decompressobj(-zlib.MAX_WBITS)
        try:
            payload = inflater.decompress(data[start:], limits.max_entry_size + 1)
        except zlib.error:
            return None, start, "member_dropped_corrupt"
        if len(payload) > limits.max_entry_size:
            return None, start, "member_dropped_size_limit"
        if not inflater.eof:
            return None, len(data), "member_dropped_truncated"
        end = len(data) - len(inflater.unused_data)
    else:
        return None, start, "member_dropped_compression_unsupported"
    if descriptor:
        tail = data[end : end + 16]
        if tail.startswith(_DATA_DESCRIPTOR):
            tail = tail[4:]
        if len(tail) >= 4:
            crc = struct.unpack_from("<I", tail)[0]
    if zlib.crc32(payload) != crc:
        return None, end, "member_dropped_crc_invalid"
    return payload, end, None


def _parse(data: bytes, limits: PackageLimits) -> ET.Element:
    return _parse_bounded_xml(data, limits=limits, total_elements=[0], preserve_qnames=True)


def _serialize(root: ET.Element) -> bytes:
    from .inspect import serialize_xml_with_preserved_qnames

    return serialize_xml_with_preserved_qnames(root)


def _repair_xml_parts(
    members: dict[str, bytes], limits: PackageLimits, actions: list[RecoveryAction]
) -> None:
    main_parts = {"word/document.xml"}
    if "[Content_Types].xml" in members:
        try:
            types = _parse(members["[Content_Types].xml"], limits)
            main_parts.update(
                node.get("PartName", "").lstrip("/")
                for node in types.findall(f"{{{_CT}}}Override")
                if node.get("ContentType") in _MAIN_CONTENT_TYPES
            )
        except ET.ParseError:
            pass
        except _XmlSafetyError as exc:
            raise _Damage(exc.code, "[Content_Types].xml") from None
    for name in list(members):
        if not name.endswith((".xml", ".rels")):
            continue
        try:
            _parse(members[name], limits)
            continue
        except ET.ParseError:
            pass
        except _XmlSafetyError as exc:
            # Recovery never bypasses parser safety limits or forbidden declarations.
            raise _Damage(exc.code, name) from None
        if name in main_parts:
            salvaged = _salvage_story(members[name])
            if salvaged is not None:
                payload, blocks = salvaged
                try:
                    _parse(payload, limits)
                except (ET.ParseError, _XmlSafetyError):
                    pass
                else:
                    members[name] = payload
                    actions.append(
                        RecoveryAction("main_document_truncated", f"{blocks}_blocks_kept")
                    )
                    continue
            raise _Damage("recovery_main_document_unreadable", name)
        repaired = _salvage_truncated_part(
            members[name], limits, relationships=name.endswith(".rels")
        )
        if repaired is not None:
            members[name] = repaired
            actions.append(
                RecoveryAction(
                    "relationships_malformed_xml_repaired"
                    if name.endswith(".rels")
                    else "part_truncated_xml_repaired",
                    name,
                )
            )
            continue
        # XML-forbidden literal characters can be replaced without changing its tree.
        try:
            text = members[name].decode("utf-8")
        except UnicodeError:
            text = ""
        if text:
            cleaned = "".join(
                char
                if (
                    ord(char) in {9, 10, 13}
                    or 0x20 <= ord(char) <= 0xD7FF
                    or 0xE000 <= ord(char) <= 0xFFFD
                    or 0x10000 <= ord(char) <= 0x10FFFF
                )
                else "\ufffd"
                for char in text
            )
            if cleaned != text:
                payload = cleaned.encode("utf-8")
                try:
                    _parse(payload, limits)
                except (ET.ParseError, _XmlSafetyError):
                    pass
                else:
                    members[name] = payload
                    actions.append(RecoveryAction("part_invalid_characters_replaced", name))
                    continue
        del members[name]
        actions.append(RecoveryAction("part_dropped_malformed_xml", name))


def _salvage_truncated_part(
    data: bytes, limits: PackageLimits, *, relationships: bool
) -> bytes | None:
    """Keep whole root children of a truncated ancillary part.

    No partial child is retained. Malformed relationship suffixes may be cut
    after a complete relationship; other XML admits only EOF truncation.
    """
    if data.startswith((b"\xff\xfe", b"\xfe\xff", b"<\0", b"\0<")):
        return None
    parser = expat.ParserCreate()
    stack: list[str] = []
    root_name = ""
    cut = 0
    children = 0

    def start(name: str, attrs: object) -> None:
        nonlocal root_name
        if not stack:
            root_name = name
        stack.append(name)

    def end(name: str) -> None:
        nonlocal cut, children
        if len(stack) == 2:
            position = parser.CurrentByteIndex
            if data[position : position + 2] == b"</":
                position = data.find(b">", position) + 1
            cut = position
            children += 1
        stack.pop()

    def forbidden(*args: Any) -> None:
        raise _Damage("xml_dtd_forbidden")

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = forbidden
    parser.EntityDeclHandler = forbidden
    try:
        parser.Parse(data, True)
        return None
    except expat.ExpatError as error:
        truncation = {
            expat.errors.codes[expat.errors.XML_ERROR_NO_ELEMENTS],
            expat.errors.codes[expat.errors.XML_ERROR_UNCLOSED_TOKEN],
            expat.errors.codes[expat.errors.XML_ERROR_PARTIAL_CHAR],
        }
        if error.code not in truncation and not relationships:
            return None
    if not root_name or not children or cut <= 0:
        return None
    repaired = data[:cut] + f"</{root_name}>".encode()
    try:
        root = _parse(repaired, limits)
    except (ET.ParseError, _XmlSafetyError):
        return None
    if relationships and root.tag != f"{{{_PR}}}Relationships":
        return None
    return repaired


def _salvage_story(data: bytes) -> tuple[bytes, int] | None:
    """Cut a malformed main story after its last whole body-level block."""
    parser = expat.ParserCreate()
    names: list[str] = []
    state = {"end": 0, "blocks": 0, "closed": 0}

    def start(name: str, _attrs: object) -> None:
        names.append(name)

    def end(name: str) -> None:
        depth = len(names)
        names.pop()
        if depth == 3:
            close = data.find(b">", parser.CurrentByteIndex)
            if close != -1:
                state["end"] = close + 1
                state["blocks"] += 1
        elif depth == 1:
            close = data.find(b">", parser.CurrentByteIndex)
            state["closed"] = close + 1

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except expat.ExpatError:
        pass
    if state["closed"]:
        return data[: state["closed"]], state["blocks"]
    if len(names) < 2 or not state["blocks"]:
        return None
    root, body = names[0], names[1]
    return data[: state["end"]] + f"</{body}></{root}>".encode(), state["blocks"]


def _repair_content_types(
    members: dict[str, bytes], limits: PackageLimits, actions: list[RecoveryAction]
) -> ET.Element:
    name = "[Content_Types].xml"
    original = members.get(name)
    root: ET.Element | None = None
    if name in members:
        try:
            root = _parse(members[name], limits)
        except (ET.ParseError, _XmlSafetyError):
            root = None
        if root is not None and root.tag != f"{{{_CT}}}Types":
            root = None
    original_tree = None if root is None else _serialize(root)
    if root is None:
        actions.append(RecoveryAction("content_types_rebuilt"))
        root = ET.Element(f"{{{_CT}}}Types")
    defaults = {
        node.get("Extension", "").lower(): node for node in root.findall(f"{{{_CT}}}Default")
    }
    for node in root.findall(f"{{{_CT}}}Override"):
        part = node.get("PartName", "").lstrip("/")
        if part not in members:
            root.remove(node)
            actions.append(RecoveryAction("content_type_override_dropped", part or None))
    overrides = {
        node.get("PartName", "").lstrip("/") for node in root.findall(f"{{{_CT}}}Override")
    }
    for extension in ("rels", "xml"):
        if extension not in defaults:
            defaults[extension] = ET.SubElement(
                root,
                f"{{{_CT}}}Default",
                {
                    "Extension": extension,
                    "ContentType": _DEFAULT_CONTENT_TYPES[extension],
                },
            )
    for part in list(members):
        if part == name or part in overrides:
            continue
        known = _known_content_type(part)
        if known is not None:
            ET.SubElement(
                root,
                f"{{{_CT}}}Override",
                {"PartName": f"/{part}", "ContentType": known},
            )
            continue
        extension = part.rsplit(".", 1)[-1].lower() if "." in part else ""
        if extension in defaults:
            continue
        if extension in _DEFAULT_CONTENT_TYPES:
            defaults[extension] = ET.SubElement(
                root,
                f"{{{_CT}}}Default",
                {
                    "Extension": extension,
                    "ContentType": _DEFAULT_CONTENT_TYPES[extension],
                },
            )
            continue
        del members[part]
        actions.append(RecoveryAction("part_dropped_content_type_unknown", part))
    # Overrides first named in a well-formed table are kept; only missing ones are added.
    if original is not None and original_tree == _serialize(root):
        members[name] = original
    else:
        members.pop(name, None)
    return root


def _known_content_type(part: str) -> str | None:
    known = _KNOWN_PARTS.get(part)
    if known is not None:
        return known[0]
    match = _HEADER_FOOTER.fullmatch(part)
    if match is not None:
        return f"{_WML}.{match.group(1)}+xml"
    return None


def _content_type(root: ET.Element, part: str) -> str | None:
    for node in root.findall(f"{{{_CT}}}Override"):
        if node.get("PartName", "").lstrip("/") == part:
            return node.get("ContentType")
    extension = part.rsplit(".", 1)[-1].lower()
    for node in root.findall(f"{{{_CT}}}Default"):
        if node.get("Extension", "").lower() == extension:
            return node.get("ContentType")
    return None


def _relationships(
    members: dict[str, bytes], name: str, limits: PackageLimits
) -> ET.Element | None:
    if name not in members:
        return None
    try:
        root = _parse(members[name], limits)
    except (ET.ParseError, _XmlSafetyError):
        return None
    return root if root.tag == f"{{{_PR}}}Relationships" else None


def _repair_package_relationships(
    members: dict[str, bytes],
    content_types: ET.Element,
    limits: PackageLimits,
    actions: list[RecoveryAction],
) -> str:
    name = "_rels/.rels"
    root = _relationships(members, name, limits)
    main: str | None = None
    if root is not None:
        for node in root.findall(f"{{{_PR}}}Relationship"):
            if node.get("Type") == _OFFICE_DOCUMENT and node.get("TargetMode") != "External":
                target = node.get("Target", "").lstrip("/")
                if (
                    target in members
                    and _content_type(content_types, target) in _MAIN_CONTENT_TYPES
                ):
                    main = target
    if main is None:
        main = next(
            (part for part in members if _content_type(content_types, part) in _MAIN_CONTENT_TYPES),
            None,
        )
        if main is None:
            raise _Damage("recovery_main_document_missing")
        root = ET.Element(f"{{{_PR}}}Relationships")
        ET.SubElement(
            root,
            f"{{{_PR}}}Relationship",
            {"Id": "rId1", "Type": _OFFICE_DOCUMENT, "Target": main},
        )
        for index, (part, relationship_type) in enumerate(_PACKAGE_LEVEL_RELATIONSHIPS.items(), 2):
            if part in members:
                ET.SubElement(
                    root,
                    f"{{{_PR}}}Relationship",
                    {"Id": f"rId{index}", "Type": relationship_type, "Target": part},
                )
        members[name] = _serialize(root)
        actions.append(RecoveryAction("package_relationships_rebuilt"))
    return main


def _repair_main_relationships(
    members: dict[str, bytes], main: str, actions: list[RecoveryAction]
) -> None:
    """Re-attach the known Word parts to a main story whose relationship part is lost.

    References the story made to parts that cannot be re-identified (images,
    headers, hyperlinks) are removed later as dangling references.
    """
    folder, _, leaf = main.rpartition("/")
    name = f"{folder}/_rels/{leaf}.rels" if folder else f"_rels/{leaf}.rels"
    if name in members:
        return
    prefix = f"{folder}/" if folder else ""
    known = [
        (part, relationship_type)
        for part, (_, relationship_type) in _KNOWN_PARTS.items()
        if relationship_type is not None and part in members and part.startswith(prefix)
    ]
    if not known:
        return
    root = ET.Element(f"{{{_PR}}}Relationships")
    for index, (part, relationship_type) in enumerate(known, 1):
        ET.SubElement(
            root,
            f"{{{_PR}}}Relationship",
            {
                "Id": f"rIdRecovered{index}",
                "Type": relationship_type,
                "Target": part[len(prefix) :],
            },
        )
    members[name] = _serialize(root)
    actions.append(RecoveryAction("relationships_rebuilt", name))


def _source_part(rels_name: str) -> str:
    folder, _, leaf = rels_name.rpartition("/")
    parent = folder.rpartition("/")[0] if folder.endswith("_rels") else folder
    if folder == "_rels":
        parent = ""
    source = leaf[: -len(".rels")]
    return f"{parent}/{source}" if parent else source


def _resolve(source_part: str, target: str) -> str:
    import posixpath

    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target))


def _prune_relationships(
    members: dict[str, bytes], limits: PackageLimits, actions: list[RecoveryAction]
) -> None:
    for name in [part for part in members if part.endswith(".rels")]:
        root = _relationships(members, name, limits)
        if root is None:
            del members[name]
            actions.append(RecoveryAction("part_dropped_malformed_xml", name))
            continue
        source = _source_part(name)
        if source and source not in members:
            del members[name]
            actions.append(RecoveryAction("relationships_dropped_orphan", name))
            continue
        changed = False
        seen: set[str] = set()
        for node in list(root.findall(f"{{{_PR}}}Relationship")):
            identifier = node.get("Id", "")
            target = node.get("Target", "")
            external = node.get("TargetMode") == "External"
            valid = bool(identifier and node.get("Type") and target) and identifier not in seen
            if valid and not external:
                resolved = _resolve(source, target)
                valid = not resolved.startswith("..") and resolved in members
            if not valid:
                root.remove(node)
                changed = True
                actions.append(RecoveryAction("relationship_dropped", f"{name}#{identifier}"))
            else:
                seen.add(identifier)
        if changed:
            members[name] = _serialize(root)


def _strip_dangling_references(
    members: dict[str, bytes], limits: PackageLimits, actions: list[RecoveryAction]
) -> None:
    for name in list(members):
        if not name.endswith(".xml"):
            continue
        folder, _, leaf = name.rpartition("/")
        rels_name = f"{folder}/_rels/{leaf}.rels" if folder else f"_rels/{leaf}.rels"
        rels = _relationships(members, rels_name, limits)
        identifiers = (
            set()
            if rels is None
            else {node.get("Id", "") for node in rels.findall(f"{{{_PR}}}Relationship")}
        )
        try:
            root = _parse(members[name], limits)
        except (ET.ParseError, _XmlSafetyError):
            continue
        parents = {child: parent for parent in root.iter() for child in parent}
        dangling = [
            (node, key)
            for node in root.iter()
            for key, value in node.attrib.items()
            if key.startswith(f"{{{_R}}}") and value and value not in identifiers
        ]
        if not dangling:
            continue
        removed: set[int] = set()
        for node, key in dangling:
            if any(id(ancestor) in removed for ancestor in _ancestors(node, parents)):
                continue
            local = node.tag.rpartition("}")[2]
            reference = node.get(key)
            target = _reference_owner(node, parents)
            if local == "hyperlink" and node.tag == f"{{{_W}}}hyperlink":
                parent = parents.get(node)
                if parent is not None:
                    index = list(parent).index(node)
                    parent.remove(node)
                    for offset, child in enumerate(list(node)):
                        parent.insert(index + offset, child)
                        parents[child] = parent
                    removed.add(id(node))
            elif target is None:
                del node.attrib[key]
            else:
                parent = parents[target]
                parent.remove(target)
                removed.add(id(target))
            actions.append(RecoveryAction("reference_dropped", f"{name}#{reference}"))
        members[name] = _serialize(root)


def _ancestors(node: ET.Element, parents: dict[ET.Element, ET.Element]) -> Iterator[ET.Element]:
    current: ET.Element | None = node
    while current is not None:
        yield current
        current = parents.get(current)


def _reference_owner(node: ET.Element, parents: dict[ET.Element, ET.Element]) -> ET.Element | None:
    """The element to remove for a dangling reference; ``None`` keeps the element."""
    for ancestor in _ancestors(node, parents):
        if ancestor.tag.rpartition("}")[2] in _UNSUPPORTED_REFERENCE_PARENTS and (
            ancestor.tag.startswith(f"{{{_W}}}")
        ):
            return ancestor if ancestor in parents else None
    return node if node in parents else None


def _rebuild(members: dict[str, bytes], content_types: ET.Element) -> bytes:
    buffer = io.BytesIO()
    ordered = ["_rels/.rels"] + [
        name for name in members if name not in {"_rels/.rels", "[Content_Types].xml"}
    ]
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        _write_member(
            archive,
            "[Content_Types].xml",
            members.get("[Content_Types].xml", _serialize(content_types)),
        )
        for name in ordered:
            if name in members:
                _write_member(archive, name, members[name])
    return buffer.getvalue()


def _write_member(archive: zipfile.ZipFile, name: str, payload: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=_FIXED_ZIP_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o600 << 16
    archive.writestr(info, payload)


# -- text helpers shared by plain-text and HTML import --------------------------------


def first_strong_direction(text: str) -> str | None:
    """``"rtl"`` or ``"ltr"`` for the first strong bidirectional character, else None."""
    for character in text:
        kind = unicodedata.bidirectional(character)
        if kind in {"R", "AL"}:
            return "rtl"
        if kind == "L":
            return "ltr"
    return None


def contains_right_to_left(text: str) -> bool:
    return any(unicodedata.bidirectional(character) in {"R", "AL"} for character in text)


def detect_text_encoding(data: bytes, fallback: str | None = "cp1252") -> str:
    """Deterministic UTF detection, then an explicitly named legacy fallback.

    BOM-less UTF-16 needs an even length and a majority of NUL bytes on exactly
    one byte lane. Ambiguous non-ASCII UTF-16 requires the caller's ``encoding``;
    no statistical language or code-page guess is made.
    """
    if data.startswith(b"\xef\xbb\xbf"):
        return "utf-8"
    if data.startswith(b"\xff\xfe"):
        return "utf-16-le"
    if data.startswith(b"\xfe\xff"):
        return "utf-16-be"
    sample = data[:8192]
    if len(data) % 2 == 0 and len(sample) >= 4:
        even, odd = sample[::2], sample[1::2]
        for name, zero_lane, other_lane in (
            ("utf-16-le", odd, even),
            ("utf-16-be", even, odd),
        ):
            if (
                zero_lane.count(0) > len(zero_lane) / 2
                and other_lane.count(0) < len(other_lane) / 4
            ):
                try:
                    data.decode(name)
                except UnicodeError:
                    break
                return name
    try:
        data.decode("utf-8-sig")
        return "utf-8"
    except UnicodeError:
        if fallback is None:
            return "utf-8"
    try:
        codec = codecs.lookup(fallback)
    except (LookupError, TypeError):
        raise LoadError("load_legacy_encoding_invalid", "Legacy encoding is invalid.") from None
    if not getattr(codec, "_is_text_encoding", True):
        raise LoadError("load_legacy_encoding_invalid", "Legacy encoding is invalid.")
    return codec.name


# -- format dispatch -------------------------------------------------------------------


@dataclass(frozen=True)
class FormatDetection:
    """Container-only inspection; encrypted package macro contents are unknown.

    ``has_macros=None`` means the package hides its macro metadata behind
    encryption, not that the source has no macros.

    ``fib_version`` is the effective MS-DOC nFib (nFibNew takes precedence).
    ``doc_version`` names the recognized file-format generation, not the authoring
    application. Unknown versions retain their numeric FIB value but have no
    generation label; unreadable FIBs and non-DOC inputs have neither. Inspection
    reads at most 4096 WordDocument bytes and 64 KiB of an XML preamble.
    """

    format: str
    encrypted: bool = False
    has_macros: bool | None = False
    fib_version: int | None = None
    doc_version: str | None = None
    encoding: str | None = None


_DOC_VERSIONS = {0x00C1: "Word 97", 0x00D9: "Word 2000", 0x0101: "Word 2002", 0x010C: "Word 2003"}
_WORDML_ROOT = "http://schemas.microsoft.com/office/word/2003/wordml}wordDocument"
_FIB_INSPECTION_BYTES = 4096
_XML_SIGNATURE_BYTES = 64 * 1024


def _doc_fib_prefix(container: _CompoundFile) -> bytes:
    """Read only the bounded FIB prefix, including CFB mini-stream inputs."""
    _feature_legacy_doc = feature_module("legacy_doc")
    LegacyDocError = _feature_legacy_doc.LegacyDocError

    entry = next(
        (item for item in container._entries if item.kind == 2 and item.name == "WordDocument"),
        None,
    )
    if entry is None or entry.size > len(container._payload):
        raise LegacyDocError("legacy_doc_corrupt")
    size = min(entry.size, _FIB_INSPECTION_BYTES)
    if size == 0:
        return b""
    if entry.size < container._mini_cutoff:
        return container.read_stream("WordDocument")[:size]
    count = (size + container._sector_size - 1) // container._sector_size
    return container._read_regular_chain(entry.start_sector, count)[:size]


def _doc_fib_version(prefix: bytes) -> int | None:
    """MS-DOC 2.5.14: locate cswNew by bounded counts, then prefer nFibNew."""
    if len(prefix) < 34 or struct.unpack_from("<H", prefix)[0] != 0xA5EC:
        return None
    position = 32
    for width, expected in ((2, 14), (4, 22), (8, None)):
        if position + 2 > len(prefix):
            return None
        count = struct.unpack_from("<H", prefix, position)[0]
        if expected is not None and count != expected:
            return None
        position += 2 + count * width
    if position + 2 > len(prefix):
        return None
    csw_new = struct.unpack_from("<H", prefix, position)[0]
    if csw_new == 0:
        version = struct.unpack_from("<H", prefix, 2)[0]
    elif csw_new < 2 or position + 2 + csw_new * 2 > len(prefix):
        return None
    else:
        version = struct.unpack_from("<H", prefix, position + 2)[0]
    expected_layout = {0x00C1: (93, 0), 0x00D9: (108, 2), 0x0101: (136, 2), 0x010C: (164, 2)}
    if version in expected_layout and (count, csw_new) != expected_layout[version]:
        return None
    return version


def _xml_root_name(data: bytes) -> str | None:
    """Resolve the first expanded XML name without loading or expanding a story."""
    candidate = data[:_XML_SIGNATURE_BYTES].replace(b"\0", b"")
    candidate = candidate.removeprefix(b"\xef\xbb\xbf").removeprefix(b"\xff\xfe")
    candidate = candidate.removeprefix(b"\xfe\xff").lstrip()
    if not candidate.startswith(b"<"):
        return None
    parser = expat.ParserCreate(namespace_separator="}")
    root: str | None = None

    def start(name: str, attributes: dict[str, str]) -> None:
        nonlocal root
        root = name
        raise StopIteration

    def doctype(*args: Any) -> None:
        raise LoadError("xml_dtd_forbidden", "Format input contains a forbidden DTD.")

    parser.StartElementHandler = start
    parser.StartDoctypeDeclHandler = doctype
    try:
        parser.Parse(data[:_XML_SIGNATURE_BYTES], len(data) <= _XML_SIGNATURE_BYTES)
    except StopIteration:
        return root
    except expat.ExpatError:
        return None
    if len(data) > _XML_SIGNATURE_BYTES:
        raise LoadError("xml_signature_limit", "XML root exceeds the 64 KiB inspection limit.")
    return None


def inspect_format(
    data: bytes, name: str = "", *, legacy_encoding: str = "cp1252"
) -> FormatDetection:
    """Identify signatures and bounded package metadata without loading stories.

    Package content types take precedence over misleading filename suffixes.
    Markdown has no unique signature, so its filename extension selects it.
    Damaged ZIPs retain their bounded ODT signature or package-family suffix for
    the load recovery policy.
    """
    if not isinstance(data, bytes):
        raise TypeError("format_input_invalid")
    if len(data) > PackageLimits().max_archive_size:
        raise LoadError("archive_size_limit", "Format input exceeds limits.")
    suffix = Path(name).suffix.lower()
    head = data[:2048]
    stripped = head.removeprefix(b"\xef\xbb\xbf").lstrip()
    lowered = stripped[:512].lower()
    if head.startswith(_CFB_SIGNATURE):
        _feature_legacy_doc = feature_module("legacy_doc")
        LegacyDocError = _feature_legacy_doc.LegacyDocError
        _CompoundFile = _feature_legacy_doc._CompoundFile

        encrypted = _ENCRYPTED_PACKAGE_NAME in data
        macros = False
        fib_version = None
        kind = "encrypted_docx" if encrypted else "doc"
        try:
            container = _CompoundFile(data)
            names = {entry.name.casefold() for entry in container._entries}
            macros = bool(names & {"vba", "_vba_project", "macros"})
            if "worddocument" in names:
                prefix = _doc_fib_prefix(container)
                if len(prefix) >= 12:
                    encrypted = bool(struct.unpack_from("<H", prefix, 10)[0] & 0x8100)
                fib_version = _doc_fib_version(prefix)
                kind = "doc"
            else:
                encrypted = "encryptedpackage" in names
                kind = "encrypted_docx" if encrypted else "doc"
        except LegacyDocError:
            pass
        return FormatDetection(
            kind,
            encrypted,
            None if kind == "encrypted_docx" else macros,
            fib_version,
            _DOC_VERSIONS.get(fib_version),
        )
    if head.startswith((_LOCAL_HEADER, b"PK\x05\x06")):
        kind = (
            "odt"
            if b"mimetypeapplication/vnd.oasis.opendocument.text" in data[:256]
            else suffix[1:]
            if suffix in {".docx", ".docm", ".dotx", ".dotm"}
            else "docx"
        )
        macros = False
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                entries = archive.infolist()
                if len(entries) > PackageLimits().max_entries:
                    raise LoadError("entry_limit", "Format container exceeds entry limits.")
                names = {entry.filename for entry in entries}
                if any(item.flag_bits & 1 for item in entries):
                    return FormatDetection(kind, encrypted=True, has_macros=None)
                macros = any(
                    item.filename.casefold().endswith("/vbaproject.bin") for item in entries
                )
                if "mimetype" in names:
                    entry = archive.getinfo("mimetype")
                    if entry.file_size <= 256:
                        with archive.open(entry) as stream:
                            if stream.read(257) == b"application/vnd.oasis.opendocument.text":
                                return FormatDetection("odt", has_macros=macros)
                if "[Content_Types].xml" in names:
                    entry = archive.getinfo("[Content_Types].xml")
                    limits = PackageLimits(max_xml_part_size=1024 * 1024)
                    if entry.file_size > limits.max_xml_part_size:
                        raise LoadError("xml_size_limit", "Format metadata exceeds limits.")
                    with archive.open(entry) as stream:
                        payload = stream.read(limits.max_xml_part_size + 1)
                    root = _parse_bounded_xml(
                        payload, limits=limits, total_elements=[0], preserve_qnames=False
                    )
                    for item in root:
                        content_type = item.get("ContentType", "")
                        if content_type.endswith("wordprocessingml.document.main+xml"):
                            kind = "docx"
                        elif content_type.endswith("wordprocessingml.template.main+xml"):
                            kind = "dotx"
                        elif content_type.endswith("word.document.macroEnabled.main+xml"):
                            kind, macros = "docm", True
                        elif content_type.endswith("word.template.macroEnabledTemplate.main+xml"):
                            kind, macros = "dotm", True
        except (zipfile.BadZipFile, KeyError, OSError, ET.ParseError):
            pass
        except _XmlSafetyError as exc:
            raise LoadError(exc.code, "Unsafe format metadata.") from None
        return FormatDetection(kind, has_macros=macros)
    if stripped.startswith(b"{\\rtf"):
        return FormatDetection("rtf")
    if stripped.startswith(b"%PDF-"):
        return FormatDetection("pdf")
    if lowered.startswith(b"mime-version:") or b"multipart/related" in lowered:
        return FormatDetection("mhtml")
    if lowered.startswith((b"<!doctype html", b"<html")):
        return FormatDetection("html")
    if _xml_root_name(data) == _WORDML_ROOT:
        return FormatDetection("wordml")
    if lowered.startswith(b"<?xml") and b"pkg:package" in lowered:
        return FormatDetection("flat_opc")
    if suffix in {".htm", ".html", ".xhtml"}:
        return FormatDetection("html")
    if suffix in {".docx", ".docm", ".dotx", ".dotm"}:
        return FormatDetection(suffix[1:])
    return FormatDetection(
        "markdown" if suffix in {".md", ".markdown"} else "txt",
        encoding=detect_text_encoding(data, legacy_encoding),
    )


def detect_format(data: bytes, name: str = "") -> str:
    """Return the source family; see :func:`inspect_format` for security flags."""
    return inspect_format(data, name).format


def _canonical_import_bytes(data: bytes) -> bytes:
    """Fixed ZIP metadata for generated imports; member payloads stay identical."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(buffer, "w") as target:
        for entry in source.infolist():
            _write_member(target, entry.filename, source.read(entry))
    return buffer.getvalue()


def _save_import_builder(builder: DocumentBuilder, output: Path) -> DocumentSaveReport:
    """Stage generated bytes with fixed metadata before non-clobbering publication."""
    import os
    import tempfile

    with tempfile.TemporaryDirectory(prefix=".wolfdocx-import-", dir=output.parent) as scratch:
        staged = Path(scratch) / "document.docx"
        report = builder.save(staged)
        data = _canonical_import_bytes(staged.read_bytes())
        staged.write_bytes(data)
        os.link(staged, output)
    return replace(report, destination=output, sha256=hashlib.sha256(data).hexdigest())


def load_document(
    source: str | Path,
    options: LoadOptions | None = None,
    *,
    limits: PackageLimits | None = None,
) -> Document:
    """Open any supported source as an in-memory :class:`~wolfdocx.Document`.

    The format is detected from the source bytes.  DOCX and encrypted DOCX open
    directly; every other format is converted by its importer into a private
    staged package that is then admitted strictly.  Refusals raise
    :class:`LoadError` with the importer's named code.
    """
    import tempfile

    from .document import Document

    options = options or LoadOptions()
    path = Path(source)
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise LoadError("load_source_unreadable", "The source could not be read.") from exc
    kind = detect_format(data, path.name)
    if kind in {"docx", "docm", "dotx", "dotm", "encrypted_docx"}:
        document = Document.from_bytes(data, path=path, limits=limits, load_options=options)
        if document.load_report is not None and kind != document.load_report.source_format:
            document.load_report = replace(document.load_report, source_format=kind)
        return document
    source_sha256 = hashlib.sha256(data).hexdigest()
    options.notify("read", 1, 3)
    with tempfile.TemporaryDirectory(prefix="wolfdocx-load-", dir=options.temp_folder) as staging:
        output = Path(staging) / "converted.docx"
        options.notify(f"{kind}.import", 0, 1)
        code, recovery = _convert(kind, path, data, output, options)
        options.notify(f"{kind}.import", 1, 1)
        if code is not None:
            report = LoadReport(
                "refused",
                code,
                kind,
                source_sha256,
                None,
                False,
                options.recovery.value,
            )
            raise LoadError(code, f"Load refused: {code}.", report=report)
        options.notify("convert", 2, 3)
        converted = _canonical_import_bytes(output.read_bytes())
    strict = replace(
        options,
        password=None,
        recovery=RecoveryMode.REFUSE,
        defer_part_validation=False,
        progress=None,
    )
    document = Document.from_bytes(converted, path=path, limits=limits, load_options=strict)
    package_report = document.load_report
    assert package_report is not None
    document.load_report = LoadReport(
        "recovered" if recovery else "opened",
        None,
        kind,
        source_sha256,
        package_report.package_sha256,
        False,
        options.recovery.value,
        tuple(RecoveryAction(item) for item in recovery),
        package_report.editing_language,
        package_report.editing_language_source,
    )
    options.notify("complete", 3, 3)
    return document


def _convert(
    kind: str, path: Path, data: bytes, output: Path, options: LoadOptions
) -> tuple[str | None, tuple[str, ...]]:
    if kind == "txt":
        _feature_txt_import = feature_module("txt_import")
        plaintext_to_docx = _feature_txt_import.plaintext_to_docx

        text_report = plaintext_to_docx(io.BytesIO(data), output, load_options=options)
        return text_report.code, text_report.recovery
    if kind == "markdown":
        _feature_builder = feature_module("builder")
        DocumentBuilder = _feature_builder.DocumentBuilder
        _feature_markdown = feature_module("markdown")
        insert_markdown = _feature_markdown.insert_markdown

        blank = output.with_name("blank.docx")
        DocumentBuilder().save(blank)
        report = insert_markdown(blank, data, output, load_options=options)
        return report.code, ()
    if kind == "html":
        _feature_builder = feature_module("builder")
        DocumentBuilder = _feature_builder.DocumentBuilder
        _feature_html_import = feature_module("html_import")
        insert_html = _feature_html_import.insert_html

        blank = output.with_name("blank.docx")
        DocumentBuilder().save(blank)
        html_report = insert_html(blank, data, output, load_options=options)
        return html_report.code, html_report.load_recovery
    if kind == "flat_opc":
        _feature_flat_opc = feature_module("flat_opc")
        FlatOpcError = _feature_flat_opc.FlatOpcError
        import_flat_opc = _feature_flat_opc.import_flat_opc

        try:
            import_flat_opc(path, output)
        except FlatOpcError as exc:
            return str(exc), ()
        return None, ()
    converters: dict[str, Callable[[Path, Path], Any]] = {}
    if kind == "rtf":
        _feature_rtf = feature_module("rtf")
        rtf_to_docx = _feature_rtf.rtf_to_docx

        report = rtf_to_docx(path, output, load_options=options)
        return report.code, report.omissions
    elif kind == "odt":
        _feature_odt = feature_module("odt")
        odt_to_docx = _feature_odt.odt_to_docx

        converters[kind] = odt_to_docx
    elif kind == "mhtml":
        _feature_mhtml = feature_module("mhtml")
        mhtml_to_docx = _feature_mhtml.mhtml_to_docx

        converters[kind] = mhtml_to_docx
    elif kind == "wordml":
        _feature_wordml = feature_module("wordml")
        wordml_to_docx = _feature_wordml.wordml_to_docx

        # Explicit conversion formats must obey the same DTD refusal as detection.
        _xml_root_name(data)
        converters[kind] = wordml_to_docx
    elif kind == "doc":
        _feature_legacy_doc = feature_module("legacy_doc")
        doc_to_docx = _feature_legacy_doc.doc_to_docx

        converters[kind] = doc_to_docx
    elif kind == "pdf":
        _feature_pdf_import = feature_module("pdf_import")
        pdf_to_docx = _feature_pdf_import.pdf_to_docx

        report = pdf_to_docx(path, output, pages=options.pdf_pages, load_options=options)
        return (None if report.status == "published" else report.code or "load_import_refused"), ()
    else:
        return "load_format_unsupported", ()
    report = converters[kind](path, output)
    return (None if report.status == "published" else report.code or "load_import_refused"), ()
