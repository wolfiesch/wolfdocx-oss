"""Strict no-follow verification and publication of declared worker outputs.

Worker directories are hostile until this module has opened every declared file
through directory descriptors, verified it is an unlinked regular file owned by
the declared worker identity, and bound its content to a SHA-256 receipt.  The
transport carries only relative display paths and artifact identities; no local
source or destination path is ever returned in a diagnostic.
"""

from __future__ import annotations

import codecs
import contextlib
import hashlib
import os
import secrets
import stat
import unicodedata
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field, fields, replace
from enum import Enum, IntEnum
from types import MappingProxyType
from typing import TYPE_CHECKING, BinaryIO, Literal, Self

from ._optional import feature_module
from .contracts import ArtifactIdentity, ContractArtifact, ContractValidationError
from .errors import ErrorCategory, WolfDocxError

if TYPE_CHECKING:
    from pathlib import Path

    from .document import Document, SessionSaveReport

    _feature_native_render = feature_module("native_render")
    FontAsset = _feature_native_render.FontAsset
    NativeRenderReport = _feature_native_render.NativeRenderReport
    _feature_xlsx_export = feature_module("xlsx_export")
    XlsxExportReport = _feature_xlsx_export.XlsxExportReport

OUTPUT_RECEIPT_SCHEMA_VERSION = 1
MAX_OUTPUT_DECLARATIONS = 1024
_COPY_CHUNK_BYTES = 1024 * 1024
_REGULAR_FILE = "regular_file"
_DECLARATION_VALIDATION_IDENTITY = ArtifactIdentity("0" * 64, 0)
_SECURE_DIR_FD_AVAILABLE = all(
    function in os.supports_dir_fd
    for function in (os.open, os.stat, os.link, os.mkdir, os.unlink, os.rmdir)
)
_STAT_NOFOLLOW_AVAILABLE = os.stat in os.supports_follow_symlinks


DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@dataclass(frozen=True, kw_only=True)
class SaveOutputParameters:
    """Format metadata shared by successful saves and typed refusal receipts."""

    content_type: str = DOCX_CONTENT_TYPE
    written_format: str = "docx"


class SaveOptionsError(ValueError):
    """Invalid save configuration or content that cannot be exported faithfully."""

    def __init__(self, code: str) -> None:
        self.code = code
        super().__init__(code)


class TxtExportHeadersFootersMode(str, Enum):  # noqa: UP042, RUF100
    NONE = "none"
    PRIMARY_ONLY = "primary_only"
    ALL = "all"


class TxtOfficeMathExportMode(str, Enum):  # noqa: UP042, RUF100
    TEXT = "text"
    LINEAR = "linear"


class OoxmlCompliance(str, Enum):  # noqa: UP042, RUF100
    TRANSITIONAL = "transitional"
    STRICT = "strict"
    ECMA_376_2006 = "ecma_376_2006"


class CompressionLevel(IntEnum):
    NO_COMPRESSION = 0
    FASTEST = 1
    OPTIMAL = 6
    MAXIMUM = 9


class Zip64Mode(str, Enum):  # noqa: UP042, RUF100
    NEVER = "never"
    AS_NEEDED = "as_needed"
    ALWAYS = "always"


@dataclass(frozen=True)
class SaveOptions:
    """Shared immutable model for every format's save policy.

    Use a format-specific subclass with its exporter. Options never modify the
    source package; unsupported policies refuse before output publication.
    """

    progress_callback: Callable[[str], None] | None = field(default=None, kw_only=True)
    property_updates: Mapping[str, object] | None = field(default=None, kw_only=True)
    _save_format: str | None = field(default=None, init=False, compare=False, repr=False)

    def _validate_lifecycle(self) -> None:
        if self.progress_callback is not None and not callable(self.progress_callback):
            raise SaveOptionsError("save_progress_callback_invalid")
        if self.property_updates is not None:
            if not isinstance(self.property_updates, Mapping):
                raise SaveOptionsError("save_property_updates_invalid")
            from .properties import _apply_core_properties

            try:
                _apply_core_properties(None, self.property_updates)
            except (TypeError, ValueError) as exc:
                raise SaveOptionsError(str(exc)) from exc
            object.__setattr__(
                self, "property_updates", MappingProxyType(dict(self.property_updates))
            )

    def _progress(self, stage: str) -> None:
        if self.progress_callback is not None:
            self.progress_callback(stage)

    def save(
        self,
        source: str | Path | Document,
        destination: str | Path | BinaryIO,
        *,
        fonts: Sequence[FontAsset] = (),
    ) -> SessionSaveReport | NativeRenderReport | XlsxExportReport:
        """Save current edits with read/render/publish callbacks before publication.

        ``property_updates`` uses the existing core-property names and validation.
        Apply updates to the conversion snapshot only. Each existing exporter
        retains its metadata scope; body-only exports do not serialize core
        properties. Explicit dates keep saves deterministic. Render formats
        publish a directory, as ``render_native`` does; other formats publish
        a new path or empty seekable binary stream. Callback exceptions propagate
        unchanged, before publication.
        """
        return _save_with_lifecycle(self, source, destination, fonts=fonts)

    @staticmethod
    def create(format_or_extension: str, **policies: object) -> SaveOptions:
        "Create a format policy, resolving optional formats only when requested."
        from pathlib import PurePath

        if not isinstance(format_or_extension, str) or not format_or_extension:
            raise SaveOptionsError("save_format_invalid")
        name = format_or_extension.casefold()
        name = (PurePath(name).suffix or name).lstrip(".")
        core = {
            "docx": OoxmlSaveOptions,
            "docm": OoxmlSaveOptions,
            "txt": TxtSaveOptions,
            "text": TxtSaveOptions,
        }
        constructor = core.get(name)
        if constructor is None:
            formats = {
                "html": ("html_export", "HtmlSaveOptions"),
                "htm": ("html_export", "HtmlSaveOptions"),
                "md": ("markdown", "MarkdownExportOptions"),
                "markdown": ("markdown", "MarkdownExportOptions"),
                "rtf": ("rtf_export", "RtfSaveOptions"),
                "doc": ("legacy_doc_export", "DocSaveOptions"),
                "odt": ("odt", "OdtSaveOptions"),
                "xlsx": ("xlsx_export", "XlsxSaveOptions"),
                **dict.fromkeys(
                    ("pdf", "svg", "xps", "png", "jpeg", "jpg", "bmp", "tiff", "tif"),
                    ("native_render", "RenderSaveOptions"),
                ),
            }
            target = formats.get(name)
            if target is None:
                raise SaveOptionsError("save_format_unsupported")
            constructor = getattr(feature_module(target[0]), target[1])
        supported = {field.name for field in fields(constructor) if field.init}
        if policies.keys() - supported:
            raise SaveOptionsError("save_policy_unsupported")
        options = constructor(**policies)
        options._validate_lifecycle()
        object.__setattr__(options, "_save_format", name)
        return options


@dataclass(frozen=True)
class TxtListIndentation:
    """Repeat ``character`` ``count`` times per zero-based list nesting level."""

    character: str = " "
    count: int = 2

    def __post_init__(self) -> None:
        if not isinstance(self.character, str) or len(self.character) != 1:
            raise SaveOptionsError("text_list_indentation_character_invalid")
        if type(self.count) is not int or not 0 <= self.count <= 1024:
            raise SaveOptionsError("text_list_indentation_count_invalid")


@dataclass(frozen=True)
class TxtSaveOptionsBase(SaveOptions):
    encoding: str = "utf-8"
    header_footer_mode: TxtExportHeadersFootersMode = TxtExportHeadersFootersMode.NONE
    list_indentation: TxtListIndentation = TxtListIndentation()
    export_list_labels: bool = True
    office_math_mode: TxtOfficeMathExportMode = TxtOfficeMathExportMode.TEXT
    export_page_breaks: bool = True
    paragraph_break: str = "\n"
    table_cell_separator: str = "\t"
    export_tables: bool = True

    def __post_init__(self) -> None:
        try:
            codecs.lookup(self.encoding)
            "".encode(self.encoding)
        except (LookupError, TypeError) as exc:
            raise SaveOptionsError("text_encoding_invalid") from exc
        if self.header_footer_mode not in tuple(TxtExportHeadersFootersMode):
            raise SaveOptionsError("text_header_footer_mode_invalid")
        if not isinstance(self.list_indentation, TxtListIndentation):
            raise SaveOptionsError("text_list_indentation_invalid")
        if self.office_math_mode not in tuple(TxtOfficeMathExportMode):
            raise SaveOptionsError("text_office_math_mode_invalid")
        if type(self.export_list_labels) is not bool or type(self.export_page_breaks) is not bool:
            raise SaveOptionsError("text_export_flag_invalid")
        if not isinstance(self.paragraph_break, str):
            raise SaveOptionsError("text_paragraph_break_invalid")
        if not isinstance(self.table_cell_separator, str):
            raise SaveOptionsError("text_table_cell_separator_invalid")
        if type(self.export_tables) is not bool:
            raise SaveOptionsError("text_export_tables_invalid")


@dataclass(frozen=True)
class TxtSaveOptions(TxtSaveOptionsBase):
    """TXT wrapping, bidi, table alignment, list, math and story placement policy.

    Both ``options.save`` and ``Document.save_text`` consume this same policy.
    """

    maximum_line_length: int = 0
    add_bidi_marks: bool = False
    preserve_table_layout: bool = False
    table_alignment: str = "left"
    simplify_list_labels: bool = False
    math_format: str = "legacy"
    headers_footers_at_end: bool = False
    progress_callback: Callable[[str], None] | None = None

    def __post_init__(self) -> None:
        super().__post_init__()
        object.__setattr__(self, "paragraph_break", str(self.paragraph_break))
        if type(self.maximum_line_length) is not int or not 0 <= self.maximum_line_length <= 100000:
            raise SaveOptionsError("text_maximum_line_length_invalid")
        for name in (
            "add_bidi_marks",
            "preserve_table_layout",
            "simplify_list_labels",
            "headers_footers_at_end",
        ):
            if type(getattr(self, name)) is not bool:
                raise SaveOptionsError(f"text_{name}_invalid")
        if self.table_alignment not in {"left", "center", "right", "auto"}:
            raise SaveOptionsError("text_table_alignment_invalid")
        if self.table_alignment != "left" and not self.preserve_table_layout:
            raise SaveOptionsError("text_table_alignment_requires_layout")
        if self.math_format not in {"legacy", "latex", "mathml"}:
            raise SaveOptionsError("text_math_format_invalid")


@dataclass(frozen=True)
class OoxmlSaveOptions(SaveOptions):
    """Explicit namespace conversion and deterministic ZIP container policy.

    ``compression_level=None`` preserves source compression, like a default save.
    An integer 0–9 selects DEFLATE for all members. ``always`` forces ZIP64 local
    member headers even for small packages; ``never`` forbids ZIP64 records.
    ``ECMA_376_2006`` downgrades admitted directional property names and removes
    the later conformance attribute. Later extensions, contextual alignment and
    unsupported bidi downgrades refuse with ``ooxml_2006_*`` codes.
    ``export_legacy_control_characters`` represents XML-safe embedded tab,
    LF and CR text as native run nodes. XML-forbidden controls remain refused.
    ``signing_details`` accepts certificate/private_key paths, an optional
    environment-only passphrase name, signing_time, comments and signature-line
    ID. Supply an aware signing_time for deterministic RSA-SHA256 OPC XMLDSIG.
    Routing slips have no admitted native OOXML representation and refuse.
    """

    compliance: OoxmlCompliance = OoxmlCompliance.TRANSITIONAL
    compression_level: int | None = None
    zip64_mode: Zip64Mode = Zip64Mode.AS_NEEDED
    signing_details: Mapping[str, object] | None = None
    routing_slip: object | None = None
    export_legacy_control_characters: bool = False

    def __post_init__(self) -> None:
        self._validate_lifecycle()
        if self.routing_slip is not None:
            raise SaveOptionsError("ooxml_routing_slip_unsupported")
        if type(self.export_legacy_control_characters) is not bool:
            raise SaveOptionsError("ooxml_legacy_control_policy_invalid")
        if self.signing_details is not None:
            if not isinstance(self.signing_details, Mapping):
                raise SaveOptionsError("ooxml_signing_details_invalid")
            allowed = {
                "certificate",
                "private_key",
                "private_key_password_env",
                "signature_line_id",
                "comments",
                "signing_time",
            }
            if (
                self.signing_details.keys() - allowed
                or not {"certificate", "private_key"} <= self.signing_details.keys()
            ):
                raise SaveOptionsError("ooxml_signing_details_invalid")
            from pathlib import Path

            if any(
                not isinstance(self.signing_details[name], (str, Path))
                for name in ("certificate", "private_key")
            ):
                raise SaveOptionsError("ooxml_signing_details_invalid")
            object.__setattr__(
                self, "signing_details", MappingProxyType(dict(self.signing_details))
            )
        if self.compliance not in tuple(OoxmlCompliance):
            raise SaveOptionsError("ooxml_compliance_invalid")
        if self.zip64_mode not in tuple(Zip64Mode):
            raise SaveOptionsError("ooxml_zip64_mode_invalid")
        if self.compression_level is not None and (
            isinstance(self.compression_level, bool)
            or not isinstance(self.compression_level, int)
            or not 0 <= self.compression_level <= 9
        ):
            raise SaveOptionsError("ooxml_compression_level_invalid")


def _apply_save_properties(data: bytes, updates: Mapping[str, object] | None) -> bytes:
    if not updates:
        return data
    import tempfile
    from pathlib import Path

    from .properties import PropertiesRequest, author_properties

    with tempfile.TemporaryDirectory(prefix="wolfdocx-save-properties-") as directory:
        source = Path(directory) / "source.docx"
        destination = Path(directory) / "properties.docx"
        source.write_bytes(data)
        report = author_properties(
            PropertiesRequest(
                source, hashlib.sha256(data).hexdigest(), destination, core=dict(updates)
            )
        )
        if report.status != "published":
            raise SaveOptionsError(report.code or "save_property_update_failed")
        return destination.read_bytes()


def _save_with_lifecycle(
    options: SaveOptions,
    source: str | Path | Document,
    destination: str | Path | BinaryIO,
    *,
    fonts: Sequence[FontAsset] = (),
) -> SessionSaveReport | NativeRenderReport | XlsxExportReport:
    import tempfile
    from pathlib import Path

    from .document import _PART_ERRORS, Document, SessionSaveReport

    options._validate_lifecycle()
    if isinstance(options, OoxmlSaveOptions):
        if isinstance(source, Document):
            return source.save(destination, options=options)
        with Document(source) as document:
            return document.save(destination, options=options)
    _feature_html_export = feature_module("html_export")
    HtmlSaveOptions = _feature_html_export.HtmlSaveOptions
    docx_to_html = _feature_html_export.docx_to_html
    _feature_legacy_doc_export = feature_module("legacy_doc_export")
    DocSaveOptions = _feature_legacy_doc_export.DocSaveOptions
    export_doc = _feature_legacy_doc_export.export_doc
    _feature_markdown = feature_module("markdown")
    MarkdownExportOptions = _feature_markdown.MarkdownExportOptions
    docx_to_markdown_report = _feature_markdown.docx_to_markdown_report
    _feature_native_layout = feature_module("native_layout")
    NativeLayoutError = _feature_native_layout.NativeLayoutError
    _feature_native_render = feature_module("native_render")
    NativeRenderRequest = _feature_native_render.NativeRenderRequest
    RenderLimits = _feature_native_render.RenderLimits
    RenderSaveOptions = _feature_native_render.RenderSaveOptions
    _publish_directory = _feature_native_render._publish_directory
    render_native = _feature_native_render.render_native
    _feature_odt = feature_module("odt")
    OdtSaveOptions = _feature_odt.OdtSaveOptions
    docx_to_odt = _feature_odt.docx_to_odt
    _feature_page_text = feature_module("page_text")
    export_text = _feature_page_text.export_text
    from .review import _link_no_clobber

    _feature_rtf_export = feature_module("rtf_export")
    RtfSaveOptions = _feature_rtf_export.RtfSaveOptions
    docx_to_rtf = _feature_rtf_export.docx_to_rtf
    _feature_xlsx_export = feature_module("xlsx_export")
    XlsxExportError = _feature_xlsx_export.XlsxExportError
    XlsxExportReport = _feature_xlsx_export.XlsxExportReport
    XlsxSaveOptions = _feature_xlsx_export.XlsxSaveOptions
    _convert = _feature_xlsx_export._convert

    options._progress("read")
    source_sha = source.source_sha256 if isinstance(source, Document) else None
    with tempfile.TemporaryDirectory(prefix="wolfdocx-save-") as directory:
        root = Path(directory)
        package = root / "source.docx"
        try:
            if isinstance(options, XlsxSaveOptions):
                # A private conversion snapshot needs no publication of its own.
                if isinstance(source, Document):
                    package.write_bytes(source._materialize())
                else:
                    with Document(source) as document:
                        source_sha = document.source_sha256
                        package.write_bytes(document._materialize())
            elif isinstance(source, Document):
                report = source.save(package)
                if report.status != "published":
                    raise SaveOptionsError(report.code or "save_source_invalid")
            else:
                with Document(source) as document:
                    source_sha = document.source_sha256
                    report = document.save(package)
                    if report.status != "published":
                        raise SaveOptionsError(report.code or "save_source_invalid")
        except _PART_ERRORS as exc:
            if isinstance(options, XlsxSaveOptions):
                raise XlsxExportError("xlsx_source_package_invalid") from exc
            raise
        if options.property_updates:
            package.write_bytes(
                _apply_save_properties(package.read_bytes(), options.property_updates)
            )
        if isinstance(options, XlsxSaveOptions) and options.refresh_fields:
            refreshed = root / "refreshed.docx"
            with Document(package) as snapshot:
                refresh = snapshot.refresh_fields(fonts=fonts)
                if not refresh.passed:
                    raise XlsxExportError(refresh.code or "xlsx_field_refresh_failed")
                saved = snapshot.save(refreshed)
                if not saved.passed:
                    raise XlsxExportError(saved.code or "xlsx_field_refresh_failed")
            package = refreshed
        selected = replace(options, progress_callback=None, property_updates=None)
        options._progress("render")
        staged = root / "output"
        media = "application/octet-stream"
        workbook_exporter = None
        if isinstance(options, TxtSaveOptions):
            # Keep caller callback exceptions distinct from content refusals.
            callback_error: Exception | None = None

            def text_progress(stage: str) -> None:
                nonlocal callback_error
                if stage == "publish":
                    try:
                        options._progress(stage)
                    except Exception as exc:
                        callback_error = exc
                        raise

            text_options = replace(
                selected,
                progress_callback=text_progress if options.progress_callback is not None else None,
            )
            try:
                report = export_text(package, destination, options=text_options)
            except (NativeLayoutError,) + _PART_ERRORS as exc:
                if exc is callback_error:
                    raise
                raise SaveOptionsError("text_content_unsupported") from exc
            return replace(
                report,
                source_sha256=source_sha,
                edits=tuple(source._edits) if isinstance(source, Document) else (),
                load_recovery=source.load_recovery if isinstance(source, Document) else None,
            )
        elif isinstance(options, HtmlSaveOptions):
            staged.write_bytes(docx_to_html(package, selected).encode(selected.encoding))
            format_name, media = "html", "text/html"
        elif isinstance(options, MarkdownExportOptions):
            staged.write_text(
                docx_to_markdown_report(package, options=selected).markdown, encoding="utf-8"
            )
            format_name, media = "markdown", "text/markdown"
        elif isinstance(options, RtfSaveOptions):
            staged.write_bytes(docx_to_rtf(package, options=selected))
            format_name, media = "rtf", "application/rtf"
        elif isinstance(options, DocSaveOptions):
            report = export_doc(package, staged, options=selected)
            if report.status != "published":
                raise SaveOptionsError(report.code or "save_doc_failed")
            format_name, media = "doc", "application/msword"
        elif isinstance(options, OdtSaveOptions):
            report = docx_to_odt(package, staged, options=selected)
            if report.status != "published":
                raise SaveOptionsError(report.code or "save_odt_failed")
            format_name, media = "odt", "application/vnd.oasis.opendocument.text"
        elif isinstance(options, XlsxSaveOptions):
            payload, workbook_exporter = _convert(package, selected)
            staged.write_bytes(payload)
            format_name = "xlsx"
            media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif isinstance(options, RenderSaveOptions):
            format_name = options._save_format
            if format_name is None or hasattr(destination, "write"):
                raise SaveOptionsError("save_render_destination_requires_format_directory")
            format_name = {"jpg": "jpeg", "tif": "tiff"}.get(format_name, format_name)
            report = render_native(
                NativeRenderRequest(
                    source=package,
                    source_sha256=hashlib.sha256(package.read_bytes()).hexdigest(),
                    output_dir=staged,
                    fonts=tuple(fonts),
                    formats=(format_name,),
                    scale=1.0,
                    limits=RenderLimits(),
                    save_options=selected,
                )
            )
            if report.status != "published":
                raise SaveOptionsError(report.code or "save_render_failed")
            options._progress("publish")
            target = Path(destination)
            if target.exists() or target.is_symlink() or not target.parent.is_dir():
                raise SaveOptionsError("save_destination_invalid")
            _publish_directory(staged, target)
            return replace(report, source_sha256=source_sha, output_dir=str(target))
        else:
            raise SaveOptionsError("save_format_unsupported")
        payload = staged.read_bytes()
        options._progress("publish")
        publication_error = XlsxExportError if workbook_exporter is not None else SaveOptionsError
        invalid_destination = (
            "xlsx_destination_invalid"
            if workbook_exporter is not None
            else "save_destination_invalid"
        )
        unavailable_publication = (
            "xlsx_publication_unavailable"
            if workbook_exporter is not None
            else "save_publication_unavailable"
        )
        if hasattr(destination, "write"):
            try:
                if not destination.seekable() or destination.tell() != 0:
                    raise publication_error(invalid_destination)
                length = destination.seek(0, 2)
                destination.seek(0)
                if length:
                    raise publication_error(invalid_destination)
            except (AttributeError, OSError, ValueError) as exc:
                if isinstance(exc, publication_error):
                    raise
                raise publication_error(invalid_destination) from exc
            try:
                if destination.write(payload) != len(payload):
                    raise OSError("short stream write")
            except (OSError, TypeError, ValueError) as exc:
                destination.seek(0)
                destination.truncate(0)
                raise publication_error(unavailable_publication) from exc
            display = "<stream>"
        else:
            target = Path(destination)
            if target.is_symlink() or not target.parent.is_dir():
                raise publication_error(invalid_destination)
            code = _link_no_clobber(
                staged,
                target,
                exists_code=(
                    "xlsx_destination_invalid"
                    if workbook_exporter is not None
                    else "document_destination_exists"
                ),
                unavailable_code=unavailable_publication,
            )
            if code:
                raise publication_error(code)
            display = str(target)
        if workbook_exporter is not None:
            return XlsxExportReport(
                "published",
                source_sha,
                hashlib.sha256(payload).hexdigest(),
                tuple(sheet.name for sheet in workbook_exporter.sheets),
                workbook_exporter.paragraphs,
                workbook_exporter.tables,
                tuple(workbook_exporter.omissions),
                destination=display,
                edits=tuple(source._edits) if isinstance(source, Document) else (),
                load_recovery=source.load_recovery if isinstance(source, Document) else (),
            )
        return SessionSaveReport(
            "published",
            None,
            source_sha,
            display,
            destination_sha256=hashlib.sha256(payload).hexdigest(),
            content_type=media,
            written_format=format_name,
            edits=tuple(source._edits) if isinstance(source, Document) else (),
            load_recovery=source.load_recovery if isinstance(source, Document) else None,
        )


class OutputAllowlistError(WolfDocxError):
    """A safe output-boundary rejection that exposes no filesystem detail."""

    category = ErrorCategory.PRECONDITION_FAILED


@dataclass(frozen=True, slots=True)
class OutputDeclaration:
    """One caller-declared regular-file output contract.

    ``display_path`` must already be NFC-normalized and relative.  The boundary
    rejects rather than rewrites a path so the caller's declaration is exactly
    the path recorded in the receipt.  ``expected_identity`` is optional for
    operations whose output bytes cannot be known before worker execution.
    """

    role: str
    display_path: str
    file_type: Literal["regular_file"]
    media_type: str
    expected_identity: ArtifactIdentity | None = None

    def __post_init__(self) -> None:
        if self.file_type != _REGULAR_FILE or type(self.file_type) is not str:
            raise OutputAllowlistError("output_declaration_invalid")
        if type(self.display_path) is not str or self.display_path != unicodedata.normalize(
            "NFC", self.display_path
        ):
            raise OutputAllowlistError("output_path_invalid")
        if self.expected_identity is not None and not isinstance(
            self.expected_identity, ArtifactIdentity
        ):
            raise OutputAllowlistError("output_declaration_invalid")
        try:
            ContractArtifact(
                role=self.role,
                media_type=self.media_type,
                display_path=self.display_path,
                identity=(
                    _DECLARATION_VALIDATION_IDENTITY
                    if self.expected_identity is None
                    else self.expected_identity
                ),
            )
        except ContractValidationError as exc:
            raise OutputAllowlistError("output_declaration_invalid") from exc


@dataclass(frozen=True, slots=True)
class OutputPublicationReceipt:
    """Path-safe verified identities for every published output."""

    owner_uid: int
    artifacts: tuple[ContractArtifact, ...]

    def __post_init__(self) -> None:
        if type(self.owner_uid) is not int or self.owner_uid < 0:
            raise OutputAllowlistError("output_owner_invalid")
        if (
            type(self.artifacts) is not tuple
            or not self.artifacts
            or len(self.artifacts) > MAX_OUTPUT_DECLARATIONS
            or any(not isinstance(artifact, ContractArtifact) for artifact in self.artifacts)
        ):
            raise OutputAllowlistError("output_receipt_invalid")
        roles: set[str] = set()
        paths: set[str] = set()
        for artifact in self.artifacts:
            casefolded = artifact.display_path.casefold()
            if artifact.role in roles or casefolded in paths:
                raise OutputAllowlistError("output_receipt_invalid")
            roles.add(artifact.role)
            paths.add(casefolded)

    def to_dict(self) -> dict[str, object]:
        return {
            "schema_version": OUTPUT_RECEIPT_SCHEMA_VERSION,
            "owner_uid": self.owner_uid,
            "artifacts": [artifact.to_dict() for artifact in self.artifacts],
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        if type(value) is not dict or set(value) != {
            "schema_version",
            "owner_uid",
            "artifacts",
        }:
            raise OutputAllowlistError("output_receipt_invalid")
        if value["schema_version"] != OUTPUT_RECEIPT_SCHEMA_VERSION:
            raise OutputAllowlistError("output_receipt_invalid")
        owner_uid = value["owner_uid"]
        artifacts = value["artifacts"]
        if type(owner_uid) is not int or type(artifacts) is not list:
            raise OutputAllowlistError("output_receipt_invalid")
        if any(not isinstance(item, Mapping) for item in artifacts):
            raise OutputAllowlistError("output_receipt_invalid")
        try:
            return cls(
                owner_uid=owner_uid,
                artifacts=tuple(ContractArtifact.from_dict(item) for item in artifacts),
            )
        except ContractValidationError as exc:
            raise OutputAllowlistError("output_receipt_invalid") from exc


@dataclass(slots=True)
class _VerifiedFile:
    declaration: OutputDeclaration
    descriptor: int
    identity: ArtifactIdentity

    @property
    def artifact(self) -> ContractArtifact:
        return ContractArtifact(
            role=self.declaration.role,
            media_type=self.declaration.media_type,
            display_path=self.declaration.display_path,
            identity=self.identity,
        )


@dataclass(frozen=True, slots=True)
class _TrackedEntry:
    """An inode created by this call and therefore eligible for safe rollback."""

    display_path: str
    device: int
    inode: int


def verify_output_allowlist(
    source_directory: str | os.PathLike[str],
    declarations: Sequence[OutputDeclaration],
    *,
    owner_uid: int,
) -> OutputPublicationReceipt:
    """Verify a hostile worker output tree without publishing any file."""

    verified = _collect_verified_files(source_directory, declarations, owner_uid=owner_uid)
    try:
        return OutputPublicationReceipt(
            owner_uid=owner_uid,
            artifacts=tuple(item.artifact for item in verified),
        )
    finally:
        _close_verified_files(verified)


def publish_verified_outputs(
    source_directory: str | os.PathLike[str],
    destination_directory: str | os.PathLike[str],
    declarations: Sequence[OutputDeclaration],
    *,
    owner_uid: int,
) -> OutputPublicationReceipt:
    """Atomically publish only previously verified regular files.

    ``destination_directory`` must be an existing empty directory.  Every file
    is copied to a private sibling staging tree, re-hashed, and atomically
    linked into its declared final name without replacement.  The temporary
    hardlink is removed before this function returns, leaving regular files
    with a single link at their normalized declared paths.
    """

    verified = _collect_verified_files(source_directory, declarations, owner_uid=owner_uid)
    destination_fd: int | None = None
    staging_entry: _TrackedEntry | None = None
    staged_files: list[_TrackedEntry] = []
    staging_directories: dict[str, _TrackedEntry] = {}
    published_files: list[_TrackedEntry] = []
    destination_directories: dict[str, _TrackedEntry] = {}
    try:
        destination_fd = _open_directory(destination_directory)
        if os.listdir(destination_fd):
            raise OutputAllowlistError("output_destination_not_empty")
        staging_name = f".wolfdocx-output-stage-{secrets.token_hex(16)}"
        os.mkdir(staging_name, 0o700, dir_fd=destination_fd)
        staging_fd = _open_directory_at(destination_fd, staging_name)
        try:
            staging_entry = _tracked_entry(staging_name, os.fstat(staging_fd))
            for item in verified:
                staged_files.append(_stage_verified_file(staging_fd, item, staging_directories))
            for item in verified:
                published_files.append(
                    _publish_staged_file(
                        staging_fd,
                        destination_fd,
                        item.declaration.display_path,
                        destination_directories,
                    )
                )
        finally:
            os.close(staging_fd)
        _discard_tracked_tree(
            destination_fd,
            staging_entry,
            staged_files,
            staging_directories,
        )
        staging_entry = None
        os.fsync(destination_fd)
        return OutputPublicationReceipt(
            owner_uid=owner_uid,
            artifacts=tuple(item.artifact for item in verified),
        )
    except OutputAllowlistError:
        if destination_fd is not None:
            _rollback_publication(
                destination_fd,
                published_files,
                destination_directories,
            )
        raise
    except OSError as exc:
        if destination_fd is not None:
            _rollback_publication(
                destination_fd,
                published_files,
                destination_directories,
            )
        raise OutputAllowlistError("output_publication_failed") from exc
    finally:
        if staging_entry is not None and destination_fd is not None:
            _rollback_tracked_tree(
                destination_fd,
                staging_entry,
                staged_files,
                staging_directories,
            )
        if destination_fd is not None:
            os.close(destination_fd)
        _close_verified_files(verified)


def _collect_verified_files(
    source_directory: str | os.PathLike[str],
    declarations: Sequence[OutputDeclaration],
    *,
    owner_uid: int,
) -> list[_VerifiedFile]:
    _require_secure_primitives()
    declaration_list = _validated_declarations(declarations)
    _require_owner_uid(owner_uid)
    source_fd = _open_directory(source_directory)
    verified: list[_VerifiedFile] = []
    try:
        declared_paths = {item.display_path for item in declaration_list}
        declared_directories = _declared_directories(declared_paths)
        _verify_declared_tree(source_fd, "", declared_paths, declared_directories)
        for declaration in declaration_list:
            descriptor, identity = _open_and_hash_source_file(
                source_fd,
                declaration.display_path,
                owner_uid=owner_uid,
            )
            if (
                declaration.expected_identity is not None
                and declaration.expected_identity != identity
            ):
                os.close(descriptor)
                raise OutputAllowlistError("output_identity_mismatch")
            verified.append(_VerifiedFile(declaration, descriptor, identity))
        return verified
    except Exception:
        _close_verified_files(verified)
        raise
    finally:
        os.close(source_fd)


def _validated_declarations(
    declarations: Sequence[OutputDeclaration],
) -> tuple[OutputDeclaration, ...]:
    if isinstance(declarations, (str, bytes)) or not isinstance(declarations, Sequence):
        raise OutputAllowlistError("output_declaration_invalid")
    values = tuple(declarations)
    if not values or len(values) > MAX_OUTPUT_DECLARATIONS:
        raise OutputAllowlistError("output_declaration_invalid")
    paths: set[str] = set()
    roles: set[str] = set()
    for declaration in values:
        if not isinstance(declaration, OutputDeclaration):
            raise OutputAllowlistError("output_declaration_invalid")
        path_key = declaration.display_path.casefold()
        if path_key in paths:
            raise OutputAllowlistError("output_path_ambiguous")
        if declaration.role in roles:
            raise OutputAllowlistError("output_declaration_invalid")
        paths.add(path_key)
        roles.add(declaration.role)
    return tuple(sorted(values, key=lambda item: item.display_path))


def _declared_directories(paths: Iterable[str]) -> set[str]:
    directories: set[str] = set()
    for path in paths:
        parts = path.split("/")
        directories.update("/".join(parts[:index]) for index in range(1, len(parts)))
    return directories


def _verify_declared_tree(
    directory_fd: int,
    prefix: str,
    declared_paths: set[str],
    declared_directories: set[str],
) -> None:
    try:
        names = os.listdir(directory_fd)
    except OSError as exc:
        raise OutputAllowlistError("output_source_invalid") from exc
    for name in names:
        relative_path = name if not prefix else f"{prefix}/{name}"
        entry = _stat_at(directory_fd, name)
        if relative_path in declared_paths:
            _require_regular_entry(entry)
            continue
        if relative_path in declared_directories:
            _require_directory_entry(entry)
            child_fd = _open_directory_at(directory_fd, name, expected=entry)
            try:
                _verify_declared_tree(
                    child_fd,
                    relative_path,
                    declared_paths,
                    declared_directories,
                )
            finally:
                os.close(child_fd)
            continue
        raise OutputAllowlistError("output_undeclared")


def _open_and_hash_source_file(
    source_fd: int,
    display_path: str,
    *,
    owner_uid: int,
) -> tuple[int, ArtifactIdentity]:
    parent_fd = _open_existing_directory_path(source_fd, display_path.split("/")[:-1])
    name = display_path.rsplit("/", maxsplit=1)[-1]
    try:
        before = _stat_at(parent_fd, name)
        _require_regular_entry(before)
        descriptor = _open_regular_at(parent_fd, name, expected=before)
    finally:
        os.close(parent_fd)
    try:
        initial = os.fstat(descriptor)
        _require_regular_entry(initial)
        if initial.st_uid != owner_uid:
            raise OutputAllowlistError("output_owner_mismatch")
        identity = _hash_descriptor(descriptor)
        final = os.fstat(descriptor)
        _require_regular_entry(final)
        if not _same_file(initial, final) or initial.st_size != final.st_size:
            raise OutputAllowlistError("output_race_replaced")
        if identity.byte_count != final.st_size:
            raise OutputAllowlistError("output_race_replaced")
        return descriptor, identity
    except Exception:
        os.close(descriptor)
        raise


def _stage_verified_file(
    staging_fd: int,
    item: _VerifiedFile,
    created_directories: dict[str, _TrackedEntry],
) -> _TrackedEntry:
    display_path = item.declaration.display_path
    parent_fd = _create_directory_path(
        staging_fd,
        display_path.split("/")[:-1],
        created_directories,
    )
    name = display_path.rsplit("/", maxsplit=1)[-1]
    try:
        descriptor = os.open(
            name,
            os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW | os.O_CLOEXEC,
            0o600,
            dir_fd=parent_fd,
        )
    except OSError as exc:
        raise OutputAllowlistError("output_publication_failed") from exc
    finally:
        os.close(parent_fd)
    staged_entry: _TrackedEntry | None = None
    try:
        initial = os.fstat(descriptor)
        _require_regular_entry(initial)
        staged_entry = _tracked_entry(display_path, initial)
        copied_identity = _copy_descriptor(item.descriptor, descriptor)
        copied_stat = os.fstat(descriptor)
        _require_regular_entry(copied_stat)
        if (
            copied_identity != item.identity
            or copied_stat.st_size != item.identity.byte_count
            or not _same_tracked_entry(copied_stat, staged_entry)
        ):
            raise OutputAllowlistError("output_race_replaced")
        os.fsync(descriptor)
        assert staged_entry is not None
        return staged_entry
    except OutputAllowlistError:
        if staged_entry is not None:
            _unlink_tracked_regular(
                staging_fd,
                staged_entry,
                created_directories,
            )
        raise
    except OSError as exc:
        if staged_entry is not None:
            _unlink_tracked_regular(
                staging_fd,
                staged_entry,
                created_directories,
            )
        raise OutputAllowlistError("output_publication_failed") from exc
    finally:
        os.close(descriptor)


def _publish_staged_file(
    staging_fd: int,
    destination_fd: int,
    display_path: str,
    destination_directories: dict[str, _TrackedEntry],
) -> _TrackedEntry:
    parents = display_path.split("/")[:-1]
    name = display_path.rsplit("/", maxsplit=1)[-1]
    staging_parent = _open_existing_directory_path(staging_fd, parents)
    destination_parent = _create_directory_path(
        destination_fd,
        parents,
        destination_directories,
    )
    linked = False
    published_entry: _TrackedEntry | None = None
    try:
        staged = _stat_at(staging_parent, name)
        _require_regular_entry(staged)
        published_entry = _tracked_entry(display_path, staged)
        os.link(
            name,
            name,
            src_dir_fd=staging_parent,
            dst_dir_fd=destination_parent,
            follow_symlinks=False,
        )
        linked = True
        published = _stat_at(destination_parent, name)
        _require_regular_entry(published, reject_hardlink=False)
        if not _same_tracked_entry(published, published_entry):
            raise OutputAllowlistError("output_race_replaced")
        assert published_entry is not None
        return published_entry
    except OutputAllowlistError:
        if linked and published_entry is not None:
            _unlink_tracked_regular(
                destination_fd,
                published_entry,
                destination_directories,
            )
        raise
    except OSError as exc:
        if linked and published_entry is not None:
            _unlink_tracked_regular(
                destination_fd,
                published_entry,
                destination_directories,
            )
        raise OutputAllowlistError("output_publication_failed") from exc
    finally:
        os.close(destination_parent)
        os.close(staging_parent)


def _open_directory(path: str | os.PathLike[str]) -> int:
    try:
        descriptor = os.open(path, os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW | os.O_CLOEXEC)
    except OSError as exc:
        raise OutputAllowlistError("output_root_invalid") from exc
    try:
        _require_directory_entry(os.fstat(descriptor))
    except Exception:
        os.close(descriptor)
        raise
    return descriptor


def _open_directory_at(parent_fd: int, name: str, *, expected: os.stat_result | None = None) -> int:
    before = _stat_at(parent_fd, name) if expected is None else expected
    _require_directory_entry(before)
    try:
        descriptor = os.open(
            name,
            os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW | os.O_CLOEXEC,
            dir_fd=parent_fd,
        )
    except OSError as exc:
        raise OutputAllowlistError("output_source_invalid") from exc
    try:
        after = os.fstat(descriptor)
        _require_directory_entry(after)
        if not _same_file(before, after):
            raise OutputAllowlistError("output_race_replaced")
    except Exception:
        os.close(descriptor)
        raise
    return descriptor


def _open_regular_at(parent_fd: int, name: str, *, expected: os.stat_result) -> int:
    try:
        descriptor = os.open(
            name,
            os.O_RDONLY | os.O_NOFOLLOW | os.O_CLOEXEC,
            dir_fd=parent_fd,
        )
    except OSError as exc:
        raise OutputAllowlistError("output_source_invalid") from exc
    try:
        actual = os.fstat(descriptor)
        _require_regular_entry(actual)
        if not _same_file(expected, actual):
            raise OutputAllowlistError("output_race_replaced")
    except Exception:
        os.close(descriptor)
        raise
    return descriptor


def _open_existing_directory_path(root_fd: int, components: Sequence[str]) -> int:
    descriptor = os.dup(root_fd)
    try:
        for component in components:
            child = _open_directory_at(descriptor, component)
            os.close(descriptor)
            descriptor = child
        return descriptor
    except Exception:
        os.close(descriptor)
        raise


def _create_directory_path(
    root_fd: int,
    components: Sequence[str],
    created_directories: dict[str, _TrackedEntry],
) -> int:
    descriptor = os.dup(root_fd)
    path_components: list[str] = []
    try:
        for component in components:
            path_components.append(component)
            display_path = "/".join(path_components)
            existing = created_directories.get(display_path)
            if existing is None:
                try:
                    os.mkdir(component, 0o700, dir_fd=descriptor)
                except FileExistsError as exc:
                    raise OutputAllowlistError("output_race_replaced") from exc
                except OSError as exc:
                    raise OutputAllowlistError("output_publication_failed") from exc
                child = _open_directory_at(descriptor, component)
                created_directories[display_path] = _tracked_entry(
                    display_path,
                    os.fstat(child),
                )
            else:
                child = _open_directory_at(descriptor, component)
                if not _same_tracked_entry(os.fstat(child), existing):
                    os.close(child)
                    raise OutputAllowlistError("output_race_replaced")
            os.close(descriptor)
            descriptor = child
        return descriptor
    except Exception:
        os.close(descriptor)
        raise


def _hash_descriptor(descriptor: int) -> ArtifactIdentity:
    try:
        os.lseek(descriptor, 0, os.SEEK_SET)
        digest = hashlib.sha256()
        byte_count = 0
        while chunk := os.read(descriptor, _COPY_CHUNK_BYTES):
            digest.update(chunk)
            byte_count += len(chunk)
        return ArtifactIdentity(digest.hexdigest(), byte_count)
    except OSError as exc:
        raise OutputAllowlistError("output_source_invalid") from exc


def _copy_descriptor(source_fd: int, destination_fd: int) -> ArtifactIdentity:
    try:
        os.lseek(source_fd, 0, os.SEEK_SET)
        digest = hashlib.sha256()
        byte_count = 0
        while chunk := os.read(source_fd, _COPY_CHUNK_BYTES):
            _write_all(destination_fd, chunk)
            digest.update(chunk)
            byte_count += len(chunk)
        return ArtifactIdentity(digest.hexdigest(), byte_count)
    except OSError as exc:
        raise OutputAllowlistError("output_publication_failed") from exc


def _write_all(descriptor: int, value: bytes) -> None:
    offset = 0
    while offset < len(value):
        written = os.write(descriptor, value[offset:])
        if written <= 0:
            raise OutputAllowlistError("output_publication_failed")
        offset += written


def _stat_at(parent_fd: int, name: str) -> os.stat_result:
    try:
        return os.stat(name, dir_fd=parent_fd, follow_symlinks=False)
    except OSError as exc:
        raise OutputAllowlistError("output_source_invalid") from exc


def _require_directory_entry(value: os.stat_result) -> None:
    if stat.S_ISLNK(value.st_mode):
        raise OutputAllowlistError("output_symlink")
    if not stat.S_ISDIR(value.st_mode):
        raise OutputAllowlistError("output_type_invalid")


def _require_regular_entry(value: os.stat_result, *, reject_hardlink: bool = True) -> None:
    if stat.S_ISLNK(value.st_mode):
        raise OutputAllowlistError("output_symlink")
    if not stat.S_ISREG(value.st_mode):
        raise OutputAllowlistError("output_type_invalid")
    if reject_hardlink and value.st_nlink != 1:
        raise OutputAllowlistError("output_hardlink")


def _same_file(left: os.stat_result, right: os.stat_result) -> bool:
    return left.st_dev == right.st_dev and left.st_ino == right.st_ino


def _tracked_entry(display_path: str, value: os.stat_result) -> _TrackedEntry:
    return _TrackedEntry(
        display_path=display_path,
        device=value.st_dev,
        inode=value.st_ino,
    )


def _same_tracked_entry(value: os.stat_result, expected: _TrackedEntry) -> bool:
    return value.st_dev == expected.device and value.st_ino == expected.inode


def _try_stat_at(parent_fd: int, name: str) -> os.stat_result | None:
    try:
        return os.stat(name, dir_fd=parent_fd, follow_symlinks=False)
    except OSError:
        return None


def _open_tracked_parent(
    root_fd: int,
    components: Sequence[str],
    directories: Mapping[str, _TrackedEntry],
) -> int | None:
    try:
        descriptor = os.dup(root_fd)
    except OSError:
        return None
    path_components: list[str] = []
    for component in components:
        path_components.append(component)
        expected = directories.get("/".join(path_components))
        if expected is None:
            os.close(descriptor)
            return None
        try:
            child = _open_directory_at(descriptor, component)
        except OutputAllowlistError:
            os.close(descriptor)
            return None
        os.close(descriptor)
        descriptor = child
        if not _same_tracked_entry(os.fstat(descriptor), expected):
            os.close(descriptor)
            return None
    return descriptor


def _unlink_tracked_regular(
    root_fd: int,
    entry: _TrackedEntry,
    directories: Mapping[str, _TrackedEntry],
) -> bool:
    components = entry.display_path.split("/")
    parent_fd = _open_tracked_parent(root_fd, components[:-1], directories)
    if parent_fd is None:
        return False
    try:
        current = _try_stat_at(parent_fd, components[-1])
        if (
            current is None
            or not stat.S_ISREG(current.st_mode)
            or not _same_tracked_entry(current, entry)
        ):
            return False
        os.unlink(components[-1], dir_fd=parent_fd)
        return True
    except OSError:
        return False
    finally:
        os.close(parent_fd)


def _rmdir_tracked(
    root_fd: int,
    entry: _TrackedEntry,
    directories: Mapping[str, _TrackedEntry],
) -> bool:
    components = entry.display_path.split("/")
    parent_fd = _open_tracked_parent(root_fd, components[:-1], directories)
    if parent_fd is None:
        return False
    try:
        current = _try_stat_at(parent_fd, components[-1])
        if (
            current is None
            or not stat.S_ISDIR(current.st_mode)
            or not _same_tracked_entry(current, entry)
        ):
            return False
        os.rmdir(components[-1], dir_fd=parent_fd)
        return True
    except OSError:
        return False
    finally:
        os.close(parent_fd)


def _open_tracked_directory(root_fd: int, entry: _TrackedEntry) -> int | None:
    current = _try_stat_at(root_fd, entry.display_path)
    if (
        current is None
        or not stat.S_ISDIR(current.st_mode)
        or not _same_tracked_entry(current, entry)
    ):
        return None
    try:
        descriptor = _open_directory_at(
            root_fd,
            entry.display_path,
            expected=current,
        )
    except OutputAllowlistError:
        return None
    if not _same_tracked_entry(os.fstat(descriptor), entry):
        os.close(descriptor)
        return None
    return descriptor


def _discard_tracked_tree(
    root_fd: int,
    root_entry: _TrackedEntry,
    files: Sequence[_TrackedEntry],
    directories: Mapping[str, _TrackedEntry],
) -> None:
    tree_fd = _open_tracked_directory(root_fd, root_entry)
    if tree_fd is None:
        raise OutputAllowlistError("output_race_replaced")
    try:
        for entry in reversed(files):
            if not _unlink_tracked_regular(tree_fd, entry, directories):
                raise OutputAllowlistError("output_race_replaced")
        for entry in reversed(tuple(directories.values())):
            if not _rmdir_tracked(tree_fd, entry, directories):
                raise OutputAllowlistError("output_race_replaced")
    finally:
        os.close(tree_fd)
    if not _rmdir_tracked(root_fd, root_entry, {}):
        raise OutputAllowlistError("output_race_replaced")


def _rollback_tracked_tree(
    root_fd: int,
    root_entry: _TrackedEntry,
    files: Sequence[_TrackedEntry],
    directories: Mapping[str, _TrackedEntry],
) -> None:
    tree_fd = _open_tracked_directory(root_fd, root_entry)
    if tree_fd is None:
        return
    try:
        for entry in reversed(files):
            _unlink_tracked_regular(tree_fd, entry, directories)
        for entry in reversed(tuple(directories.values())):
            _rmdir_tracked(tree_fd, entry, directories)
    finally:
        os.close(tree_fd)
    _rmdir_tracked(root_fd, root_entry, {})


def _rollback_publication(
    destination_fd: int,
    files: Sequence[_TrackedEntry],
    directories: Mapping[str, _TrackedEntry],
) -> None:
    for entry in reversed(files):
        _unlink_tracked_regular(destination_fd, entry, directories)
    for entry in reversed(tuple(directories.values())):
        _rmdir_tracked(destination_fd, entry, directories)


def _close_verified_files(verified: Iterable[_VerifiedFile]) -> None:
    for item in verified:
        with contextlib.suppress(OSError):
            os.close(item.descriptor)


def _require_owner_uid(owner_uid: int) -> None:
    if type(owner_uid) is not int or owner_uid < 0:
        raise OutputAllowlistError("output_owner_invalid")


def _require_secure_primitives() -> None:
    if (
        not hasattr(os, "O_NOFOLLOW")
        or not hasattr(os, "O_DIRECTORY")
        or not hasattr(os, "O_CLOEXEC")
        or not _SECURE_DIR_FD_AVAILABLE
        or not _STAT_NOFOLLOW_AVAILABLE
    ):
        raise OutputAllowlistError("output_nofollow_unavailable")
