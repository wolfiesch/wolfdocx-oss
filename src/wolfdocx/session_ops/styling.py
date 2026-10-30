"""Session text and paragraph styling through :func:`~wolfdocx.formatting.author_format`."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ._bridge import EditHost, SessionEditReport, Staged, bound_fingerprint

if TYPE_CHECKING:
    from ..formatting import FormatRequest

_DOCUMENT = "word/document.xml"


def _format_target(staged: Staged, paragraph: int, part: str) -> dict[str, Any]:
    from ..formatting import discover_format_targets

    targets: list[dict[str, Any]] = discover_format_targets(staged.source, part=part).get(
        "targets", []
    )
    return next((target for target in targets if target["paragraph"] == paragraph), {})


class StylingEdits(EditHost):
    def format_text(
        self,
        paragraph: int,
        start: int,
        end: int,
        *,
        expected_old_text: str | None = None,
        part: str = _DOCUMENT,
        **formatting: Any,
    ) -> SessionEditReport:
        """Apply run formatting to visible characters ``[start, end)`` of ``paragraph``.

        ``formatting`` takes the :class:`~wolfdocx.formatting.RunFormatEdit` fields
        (``bold``, ``italic``, ``underline``, ``color``, ``font_size``, ``style``, and
        the rest). ``expected_old_text`` is an optional guard; when omitted it is the
        session's current text of that range.
        """
        from ..formatting import FormatRequest, RunFormatEdit, author_format

        def build(staged: Staged) -> FormatRequest:
            target = _format_target(staged, paragraph, part)
            text = target.get("text", "")
            return FormatRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                run_edits=(
                    RunFormatEdit(
                        paragraph=paragraph,
                        start=start,
                        end=end,
                        expected_old_text=(
                            text[start:end] if expected_old_text is None else expected_old_text
                        ),
                        structure_fingerprint_sha256=bound_fingerprint(
                            target.get("structure_fingerprint_sha256")
                        ),
                        part=part,
                        **formatting,
                    ),
                ),
            )

        return self._apply_operation("format", build, author_format)

    def format_paragraph(
        self, paragraph: int, *, part: str = _DOCUMENT, **formatting: Any
    ) -> SessionEditReport:
        """Apply paragraph formatting to ``paragraph``.

        ``formatting`` takes the :class:`~wolfdocx.formatting.ParagraphFormatEdit`
        fields (``style``, ``alignment``, ``spacing``, ``indent``, ``keep_next``, and
        the rest).
        """
        from ..formatting import FormatRequest, ParagraphFormatEdit, author_format

        def build(staged: Staged) -> FormatRequest:
            target = _format_target(staged, paragraph, part)
            return FormatRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                paragraph_edits=(
                    ParagraphFormatEdit(
                        paragraph=paragraph,
                        structure_fingerprint_sha256=bound_fingerprint(
                            target.get("structure_fingerprint_sha256")
                        ),
                        part=part,
                        **formatting,
                    ),
                ),
            )

        return self._apply_operation("format", build, author_format)
