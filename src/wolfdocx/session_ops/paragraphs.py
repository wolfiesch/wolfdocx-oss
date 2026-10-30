"""Session paragraph edits through the stateless operation's in-memory planner."""

from __future__ import annotations

from ._bridge import EditHost, SessionEditReport

_DOCUMENT = "word/document.xml"


class ParagraphEdits(EditHost):
    """Paragraph edits addressed as :class:`~wolfdocx.ParagraphRequest` addresses them:
    the 1-based ``paragraph`` counts every ``w:p`` of ``part`` in document order."""

    def insert_paragraph(
        self,
        paragraph: int,
        text: str = "",
        *,
        position: str = "after",
        style: str | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert a paragraph before or after ``paragraph``."""
        return self._paragraph(
            paragraph, part, action="insert", position=position, text=text, style=style
        )

    def delete_paragraph(self, paragraph: int, *, part: str = _DOCUMENT) -> SessionEditReport:
        """Delete ``paragraph``."""
        return self._paragraph(paragraph, part, action="delete")

    def split_paragraph(
        self, paragraph: int, offset: int, *, part: str = _DOCUMENT
    ) -> SessionEditReport:
        """Split ``paragraph`` at the visible-text ``offset``."""
        return self._paragraph(paragraph, part, action="split", offset=offset)

    def _paragraph(self, paragraph: int, part: str, **fields: object) -> SessionEditReport:
        return self._edit_paragraph(paragraph, part, **fields)
