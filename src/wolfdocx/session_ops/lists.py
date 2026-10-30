"""Session list items through :func:`~wolfdocx.numbering_authoring.author_numbering`.

A paragraph becomes a list item by taking an existing numbering instance
(``num_id``) at a ``level``; it stops being one when its numbering is removed.
Paragraphs are 1-based over every ``w:p`` of ``word/document.xml``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ._bridge import SessionEditReport, Staged, bound_fingerprint
from .paragraphs import ParagraphEdits

if TYPE_CHECKING:
    from ..numbering_authoring import NumberingRequest

_DOCUMENT = "word/document.xml"


def _numbering_fingerprint(staged: Staged, paragraph: int) -> str:
    from ..numbering_authoring import discover_numbering_targets

    discovered = discover_numbering_targets(staged.source)
    for target in (*discovered.get("targets", ()), *discovered.get("unsupported", ())):
        if target["paragraph"] == paragraph:
            return bound_fingerprint(target["structure_fingerprint_sha256"])
    return bound_fingerprint(None)


class ListEdits(ParagraphEdits):
    def apply_list(self, paragraph: int, num_id: int, *, level: int = 0) -> SessionEditReport:
        """Make ``paragraph`` a list item of numbering instance ``num_id`` at ``level``."""
        return self._numbering(paragraph, action="apply", num_id=num_id, level=level)

    def remove_list(self, paragraph: int) -> SessionEditReport:
        """Remove ``paragraph`` from its list."""
        return self._numbering(paragraph, action="remove")

    def restart_list(
        self,
        paragraph: int,
        *,
        num_id: int | None = None,
        abstract_num_id: int | None = None,
        level: int = 0,
        start_override: int | None = None,
    ) -> SessionEditReport:
        """Restart numbering at ``paragraph`` through a new numbering instance."""
        return self._numbering(
            paragraph,
            action="restart",
            num_id=num_id,
            abstract_num_id=abstract_num_id,
            level=level,
            start_override=start_override,
        )

    def insert_list_item(
        self,
        paragraph: int,
        text: str,
        *,
        num_id: int,
        level: int = 0,
        position: str = "after",
        style: str | None = None,
    ) -> SessionEditReport:
        """Insert a paragraph next to ``paragraph`` and make it a list item.

        Both steps land or neither does.
        """
        inserted = paragraph + 1 if position == "after" else paragraph
        return self._apply_steps(
            "list-item-insert",
            (
                lambda: self.insert_paragraph(paragraph, text, position=position, style=style),
                lambda: self.apply_list(inserted, num_id, level=level),
            ),
        )

    def _numbering(self, paragraph: int, **fields: Any) -> SessionEditReport:
        from ..numbering_authoring import (
            NumberingEdit,
            NumberingRequest,
            author_numbering,
        )

        def build(staged: Staged) -> NumberingRequest:
            return NumberingRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                edits=(
                    NumberingEdit(
                        paragraph=paragraph,
                        structure_fingerprint_sha256=_numbering_fingerprint(staged, paragraph),
                        part=_DOCUMENT,
                        **fields,
                    ),
                ),
            )

        return self._apply_operation("list", build, author_numbering)
