"""Session sections, page setup, and headers and footers through their operation modules."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import TYPE_CHECKING, Any

from ._bridge import EditHost, SessionEditReport, Staged, bound_fingerprint

if TYPE_CHECKING:
    from ..page_setup import PageSetupRequest
    from ..regions import RegionParagraph, RegionRequest
    from ..section_break import SectionBreakRequest

_DOCUMENT = "word/document.xml"


class SectionEdits(EditHost):
    def insert_section_break(
        self, after: int, *, break_type: str = "nextPage", part: str = _DOCUMENT
    ) -> SessionEditReport:
        """End a section after body item ``after`` (a 1-based direct ``w:body`` child)."""
        return self._section_break(after, part, action="insert", break_type=break_type)

    def delete_section_break(self, paragraph: int, *, part: str = _DOCUMENT) -> SessionEditReport:
        """Remove the section break carried by body item ``paragraph``."""
        return self._section_break(paragraph, part, action="delete")

    def set_page_setup(
        self, section: int = 1, *, part: str = _DOCUMENT, **settings: Any
    ) -> SessionEditReport:
        """Change the page setup of ``section``; ``settings`` are the
        :class:`~wolfdocx.page_setup.PageSetupRequest` fields (``page_size``,
        ``orientation``, ``margins``, ``columns``, and the rest)."""
        from ..page_setup import (
            PageSetupRequest,
            author_page_setup,
            page_setup_structure_fingerprint,
        )

        def build(staged: Staged) -> PageSetupRequest:
            return PageSetupRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                structure_fingerprint_sha256=bound_fingerprint(
                    page_setup_structure_fingerprint(staged.source, section=section, part=part)
                ),
                section=section,
                part=part,
                **settings,
            )

        return self._apply_operation("page-setup", build, author_page_setup)

    def edit_header_footer(
        self,
        kind: str = "header",
        action: str = "create",
        *,
        section: int = 1,
        variant: str = "default",
        paragraphs: Sequence[RegionParagraph | Mapping[str, Any]] = (),
        link_to_previous: bool | None = None,
    ) -> SessionEditReport:
        """Create, replace, remove, link, or unlink one section's header or footer.

        ``kind`` is ``"header"`` or ``"footer"``; ``variant`` is ``"default"``,
        ``"first"``, or ``"even"``; ``paragraphs`` follow
        :class:`~wolfdocx.regions.RegionParagraph`.
        """
        from ..regions import RegionRequest, author_region

        def build(staged: Staged) -> RegionRequest:
            return RegionRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                section=section,
                kind=kind,
                variant=variant,
                paragraphs=paragraphs,
                action=action,
                link_to_previous=link_to_previous,
            )

        return self._apply_operation("header-footer", build, author_region)

    def _section_break(self, item: int, part: str, **fields: Any) -> SessionEditReport:
        from ..section_break import (
            SectionBreakRequest,
            author_section_break,
            section_break_anchor_fingerprint,
        )

        def build(staged: Staged) -> SectionBreakRequest:
            return SectionBreakRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                insert_after=item,
                anchor_fingerprint_sha256=bound_fingerprint(
                    section_break_anchor_fingerprint(staged.source, insert_after=item, part=part)
                ),
                part=part,
                **fields,
            )

        return self._apply_operation("section-break", build, author_section_break)
