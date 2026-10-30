"""Session run edits using the stateless planner only when an edit is requested."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..package import PackageGraph
from ._bridge import EditHost, SessionEditReport

if TYPE_CHECKING:
    from .runs import Loader, RunPlan

_DOCUMENT = "word/document.xml"


class RunEdits(EditHost):
    def insert_run(
        self,
        paragraph: int,
        text: str,
        *,
        run: int | None = None,
        position: str = "after",
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert a run of ``text`` before or after run ``run`` of ``paragraph``;
        without ``run``, append it after the paragraph's last run."""
        return self._run_edit(paragraph, "insert", run, position, text, part, None)

    def delete_run(self, paragraph: int, run: int, *, part: str = _DOCUMENT) -> SessionEditReport:
        """Delete run ``run`` of ``paragraph``."""
        return self._run_edit(paragraph, "delete", run, "after", None, part, None)

    def _run_edit(
        self,
        paragraph: int,
        action: str,
        run: int | None,
        position: str,
        text: str | None,
        part: str,
        fingerprint: str | None,
    ) -> SessionEditReport:
        from .runs import _check_run_fields, _plan_run_edit, _target

        _check_run_fields(paragraph, action, run, position, text, part)

        def plan(graph: PackageGraph, load: Loader) -> RunPlan:
            return _plan_run_edit(
                graph,
                load,
                paragraph=paragraph,
                action=action,
                run=run,
                position=position,
                text=text,
                part=part,
                fingerprint=fingerprint,
            )

        return self._edit_part(
            "run",
            part,
            plan,
            failure_code="run_staged_verification_failed",
            detail={
                "target": _target(part, paragraph, run, action, position),
                "action": action,
            },
        )
