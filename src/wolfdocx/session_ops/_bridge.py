"""Run a file-to-file operation against a session's current package.

Every bridged session edit builds the stateless operation's own request and calls
the operation's own ``author_*`` function, so the session result is exactly the
file-to-file result. The operation reads the session's current package from a
private scratch directory and writes its output there; nothing reaches the
session's source or any destination before :meth:`~wolfdocx.Document.save`. A
refusal leaves the session unchanged.
"""

from __future__ import annotations

import tempfile
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any

from ..canonical import sha256_bytes

__all__ = ["SessionEditReport"]

_APPLIED = "applied"
_REFUSED = "refused"
_UNBOUND_FINGERPRINT = "0" * 64
# Report fields that only describe the scratch files an operation ran against.
_SCRATCH_FIELDS = frozenset(
    {
        "source",
        "destination",
        "source_sha256",
        "destination_sha256",
        "status",
        "passed",
        "code",
    }
)


@dataclass(frozen=True)
class SessionEditReport:
    """Content-safe result of one session edit.

    ``detail`` carries the stateless operation's own report fields, without the
    fields that only describe the scratch package it ran against.
    """

    status: str
    code: str | None
    operation: str
    changed_parts: tuple[str, ...] = ()
    removed_parts: tuple[str, ...] = ()
    detail: Mapping[str, Any] = field(default_factory=dict)

    @property
    def passed(self) -> bool:
        return self.status == _APPLIED

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "operation": self.operation,
            "passed": self.passed,
            "status": self.status,
            "code": self.code,
            "changed_parts": list(self.changed_parts),
            "removed_parts": list(self.removed_parts),
            "detail": dict(self.detail),
        }


@dataclass(frozen=True)
class Staged:
    """The scratch paths and digest one bridged operation request binds to."""

    source: Path
    destination: Path
    source_sha256: str


def bound_fingerprint(fingerprint: str | None) -> str:
    """A discovered fingerprint, or a well-formed one that binds to nothing.

    An absent target then refuses inside the operation, with the operation's own
    code, instead of failing request construction.
    """
    return fingerprint or _UNBOUND_FINGERPRINT


def operation_detail(report: Any) -> dict[str, Any]:
    return {key: value for key, value in report.to_dict().items() if key not in _SCRATCH_FIELDS}


def refused(operation: str, code: str) -> SessionEditReport:
    return SessionEditReport(_REFUSED, code, operation)


class EditHost:
    """Session hooks the edit mixins rely on; :class:`~wolfdocx.Document` provides them."""

    def _apply_operation(
        self,
        operation: str,
        build: Callable[[Staged], Any],
        author: Callable[[Any], Any],
    ) -> SessionEditReport:
        """Run ``author(build(staged))`` on the current package and adopt its output.

        ``build`` may raise the request's own validation errors; the session is then
        unchanged, as it is after a refused report.
        """
        current = self._materialize()
        with tempfile.TemporaryDirectory(prefix="wolfdocx-session-") as scratch:
            staged = Staged(
                Path(scratch, "session.docx"),
                Path(scratch, "edited.docx"),
                sha256_bytes(current),
            )
            staged.source.write_bytes(current)
            report = author(build(staged))
            detail = operation_detail(report)
            if not report.passed:
                return SessionEditReport(_REFUSED, report.code, operation, detail=detail)
            edited = staged.destination.read_bytes()
        return self._adopt(operation, edited, detail)

    def _apply_steps(
        self, operation: str, steps: Sequence[Callable[[], SessionEditReport]]
    ) -> SessionEditReport:
        """Apply several edits as one: a refusal or error in any step restores the
        state before the first, so the session is unchanged unless every step lands."""
        snapshot = self._snapshot()
        reports: list[SessionEditReport] = []
        try:
            for step in steps:
                report = step()
                if not report.passed:
                    self._restore(snapshot)
                    return replace(report, operation=operation)
                reports.append(report)
        except BaseException:
            self._restore(snapshot)
            raise
        return SessionEditReport(
            _APPLIED,
            None,
            operation,
            changed_parts=tuple(sorted({p for r in reports for p in r.changed_parts})),
            removed_parts=tuple(sorted({p for r in reports for p in r.removed_parts})),
            detail={"steps": [r.to_dict() for r in reports]},
        )
