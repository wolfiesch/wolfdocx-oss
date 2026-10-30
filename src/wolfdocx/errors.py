"""Typed public error taxonomy and the deterministic CLI exit contract.

Every WolfDocx failure is classified into one stable category with a safe
machine code, and every CLI command selects its exit status through
:func:`select_exit` so equivalent API and CLI failures expose the same
classification. Failure data never carries paths, document content, raw tool
output, or exception strings.
"""

from __future__ import annotations

import subprocess
from dataclasses import dataclass
from enum import IntEnum, StrEnum, unique

from .canonical import safe_diagnostic_code


@unique
class ExitStatus(IntEnum):
    """Deterministic CLI exit codes.

    - ``SUCCESS``: the command completed with a positive outcome.
    - ``NEGATIVE_DECISION``: the command completed a decision or validation
      and the answer was negative, with no execution failure.
    - ``EXECUTION_FAILURE``: invalid input or precondition, missing or
      unreadable required artifacts, unavailable or failed tools, timeouts,
      or internal execution failures.
    """

    SUCCESS = 0
    NEGATIVE_DECISION = 1
    EXECUTION_FAILURE = 2


@unique
class ErrorCategory(StrEnum):
    """Public failure categories; all map to :attr:`ExitStatus.EXECUTION_FAILURE`."""

    INVALID_INPUT = "invalid_input"
    MISSING_ARTIFACT = "missing_artifact"
    UNREADABLE_ARTIFACT = "unreadable_artifact"
    PRECONDITION_FAILED = "precondition_failed"
    TOOL_UNAVAILABLE = "tool_unavailable"
    TOOL_FAILURE = "tool_failure"
    TIMEOUT = "timeout"
    INTERNAL_ERROR = "internal_error"

    @property
    def exit_status(self) -> ExitStatus:
        return ExitStatus.EXECUTION_FAILURE


def safe_code(value: object, category: ErrorCategory) -> str:
    """Return ``value`` only when it is a stable snake-case machine code.

    Anything else — including exception text, paths, and document content —
    collapses to the category's own stable code, so unsafe detail can never
    enter a public error envelope.
    """

    return safe_diagnostic_code(value, fallback=str(category))


@dataclass(frozen=True)
class Failure:
    """Immutable machine-readable failure with a stable safe code."""

    code: str
    category: ErrorCategory

    @property
    def exit_status(self) -> ExitStatus:
        return self.category.exit_status

    def to_dict(self) -> dict[str, object]:
        """Content-safe serialization: code, category, and exit status only."""

        return {
            "code": self.code,
            "category": str(self.category),
            "exit_code": int(self.exit_status),
        }


class WolfDocxError(Exception):
    """Typed WolfDocx execution failure carrying a :class:`Failure`.

    The exception never carries message content; its public surface is the
    immutable :attr:`failure` value and its stable code.
    """

    category: ErrorCategory = ErrorCategory.INTERNAL_ERROR

    def __init__(self, code: str | None = None) -> None:
        resolved = safe_code(code, self.category)
        super().__init__(resolved)
        self._code = resolved

    @property
    def code(self) -> str:
        return self._code

    @property
    def failure(self) -> Failure:
        return Failure(self._code, self.category)

    def to_dict(self) -> dict[str, object]:
        return self.failure.to_dict()


def _trusted_code_errors() -> tuple[type[Exception], ...]:
    """Domain exceptions whose stable machine codes can be exposed safely."""
    from .content import ContentPolicyError

    return WolfDocxError, ContentPolicyError


def classify_failure(error: Exception) -> Failure:
    """Classify any exception into the public taxonomy without leaking content.

    Custom machine codes are honored only on the explicit trusted hierarchy
    (:class:`WolfDocxError` and WolfDocx domain errors that carry a ``code``
    attribute). A foreign exception's ``code`` attribute or message text is
    never trusted and never becomes a public code: untrusted failures always
    reduce to their category's stable default code.
    """

    if isinstance(error, WolfDocxError):
        return error.failure
    if isinstance(error, (subprocess.TimeoutExpired, TimeoutError)):
        category = ErrorCategory.TIMEOUT
    elif isinstance(error, FileNotFoundError):
        category = ErrorCategory.MISSING_ARTIFACT
    elif isinstance(error, OSError):
        category = ErrorCategory.UNREADABLE_ARTIFACT
    elif isinstance(error, (TypeError, ValueError)):
        category = ErrorCategory.INVALID_INPUT
    else:
        category = ErrorCategory.INTERNAL_ERROR
    code: object = category
    if isinstance(error, _trusted_code_errors()):
        code = getattr(error, "code", category)
    return Failure(safe_code(code, category), category)


def failure_envelope(failure: Failure) -> dict[str, object]:
    """Build the public pathless, contentless JSON envelope for a failure."""

    return {
        "schema_version": 1,
        "passed": False,
        "execution_error": True,
        "error": failure.to_dict(),
    }


def select_exit(*, decision_passed: bool, execution_failure: bool = False) -> ExitStatus:
    """The one centralized exit selection used by every CLI command.

    A negative decision is never hidden as an execution failure: only an
    explicit execution failure yields :attr:`ExitStatus.EXECUTION_FAILURE`.
    """

    if execution_failure:
        return ExitStatus.EXECUTION_FAILURE
    if decision_passed:
        return ExitStatus.SUCCESS
    return ExitStatus.NEGATIVE_DECISION
