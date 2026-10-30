"""Edits a :class:`~wolfdocx.Document` session applies through the operation modules.

Each mixin adds one family of session edits. Bridged edits run the stateless
operation itself against the session's current package; run edits plan in memory
with the same planner :func:`~wolfdocx.session_ops.runs.author_run` publishes.
"""

from importlib import import_module
from typing import TYPE_CHECKING, Any

from ._bridge import SessionEditReport
from .lists import ListEdits
from .paragraphs import ParagraphEdits
from .run_edits import RunEdits
from .sections import SectionEdits
from .styling import StylingEdits
from .tables import TableEdits

if TYPE_CHECKING:
    from .runs import RunReport, RunRequest, author_run

__all__ = [
    "ListEdits",
    "ParagraphEdits",
    "RunEdits",
    "RunReport",
    "RunRequest",
    "SectionEdits",
    "SessionEditReport",
    "StylingEdits",
    "TableEdits",
    "author_run",
]


def __getattr__(name: str) -> Any:
    if name in {"RunReport", "RunRequest", "author_run"}:
        value = getattr(import_module(".runs", __name__), name)
        globals()[name] = value
        return value
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))
