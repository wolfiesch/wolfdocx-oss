"""Read, write, and edit DOCX packages with the Python standard library."""

from .document import Document
from .model import DocumentVisitor, Node, NodeEditError, NodeType, StoryType, VisitorAction
from .package import PackageError, PackageGraph, PackageLimits
from .session_ops import SessionEditReport
from .session_ops.rich_objects import SessionSaveReport

__version__ = "0.1.0"

__all__ = [
    "Document",
    "DocumentVisitor",
    "Node",
    "NodeEditError",
    "NodeType",
    "StoryType",
    "VisitorAction",
    "PackageError",
    "PackageGraph",
    "PackageLimits",
    "SessionEditReport",
    "SessionSaveReport",
    "__version__",
]
