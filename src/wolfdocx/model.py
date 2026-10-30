from __future__ import annotations

import copy
import re
from collections.abc import Callable, Iterator, Sequence
from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree as ET

if TYPE_CHECKING:
    from .document import Document


@dataclass(frozen=True)
class Issue:
    code: str
    severity: str
    message: str
    part: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class PartInfo:
    name: str
    size: int
    compressed_size: int
    sha256: str
    content_type: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class StyleInfo:
    style_id: str
    style_type: str | None
    name: str | None
    based_on: str | None
    run_properties: dict[str, str]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class DocumentFingerprint:
    document_text_sha256: str = ""
    document_nontext_sha256: str = ""
    structure_sha256: str = ""
    styles_sha256: str = ""
    numbering_sha256: str = ""
    sections_sha256: str = ""
    headers_sha256: str = ""
    footers_sha256: str = ""
    comments_sha256: str = ""
    fields_sha256: str = ""
    drawings_sha256: str = ""
    content_controls_sha256: str = ""
    formatting_sha256: str = ""
    tables_sha256: str = ""
    features_sha256: str = ""
    paragraph_count: int = 0
    table_count: int = 0
    row_count: int = 0
    cell_count: int = 0
    feature_counts: dict[str, int] = field(default_factory=dict)
    paragraphs: list[dict[str, Any]] = field(default_factory=list)
    tables: list[dict[str, Any]] = field(default_factory=list)
    styles: list[StyleInfo] = field(default_factory=list)

    def to_dict(self, *, full: bool = False) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "document_text_sha256": self.document_text_sha256,
            "document_nontext_sha256": self.document_nontext_sha256,
            "structure_sha256": self.structure_sha256,
            "styles_sha256": self.styles_sha256,
            "numbering_sha256": self.numbering_sha256,
            "sections_sha256": self.sections_sha256,
            "headers_sha256": self.headers_sha256,
            "footers_sha256": self.footers_sha256,
            "comments_sha256": self.comments_sha256,
            "fields_sha256": self.fields_sha256,
            "drawings_sha256": self.drawings_sha256,
            "content_controls_sha256": self.content_controls_sha256,
            "formatting_sha256": self.formatting_sha256,
            "tables_sha256": self.tables_sha256,
            "features_sha256": self.features_sha256,
            "paragraph_count": self.paragraph_count,
            "table_count": self.table_count,
            "row_count": self.row_count,
            "cell_count": self.cell_count,
            "feature_counts": self.feature_counts,
        }
        if full:
            payload["paragraphs"] = self.paragraphs
            payload["tables"] = self.tables
            payload["styles"] = [style.to_dict() for style in self.styles]
        return payload


class NodeType(str, Enum):
    """Semantic nodes; XML preserves otherwise unclassified markup by name."""

    DOCUMENT = "document"
    STORY = "story"
    PARAGRAPH = "paragraph"
    RUN = "run"
    TABLE = "table"
    ROW = "row"
    CELL = "cell"
    SPECIAL_CHARACTER = "special_character"
    FIELD = "field"
    BOOKMARK = "bookmark"
    COMMENT = "comment"
    SHAPE = "shape"
    STRUCTURED_TAG = "structured_tag"
    SECTION = "section"
    HEADER_FOOTER = "header_footer"
    OFFICE_MATH = "office_math"
    SMART_TAG = "smart_tag"
    XML = "xml"


class StoryType(str, Enum):
    BODY = "body"
    HEADER = "header"
    FOOTER = "footer"
    FOOTNOTE = "footnote"
    ENDNOTE = "endnote"
    COMMENT = "comment"
    TEXT_BOX = "text_box"
    GLOSSARY = "glossary"
    FIRST_HEADER = "first_header"
    EVEN_HEADER = "even_header"
    FIRST_FOOTER = "first_footer"
    EVEN_FOOTER = "even_footer"
    FOOTNOTE_SEPARATOR = "footnote_separator"
    FOOTNOTE_CONTINUATION_SEPARATOR = "footnote_continuation_separator"
    FOOTNOTE_CONTINUATION_NOTICE = "footnote_continuation_notice"
    ENDNOTE_SEPARATOR = "endnote_separator"
    ENDNOTE_CONTINUATION_SEPARATOR = "endnote_continuation_separator"
    ENDNOTE_CONTINUATION_NOTICE = "endnote_continuation_notice"

    @property
    def header_footer_kind(self) -> str | None:
        if self in {self.HEADER, self.FIRST_HEADER, self.EVEN_HEADER}:
            return "header"
        if self in {self.FOOTER, self.FIRST_FOOTER, self.EVEN_FOOTER}:
            return "footer"
        return None


class VisitorAction(str, Enum):
    CONTINUE = "continue"
    SKIP_CHILDREN = "skip_children"
    STOP = "stop"


class NodeEditError(ValueError):
    """A byte-preserving edit refused without changing the package."""

    def __init__(self, code: str) -> None:
        self.code = code
        super().__init__(code)


class DocumentVisitor:
    """Override ``visit_<node_type>_start/end`` or the generic defaults.

    Both callbacks must return a :class:`VisitorAction`. SKIP_CHILDREN at start
    still calls that node's end; at end it is equivalent to CONTINUE. STOP ends
    the entire traversal immediately, without synthetic end callbacks.
    """

    def visit_node_start(self, node: Node) -> VisitorAction:
        return VisitorAction.CONTINUE

    def visit_node_end(self, node: Node) -> VisitorAction:
        return VisitorAction.CONTINUE

    def visit_document_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_document_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_story_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_story_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_paragraph_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_paragraph_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_run_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_run_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_table_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_table_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_row_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_row_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_cell_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_cell_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_special_character_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_special_character_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_xml_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_xml_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_field_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_field_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_bookmark_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_bookmark_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_comment_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_comment_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_shape_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_shape_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_structured_tag_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_structured_tag_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_section_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_section_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_header_footer_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_header_footer_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_office_math_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_office_math_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)

    def visit_smart_tag_start(self, node: Node) -> VisitorAction:
        return self.visit_node_start(node)

    def visit_smart_tag_end(self, node: Node) -> VisitorAction:
        return self.visit_node_end(node)


class ChildNodeCollection(Sequence["Node"]):
    """Live direct children. Mutation uses the owner's atomic structural gate."""

    def __init__(self, parent: Node) -> None:
        self._parent = parent

    def __len__(self) -> int:
        return len(self._parent._children)

    def __getitem__(self, index: int | slice) -> Node | tuple[Node, ...]:
        return self._parent._children[index]

    def __iter__(self) -> Iterator[Node]:
        return iter(self._parent._children)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, ChildNodeCollection):
            return self._parent._children == other._parent._children
        return self._parent._children == other

    def add(self, child: Node) -> Node:
        return self._parent.append_child(child)

    def insert(self, index: int, child: Node) -> Node:
        return self._parent.insert_child(index, child)

    def remove(self, child: Node) -> Node:
        return self._parent.remove_child(child)

    def clear(self) -> None:
        self._parent.remove_all_children()


class Node:
    """A navigable package snapshot supporting validated generic tree edits.

    ``children``, ``paragraphs``, ``runs`` and ``tables`` are direct children.
    Use ``get_child_nodes(type, deep=True)`` for descendants in preorder.
    Glossary metadata and entries are XML nodes; each docPartBody is a GLOSSARY
    story whose content uses the same semantic nodes and mutation APIs.
    ``clone`` creates an independent, detached subtree. Insert detached nodes
    through their destination parent; attached nodes must be removed first.
    Structural edits persist through the owning Document and keep this tree's
    navigation coherent. Other Document operations invalidate the snapshot.
    Run setters preserve bytes outside text payloads until a structural edit.
    ``custom_node_id`` is a caller-assigned integer, initially zero. It stays on
    the handle through detach/reinsert and is never written to the package.
    All clones, including document clones and their descendants, start at zero.
    Complex field markers are ``FieldStart``, ``FieldSeparator`` and ``FieldEnd``
    nodes, all selected by ``NodeType.FIELD``. Their marker flags and immutable
    logical-field snapshots are available without a second navigation tree.
    """

    def __init__(
        self,
        node_type: NodeType,
        *,
        part: str | None = None,
        element: ET.Element | None = None,
        story_type: StoryType | None = None,
        rewrite: Callable[[ET.Element, str], None] | None = None,
        edit: Callable[[str, Node, Node, int], None] | None = None,
        clone_document: Callable[[bool], Node] | None = None,
        document: Document | None = None,
    ) -> None:
        self._node_type = node_type
        self._part = part
        self._element = element
        self._story_type = story_type
        self._rewrite = rewrite
        self._edit = edit
        self._clone_document = clone_document
        self._document = document
        self._live = document is not None
        self._parent: Node | None = None
        self._children: tuple[Node, ...] = ()
        self._index = 0
        self._custom_node_id = 0
        self._collection = ChildNodeCollection(self)
        self._bind: Callable[[ET.Element, str, StoryType | None], Node] | None = None

    @property
    def node_type(self) -> NodeType:
        return self._node_type

    @property
    def custom_node_id(self) -> int:
        """Lifetime-only caller metadata; clones and reopened nodes start at zero."""
        return self._custom_node_id

    @custom_node_id.setter
    def custom_node_id(self, value: int) -> None:
        if not isinstance(value, int) or isinstance(value, bool):
            raise NodeEditError("node_custom_id_invalid")
        self._custom_node_id = value

    @property
    def part(self) -> str | None:
        return self._part

    @property
    def xml_name(self) -> str | None:
        return None if self._element is None else self._element.tag

    @property
    def story_type(self) -> StoryType | None:
        return self._story_type

    @property
    def parent(self) -> Node | None:
        return self._parent

    @property
    def document(self) -> Document | None:
        """The owning session, including the independent session of a root clone."""
        return self._document

    @property
    def children(self) -> ChildNodeCollection:
        return self._collection

    @property
    def first_child(self) -> Node | None:
        return self.children[0] if self.children else None

    @property
    def last_child(self) -> Node | None:
        return self.children[-1] if self.children else None

    @property
    def previous_sibling(self) -> Node | None:
        if self.parent is None or self._index == 0:
            return None
        return self.parent.children[self._index - 1]

    @property
    def next_sibling(self) -> Node | None:
        if self.parent is None or self._index + 1 == len(self.parent.children):
            return None
        return self.parent.children[self._index + 1]

    def get_child_nodes(self, node_type: NodeType | None, deep: bool = False) -> tuple[Node, ...]:
        """Select direct children or preorder descendants; None selects all types."""
        if node_type is not None and not isinstance(node_type, NodeType):
            raise TypeError("node_type_invalid")
        if not isinstance(deep, bool):
            raise NodeEditError("node_depth_invalid")
        selected: list[Node] = []
        pending = list(reversed(self.children))
        while pending:
            node = pending.pop()
            if node_type is None or node.node_type == node_type:
                selected.append(node)
            if deep:
                pending.extend(reversed(node.children))
        return tuple(selected)

    def get_child(self, node_type: NodeType | None, index: int, deep: bool = False) -> Node | None:
        """Zero-based index in the selected children; return None past the end."""
        if not isinstance(index, int) or isinstance(index, bool) or index < 0:
            raise NodeEditError("node_index_invalid")
        nodes = self.get_child_nodes(node_type, deep)
        return nodes[index] if index < len(nodes) else None

    def get_ancestor(self, node_type: NodeType) -> Node | None:
        """Nearest matching parent, excluding this node."""
        if not isinstance(node_type, NodeType):
            raise NodeEditError("node_type_invalid")
        ancestor = self.parent
        while ancestor is not None:
            if ancestor.node_type == node_type:
                return ancestor
            ancestor = ancestor.parent
        return None

    def _check_traversal_root(self, root: Node) -> None:
        if not isinstance(root, Node):
            raise NodeEditError("node_traversal_root_invalid")
        ancestor: Node | None = self
        while ancestor is not None:
            if ancestor is root:
                return
            ancestor = ancestor.parent
        raise NodeEditError("node_traversal_root_not_ancestor")

    def next_pre_order(self, root: Node) -> Node | None:
        """Next live handle in preorder, never leaving the supplied subtree."""
        self._check_traversal_root(root)
        if self.first_child is not None:
            return self.first_child
        current = self
        while current is not root:
            if current.next_sibling is not None:
                return current.next_sibling
            assert current.parent is not None
            current = current.parent
        return None

    def previous_pre_order(self, root: Node) -> Node | None:
        """Previous live handle in preorder; the supplied root has no predecessor."""
        self._check_traversal_root(root)
        if self is root:
            return None
        current = self.previous_sibling
        if current is None:
            return self.parent
        while current.last_child is not None:
            current = current.last_child
        return current

    @property
    def paragraphs(self) -> tuple[Node, ...]:
        return self.get_child_nodes(NodeType.PARAGRAPH)

    @property
    def runs(self) -> tuple[Node, ...]:
        return self.get_child_nodes(NodeType.RUN)

    @property
    def tables(self) -> tuple[Node, ...]:
        return self.get_child_nodes(NodeType.TABLE)

    @property
    def special_character(self) -> str | None:
        if self.node_type != NodeType.SPECIAL_CHARACTER or self._element is None:
            return None
        return self._element.tag.rsplit("}", 1)[-1]

    @property
    def text(self) -> str:
        """Aggregate story text, or a run's direct text payload."""
        if self.node_type == NodeType.STORY:
            return self.get_text()
        if self.node_type != NodeType.RUN or self._element is None:
            raise TypeError("node_is_not_run")
        namespace = self._element.tag.rsplit("}", 1)[0] + "}"
        return "".join(child.text or "" for child in self._element if child.tag == namespace + "t")

    @text.setter
    def text(self, value: str) -> None:
        if self.node_type == NodeType.STORY:
            self.set_text(value)
            return
        if self.node_type != NodeType.RUN or self._element is None or self._rewrite is None:
            raise TypeError("node_is_not_run")
        self._rewrite(self._element, value)

    def clone(self, deep: bool = True) -> Node:
        """Copy this subtree without attaching it; a document clone owns a new session.

        Shallow clones retain attributes but no child XML or run text. They must
        be made structurally valid before insertion into a package.
        """
        if not isinstance(deep, bool):
            raise NodeEditError("node_clone_depth_invalid")
        if self._edit is not None and self.node_type != NodeType.DOCUMENT:
            self._edit("check", self, self, 0)
        if self.node_type == NodeType.DOCUMENT:
            if self._clone_document is None:
                raise NodeEditError("node_document_clone_invalid")
            return self._clone_document(deep)
        memo: dict[int, Any] = {}
        element = copy.deepcopy(self._element, memo)
        if not deep and element is not None:
            element[:] = []
            element.text = None

        def duplicate(source: Node) -> Node:
            target = type(source)(
                source.node_type,
                part=source.part,
                element=memo.get(id(source._element)),
                story_type=source.story_type,
                rewrite=source._rewrite,
                edit=source._edit,
                document=source.document,
            )
            target._live = False
            target._bind = source._bind
            if deep:
                target._attach(tuple(duplicate(child) for child in source.children))
            return target

        clone = duplicate(self)
        clone._element = element
        return clone

    def insert_child(self, index: int, child: Node) -> Node:
        """Insert a detached node at a direct-child index, or refuse atomically."""
        if not isinstance(child, Node):
            raise NodeEditError("node_child_invalid")
        if (
            not isinstance(index, int)
            or isinstance(index, bool)
            or not 0 <= index <= len(self.children)
        ):
            raise NodeEditError("node_index_invalid")
        if child.parent is not None or child._live:
            raise NodeEditError("node_child_attached")
        ancestor: Node | None = self
        while ancestor is not None:
            if ancestor is child:
                raise NodeEditError("node_cycle")
            ancestor = ancestor.parent
        if self._edit is None:
            raise NodeEditError("node_edit_unbound")
        self._edit("insert", self, child, index)
        return child

    def append_child(self, child: Node) -> Node:
        """Append a detached node before final body or glossary section data."""
        if not isinstance(child, Node):
            raise NodeEditError("node_child_invalid")
        index = len(self.children)
        if (
            self.story_type in {StoryType.BODY, StoryType.GLOSSARY}
            and self.children
            and self.children[-1].xml_name is not None
            and self.children[-1].xml_name.endswith("}sectPr")
            and (child.xml_name is None or not child.xml_name.endswith("}sectPr"))
        ):
            index -= 1
        return self.insert_child(index, child)

    def insert_before(self, child: Node, reference: Node) -> Node:
        if reference not in self.children:
            raise NodeEditError("node_reference_invalid")
        return self.insert_child(reference._index, child)

    def insert_after(self, child: Node, reference: Node) -> Node:
        if reference not in self.children:
            raise NodeEditError("node_reference_invalid")
        return self.insert_child(reference._index + 1, child)

    def remove_child(self, child: Node) -> Node:
        """Detach a direct child after validating the remaining package structure."""
        if child not in self.children:
            raise NodeEditError("node_reference_invalid")
        if self._edit is None:
            raise NodeEditError("node_edit_unbound")
        self._edit("remove", self, child, child._index)
        return child

    def remove(self) -> Node:
        if self.parent is None:
            raise NodeEditError("node_parent_missing")
        return self.parent.remove_child(self)

    def remove_all_children(self) -> None:
        """Clear children atomically; required story/cell blocks refuse unchanged."""
        if self._edit is None:
            raise NodeEditError("node_edit_unbound")
        self._edit("clear", self, self, 0)

    def select_nodes(self, path: str, namespaces: dict[str, str] | None = None) -> tuple[Node, ...]:
        """Local XPath 1.0 over the current tree, returning live handles in preorder.

        The document has a synthetic, unnamespaced ``document`` element containing
        its stories in navigation order. Detached subtrees are independent XPath
        roots. Standard ``w``, ``r`` and ``m`` prefixes are supplied; pass additional
        namespace bindings as needed. Text payloads, attributes, namespaces and
        scalar results refuse. Only XPath 1.0 built-ins are callable, with no
        extensions, variables, entity resolution or external resource loading.
        Install lxml (also included in the signatures extra) to use selection.
        """
        if not isinstance(path, str) or not path:
            raise NodeEditError("node_path_invalid")
        if namespaces is not None and (
            not isinstance(namespaces, dict)
            or any(
                not isinstance(key, str) or not key or not isinstance(value, str) or not value
                for key, value in namespaces.items()
            )
        ):
            raise NodeEditError("node_path_namespace_invalid")
        # Ignore quoted literals when recognizing function calls. An explicit
        # allowlist also excludes globally registered lxml extension functions.
        tokens = re.sub(r"""('[^']*'|"[^"]*")""", "0", path)
        tokens = re.sub(r"\b[\w-]+::", "", tokens)
        builtins = {
            "last",
            "position",
            "count",
            "id",
            "local-name",
            "namespace-uri",
            "name",
            "string",
            "concat",
            "starts-with",
            "contains",
            "substring-before",
            "substring-after",
            "substring",
            "string-length",
            "normalize-space",
            "translate",
            "boolean",
            "not",
            "true",
            "false",
            "lang",
            "number",
            "sum",
            "floor",
            "ceiling",
            "round",
            "node",
            "text",
            "comment",
            "processing-instruction",
        }
        for call in re.finditer(r"""([^\s/()\[\]@,=<>+*|'"$]+)\s*\(""", tokens):
            name = call.group(1)
            if name == "-":  # Unary minus before a grouped expression.
                continue
            previous = tokens[: call.start()].rstrip()
            # XPath word operators may precede a grouped operand. At a primary
            # expression position, the same spelling is an extension function.
            if (
                name in {"and", "or", "mod", "div"}
                and previous
                and (previous[-1].isalnum() or previous[-1] in "_-.)]*")
            ):
                continue
            if name not in builtins:
                raise NodeEditError("node_path_function_unsupported")
        try:
            from lxml import etree
        except ImportError as exc:
            raise NodeEditError("node_path_dependency_missing") from exc
        from .inspect import serialize_xml_with_preserved_qnames

        top = self
        while top.parent is not None:
            top = top.parent
        nodes = (top,) + top.get_child_nodes(None, deep=True)
        handles = {id(node._element): node for node in nodes if node._element is not None}
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        mapped: dict[Any, Node] = {}

        def mirror(node: Node) -> Any:
            assert node._element is not None
            copied = etree.fromstring(serialize_xml_with_preserved_qnames(node._element), parser)
            for original, duplicate in zip(node._element.iter(), copied.iter(), strict=True):
                handle = handles.get(id(original))
                if handle is not None:
                    mapped[duplicate] = handle
            return copied

        if top.node_type == NodeType.DOCUMENT:
            tree = etree.Element("document")
            mapped[tree] = top
            for story in top.children:
                tree.append(mirror(story))
        elif top._element is not None:
            tree = mirror(top)
        else:
            raise NodeEditError("node_path_invalid")
        context = next((element for element, node in mapped.items() if node is self), None)
        if context is None:
            raise NodeEditError("node_path_payload_not_node")
        word_namespace = next(
            (
                node.xml_name[1:].partition("}")[0]
                for node in nodes
                if node.xml_name
                and node.xml_name.startswith("{")
                and node.xml_name[1:].partition("}")[0]
                in {
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "http://purl.oclc.org/ooxml/wordprocessingml/main",
                }
            ),
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        )
        strict = word_namespace.startswith("http://purl.oclc.org/")
        bindings = {
            "w": word_namespace,
            "r": (
                "http://purl.oclc.org/ooxml/officeDocument/relationships"
                if strict
                else "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            ),
            "m": (
                "http://purl.oclc.org/ooxml/officeDocument/math"
                if strict
                else "http://schemas.openxmlformats.org/officeDocument/2006/math"
            ),
        }
        if namespaces is not None:
            bindings.update(namespaces)
        try:
            expression = etree.XPath(path, namespaces=bindings, extensions={}, regexp=False)
            results = expression(context)
        except (etree.XPathError, TypeError, ValueError) as exc:
            raise NodeEditError("node_path_invalid") from exc
        if not isinstance(results, list):
            raise NodeEditError("node_path_payload_not_node")
        if any(not isinstance(item, etree._Element) or item not in mapped for item in results):
            raise NodeEditError("node_path_payload_not_node")
        selected = {mapped[item] for item in results}
        return tuple(node for node in nodes if node in selected)

    def select_single_node(
        self, path: str, namespaces: dict[str, str] | None = None
    ) -> Node | None:
        return next(iter(self.select_nodes(path, namespaces)), None)

    def get_text(self) -> str:
        """OOXML display text; all enclosing field instructions hide operand caches."""
        if self._element is None:
            return "".join(child.get_text() for child in self.children)

        def extract(item: ET.Element, state: _FieldTextState) -> str:
            if not isinstance(item.tag, str):
                return ""
            local = item.tag.rpartition("}")[2]
            if local in {"instrText", "delInstrText", "pPr", "rPr", "sectPr"}:
                return ""
            if local == "fldChar":
                state.marker(item)
                return ""
            if local in {"txbxContent", "footnote", "endnote", "comment"}:
                nested = _FieldTextState("node_field_incomplete", NodeEditError)
                text = "".join(extract(child, nested) for child in item)
                nested.finish()
                return text if state.visible else ""
            if local in {"t", "delText", "tab", "ptab", "br", "cr", "noBreakHyphen", "softHyphen"}:
                if not state.visible:
                    return ""
                return {
                    "t": item.text or "",
                    "delText": item.text or "",
                    "tab": "\t",
                    "ptab": "\t",
                    "br": "\n",
                    "cr": "\n",
                    "noBreakHyphen": "\u2011",
                    "softHyphen": "\u00ad",
                }[local]
            text = "".join(extract(child, state) for child in item)
            return text + ({"p": "\r", "tc": "\x07"}.get(local, ""))

        state = _FieldTextState("node_field_incomplete", NodeEditError)
        text = extract(self._element, state)
        state.finish()
        return text

    def to_string(self, save_format: str = "txt") -> str:
        """Return plain text or the node's XML representation."""
        if save_format == "txt":
            return self.get_text().replace("\r", "\n").replace("\x07", "\t")
        if save_format == "xml" and self._element is not None:
            from .inspect import serialize_xml_with_preserved_qnames

            return serialize_xml_with_preserved_qnames(self._element).decode("utf-8")
        raise NodeEditError("node_string_format_unsupported")

    

    def set_text(self, text: str) -> None:
        """Replace one story's content with plain paragraphs, retaining its identity."""
        if self.node_type != NodeType.STORY or self._element is None:
            raise NodeEditError("node_story_required")
        if self._edit is None or self._bind is None or self.part is None:
            raise NodeEditError("node_edit_unbound")
        if not isinstance(text, str):
            raise NodeEditError("story_text_invalid")
        from .authoring import _MAX_TEXT_CHARACTERS
        from .paragraph_authoring import _is_valid_xml_text

        if len(text) > _MAX_TEXT_CHARACTERS:
            raise NodeEditError("story_text_limit_exceeded")

        if not _is_valid_xml_text(text):
            raise NodeEditError("story_text_invalid")
        namespace = self._element.tag[1:].partition("}")[0]
        payload = copy.copy(self._element)
        payload[:] = []
        # Final section properties are structural metadata, not story text.
        if self.story_type in {StoryType.BODY, StoryType.GLOSSARY}:
            payload.extend(
                copy.deepcopy(child)
                for child in self._element
                if child.tag == f"{{{namespace}}}sectPr"
            )
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            paragraph = ET.Element(f"{{{namespace}}}p")
            run = ET.SubElement(paragraph, f"{{{namespace}}}r")
            chunks = line.split("\t")
            for index, chunk in enumerate(chunks):
                if index:
                    ET.SubElement(run, f"{{{namespace}}}tab")
                item = ET.SubElement(run, f"{{{namespace}}}t")
                item.text = chunk
                item.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            payload.insert(
                len(payload)
                - int(
                    self.story_type in {StoryType.BODY, StoryType.GLOSSARY}
                    and len(payload) > 0
                    and payload[-1].tag == f"{{{namespace}}}sectPr"
                ),
                paragraph,
            )
        replacement = self._bind(payload, self.part, self.story_type)
        self._edit("replace-content", self, replacement, 0)

    def append_paragraph(self, text: str = "") -> Node:
        """Append a plain paragraph to a story using the same atomic node gate."""
        from .authoring import _MAX_TEXT_CHARACTERS
        from .paragraph_authoring import _is_valid_xml_text

        if not isinstance(text, str) or not _is_valid_xml_text(text):
            raise NodeEditError("story_text_invalid")
        if len(text) > _MAX_TEXT_CHARACTERS:
            raise NodeEditError("story_text_limit_exceeded")
        if (
            self.node_type != NodeType.STORY
            or self._element is None
            or self._bind is None
            or self.part is None
        ):
            raise NodeEditError("node_story_required")
        namespace = self._element.tag[1:].partition("}")[0]
        paragraph = self._bind(ET.Element(f"{{{namespace}}}p"), self.part, self.story_type)
        element = ET.Element(f"{{{namespace}}}r")
        payload = ET.SubElement(element, f"{{{namespace}}}t")
        payload.text = text
        payload.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        if self._bind is None or self.part is None:
            raise NodeEditError("node_edit_unbound")
        paragraph.append_child(self._bind(element, self.part, self.story_type))
        return self.append_child(paragraph)

    def _attach(self, children: tuple[Node, ...]) -> None:
        for child in self._children:
            if child not in children:
                child._parent = None
                child._index = 0
        self._children = children
        for index, child in enumerate(children):
            child._parent = self
            child._index = index

    def accept(self, visitor: DocumentVisitor) -> VisitorAction:
        """Walk preorder/start and postorder/end; see Document.accept for stories."""
        pending = [(self, False)]
        while pending:
            node, ending = pending.pop()
            phase = "end" if ending else "start"
            method = getattr(visitor, f"visit_{node.node_type.value}_{phase}")
            action = method(node)
            if not isinstance(action, VisitorAction):
                raise TypeError("visitor_action_invalid")
            if action == VisitorAction.STOP:
                return action
            if not ending:
                pending.append((node, True))
                if action == VisitorAction.CONTINUE:
                    pending.extend((child, False) for child in reversed(node.children))
        return VisitorAction.CONTINUE


class FieldCharacter(Node):
    """A complex field marker with navigation and read-only character flags.

    Field instructions and cached text remain opaque package content. Clone,
    navigation, visitors and structural mutation use the normal Node API.
    """

    @property
    def field_character(self) -> str:
        if self._element is None:
            raise NodeEditError("node_field_incomplete")
        namespace = self._element.tag.rpartition("}")[0] + "}"
        kind = self._element.get(namespace + "fldCharType")
        if kind not in {"begin", "separate", "end"}:
            raise NodeEditError("node_field_incomplete")
        return kind

    def _field_flag(self, name: str) -> bool:
        if self._element is None:
            raise NodeEditError("node_field_incomplete")
        namespace = self._element.tag.rpartition("}")[0] + "}"
        return self._element.get(namespace + name) in {"true", "1", "on"}

    @property
    def dirty(self) -> bool:
        return self._field_flag("dirty")

    @property
    def locked(self) -> bool:
        return self._field_flag("fldLock")

    

    

    

    


class FieldStart(FieldCharacter):
    """The ``w:fldCharType="begin"`` character; selected as ``NodeType.FIELD``."""


class FieldSeparator(FieldCharacter):
    """The ``w:fldCharType="separate"`` character before the cached display result."""


class FieldEnd(FieldCharacter):
    """The ``w:fldCharType="end"`` character closing a complex field."""


class _FieldTextState:
    """Instruction visibility shared by the node and package plain-text exporters."""

    def __init__(self, code: str, error: type[ValueError]) -> None:
        self._frames: list[bool] = []
        self._code = code
        self._error = error

    @property
    def visible(self) -> bool:
        return all(self._frames)

    def marker(self, element: ET.Element) -> None:
        namespace = element.tag.rpartition("}")[0] + "}"
        kind = element.get(namespace + "fldCharType")
        if kind == "begin":
            self._frames.append(False)
        elif kind == "separate" and self._frames and not self._frames[-1]:
            self._frames[-1] = True
        elif kind == "end" and self._frames:
            self._frames.pop()
        else:
            raise self._error(self._code)

    def finish(self) -> None:
        if self._frames:
            raise self._error(self._code)
