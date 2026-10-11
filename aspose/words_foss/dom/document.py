"""Explicit original-package editing entry point; legacy Document remains unchanged."""

from datetime import datetime

from aspose.words_foss._io import MAX_TABLE_COLUMNS
from aspose.words_foss._opc import resolve_target
from aspose.words_foss.dom.importing import ImportFormatMode
from aspose.words_foss.dom.nodes import (
    _NODE_CLASSES,
    Body,
    Node,
    Paragraph,
    Run,
    Table,
    UnknownNode,
    W,
    _elements,
    _find,
    _is,
    _new,
    _validate_text,
)
from aspose.words_foss.dom.package import DocxPackage


class DocxDocument:
    """Editable Transitional DOCX, retaining unsupported parts and XML.

    This is a bounded DOM, not an Aspose.Words-complete implementation or a
    sanitizer. Macros, external relationships and embedded resources are retained
    but never executed or fetched. Strict OOXML and signed packages are rejected.
    """

    def __init__(self, source):
        self._package = DocxPackage(source)
        self._nodes = {}
        self._note_graph = None
        from aspose.words_foss.dom.comment_utc import metadata_links, typed_parts
        if metadata_links(self._package):
            typed_parts(self._package)
        root = self._package.tree("word/document.xml").documentElement
        if not _is(root, "document") or _find(root, "body") is None:
            raise ValueError("Expected a Transitional Word document with a body")
        from aspose.words_foss.dom.row_format import validate_row_properties
        for name in self._package.story_part_names():
            validate_row_properties(self._package.tree(name))

    def _annotations(self):
        if self._note_graph is None:
            from aspose.words_foss.dom.notes import NoteGraph
            self._note_graph = NoteGraph(self)
        return self._note_graph

    def _wrap(self, part_name, element, *, storage=False):
        storage = storage or element.namespaceURI == W and element.localName in {"footnotes", "endnotes", "comments"}
        key = (element, True) if storage else element
        if key not in self._nodes:
            cls = _NODE_CLASSES.get(element.localName, UnknownNode) if element.namespaceURI == W else UnknownNode
            self._nodes[key] = cls(self, part_name, element)
            self._nodes[key]._storage_view = storage
        return self._nodes[key]

    @property
    def body(self) -> Body:
        root = self._package.tree("word/document.xml").documentElement
        return self._wrap("word/document.xml", _find(root, "body"))

    @property
    def styles(self):
        from aspose.words_foss.dom.styles import StyleCollection

        return StyleCollection(self)

    @property
    def part_names(self):
        return self._package.part_names

    @property
    def default_tab_stop(self) -> float:
        """Default tab interval in points from the related settings part."""
        from aspose.words_foss.dom.settings import get_default_tab_stop
        return get_default_tab_stop(self._package)

    @default_tab_stop.setter
    def default_tab_stop(self, value: float) -> None:
        from aspose.words_foss.dom.settings import set_default_tab_stop
        set_default_tab_stop(self._package, value)

    def part_xml(self, name):
        """Read-only XML snapshot; parsing alone does not mark a part modified."""
        return self._package.tree(name).documentElement.toxml()

    def preservation_report(self):
        """List added/modified/removed/unchanged parts since loading (not sanitization)."""
        overrides, _ = self._prepare_note_projection()
        return self._package.preservation_report(overrides)

    def story(self, part_name):
        """Access a body, header/footer, or related note/comment part."""
        if part_name == "word/document.xml":
            return self.body
        root = self._package.tree(part_name).documentElement
        if not (_is(root, "hdr") or _is(root, "ftr") or
                root.namespaceURI == W and root.localName in {"footnotes", "endnotes", "comments"}
                and part_name in self._package.story_part_names()):
            raise NotImplementedError("Only typed Word stories are editable")
        return self._wrap(part_name, root)

    def get_child_nodes(self, node_type=0, deep=False):
        """Traverse the main story; other stories are accessed explicitly."""
        return self.body.get_child_nodes(node_type, deep)

    def create_footnote(self, footnote_type):
        from aspose.words_foss.dom.notes import FootnoteType
        if not isinstance(footnote_type, FootnoteType):
            raise TypeError("footnote_type requires FootnoteType")
        return self._annotations().create("footnotes" if footnote_type == FootnoteType.FOOTNOTE else "endnotes")

    def create_comment(self, author="", initial="", date_time=datetime.min):
        from aspose.words_foss.dom.comment_metadata import normalize_date, normalize_text
        author, initial = normalize_text(author), normalize_text(initial)
        date_time = normalize_date(date_time, constructor=True)
        return self._annotations().create("comments", author=author, initial=initial, date_time=date_time)

    def _create_comment_range(self, local_name, identifier):
        if type(identifier) is not int:
            raise TypeError("Comment range id requires int")
        if not -(2**31) <= identifier < 2**31:
            raise OverflowError("Comment range id exceeds INT32")
        from aspose.words_foss.dom.importing import _set_word_attribute
        element = _new(self.body._element, local_name)
        _set_word_attribute(element, "id", str(identifier))
        element._logical_comment_id = identifier
        return self._wrap("word/document.xml", element)

    def create_comment_range_start(self, identifier):
        return self._create_comment_range("commentRangeStart", identifier)

    def create_comment_range_end(self, identifier):
        return self._create_comment_range("commentRangeEnd", identifier)

    def _creation_context(self, part_name):
        if self._note_graph is not None:
            for name, tree in self._note_graph.pending_trees.values():
                if name == part_name:
                    return tree.documentElement
        return self.story(part_name)._element

    def create_run(self, text="", *, part_name="word/document.xml") -> Run:
        _validate_text(text)
        element = _new(self._creation_context(part_name), "r")
        run = self._wrap(part_name, element)
        run.text = text
        return run

    def import_node(self, src_node: Node, is_import_children: bool,
                    import_format_mode: ImportFormatMode = ImportFormatMode.USE_DESTINATION_STYLES) -> Node:
        """Return a detached copy owned by this document; unsupported references fail before commit."""
        from aspose.words_foss.dom.importing import import_node

        return import_node(self, src_node, is_import_children, import_format_mode)

    def create_paragraph(self, text="", *, part_name="word/document.xml") -> Paragraph:
        _validate_text(text)
        element = _new(self._creation_context(part_name), "p")
        paragraph = self._wrap(part_name, element)
        if text:
            paragraph.append_child(self.create_run(text, part_name=part_name))
        return paragraph

    def create_table(self, rows, columns, *, part_name="word/document.xml") -> Table:
        """Create a plain rectangular table; existing merged grids are not edited."""
        if (type(rows) is not int or type(columns) is not int or rows < 1 or
                not 1 <= columns <= MAX_TABLE_COLUMNS):
            raise ValueError("Table requires positive integer dimensions within the column limit")
        element = _new(self._creation_context(part_name), "tbl")
        grid = _new(element, "tblGrid")
        element.appendChild(grid)
        for _ in range(columns):
            grid.appendChild(_new(element, "gridCol"))
        for _ in range(rows):
            row = _new(element, "tr")
            element.appendChild(row)
            for _ in range(columns):
                cell = _new(element, "tc")
                row.appendChild(cell)
                cell.appendChild(_new(element, "p"))
        return self._wrap(part_name, element)

    def _styles_root(self):
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in self.part_names:
            return None
        rels = self._package.tree(rels_name).documentElement
        if (rels.namespaceURI != "http://schemas.openxmlformats.org/package/2006/relationships" or
                rels.localName != "Relationships"):
            raise ValueError("Expected an OPC relationships part")
        links = [node for node in _elements(rels)
                 if node.namespaceURI == "http://schemas.openxmlformats.org/package/2006/relationships"
                 and node.localName == "Relationship" and node.getAttribute("Type") ==
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"]
        if not links:
            return None
        if len(links) != 1 or links[0].getAttribute("TargetMode") not in {"", "Internal"}:
            raise ValueError("Expected one internal styles relationship")
        target = links[0].getAttribute("Target")
        name = resolve_target("word/document.xml", target)
        if name not in self.part_names:
            raise ValueError("Styles relationship points to a missing or unsafe part")
        root = self._package.tree(name).documentElement
        if not _is(root, "styles"):
            raise ValueError("Expected a Transitional Word styles part")
        return root

    def _has_style(self, style_id, style_type="paragraph"):
        from aspose.words_foss.dom.styles import StyleResolver

        return StyleResolver(self).has_style(style_id, style_type)

    def _prepare_note_projection(self):
        from aspose.words_foss.dom.comment_utc import metadata_links
        if self._note_graph is None and metadata_links(self._package):
            self._annotations()
        return self._note_graph.prepare_projection() if self._note_graph is not None else ({}, lambda: None)

    def _prepared_output(self, *, flat=False):
        from aspose.words_foss.dom.row_format import prepare_row_commit
        overrides, commit_notes = self._prepare_note_projection()
        data = self._package.to_bytes(overrides) if overrides else self._package.to_bytes()
        if flat:
            from aspose.words_foss._flat_opc import encode
            data = encode(data)
        commit_rows = prepare_row_commit(self)
        theme_parts = self._package._default_theme_parts(overrides)
        self._package.validate_parts(theme_parts)
        return data, commit_notes, commit_rows, theme_parts

    def _save(self, destination, *, flat=False):
        from aspose.words_foss._io import atomic_output
        data, commit_notes, commit_rows, theme_parts = self._prepared_output(flat=flat)
        with atomic_output(destination) as temporary:
            temporary.write_bytes(data)
        commit_notes()
        commit_rows()
        self._package._commit_parts(theme_parts)
        if theme_parts:
            self._package._injected_default_theme = True

    def save(self, destination):
        """Atomically save supported projections, updating held objects only on success."""
        self._save(destination)

    def to_bytes(self):
        return self._prepared_output()[0]

    def to_flat_opc(self):
        """Return Flat OPC bytes without committing any warm-object state."""
        return self._prepared_output(flat=True)[0]

    def save_flat_opc(self, destination):
        self._save(destination, flat=True)

    def to_light_document(self):
        """Produce an independent, potentially lossy conversion snapshot of current edits.

        Changes to the snapshot are not written back to this DOM.
        """
        from aspose.words_foss.docx_reader import DocumentReader

        reader = DocumentReader()
        reader.load_bytes(self.to_bytes())
        return reader.to_light_document()
