"""Explicit original-package editing entry point; legacy Document remains unchanged."""

import posixpath
from urllib.parse import unquote, urlsplit

from aspose.words_foss._io import MAX_TABLE_COLUMNS
from aspose.words_foss.dom.nodes import (
    _NODE_CLASSES,
    Body,
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
        root = self._package.tree("word/document.xml").documentElement
        if not _is(root, "document") or _find(root, "body") is None:
            raise ValueError("Expected a Transitional Word document with a body")

    def _wrap(self, part_name, element):
        if element not in self._nodes:
            cls = _NODE_CLASSES.get(element.localName, UnknownNode) if element.namespaceURI == W else UnknownNode
            self._nodes[element] = cls(self, part_name, element)
        return self._nodes[element]

    @property
    def body(self) -> Body:
        root = self._package.tree("word/document.xml").documentElement
        return self._wrap("word/document.xml", _find(root, "body"))

    @property
    def part_names(self):
        return self._package.part_names

    def part_xml(self, name):
        """Read-only XML snapshot; parsing alone does not mark a part modified."""
        return self._package.tree(name).documentElement.toxml()

    def story(self, part_name):
        """Access the main body or a header/footer by its original part name."""
        if part_name == "word/document.xml":
            return self.body
        root = self._package.tree(part_name).documentElement
        if not (_is(root, "hdr") or _is(root, "ftr")):
            raise NotImplementedError("Only body/header/footer stories are editable")
        return self._wrap(part_name, root)

    def get_child_nodes(self, node_type=0, deep=False):
        """Traverse the main story; other stories are accessed explicitly."""
        return self.body.get_child_nodes(node_type, deep)

    def create_run(self, text="", *, part_name="word/document.xml") -> Run:
        _validate_text(text)
        element = _new(self.story(part_name)._element, "r")
        run = self._wrap(part_name, element)
        run.text = text
        return run

    def create_paragraph(self, text="", *, part_name="word/document.xml") -> Paragraph:
        _validate_text(text)
        element = _new(self.story(part_name)._element, "p")
        paragraph = self._wrap(part_name, element)
        if text:
            paragraph.append_child(self.create_run(text, part_name=part_name))
        return paragraph

    def create_table(self, rows, columns, *, part_name="word/document.xml") -> Table:
        """Create a plain rectangular table; existing merged grids are not edited."""
        if (type(rows) is not int or type(columns) is not int or rows < 1 or
                not 1 <= columns <= MAX_TABLE_COLUMNS):
            raise ValueError("Table requires positive integer dimensions within the column limit")
        element = _new(self.story(part_name)._element, "tbl")
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
        uri = urlsplit(target)
        if not target or uri.scheme or uri.netloc or uri.query or uri.fragment:
            raise ValueError("Invalid styles relationship target")
        path = unquote(uri.path)
        name = posixpath.normpath(path.lstrip("/") if path.startswith("/") else "word/" + path)
        if name.startswith("../") or "\\" in name or name not in self.part_names:
            raise ValueError("Styles relationship points to a missing or unsafe part")
        root = self._package.tree(name).documentElement
        if not _is(root, "styles"):
            raise ValueError("Expected a Transitional Word styles part")
        return root

    def _has_style(self, style_id, style_type="paragraph"):
        from aspose.words_foss.dom.styles import StyleResolver

        return StyleResolver(self).has_style(style_id, style_type)

    def save(self, destination):
        """Atomically save original payloads except parts modified through this DOM."""
        self._package.save(destination)

    def to_bytes(self):
        return self._package.to_bytes()

    def to_light_document(self):
        """Produce an independent, potentially lossy conversion snapshot of current edits.

        Changes to the snapshot are not written back to this DOM.
        """
        from aspose.words_foss.docx_reader import DocumentReader

        reader = DocumentReader()
        reader.load_bytes(self.to_bytes())
        return reader.to_light_document()
