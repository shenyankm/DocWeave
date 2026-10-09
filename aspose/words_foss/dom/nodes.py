"""Typed views over the authoritative OOXML tree, not a second document model."""

from math import isfinite
from xml.dom import Node as XmlNode

from aspose.words_foss._opc import bind_namespace_context as _bind_namespace_context
from aspose.words_foss.light_document_model import NodeType

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
XMLNS = "http://www.w3.org/2000/xmlns/"


def _elements(node):
    return [child for child in node.childNodes if child.nodeType == XmlNode.ELEMENT_NODE]


def _is(node, name):
    return node.namespaceURI == W and node.localName == name


def _find(node, name):
    return next((child for child in _elements(node) if _is(child, name)), None)


def _new(node, name):
    prefix = node.prefix or "w"
    element = node.ownerDocument.createElementNS(W, f"{prefix}:{name}")
    if node.namespaceURI == W and not node.prefix:
        element = node.ownerDocument.createElementNS(W, name)
    return element


def _validate_text(value):
    if not isinstance(value, str) or any(
        ord(char) < 32 or 0xD800 <= ord(char) <= 0xDFFF or ord(char) in (0xFFFE, 0xFFFF)
        for char in value
    ):
        raise ValueError("Text must be a string of XML characters without tabs or line breaks")


def _text(node):
    return "".join(child.data for child in node.childNodes
                   if child.nodeType in (XmlNode.TEXT_NODE, XmlNode.CDATA_SECTION_NODE))


def _set_text(node, value):
    text_nodes = [child for child in node.childNodes
                  if child.nodeType in (XmlNode.TEXT_NODE, XmlNode.CDATA_SECTION_NODE)]
    replacement = node.ownerDocument.createTextNode(value)
    if text_nodes:
        node.replaceChild(replacement, text_nodes[0])
        for child in text_nodes[1:]:
            node.removeChild(child)
    else:
        node.appendChild(replacement)
    if value[:1].isspace() or value[-1:].isspace():
        node.setAttributeNS(XML, "xml:space", "preserve")


def _text_spans(nodes):
    spans, offset = [], 0
    for node in nodes:
        end = offset + len(_text(node))
        spans.append((offset, end, node))
        offset = end
    return spans


def _replace_span(spans, start, end, value):
    if start == end:
        target = next(((a, node) for a, b, node in spans if a <= start < b), None)
        if target is None:
            a, _, node = next((span for span in reversed(spans) if span[0] < span[1]), spans[-1])
        else:
            a, node = target
        text = _text(node)
        offset = start - a
        _set_text(node, text[:offset] + value + text[offset:])
        return
    covered = [(a, b, node) for a, b, node in spans if a < end and b > start]
    for i, (a, b, node) in enumerate(covered):
        text = _text(node)
        left, right = max(start - a, 0), min(end - a, b - a)
        _set_text(node, text[:left] + (value if i == 0 else "") + text[right:])


def _onoff(element):
    value = element.getAttributeNS(W, "val") if element is not None else "0"
    if value not in {"", "0", "1", "true", "false", "on", "off"}:
        raise ValueError("Invalid OOXML on/off value")
    return value in {"", "1", "true", "on"}


def _validate_toggle(value):
    if value is not None and not isinstance(value, bool):
        raise ValueError("Direct toggle formatting must be True, False or None")


def _size_value(value):
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError("Font size must be a positive number in half-point increments or None")
    half_points = value * 2
    if (not 0 < half_points <= 2**64 - 1 or
            isinstance(half_points, float) and (not isfinite(half_points) or not half_points.is_integer())):
        raise ValueError("Font size must be a positive number in half-point increments or None")
    return str(int(half_points))


def _read_size(element):
    if element is None:
        return None
    value = element.getAttributeNS(W, "val")
    if not value.isascii() or not value.isdecimal() or len(value) > 20:
        raise ValueError("Expected an OOXML half-point integer font size")
    half_points = int(value)
    if not 0 < half_points <= 2**64 - 1:
        raise ValueError("Invalid OOXML font size")
    return half_points / 2


# Moving/deleting complex content needs relationship and range semantics first.
_SAFE_STRUCTURE = {"p", "r", "t", "tbl", "tr", "tc", "tblGrid", "gridCol",
                   "pPr", "rPr", "tblPr", "trPr", "tcPr"}
_PROPERTY_ROOTS = {"pPr", "rPr", "tblPr", "trPr", "tcPr"}
_ALLOWED_CHILDREN = {"body": {"p", "tbl"}, "hdr": {"p", "tbl"}, "ftr": {"p", "tbl"},
                     "p": {"r"}, "tbl": {"tr"}, "tr": {"tc"}, "tc": {"p", "tbl"}}


def _safe_structure(node, properties=False):
    if node.namespaceURI != W:
        return False
    if node.localName == "sectPr":
        return False
    for i in range(node.attributes.length):
        attribute = node.attributes.item(i)
        if (attribute.namespaceURI not in (None, W, XML, XMLNS) or
                attribute.namespaceURI == XML and attribute.localName not in {"space", "lang"}):
            return False
    if not properties and node.localName not in _SAFE_STRUCTURE:
        return False
    properties = properties or node.localName in _PROPERTY_ROOTS
    return all(_safe_structure(child, properties) for child in _elements(node))


class Node:
    node_type = -1

    def __init__(self, document, part_name, element):
        self._document = document
        self._part_name = part_name
        self._element = element

    @property
    def owner_document(self):
        return self._document

    @property
    def part_name(self):
        return self._part_name

    @property
    def parent_node(self):
        parent = self._element.parentNode
        if parent is None or parent.nodeType != XmlNode.ELEMENT_NODE:
            return None
        return self._document._wrap(self._part_name, parent)

    @property
    def xml(self):
        """Read-only XML snapshot, including unsupported content and properties."""
        return self._element.toxml()

    @property
    def child_nodes(self):
        return tuple(self._document._wrap(self._part_name, child)
                     for child in _elements(self._element)
                     if child.namespaceURI != W or (
                         child.localName not in _PROPERTY_ROOTS and
                         not (_is(self._element, "tbl") and _is(child, "tblGrid"))))

    def get_child_nodes(self, node_type=NodeType.ANY, deep=False):
        found = []
        for child in self.child_nodes:
            if node_type == NodeType.ANY or child.node_type == node_type:
                found.append(child)
            if deep:
                found.extend(child.get_child_nodes(node_type, deep=True))
        return found

    def _editable(self):
        node = self._element
        while node is not None and node.nodeType == XmlNode.ELEMENT_NODE:
            if node.namespaceURI != W or node.localName not in {
                "document", "body", "hdr", "ftr", "p", "r", "hyperlink", "tbl", "tr", "tc",
            }:
                raise NotImplementedError("Editing inside unsupported containers is not supported")
            node = node.parentNode

    def _changed(self):
        root = self._element
        while root.parentNode is not None:
            root = root.parentNode
        if root.nodeType == XmlNode.DOCUMENT_NODE:
            self._document._package._dirty.add(self._part_name)

    def _structural_editable(self):
        self._editable()
        root = self._element
        while root.parentNode is not None and root.parentNode.nodeType == XmlNode.ELEMENT_NODE:
            root = root.parentNode
        ranges = {"bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
                  "commentReference", "fldChar", "fldSimple", "ins", "del", "moveFrom", "moveTo",
                  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
                  "permStart", "permEnd", "customXmlInsRangeStart", "customXmlInsRangeEnd",
                  "customXmlDelRangeStart", "customXmlDelRangeEnd",
                  "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
                  "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd", "rPrChange", "pPrChange",
                  "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "sectPrChange",
                  "cellIns", "cellDel", "cellMerge", "numberingChange"}
        if any(node.localName in ranges for node in root.getElementsByTagNameNS(W, "*")):
            raise NotImplementedError("Structural edits in stories with ranges/revisions are unsupported")

    def append_child(self, child):
        return self.insert_before(child, None)

    def insert_before(self, child, reference):
        self._structural_editable()
        if not isinstance(child, Node):
            raise TypeError("child must be a DOM node")
        if child.owner_document is not self.owner_document or child.part_name != self.part_name:
            raise ValueError("Cross-document/part insertion requires resource import and is unsupported")
        child._structural_editable()
        if reference is not None and reference.parent_node is not self:
            raise ValueError("Reference must be a direct child")
        current = self
        while current is not None:
            if current is child:
                raise ValueError("Cannot insert a node into itself or its descendant")
            current = current.parent_node
        if (child._element.namespaceURI != W or child._element.localName not in
                _ALLOWED_CHILDREN.get(self._element.localName, set())):
            raise ValueError("Invalid parent/child node types")
        if not _safe_structure(child._element):
            raise NotImplementedError("Moving complex content requires relationship/range support")
        if child is reference:
            return child
        old_parent = child.parent_node
        if old_parent is not None:
            old_parent._validate_removal(child)
        anchor = reference._element if reference is not None else None
        if _is(self._element, "body"):
            section = _find(self._element, "sectPr")
            if reference is not None and section is not None:
                siblings = _elements(self._element)
                if siblings.index(reference._element) > siblings.index(section):
                    raise ValueError("Body content must precede the final section properties")
            if anchor is None:
                anchor = section
        _bind_namespace_context(child._element, child._element, self._element)
        self._element.insertBefore(child._element, anchor)
        self._changed()
        return child

    def _validate_removal(self, child):
        if _is(self._element, "tc") and isinstance(child, Paragraph):
            remaining = [node for node in self.child_nodes if node is not child]
            if not remaining or not isinstance(remaining[-1], Paragraph):
                raise ValueError("A table cell must end with a paragraph")
        if _is(self._element, "tbl"):
            raise NotImplementedError("Changing existing table rows requires grid/merge support")
        if _is(self._element, "tr"):
            raise NotImplementedError("Changing existing table columns requires grid/merge support")

    def remove(self):
        self._structural_editable()
        parent = self.parent_node
        if parent is None:
            raise ValueError("Node has no parent")
        if not _safe_structure(self._element):
            raise NotImplementedError("Deleting complex content requires relationship/range support")
        parent._validate_removal(self)
        parent._element.removeChild(self._element)
        parent._changed()
        return self

    def clone(self, deep: bool = True):
        """Copy this node detached, retaining its owner and direct formatting."""
        if not isinstance(deep, bool):
            raise TypeError("deep must be a boolean")
        self._structural_editable()
        if not _safe_structure(self._element):
            raise NotImplementedError("Cloning complex content requires resource import support")
        cloned = self._element.cloneNode(deep=True)
        if not deep and not isinstance(self, Run):
            # OOXML property subtrees are formatting, not composite DOM children.
            for child in list(cloned.childNodes):
                if (child.nodeType != XmlNode.ELEMENT_NODE or child.namespaceURI != W or
                        child.localName not in _PROPERTY_ROOTS | {"tblGrid"}):
                    cloned.removeChild(child)
        _bind_namespace_context(self._element, cloned)
        return self._document._wrap(self.part_name, cloned)


class UnknownNode(Node):
    """Unsupported XML is retained and inspectable, but not structurally editable."""


class Body(Node):
    node_type = NodeType.BODY

    @property
    def paragraphs(self):
        return tuple(node for node in self.child_nodes if isinstance(node, Paragraph))

    @property
    def tables(self):
        return tuple(node for node in self.child_nodes if isinstance(node, Table))


class Paragraph(Node):
    node_type = NodeType.PARAGRAPH

    @property
    def runs(self):
        return tuple(node for node in self.child_nodes if isinstance(node, Run))

    @property
    def text(self):
        """Raw w:t text, including hidden/nested text; no revision filtering or field evaluation."""
        return "".join(_text(node) if _is(node, "t") else
                       "\t" if _is(node, "tab") else "\n"
                       for node in self._element.getElementsByTagNameNS(W, "*")
                       if node.localName in {"t", "tab", "br", "cr"})

    def add_hyperlink(self, text, target):
        """Append a link; URLs are never fetched. Complex range stories remain rejected."""
        from aspose.words_foss.dom.resources import add_hyperlink

        return add_hyperlink(self, text, target)

    def add_picture(self, source, *, width=None, height=None, alternative_text=""):
        """Append an inline PNG/JPEG from bytes, a stream or a path; sizes are points."""
        from aspose.words_foss.dom.resources import add_picture

        return add_picture(self, source, width, height, alternative_text)

    def replace_text(self, old, new):
        """Replace non-overlapping literals across direct runs; keep first-run formatting.

        Complex inline containers, fields, markers, tabs and breaks are rejected.
        Existing run/text nodes are retained, including empty nodes after replacement.
        """
        _validate_text(old)
        _validate_text(new)
        if not old:
            raise ValueError("Search text must not be empty")
        text_nodes = self._plain_text_nodes()
        spans = _text_spans(text_nodes)
        text = "".join(_text(node) for node in text_nodes)
        matches, start = [], 0
        while (index := text.find(old, start)) != -1:
            matches.append(index)
            start = index + len(old)
        for start in reversed(matches):
            _replace_span(spans, start, start + len(old), new)
        if matches:
            self._changed()
        return len(matches)

    def _plain_text_nodes(self):
        self._editable()
        if any(not (_is(node, "r") or _is(node, "pPr")) for node in _elements(self._element)):
            raise NotImplementedError("Text ranges across complex inline content are unsupported")
        return [node for run in self.runs for node in run._plain_text_nodes()]

    def range(self, start=0, end=None):
        """Select [start, end) in Python Unicode code points within a plain paragraph."""
        from aspose.words_foss.dom.ranges import TextRange

        return TextRange(self, start, end)

    @property
    def paragraph_format(self):
        return ParagraphFormat(self)

    @property
    def effective_paragraph_format(self):
        from aspose.words_foss.dom.styles import StyleResolver

        return StyleResolver(self.owner_document).paragraph_format(self)


class Run(Node):
    node_type = NodeType.RUN

    @property
    def text(self):
        return "".join(_text(node) if _is(node, "t") else
                       "\t" if _is(node, "tab") else "\n"
                       for node in _elements(self._element)
                       if node.namespaceURI == W and node.localName in {"t", "tab", "br", "cr"})

    def _plain_text_nodes(self):
        if any(not (_is(node, "rPr") or _is(node, "t")) for node in _elements(self._element)):
            raise NotImplementedError("Run contains non-text content")
        nodes = [node for node in _elements(self._element) if _is(node, "t")]
        if any(_elements(node) for node in nodes):
            raise ValueError("Nested elements in Word text nodes are unsupported")
        return nodes

    @text.setter
    def text(self, value):
        _validate_text(value)
        self._editable()
        nodes = self._plain_text_nodes()
        if not nodes:
            nodes = [_new(self._element, "t")]
            self._element.appendChild(nodes[0])
        for i, node in enumerate(nodes):
            _set_text(node, value if i == 0 else "")
        self._changed()

    def split(self, offset):
        """Split a plain, attached run; return the right run (self at 0, None at end)."""
        self._structural_editable()
        nodes = self._plain_text_nodes()
        length = sum(len(_text(node)) for node in nodes)
        if type(offset) is not int or not 0 <= offset <= length:
            raise ValueError("Split offset must be an integer within the run")
        parent = self.parent_node
        if parent is None or not _is(parent._element, "p"):
            raise ValueError("Run splitting requires a direct paragraph parent")
        if offset == 0:
            return self
        if offset == length:
            return None
        if not _safe_structure(self._element):
            raise NotImplementedError("Splitting a run with unsupported XML metadata is unsafe")
        right = self._element.cloneNode(deep=False)
        properties = _find(self._element, "rPr")
        if properties is not None:
            right.appendChild(properties.cloneNode(deep=True))
        cursor = 0
        for child in list(self._element.childNodes):
            if child is properties:
                continue
            if _is(child, "t"):
                value = _text(child)
                if cursor >= offset:
                    right.appendChild(child)
                elif cursor + len(value) > offset:
                    suffix = child.cloneNode(deep=False)
                    _set_text(suffix, value[offset - cursor:])
                    _set_text(child, value[:offset - cursor])
                    right.appendChild(suffix)
                cursor += len(value)
            elif cursor >= offset:
                right.appendChild(child)
        parent._element.insertBefore(right, self._element.nextSibling)
        self._changed()
        return self._document._wrap(self.part_name, right)

    @property
    def font(self):
        return Font(self)

    @property
    def effective_font(self):
        from aspose.words_foss.dom.styles import StyleResolver

        return StyleResolver(self.owner_document).font(self)


class Hyperlink(Node):
    @property
    def runs(self):
        return tuple(node for node in self.child_nodes if isinstance(node, Run))

    @property
    def text(self):
        return "".join(run.text for run in self.runs)

    @property
    def target(self):
        from aspose.words_foss.docx_writer.constants import R_URI, REL_HYPERLINK
        from aspose.words_foss.dom.resources import relationship_root

        rid = self._element.getAttributeNS(R_URI, "id")
        target = ""
        if rid:
            root = relationship_root(self.owner_document._package, self.part_name)
            relation = next((node for node in _elements(root) if node.getAttribute("Id") == rid), None)
            if (relation is None or relation.getAttribute("Type") != REL_HYPERLINK or
                    relation.getAttribute("TargetMode") != "External"):
                raise ValueError("Missing or unsupported hyperlink relationship")
            target = relation.getAttribute("Target")
        anchor = self._element.getAttributeNS(W, "anchor")
        return target + ("#" + anchor if anchor else "")

    @target.setter
    def target(self, value):
        from aspose.words_foss.dom.resources import set_link_target

        set_link_target(self, value)


class Table(Node):
    node_type = NodeType.TABLE

    # ponytail: row/column edits and vertical merges still require a complete grid editor.
    def insert_before(self, child, reference):
        raise NotImplementedError("Changing existing table rows requires grid/merge support")

    @property
    def rows(self):
        return tuple(node for node in self.child_nodes if isinstance(node, Row))

    def merge_cells(self, row_index, start_column, end_column):
        """Merge a horizontal grid range [start, end); reject vertical/legacy merges."""
        from aspose.words_foss._io import MAX_TABLE_COLUMNS

        self._structural_editable()
        if not _safe_structure(self._element):
            raise NotImplementedError("Merging tables with complex content is unsupported")
        rows = self.rows
        if (any(type(value) is not int for value in (row_index, start_column, end_column)) or
                not 0 <= row_index < len(rows) or not 0 <= start_column < end_column):
            raise ValueError("Expected a valid row and nonempty integer column range")
        grid = _find(self._element, "tblGrid")
        columns = len(_elements(grid)) if grid is not None else 0
        if not 1 <= columns <= MAX_TABLE_COLUMNS:
            raise ValueError("Expected a bounded table grid")
        selected = []
        for index, row in enumerate(rows):
            tr_pr = _find(row._element, "trPr")
            if tr_pr is not None and any(_find(tr_pr, tag) is not None for tag in ("gridBefore", "gridAfter")):
                raise NotImplementedError("Rows with omitted grid cells are unsupported")
            column = 0
            for cell in row.cells:
                properties = _find(cell._element, "tcPr")
                if properties is not None and any(_find(properties, tag) is not None for tag in ("vMerge", "hMerge")):
                    raise NotImplementedError("Vertical and legacy horizontal merges are unsupported")
                span_node = _find(properties, "gridSpan") if properties is not None else None
                raw = span_node.getAttributeNS(W, "val") if span_node is not None else "1"
                if not raw.isascii() or not raw.isdecimal() or len(raw) > 4 or not 1 <= int(raw) <= MAX_TABLE_COLUMNS:
                    raise ValueError("Invalid grid span")
                span = int(raw)
                if not cell.child_nodes or not isinstance(cell.child_nodes[-1], Paragraph):
                    raise ValueError("A table cell must end with a paragraph")
                if index == row_index and column < end_column and column + span > start_column:
                    if column < start_column or column + span > end_column:
                        raise ValueError("Merge boundaries cannot split an existing cell")
                    selected.append(cell)
                column += span
            if column != columns:
                raise ValueError("Table rows do not match the grid")
        if end_column > columns or not selected:
            raise ValueError("Merge range exceeds the grid")
        if len(selected) == 1:
            return selected[0]
        target = selected[0]
        properties = _find(target._element, "tcPr")
        if properties is None:
            properties = _new(target._element, "tcPr")
            target._element.insertBefore(properties, target._element.firstChild)
        span = _find(properties, "gridSpan")
        if span is None:
            span = _new(properties, "gridSpan")
            order = ("vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark")
            anchor = next((node for node in _elements(properties) if node.localName in order), None)
            properties.insertBefore(span, anchor)
        span.setAttributeNS(XMLNS, "xmlns:w", W)
        span.setAttributeNS(W, "w:val", str(end_column - start_column))
        width = _find(properties, "tcW")
        if width is not None:
            properties.removeChild(width)  # The grid/span now determines width, not the old single cell.
        for cell in selected[1:]:
            for child in cell.child_nodes:
                _bind_namespace_context(child._element, child._element, target._element)
                target._element.appendChild(child._element)
            cell._element.parentNode.removeChild(cell._element)
        self._changed()
        return target


class Row(Node):
    node_type = NodeType.ROW

    @property
    def cells(self):
        return tuple(node for node in self.child_nodes if isinstance(node, Cell))

    def insert_before(self, child, reference):
        raise NotImplementedError("Changing table columns requires grid/merge support")


class Cell(Body):
    node_type = NodeType.CELL

    def insert_before(self, child, reference):
        if isinstance(child, Table) and reference is None:
            paragraphs = self.paragraphs
            if not paragraphs:
                raise ValueError("A table cell must end with a paragraph")
            reference = paragraphs[-1]
        return super().insert_before(child, reference)


class HeaderFooter(Body):
    node_type = NodeType.HEADER_FOOTER


class _Format:
    property_name = ""
    order = ()

    def __init__(self, node):
        self._node = node

    def _get(self, name):
        properties = _find(self._node._element, self.property_name)
        return _find(properties, name) if properties is not None else None

    def _set(self, name, value):
        node = self._node
        node._editable()
        properties = _find(node._element, self.property_name)
        if properties is None:
            if value is None:
                return
            properties = _new(node._element, self.property_name)
            node._element.insertBefore(properties, node._element.firstChild)
        element = _find(properties, name)
        if value is None:
            if element is not None:
                properties.removeChild(element)
                node._changed()
            return
        if element is None:
            element = _new(node._element, name)
            later = set(self.order[self.order.index(name) + 1:])
            anchor = next((item for item in _elements(properties)
                           if item.namespaceURI == W and item.localName in later), None)
            properties.insertBefore(element, anchor)
        prefix = node._element.prefix or "w"
        if not node._element.prefix:
            # Attributes never inherit a default namespace.
            element.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:w", W)
        element.setAttributeNS(W, f"{prefix}:val", value)
        node._changed()


class Font(_Format):
    """Direct formatting only: None means inherited, not false."""

    property_name = "rPr"
    order = ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
             "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
             "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
             "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
             "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange")

    def _toggle(self, name):
        element = self._get(name)
        return None if element is None else _onoff(element)

    def _set_toggle(self, name, value):
        _validate_toggle(value)
        self._set(name, None if value is None else "1" if value else "0")

    bold = property(lambda self: self._toggle("b"), lambda self, value: self._set_toggle("b", value))
    italic = property(lambda self: self._toggle("i"), lambda self, value: self._set_toggle("i", value))

    @property
    def size(self):
        return _read_size(self._get("sz"))

    @size.setter
    def size(self, value):
        self._set("sz", _size_value(value))

    @property
    def style_id(self):
        element = self._get("rStyle")
        return element.getAttributeNS(W, "val") if element is not None else None

    @style_id.setter
    def style_id(self, value):
        if value is not None:
            _validate_text(value)
            if not value or not self._node.owner_document._has_style(value, "character"):
                raise ValueError("Character style ID must exist in the related styles part")
        self._set("rStyle", value)


class ParagraphFormat(_Format):
    property_name = "pPr"
    order = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
             "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
             "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
             "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
             "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange")

    @property
    def alignment(self):
        element = self._get("jc")
        return element.getAttributeNS(W, "val") if element is not None else None

    @alignment.setter
    def alignment(self, value):
        if value is not None and value not in {"left", "right", "center", "both", "distribute",
                                               "start", "end", "numTab", "highKashida",
                                               "mediumKashida", "lowKashida", "thaiDistribute"}:
            raise ValueError("Unsupported OOXML paragraph alignment")
        self._set("jc", value)

    @property
    def style_id(self):
        element = self._get("pStyle")
        return element.getAttributeNS(W, "val") if element is not None else None

    @style_id.setter
    def style_id(self, value):
        if value is not None:
            _validate_text(value)
            if not value or not self._node.owner_document._has_style(value):
                raise ValueError("Paragraph style ID must exist in the related styles part")
        self._set("pStyle", value)


_NODE_CLASSES = {"body": Body, "p": Paragraph, "r": Run, "tbl": Table, "tr": Row,
                 "tc": Cell, "hdr": HeaderFooter, "ftr": HeaderFooter, "hyperlink": Hyperlink}
