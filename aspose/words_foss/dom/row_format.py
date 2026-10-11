"""Pure row-height save projection and success-only live updates."""

from aspose.words_foss.dom.nodes import W, _elements, _is
from aspose.words_foss.dom.styles import _child
from aspose.words_foss.utils.table_row import MAX_SAVED_ROW_HEIGHT, parse_row_height, parse_row_split


def last_row_property(properties, name):
    if properties is None:
        return None
    return next((node for node in reversed(properties.childNodes)
                 if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == W and node.localName == name), None)


def supported_row(row):
    """Recognize Word content ancestry; opaque containers are not formatting targets."""
    root = row.ownerDocument.documentElement
    entries = {"document": "body", "hdr": "hdr", "ftr": "ftr", "footnotes": "footnote",
               "endnotes": "endnote", "comments": "comment"}
    if root.namespaceURI != W or root.localName not in entries:
        return False
    ordinary = {"p", "r", "hyperlink", "tbl", "tr", "tc", entries[root.localName]}
    v = "urn:schemas-microsoft-com:vml"
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    paths = [((v, "textbox"), (v, "shape"), (W, "pict")),
             ((wps, "txbx"), (wps, "wsp"), (a, "graphicData"), (a, "graphic"), (wp, "inline"), (W, "drawing")),
             ((wps, "txbx"), (wps, "wsp"), (a, "graphicData"), (a, "graphic"), (wp, "anchor"), (W, "drawing"))]
    node = row
    while node is not root:
        if node is None or node.nodeType != node.ELEMENT_NODE:
            return False
        if _is(node, "txbxContent"):
            for path in paths:
                ancestor = node.parentNode
                for namespace, local in path:
                    if ancestor is None or ancestor.nodeType != ancestor.ELEMENT_NODE or (ancestor.namespaceURI, ancestor.localName) != (namespace, local):
                        break
                    if namespace == a and local == "graphicData" and ancestor.getAttribute("uri") != wps:
                        break
                    ancestor = ancestor.parentNode
                else:
                    node = ancestor
                    break
            else:
                return False
        elif node.namespaceURI == W and node.localName in ordinary:
            node = node.parentNode
        else:
            return False
    return True


def story_rows(tree):
    return (row for row in tree.getElementsByTagNameNS(W, "tr") if supported_row(row))


def validate_row_properties(tree):
    for row in story_rows(tree):
        for properties in (node for node in _elements(row) if _is(node, "trPr")):
            for element in _elements(properties):
                if _is(element, "trHeight"):
                    parse_row_height(element.getAttributeNS(W, "val"), element.getAttributeNS(W, "hRule"))
                elif _is(element, "cantSplit"):
                    parse_row_split(element.getAttributeNS(W, "val"))


def _oversized_rows(tree):
    for row in story_rows(tree):
        element = last_row_property(_child(row, "trPr"), "trHeight")
        if element is not None:
            height, _ = parse_row_height(element.getAttributeNS(W, "val"), element.getAttributeNS(W, "hRule"))
            if height > MAX_SAVED_ROW_HEIGHT:
                yield row, element


def projected_row_tree(tree):
    if not any(_oversized_rows(tree)):
        return tree
    projected = tree.cloneNode(deep=True)
    for _, element in _oversized_rows(projected):
        prefix = element.prefix or "w"
        element.setAttributeNS(W, f"{prefix}:val", str(round(MAX_SAVED_ROW_HEIGHT * 20)))
    return projected


def prepare_row_commit(document):
    package = document._package
    plan = []
    for name in sorted(package.story_part_names()):
        tree = package.tree(name)
        validate_row_properties(tree)
        plan.extend((name, row, element) for row, element in _oversized_rows(tree))

    def commit():
        for name, row, element in plan:
            prefix = element.prefix or "w"
            element.setAttributeNS(W, f"{prefix}:val", str(round(MAX_SAVED_ROW_HEIGHT * 20)))
            warm = getattr(row, "_row_format_warm", None)
            if warm is not None and warm[0] is element:
                row._row_format_warm = (element, element.getAttributeNS(W, "val"),
                    element.getAttributeNS(W, "hRule"), MAX_SAVED_ROW_HEIGHT, warm[4])
            package._dirty.add(name)

    return commit
