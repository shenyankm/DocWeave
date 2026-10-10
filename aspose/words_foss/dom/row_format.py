"""Pure row-height save projection and success-only live updates."""

from aspose.words_foss.dom.nodes import W, _elements, _is
from aspose.words_foss.dom.styles import _child
from aspose.words_foss.utils.table_row import MAX_SAVED_ROW_HEIGHT, parse_row_height, parse_row_split


def last_row_property(properties, name):
    if properties is None:
        return None
    return next((node for node in reversed(properties.childNodes)
                 if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == W and node.localName == name), None)


def validate_row_properties(tree):
    for properties in tree.getElementsByTagNameNS(W, "trPr"):
        for element in _elements(properties):
            if _is(element, "trHeight"):
                parse_row_height(element.getAttributeNS(W, "val"), element.getAttributeNS(W, "hRule"))
            elif _is(element, "cantSplit"):
                parse_row_split(element.getAttributeNS(W, "val"))


def _oversized_rows(tree):
    for row in tree.getElementsByTagNameNS(W, "tr"):
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
