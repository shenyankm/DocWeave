"""Apply outline opening state before fpdf2 serializes its PDF objects."""

from fpdf.output import OutputProducer


class OutlineOutputProducer(OutputProducer):
    def __init__(self, pdf, expanded_levels):
        super().__init__(pdf)
        self.expanded_levels = expanded_levels

    # shortcut: fpdf2 has no public outline-count hook; verify this override on dependency upgrades.
    def _add_document_outline(self):
        root, items = super()._add_document_outline()
        if root is None:
            return root, items
        depths = {id(root): 0}
        root.count = 0
        for item in items:
            depths[id(item)] = depths[id(item.parent)] + 1
            item.count = 0
        for item in reversed(items):
            if depths[id(item)] > self.expanded_levels:
                item.count = -item.count
            item.parent.count += 1 + max(item.count, 0)
            if not item.count:
                item.count = None
        if not any(item.count and item.count > 0 for item in items):
            root.count = None
        return root, items
