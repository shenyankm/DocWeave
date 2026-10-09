"""Apply outline opening state before fpdf2 serializes its PDF objects."""

from fpdf.output import OutputProducer


class OutlineOutputProducer(OutputProducer):
    def __init__(self, pdf, expanded_levels, zoom_destination=None):
        super().__init__(pdf)
        self.expanded_levels = expanded_levels
        self.zoom_destination = zoom_destination

    # shortcut: fpdf2 has no public OpenAction hook; verify this override on dependency upgrades.
    def _finalize_catalog(self, catalog_obj, pages_root_obj, first_page_obj, *args, **kwargs):
        super()._finalize_catalog(catalog_obj, pages_root_obj, first_page_obj, *args, **kwargs)
        if self.zoom_destination is not None:
            catalog_obj.open_action = f"[{first_page_obj.id} 0 R /XYZ null null {self.zoom_destination}]"

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
