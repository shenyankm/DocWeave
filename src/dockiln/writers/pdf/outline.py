"""Apply outline opening state before fpdf2 serializes its PDF objects."""

from collections.abc import Mapping

from fpdf.output import OutputProducer


def _validate_outline_options(options):
    for name in ("headings_outline_levels", "expanded_outline_levels", "default_bookmarks_outline_level"):
        value = getattr(options, name)
        if (isinstance(value, bool) or not isinstance(value, int) or value < 0
                or (name != "headings_outline_levels" and value > 9)):
            raise ValueError(f"{name} must be a nonnegative integer" +
                             ("" if name == "headings_outline_levels" else " from 0 to 9"))
    levels = options.bookmarks_outline_levels
    if not isinstance(levels, Mapping) or any(
        not isinstance(name, str) or isinstance(level, bool) or not isinstance(level, int) or not 0 <= level <= 9
        for name, level in levels.items()
    ):
        raise ValueError("bookmarks_outline_levels must map string names to integer levels from 0 to 9")
    for name in ("create_missing_outline_levels", "create_outlines_for_headings_in_tables"):
        if not isinstance(getattr(options, name), bool):
            raise ValueError(f"{name} must be a bool")


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
