"""Keep fpdf2's page registry while serializing logical table relationships."""

from contextlib import contextmanager
from dataclasses import dataclass

from fpdf.structure_tree import StructElem, StructureTreeBuilder
from fpdf.syntax import PDFArray, Raw

from aspose.words_foss.light_document_model import iter_grid_cells


@dataclass(slots=True)
class _CellSpan:
    row: int
    rows: int
    columns: int


@dataclass(slots=True)
class _ListLevel:
    list_id: int
    level: int
    element: StructElem
    item: StructElem | None = None


class _MarkedContentReference:
    __slots__ = ('element', 'mcid')

    def __init__(self, element, mcid):
        self.element = element
        self.mcid = mcid

    def serialize(self, _security_handler=None, _obj_id=None):
        return f'<< /Type /MCR /Pg {self.element.pg.ref} /MCID {self.mcid} >>'


class _DocumentChildren(PDFArray):
    def __init__(self):
        super().__init__()
        self.roots = PDFArray()

    # shortcut: fpdf2 iterates this flat registry to assign /Pg; recheck on upgrades.
    def serialize(self, _security_handler=None, _obj_id=None):
        return self.roots.serialize(_security_handler, _obj_id)


class TableStructureBuilder(StructureTreeBuilder):
    def __init__(self, pdf):
        super().__init__()
        self.pdf = pdf
        self.doc_struct_elem.k = _DocumentChildren()
        self.parents = {False: self.doc_struct_elem, True: self.doc_struct_elem}
        self.groups = {}
        self.sources = []
        self.cell_spans = {}
        self.fragments = set()
        self.list_paths = {}

    def _in_band(self):
        return bool(getattr(self.pdf, '_in_header_render', False)
                    or getattr(self.pdf, '_in_footer_render', False))

    def _append_child(self, parent, child):
        children = parent.k.roots if parent is self.doc_struct_elem else parent.k
        children.append(child)

    @contextmanager
    def group(self, struct_type, source, position=None, *, parent=None, attributes=None):
        band = self._in_band()
        original = self.parents[band]
        parent = original if parent is None else parent
        key = (id(parent), id(source), struct_type, position)
        element = self.groups.get(key)
        if element is None:
            self.sources.append(source)
            element = StructElem(struct_type, parent, [])
            if attributes:
                element.a = Raw('<< /O /Table ' + ' '.join(f'/{key} {value}' for key, value in attributes.items()) + ' >>')
            self.groups[key] = element
            self._append_child(parent, element)
        self.parents[band] = element
        try:
            yield
        finally:
            self.parents[band] = original

    @contextmanager
    def list_item(self, list_format, source=None):
        parent = self.parents[self._in_band()]
        path = self.list_paths.setdefault(id(parent), [])
        if path and path[0].list_id != list_format.list_id:
            path.clear()
        level = list_format.list_level_number
        while path and path[-1].level > level:
            path.pop()
        if not path or path[-1].level < level:
            nested_parent = path[-1].item.k[-1] if path else parent
            with self.group('/L', object(), parent=nested_parent):
                path.append(_ListLevel(list_format.list_id, level, self.parents[self._in_band()]))
        with self.group('/LI', source if source is not None else object(), parent=path[-1].element):
            path[-1].item = self.parents[self._in_band()]
            yield

    def end_list(self):
        self.list_paths.pop(id(self.parents[self._in_band()]), None)

    @contextmanager
    def cell(self, table, row_index, column):
        records = self.cell_spans.get(id(table))
        if records is None:
            records, previous = {}, {}
            for index, row in enumerate(table.rows):
                current = {}
                for cell, col, width in iter_grid_cells(row):
                    merge = cell.cell_format.vertical_merge
                    record = previous.get((col, width)) if merge == 2 else None
                    continuing = record is not None
                    if continuing:
                        record.rows += 1
                    else:
                        record = _CellSpan(index, 1, width)
                    records[index, col] = record
                    if merge == 1 or continuing:
                        current[col, width] = record
                previous = current
            self.cell_spans[id(table)] = records
        span = records[row_index, column]
        row = self.parents[self._in_band()]
        parent = self.groups[id(row.p), id(table), '/TR', span.row]
        role = '/TH' if table.rows[span.row].row_format.heading_format else '/TD'
        attributes = {}
        if span.rows > 1:
            attributes['RowSpan'] = span.rows
        if span.columns > 1:
            attributes['ColSpan'] = span.columns
        with self.group(role, table, column, parent=parent, attributes=attributes):
            yield

    def add_marked_content(self, *, continuation=None, **kwargs):
        element, spid = super().add_marked_content(**kwargs)
        if continuation is not None:
            if isinstance(continuation.k[0], int):
                continuation.k = PDFArray(_MarkedContentReference(continuation, mcid) for mcid in continuation.k)
            continuation.k.append(_MarkedContentReference(element, kwargs['mcid']))
            self.struct_tree_root.parent_tree.nums[spid][-1] = continuation
            self.fragments.add(id(element))
            return continuation, spid
        element.p = self.parents[self._in_band()]
        self._append_child(element.p, element)
        return element, spid

    def next_mcid_for_page(self, page_number):
        spid = self.spid_per_page_number.get(page_number)
        return 0 if spid is None else len(self.struct_tree_root.parent_tree.nums[spid])

    def __iter__(self):
        # Fragment records stay in the flat page registry but are not PDF objects.
        yield from (element for element in super().__iter__() if id(element) not in self.fragments)
        yield from self.groups.values()

    def empty(self):
        return not self.doc_struct_elem.k.roots
