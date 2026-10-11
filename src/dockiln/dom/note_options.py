"""Live note options over selected settings and top-level section properties."""
from dockiln import light_document_model as ldm
from dockiln.dom.nodes import XMLNS, W, _elements, _find, _is, _new
from dockiln.dom.settings import edit_settings, settings_target
from dockiln.notes import (
    _FIELDS,
    _PROPERTY_ORDER,
    FOOTNOTE_COLUMNS_NS,
    EndnoteOptions,
    FootnoteOptions,
    _get,
    _property,
    _set,
)
from dockiln.utils.xml_helpers import serialize_xml


def section_roots(document):
    body = document.body._element
    roots = []
    for child in _elements(body):
        if _is(child, 'p'):
            properties = _find(child, 'pPr')
            if properties is not None:
                root = _find(properties, 'sectPr')
                if root is not None:
                    roots.append(root)
    final = _find(body, 'sectPr')
    roots.append(final)
    return roots


def _read(root, warm):
    result = []
    if root is not None:
        for kind in ('footnote', 'endnote'):
            prop = _find(root, kind + 'Pr')
            if prop is None:
                continue
            children = []
            for child in _elements(prop):
                if child.namespaceURI == W and child.localName in _FIELDS.values():
                    children.append(ldm.NoteProperty(kind=child.localName, value=child.getAttributeNS(W, 'val')))
            result.append(ldm.NoteProperties(kind=kind, children=children))
    if root is not None and _is(root, 'sectPr'):
        from dockiln._notes import note_id
        nodes = [n for n in _elements(root) if n.namespaceURI == FOOTNOTE_COLUMNS_NS and n.localName == 'footnoteColumns']
        if nodes:
            prop = _property(result, 'footnote')
            if prop is None:
                prop = ldm.NoteProperties(kind='footnote')
                result.append(prop)
            prop.columns = max(0, int(note_id(nodes[-1].getAttributeNS(W, 'val'))))
    for old in warm:
        prop = _property(result, old.kind)
        if prop is None:
            prop = ldm.NoteProperties(kind=old.kind)
            result.append(prop)
        prop.columns, prop.number_style_override = old.columns, old.number_style_override
    return result


def _edit(root, kind, field, value, mode=12):
    if field == 'columns':
        for item in list(_elements(root)):
            if item.namespaceURI == FOOTNOTE_COLUMNS_NS and item.localName == 'footnoteColumns':
                root.removeChild(item)
        if value > 0 and mode >= 15:
            item = root.ownerDocument.createElementNS(FOOTNOTE_COLUMNS_NS, 'w15:footnoteColumns')
            item.setAttributeNS(XMLNS, 'xmlns:w15', FOOTNOTE_COLUMNS_NS)
            item.setAttributeNS(XMLNS, 'xmlns:w', W)
            item.setAttributeNS(W, 'w:val', str(value))
            root.appendChild(item)
        return
    model = []
    _set(model, kind, field, value)
    prop = _find(root, kind + 'Pr')
    if prop is None:
        prop = _new(root, kind + 'Pr')
        root.appendChild(prop)
    tag = _FIELDS[field]
    for child in list(_elements(prop)):
        if _is(child, tag):
            prop.removeChild(child)
    for child in model[0].children:
        item = _new(prop, child.kind)
        if not root.prefix:
            item.setAttributeNS(XMLNS, 'xmlns:w', W)
        item.setAttributeNS(W, (root.prefix or 'w') + ':val', child.value)
        following = next((node for node in _elements(prop)
                          if node.namespaceURI == W and
                          _PROPERTY_ORDER.get(node.localName, -1) > _PROPERTY_ORDER[child.kind]), None)
        prop.insertBefore(item, following)


class NoteBackend:
    def __init__(self, document, section=None):
        self.document, self.section = document, section
        if not hasattr(document, '_note_option_warm'):
            document._note_option_warm = {}

    def _warm(self, section):
        return self.document._note_option_warm.get(section, [])

    def _root(self, section):
        if section is not None:
            return section_roots(self.document)[section]
        name = settings_target(self.document._package)
        root = self.document._package.tree(name).documentElement if name else None
        return root if root is not None and _is(root, 'settings') else None

    def get(self, kind, field):
        properties = _read(self._root(self.section), self._warm(self.section))
        defaults = _read(self._root(None), self._warm(None)) if self.section is not None else None
        return _get(properties, defaults, kind, field)

    def set(self, kind, field, value):
        package = self.document._package
        settings = self._root(None)
        mode = 12
        compat = _find(settings, 'compat') if settings is not None else None
        if compat is not None:
            from dockiln._notes import note_id
            for node in _elements(compat):
                if _is(node, 'compatSetting') and node.getAttributeNS(W, 'name') == 'compatibilityMode' and node.getAttributeNS(W, 'uri') == 'http://schemas.microsoft.com/office/word':
                    mode = int(note_id(node.getAttributeNS(W, 'val')))
        sections = section_roots(self.document)
        document_scope = self.section is None or kind == 'endnote' and field == 'position'
        indices = range(len(sections)) if document_scope else [self.section]
        clones = []
        for index in indices:
            existing = sections[index]
            root = existing.cloneNode(deep=True) if existing is not None else _new(self.document.body._element, 'sectPr')
            _edit(root, kind, field, value, mode)
            clones.append((index, existing, root))
        parts = edit_settings(package, lambda root: _edit(root, kind, field, value), commit=False) if document_scope and field != 'columns' else {}
        # Check the complete prospective package before mutating any live body node.
        body_copy = self.document.body._element.ownerDocument.cloneNode(deep=True)
        main = body_copy.documentElement
        # Work with corresponding section roots without constructing a package view.
        body = _find(main, 'body')
        roots = []
        for child in _elements(body):
            if _is(child, 'p'):
                ppr = _find(child, 'pPr')
                if ppr is not None and _find(ppr, 'sectPr') is not None:
                    roots.append(_find(ppr, 'sectPr'))
        roots.append(_find(body, 'sectPr'))
        for index, _, prepared in clones:
            old = roots[index]
            if old is None:
                body.appendChild(body_copy.importNode(prepared, deep=True))
            else:
                old.parentNode.replaceChild(body_copy.importNode(prepared, deep=True), old)
        candidate = dict(parts)
        candidate['word/document.xml'] = serialize_xml(main)
        package.validate_parts(candidate)
        package.set_parts(parts)
        for index, existing, prepared in clones:
            if field != 'columns' or existing is not None or value > 0 and mode >= 15:
                if existing is None:
                    self.document.body._element.appendChild(prepared)
                else:
                    # Retain original section and paragraph wrappers during property edits.
                    _edit(existing, kind, field, value, mode)
                self.document.body._changed()
            warm = self.document._note_option_warm.setdefault(index, [])
            _set(warm, kind, field, value)
        if document_scope:
            _set(self.document._note_option_warm.setdefault(None, []), kind, field, value)


class PageSetup:
    def __init__(self, document, index):
        self._backend = NoteBackend(document, index)

    @property
    def footnote_options(self):
        return FootnoteOptions._for_owner(self._backend, 'footnote', dom=True)

    @property
    def endnote_options(self):
        return EndnoteOptions._for_owner(self._backend, 'endnote', dom=True)


class Section:
    def __init__(self, document, index):
        self.page_setup = PageSetup(document, index)
