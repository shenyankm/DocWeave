"""Edit related settings without rebuilding unrelated package parts."""
from defusedxml.minidom import parseString

from dockiln._opc import (
    ordered_relationships,
    relationships_path,
    resolve_settings_target,
    resolve_target,
)
from dockiln.writers.docx.constants import (
    CT_SETTINGS,
    CT_URI,
    PKG_RELS_URI,
    REL_SETTINGS,
)
from dockiln.dom.nodes import XMLNS, W, _elements, _is, _new
from dockiln.dom.resources import add_relationship
from dockiln.utils.tab_stops import (
    coerce_default_tab_stop,
    parse_default_tab_stop,
)
from dockiln.utils.xml_helpers import serialize_xml


def _relationship_root(package, overrides=None):
    overrides = overrides or {}
    name = relationships_path('word/document.xml')
    root = (parseString(overrides[name]).documentElement if name in overrides else
            package.tree(name).documentElement if name in package.part_names else
            parseString(f'<Relationships xmlns="{PKG_RELS_URI}"/>').documentElement)
    if root.namespaceURI != PKG_RELS_URI or root.localName != 'Relationships':
        raise ValueError('Expected an OPC relationships part')
    return root


def _selected_link(package, overrides=None):
    links = [node for node in _elements(_relationship_root(package, overrides))
             if node.namespaceURI == PKG_RELS_URI and node.localName == 'Relationship']
    ordered = ordered_relationships(links, lambda node, key: node.getAttribute(key))
    link = next((node for node in ordered if node.getAttribute('Type') == REL_SETTINGS), None)
    return link


def settings_target(package):
    link = _selected_link(package)
    if link is None:
        return None
    name = resolve_settings_target(link.getAttribute('Target'))
    return name if name in package.part_names else None


def get_default_tab_stop(package):
    name = settings_target(package)
    if name is None:
        if 'word/settings.xml' in package.part_names:
            package.tree('word/settings.xml')  # Validate orphan XML without applying it.
        return 36.0
    root = package.tree(name).documentElement
    if not _is(root, 'settings'):
        return 36.0
    values = [node.getAttributeNS(W, 'val') for node in _elements(root) if _is(node, 'defaultTabStop')]
    return parse_default_tab_stop(values[-1]) if values else 36.0


def edit_settings(package, edit, *, commit=True):
    """Prepare a selected settings edit while preserving unknown XML and links."""
    name = settings_target(package)
    parts = {}
    if name is not None and not _is(package.tree(name).documentElement, 'settings'):
        name = None
    if name is not None:
        root = package.tree(name).documentElement.cloneNode(deep=True)
    else:
        name, number = 'word/settings.xml', 1
        while name in package.part_names:
            number += 1
            name = f'word/settings{number}.xml'
        root = parseString(f'<w:settings xmlns:w="{W}"/>').documentElement
        links = _relationship_root(package).cloneNode(deep=True)
        ordered = ordered_relationships(
            [node for node in _elements(links) if node.namespaceURI == PKG_RELS_URI and node.localName == 'Relationship'],
            lambda node, key: node.getAttribute(key))
        selected = next((node for node in ordered if node.getAttribute('Type') == REL_SETTINGS), None)
        if selected is not None:
            selected.setAttribute('Target', name.removeprefix('word/'))
            if selected.hasAttribute('TargetMode'):
                selected.removeAttribute('TargetMode')
        else:
            add_relationship(links, REL_SETTINGS, name.removeprefix('word/'))
        parts[relationships_path('word/document.xml')] = serialize_xml(links)
        types = package.tree('[Content_Types].xml').documentElement.cloneNode(deep=True)
        if types.namespaceURI != CT_URI or types.localName != 'Types':
            raise ValueError('Expected an OPC content types part')
        existing = [node for node in _elements(types) if node.namespaceURI == CT_URI
                    and node.localName == 'Override' and node.getAttribute('PartName') == '/' + name]
        if len(existing) > 1 or existing and existing[0].getAttribute('ContentType') != CT_SETTINGS:
            raise ValueError('Conflicting settings content type')
        if not existing:
            item = types.ownerDocument.createElementNS(CT_URI, f'{types.prefix}:Override' if types.prefix else 'Override')
            item.setAttribute('PartName', '/' + name)
            item.setAttribute('ContentType', CT_SETTINGS)
            types.appendChild(item)
        parts['[Content_Types].xml'] = serialize_xml(types)
    edit(root)
    parts[name] = serialize_xml(root)
    package.validate_parts(parts)
    if commit:
        package.set_parts(parts)
    return parts


def set_default_tab_stop(package, value):
    value = coerce_default_tab_stop(value)

    def edit(root):
        for item in list(_elements(root)):
            if _is(item, 'defaultTabStop'):
                root.removeChild(item)
        item = _new(root, 'defaultTabStop')
        prefix = root.prefix or 'w'
        if not root.prefix:
            item.setAttributeNS(XMLNS, 'xmlns:w', W)
        item.setAttributeNS(W, prefix + ':val', str(round(value * 20)))
        root.appendChild(item)

    edit_settings(package, edit)


def prepare_settings_projection(package, overrides):
    """Prepare the selected settings role; publish source changes only after IO."""
    names = (set(package.part_names) | set(overrides)) - {
        name for name, data in overrides.items() if data is None}
    def payload(name):
        return overrides[name] if name in overrides else package.payload(name)

    links = _relationship_root(package, overrides).cloneNode(deep=True)
    records = [node for node in _elements(links)
               if node.namespaceURI == PKG_RELS_URI and node.localName == 'Relationship']
    # Other prepared roles may rename an ID: never resurrect a source loser.
    selected = _selected_link(package)
    typed = [node for node in records if node.getAttribute('Type') == REL_SETTINGS]
    types = parseString(payload('[Content_Types].xml')).documentElement
    if types.namespaceURI != CT_URI or types.localName != 'Types':
        raise ValueError('Expected an OPC content types part')
    declarations = [node for node in _elements(types) if node.namespaceURI == CT_URI
                    and node.localName == 'Override' and node.getAttribute('ContentType') == CT_SETTINGS]
    # An entirely absent settings role needs no package-preserving projection.
    if not typed and not declarations:
        return overrides, lambda: None

    name = None
    if selected is not None:
        candidate = resolve_settings_target(selected.getAttribute('Target'))
        if candidate in names and _is(parseString(payload(candidate)).documentElement, 'settings'):
            name = candidate

    updates = {}
    if name is None:
        name, number = 'word/settings.xml', 1
        while name in names:
            number += 1
            name = f'word/settings{number}.xml'
        updates[name] = f'<w:settings xmlns:w="{W}"/>'.encode()
        existing = [node for node in _elements(types) if node.namespaceURI == CT_URI
                    and node.localName == 'Override' and node.getAttribute('PartName') == '/' + name]
        if len(existing) > 1 or existing and existing[0].getAttribute('ContentType') != CT_SETTINGS:
            raise ValueError('Conflicting settings content type')
        if not existing:
            item = types.ownerDocument.createElementNS(CT_URI, f'{types.prefix}:Override' if types.prefix else 'Override')
            item.setAttribute('PartName', '/' + name)
            item.setAttribute('ContentType', CT_SETTINGS)
            types.appendChild(item)

    unchanged_link = (len(typed) == 1 and selected is not None and
                      selected.getAttribute('Target') == name.removeprefix('word/') and
                      not selected.getAttribute('TargetMode') and
                      sum(node.getAttribute('Id') == selected.getAttribute('Id') for node in records) == 1)
    for node in typed:
        if not unchanged_link:
            links.removeChild(node)
    chosen = selected.cloneNode(deep=True) if selected is not None else None
    ids = {node.getAttribute('Id') for node in _elements(links)
           if node.namespaceURI == PKG_RELS_URI and node.localName == 'Relationship'}
    if unchanged_link:
        pass
    elif chosen is None:
        add_relationship(links, REL_SETTINGS, name.removeprefix('word/'))
    else:
        if chosen.getAttribute('Id') in ids:
            number = 1
            while f'rId{number}' in ids:
                number += 1
            chosen.setAttribute('Id', f'rId{number}')
        chosen.setAttribute('Target', name.removeprefix('word/'))
        if chosen.hasAttribute('TargetMode'):
            chosen.removeAttribute('TargetMode')
        links.appendChild(chosen)

    # Only declared, typed Word settings are eligible for removal. Foreign XML,
    # binary payloads and resources still referenced by retained edges stay intact.
    candidates = {node.getAttribute('PartName').lstrip('/') for node in declarations} - {name}
    retained = set()
    import posixpath

    from dockiln.writers.docx.constants import CT_RELS
    content_types = {node.getAttribute('PartName').lstrip('/'): node.getAttribute('ContentType')
                     for node in _elements(types) if node.namespaceURI == CT_URI and node.localName == 'Override'}
    defaults = {node.getAttribute('Extension'): node.getAttribute('ContentType')
                for node in _elements(types) if node.namespaceURI == CT_URI and node.localName == 'Default'}
    for part in names:
        if content_types.get(part, defaults.get(part.rsplit('.', 1)[-1])) != CT_RELS:
            continue
        root = links if part == relationships_path('word/document.xml') else parseString(payload(part)).documentElement
        for link in _elements(root):
            if link.namespaceURI != PKG_RELS_URI or link.localName != 'Relationship' or not link.getAttribute('Target'):
                continue
            try:
                directory, filename = posixpath.split(part)
                source = posixpath.join(posixpath.dirname(directory), filename.removesuffix('.rels'))
                target = resolve_target(source, link.getAttribute('Target'))
            except ValueError:
                continue  # An unresolved URI cannot refer to a package candidate.
            else:
                retained.add(target)
    for candidate in candidates - retained:
        if candidate in names and _is(parseString(payload(candidate)).documentElement, 'settings'):
            related = relationships_path(candidate)
            if related in names:
                kind = content_types.get(related, defaults.get(related.rsplit('.', 1)[-1]))
                if kind != CT_RELS:
                    continue  # Keep the owner of an opaque part at the .rels path.
                root = parseString(payload(related)).documentElement
                if root.namespaceURI != PKG_RELS_URI or root.localName != 'Relationships':
                    continue
                updates[related] = None
            updates[candidate] = None
            removed = {candidate, related} if related in updates else {candidate}
            for node in list(_elements(types)):
                if node.namespaceURI == CT_URI and node.localName == 'Override' and node.getAttribute('PartName').lstrip('/') in removed:
                    types.removeChild(node)
    rel_name = relationships_path('word/document.xml')
    rel_data, ct_data = serialize_xml(links), serialize_xml(types)
    # Avoid marking an ordinary, already selected role modified just by parsing it.
    if links.toxml() != _relationship_root(package, overrides).toxml():
        updates[rel_name] = rel_data
    if types.toxml() != parseString(payload('[Content_Types].xml')).documentElement.toxml():
        updates['[Content_Types].xml'] = ct_data
    if not updates:
        return overrides, lambda: None
    from dockiln.diagnostics import ContentLossWarning, warn
    warn('Non-selected settings relationships and unreferenced settings parts are normalized',
         ContentLossWarning, code='settings.relationships_normalized')
    combined = {**overrides, **updates}
    package.validate_parts(combined, _allow_removed=True)
    return combined, lambda: package._commit_parts(updates)
