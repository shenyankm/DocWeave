"""Edit related settings without rebuilding unrelated package parts."""
from defusedxml.minidom import parseString
from aspose.words_foss._opc import resolve_target, relationships_path
from aspose.words_foss.docx_writer.constants import CT_URI, CT_SETTINGS, REL_SETTINGS, PKG_RELS_URI
from aspose.words_foss.dom.nodes import W, XMLNS, _elements, _is, _new
from aspose.words_foss.dom.resources import relationship_root, add_relationship
from aspose.words_foss.utils.tab_stops import coerce_default_tab_stop, parse_default_tab_stop
from aspose.words_foss.utils.xml_helpers import serialize_xml


def settings_target(package):
    targets = []
    for link in _elements(relationship_root(package, 'word/document.xml')):
        if link.namespaceURI != PKG_RELS_URI or link.localName != 'Relationship' or link.getAttribute('Type') != REL_SETTINGS:
            continue
        try:
            name = resolve_target('word/document.xml', link.getAttribute('Target'))
        except ValueError:
            if link.getAttribute('TargetMode') == 'External':
                continue
            raise
        if link.getAttribute('TargetMode') == 'External' and name not in package.part_names:
            continue
        targets.append(name)
    if len(targets) > 1:
        raise ValueError('Duplicate document relationship: settings')
    return targets[0] if targets and targets[0] in package.part_names else None


def get_default_tab_stop(package):
    name = settings_target(package)
    if name is None:
        if 'word/settings.xml' in package.part_names:
            package.tree('word/settings.xml')  # Validate orphan XML without applying it.
        return 36.0
    root = package.tree(name).documentElement
    values = [node.getAttributeNS(W, 'val') for node in _elements(root) if _is(node, 'defaultTabStop')]
    return parse_default_tab_stop(values[-1]) if values else 36.0


def set_default_tab_stop(package, value):
    value = coerce_default_tab_stop(value)
    name = settings_target(package)
    parts = {}
    if name is not None:
        root = package.tree(name).documentElement.cloneNode(deep=True)
        if not _is(root, 'settings'):
            raise ValueError('Expected a Word settings part')
    else:
        name, number = 'word/settings.xml', 1
        while name in package.part_names:
            number += 1
            name = f'word/settings{number}.xml'
        root = parseString(f'<w:settings xmlns:w="{W}"/>').documentElement
        links = relationship_root(package, 'word/document.xml').cloneNode(deep=True)
        existing_links = [node for node in _elements(links) if node.namespaceURI == PKG_RELS_URI
                          and node.localName == 'Relationship' and node.getAttribute('Type') == REL_SETTINGS]
        if len(existing_links) > 1:
            raise ValueError('Duplicate document relationship: settings')
        if existing_links:
            existing_links[0].setAttribute('Target', name.removeprefix('word/'))
            if existing_links[0].hasAttribute('TargetMode'):
                existing_links[0].removeAttribute('TargetMode')
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
    for item in list(_elements(root)):
        if _is(item, 'defaultTabStop'):
            root.removeChild(item)
    item = _new(root, 'defaultTabStop')
    prefix = root.prefix or 'w'
    if not root.prefix:
        item.setAttributeNS(XMLNS, 'xmlns:w', W)
    item.setAttributeNS(W, prefix + ':val', str(round(value * 20)))
    root.appendChild(item)
    parts[name] = serialize_xml(root)
    package.set_parts(parts)
