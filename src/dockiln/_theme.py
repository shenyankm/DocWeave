"""Default Office-theme save projection shared by source and package DOM writers."""
from functools import lru_cache
from importlib import resources

THEME_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml'
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'


@lru_cache(maxsize=1)
def default_theme_bytes() -> bytes:
    # MIT python-docx 1.2.0 template; calibrated Latin defaults, 29-script group,
    # and independently specified third-fill gradient stops/colour transforms.
    return (resources.files('dockiln.writers.docx') / 'resources' / 'default_theme.xml').read_bytes()


def missing_theme_parts(payloads: dict[str, bytes]) -> dict[str, bytes]:
    """Prepare new OPC parts without changing an existing or malformed theme relationship."""
    from defusedxml.minidom import parseString
    from dockiln.dom.resources import add_relationship
    from dockiln.utils.xml_helpers import serialize_xml

    rels_name = 'word/_rels/document.xml.rels'
    rels = parseString(payloads[rels_name] if rels_name in payloads else
                       f'<Relationships xmlns="{R[1:-1]}"/>')
    root = rels.documentElement
    if root.namespaceURI != R[1:-1] or root.localName != 'Relationships':
        raise ValueError('Expected an OPC relationships part')
    if any(rel.getAttribute('Type') == THEME_TYPE for rel in
           root.getElementsByTagNameNS(R[1:-1], 'Relationship')):
        return {}
    # An orphan theme remains source material; never replace it with generated defaults.
    if any(name.startswith('word/theme/') and name.endswith('.xml') for name in payloads):
        return {}
    add_relationship(root, THEME_TYPE, 'theme/theme1.xml')
    types = parseString(payloads['[Content_Types].xml'])
    type_root = types.documentElement
    if type_root.namespaceURI != CT[1:-1] or type_root.localName != 'Types':
        raise ValueError('Expected an OPC content types part')
    tag = f'{type_root.prefix}:Override' if type_root.prefix else 'Override'
    override = types.createElementNS(CT[1:-1], tag)
    override.setAttribute('PartName', '/word/theme/theme1.xml')
    override.setAttribute('ContentType', THEME_CONTENT_TYPE)
    type_root.appendChild(override)
    return {rels_name: serialize_xml(rels),
            '[Content_Types].xml': serialize_xml(types),
            'word/theme/theme1.xml': default_theme_bytes()}
