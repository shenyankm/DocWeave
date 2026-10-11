"""Only document-associated settings affect the conversion model."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as E

import pytest
import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln.writers.docx import LdmDocxWriter
from dockiln._flat_opc import encode

R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
SETTINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'


def settings_package(mode, malformed=None):
    raw = LdmDocxWriter().write_to_bytes(ldm.Document(sections=[ldm.Section(body=ldm.Body(
        children=[ldm.Paragraph(children=[ldm.Run(text='Owned settings')])]))]))
    with ZipFile(BytesIO(raw)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    rels = E.fromstring(parts['word/_rels/document.xml.rels'])
    for link in list(rels):
        if link.get('Type') == SETTINGS:
            rels.remove(link)
    name = 'word/custom-settings.xml' if mode == 'custom' else 'word/settings.xml'
    if name != 'word/settings.xml':
        parts.pop('word/settings.xml', None)
    parts[name] = malformed if malformed is not None else (
        f'<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="1234"/>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>').encode()
    if mode != 'orphan':
        attrs = {'Id': 'owned-settings', 'Type': SETTINGS, 'Target': name.removeprefix('word/')}
        if mode in {'external', 'external-url'}:
            attrs['TargetMode'] = 'External'
            if mode == 'external-url':
                attrs['Target'] = 'https://example.invalid/settings.xml'
        E.SubElement(rels, f'{{{R}}}Relationship', attrs)
    parts['word/_rels/document.xml.rels'] = E.tostring(rels)
    types = E.fromstring(parts['[Content_Types].xml'])
    for item in list(types):
        if item.get('PartName') in {'/word/settings.xml', '/word/custom-settings.xml'}:
            types.remove(item)
    E.SubElement(types, f'{{{CT}}}Override', PartName='/' + name,
                 ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
    parts['[Content_Types].xml'] = E.tostring(types)
    stream = BytesIO()
    with ZipFile(stream, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return stream.getvalue()


@pytest.mark.parametrize('mode', ['orphan', 'external', 'external-url', 'internal', 'custom'])
@pytest.mark.parametrize('flat', [False, True])
def test_only_related_settings_change_default_tabs(tmp_path, mode, flat):
    raw = settings_package(mode)
    raw = encode(raw) if flat else raw
    (tmp_path / (mode + ('.xml' if flat else '.docx'))).write_bytes(raw)
    doc = dk.Document(BytesIO(raw))
    assert doc.get_text().strip() == 'Owned settings'
    assert doc.light_document_model.default_tab_stop == (61.7 if mode in {'external', 'internal', 'custom'} else 36)
    assert doc.light_document_model.compatibility_mode == (15 if mode in {'external', 'internal', 'custom'} else 12)


@pytest.mark.parametrize('mode', ['orphan', 'internal'])
@pytest.mark.parametrize('malformed', [b'<broken', b'<!DOCTYPE x [<!ENTITY e "Owned entity">]><x>&e;</x>'])
def test_malformed_settings_remains_rejected(mode, malformed):
    from defusedxml.common import DefusedXmlException
    with pytest.raises((E.ParseError, DefusedXmlException)):
        dk.Document(BytesIO(settings_package(mode, malformed=malformed)))
