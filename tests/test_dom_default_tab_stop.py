"""Default tabs preserve unrelated OPC payloads and edit atomically."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as E
import pytest
import aspose.words_foss as aw
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss._flat_opc import encode
from tests.test_settings_relationship_loading import settings_package, W


def package_with_values(values):
    with ZipFile(BytesIO(settings_package('internal'))) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = E.fromstring(parts['word/settings.xml'])
    for node in list(root):
        if node.tag == '{' + W + '}defaultTabStop':
            root.remove(node)
    for value in values:
        E.SubElement(root, '{' + W + '}defaultTabStop', {'{' + W + '}val': value})
    E.SubElement(root, '{urn:owned}retain', {'token': 'keep'})
    parts['word/settings.xml'] = E.tostring(root)
    out = BytesIO()
    with ZipFile(out, 'w') as z:
        for n, raw in parts.items():
            z.writestr(n, raw)
    return out.getvalue()


@pytest.mark.parametrize('mode', ['internal', 'custom', 'orphan', 'external', 'external-url'])
@pytest.mark.parametrize('flat', [False, True])
def test_related_default_tabs_preserve_source_and_unrelated_parts(tmp_path, mode, flat):
    raw = settings_package(mode)
    doc = DocxDocument(BytesIO(encode(raw) if flat else raw))
    before = {n: doc._package.payload(n) for n in doc.part_names}
    assert doc.default_tab_stop == (61.7 if mode in {'internal', 'custom', 'external'} else 36)
    assert all(doc._package.payload(n) == b for n, b in before.items())
    doc.default_tab_stop = 42.075
    assert doc.default_tab_stop == 42.1
    for n, b in before.items():
        if n not in {'[Content_Types].xml', 'word/_rels/document.xml.rels',
                      'word/settings.xml' if mode in {'internal', 'external'} else 'word/custom-settings.xml'}:
            assert doc._package.payload(n) == b
    if mode in {'orphan', 'external-url'}:
        assert doc._package.payload('word/settings.xml') == before['word/settings.xml']
    if mode == 'external-url':
        from tests.test_settings_relationship_loading import R, SETTINGS
        links = [n for n in E.fromstring(doc._package.payload('word/_rels/document.xml.rels'))
                 if n.tag == '{' + R + '}Relationship' and n.get('Type') == SETTINGS]
        assert len(links) == 1
        assert links[0].get('Id') == 'owned-settings'
        assert links[0].get('Target') == 'settings2.xml'
        assert links[0].get('TargetMode') is None
    dest = tmp_path / ('tabs.xml' if flat else 'tabs.docx')
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(dest)
        assert doc.default_tab_stop == 42.1
        assert DocxDocument(dest).default_tab_stop == 42.1
        assert aw.Document(dest).default_tab_stop == 42.1


@pytest.mark.parametrize('values,expected', [
    ([], 36), ([''], 0), (['bogus'], 0), (['-20'], -1),
    (['2147483648'], -107374182.4), (['99999999999999'], 13822361.55),
    (['720', '1234'], 61.7), (['1234', '720'], 36),
    (['1in'], 72), (['1cm'], 28.35), (['1mm'], 2.85),
    (['1.25'], .05), (['1e3'], 50), (['720\n'], 0),
])
def test_two_models_share_source_measure_semantics(tmp_path, values, expected):
    raw = package_with_values(values)
    dom = DocxDocument(BytesIO(raw))
    assert dom.default_tab_stop == aw.Document(BytesIO(raw)).default_tab_stop == expected
    dom.default_tab_stop = 61.7
    name = 'word/settings.xml'
    root = E.fromstring(dom._package.payload(name))
    assert root.find('{urn:owned}retain').get('token') == 'keep'
    assert len(root.findall('{' + W + '}defaultTabStop')) == 1
    dest = tmp_path / 'retained.docx'
    dom.save(dest)
    assert DocxDocument(dest).default_tab_stop == 61.7


@pytest.mark.parametrize('value', ['720 ', ' 720 ', '720PT', '1px'])
def test_invalid_source_units_are_explicit_errors(value):
    raw = package_with_values([value])
    with pytest.raises(RuntimeError):
        aw.Document(BytesIO(raw))
    with pytest.raises(RuntimeError):
        _ = DocxDocument(BytesIO(raw)).default_tab_stop


@pytest.mark.parametrize('mode', ['internal', 'orphan'])
def test_failed_set_keeps_package_and_raw_xml(monkeypatch, mode):
    doc = DocxDocument(BytesIO(settings_package(mode)))
    before = {n: doc._package.payload(n) for n in doc.part_names}
    with pytest.raises(RuntimeError):
        doc.default_tab_stop = -.01
    from aspose.words_foss import _io
    monkeypatch.setattr(_io, 'MAX_PART_BYTES', 1)
    with pytest.raises(ValueError, match='safety limit'):
        doc.default_tab_stop = 42
    assert doc.default_tab_stop == (61.7 if mode == 'internal' else 36)
    assert {n: doc._package.payload(n) for n in doc.part_names} == before


@pytest.mark.parametrize('value,expected', [(.025, 0), (.026, .05), (0, 0),
    (1e300, 107374182.35), (float('nan'), -107374182.4),
    (float('inf'), -107374182.4)])
@pytest.mark.parametrize('flat', [False, True])
def test_dom_setter_uses_the_public_numeric_contract(tmp_path, value, expected, flat):
    doc = DocxDocument(BytesIO(settings_package('internal')))
    doc.default_tab_stop = value
    assert doc.default_tab_stop == expected
    dest = tmp_path / ('tabs.xml' if flat else 'tabs.docx')
    (doc.save_flat_opc if flat else doc.save)(dest)
    assert DocxDocument(dest).default_tab_stop == aw.Document(dest).default_tab_stop == expected


def test_new_attribute_binding_does_not_rebind_foreign_prefix():
    raw = settings_package('internal')
    with ZipFile(BytesIO(raw)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts['word/settings.xml'] = (
        f'<settings xmlns="{W}" xmlns:w="urn:owned"><w:retain token="keep"/></settings>').encode()
    out = BytesIO()
    with ZipFile(out, 'w') as z:
        for n, b in parts.items():
            z.writestr(n, b)
    doc = DocxDocument(BytesIO(out.getvalue()))
    doc.default_tab_stop = 61.7
    root = E.fromstring(doc._package.payload('word/settings.xml'))
    assert root.find('{urn:owned}retain').get('token') == 'keep'
    assert root.find('{' + W + '}defaultTabStop').get('{' + W + '}val') == '1234'


@pytest.mark.parametrize('conflict', [False, True])
def test_missing_settings_creation_checks_content_types_before_commit(conflict):
    with ZipFile(BytesIO(settings_package('orphan'))) as z:
        parts = {n: z.read(n) for n in z.namelist() if n != 'word/settings.xml'}
    if conflict:
        root = E.fromstring(parts['[Content_Types].xml'])
        for node in root:
            if node.get('PartName') == '/word/settings.xml':
                node.set('ContentType', 'application/octet-stream')
        parts['[Content_Types].xml'] = E.tostring(root)
    raw = BytesIO()
    with ZipFile(raw, 'w') as z:
        for n, b in parts.items():
            z.writestr(n, b)
    doc = DocxDocument(BytesIO(raw.getvalue()))
    before = {n: doc._package.payload(n) for n in doc.part_names}
    assert doc.default_tab_stop == 36
    if conflict:
        with pytest.raises(ValueError, match='content type'):
            doc.default_tab_stop = 42
        assert before == {n: doc._package.payload(n) for n in doc.part_names}
    else:
        doc.default_tab_stop = 42
        assert doc.default_tab_stop == 42
        root = E.fromstring(doc._package.payload('[Content_Types].xml'))
        assert len([n for n in root if n.get('PartName') == '/word/settings.xml']) == 1
