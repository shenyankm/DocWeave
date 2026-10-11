"""Native-compatible modern comment identifiers from self-contained XML."""
from datetime import timezone
from io import BytesIO

import pytest

from dockiln.dom import DocxDocument
from dockiln.dom.comment_utc import CID, CEX, W14, hex_id
from .test_dom_comment_utc import modern
from .test_dom_comment_metadata import comment


@pytest.mark.parametrize('raw, expected', [
    ('00000000000012345678', 0x12345678), (' \t12345678\r\n', 0x12345678),
    ('000000000', 0), ('FFFFFFFF', 0xFFFFFFFF),
    ('123456789', None), ('\u00a012345678\u00a0', None),
    ('', None), ('  ', None), ('+12345678', None), ('0x12345678', None),
    ('1234_5678', None), ('Owned invalid identifier', None),
])
def test_identifier_lexical_recovery(raw, expected):
    assert hex_id(raw) == expected


@pytest.mark.parametrize('left, right, year', [
    (None, '0', 2030), ('0', None, 2030), (None, 'Owned invalid', 2026),
    ('Owned invalid', None, 2026), ('', '0', 2026), ('0', '', 2026),
    ('Owned invalid', '80000000', 2030), ('80000000', 'Owned invalid', 2030),
    ('123456789', '80000000', 2030), ('80000000', '123456789', 2030),
    ('000000000', '0', 2030), (' 23456789 ', '23456789', 2030),
    ('0000000023456789', '23456789', 2030),
    ('\u00a023456789\u00a0', '23456789', 2026),
])
@pytest.mark.parametrize('flat', [False, True])
def test_durable_identifier_binding_and_two_saves(tmp_path, left, right, year, flat):
    doc = modern(tmp_path)
    for role, ns, value in [('commentsIds', CID, left), ('commentsExtensible', CEX, right)]:
        item = doc._package.tree('word/' + role + '.xml').documentElement.firstChild
        if value is None:item.removeAttributeNS(ns, 'durableId')
        else:item.setAttributeNS(ns, 'm:durableId', value)
    held = comment(doc);expected = held.date_time_utc
    assert expected.year == year
    original = {name: doc.part_xml(name) for name in ['word/commentsIds.xml', 'word/commentsExtensible.xml']}
    cold = comment(DocxDocument(BytesIO(doc.to_flat_opc() if flat else doc.to_bytes())))
    assert cold.date_time_utc == expected.replace(microsecond=0, tzinfo=timezone.utc)
    assert {name: doc.part_xml(name) for name in original} == original
    for index in range(2):
        out = tmp_path / str(index)
        (doc.save_flat_opc if flat else doc.save)(out)
        assert comment(doc) is held and held.date_time_utc == expected
        assert comment(DocxDocument(out)).date_time_utc == expected.replace(microsecond=0, tzinfo=timezone.utc)


@pytest.mark.parametrize('raw, year', [
    ('00000000000012345678', 2030), (' \t12345678\r\n', 2030),
    ('\u00a012345678\u00a0', 2026), ('\u200312345678\u2003', 2026),
])
@pytest.mark.parametrize('flat', [False, True])
def test_paragraph_identifier_recovers_without_unicode_trimming(tmp_path, raw, year, flat):
    doc = modern(tmp_path)
    p = doc._package.tree('word/comments.xml').getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'p')[0]
    p.setAttributeNS(W14, 'w14:paraId', raw);p.setAttribute('xmlns:w14', W14)
    node = doc._package.tree('word/commentsIds.xml').documentElement.firstChild
    node.setAttributeNS(CID, 'm:paraId', raw)
    assert comment(doc).date_time_utc.year == year
    out = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(out)
    assert comment(DocxDocument(out)).date_time_utc.year == year


@pytest.mark.parametrize('flat', [False, True])
def test_recovered_identifiers_survive_failed_publication(tmp_path, monkeypatch, flat):
    doc = modern(tmp_path)
    for role, ns in [('commentsIds', CID), ('commentsExtensible', CEX)]:
        item = doc._package.tree('word/' + role + '.xml').documentElement.firstChild
        item.setAttributeNS(ns, 'm:durableId', 'Owned invalid')
    held = comment(doc);value = held.date_time_utc
    originals = {name: doc._package.payload(name) for name in doc._package.part_names}
    output = tmp_path / 'existing';output.write_bytes(b'owned original output')
    def fail(*args):raise OSError('owned publication failure')
    monkeypatch.setattr('os.replace', fail)
    with pytest.raises(OSError, match='owned publication failure'):
        (doc.save_flat_opc if flat else doc.save)(output)
    assert output.read_bytes() == b'owned original output'
    assert comment(doc) is held and held.date_time_utc == value
    assert {name: doc._package.payload(name) for name in originals} == originals
