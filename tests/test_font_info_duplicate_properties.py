"""Duplicate metadata declarations have a source-preserving effective public view."""
from io import BytesIO

import pytest

from aspose.words_foss import Document, SaveFormat, fonts
from .test_font_info_values import package
from .test_source_font_table import W, TABLE, archive, owned_parts, relationships_path


CASES = [
    ('altName', 'First', 'Second', 'alt_name', 'Second'),
    ('altName', 'First', '', 'alt_name', 'First'),
    ('altName', 'First', None, 'alt_name', 'First'),
    ('altName', '', 'Second', 'alt_name', 'Second'),
    ('family', 'swiss', 'roman', 'family', fonts.FontFamily.ROMAN),
    ('family', 'swiss', 'auto', 'family', fonts.FontFamily.AUTO),
    ('family', 'auto', 'roman', 'family', fonts.FontFamily.ROMAN),
    ('family', 'swiss', 'unknown', 'family', fonts.FontFamily.AUTO),
    ('family', 'swiss', None, 'family', fonts.FontFamily.AUTO),
    ('pitch', 'fixed', 'variable', 'pitch', fonts.FontPitch.VARIABLE),
    ('pitch', 'fixed', 'default', 'pitch', fonts.FontPitch.DEFAULT),
    ('pitch', 'fixed', None, 'pitch', fonts.FontPitch.DEFAULT),
    ('charset', 'CC', 'EE', 'charset', 238),
    ('charset', 'CC', '00', 'charset', 0),
    ('charset', 'CC', '', 'charset', 204),
    ('charset', 'CC', None, 'charset', 204),
    ('notTrueType', 'true', 'false', 'is_true_type', True),
    ('notTrueType', 'false', 'true', 'is_true_type', False),
    ('notTrueType', 'false', '', 'is_true_type', False),
    ('notTrueType', 'false', None, 'is_true_type', False),
    ('panose1', '020B', '030C', 'panose', bytearray([3, 12] + [0] * 8)),
    ('panose1', '020B', '', 'panose', bytearray([2, 11] + [0] * 8)),
    ('panose1', '020B', None, 'panose', bytearray([2, 11] + [0] * 8)),
    ('panose1', '', '030C', 'panose', bytearray([3, 12] + [0] * 8)),
    ('panose1', '020B', 'GG', 'panose', bytearray(10)),
]


def declarations(tag, first, last):
    def element(value):
        return f'<w:{tag}/>' if value is None else f'<w:{tag} w:val="{value}"/>'
    return element(first) + element(last)


@pytest.mark.parametrize('tag,first,last,attribute,expected', CASES)
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_duplicate_properties_getter_edit_and_cold_save(tmp_path, tag, first, last, attribute, expected, format):
    doc = Document(BytesIO(package(declarations(tag, first, last))))
    source = doc.light_document_model.source_font_table
    collection = doc.font_infos
    font = collection[0]
    assert getattr(font, attribute) == expected
    assert doc.light_document_model.source_font_table is source
    assert Document(BytesIO(doc.to_bytes(format))).font_infos[0].name == 'Owned'
    assert getattr(Document(BytesIO(doc.to_bytes(format))).font_infos[0], attribute) == expected
    replacement = {'alt_name': 'Edited', 'family': fonts.FontFamily.MODERN,
                   'pitch': fonts.FontPitch.FIXED, 'charset': 204,
                   'is_true_type': False, 'panose': bytearray(range(10))}[attribute]
    setattr(font, attribute, replacement)
    output = tmp_path / 'edited'; doc.save(output, format)
    assert doc.font_infos is collection and collection[0] is font
    assert getattr(font, attribute) == replacement
    assert getattr(Document(output).font_infos[0], attribute) == replacement


@pytest.mark.parametrize('tag,bad,valid', [('charset', 'GG', 'CC'), ('notTrueType', 'UNKNOWN', 'false')])
def test_shadowed_bad_scalar_still_fails_at_load(tag, bad, valid):
    with pytest.raises(RuntimeError):
        Document(BytesIO(package(declarations(tag, bad, valid))))


def test_within_registration_order_is_resolved_before_cross_registration_merging():
    parts = owned_parts(None)
    parts[TABLE] = (f'<w:fonts xmlns:w="{W}"><w:font w:name="Owned">'
                   '<w:family w:val="swiss"/><w:family w:val="auto"/>'
                   '<w:altName w:val="First"/><w:altName w:val="Second"/>'
                   '</w:font><w:font w:name="OWNED"><w:family w:val="roman"/>'
                   '<w:altName w:val="Third"/></w:font></w:fonts>').encode()
    del parts[relationships_path(TABLE)]
    doc = Document(BytesIO(archive(parts)))
    assert doc.font_infos.count == 1
    assert doc.font_infos[0].family == fonts.FontFamily.ROMAN
    assert doc.font_infos[0].alt_name == 'Second'
