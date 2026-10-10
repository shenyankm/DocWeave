"""Font value normalization preserves valid Unicode and rejects corrupt scalar declarations."""
from io import BytesIO

import pytest

from aspose.words_foss import Document, SaveFormat, fonts
from aspose.words_foss.docx_writer import LdmDocxWriter
from .test_source_font_table import W, TABLE, archive, owned_parts, relationships_path


def package(fields):
    parts = owned_parts(None)
    parts[TABLE] = (f'<w:fonts xmlns:w="{W}"><w:font w:name="Owned">' + fields + '</w:font></w:fonts>').encode()
    del parts[relationships_path(TABLE)]
    return archive(parts)


@pytest.mark.parametrize('raw,expected', [('00000000', 0), (' CC ', 204), ('0xCC', 204),
                                        ('0000000CC', 204), ('FFFFFFFF', -1), ('', 0)])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_charset_values_load_and_cold_save(raw, expected, format):
    doc = Document(BytesIO(package(f'<w:charset w:val="{raw}"/>')))
    assert doc.font_infos[0].charset == expected
    assert Document(BytesIO(doc.to_bytes(format))).font_infos[0].charset == expected


@pytest.mark.parametrize('raw', ['GG', '+CC', 'C_C', '100000000'])
def test_invalid_charset_fails_at_document_load(raw):
    with pytest.raises(RuntimeError):
        Document(BytesIO(package(f'<w:charset w:val="{raw}"/>')))


@pytest.mark.parametrize('raw', ['OFF', 'unknown'])
def test_invalid_boolean_fails_at_document_load(raw):
    with pytest.raises(RuntimeError):
        Document(BytesIO(package(f'<w:notTrueType w:val="{raw}"/>')))


@pytest.mark.parametrize('raw,expected', [('', False), ('0', True), ('1', False),
                                       ('true', False), ('false', True), ('on', False), ('off', True)])
def test_onoff_values(raw, expected):
    doc = Document(BytesIO(package(f'<w:notTrueType w:val="{raw}"/>')))
    assert doc.font_infos[0].is_true_type is expected
    assert Document(BytesIO(doc.to_bytes(SaveFormat.DOCX))).font_infos[0].is_true_type is expected


@pytest.mark.parametrize('tag,raw', [('family', 'fantasy'), ('family', 'SWISS'),
                                   ('pitch', 'wide'), ('pitch', 'FIXED')])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_unknown_classification_has_default_view_without_changing_source(tag, raw, format):
    doc = Document(BytesIO(package(f'<w:{tag} w:val="{raw}"/>')))
    before = doc.light_document_model.source_font_table.data
    assert doc.font_infos[0].family == fonts.FontFamily.AUTO
    assert doc.font_infos[0].pitch == fonts.FontPitch.DEFAULT
    assert doc.light_document_model.source_font_table.data == before
    cold = Document(BytesIO(doc.to_bytes(format))).font_infos[0]
    assert cold.family == fonts.FontFamily.AUTO and cold.pitch == fonts.FontPitch.DEFAULT


@pytest.mark.parametrize('raw,prefix', [('020B', [2, 11]), ('GG0B0604020202020204', [11, 6, 4, 2, 2, 2, 2, 2, 4]),
    ('020B06040202020202GG', [2, 11, 6, 4, 2, 2, 2, 2, 2]),
    ('020B060402020202020400', [2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
    ('02 0B 06 04 02 02 02 02 02 04', [2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
    ('2', []), ('020', [2]), ('0 G2', [2]), ('020 B06', [2, 11, 6]),
    ('0x020B', [0, 32]), ('A-B', [171]), ('020B🙂0604', [2, 11, 6, 4]), ('GG', [])])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_panose_extracts_first_ten_hex_pairs_and_pads(raw, prefix, format):
    doc = Document(BytesIO(package(f'<w:panose1 w:val="{raw}"/>')))
    expected = bytearray(prefix + [0] * (10 - len(prefix)))
    assert doc.font_infos[0].panose == expected
    assert Document(BytesIO(doc.to_bytes(format))).font_infos[0].panose == expected


def test_empty_panose_and_charset_tags():
    doc = Document(BytesIO(package('<w:panose1/><w:charset/><w:notTrueType/>')))
    font = doc.font_infos[0]
    assert font.panose is None and font.charset == 0 and font.is_true_type is False


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_alias_legitimate_whitespace_and_unicode_survive_getter_edit_save(tmp_path, format):
    doc = Document(BytesIO(package('<w:altName w:val=" &#9;&#10;&#13; 字体🙂 "/>')))
    font = doc.font_infos[0]
    value = ' \t\n\r 字体🙂 '
    assert font.alt_name == value
    font.alt_name = value + '<&"'
    output = tmp_path / 'unicode.out'; doc.save(output, format)
    assert font.alt_name == value + '<&"'
    assert Document(output).font_infos[0].alt_name == value + '<&"'


@pytest.mark.parametrize('field', ['<w:charset w:val="GG"/>', '<w:notTrueType w:val="UNKNOWN"/>'])
def test_untrusted_model_metadata_fails_before_atomic_replacement(tmp_path, field):
    doc = Document(BytesIO(package('')))
    source = doc.light_document_model.source_font_table
    data = source.data.replace(b'</w:font>', field.encode() + b'</w:font>')
    doc.light_document_model.source_font_table = source.model_copy(update={'data': data,
        'parts': tuple(part.model_copy(update={'data': data}) if part.name == source.part_name else part for part in source.parts)})
    output = tmp_path / 'existing.docx'; output.write_bytes(b'existing')
    before = doc.light_document_model.model_dump_json()
    with pytest.raises(RuntimeError): LdmDocxWriter().write(doc.light_document_model, output)
    assert output.read_bytes() == b'existing'
    assert doc.light_document_model.model_dump_json() == before
