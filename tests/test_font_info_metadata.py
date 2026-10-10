"""Public metadata edits preserve source resources and held object identities."""
from contextlib import contextmanager
from io import BytesIO

import pytest

from aspose.words_foss import Document, SaveFormat, fonts
from aspose.words_foss import light_document_model as ldm
from .test_source_font_table import W, TABLE, FONT, archive, owned_parts, relationships_path


def document(flag=None, metadata=None):
    parts = owned_parts(flag)
    if metadata is not None:
        parts[TABLE] = (f'<w:fonts xmlns:w="{W}">' + metadata + '</w:fonts>').encode()
        del parts[relationships_path(TABLE)]
    return Document(BytesIO(archive(parts)))


@pytest.mark.parametrize('flag', [None, False, True])
def test_lookup_defaults_and_flags(flag):
    doc = document(flag)
    collection = doc.font_infos
    font = collection[0]
    assert doc.font_infos is collection
    assert collection.count == len(collection) == 1
    assert list(collection) == [font]
    assert collection.get_by_name('OWNED TINY') is font
    assert collection.contains('owned tiny')
    assert collection.get_by_name('') is None
    assert not collection.contains('missing')
    assert [collection.embed_true_type_fonts, collection.embed_system_fonts, collection.save_subset_fonts] == [bool(flag), False, True]
    assert font.name == 'Owned Tiny'
    assert font.alt_name == 'Owned Alias'
    assert font.family == fonts.FontFamily.SWISS
    assert font.pitch == fonts.FontPitch.VARIABLE
    assert font.charset == 0
    assert font.is_true_type is True
    assert font.panose == bytearray.fromhex('020B0604020202020204')


@pytest.mark.parametrize('call,error', [
    (lambda c: c[-1], IndexError), (lambda c: c[c.count], IndexError),
    (lambda c: c['Owned Tiny'], TypeError), (lambda c: c.get_by_name(None), RuntimeError),
    (lambda c: c.get_by_name(1), TypeError),
    (lambda c: setattr(c[0], 'name', 'new'), AttributeError),
    (lambda c: setattr(c[0], 'alt_name', None), RuntimeError),
    (lambda c: setattr(c[0], 'family', 1), TypeError),
    (lambda c: setattr(c[0], 'pitch', 1), TypeError),
    (lambda c: setattr(c[0], 'charset', True), TypeError),
    (lambda c: setattr(c[0], 'charset', 1 << 31), OverflowError),
    (lambda c: setattr(c[0], 'panose', b''), RuntimeError),
    (lambda c: setattr(c[0], 'is_true_type', 1), TypeError),
    (lambda c: setattr(c, 'embed_true_type_fonts', 1), TypeError),
])
def test_invalid_edit_keeps_model(call, error):
    doc = document(); collection = doc.font_infos
    before = doc.light_document_model.model_dump_json()
    with pytest.raises(error): call(collection)
    assert doc.light_document_model.model_dump_json() == before


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_edits_two_successful_saves_keep_metadata_resources_and_held_objects(tmp_path, format):
    doc = document(True); model = doc.light_document_model
    collection = doc.font_infos; font = collection[0]
    original_font = next(part.data for part in model.source_font_table.parts if part.name == FONT)
    font.alt_name = 'New<&Alias'; font.family = fonts.FontFamily.ROMAN
    font.pitch = fonts.FontPitch.FIXED; font.charset = 204
    font.is_true_type = False; font.panose = bytes(range(10))
    collection.embed_true_type_fonts = False; collection.embed_system_fonts = True
    collection.save_subset_fonts = False
    before = model.model_dump_json()
    doc.to_bytes(format)
    assert model.model_dump_json() == before
    for i in range(2):
        output = tmp_path / f'{i}.out'
        doc.save(output, format)
        assert doc.light_document_model is model
        assert doc.font_infos is collection and collection[0] is font
        assert model.source_font_table.current_registrations_data is not None
        cold = Document(output)
        actual = cold.font_infos[0]
        assert (actual.alt_name, actual.family, actual.pitch, actual.charset) == ('New<&Alias', fonts.FontFamily.ROMAN, fonts.FontPitch.FIXED, 204)
        assert actual.is_true_type is False and actual.panose == bytearray(range(10))
        assert [cold.font_infos.embed_true_type_fonts, cold.font_infos.embed_system_fonts, cold.font_infos.save_subset_fonts] == [False, True, False]
        assert next(p.data for p in cold.light_document_model.source_font_table.parts if p.content_type.endswith('obfuscatedFont')) == original_font
        roundtrip = ldm.Document.model_validate_json(model.model_dump_json())
        assert roundtrip.source_font_table == model.source_font_table
    font.alt_name = 'After save'
    assert collection[0].alt_name == 'After save'
    assert Document(BytesIO(doc.to_bytes(format))).font_infos[0].alt_name == 'After save'


@pytest.mark.parametrize('failure', ['prepare', 'write', 'publish', 'graph', 'budget'])
def test_failed_save_keeps_output_snapshot_projection_and_held_objects(tmp_path, monkeypatch, failure):
    from aspose.words_foss import document as document_module, _io
    doc = document(); collection = doc.font_infos; font = collection[0]
    output = tmp_path / 'existing.docx'; doc.save(output)
    font.alt_name = 'Pending'
    if failure == 'prepare':
        def fail(*args): raise RuntimeError('prepare failed')
        monkeypatch.setattr(fonts, 'prepare_font_table_commit', fail)
    elif failure == 'write':
        @contextmanager
        def fail(path):
            class Broken:
                def write_bytes(self, data): raise OSError('write failed')
            yield Broken()
        monkeypatch.setattr(document_module, 'atomic_output', fail)
    elif failure == 'publish':
        def fail(*args): raise OSError('publish failed')
        monkeypatch.setattr(_io.os, 'replace', fail)
    elif failure == 'graph':
        source = doc.light_document_model.source_font_table
        doc.light_document_model.source_font_table = source.model_copy(update={'parts': source.parts[:-1]})
    else: monkeypatch.setattr(_io, 'MAX_INPUT_BYTES', 10)
    before = doc.light_document_model.model_dump_json(); content = output.read_bytes()
    with pytest.raises((ValueError, RuntimeError, OSError)): doc.save(output)
    assert output.read_bytes() == content
    assert doc.light_document_model.model_dump_json() == before
    assert doc.font_infos is collection
    if failure != 'budget':
        assert collection[0] is font


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('value,cold', [('A\0B', 'A'), ('A\x01B', 'AB'), ('', '')])
def test_alias_warm_input_and_cold_xml_normalization(tmp_path, format, value, cold):
    doc = document(); font = doc.font_infos[0]; font.alt_name = value
    output = tmp_path / 'alias.out'; doc.save(output, format)
    assert font.alt_name == value
    assert Document(output).font_infos[0].alt_name == cold


@pytest.mark.parametrize('value,expected', [(-1, 0), (-2, -2), (256, 256), ((1 << 31)-1, (1 << 31)-1)])
def test_charset_signed_values(value, expected):
    doc = document(); doc.font_infos[0].charset = value
    assert doc.font_infos[0].charset == expected
    assert Document(BytesIO(doc.to_bytes(SaveFormat.DOCX))).font_infos[0].charset == expected


def test_missing_and_duplicate_registration_metadata():
    collection = document(metadata='<w:font w:name="Name"/>').font_infos
    font = collection[0]
    assert (font.alt_name, font.charset, font.family, font.pitch, font.panose) == ('', 0, fonts.FontFamily.AUTO, fonts.FontPitch.DEFAULT, None)
    duplicates = document(metadata='<w:font w:name="Name"><w:altName w:val="First"/><w:family w:val="swiss"/></w:font><w:font w:name="name"><w:altName w:val="Second"/><w:pitch w:val="fixed"/></w:font>')
    collection = duplicates.font_infos
    assert collection.count == 1
    assert collection[0].name == 'Name'
    assert collection[0].alt_name == 'First'
    assert collection[0].family == fonts.FontFamily.SWISS
    assert collection[0].pitch == fonts.FontPitch.FIXED
    collection[0].alt_name = ''
    assert Document(BytesIO(duplicates.to_bytes(SaveFormat.DOCX))).font_infos[0].alt_name == ''
    assert document(metadata='').font_infos.count == 0


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_missing_font_table_getter_does_not_change_rendering_and_success_commits(tmp_path, format):
    from xml.etree import ElementTree as ET
    parts = owned_parts(None)
    del parts[TABLE]; del parts[relationships_path(TABLE)]
    relationships = ET.fromstring(parts['word/_rels/document.xml.rels'])
    relationships.remove(relationships[0])
    parts['word/_rels/document.xml.rels'] = ET.tostring(relationships)
    parts['word/settings.xml'] = f'<w:settings xmlns:w="{W}"/>'.encode()
    doc = Document(BytesIO(archive(parts)))
    before = doc.light_document_model.model_dump_json()
    collection = doc.font_infos
    assert collection.count == 0
    assert [collection.embed_true_type_fonts, collection.embed_system_fonts, collection.save_subset_fonts] == [False, False, False]
    assert doc.light_document_model.model_dump_json() == before
    rendered = doc.to_bytes(format)
    assert doc.light_document_model.model_dump_json() == before
    output = tmp_path / 'generated.out'; doc.save(output, format)
    assert doc.font_infos is collection
    assert [f.name for f in collection] == [f.name for f in Document(BytesIO(rendered)).font_infos]
    assert [f.name for f in collection] == [f.name for f in Document(output).font_infos]


@pytest.mark.parametrize('value', [None, False, True])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_embed_system_fonts_original_presence_and_setter(value, format):
    parts = owned_parts(None)
    element = '' if value is None else '<w:embedSystemFonts' + ('/>' if value else ' w:val="0"/>')
    parts['word/settings.xml'] = (f'<w:settings xmlns:w="{W}">' + element + '</w:settings>').encode()
    doc = Document(BytesIO(archive(parts)))
    collection = doc.font_infos
    assert collection.embed_system_fonts is bool(value)
    source = doc.light_document_model.font_embedding
    assert (source.embed_system_fonts if source else None) is value
    collection.embed_system_fonts = not bool(value)
    assert Document(BytesIO(doc.to_bytes(format))).font_infos.embed_system_fonts is not bool(value)


def test_duplicate_default_metadata_fills_but_existing_panose_stays():
    doc = document(metadata='<w:font w:name="Name"><w:altName w:val=""/><w:family w:val="auto"/><w:pitch w:val="default"/><w:charset w:val="00"/><w:notTrueType w:val="0"/><w:panose1 w:val="00000000000000000000"/></w:font><w:font w:name="name"><w:altName w:val="Second"/><w:family w:val="swiss"/><w:pitch w:val="fixed"/><w:charset w:val="CC"/><w:notTrueType w:val="1"/><w:panose1 w:val="020B0604020202020204"/></w:font>')
    font = doc.font_infos[0]
    assert (font.alt_name, font.family, font.pitch, font.charset, font.is_true_type) == ('Second', fonts.FontFamily.SWISS, fonts.FontPitch.FIXED, 204, False)
    assert font.panose == bytearray(10)
    font.alt_name = 'Edited'
    cold = Document(BytesIO(doc.to_bytes(SaveFormat.DOCX))).font_infos[0]
    assert (cold.alt_name, cold.family, cold.pitch, cold.charset, cold.is_true_type) == ('Edited', fonts.FontFamily.SWISS, fonts.FontPitch.FIXED, 204, False)
    assert cold.panose == bytearray(10)


def test_enumeration_does_not_reparse_immutable_snapshot_per_font(monkeypatch):
    doc = document(metadata=''.join(f'<w:font w:name="Owned {i}"/>' for i in range(20)))
    calls = []
    parse = fonts._root
    def record(data):
        calls.append(data)
        return parse(data)
    monkeypatch.setattr(fonts, '_root', record)
    assert len(list(doc.font_infos)) == 20
    assert len(calls) == 1
