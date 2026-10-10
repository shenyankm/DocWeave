"""Blank construction uses the packaged document, without shared mutable state."""
from io import BytesIO

import pytest

from aspose.words_foss import Document, LoadFormat, LoadOptions, SaveFormat, fonts
from aspose.words_foss import light_document_model as ldm


def metadata(collection):
    return [(f.name, f.alt_name, f.family, f.pitch, f.charset,
             bytes(f.panose), f.is_true_type) for f in collection]


def test_blank_document_defaults_and_independent_instances():
    first, second = Document(), Document()
    assert len(first.sections) == 1
    assert first.first_section is first.last_section
    assert len(first.first_section.body.children) == 1
    paragraph = first.first_section.body.paragraphs[0]
    assert paragraph.text == '' and paragraph.runs == []
    assert paragraph.paragraph_format.style_name == 'Normal'
    assert (first.first_section.page_setup.page_width,
            first.first_section.page_setup.page_height) == (612, 792)
    assert {s.name for s in first.styles} == {
        'Normal', 'Default Paragraph Font', 'Table Normal', 'No List'}
    assert metadata(first.font_infos) == [
        ('Times New Roman', '', fonts.FontFamily.ROMAN, fonts.FontPitch.VARIABLE,
         204, bytes.fromhex('02020603050405020304'), True),
        ('Symbol', '', fonts.FontFamily.ROMAN, fonts.FontPitch.VARIABLE,
         2, bytes.fromhex('05050102010706020507'), True),
        ('Arial', '', fonts.FontFamily.SWISS, fonts.FontPitch.VARIABLE,
         204, bytes.fromhex('020B0604020202020204'), True),
    ]
    collection = first.font_infos
    assert collection.count == 3 and collection.get_by_name('arial') is collection[2]
    assert [collection.embed_true_type_fonts, collection.embed_system_fonts,
            collection.save_subset_fonts] == [False, False, False]
    collection[0].alt_name = 'First only'
    first.first_section.body.children.clear()
    assert second.font_infos[0].alt_name == ''
    assert len(second.first_section.body.paragraphs) == 1


def test_load_options_do_not_reinterpret_internal_blank_template():
    options = LoadOptions()
    options.load_format = LoadFormat.TEXT
    assert Document(load_options=options).font_infos.count == 3


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('edited', [False, True])
def test_new_document_two_saves_and_cold_metadata(tmp_path, format, edited):
    document = Document()
    model, collection = document.light_document_model, document.font_infos
    held = collection[0]
    if edited:
        model.sections[0].body.children[0] = ldm.Paragraph(children=[ldm.Run(text='Owned text')])
        held.alt_name = 'Owned Alias<&'
        held.charset = 0
        collection.embed_system_fonts = True
        collection.save_subset_fonts = True
    expected = metadata(collection)
    before = model.model_dump_json()
    assert metadata(Document(BytesIO(document.to_bytes(format))).font_infos) == expected
    assert model.model_dump_json() == before
    for number in range(2):
        output = tmp_path / f'{number}.out'
        document.save(output, format)
        cold = Document(output)
        assert document.light_document_model is model
        assert document.font_infos is collection and collection[0] is held
        assert metadata(cold.font_infos) == expected
        assert len(cold.sections) == len(cold.first_section.body.paragraphs) == 1
        assert cold.first_section.body.paragraphs[0].text == ('Owned text' if edited else '')
        assert [cold.font_infos.embed_true_type_fonts, cold.font_infos.embed_system_fonts,
                cold.font_infos.save_subset_fonts] == [False, edited, edited]


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('failure', ['prepare', 'publish', 'budget'])
def test_new_document_failed_save_preserves_model_and_existing_output(
        tmp_path, monkeypatch, format, failure):
    from aspose.words_foss import _io
    document = Document()
    collection, model = document.font_infos, document.light_document_model
    held = collection[0]
    held.alt_name = 'Pending'
    output = tmp_path / 'existing.out'
    output.write_bytes(b'original')
    if failure == 'prepare':
        def fail(*args):
            raise RuntimeError('prepare failed')
        monkeypatch.setattr(fonts, 'prepare_font_table_commit', fail)
    elif failure == 'publish':
        def fail(*args):
            raise OSError('publish failed')
        monkeypatch.setattr(_io.os, 'replace', fail)
    else:
        monkeypatch.setattr(_io, 'MAX_INPUT_BYTES', 10)
    before = model.model_dump_json()
    with pytest.raises((RuntimeError, OSError, ValueError)):
        document.save(output, format)
    assert output.read_bytes() == b'original'
    assert model.model_dump_json() == before
    assert document.font_infos is collection
    if failure != 'budget':
        assert collection[0] is held
