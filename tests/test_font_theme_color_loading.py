"""Raw theme getters and layout RGB have distinct, calibrated meanings."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pymupdf
import pytest

from aspose.words_foss import Document
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.ldm_builder.cascading import FontResolver

ROOT = Path(__file__).parents[1] / 'docs/benchmarks'
REPORT = json.loads((ROOT / 'font-theme-colors-26.9.json').read_text())
ROWS = REPORT['native']['records']


def getter(value):
    if value['is_empty']:
        return 'Color [Empty]'
    return f"Color [A={value['a']}, R={value['r']}, G={value['g']}, B={value['b']}]"


@pytest.mark.parametrize('row', ROWS, ids=lambda r: r['input'])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_theme_raw_color_getters_and_all_modifiers_render(row, json_roundtrip):
    with ZipFile(ROOT / REPORT['corpus']) as archive:
        model = Document(BytesIO(archive.read(row['input']))).light_document_model
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    font = model.sections[0].body.paragraphs[0].runs[0].font
    assert font.color == getter(row['getter']) and font.color_explicit is True
    combined = 'themeTint' in row['modifiers'] and 'themeShade' in row['modifiers']
    assert (font.color_rendering is None) is combined
    from aspose.words_foss.pdf_writer import LdmPdfWriter
    raw = LdmPdfWriter().write_to_bytes(model)
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        colors = [s['color'] for p in pdf for b in p.get_text('dict')['blocks']
                  for line in b.get('lines', []) for s in line['spans'] if s['text'] == 'IMPORT']
    assert colors == [row['pdf_target_color']]


@pytest.mark.parametrize('copy', [False, True])
def test_color_assignment_discards_inherited_theme_rendering(copy):
    font = ldm.Font(color='Color [Empty]', source_color=ldm.SourceColor(value='auto', theme_color='accent1'),
                    color_rendering='Color [A=255, R=18, G=52, B=86]')
    if copy:
        font = font.model_copy(update={'color': font.color})
    else:
        font.color = font.color
    assert font.color_rendering is None and font.render_color == font.color
    assert font.color_explicit is True and font.source_color is None


def test_direct_rgb_clears_inherited_theme_but_absent_color_keeps_it():
    font = ldm.Font(color='Color [Empty]', color_rendering='Color [A=255, R=18, G=52, B=86]')
    FontResolver.merge(font, ldm.Font(bold=True))
    assert font.color_rendering is not None
    FontResolver.merge(font, ldm.Font(color='Color [A=255, R=255, G=0, B=0]'))
    assert font.color_rendering is None and font.render_color == font.color


@pytest.mark.parametrize('row', ROWS, ids=lambda r: r['input'])
@pytest.mark.parametrize('format', ['docx', 'flat_opc'])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_saved_theme_declarations_and_layout_survive_json(row, format, json_roundtrip):
    from aspose.words_foss import SaveFormat
    from aspose.words_foss.docx_writer import LdmDocxWriter
    from aspose.words_foss.saving import OoxmlSaveOptions
    with ZipFile(ROOT / REPORT['corpus']) as archive:
        model = Document(BytesIO(archive.read(row['input']))).light_document_model
    source = model.sections[0].body.paragraphs[0].runs[0].font.source_color
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    saved = LdmDocxWriter(OoxmlSaveOptions(SaveFormat.DOCX if format == 'docx' else SaveFormat.FLAT_OPC)).write_to_bytes(model)
    reopened = Document(BytesIO(saved))
    reread = reopened.light_document_model
    assert reread.sections[0].body.paragraphs[0].runs[0].font.source_color == source
    assert reread.sections[0].body.paragraphs[0].runs[0].font.color == getter(row['getter'])
    from xml.etree import ElementTree as ET
    assert ET.tostring(ET.fromstring(reread.source_theme.data)) == ET.tostring(ET.fromstring(model.source_theme.data))
    if format == 'docx':
        assert reread.source_theme.data == model.source_theme.data
    raw = reopened.to_bytes(SaveFormat.PDF)
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        colors = [s['color'] for p in pdf for b in p.get_text('dict')['blocks']
                  for line in b.get('lines', []) for s in line['spans'] if s['text'] == 'IMPORT']
    assert colors == [row['pdf_target_color']]


@pytest.mark.parametrize('raw', [b'<private/>', b'<broken PRIVATE_THEME_VALUE',
                               b'<!DOCTYPE theme [<!ENTITY secret "PRIVATE_THEME_VALUE">]><theme>&secret;</theme>'])
def test_invalid_theme_fails_before_replacing_output(tmp_path, raw):
    from aspose.words_foss.docx_writer import LdmDocxWriter
    model = ldm.Document(source_theme=ldm.SourceTheme(data=raw))
    output = tmp_path / 'existing.docx'
    output.write_bytes(b'existing')
    with pytest.raises(ValueError) as error:
        LdmDocxWriter().write(model, output)
    assert 'PRIVATE_THEME_VALUE' not in str(error.value)
    assert output.read_bytes() == b'existing'


def test_theme_relationship_resources_fail_explicitly_before_output(tmp_path):
    from aspose.words_foss.docx_writer import LdmDocxWriter
    with ZipFile(ROOT / REPORT['corpus']) as archive:
        model = Document(BytesIO(archive.read(ROWS[0]['input']))).light_document_model
    relationships = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="image" Target="../media/image.png"/></Relationships>'
    model.source_theme = ldm.SourceTheme(data=model.source_theme.data, relationships=relationships)
    output = tmp_path / 'existing.docx'
    output.write_bytes(b'existing')
    with pytest.raises(ValueError, match='Theme relationship resources'):
        LdmDocxWriter().write(model, output)
    assert output.read_bytes() == b'existing'


def test_source_color_snapshot_is_immutable_and_assignment_validated():
    from pydantic import ValidationError
    source = ldm.SourceColor(value='auto', theme_color='accent1')
    with pytest.raises(ValidationError):
        source.theme_color = 'accent2'
    font = ldm.Font(source_color=source)
    with pytest.raises(TypeError):
        font.source_color = {'theme_color': 'accent2'}
    assert font.source_color is source
