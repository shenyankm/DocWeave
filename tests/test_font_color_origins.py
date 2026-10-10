"""Automatic color overrides RGB inheritance and retains its direct declaration."""

import json
import re
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.saving import OoxmlSaveOptions

ROOT = Path(__file__).parents[1]
REPORT = json.loads((ROOT / 'docs/benchmarks/font-color-categories-26.9.json').read_text())


def observed(color):
    if color == 'Color [Empty]':
        return {'a': 0, 'r': 0, 'g': 0, 'b': 0, 'is_empty': True}
    return dict(zip(('a', 'r', 'g', 'b'), map(int, re.findall(r'\d+', color)))) | {'is_empty': False}


@pytest.mark.parametrize('row', REPORT['records'], ids=lambda r: r['input'])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_color_getter_and_direct_origin_survive_json_and_save(row, format, json_roundtrip):
    with ZipFile(ROOT / 'docs/benchmarks' / REPORT['corpus']) as archive:
        raw = archive.read(row['input'])
    model = Document(BytesIO(raw)).light_document_model
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    font = model.sections[0].body.paragraphs[0].runs[0].font
    assert observed(font.color) == row['observed']
    assert font.color_explicit is (row['values'][3] is not None)
    saved = LdmDocxWriter(OoxmlSaveOptions(format)).write_to_bytes(model)
    reopened = Document(BytesIO(saved)).light_document_model.sections[0].body.paragraphs[0].runs[0].font
    assert observed(reopened.color) == row['observed']
    if row['values'][3] is not None:
        assert reopened.color_explicit is True


@pytest.mark.parametrize('copy', [False, True])
@pytest.mark.parametrize('color', ['Color [Empty]', 'Color [A=255, R=255, G=0, B=0]'])
def test_same_value_color_assignment_is_explicit(copy, color):
    font = ldm.Font(color=color, color_explicit=False)
    if copy:
        font = font.model_copy(update={'color': color})
    else:
        font.color = color
    font = ldm.Font.model_validate_json(font.model_dump_json())
    assert font.color_explicit is True
    from aspose.words_foss.docx_writer.runs import render_rPr
    assert 'w:color' in render_rPr(font, base=ldm.Font(color=color))


def test_color_evidence_verifies():
    verify = runpy.run_path(str(ROOT / 'scripts/verify_commercial_baseline.py'))
    assert verify['verify_font_color_categories'](ROOT / 'docs/benchmarks') == 1024


def test_automatic_run_color_overrides_red_style_in_painted_pdf(tmp_path):
    import pymupdf
    with ZipFile(ROOT / 'docs/benchmarks' / REPORT['corpus']) as archive:
        raw = archive.read('n-n-FF0000-auto.docx')
    output = tmp_path / 'automatic.pdf'
    Document(BytesIO(raw)).save(output, SaveFormat.PDF)
    with pymupdf.open(output) as pdf:
        page = next(p for p in pdf if any(w[4] == 'IMPORT' for w in p.get_text('words')))
        word = next(w for w in page.get_text('words') if w[4] == 'IMPORT')
        spans = [s for b in page.get_text('dict')['blocks'] for line in b.get('lines', [])
                 for s in line['spans'] if s['text'] == 'IMPORT']
        assert spans and all(s['color'] == 0 for s in spans)
        pix = page.get_pixmap(matrix=pymupdf.Matrix(3, 3), clip=pymupdf.Rect(word[:4]), alpha=False)
        pixels = list(zip(*(iter(pix.samples),) * 3))
        assert any(max(r, g, b) < 100 for r, g, b in pixels)
        assert not any(r > g + 60 and r > b + 60 for r, g, b in pixels)


def test_conditional_table_font_preserves_automatic_color_declaration():
    from xml.etree import ElementTree as ET

    from aspose.words_foss.docx_reader.ldm_builder.definitions import _tsp_build_font
    from aspose.words_foss.docx_writer.runs import render_rPr
    font = _tsp_build_font(ET.fromstring(
        '<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:color w:val="auto"/></w:rPr>'))
    assert font.color == 'Color [Empty]' and font.color_explicit is True
    assert 'w:color w:val="auto"' in render_rPr(font, for_style=True)
