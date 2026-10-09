"""Ordinary size decoding matches frozen native getters without silent fallback."""

import json
import os
import runpy
import subprocess
import sys
import traceback
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, DocxDocument
from aspose.words_foss.docx_reader.ldm_builder.definitions import _tsp_build_font
from aspose.words_foss.utils.xml_helpers import W_NS

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / 'docs/benchmarks'
REPORT = json.loads((BENCHMARKS / 'font-size-loading-26.9.json').read_text())
ROWS = [row for row in REPORT['records'] if 'loaded' in row and 0 < row['loaded']['run_size'] < 1000]


@pytest.mark.parametrize('row', ROWS, ids=[row['input'] for row in ROWS])
def test_positive_ordinary_sizes_match_native_and_roundtrip(row):
    with ZipFile(BENCHMARKS / REPORT['corpus']) as archive:
        data = archive.read(row['input'])
    document = DocxDocument(BytesIO(data))
    run = document.body.paragraphs[0].runs[0]
    expected = row['loaded']['run_size']
    assert run.effective_font.size == expected
    if row['scope'] == 'run':
        assert run.font.size == expected
    else:
        assert run.font.size is None
        assert document.styles.get_by_name('Derived').font.size == expected
    with ZipFile(BytesIO(document.to_bytes())) as saved, ZipFile(BytesIO(data)) as original:
        assert {name: saved.read(name) for name in saved.namelist()} == {
            name: original.read(name) for name in original.namelist()}
    model = document.to_light_document()
    assert model.sections[0].body.paragraphs[0].runs[0].font.size == expected
    output = Document(BytesIO(data)).to_bytes('docx')
    reopened = DocxDocument(BytesIO(output))
    assert reopened.body.paragraphs[0].runs[0].effective_font.size == row['cold']['run_size']
    assert 'IMPORT' in reopened.body.paragraphs[0].text


def test_font_size_roundtrip_accepts_windows_zip_metadata(monkeypatch):
    from zipfile import ZipInfo

    original = ZipInfo.__init__

    def windows_info(self, *args, **kwargs):
        original(self, *args, **kwargs)
        self.create_system = 0

    monkeypatch.setattr(ZipInfo, '__init__', windows_info)
    test_positive_ordinary_sizes_match_native_and_roundtrip(ROWS[0])


BAD = ('PRIVATE FONT VALUE', '', '0', '-2', '2_4', '٢٤', 'NaN', 'Infinity', '24hp', ' 24 ', '9' * 400)


@pytest.mark.parametrize('scope', ['run', 'style', 'default'])
@pytest.mark.parametrize('value', BAD)
def test_invalid_sizes_are_private_and_never_silently_inherited(scope, value):
    inputs = runpy.run_path(str(ROOT / 'docs/probes/font_size_loading.py'))['inputs']
    data = next(data for _, kind, raw, data in inputs() if kind == scope and raw == value)
    document = DocxDocument(BytesIO(data))
    before = document.to_bytes()
    for read in (lambda: document.body.paragraphs[0].runs[0].effective_font.size,
                 document.to_light_document, lambda: Document(BytesIO(data))):
        with pytest.raises(ValueError, match='OOXML font size') as caught:
            read()
        trace = ''.join(traceback.format_exception(caught.value))
        if len(value) > 1:
            assert value not in trace
        assert 'IMPORT' not in trace and caught.value.__suppress_context__
        assert document.to_bytes() == before


@pytest.mark.parametrize('value,expected', [('24.0', 12), ('6.25pt', 6), ('0.5in', 36)])
def test_conditional_table_style_uses_the_same_size_decoder(value, expected):
    rpr = ET.fromstring(f'<w:rPr xmlns:w="{W_NS[1:-1]}"><w:sz w:val="{value}"/></w:rPr>')
    assert _tsp_build_font(rpr).size == expected


def test_conditional_table_style_does_not_hide_invalid_size():
    rpr = ET.fromstring(f'<w:rPr xmlns:w="{W_NS[1:-1]}"><w:sz w:val="PRIVATE FONT VALUE"/></w:rPr>')
    with pytest.raises(ValueError, match='OOXML font size'):
        _tsp_build_font(rpr)


def test_cli_keeps_existing_output_on_invalid_font_size(tmp_path):
    inputs = runpy.run_path(str(ROOT / 'docs/probes/font_size_loading.py'))['inputs']
    data = next(data for _, kind, raw, data in inputs() if kind == 'run' and raw == 'PRIVATE FONT VALUE')
    source, output = tmp_path / 'input.docx', tmp_path / 'output.pdf'
    source.write_bytes(data)
    output.write_bytes(b'existing')
    result = subprocess.run([sys.executable, '-m', 'aspose.words_foss.convert', str(source), str(output)],
                            capture_output=True, text=True, check=False, timeout=20)
    error = 'requires POSIX process groups' if os.name == 'nt' else 'OOXML font size'
    assert result.returncode != 0 and error in result.stderr
    assert 'PRIVATE FONT VALUE' not in result.stderr and 'IMPORT' not in result.stderr
    assert output.read_bytes() == b'existing' and not list(tmp_path.glob('.conversion-*'))


@pytest.mark.parametrize('index', [7, 8, 10])
def test_rendered_size_is_measured_independently(index):
    import pymupdf

    row = next(row for row in ROWS if row['input'] == f'run-{index:02d}.docx')
    with ZipFile(BENCHMARKS / REPORT['corpus']) as archive:
        data = archive.read(row['input'])
    raw = Document(BytesIO(data)).to_bytes('pdf')
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        spans = [span for page in pdf for block in page.get_text('dict')['blocks']
                 for line in block.get('lines', []) for span in line['spans'] if span['text'] == 'IMPORT']
    assert len(spans) == 1
    assert spans[0]['size'] == pytest.approx(row['loaded']['run_size'], abs=.001)


@pytest.mark.parametrize('scope', ['run', 'style', 'default'])
def test_font_size_units_survive_cross_document_import(scope):
    with ZipFile(BENCHMARKS / REPORT['corpus']) as archive:
        source = DocxDocument(BytesIO(archive.read(f'{scope}-07.docx')))
    helpers = runpy.run_path(str(ROOT / 'docs/probes/import_style_conflicts.py'))
    defaults = ('<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="6.25pt"/></w:rPr></w:rPrDefault></w:docDefaults>'
                if scope == 'default' else '')
    data = helpers['document'](defaults + helpers['style']('Target', 'paragraph', {}),
                               '<w:pStyle w:val="Target"/>', '')
    target = DocxDocument(BytesIO(data))
    before = source.to_bytes()
    copied = target.import_node(source.body.paragraphs[0], True)
    assert copied.owner_document is target and copied.parent_node is None
    target.body.append_child(copied)
    assert source.to_bytes() == before
    reopened = DocxDocument(BytesIO(target.to_bytes()))
    assert reopened.body.paragraphs[-1].runs[0].effective_font.size == 6


def test_size_unit_conversion_does_not_depend_on_decimal_context():
    from decimal import localcontext

    from aspose.words_foss.utils.xml_helpers import parse_font_size

    with localcontext() as context:
        context.prec = 2
        assert parse_font_size('25.4mm') == parse_font_size('2.54cm') == 72
        assert parse_font_size('6.25pt') == 6
