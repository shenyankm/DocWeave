"""Font readers close promptly on every public conversion exit path."""

from io import BytesIO
import gc
import warnings
import weakref

from fpdf import FPDF
from pypdf import PdfReader
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import PdfFontSubstitutionWarning
from aspose.words_foss.pdf_writer.writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def document(font_name=''):
    return ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[
        ldm.Paragraph(children=[ldm.Run(text='RECOVERED', font=ldm.Font(name=font_name))])]))])


@pytest.fixture
def font_files(monkeypatch):
    files = []
    original = FPDF.add_font

    def track(pdf, *args, **kwargs):
        result = original(pdf, *args, **kwargs)
        files.extend(font.ttfont.reader.file for font in pdf.fonts.values()
                     if font.type == 'TTF' and font.ttfont.reader is not None
                     and font.ttfont.reader.file not in files)
        return result

    monkeypatch.setattr(FPDF, 'add_font', track)
    yield files
    for file in files:
        file.close()


def test_fatal_font_warning_closes_readers_and_preserves_output(tmp_path, font_files):
    writer = LdmPdfWriter()
    output = tmp_path / 'existing.pdf'
    output.write_bytes(b'KEEP')
    with warnings.catch_warnings():
        warnings.simplefilter('error', PdfFontSubstitutionWarning)
        with pytest.raises(PdfFontSubstitutionWarning):
            writer.write(document('SimSun'), output)
    assert font_files and all(file.closed for file in font_files)
    assert output.read_bytes() == b'KEEP'
    assert writer._measurement_pdf is None
    raw = writer.write_to_bytes(document())
    assert 'RECOVERED' in PdfReader(BytesIO(raw)).pages[0].extract_text()
    assert all(file.closed for file in font_files)


def test_partial_font_registration_failure_closes_existing_readers(tmp_path, font_files):
    options = PdfSaveOptions()
    options.fallback_fonts = [str(tmp_path / 'missing.ttf')]
    writer = LdmPdfWriter(options)
    with pytest.raises(FileNotFoundError):
        writer.write_to_bytes(document())
    assert font_files and all(file.closed for file in font_files)
    options.fallback_fonts = []
    assert writer.write_to_bytes(document()).startswith(b'%PDF')


def test_output_failure_closes_readers_and_clears_measurements(monkeypatch, font_files):
    writer = LdmPdfWriter()
    original = FPDF.output

    def fail(pdf, *args, **kwargs):
        raise RuntimeError('output failed')

    monkeypatch.setattr(FPDF, 'output', fail)
    with pytest.raises(RuntimeError, match='output failed'):
        writer.write_to_bytes(document())
    assert font_files and all(file.closed for file in font_files)
    assert writer._measurement_pdf is None and writer._measurement_writer is None
    monkeypatch.setattr(FPDF, 'output', original)
    assert writer.write_to_bytes(document()).startswith(b'%PDF')


def test_outline_output_failure_closes_fonts_preserves_file_and_recovers(tmp_path, monkeypatch, font_files):
    from aspose.words_foss.pdf_writer.outline import OutlineOutputProducer

    options = PdfSaveOptions()
    options.outline_options.headings_outline_levels = 6
    writer = LdmPdfWriter(options)
    source = document()
    source.sections[0].body.paragraphs[0].paragraph_format.is_heading = True
    source.sections[0].body.paragraphs[0].paragraph_format.outline_level = 0
    output = tmp_path / 'existing.pdf'
    output.write_bytes(b'KEEP')
    original = OutlineOutputProducer._add_document_outline

    def fail(producer):
        assert original(producer)[0] is not None
        raise RuntimeError('outline output failed')

    monkeypatch.setattr(OutlineOutputProducer, '_add_document_outline', fail)
    with pytest.raises(RuntimeError, match='outline output failed'):
        writer.write(source, output)
    assert output.read_bytes() == b'KEEP'
    assert font_files and all(file.closed for file in font_files)
    assert writer._measurement_pdf is None and writer._measurement_writer is None
    monkeypatch.setattr(OutlineOutputProducer, '_add_document_outline', original)
    assert 'RECOVERED' in PdfReader(BytesIO(writer.write_to_bytes(source))).pages[0].extract_text()


def test_invalid_expansion_releases_standalone_measurement_fonts(font_files):
    writer = LdmPdfWriter()
    writer._estimate_paragraph_height(document().sections[0].body.paragraphs[0], 60)
    assert font_files and not any(file.closed for file in font_files)
    writer.options.outline_options.expanded_outline_levels = 10
    with pytest.raises(ValueError, match='expanded_outline_levels'):
        writer.write_to_bytes(document())
    assert all(file.closed for file in font_files)
    assert writer._measurement_pdf is None and writer._measurement_writer is None


def test_navigation_hook_restores_on_layout_failure(tmp_path, monkeypatch, font_files):
    writer = LdmPdfWriter()
    source = document()
    source.sections[0].body.paragraphs[0]._children.insert(0, ldm.BookmarkStart(name='Target'))
    target = tmp_path / 'existing.pdf'
    target.write_bytes(b'KEEP')
    original = writer._paragraph_renderer._render_paragraph_body
    captured = []

    def fail(pdf, para):
        assert '_perform_page_break_if_need_be' in pdf.__dict__
        captured.append(pdf)
        raise RuntimeError('navigation layout failed')

    monkeypatch.setattr(writer._paragraph_renderer, '_render_paragraph_body', fail)
    with pytest.raises(RuntimeError, match='navigation layout failed'):
        writer.write(source, target)
    assert '_perform_page_break_if_need_be' not in captured[0].__dict__
    assert target.read_bytes() == b'KEEP'
    assert font_files and all(file.closed for file in font_files)
    monkeypatch.setattr(writer._paragraph_renderer, '_render_paragraph_body', original)
    assert 'RECOVERED' in PdfReader(BytesIO(writer.write_to_bytes(source))).pages[0].extract_text()


@pytest.mark.parametrize('kind', ['paragraph', 'table'])
def test_standalone_measurement_fonts_close_before_conversion(font_files, kind):
    writer = LdmPdfWriter()
    para = document().sections[0].body.paragraphs[0]
    if kind == 'paragraph':
        writer._estimate_paragraph_height(para, 60)
    else:
        table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[para])])])
        writer._estimate_table_height(table, 60)
    old_files = list(font_files)
    assert old_files and not any(file.closed for file in old_files)
    assert writer.write_to_bytes(document()).startswith(b'%PDF')
    assert all(file.closed for file in font_files)


def test_layout_failure_closes_readers_and_allows_retry(monkeypatch, font_files):
    writer = LdmPdfWriter()
    original = writer._paragraph_renderer.render_paragraph

    def fail(pdf, paragraph):
        raise RuntimeError('layout failed')

    monkeypatch.setattr(writer._paragraph_renderer, 'render_paragraph', fail)
    with pytest.raises(RuntimeError, match='layout failed'):
        writer.write_to_bytes(document())
    assert font_files and all(file.closed for file in font_files)
    assert writer._measurement_pdf is None and writer._measurement_writer is None
    monkeypatch.setattr(writer._paragraph_renderer, 'render_paragraph', original)
    assert writer.write_to_bytes(document()).startswith(b'%PDF')


def test_failure_after_drawing_does_not_retain_font_subset_cache(monkeypatch, font_files):
    writer = LdmPdfWriter()
    original = writer._paragraph_renderer.render_paragraph
    subsets = []

    def fail_after_drawing(pdf, paragraph):
        original(pdf, paragraph)
        subsets.extend(weakref.ref(font.subset) for font in pdf.fonts.values())
        raise RuntimeError('after drawing')

    monkeypatch.setattr(writer._paragraph_renderer, 'render_paragraph', fail_after_drawing)
    with pytest.raises(RuntimeError, match='after drawing'):
        writer.write_to_bytes(document())
    assert all(file.closed for file in font_files)
    gc.collect()
    assert subsets and all(subset() is None for subset in subsets)


def test_shaped_font_preparation_failure_closes_readers_and_recovers(monkeypatch, font_files):
    pytest.importorskip('uharfbuzz')
    from fontTools.ttLib.sfnt import SFNTWriter

    original = SFNTWriter.close

    def fail(writer):
        raise RuntimeError('font table preparation failed')

    options = PdfSaveOptions()
    options.text_shaping = True
    writer = LdmPdfWriter(options)
    monkeypatch.setattr(SFNTWriter, 'close', fail)
    with pytest.raises(RuntimeError, match='font table preparation failed'):
        writer.write_to_bytes(document())
    assert font_files and all(file.closed for file in font_files)
    assert writer._measurement_pdf is None
    monkeypatch.setattr(SFNTWriter, 'close', original)
    assert 'RECOVERED' in PdfReader(BytesIO(writer.write_to_bytes(document()))).pages[0].extract_text()
    assert all(file.closed for file in font_files)
