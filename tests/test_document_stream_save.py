from io import BytesIO, StringIO

import pytest
import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln.saving import SaveOutputParameters

FORMATS = [dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC, dk.SaveFormat.FLAT_OPC_MACRO_ENABLED,
           dk.SaveFormat.FLAT_OPC_TEMPLATE, dk.SaveFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED,
           dk.SaveFormat.PDF]
MIMES = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document',
         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
         'application/vnd.ms-word.document.macroEnabled.12',
         'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
         'application/vnd.ms-word.template.macroEnabled.12', 'application/pdf']


def document():
    doc = dk.Document()
    doc.light_document_model.sections[0].body.children = [ldm.Paragraph(children=[ldm.Run(text='OWNED STREAM')])]
    return doc


@pytest.mark.parametrize('fmt,mime', list(zip(FORMATS, MIMES)))
@pytest.mark.parametrize('file_stream', [False, True])
def test_stream_position_tail_and_metadata(tmp_path, fmt, mime, file_stream):
    path = tmp_path/'owned.bin'
    initial = b'Z'*200000
    path.write_bytes(initial)
    stream = path.open('r+b') if file_stream else BytesIO(initial)
    try:
        stream.seek(17)
        result = document().save(stream, fmt)
        end = stream.tell()
        assert not stream.closed
        stream.seek(0)
        data = stream.read()
        start = 0 if fmt == dk.SaveFormat.DOCX else 17
        assert data[:start] == initial[:start]
        assert data[end:] == initial[end:] and len(data) >= len(initial)
        output = data[start:end]
        if fmt == dk.SaveFormat.PDF:
            import fitz
            with fitz.open(stream=output, filetype='pdf') as pdf:
                assert 'OWNED STREAM' in ''.join(page.get_text() for page in pdf)
        else:
            assert 'OWNED STREAM' in dk.Document(BytesIO(output)).get_text()
        assert isinstance(result, SaveOutputParameters) and result.content_type == mime
        with pytest.raises(AttributeError):result.content_type = 'changed'
    finally:
        stream.close()


@pytest.mark.parametrize('fmt,mime', list(zip(FORMATS, MIMES)))
def test_path_metadata_keeps_existing_save_contract(tmp_path, fmt, mime):
    path = tmp_path/'owned.output'
    assert document().save(path, fmt).content_type == mime
    assert path.is_file()


def test_output_metadata_has_no_public_constructor():
    with pytest.raises(TypeError):SaveOutputParameters()


@pytest.mark.parametrize('fmt', [None, dk.SaveFormat.DOC, 'unknown'])
def test_invalid_format_does_not_change_stream(fmt):
    stream = BytesIO(b'OWNED old bytes');stream.seek(3)
    with pytest.raises(TypeError if fmt is None else ValueError):document().save(stream, fmt)
    assert stream.tell() == 3 and stream.getvalue() == b'OWNED old bytes'


def test_nonbinary_targets_are_rejected_before_writing():
    target = StringIO('OWNED');target.seek(2)
    with pytest.raises(TypeError):document().save(target, dk.SaveFormat.PDF)
    assert target.getvalue() == 'OWNED' and target.tell() == 2


@pytest.mark.parametrize('closed', [False, True])
def test_unwritable_stream_errors_leave_caller_ownership(tmp_path, closed):
    path = tmp_path/'owned';path.write_bytes(b'OWNED')
    with path.open('rb') as target:
        if closed:target.close()
        with pytest.raises(RuntimeError):document().save(target, dk.SaveFormat.PDF)
        assert target.closed == closed
    assert path.read_bytes() == b'OWNED'


class Nonseek(BytesIO):
    def seek(self, *args):raise OSError('OWNED seek unavailable')


def test_pdf_does_not_require_seek_but_docx_does():
    stream = Nonseek()
    document().save(stream, dk.SaveFormat.PDF)
    from pypdf import PdfReader
    assert 'OWNED STREAM' in PdfReader(BytesIO(stream.getvalue())).pages[0].extract_text()
    before = stream.getvalue()
    with pytest.raises(RuntimeError):document().save(stream, dk.SaveFormat.DOCX)
    assert stream.getvalue() == before and not stream.closed


class Short(BytesIO):
    def write(self, data):return super().write(data[:7])


def test_short_writes_report_failure_without_closing():
    stream = Short()
    with pytest.raises(RuntimeError):document().save(stream, dk.SaveFormat.DOCX)
    assert len(stream.getvalue()) == 7 and not stream.closed


class Stalled(BytesIO):
    def write(self, data):return 0


def test_stalled_stream_reports_failure_without_closing():
    stream = Stalled()
    with pytest.raises(RuntimeError):document().save(stream, dk.SaveFormat.PDF)
    assert not stream.closed and stream.getvalue() == b''


@pytest.mark.parametrize('option,mime', [(dk.saving.OoxmlSaveOptions(), MIMES[0]), (dk.saving.PdfSaveOptions(), 'application/pdf')])
def test_options_overload(option, mime):
    stream = BytesIO()
    result = document().save(stream, option)
    assert stream.getvalue() and isinstance(result, SaveOutputParameters)
    assert result.content_type == mime






@pytest.mark.parametrize('option', [dk.saving.OoxmlSaveOptions()])
def test_mutated_options_fail_before_touching_stream(option):
    option.save_format = dk.SaveFormat.PDF
    stream = BytesIO(b'OWNED unchanged');stream.seek(4)
    with pytest.raises(ValueError):document().save(stream, option)
    assert stream.tell() == 4 and stream.getvalue() == b'OWNED unchanged'


@pytest.mark.parametrize('fmt', [dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('failed', [False, True])
def test_source_metadata_commits_only_after_successful_write(fmt, failed):
    doc = document()
    row = ldm.Row(row_format=ldm.RowFormat(height=2000, height_rule=1),
                  cells=[ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(text='OWNED ROW')])])])
    doc.light_document_model.sections[0].body.children.append(ldm.Table(rows=[row]))
    before = doc.light_document_model.model_dump_json()
    stream = Stalled() if failed else BytesIO()
    if failed:
        with pytest.raises(RuntimeError):doc.save(stream, fmt)
        assert doc.light_document_model.model_dump_json() == before
    else:
        doc.save(stream, fmt)
        assert row.row_format.height == 1584
        assert doc.light_document_model.source_font_table is not None
        cold = dk.Document(BytesIO(stream.getvalue())).light_document_model
        assert cold.sections[0].body.children[1].rows[0].row_format.height == 1584
