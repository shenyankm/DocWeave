"""DOCX notes export keeps inline positions, note bodies, and explicit losses."""

from io import BytesIO
from contextlib import nullcontext
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.md_writer import LdmMarkdownWriter
from aspose.words_foss.models import ConversionOptions
from .test_content_integrity import source_package, load_with_loss


def note_document(body=None):
    if body is None:
        return load_with_loss(source_package())
    with ZipFile(BytesIO(source_package())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    start, rest = parts['word/document.xml'].decode().split('<w:body>', 1)
    parts['word/document.xml'] = (start + '<w:body>' + body + '</w:body></w:document>').encode()
    stream = BytesIO()
    with ZipFile(stream, 'w') as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return load_with_loss(stream.getvalue())


def options():
    opts = aw.saving.MarkdownSaveOptions()
    opts.export_notes = True
    opts.paragraph_break = '\n'
    opts.link_export_mode = aw.saving.MarkdownLinkExportMode.INLINE
    return opts


def test_notes_have_inline_references_and_definitions_without_mutating_model(tmp_path):
    doc = note_document()
    before = doc.light_document_model.model_dump_json(by_alias=True)
    text = doc.to_bytes(options()).decode()
    assert 'body[^note1][^note2]' in text
    assert '[^note1]: footnote text' in text
    assert '[^note2]: endnote text' in text
    assert not any(item.code == 'markdown.notes_omitted' for item in doc.diagnostics)
    assert doc.light_document_model.model_dump_json(by_alias=True) == before
    path = tmp_path / 'notes.md'
    doc.save(path, options())
    assert path.read_text() == text
    restored = ldm.Document.model_validate_json(before)
    writer = LdmMarkdownWriter(ConversionOptions(export_notes=True))
    assert writer.write(restored) == text
    assert writer.write(restored) == text


@pytest.mark.parametrize('wrapper', ['{}', '<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>',
                                    '<w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>'])
def test_notes_between_text_in_one_run_and_in_nested_cells(wrapper):
    body = ('<w:p><w:r><w:t>before</w:t><w:footnoteReference w:id="7"/>'
            '<w:t>after</w:t><w:endnoteReference w:id="9"/></w:r></w:p>')
    doc = note_document(wrapper.format(body))
    with pytest.warns(aw.ContentLossWarning) if '<w:tbl><w:tr><w:tc><w:tbl>' in wrapper else nullcontext():
        text = doc.to_bytes(options()).decode()
    assert 'before[^note1]after[^note2]' in text
    assert '[^note1]: footnote text' in text


def test_notes_in_hyperlinks_keep_both_destinations_and_inline_order():
    doc = note_document('<w:p><w:hyperlink w:anchor="target"><w:r><w:t>before</w:t>'
                        '<w:footnoteReference w:id="7"/><w:t>after</w:t></w:r></w:hyperlink></w:p>')
    text = doc.to_bytes(options()).decode()
    assert '[before](#target)[^note1][after](#target)' in text
    assert '[^note1]: footnote text' in text


def test_missing_note_body_reports_specific_location():
    doc = note_document('<w:p><w:r><w:t>body</w:t><w:footnoteReference w:id="999"/></w:r></w:p>')
    with pytest.warns(aw.ContentLossWarning, match='body is missing'):
        doc.to_bytes(options())
    assert any(item.code == 'markdown.note_missing' and item.location == 'footnote:999'
               for item in doc.diagnostics)


def test_duplicate_note_ids_fail_before_atomic_publish(tmp_path):
    doc = note_document()
    doc.light_document_model.source_stories.append(next(
        story for story in doc.light_document_model.source_stories if story.kind == 'footnote'))
    output = tmp_path / 'notes.md'
    output.write_bytes(b'original')
    with pytest.raises(ValueError, match='Duplicate note'):
        doc.save(output, options())
    assert output.read_bytes() == b'original'


def test_hidden_references_do_not_expose_note_bodies():
    doc = note_document('<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t>'
                        '<w:footnoteReference w:id="7"/></w:r><w:r><w:t>public</w:t></w:r></w:p>')
    text = doc.to_bytes(options()).decode()
    assert text.strip() == 'public'


def test_code_note_reference_is_relocated_with_explicit_diagnostic():
    doc = note_document()
    opts = options()
    opts.style_map = {'Normal': 'Code'}
    with pytest.warns(aw.ContentLossWarning, match='moved after'):
        text = doc.to_bytes(opts).decode()
    assert '```\nbody\n```\n\n[^note1][^note2]' in text
    assert any(item.code == 'markdown.note_reference_relocated' for item in doc.diagnostics)


def test_inline_only_model_anchor_still_reports_default_omission():
    doc = aw.Document()
    doc._document = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[
        ldm.Paragraph(children=[ldm.Run(text='body'), ldm.NoteReference(kind='footnote', identifier='7')])]))])
    with pytest.warns(aw.ContentLossWarning, match='Footnotes/endnotes'):
        doc.to_bytes('md')
