"""Separate note stories use live size declarations without changing public traversal."""
from copy import deepcopy
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pymupdf
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._comments import _comment_story_nodes
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import OoxmlSaveOptions
from tests.test_conversion_notes import W, owned_notes, package


def owned_model(kind, size):
    with ZipFile(BytesIO(owned_notes(kind))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    styles = ET.fromstring(parts['word/styles.xml'])
    normal = next(s for s in styles.findall(W + 'style') if s.get(W + 'styleId') == 'Normal')
    normal.set(W + 'default', 'off')
    if size is not None:
        props = ET.SubElement(normal, W + 'rPr')
        ET.SubElement(props, W + 'sz', {W + 'val': str(size * 2)})
    parts['word/styles.xml'] = ET.tostring(styles)
    return aw.Document(BytesIO(package(parts))).light_document_model


def story(model):
    return next(s for s in model.source_stories if s.kind in ('footnote', 'endnote') and s.note_type == 'normal')


def fonts(model):
    nodes = _comment_story_nodes(story(model).children)
    result = [n.font for n in nodes if isinstance(n, (ldm.Run, ldm.NoteMarker))]
    result.extend(n.font for p in model.get_child_nodes(ldm.NodeType.PARAGRAPH, True)
                  for n in p._children if isinstance(n, ldm.NoteReference))
    return result


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
@pytest.mark.parametrize('size', [None, 0, 12, 16])
@pytest.mark.parametrize('route', ['original', 'json', 'deep'])
def test_note_size_views_edit_and_two_real_cold_generations(kind, size, route):
    model = owned_model(kind, size)
    if route == 'json':
        model = ldm.Document.model_validate_json(model.model_dump_json())
    elif route == 'deep':
        model = model.model_copy(deep=True)
    held = fonts(model)
    snapshot = model.model_dump_json()
    expected = 11 if size is None else size
    assert {f.size for f in held} == {expected}
    assert {f.render_size for f in held} == {expected or .5}
    assert model.model_dump_json() == snapshot
    next(s for s in model.styles if s.name == 'Normal').font.size = 28
    assert {f.size for f in held} == {28}
    for fmt in (aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC):
        current = model.model_copy(deep=True)
        for _ in range(2):
            raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
            current = aw.Document(BytesIO(raw)).light_document_model
            assert {f.size for f in fonts(current)} == {28}
    assert {f.size for f in held} == {28}


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_note_ownership_adoption_copy_and_shallow_alias(kind):
    model = owned_model(kind, 16)
    held = fonts(model)
    shallow = model.model_copy()
    deep = model.model_copy(deep=True)
    raw = {'styles': model.styles, 'source_stories': model.source_stories, 'sections': model.sections}
    declaration = model.model_dump_json()
    other = ldm.Document.model_validate(raw)
    assert other.source_stories[0] is not model.source_stories[0]
    assert model.model_dump_json() == declaration
    next(s for s in other.styles if s.name == 'Normal').font.size = 22
    assert {f.size for f in held} == {16}
    assert {f.size for f in fonts(other)} == {22}
    next(s for s in model.styles if s.name == 'Normal').font.size = 20
    assert fonts(shallow)[0] is held[0]
    assert {f.size for f in fonts(shallow)} == {20}
    assert {f.size for f in fonts(deep)} == {16}


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_first_unbound_story_adoption_keeps_caller_alias(kind):
    original = owned_model(kind, 16)
    typed = deepcopy({'styles': original.styles, 'source_stories': original.source_stories,
                      'sections': original.sections})
    owned = next(s for s in typed['source_stories'] if s.note_type == 'normal')
    caller = owned.children[0].runs[0].font
    raw = caller.model_dump()
    model = ldm.Document.model_validate(typed)
    assert story(model) is owned
    assert caller is fonts(model)[1]  # Body mark is the first ordered note child.
    assert caller.model_dump() == raw
    assert caller.size == 16
    next(s for s in model.styles if s.name == 'Normal').font.size = 20
    assert caller.size == 20


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_held_note_run_follows_move_to_body_and_back(kind):
    model = owned_model(kind, 16)
    note = story(model).children[0]
    node = note.runs[0]
    held = node.font
    destination = model.sections[0].body.paragraphs[0]
    model.styles.append(ldm.Style(name='OwnedParagraph', type=1,
                                font=ldm.Font(size=22, size_explicit=True)))
    destination.paragraph_format.style_name = 'OwnedParagraph'
    note._children.remove(node)
    destination._children.append(node)
    assert held.size == held.render_size == 22
    with pytest.warns(aw.ContentLossWarning, match='Footnotes/endnotes'):
        raw = LdmPdfWriter().write_to_bytes(model)
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        spans = [span for page in pdf for span in page.get_texttrace()
                 if 'OWNED NOTE BEFORE' in ''.join(chr(c[0]) for c in span['chars'])]
        assert len(spans) == 1
        assert spans[0]['size'] == pytest.approx(22, abs=.011)
    destination._children.remove(node)
    note._children.append(node)
    assert held.size == held.render_size == 16
    node.font.size = 0
    node.font.size_explicit = True
    assert held.size == 0
    assert held.render_size == .5


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_held_note_marker_and_anchor_follow_paragraph_moves(kind):
    model = owned_model(kind, 16)
    ordinary = story(model).children[0]
    marker = next(n for n in ordinary._children if isinstance(n, ldm.NoteMarker))
    body = model.sections[0].body.paragraphs[0]
    anchor = next(n for n in body._children if isinstance(n, ldm.NoteReference))
    model.styles.append(ldm.Style(name='OwnedDestination', type=1,
                                font=ldm.Font(size=30, size_explicit=True)))
    destination = ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(style_name='OwnedDestination'))
    story(model).children.append(destination)
    held = marker.font
    ordinary._children.remove(marker)
    destination._children.append(marker)
    assert held.size == held.render_size == 30
    destination._children.remove(marker)
    ordinary._children.append(marker)
    assert held.size == 16
    moved_anchor = ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(style_name='OwnedDestination'))
    model.sections[0].body.children.append(moved_anchor)
    body._children.remove(anchor)
    moved_anchor._children.append(anchor)
    assert anchor.font.size == anchor.font.render_size == 30


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_held_note_table_run_uses_live_table_context(kind):
    model = owned_model(kind, None)
    table = next(n for n in story(model).children if isinstance(n, ldm.Table))
    paragraph = table.rows[0].cells[0].paragraphs[0]
    node = paragraph.runs[0]
    model.styles.append(ldm.Style(name='OwnedTable', type=3,
                                font=ldm.Font(size=24, size_explicit=True)))
    table.style_name = 'OwnedTable'
    assert node.font.size == node.font.render_size == 24
    destination = story(model).children[0]
    paragraph._children.remove(node)
    destination._children.append(node)
    assert node.font.size == node.font.render_size == 11
    destination._children.remove(node)
    paragraph._children.append(node)
    assert node.font.size == node.font.render_size == 24


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_failed_save_keeps_note_size_views_and_declarations(kind, tmp_path, monkeypatch):
    document = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(owned_model(kind, 16))))
    held = fonts(document.light_document_model)
    next(s for s in document.styles if s.name == 'Normal').font.size = 28
    snapshot = document.light_document_model.model_dump_json()
    def fail(*args, **kwargs):
        raise OSError('owned failure')
    monkeypatch.setattr('aspose.words_foss.docx_writer.writer.LdmDocxWriter.write_to_bytes', fail)
    with pytest.raises(OSError, match='owned failure'):
        document.save(str(tmp_path / 'failed.docx'), aw.SaveFormat.DOCX)
    assert not (tmp_path / 'failed.docx').exists()
    assert document.light_document_model.model_dump_json() == snapshot
    assert {f.size for f in held} == {28}
