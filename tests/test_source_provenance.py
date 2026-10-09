"""Resolve provenance against independently parsed original XML, not model paths."""

from io import BytesIO
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader import DocumentReader
from aspose.words_foss.docx_writer import LdmDocxWriter
from .test_content_integrity import source_package, load_with_loss, interleaved_document


def blocks(items):
    for block in items:
        yield block
        if block['type'] == 'table':
            for row in block['rows']:
                for cell in row['cells']:
                    yield from blocks(cell['blocks'])


def verify_locations(raw, exported):
    count = 0
    with ZipFile(BytesIO(raw)) as archive:
        for block in blocks(exported['blocks']):
            provenance = block['provenance']
            element = fromstring(archive.read(provenance['part_name']))
            for index in provenance['child_path']:
                element = element[index]
            assert element.tag.endswith('}p' if block['type'] == 'paragraph' else '}tbl')
            if block['type'] == 'paragraph':
                assert ''.join(element.itertext()) == block['text']
            count += 1
        for story in exported['source_stories']:
            for block in blocks(story['blocks']):
                provenance = block['provenance']
                assert provenance['part_name'] == story['part_name']
                element = fromstring(archive.read(provenance['part_name']))
                for index in provenance['child_path']:
                    element = element[index]
                assert element.tag.endswith('}p' if block['type'] == 'paragraph' else '}tbl')
                count += 1
    return count


def test_body_and_nested_table_provenance_resolves_to_original_elements():
    raw = LdmDocxWriter().write_to_bytes(interleaved_document())
    doc = aw.Document(BytesIO(raw))
    assert verify_locations(raw, doc.to_dict()) == 5
    restored = ldm.Document.model_validate_json(doc.light_document_model.model_dump_json(by_alias=True))
    doc._document = restored
    assert verify_locations(raw, doc.to_dict()) == 5


def test_notes_and_nonstandard_header_part_locations_and_reader_reuse():
    raw = source_package()
    assert verify_locations(raw, load_with_loss(raw).to_dict()) >= 6
    reader = DocumentReader()
    with pytest.warns(aw.ContentLossWarning):
        reader.load_bytes(raw)
    first = reader.to_light_document()
    raw2 = LdmDocxWriter().write_to_bytes(interleaved_document())
    reader.load_bytes(raw2)
    second = reader.to_light_document()
    assert not second.source_stories
    assert all(p.source_location.part_name == 'word/document.xml' for p in second.all_paragraphs)
    assert any(story.part_name == 'notes/fn text.xml' for story in first.source_stories)


def test_plain_text_has_no_invented_xml_provenance():
    doc = aw.Document(BytesIO(b'plain text'))
    assert doc.to_dict()['blocks'][0]['provenance'] is None
