"""Closing embedding strips output resources without erasing original recoverable snapshots."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat
from aspose.words_foss._flat_opc import decode, is_flat_opc
from aspose.words_foss._opc import resolve_target
from .test_source_font_table import W, R, TABLE, FONT, archive, owned_parts, relationships_path, rels


def unpack(data):
    with ZipFile(BytesIO(decode(data) if is_flat_opc(data) else data)) as package:
        return {name: package.read(name) for name in package.namelist()}


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('original', [None, False, True])
def test_disable_save_twice_preserves_original_graph_and_held_metadata(tmp_path, format, original):
    doc = Document(BytesIO(archive(owned_parts(original))))
    held = doc.font_infos[0]; collection = doc.font_infos
    collection.embed_true_type_fonts = False
    source = doc.light_document_model.source_font_table
    held.alt_name = 'Edited'; source = doc.light_document_model.source_font_table
    before = doc.light_document_model.model_dump_json()
    doc.to_bytes(format)
    assert doc.light_document_model.model_dump_json() == before
    for i in range(2):
        output = tmp_path / str(i); doc.save(output, format)
        parts = unpack(output.read_bytes())
        assert not any(name.endswith('.odttf') for name in parts)
        assert 'word/_rels/fontTable.xml.rels' not in parts
        assert ET.fromstring(parts['word/fontTable.xml']).find(f'.//{{{W}}}embedRegular') is None
        assert not any(item.get('ContentType', '').endswith('obfuscatedFont') for item in ET.fromstring(parts['[Content_Types].xml']))
        assert doc.font_infos is collection and collection[0] is held and held.alt_name == 'Edited'
        cold = Document(output)
        assert cold.font_infos[0].alt_name == 'Edited' and cold.font_infos.embed_true_type_fonts is False
        assert doc.light_document_model.source_font_table.data == source.data
        assert doc.light_document_model.source_font_table.parts == source.parts
    # Turning embedding back on still has the original bytes available. Native
    # true-case font selection and reconstruction remain a separate stage.
    collection.embed_true_type_fonts = True
    restored = Document(BytesIO(doc.to_bytes(format))).light_document_model.source_font_table
    assert [p.data for p in restored.parts if p.content_type.endswith('obfuscatedFont')] == [p.data for p in source.parts if p.name == FONT]


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_shared_resource_with_nonfont_reference_is_retained(format):
    parts = owned_parts(False)
    root = ET.fromstring(parts[TABLE]);font = root.find(f'{{{W}}}font')
    ET.SubElement(font, f'{{{W}}}custom', {f'{{{R}}}id': 'other'})
    parts[TABLE] = ET.tostring(root)
    parts[relationships_path(TABLE)] = rels(('font', 'font', '../fonts/tiny.odttf'), ('other', 'image', '../fonts/tiny.odttf'))
    doc = Document(BytesIO(archive(parts)))
    saved = unpack(doc.to_bytes(format))
    edges = ET.fromstring(saved['word/_rels/fontTable.xml.rels'])
    assert [edge.get('Id') for edge in edges] == ['other']
    resource = resolve_target('word/fontTable.xml', edges[0].get('Target'))
    assert saved[resource] == parts[FONT]
    assert any(item.get('PartName') == '/' + resource for item in ET.fromstring(saved['[Content_Types].xml']))


@pytest.mark.parametrize('damage', ['missing', 'bad-id', 'duplicate'])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_disabled_embedding_still_validates_original_graph_atomically(tmp_path, damage, format):
    doc = Document(BytesIO(archive(owned_parts(False))))
    source = doc.light_document_model.source_font_table
    updates = {'parts': source.parts[:-1]} if damage == 'missing' else {'parts': source.parts + source.parts[:1]}
    if damage == 'bad-id':
        data = source.data.replace(b'r:id="font"', b'r:id="missing"')
        updates = {'data': data, 'parts': tuple(p.model_copy(update={'data': data}) if p.name == TABLE else p for p in source.parts)}
    doc.light_document_model.source_font_table = source.model_copy(update=updates)
    before = doc.light_document_model.model_dump_json(); collection = doc.font_infos; font = collection[0]
    output = tmp_path / 'existing'; output.write_bytes(b'existing')
    with pytest.raises(ValueError): doc.save(output, format)
    assert output.read_bytes() == b'existing'
    assert doc.light_document_model.model_dump_json() == before
    assert doc.font_infos is collection and collection[0] is font
