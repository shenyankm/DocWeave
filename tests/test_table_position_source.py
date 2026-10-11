"""Floating source declarations survive copies without becoming effective values."""
import json
from io import BytesIO
from xml.etree import ElementTree as E
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from pydantic import ValidationError

import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln.writers.docx import LdmDocxWriter
from dockiln.saving import OoxmlSaveOptions
from .test_docx_dom import package

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


@pytest.mark.parametrize('position', [
    {'horzAnchor': 'margin', 'vertAnchor': 'paragraph', 'tblpXSpec': 'center', 'tblpY': '0'},
    {'horzAnchor': 'page', 'vertAnchor': 'text', 'tblpX': '-240', 'tblpY': '360',
     'leftFromText': '120', 'rightFromText': '0', 'topFromText': '40', 'bottomFromText': '80'},
])
def test_position_declarations_survive_json_and_two_saves(tmp_path, position):
    path = tmp_path / 'owned.docx'
    package(path, body=(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>OWNED_POSITION</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>'))
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(parts['word/document.xml'])
    E.SubElement(document.find('.//' + W + 'tblPr'), W + 'tblpPr',
                 {W + key: value for key, value in position.items()} | {W + 'unknown': 'ignored'})
    parts['word/document.xml'] = E.tostring(document)
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    model = dk.Document(path).light_document_model
    for original in (model, model.model_copy(deep=True),
                     ldm.Document.model_validate_json(model.model_dump_json())):
        for fmt in (dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC):
            value = original
            for _ in range(2):
                table = value.get_child_nodes(ldm.NodeType.TABLE, True)[0]
                assert table._tblp_pr_attrs == position
                assert table.text_wrapping == 1
                data = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(value)
                value = dk.Document(BytesIO(data)).light_document_model
                assert value.get_child_nodes(ldm.NodeType.TABLE, True)[0]._tblp_pr_attrs == position
    assert model.get_child_nodes(ldm.NodeType.TABLE, True)[0]._tblp_pr_attrs == position


@pytest.mark.parametrize('attributes', [
    [['unknown', 'x']], [['tblpX', 12]], [['tblpX', '0'], ['tblpX', '1']],
])
@pytest.mark.parametrize('global_assignment_validation', [False, True])
def test_invalid_position_payload_is_atomic(attributes, global_assignment_validation):
    class PositionTable(ldm.Table):
        model_config = {'validate_assignment': global_assignment_validation}

    table = PositionTable(source_position=ldm.SourceTablePosition(attributes=(('tblpX', '0'),)))
    before = table.model_dump_json()
    with pytest.raises(ValidationError):
        table.source_position = {'attributes': attributes}
    assert table.model_dump_json() == before
    with pytest.raises(ValidationError):
        ldm.Table.model_validate_json('{"source_position":{"attributes":' +
                                      json.dumps(attributes) + '}}')


def test_position_carrier_is_immutable_and_unset_is_sparse():
    table = ldm.Table()
    assert table.source_position is None
    assert table._tblp_pr_attrs == {}
    table._tblp_pr_attrs = {'tblpXSpec': 'right'}
    with pytest.raises(ValidationError):
        table.source_position.attributes = ()
    assert table._tblp_pr_attrs == {'tblpXSpec': 'right'}
    # The compatibility view cannot mutate the immutable source carrier.
    table._tblp_pr_attrs.clear()
    assert table._tblp_pr_attrs == {'tblpXSpec': 'right'}


@pytest.mark.parametrize('marker', [None, '', ' w:unknown="ignored"', ' w:tblpX=""'])
def test_empty_position_is_sparse_source_and_inline_effective(tmp_path, marker):
    path = tmp_path / 'empty-position.docx'
    position = '' if marker is None else f'<w:tblpPr{marker}/>'
    parts = package(path, body=(
        f'<w:tbl><w:tblPr>{position}</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>OWNED_POSITION</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl><w:sectPr/>'))
    model = dk.Document(path).light_document_model
    original_table = model.get_child_nodes(ldm.NodeType.TABLE, True)[0]
    assert original_table.text_wrapping == 0
    assert (original_table.source_position is None) == (marker is None)
    if marker is not None:
        assert original_table.source_position.attributes == ()
    # Package DOM keeps the raw declaration. Conversion saves normalize it.
    with ZipFile(BytesIO(dk.DocxDocument(path).to_bytes())) as archive:
        assert archive.read('word/document.xml') == parts['word/document.xml']
    for original in (model, model.model_copy(deep=True),
                     ldm.Document.model_validate_json(model.model_dump_json())):
        table = original.get_child_nodes(ldm.NodeType.TABLE, True)[0]
        assert table.text_wrapping == 0
        assert (table.source_position is None) == (marker is None)
        for fmt in (dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC):
            value = original
            for _ in range(2):
                data = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(value)
                if fmt == dk.SaveFormat.FLAT_OPC:
                    from dockiln._flat_opc import decode
                    data = decode(data)
                with ZipFile(BytesIO(data)) as archive:
                    root = E.fromstring(archive.read('word/document.xml'))
                    assert root.find('.//' + W + 'tblpPr') is None
                value = dk.Document(BytesIO(data)).light_document_model
                cold = value.get_child_nodes(ldm.NodeType.TABLE, True)[0]
                assert cold.text_wrapping == 0
                assert cold.source_position is None
    assert (original_table.source_position is None) == (marker is None)
