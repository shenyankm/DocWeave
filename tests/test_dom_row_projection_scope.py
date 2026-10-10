"""Only supported Word content participates in row validation/save projection."""
from io import BytesIO
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom import DocxDocument
from .test_dom_row_format import row_document
from .test_dom_story_row_height import W


@pytest.mark.parametrize('value', ['00040000', '1PT'])
@pytest.mark.parametrize('valid_row', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_opaque_row_declarations_are_preserved(tmp_path, value, valid_row, flat):
    source, parts = row_document(tmp_path, '<w:trHeight w:val="40000" w:hRule="exact"/>')
    opaque = (f'<ext:opaque xmlns:ext="urn:owned:opaque" xmlns:w="{W}">'
              '<w:tbl><w:tr><w:trPr>'
              f'<w:trHeight w:val="{value}" w:hRule="exact"/>'
              '</w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl></ext:opaque>').encode()
    data = parts['word/document.xml']
    if not valid_row:
        start, end = data.index(b'<w:tbl>'), data.index(b'</w:tbl>') + len(b'</w:tbl>')
        data = data[:start] + data[end:]
    parts['word/document.xml'] = data.replace(b'<w:sectPr/>', opaque + b'<w:sectPr/>')
    with ZipFile(source, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    document = DocxDocument(source)
    node = next(n for n in document.body.child_nodes if n._element.namespaceURI == 'urn:owned:opaque')
    original = node.xml
    held_tree = node._element
    row = node.get_child_nodes(aw.NodeType.ROW, True)[0]
    with pytest.raises(NotImplementedError):
        row.row_format.height = 2
    if not valid_row:
        assert document._package.payload('word/document.xml') == parts['word/document.xml']
    export = document.to_flat_opc() if flat else document.to_bytes()
    cold = DocxDocument(BytesIO(export))
    cold_node = next(n for n in cold.body.child_nodes if n._element.namespaceURI == 'urn:owned:opaque')
    assert cold_node._element.getElementsByTagNameNS(W, 'trHeight')[0].getAttributeNS(W, 'val') == value
    assert node.xml == original
    assert node._element is held_tree
    output = tmp_path / 'result'
    (document.save_flat_opc if flat else document.save)(output)
    assert node.xml == original
    assert node._element is held_tree
    if valid_row:
        assert document.body.tables[0].rows[0].row_format.height == 1584
    elif not flat:
        assert DocxDocument(output)._package.payload('word/document.xml') == parts['word/document.xml']
