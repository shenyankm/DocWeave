"""Word text whitespace follows owned Aspose 26.9 source controls."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._flat_opc import encode
from aspose.words_foss.docx_reader.document_reader import DocumentReader
from aspose.words_foss.docx_reader.utils import _collect_run_text
from aspose.words_foss.docx_writer import LdmDocxWriter

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
X = '{http://www.w3.org/XML/1998/namespace}'


@pytest.mark.parametrize('text,ordinary,preserved', [
    (' TARGET ', 'TARGET', ' TARGET '),
    ('\tTARGET\t', 'TARGET', '\tTARGET\t'),
    ('\nTARGET\n', 'TARGET', 'TARGET '),
    ('\rTARGET\r', 'TARGET', ' TARGET '),
    ('\t  TARGET \n ', 'TARGET', '\t  TARGET   '),
    ('\xa0TARGET\xa0', '\xa0TARGET\xa0', '\xa0TARGET\xa0'),
    ('\u2002TARGET\u2002', '\u2002TARGET\u2002', '\u2002TARGET\u2002'),
    ('\u3000TARGET\u3000', '\u3000TARGET\u3000', '\u3000TARGET\u3000'),
    ('\ufeffTARGET\ufeff', '\ufeffTARGET\ufeff', '\ufeffTARGET\ufeff'),
    (' \t\r\n ', '', ' \t  '),
    ('A\r\nB', 'A B', 'A B'),
    ('A\n\rB', 'A  B', 'A  B'),
    ('\r\n\nA', 'A', '  A'),
    ('\n\n A', 'A', ' A'),
    ('TARGET\tX', 'TARGET X', 'TARGET\tX'),
])
@pytest.mark.parametrize('space', [None, 'default', 'preserve'])
def test_shared_text_readers(text, ordinary, preserved, space):
    run = ET.Element(W + 'r', {X + 'space': 'preserve'})
    target = ET.SubElement(run, W + 't')
    target.text = text
    if space is not None:
        target.set(X + 'space', space)
    expected = preserved if space == 'preserve' else ordinary
    assert _collect_run_text(run) == expected
    reader = DocumentReader()
    assert reader._parse_run(run).text == expected
    paragraph = ET.Element(W + 'p')
    link = ET.SubElement(paragraph, W + 'hyperlink')
    link.append(run)
    assert reader._parse_paragraph(paragraph).text == expected


def test_field_code_text_is_not_normalized():
    run = ET.Element(W + 'r')
    ET.SubElement(run, W + 'instrText').text = ' HYPERLINK "https://example.test/" '
    assert _collect_run_text(run) == ' HYPERLINK "https://example.test/" '


def owned_text(space):
    stream = BytesIO()
    with ZipFile(stream, 'w') as package:
        package.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        package.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        document = ET.Element(W + 'document')
        body = ET.SubElement(document, W + 'body')
        paragraph = ET.SubElement(body, W + 'p')
        target = ET.SubElement(ET.SubElement(paragraph, W + 'r'), W + 't')
        target.text = '  OWNED\tTEXT  '
        if space is not None:
            target.set(X + 'space', space)
        ET.SubElement(body, W + 'sectPr')
        package.writestr('word/document.xml', ET.tostring(document))
    return stream.getvalue()


@pytest.mark.parametrize('space,expected', [(None, 'OWNED TEXT'), ('default', 'OWNED TEXT'), ('preserve', '  OWNED\tTEXT  ')])
@pytest.mark.parametrize('flat_input', [False, True])
@pytest.mark.parametrize('clone', ['direct', 'deep', 'json'])
def test_text_edit_and_cold_save(space, expected, flat_input, clone):
    raw = owned_text(space)
    model = aw.Document(BytesIO(encode(raw) if flat_input else raw)).light_document_model
    if clone == 'deep':
        model = model.model_copy(deep=True)
    elif clone == 'json':
        model = ldm.Document.model_validate_json(model.model_dump_json())
    run = model.sections[0].body.paragraphs[0].runs[0]
    assert run.text == expected
    run.text += ' EDITED'
    for _ in range(2):
        raw = LdmDocxWriter().write_to_bytes(model)
        model = aw.Document(BytesIO(encode(raw) if flat_input else raw)).light_document_model
        assert ''.join(r.text for r in model.sections[0].body.paragraphs[0].runs) == expected + ' EDITED'


def test_invalid_text_space_is_rejected():
    with pytest.raises(ValueError, match='Invalid xml:space'):
        aw.Document(BytesIO(owned_text('unexpected')))


@pytest.mark.parametrize('root_space,local_space,expected', [
    ('preserve', None, '  OWNED\tTEXT  '),
    ('preserve', 'default', '  OWNED\tTEXT  '),
    ('preserve', 'preserve', '  OWNED\tTEXT  '),
    ('default', None, 'OWNED TEXT'),
    ('default', 'default', 'OWNED TEXT'),
    ('default', 'preserve', '  OWNED\tTEXT  '),
])
def test_part_root_text_space(root_space, local_space, expected):
    raw = owned_text(local_space)
    output = BytesIO()
    with ZipFile(BytesIO(raw)) as source, ZipFile(output, 'w') as destination:
        for name in source.namelist():
            data = source.read(name)
            if name == 'word/document.xml':
                root = ET.fromstring(data)
                root.set(X + 'space', root_space)
                data = ET.tostring(root)
            destination.writestr(name, data)
    model = aw.Document(BytesIO(output.getvalue())).light_document_model
    assert model.sections[0].body.paragraphs[0].runs[0].text == expected


@pytest.mark.parametrize('root_space', [None, 'default', 'preserve'])
@pytest.mark.parametrize('local_space', [None, 'default', 'preserve'])
def test_comment_root_does_not_preserve_text(root_space, local_space):
    root = ET.Element(W + 'comments')
    run = ET.SubElement(ET.SubElement(ET.SubElement(root, W + 'comment'), W + 'p'), W + 'r')
    target = ET.SubElement(run, W + 't')
    target.text = '  OWNED SPACE  '
    if root_space is not None:
        root.set(X + 'space', root_space)
    if local_space is not None:
        target.set(X + 'space', local_space)
    DocumentReader._prepare_text_space(root)
    assert _collect_run_text(run) == ('  OWNED SPACE  ' if local_space == 'preserve' else 'OWNED SPACE')
