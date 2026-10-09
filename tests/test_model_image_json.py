"""Complete LDM JSON preserves binary images, including legacy header/footer views."""

import base64
from io import BytesIO
import json
from pathlib import Path
from zipfile import ZipFile

from PIL import Image
from pydantic import ValidationError
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter


def png():
    output = BytesIO()
    Image.new('RGB', (12, 6), 'navy').save(output, format='PNG')
    return output.getvalue()


@pytest.mark.parametrize('payload', [b'', b'plain ascii', b'\x00\xff\x89', png()])
def test_image_bytes_json_is_explicit_and_lossless(payload):
    image = ldm.ImageData(image_bytes=payload)
    encoded = json.loads(image.model_dump_json())['image_bytes']
    assert encoded == {'encoding': 'base64', 'data': base64.b64encode(payload).decode('ascii')}
    assert image.model_dump()['image_bytes'] == payload
    assert ldm.ImageData.model_validate_json(image.model_dump_json()).image_bytes == payload
    assert ldm.ImageData.model_validate(image.model_dump(mode='json')).image_bytes == payload


@pytest.mark.parametrize('text', ['旧模型文本', 'base64:AAAA'])
def test_legacy_utf8_image_string_is_not_reinterpreted(text):
    assert ldm.ImageData.model_validate_json(json.dumps({'image_bytes': text})).image_bytes == text.encode()


@pytest.mark.parametrize('encoded', [
    {'encoding': 'base64', 'data': '@@@@'},
    {'encoding': 'base64', 'data': 'abc'},
    {'encoding': 'base64', 'data': '中文'},
    {'encoding': 'base64', 'data': None},
    {'encoding': 'base64'},
    {'encoding': 'unknown', 'data': 'AAAA'},
])
def test_invalid_binary_json_is_rejected(encoded):
    with pytest.raises(ValidationError):
        ldm.ImageData.model_validate_json(json.dumps({'image_bytes': encoded}))


@pytest.mark.parametrize('by_alias', [False, True])
@pytest.mark.parametrize('mode', ['json', 'dict', 'python'])
def test_document_image_json_preserves_order_stories_and_docx_resources(by_alias, mode):
    payload = png()

    def paragraph(label):
        return ldm.Paragraph(children=[ldm.Run(text=label), ldm.Shape(
            is_inline=True, has_image=True, width=30, height=15,
            image_data=ldm.ImageData(image_type=ldm.ImageData.from_mime('image/png'),
                                     image_bytes=payload))])

    source = ldm.Document(sections=[ldm.Section(
        body=ldm.Body(children=[paragraph('BODY'), ldm.Table(rows=[ldm.Row(cells=[
            ldm.Cell(children=[paragraph('CELL')])])]), paragraph('AFTER')]),
        headers_footers=[ldm.HeaderFooter(header_footer_type=0, children=[paragraph('HEADER')]),
                        ldm.HeaderFooter(header_footer_type=1, children=[paragraph('FOOTER')])])])
    before = source.model_dump(by_alias=True)
    if mode == 'json':
        restored = ldm.Document.model_validate_json(source.model_dump_json(by_alias=by_alias))
    else:
        restored = ldm.Document.model_validate(source.model_dump(
            mode='json' if mode == 'dict' else 'python', by_alias=by_alias))
    assert restored.model_dump(by_alias=True) == before
    assert source.model_dump(by_alias=True) == before
    legacy = json.loads(source.model_dump_json(by_alias=by_alias))
    legacy['sections'][0].pop('headers_footers')
    assert ldm.Document.model_validate(legacy).model_dump(by_alias=True) == before
    assert [type(child).__name__ for child in restored.sections[0].body.children] == [
        'Paragraph', 'Table', 'Paragraph']
    shapes = restored.get_child_nodes(ldm.NodeType.SHAPE, True)
    assert len(shapes) == 5
    assert all(shape.image_data.image_bytes == payload for shape in shapes)
    output = LdmDocxWriter().write_to_bytes(restored)
    with ZipFile(BytesIO(output)) as archive:
        media = [archive.read(name) for name in archive.namelist() if name.startswith('word/media/')]
    assert media and all(data == payload for data in media)
    reread = aw.Document(BytesIO(output)).light_document_model
    assert len(reread.get_child_nodes(ldm.NodeType.SHAPE, True)) == 5


def test_existing_docx_image_corpus_round_trips_complete_model_json():
    paths = sorted((Path(__file__).parent/'data/input').glob('*.docx'))
    assert len(paths) >= 12
    for path in paths:
        source = aw.Document(path).light_document_model
        before = source.model_dump(by_alias=True)
        restored = ldm.Document.model_validate_json(source.model_dump_json(by_alias=True))
        assert restored.model_dump(by_alias=True) == before, path.name
