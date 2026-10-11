"""Section references determine header/footer ownership during source recovery."""
from io import BytesIO
from itertools import product
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def owned_story(reference, relationship, root, *, both=False):
    refs = ('header', 'footer') if both else (reference,)
    references = ''.join(f'<w:{role}Reference w:type="default" r:id="story"/>' for role in refs)
    document = (f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
                '<w:p><w:r><w:t>OWNED BODY</w:t></w:r></w:p>'
                f'<w:sectPr>{references}</w:sectPr></w:body></w:document>')
    tag = 'hdr' if root == 'header' else 'ftr'
    story = (f'<w:{tag} xmlns:w="{W}"><w:p><w:r><w:rPr><w:b/></w:rPr>'
             f'<w:t>OWNED STORY</w:t></w:r></w:p></w:{tag}>')
    stream = BytesIO()
    with ZipFile(stream, 'w') as package:
        package.writestr('[Content_Types].xml',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            f'<Override PartName="/word/story.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{relationship}+xml"/></Types>')
        package.writestr('_rels/.rels',
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>')
        package.writestr('word/_rels/document.xml.rels',
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="story" Type="{R}/{relationship}" Target="story.xml"/></Relationships>')
        package.writestr('word/document.xml', document)
        package.writestr('word/story.xml', story)
    return stream.getvalue()


CASES = [(a, b, c, False) for a, b, c in product(('header', 'footer'), repeat=3)]
CASES += [('header', kind, kind, True) for kind in ('header', 'footer')]


@pytest.mark.parametrize('reference,relationship,root,both', CASES)
@pytest.mark.parametrize('format_name', ['DOCX', 'FLAT_OPC'])
@pytest.mark.parametrize('via_json', [False, True])
def test_reference_role_survives_recovery_and_two_saves(tmp_path, reference, relationship, root, both, format_name, via_json):
    doc = aw.Document(BytesIO(owned_story(reference, relationship, root, both=both)))
    expected = {0, 1} if both else {int(reference == 'footer')}
    for step in range(3):
        model = doc.light_document_model
        assert {hf.header_footer_type for hf in model.sections[0].headers_footers} == expected
        for hf in model.sections[0].headers_footers:
            assert [p.text for p in hf.paragraphs] == ['OWNED STORY']
            assert hf.paragraphs[0]._children[0].font.bold
        if via_json:
            doc._document = ldm.Document.model_validate_json(model.model_dump_json())
        if step < 2:
            path = tmp_path / f'{step}.bin'
            doc.save(path, getattr(aw.SaveFormat, format_name))
            doc = aw.Document(path)
