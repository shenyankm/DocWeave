"""Related comment parts must not disappear without a content-loss diagnostic."""
from io import BytesIO
import warnings
from zipfile import ZipFile

from docx import Document as NativeDocument
import pytest

import aspose.words_foss as aw
from aspose.words_foss._flat_opc import encode


def commented_document(reference):
    document = NativeDocument()
    run = document.add_paragraph('PRIVATE BODY').runs[0]
    document.add_comment(run, text='PRIVATE COMMENT', author='PRIVATE AUTHOR')
    if not reference:
        for name in ('commentReference', 'commentRangeStart', 'commentRangeEnd'):
            for node in document.element.xpath('.//w:' + name):
                node.getparent().remove(node)
    output = BytesIO()
    document.save(output)
    return output.getvalue()


@pytest.mark.parametrize('reference', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_comment_part_loss_is_recorded_even_when_warnings_are_ignored(reference, flat):
    raw = commented_document(reference)
    source = encode(raw) if flat else raw
    with warnings.catch_warnings():
        warnings.simplefilter('ignore', aw.ContentLossWarning)
        document = aw.Document(BytesIO(source))
        output = document.to_bytes(aw.SaveFormat.DOCX)
    observed = [item for item in document.diagnostics if item.code == 'load.comments_omitted']
    assert len(observed) == 1
    assert all(text not in repr(observed) for text in ('PRIVATE BODY', 'PRIVATE COMMENT', 'PRIVATE AUTHOR'))
    with ZipFile(BytesIO(output)) as package:
        assert 'word/comments.xml' not in package.namelist()
    assert 'PRIVATE BODY' in document.get_text()


@pytest.mark.parametrize('reference', [False, True])
def test_strict_conversion_rejects_comment_loss_and_keeps_editable_package(reference):
    raw = commented_document(reference)
    editable = aw.DocxDocument(BytesIO(raw))
    with warnings.catch_warnings():
        warnings.simplefilter('error', aw.ContentLossWarning)
        with pytest.raises(aw.ContentLossWarning, match='Comment parts'):
            editable.to_light_document()
    assert editable.to_bytes() == raw


def test_document_without_comments_has_no_comment_loss_diagnostic():
    document = NativeDocument()
    document.add_paragraph('OWNED BODY')
    output = BytesIO()
    document.save(output)
    loaded = aw.Document(BytesIO(output.getvalue()))
    assert not any(item.code == 'load.comments_omitted' for item in loaded.diagnostics)
