"""Optional live competitors validate our output; no mocked engines or core dependencies."""

from html.parser import HTMLParser
from io import BytesIO
from zipfile import ZipFile
import warnings

from docx import Document
import pytest

from ._word_fixtures import word_document

import dockiln as dk
from .test_content_integrity import source_package
from .test_docx_corpus import corpus
from .test_reference_styles import reference

mammoth = pytest.importorskip('mammoth', reason='Install mammoth for independent ecosystem checks')
pypandoc = pytest.importorskip('pypandoc', reason='Install pypandoc and Pandoc for independent checks')
try:
    pypandoc.get_pandoc_version()
except OSError:
    pytest.skip('A runnable Pandoc executable is required', allow_module_level=True)


class HTMLText(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        self.parts.append(data)

    @classmethod
    def read(cls, html):
        parser = cls()
        parser.feed(html)
        return ' '.join(' '.join(parser.parts).split())


@pytest.mark.parametrize('case,tokens', [
    ('text_and_emphasis', ['中文报告', '普通', '粗体', '粗斜体', 'Literal *stars* _under_ [brackets]']),
    ('business_style', ['业务章节', '第一项', '第二项']),
    ('merged_table', ['表前', 'Header left', 'Header right', 'A | B', '中文表格', '表后']),
    ('links', ['链接 [说明]']),
])
def test_competitors_agree_on_common_docx_content(tmp_path, case, tokens):
    raw = corpus(case)
    source = tmp_path / 'source.docx'
    source.write_bytes(raw)
    own = dk.Document(BytesIO(raw)).get_text()
    other = mammoth.convert_to_html(BytesIO(raw), style_map="p[style-name='业务标题'] => h2:fresh").value
    pandoc = pypandoc.convert_file(str(source), 'html')
    assert all(token in own and token in HTMLText.read(other) and token in HTMLText.read(pandoc)
               for token in tokens)
    if case == 'links':
        assert 'https://example.com/a_(b)?x=1&amp;y=2' in other
        assert 'https://example.com/a_(b)?x=1&amp;y=2' in pandoc






