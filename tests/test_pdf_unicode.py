"""DOCX-to-PDF regression: Unicode content and embedded font styles."""
from pathlib import Path

import pytest
from docx import Document
from pypdf import PdfReader

import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import PdfFontSubstitutionWarning
from aspose.words_foss.pdf_writer.text import safe_text


@pytest.mark.parametrize("text", ["中文與繁體 ABC 123", "“引号”——……￥€", ""])
def test_safe_text_preserves_unicode(text):
    assert safe_text(text) == text


def test_word_symbol_bullets():
    assert safe_text("\uf0b7 \uf0a7") == "• ▪"


def test_chinese_docx_to_pdf(tmp_path):
    source = Path(__file__).parent / "data" / "input" / "chinese_pdf.docx"
    output = tmp_path / "chinese.pdf"
    doc = aw.Document(source)
    assert Document(source).styles['Normal'].font.name == 'SimSun'
    assert "你好世界，简体中文与繁體中文。English ABC 123" in doc.get_text()
    with pytest.warns(PdfFontSubstitutionWarning, match="Sources: SimSun$"):
        doc.save(output, aw.SaveFormat.PDF)

    pdf = PdfReader(output)
    assert len(pdf.pages) == 2
    text = "\n".join(page.extract_text() for page in pdf.pages)
    expected = [
        "中文报告 Report 2026",
        "你好世界，简体中文与繁體中文。English ABC 123",
        "标点：“中文引号”——省略号……金额￥100，€50。",
        "普通文本 Regular", "加粗文本 Bold", "斜体文本 Italic", "粗斜体文本 BoldItalic",
        "引用内容：中文也能倾斜。", "print('你好，中文代码')", "中文列表项目",
        "居中中文标题", "右对齐中文内容", "后续中文内容", "下划线中文",
        "姓名 Name", "结果 Result", "张三", "通过 Passed", "第二页：分页中文测试。",
    ]
    for snippet in expected:
        assert snippet in text
    assert "?" not in text
    assert text.replace("\n", "").count("长段落中文换行验证。") == 30
    for page in pdf.pages:
        page_text = page.extract_text()
        assert "页眉：中文测试 Header" in page_text
        assert "页脚：中文测试 Footer" in page_text

    font_names = set()
    for page in pdf.pages:
        for ref in page["/Resources"]["/Font"].values():
            font = ref.get_object()
            assert font["/Subtype"] == "/Type0"
            assert "/ToUnicode" in font
            descendant = font["/DescendantFonts"][0].get_object()
            descriptor = descendant["/FontDescriptor"].get_object()
            assert len(descriptor["/FontFile2"].get_data()) > 0
            font_names.add(str(font["/BaseFont"]).split("+")[-1])
    assert font_names == {
        "DocumentSansSC", "DocumentSansSCBold", "DocumentSansSCOblique", "DocumentSansSCBoldOblique"
    }
