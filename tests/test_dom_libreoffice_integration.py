"""Optional native renderer smoke test, not a Word visual-fidelity assertion."""

from io import BytesIO
from pathlib import Path
import shutil

from docx import Document
from PIL import Image
from pypdf import PdfReader
import pytest

import aspose.words_foss as aw
from aspose.words_foss.libreoffice import convert_to_pdf


AVAILABLE = bool(shutil.which("soffice") or shutil.which("libreoffice") or
                 Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").is_file())


@pytest.mark.skipif(not AVAILABLE, reason="LibreOffice is not installed")
def test_native_renderer_accepts_dom_picture_link_and_horizontal_merge(tmp_path):
    original = Document()
    original.add_paragraph("DOM 中文 smoke test")
    table = original.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "left cell"
    table.cell(0, 1).text = "right cell"
    raw = BytesIO()
    original.save(raw)
    doc = aw.DocxDocument(BytesIO(raw.getvalue()))
    paragraph = doc.body.paragraphs[0]
    paragraph.add_hyperlink("链接 [说明]", "https://example.com/a_(b)?x=1&y=2")
    image = BytesIO()
    Image.new("RGB", (100, 50), "navy").save(image, format="PNG")
    paragraph.add_picture(image.getvalue(), width=72, alternative_text="中文图片")
    doc.body.tables[0].merge_cells(0, 0, 2)
    source, output = tmp_path / "edited.docx", tmp_path / "native.pdf"
    doc.save(source)
    convert_to_pdf(source, output)
    rendered = PdfReader(output)
    text = "".join(page.extract_text() for page in rendered.pages)
    assert all(token in text for token in ("DOM", "中文", "链接", "说明", "left cell", "right cell"))
    assert any(page.images for page in rendered.pages)
    links = [item.get_object()["/A"].get("/URI") for page in rendered.pages
             for item in page.get("/Annots", []) if "/A" in item.get_object()]
    assert "https://example.com/a_(b)?x=1&y=2" in links
