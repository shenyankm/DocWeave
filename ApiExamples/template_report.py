"""Compose docxtpl with this converter; install docxtpl separately.

Usage: python ApiExamples/template_report.py template.docx context.json report.docx report.pdf
Templates must be trusted: Jinja templates are code, not a secure upload format.
"""

import argparse
import json
from pathlib import Path

import aspose.words_foss as aw
from aspose.words_foss._io import atomic_output


def render_report(template_path, context, docx_path, pdf_path):
    from docxtpl import DocxTemplate
    from jinja2 import Environment, StrictUndefined

    if not isinstance(context, dict):
        raise ValueError("Template context must be a JSON object")
    if len({Path(path).resolve() for path in (template_path, docx_path, pdf_path)}) != 3:
        raise ValueError("Template, DOCX output and PDF output paths must be distinct")
    template = DocxTemplate(template_path)
    template.render(context, jinja_env=Environment(undefined=StrictUndefined), autoescape=True)
    with atomic_output(docx_path) as temporary:
        template.save(temporary)
    document = aw.Document(docx_path)
    document.save(pdf_path, aw.SaveFormat.PDF)
    return document.diagnostics


def test_template_report(tmp_path):
    import pytest
    pytest.importorskip("docxtpl")
    from docx import Document
    from pypdf import PdfReader

    template = tmp_path / "template.docx"
    source = Document()
    source.add_paragraph("客户：{{ company }}")
    source.save(template)
    docx, pdf = tmp_path / "report.docx", tmp_path / "report.pdf"
    render_report(template, {"company": "中文 <公司> & 合作方"}, docx, pdf)
    assert "中文 <公司> & 合作方" in aw.Document(docx).get_text()
    assert "中文" in "".join(page.extract_text() for page in PdfReader(pdf).pages)
    before = docx.read_bytes()
    for docx_output, pdf_output in [(template, pdf), (docx, docx)]:
        with pytest.raises(ValueError, match="distinct"):
            render_report(template, {"company": "test"}, docx_output, pdf_output)
    assert docx.read_bytes() == before
    raw_pdf = tmp_path / "report.bin"
    render_report(template, {"company": "test"}, docx, raw_pdf)
    assert raw_pdf.read_bytes().startswith(b"%PDF-")
    before = docx.read_bytes()
    from jinja2 import UndefinedError
    with pytest.raises(UndefinedError):
        render_report(template, {}, docx, pdf)
    assert docx.read_bytes() == before


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path)
    parser.add_argument("context", type=Path)
    parser.add_argument("docx", type=Path)
    parser.add_argument("pdf", type=Path)
    args = parser.parse_args()
    diagnostics = render_report(args.template, json.loads(args.context.read_text(encoding="utf-8")), args.docx, args.pdf)
    for diagnostic in diagnostics:
        print(f"{diagnostic.code}: {diagnostic.message}")


if __name__ == "__main__":
    main()
