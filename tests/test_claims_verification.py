"""Independent acceptance checks for the enhanced-fork claims.

The original failing height checks remain unchanged as regression tests;
no discrepancy is hidden by xfail or a widened tolerance.
"""

from io import BytesIO
from pathlib import Path
import subprocess
import sys
import warnings
from zipfile import ZipFile, ZIP_DEFLATED

from fontTools.ttLib import TTFont
from fpdf import FPDF
import pymupdf
import pytest
from pypdf import PdfReader

import aspose.words_foss as aw
from aspose.words_foss import _io, light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning, PdfUnsupportedOptionWarning
from aspose.words_foss.pdf_writer.paragraph_renderer import ParagraphRenderer

ROOT = Path(__file__).resolve().parents[1]
YELLOW = "Color [A=255, R=255, G=255, B=0]"


def model(paragraphs, **page):
    return ldm.Document(
        sections=[ldm.Section(page_setup=ldm.PageSetup(**page), body=ldm.Body(children=paragraphs))]
    )


@pytest.mark.parametrize(
    "case", ["plain", "highlight", "first_indent", "left_indent", "code", "quote"]
)
def test_estimated_height_matches_actual_public_conversion(tmp_path, monkeypatch, case):
    text = "中文测量" * 19
    paragraph = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(
            first_line_indent=72 if case == "first_indent" else 0,
            left_indent=36 if case == "left_indent" else 0,
            style_name={"code": "Code", "quote": "Quote"}.get(case, ""),
        ),
        children=[
            ldm.Run(
                text=text,
                font=ldm.Font(size=14, highlight_color=YELLOW if case == "highlight" else ""),
            )
        ],
    )
    source = tmp_path / "height.docx"
    LdmDocxWriter().write(
        model(
            [paragraph],
            page_width=180,
            page_height=1440,
            left_margin=15,
            right_margin=15,
            top_margin=30,
            bottom_margin=30,
        ),
        source,
    )
    observations = []
    original = ParagraphRenderer.render_paragraph

    def record(self, pdf, para):
        writer = self._writer
        estimate = writer._estimate_paragraph_height(
            para, writer._page_width - writer._page_margin_left - writer._page_margin_right
        )
        y, page = pdf.y, pdf.page
        original(self, pdf, para)
        assert pdf.page == page, "Fixture must stay on one page for a valid measurement"
        observations.append((estimate, pdf.y - y))

    monkeypatch.setattr(ParagraphRenderer, "render_paragraph", record)
    output = tmp_path / "height.pdf"
    aw.Document(source).save(output)
    with pymupdf.open(output) as pdf:
        assert "".join(page.get_text() for page in pdf).replace("\n", "") == text
    assert len(observations) == 1
    estimate, actual = observations[0]
    assert estimate == pytest.approx(
        actual, abs=0.05
    ), f"{case}: estimate={estimate:.3f} mm, actual={actual:.3f} mm"


def test_keep_together_does_not_split_indented_paragraph(tmp_path):
    prefix = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(space_after=40),
        children=[ldm.Run(text="前置占位", font=ldm.Font(size=14))],
    )
    text = "中文测量" * 19
    target = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(first_line_indent=72, keep_together=True),
        children=[ldm.Run(text=text, font=ldm.Font(size=14))],
    )
    source = tmp_path / "keep.docx"
    LdmDocxWriter().write(
        model(
            [prefix, target],
            page_width=180,
            page_height=100 * 72 / 25.4,
            left_margin=15,
            right_margin=15,
            top_margin=30,
            bottom_margin=30,
        ),
        source,
    )
    output = tmp_path / "keep.pdf"
    aw.Document(source).save(output)
    with pymupdf.open(output) as pdf:
        texts = [page.get_text().replace("\n", "") for page in pdf]
        assert text in "".join(texts)
        assert any(
            text in page_text for page_text in texts
        ), "keep_together paragraph split across pages"


def test_highlighted_links_survive_real_page_breaks(tmp_path):
    text = "中文链接换行" * 600
    paragraph = ldm.Paragraph(
        children=[
            ldm.Run(
                text=f"[{text}](https://example.com/verify)",
                font=ldm.Font(
                    size=13, bold=True, italic=True, underline=True, highlight_color=YELLOW
                ),
            )
        ]
    )
    source = tmp_path / "links.docx"
    LdmDocxWriter().write(model([paragraph]), source)
    output = tmp_path / "links.pdf"
    aw.Document(source).save(output)
    with pymupdf.open(output) as pdf:
        assert len(pdf) >= 2
        assert "".join(page.get_text() for page in pdf).replace("\n", "") == text
        for page in pdf:
            links = page.get_links()
            assert len(links) > 2
            assert all(link["uri"] == "https://example.com/verify" for link in links)
            fills = [d for d in page.get_drawings() if d.get("fill") == (1, 1, 0)]
            assert len(fills) > 2
            for span in [
                s
                for block in page.get_text("dict")["blocks"]
                for line in block.get("lines", [])
                for s in line["spans"]
            ]:
                assert span["flags"] & 16, "bold disappeared at a page break"
                assert span["flags"] & 2, "italic disappeared at a page break"
                x0, y0, x1, y1 = span["bbox"]
                assert 0 <= x0 < x1 <= page.rect.width
                assert 0 <= y0 < y1 <= page.rect.height
            assert page.get_pixmap().samples


@pytest.mark.parametrize(
    "option",
    [
        "text_compression",
        "embed_full_fonts",
        "use_core_fonts",
        "font_embedding_mode",
        "page_mode",
        "color_mode",
        "preserve_form_fields",
        "memory_optimization",
    ],
)
def test_all_eight_explicit_unsupported_options_warn_through_document(tmp_path, option):
    source = tmp_path / "source.txt"
    source.write_text("中文诊断", encoding="utf-8")
    options = aw.saving.PdfSaveOptions()
    setattr(options, option, getattr(options, option))
    with pytest.warns(PdfUnsupportedOptionWarning, match=option):
        aw.Document(source).save(tmp_path / "options.pdf", options)


def test_actual_default_input_limit(tmp_path):
    assert _io.MAX_INPUT_BYTES == 64 * 1024 * 1024
    source = tmp_path / "oversized.txt"
    with source.open("wb") as stream:
        stream.truncate(_io.MAX_INPUT_BYTES + 1)
    with pytest.raises(ValueError, match="Input exceeds 67108864"):
        aw.Document(source)


@pytest.mark.parametrize("limit", ["part", "expanded"])
def test_actual_docx_zip_limits_without_reducing_constants(limit):
    buffer = BytesIO()
    chunk = b"x" * (2 * 1024 * 1024)
    with ZipFile(buffer, "w", ZIP_DEFLATED, compresslevel=1) as archive:
        if limit == "part":
            with archive.open("word/document.xml", "w") as part:
                for _ in range(32):
                    part.write(chunk)
                part.write(b"x")
        else:
            archive.writestr("word/document.xml", "<document/>")
            for number in range(4):
                with archive.open(f"extra{number}.bin", "w") as part:
                    for _ in range(32):
                        part.write(chunk)
    assert len(buffer.getvalue()) < _io.MAX_INPUT_BYTES
    with pytest.raises(ValueError, match="part exceeds" if limit == "part" else "expanded size"):
        aw.Document(BytesIO(buffer.getvalue()))


def test_real_fallback_font_is_embedded_with_its_actual_outline(tmp_path):
    candidates = [
        Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    fallback = next((path for path in candidates if path.is_file()), None)
    if fallback is None:
        pytest.skip("No suitable real external font installed")
    code = 0x05D0  # Hebrew alef: absent from the bundled SC font, present in Arial/DejaVu.
    original = TTFont(fallback)
    bundled = TTFont(ROOT / "aspose/words_foss/pdf_writer/fonts/DocumentSansSC-Regular.woff")
    assert code not in bundled.getBestCmap() and code in original.getBestCmap()
    source = tmp_path / "fallback.txt"
    source.write_text("中文" + chr(code), encoding="utf-8")
    options = aw.saving.PdfSaveOptions()
    options.fallback_fonts = [str(fallback)]
    output = tmp_path / "fallback.pdf"
    with warnings.catch_warnings():
        warnings.simplefilter("error", PdfMissingGlyphWarning)
        aw.Document(source).save(output, options)
    embedded = []
    for ref in PdfReader(output).pages[0]["/Resources"]["/Font"].values():
        descriptor = ref.get_object()["/DescendantFonts"][0].get_object()["/FontDescriptor"]
        embedded.append(TTFont(BytesIO(descriptor["/FontFile2"].get_object().get_data())))
    font = next(font for font in embedded if code in font.getBestCmap())
    source_glyph = original["glyf"][original.getBestCmap()[code]]
    actual_glyph = font["glyf"][font.getBestCmap()[code]]
    assert actual_glyph.getCoordinates(font["glyf"]) == source_glyph.getCoordinates(
        original["glyf"]
    )
    with pymupdf.open(output) as pdf:
        chars = [
            char["c"]
            for block in pdf[0].get_text("rawdict")["blocks"]
            for line in block.get("lines", [])
            for span in line["spans"]
            for char in span["chars"]
        ]
        assert "中文" + chr(code) == "".join(chars)
        assert pdf[0].get_pixmap().samples


@pytest.mark.parametrize("suffix", ["pdf", "docx", "txt", "md"])
def test_actual_partial_write_preserves_original(tmp_path, monkeypatch, suffix):
    source = tmp_path / "source.txt"
    source.write_text("中文输出测试", encoding="utf-8")
    output = tmp_path / f"output.{suffix}"
    output.write_bytes(b"original file must survive")
    if suffix == "docx":
        original = ZipFile.writestr

        def write_then_fail(self, *args, **kwargs):
            original(self, *args, **kwargs)
            raise OSError("injected failure after writing real ZIP data")

        monkeypatch.setattr(ZipFile, "writestr", write_then_fail)
    else:

        def write_then_fail(self, data, *args, **kwargs):
            with self.open("wb") as stream:
                stream.write(
                    data[:10] if isinstance(data, (bytes, bytearray)) else data[:10].encode()
                )
            raise OSError("injected failure after writing real partial data")

        monkeypatch.setattr(
            Path, "write_bytes" if suffix == "pdf" else "write_text", write_then_fail
        )
    with pytest.raises(OSError, match="injected failure"):
        aw.Document(source).save(output)
    assert output.read_bytes() == b"original file must survive"
    assert sorted(path.name for path in tmp_path.iterdir()) == [output.name, source.name]


@pytest.mark.parametrize("style", ["Regular", "Bold", "Oblique", "BoldOblique"])
def test_woff_really_preserves_preoptimization_font_tables(style):
    path = f"aspose/words_foss/pdf_writer/fonts/DocumentSansSC-{style}"
    result = subprocess.run(
        ["git", "show", "7f6a745:" + path + ".ttf"], cwd=ROOT, capture_output=True
    )
    if result.returncode:
        pytest.skip("Original TTF revision not present in this checkout")
    original = TTFont(BytesIO(result.stdout))
    current = TTFont(ROOT / (path + ".woff"))
    assert current.getGlyphOrder() == original.getGlyphOrder()
    assert current.getBestCmap() == original.getBestCmap()
    assert current["hmtx"].metrics == original["hmtx"].metrics
    for tag in ("glyf", "cmap", "hmtx", "name", "OS/2", "loca"):
        assert current.reader[tag] == original.reader[tag], f"Font table changed: {tag}"
    assert (ROOT / (path + ".woff")).stat().st_size < len(result.stdout)


def test_real_libreoffice_conversion_if_installed(tmp_path):
    import shutil
    from aspose.words_foss.libreoffice import convert_to_pdf

    if not (
        shutil.which("soffice")
        or shutil.which("libreoffice")
        or Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").is_file()
    ):
        pytest.skip("LibreOffice is not installed; real rendering cannot be verified")
    output = tmp_path / "native.pdf"
    convert_to_pdf(ROOT / "tests/data/input/chinese_pdf.docx", output)
    with pymupdf.open(output) as pdf:
        assert "中文" in "".join(page.get_text() for page in pdf)
        assert all(page.get_pixmap().samples for page in pdf)

    # A standard text RTF cannot use the inherited DOC/OLE reader; native input can.
    text = "中文真实原文件通道"
    source = tmp_path / "original.rtf"
    source.write_text("{\\rtf1\\ansi\\uc1 " + "".join(f"\\u{ord(c)}?" for c in text) + "}",
                      encoding="ascii")
    convert_to_pdf(source, output)
    with pymupdf.open(output) as pdf:
        assert text in "".join(page.get_text() for page in pdf).replace("\n", "")

    # Inject a real footnote into the original OOXML, not the lossy LDM.
    source = tmp_path / "footnote.docx"
    LdmDocxWriter().write(model([ldm.Paragraph(children=[ldm.Run(text="中文正文")])]), source)
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["word/document.xml"] = parts["word/document.xml"].replace(
        b"</w:p>", b'<w:r><w:footnoteReference w:id="2"/></w:r></w:p>', 1)
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
        b"</Relationships>", b'<Relationship Id="rFootnote" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
        b'Target="footnotes.xml"/></Relationships>')
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", b'<Override PartName="/word/footnotes.xml" '
        b'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
    note = "原始脚注独有内容"
    parts["word/footnotes.xml"] = (
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:footnote w:id="2"><w:p><w:r><w:t>' + note +
        '</w:t></w:r></w:p></w:footnote></w:footnotes>').encode()
    with ZipFile(source, "w", ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    original = source.read_bytes()
    with pytest.warns(aw.loading.DocumentLoadWarning, match="footnoteReference"):
        lossy = aw.Document(source)
    assert note not in lossy.get_text()
    convert_to_pdf(source, output)
    assert source.read_bytes() == original
    with pymupdf.open(output) as pdf:
        assert note in "".join(page.get_text() for page in pdf).replace("\n", "")
        assert all(page.get_pixmap().samples for page in pdf)


def test_public_save_loads_only_regular_font_for_plain_text(tmp_path, monkeypatch):
    calls = []
    original = FPDF.add_font

    def record(self, family=None, style="", fname=None, **kwargs):
        calls.append((family, style, Path(fname).suffix))
        return original(self, family=family, style=style, fname=fname, **kwargs)

    monkeypatch.setattr(FPDF, "add_font", record)
    source = tmp_path / "source.txt"
    source.write_text("中文普通正文", encoding="utf-8")
    aw.Document(source).save(tmp_path / "plain.pdf")
    assert calls == [("DocumentSansSC", "", ".woff")]


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX resource limits")
def test_worker_resource_limits_are_actually_set():
    code = (
        "import resource,json; from aspose.words_foss.convert import _set_limits; "
        "_set_limits(1024,60); print(json.dumps([resource.getrlimit(resource.RLIMIT_CPU),"
        "resource.getrlimit(resource.RLIMIT_FSIZE)]))"
    )
    import json

    result = subprocess.run(
        [sys.executable, "-c", code], capture_output=True, text=True, timeout=10
    )
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout) == [[61, 61], [256 * 1024 * 1024] * 2]
