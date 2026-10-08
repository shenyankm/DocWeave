"""Warnings, Unicode fallback and font-resource loading policy."""

import warnings
from pathlib import Path
from types import SimpleNamespace

from fpdf import FPDF
from fontTools.ttLib import TTFont
from fontTools import subset
from pypdf import PdfReader
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import (
    LdmPdfWriter,
    PdfFontSubstitutionWarning,
    PdfMissingGlyphWarning,
    PdfUnsupportedOptionWarning,
)
from aspose.words_foss.pdf_writer.font import register_fonts
from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_NAME
from aspose.words_foss.pdf_writer.diagnostics import warn_about_conversion

FONTS = Path(__file__).resolve().parents[1] / "aspose/words_foss/pdf_writer/fonts"


def plain_document(text="中文", **font):
    return ldm.Document(
        sections=[
            ldm.Section(
                body=ldm.Body(
                    children=[ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(**font))])]
                )
            )
        ]
    )


def test_plain_document_only_loads_regular_font():
    pdf = FPDF()
    register_fonts(pdf, plain_document())
    assert set(pdf.fonts) == {"documentsanssc"}


def test_text_box_styles_are_discovered():
    quote = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(style_name="Quote"), children=[ldm.Run(text="中文")]
    )
    document = ldm.Document(
        sections=[
            ldm.Section(
                body=ldm.Body(
                    children=[
                        ldm.Paragraph(
                            children=[
                                ldm.Shape(text_box={"paragraphs": [quote]}),
                                ldm.Shape(text_box={"paragraphs": None}),
                            ]
                        )
                    ]
                )
            )
        ]
    )
    pdf = FPDF()
    register_fonts(pdf, document)
    assert set(pdf.fonts) == {"documentsanssc", "documentsansscI"}


def test_font_compression_retains_full_source_coverage():
    coverage = None
    for path in FONTS.glob("*.woff"):
        font = TTFont(path)
        assert len(font.getGlyphOrder()) == 31036
        current = set(font.getBestCmap())
        assert all(ord(c) in current for c in "简体中文繁體中文ABC，。€")
        assert coverage is None or current == coverage
        coverage = current
    assert coverage


def test_default_pdf_options_do_not_warn(tmp_path):
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        LdmPdfWriter().write(plain_document(), tmp_path / "plain.pdf")
    assert not [w for w in caught if issubclass(w.category, PdfUnsupportedOptionWarning)]


@pytest.mark.parametrize(
    "option,value",
    [
        ("use_core_fonts", False),
        ("memory_optimization", True),
        ("font_embedding_mode", "embed_none"),
    ],
)
def test_explicit_unsupported_options_warn(tmp_path, option, value):
    options = aw.saving.PdfSaveOptions()
    setattr(options, option, value)
    with pytest.warns(PdfUnsupportedOptionWarning, match=option):
        LdmPdfWriter(options).write(plain_document(), tmp_path / "options.pdf")


def test_pdf_archival_compliance_is_not_silently_claimed(tmp_path):
    options = aw.saving.PdfSaveOptions()
    options.compliance = aw.saving.PdfCompliance.PDF_A1B
    with pytest.warns(PdfUnsupportedOptionWarning, match="conformance is not implemented"):
        LdmPdfWriter(options).write(plain_document(), tmp_path / "archival.pdf")


@pytest.mark.parametrize("codepoint", [0x1FAE0, 0xE123, 0x10FFFF])
def test_missing_glyph_includes_new_unassigned_and_private_use_characters(tmp_path, codepoint):
    with pytest.warns(PdfMissingGlyphWarning, match=f"U\\+{codepoint:04X}"):
        LdmPdfWriter().write(plain_document(chr(codepoint)), tmp_path / "missing.pdf")


def test_missing_list_label_warns(tmp_path):
    document = plain_document()
    paragraph = document.sections[0].body.children[0]
    paragraph.list_format = ldm.ListFormat(is_list_item=True)
    paragraph.list_label = ldm.ListLabel(label_string=chr(0x1FAE0))
    with pytest.warns(PdfMissingGlyphWarning, match="U\\+1FAE0"):
        LdmPdfWriter().write(document, tmp_path / "list.pdf")


@pytest.mark.parametrize("bold", [False, True])
@pytest.mark.parametrize("fallback", [False, True])
def test_glyph_diagnostics_do_not_scan_primary_font_cmap(bold, fallback):
    class LookupOnlyCmap(dict):
        def __iter__(self):
            raise AssertionError("Diagnostics must only look up text characters")

        def keys(self):
            raise AssertionError("Diagnostics must not scan the font cmap")

    coverage = LookupOnlyCmap({ord("A"): "A"})
    family = DEFAULT_FONT_NAME.lower()
    pdf = SimpleNamespace(fonts={
        family: SimpleNamespace(cmap=coverage),
        family + "B": SimpleNamespace(cmap=coverage),
        "fallback0": SimpleNamespace(cmap={0x1FAE0: "fallback"}),
    })
    document = plain_document("AA" + chr(0x1FAE0) + chr(0x10FFFF) + "\t\n",
                              name=DEFAULT_FONT_NAME, bold=bold)
    paragraph = document.sections[0].body.children[0]
    paragraph.list_format = ldm.ListFormat(is_list_item=True)
    paragraph.list_label = ldm.ListLabel(label_string=chr(0xE123) + chr(0x1FAE0))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        warn_about_conversion(pdf, document, aw.saving.PdfSaveOptions(),
                              ["Fallback0"] if fallback else [])
    missing = [w for w in caught if issubclass(w.category, PdfMissingGlyphWarning)]
    assert len(missing) == 1
    expected = "U+E123, U+10FFFF" if fallback else "U+E123, U+1FAE0, U+10FFFF"
    assert str(missing[0].message) == (
        f"PDF has {2 if fallback else 3} unsupported character(s): {expected}. Configure fallback_fonts."
    )
    substitutions = [w for w in caught if issubclass(w.category, PdfFontSubstitutionWarning)]
    assert len(substitutions) == int(bold and fallback)
    assert len(caught) == len(missing) + len(substitutions)


def test_source_font_substitution_warns(tmp_path):
    with pytest.warns(PdfFontSubstitutionWarning, match="SimSun"):
        LdmPdfWriter().write(plain_document(name="SimSun"), tmp_path / "substitution.pdf")


@pytest.mark.parametrize("bold", [False, True])
def test_fallback_font_fills_missing_character(tmp_path, bold):
    # A synthetic test font maps an otherwise missing codepoint to a known glyph.
    code = 0x1FAE0
    font = TTFont(FONTS / "DocumentSansSC-Regular.woff")
    glyph = font.getBestCmap()[ord("中")]
    for table in font["cmap"].tables:
        if table.format == 12:
            table.cmap[code] = glyph
    options = subset.Options()
    subsetter = subset.Subsetter(options=options)
    subsetter.populate(unicodes=[code, 32])
    subsetter.subset(font)
    fallback = tmp_path / "fallback.ttf"
    font.flavor = None
    font.save(fallback)
    save_options = aw.saving.PdfSaveOptions()
    save_options.fallback_fonts = [str(fallback)]
    output = tmp_path / "fallback.pdf"
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        LdmPdfWriter(save_options).write(plain_document("中文" + chr(code), bold=bold), output)
    assert not [w for w in caught if issubclass(w.category, PdfMissingGlyphWarning)]
    assert "中文" + chr(code) in PdfReader(output).pages[0].extract_text()
    substitutions = [w for w in caught if issubclass(w.category, PdfFontSubstitutionWarning)]
    assert bool(substitutions) == bold
