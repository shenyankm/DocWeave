"""Unknown enum requests must fail at the public option boundary."""

import pytest

from dockiln import saving

FIELDS = [
    (saving.PdfSaveOptions, "compliance", saving.PdfCompliance),
    (saving.PdfSaveOptions, "image_compression", saving.PdfImageCompression),
    (saving.PdfSaveOptions, "text_compression", saving.PdfTextCompression),
    (saving.PdfSaveOptions, "font_embedding_mode", saving.PdfFontEmbeddingMode),
    (saving.PdfSaveOptions, "page_mode", saving.PdfPageMode),
    (saving.PdfSaveOptions, "color_mode", saving.ColorMode),
    (saving.PdfSaveOptions, "zoom_behavior", saving.PdfZoomBehavior),
    (saving.OoxmlSaveOptions, "compliance", saving.OoxmlCompliance),
    (saving.OoxmlSaveOptions, "compression_level", saving.CompressionLevel),
    (saving.OoxmlSaveOptions, "zip_64_mode", saving.Zip64Mode),
]


@pytest.mark.parametrize("options_type,name,enum", FIELDS)
@pytest.mark.parametrize("invalid", ["not_a_mode", 999, None, [], 1.25])
def test_bad_enum_assignment_rejected_without_changing_option(options_type, name, enum, invalid):
    options = options_type()
    previous = getattr(options, name)
    with pytest.raises(ValueError, match=enum.__name__):
        setattr(options, name, invalid)
    assert getattr(options, name) is previous


@pytest.mark.parametrize("options_type,name,enum", FIELDS)
def test_all_known_enum_members_and_legacy_values_still_normalize(options_type, name, enum):
    options = options_type()
    for member in enum:
        for value in (member, int(member), member.name, member.name.lower()):
            setattr(options, name, value)
            assert getattr(options, name) is member


