"""Compression changes must preserve every decoded table of this font revision."""

import hashlib
from pathlib import Path

import pytest
from fontTools.ttLib.sfnt import SFNTReader

from aspose.words_foss.pdf_writer import font


@pytest.mark.parametrize(
    "face,expected",
    [
        ("Regular", "33e96821f7cf58d750f5d0f5eebe368261d04bcd5e27d20ed45d76910ed99bb3"),
        ("Bold", "785bdc701b3fd1423b60e90297b61a9f133adf1070ce2c7618625ec1413b561b"),
        ("Oblique", "d2719c789a0ccba5e81fb03eebb413e29c747c26c2f9372787a0e682006a51ef"),
        (
            "BoldOblique",
            "5f1e3eec70ee385b0234e3873cea0c7844fe955831eb92d6857e0b460b7b826e",
        ),
    ],
)
def test_bundled_font_decoded_tables(face, expected):
    path = Path(font.__file__).parent / "fonts" / f"DocumentSansSC-{face}.woff"
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        reader = SFNTReader(stream)
        for tag in sorted(reader.keys()):
            data = reader[tag]
            digest.update(str(tag).encode("ascii"))
            digest.update(len(data).to_bytes(4, "big"))
            digest.update(data)
    # Update these baselines only when intentionally changing the font revision.
    assert digest.hexdigest() == expected


@pytest.mark.parametrize(
    "face,expected",
    [
        ("Regular", "740891cfa7ec1eea262de44c8afc8969d0c686cc1eff69ddd0779d3901cc3985"),
        ("Bold", "804e11eb6c22934bb3b9ce3520e0456f4ff3519687edde5e92314cd797323638"),
        ("Oblique", "8a232959068a25e1a0dd810ccef8a03f985dfab46d63430766c05ecc793b33b5"),
        (
            "BoldOblique",
            "e3208e5d0485d889a7237f1c442cea41e0b4834353078d2703f83600e642a0a7",
        ),
    ],
)
def test_font_metrics_outlines_and_layout_tables_are_unchanged(face, expected):
    path = Path(font.__file__).parent / "fonts" / f"DocumentSansSC-{face}.woff"
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        reader = SFNTReader(stream)
        for tag in sorted(reader.keys()):
            if str(tag) == "post":
                continue
            data = reader[tag]
            if str(tag) == "head":
                data = data[:8] + bytes(4) + data[12:]
            digest.update(str(tag).encode("ascii"))
            digest.update(len(data).to_bytes(4, "big"))
            digest.update(data)
    assert digest.hexdigest() == expected
