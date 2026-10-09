"""Compression changes must preserve every decoded table of this font revision."""

import hashlib
from pathlib import Path

from fontTools.ttLib.sfnt import SFNTReader
import pytest

from aspose.words_foss.pdf_writer import font


@pytest.mark.parametrize(
    "face,expected",
    [
        ("Regular", "4efe0ec35340e13e302ad75788f594c6ea97cc098ef761a8cdece22c6c46566d"),
        ("Bold", "a92b0c1fd56025b11fea6f56174f9c8995a5b6d50b8835e92a92252fffca0b4f"),
        ("Oblique", "fa2fc5508f797660edea5a258f8591a57fcac8b7e5478b55c907d584393b1a7a"),
        ("BoldOblique", "651e95e7f2bf63188ffab24cb234a3dfbd959880cc8e1697b57b9b22983bafb1"),
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
