"""Run outside the checkout to verify wheel imports, typing marker and bundled fonts."""

from importlib.metadata import distribution
from pathlib import Path

import aspose.words_foss as aw


def main():
    dist = distribution("aspose-words-foss-enhanced")
    module = Path(aw.__file__).resolve()
    assert module == Path(dist.locate_file("aspose/words_foss/__init__.py")).resolve()
    assert "site-packages" in module.parts, f"Imported checkout instead of wheel: {module}"
    assert aw.__version__ == dist.version
    assert module.with_name("py.typed").is_file()
    fonts = module.parent / "pdf_writer" / "fonts"
    for style in ("Regular", "Bold", "Oblique", "BoldOblique"):
        assert (fonts / f"DocumentSansSC-{style}.woff").stat().st_size > 0
    assert (fonts / "OFL.txt").is_file()
    print(f"Verified {dist.metadata['Name']} {dist.version}: {module}")


if __name__ == "__main__":
    main()
