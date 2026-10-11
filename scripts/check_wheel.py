"""Run outside the checkout to verify wheel imports, typing marker and bundled fonts."""

from importlib.metadata import distribution
from pathlib import Path

import dockiln as dk


def main():
    dist = distribution("dockiln")
    module = Path(dk.__file__).resolve()
    assert module == Path(dist.locate_file("dockiln/__init__.py")).resolve()
    assert "site-packages" in module.parts, f"Imported checkout instead of wheel: {module}"
    assert dk.__version__ == dist.version
    assert module.with_name("py.typed").is_file()
    fonts = module.parent / "writers" / "pdf" / "fonts"
    for style in ("Regular", "Bold", "Oblique", "BoldOblique"):
        assert (fonts / f"DocumentSansSC-{style}.woff").stat().st_size > 0
    assert (fonts / "OFL.txt").is_file()
    print(f"Verified {dist.metadata['Name']} {dist.version}: {module}")


if __name__ == "__main__":
    main()
