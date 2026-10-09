"""Generate owned character-style imports with default-on italic or both toggles."""

import argparse
import itertools
import runpy
from pathlib import Path
from zipfile import ZipFile, ZipInfo


def inputs():
    destination = runpy.run_path(str(Path(__file__).with_name("paragraph_style_defaults.py")))["destination"]
    helpers = runpy.run_path(str(Path(__file__).with_name("import_style_conflicts.py")))
    for bold, italic in ((False, True), (True, True)):
        for property_name in ("b", "i"):
            for source, child, target in itertools.product((None, False, True), repeat=3):
                key = "-".join(map(str, ("character-default", bold, italic, property_name, source, child, target)))
                for phase, value in (("source", source), ("destination", target)):
                    styles = ('<w:docDefaults><w:rPrDefault><w:rPr>' +
                              f'<w:b w:val="{int(bold)}"/><w:i w:val="{int(italic)}"/><w:sz w:val="22"/>' +
                              '</w:rPr></w:rPrDefault></w:docDefaults>')
                    styles += helpers["style"]("Base", "character", {property_name: value})
                    if phase == "source":
                        styles += helpers["style"]("Derived", "character", {property_name: child}, "Base")
                    identifier = "Derived" if phase == "source" else "Base"
                    styles += helpers["style"]("P", "paragraph", {"b": False, "i": False})
                    data = helpers["document"](styles, '<w:pStyle w:val="P"/>', f'<w:rStyle w:val="{identifier}"/>')
                    if phase == "destination":
                        data = destination(data)
                    yield key, phase, data


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    args = parser.parse_args()
    with ZipFile(args.corpus, "w") as archive:
        count = 0
        for key, phase, data in inputs():
            archive.writestr(ZipInfo(key + "/" + phase + ".docx", (2020, 1, 1, 0, 0, 0)), data)
            count += 1
        assert count == 216
