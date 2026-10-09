"""Generate owned paragraph-style imports with default-on italic or both toggles."""

import argparse
import itertools
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo


def destination(data):
    with ZipFile(BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["word/document.xml"] = parts["word/document.xml"].replace(b">IMPORT<", b">DESTINATION<")
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        for name, payload in parts.items():
            archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), payload)
    return stream.getvalue()


def inputs():
    helpers = runpy.run_path(str(Path(__file__).with_name("import_style_conflicts.py")))
    for bold, italic in ((False, True), (True, True)):
        for property_name in ("b", "i"):
            for source, child, target in itertools.product((None, False, True), repeat=3):
                key = "-".join(map(str, ("paragraph-default", bold, italic, property_name, source, child, target)))
                for phase, value in (("source", source), ("destination", target)):
                    styles = ('<w:docDefaults><w:rPrDefault><w:rPr>' +
                              f'<w:b w:val="{int(bold)}"/><w:i w:val="{int(italic)}"/><w:sz w:val="22"/>' +
                              '</w:rPr></w:rPrDefault></w:docDefaults>')
                    styles += helpers["style"]("Base", "paragraph", {property_name: value})
                    if phase == "source":
                        styles += helpers["style"]("Derived", "paragraph", {property_name: child}, "Base")
                    identifier = "Derived" if phase == "source" else "Base"
                    data = helpers["document"](styles, f'<w:pStyle w:val="{identifier}"/>', "")
                    if phase == "destination":
                        data = destination(data)
                    yield key, phase, data
            for child, character in itertools.product((None, False, True), (False, True)):
                key = "-".join(map(str, ("paragraph-context", bold, italic, property_name, child, character)))
                common = ('<w:docDefaults><w:rPrDefault><w:rPr>' +
                          f'<w:b w:val="{int(bold)}"/><w:i w:val="{int(italic)}"/><w:sz w:val="22"/>' +
                          '</w:rPr></w:rPrDefault></w:docDefaults>' +
                          helpers["style"]("Base", "paragraph", {property_name: True}) +
                          helpers["style"]("C", "character", {property_name: character}))
                for phase in ("source", "destination"):
                    styles = common + (helpers["style"]("Derived", "paragraph", {property_name: child}, "Base") if phase == "source" else "")
                    identifier = "Derived" if phase == "source" else "Base"
                    data = helpers["document"](styles, f'<w:pStyle w:val="{identifier}"/>', '<w:rStyle w:val="C"/>')
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
        assert count == 264
