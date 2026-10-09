"""Own used/unused root style contexts in body, header and footer stories."""

import argparse
import itertools
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def inputs():
    helpers = runpy.run_path(str(Path(__file__).with_name("import_style_conflicts.py")))
    destination = runpy.run_path(str(Path(__file__).with_name("paragraph_style_defaults.py")))["destination"]
    for default, paragraph, character, used in itertools.product((False, True), (False, True), (None, False, True), (False, True)):
        if character is None and used:
            continue
        styles = (f'<w:docDefaults><w:rPrDefault><w:rPr><w:b w:val="{int(default)}"/><w:sz w:val="22"/>'
                  '</w:rPr></w:rPrDefault></w:docDefaults>' + helpers["style"]("Base", "paragraph", {"b": paragraph}) +
                  helpers["style"]("P", "paragraph", {"b": default}))
        if character is not None:
            styles += helpers["style"]("C", "character", {"b": character})
        data = helpers["document"](styles, '<w:pStyle w:val="Base"/>', '<w:rStyle w:val="C"/>' if used else "")
        key = "-".join(map(str, ("root", default, paragraph, character, used)))
        for phase in ("source", "destination"):
            yield key, phase, destination(data) if phase == "destination" else data
    for kind in ("header", "footer"):
        tag = "hdr" if kind == "header" else "ftr"
        for default, paragraph, character in itertools.product((False, True), repeat=3):
            styles = (f'<w:docDefaults><w:rPrDefault><w:rPr><w:b w:val="{int(default)}"/><w:sz w:val="22"/>'
                      '</w:rPr></w:rPrDefault></w:docDefaults>' + helpers["style"]("Base", "paragraph", {"b": paragraph}) +
                      helpers["style"]("C", "character", {"b": character}))
            source = helpers["document"](styles, '<w:pStyle w:val="Base"/>', "")
            with ZipFile(BytesIO(destination(source))) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["word/document.xml"] = parts["word/document.xml"].replace(b">DESTINATION<", b">BODY<").replace(
                b"<w:document ", f'<w:document xmlns:r="{R}" '.encode()).replace(
                b"<w:sectPr/>", f'<w:sectPr><w:{kind}Reference w:type="default" r:id="{kind}"/></w:sectPr>'.encode())
            parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
                b"</Relationships>", f'<Relationship Id="{kind}" Type="{R}/{kind}" Target="{kind}1.xml"/></Relationships>'.encode())
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(b"</Types>", (
                f'<Override PartName="/word/{kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/></Types>').encode())
            parts[f"word/{kind}1.xml"] = (f'<w:{tag} xmlns:w="{W}"><w:p><w:pPr><w:pStyle w:val="Base"/></w:pPr>'
                                         '<w:r><w:rPr><w:rStyle w:val="C"/></w:rPr><w:t>DESTINATION</w:t></w:r></w:p>'
                                         f'</w:{tag}>').encode()
            output = BytesIO()
            with ZipFile(output, "w") as archive:
                for name, payload in parts.items():
                    archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), payload)
            key = "-".join(map(str, (kind, default, paragraph, character)))
            yield key, "source", source
            yield key, "destination", output.getvalue()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    args = parser.parse_args()
    with ZipFile(args.corpus, "w") as archive:
        count = 0
        for key, phase, data in inputs():
            archive.writestr(ZipInfo(key + "/" + phase + ".docx", (2020, 1, 1, 0, 0, 0)), data)
            count += 1
        assert count == 72
