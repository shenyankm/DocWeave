"""Owned start/end indents: alias priority, inherited layers and bidi getter observations."""

import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo

VALUES = (0.0, 12.375)
PATTERNS = (
    (None, None, None, "logical"),
    (None, None, None, "both"),
    (None, "logical", "physical", None),
    (None, "physical", "logical", None),
    ("logical", None, "sibling", None),
    ("logical", "logical", "physical", "logical"),
    (None, None, None, "both-reversed"),
)


def inputs():
    helper = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))
    for prop, physical, logical in (("left_indent", "left", "start"), ("right_indent", "right", "end"), ("first_line_indent", "firstLine", "hanging")):
        for bidi in (0, 1):
            for index, layers in enumerate(PATTERNS):
                def fragment(kind, value, physical=physical, logical=logical):
                    if kind is None:
                        return ""
                    other = "end" if logical == "start" else "start"
                    if kind == "sibling":
                        return f'<w:ind w:{other}="80"/>'
                    attrs = f'w:{logical if kind != "physical" else physical}="{value}"'
                    if kind == "both":
                        attrs += f' w:{physical}="200"'
                    elif kind == "both-reversed":
                        attrs = f'w:{physical}="200" w:{logical}="{value}"'
                    return f'<w:ind {attrs}/>'

                default, base, child, direct = layers
                styles = '<w:docDefaults><w:pPrDefault><w:pPr>' + fragment(default, 240)
                styles += '</w:pPr></w:pPrDefault></w:docDefaults>'
                for name, kind, value, parent in (("Base", base, 360, ""), ("Derived", child, 480, '<w:basedOn w:val="Base"/>')):
                    styles += f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>{parent}<w:pPr>{fragment(kind, value)}</w:pPr></w:style>'
                ppr = f'<w:pStyle w:val="Derived"/><w:bidi w:val="{bidi}"/>' + fragment(direct, 600)
                raw = helper["document"](styles, ppr, "")
                with ZipFile(BytesIO(raw)) as package:
                    parts = {name: package.read(name) for name in package.namelist()}
                parts["word/document.xml"] = parts["word/document.xml"].replace(b"<w:body>", b"<w:body><w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>", 1)
                output = BytesIO()
                with ZipFile(output, "w") as package:
                    for name, data in parts.items():
                        package.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), data)
                yield f"{prop}/{bidi}-{index}.docx", prop, output.getvalue()


if __name__ == "__main__":
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_dimensions.py")))
    probe["main"](inputs, VALUES)
