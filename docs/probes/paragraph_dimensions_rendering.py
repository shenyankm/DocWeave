"""Native/current point edits in fixed-line paragraphs after a trial-page anchor."""

import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo

HELPER = runpy.run_path(str(Path(__file__).with_name("pagination_rendering.py")))
PROPERTIES = ("left_indent", "right_indent", "first_line_indent", "space_before", "space_after")


def inputs():
    _, _, raw = next(HELPER["inputs"]())
    with ZipFile(BytesIO(raw)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    for prop in PROPERTIES:
        stream = BytesIO()
        with ZipFile(stream, "w") as archive:
            for name, data in parts.items():
                if name == "word/styles.xml" and prop == "right_indent":
                    data = data.replace(b'<w:name w:val="Target"/>', b'<w:name w:val="Target"/><w:pPr><w:jc w:val="right"/></w:pPr>')
                if name == "word/document.xml":
                    data = data.replace(b'<w:sectPr>', b'<w:p><w:r><w:t>FOLLOW00</w:t></w:r></w:p><w:sectPr>')
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), data)
        yield prop + ".docx", prop, stream.getvalue()


if __name__ == "__main__":
    HELPER["main"](input_generator=inputs, values=(0.0, 12.375))
