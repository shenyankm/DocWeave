"""Owned fixed-line pagination inputs with a new-page anchor after trial content."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo

CASES = {"page_break_before": 2, "keep_together": 8, "keep_with_next": 10,
         "widow_control_orphan": 10, "widow_control_widow": 8}


def prepare_font(folder):
    from fontTools.ttLib import TTFont

    source = Path(__file__).parents[2] / "aspose/words_foss/pdf_writer/fonts/DocumentSansSC-Regular.woff"
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    with TTFont(source, recalcTimestamp=False) as font:
        font.flavor = None
        font.save(folder / "DocumentSansSC-Regular.ttf")


def pdf_snapshot(data):
    import pymupdf

    with pymupdf.open(stream=data, filetype="pdf") as document:
        lines = {}
        for index, page in enumerate(document):
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    for span in line["spans"]:
                        label = span["text"]
                        if label.startswith(("BEFORE", "ANCHOR", "FILL", "TARGET", "FOLLOW")):
                            assert label not in lines
                            lines[label] = {"page": index + 1, "origin": list(span["origin"]),
                                            "advance": span["bbox"][2] - span["bbox"][0], "size": span["size"]}
        return {"page_sizes": [[page.rect.width, page.rect.height] for page in document], "lines": lines}


def black_ink(page):
    import pymupdf
    from PIL import Image, ImageChops

    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    red, green, blue = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples).split()
    # Owned labels are black; this excludes the trial's red footer and pale logo.
    return ImageChops.lighter(ImageChops.lighter(red, green), blue).point([255 if value < 128 else 0 for value in range(256)])


def ink_difference(reference, candidate):
    import pymupdf
    from PIL import ImageChops

    with pymupdf.open(stream=reference, filetype="pdf") as native, pymupdf.open(stream=candidate, filetype="pdf") as current:
        assert len(native) == len(current)
        result = []
        for index in range(1, len(native)):
            left, right = black_ink(native[index]), black_ink(current[index])
            assert left.size == right.size
            union = ImageChops.lighter(left, right).histogram()[255]
            assert union > 0
            result.append(ImageChops.difference(left, right).histogram()[255] / union)
        return result


def inputs():
    helper = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))
    defaults = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Document Sans SC" w:hAnsi="Document Sans SC"/>'
                '<w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
                '<w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>'
                '</w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Target">'
                '<w:name w:val="Target"/></w:style>')
    raw = helper["document"](defaults, "", "")
    with ZipFile(BytesIO(raw)) as package:
        parts = {name: package.read(name) for name in package.namelist()}

    def paragraph(labels, properties=""):
        return '<w:p><w:pPr>' + properties + '</w:pPr><w:r>' + '<w:br/>'.join(f'<w:t>{label}</w:t>' for label in labels) + '</w:r></w:p>'

    for name, count in CASES.items():
        body = paragraph(["BEFORE"]) + paragraph(["ANCHOR"], '<w:pageBreakBefore/>')
        body += paragraph([f"FILL{i:02d}" for i in range(count)])
        labels = ["TARGET00"] if name == "keep_with_next" else [f"TARGET{i:02d}" for i in range(4)]
        body += paragraph(labels, '<w:pStyle w:val="Target"/>')
        if name == "keep_with_next":
            body += paragraph(["FOLLOW00", "FOLLOW01"])
        body += '<w:sectPr><w:pgSz w:w="7200" w:h="4320"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0"/></w:sectPr>'
        stream = BytesIO()
        with ZipFile(stream, "w") as package:
            for part, data in parts.items():
                if part == "word/document.xml":
                    data = f'<w:document xmlns:w="{helper["W"]}"><w:body>{body}</w:body></w:document>'.encode()
                package.writestr(ZipInfo(part, (2020, 1, 1, 0, 0, 0)), data)
        yield name + ".docx", name.split("_orphan")[0].split("_widow")[0], stream.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("backend", choices=("commercial", "docweave"))
    parser.add_argument("output", type=Path)
    parser.add_argument("--fonts", type=Path, required=True)
    args = parser.parse_args()
    commercial = args.backend == "commercial"
    if commercial:
        import aspose.words as aw
        assert version("aspose-words") == "26.9.0"
    else:
        import aspose.words_foss as aw
    args.output.mkdir(parents=True, exist_ok=True)
    rows = []
    for name, prop, raw in inputs():
        (args.output / name).write_bytes(raw)
        for value in (False, True):
            if commercial:
                document = aw.Document(BytesIO(raw))
                settings = aw.fonts.FontSettings()
                settings.set_fonts_folder(str(args.fonts), False)
                document.font_settings = settings
            else:
                document = aw.DocxDocument(BytesIO(raw))
            setattr(document.styles.get_by_name("Target").paragraph_format, prop, value)
            output = args.output / f"{Path(name).stem}-{int(value)}.pdf"
            if commercial:
                document.save(str(output), aw.SaveFormat.PDF)
            else:
                from aspose.words_foss.pdf_writer import LdmPdfWriter
                LdmPdfWriter().write(document.to_light_document(), output)
            rows.append({"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(), "property": prop,
                         "value": value, "output": output.name, "output_sha256": hashlib.sha256(output.read_bytes()).hexdigest()})
    report = {"backend": args.backend, "version": version("aspose-words" if commercial else "aspose-words-foss-enhanced"),
              "licensed": False if commercial else None, "python": platform.python_version(), "platform": platform.platform(),
              "font_sha256": hashlib.sha256((args.fonts / "DocumentSansSC-Regular.ttf").read_bytes()).hexdigest(),
              "font_source_sha256": hashlib.sha256((Path(__file__).parents[2] / "aspose/words_foss/pdf_writer/fonts/DocumentSansSC-Regular.woff").read_bytes()).hexdigest(),
              "records": rows, "full_rendering_acceptance": False}
    (args.output / "outputs.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"pdf_outputs": len(rows)}))


if __name__ == "__main__":
    main()
