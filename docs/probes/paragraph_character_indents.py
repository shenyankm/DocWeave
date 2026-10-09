"""Owned character-indent lifecycle and edits; no rendering or SDK acceptance."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROPERTIES = ("left_indent", "right_indent", "first_line_indent",
              "character_unit_left_indent", "character_unit_right_indent",
              "character_unit_first_line_indent")
FONT_CONTEXTS = ((None, None, None, None), (20, None, None, None),
                 (None, 20, None, None), (None, None, 20, None),
                 (None, None, None, 20), (12, 16, 20, 24),
                 (12, 16, None, 24), (12, None, 20, 24))
VALUES = (-1.235, 0.0, 1.235, 1.225, 12.375)


def inputs():
    helper = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))
    for index, (default, style, mark, run) in enumerate(FONT_CONTEXTS):
        def size(value):
            return f'<w:sz w:val="{value * 2}"/>' if value is not None else ""
        styles = '<w:docDefaults><w:rPrDefault><w:rPr>' + size(default)
        styles += '</w:rPr></w:rPrDefault></w:docDefaults>'
        styles += '<w:style w:type="paragraph" w:styleId="P"><w:name w:val="P"/><w:rPr>' + size(style) + '</w:rPr></w:style>'
        ppr = '<w:pStyle w:val="P"/><w:ind w:leftChars="100" w:rightChars="200" w:firstLineChars="50"/><w:rPr>' + size(mark) + '</w:rPr>'
        yield f"fonts/{index}.docx", helper["document"](styles, ppr, size(run))
    for prop, point, char in zip(PROPERTIES[3:], ("left", "right", "firstLine"), ("leftChars", "rightChars", "firstLineChars"), strict=True):
        for index, attrs in enumerate((f'w:{char}="125"', f'w:{char}="125" w:{point}="240"', f'w:{point}="240" w:{char}="125"')):
            yield f"{prop}/{index}.docx", helper["document"]("", f'<w:ind {attrs}/>', "")


def nonfirst_inputs():
    for name, raw in inputs():
        with ZipFile(BytesIO(raw)) as archive:
            parts = {key: archive.read(key) for key in archive.namelist()}
        parts["word/document.xml"] = parts["word/document.xml"].replace(b"<w:body>", b"<w:body><w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>", 1)
        stream = BytesIO()
        with ZipFile(stream, "w") as archive:
            for key, data in parts.items():
                archive.writestr(ZipInfo(key, (2020, 1, 1, 0, 0, 0)), data)
        yield name, stream.getvalue()


def snapshot(document, aw):
    paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True) if "IMPORT" in node.get_text())
    return {prop: getattr(paragraph.paragraph_format, prop) for prop in PROPERTIES}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    rows, errors = [], []
    with ZipFile(args.output / "inputs.zip", "w", compression=ZIP_DEFLATED) as corpus, \
            ZipFile(args.output / "outputs.zip", "w", compression=ZIP_DEFLATED) as outputs:
        for name, raw in nonfirst_inputs():
            corpus.writestr(name, raw)
            font_case = name.startswith("fonts/")
            cases = [(None, layout) for layout in (False, True)] if font_case else [(value, False) for value in VALUES]
            prop = None if font_case else name.split("/")[0]
            for value, layout in cases:
                document = aw.Document(BytesIO(raw))
                row = {"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(),
                       "property": prop, "value": value, "layout": layout,
                       "loaded": snapshot(document, aw)}
                if prop:
                    target = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True) if "IMPORT" in node.get_text())
                    setattr(target.paragraph_format, prop, value)
                row["after_edit"] = snapshot(document, aw)
                if layout:
                    document.update_page_layout()
                row["before_save"] = snapshot(document, aw)
                stream = BytesIO()
                document.save(stream, aw.SaveFormat.DOCX)
                saved = stream.getvalue()
                output = f"{len(rows):03d}.docx"
                outputs.writestr(output, saved)
                row.update(output=output, output_sha256=hashlib.sha256(saved).hexdigest(),
                           after_save_live=snapshot(document, aw),
                           after_reopen=snapshot(aw.Document(BytesIO(saved)), aw))
                rows.append(row)
            if prop:
                for value in (None, True, "bad"):
                    document = aw.Document(BytesIO(raw))
                    before = snapshot(document, aw)
                    target = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True) if "IMPORT" in node.get_text())
                    try:
                        setattr(target.paragraph_format, prop, value)
                    except TypeError:
                        error = "TypeError"
                    else:
                        raise AssertionError("Expected a native setter type error")
                    assert snapshot(document, aw) == before
                    errors.append({"input": name, "property": prop, "value": value, "error": error})
    assert len(rows) == 61 and len(errors) == 27
    report = {"version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "full_format_acceptance": False, "rendering_acceptance": False,
              "sdk_acceptance": False, "font_contexts": FONT_CONTEXTS,
              "records": rows, "setter_errors": errors,
              "corpus": "inputs.zip", "corpus_sha256": hashlib.sha256((args.output / "inputs.zip").read_bytes()).hexdigest(),
              "outputs": "outputs.zip", "outputs_sha256": hashlib.sha256((args.output / "outputs.zip").read_bytes()).hexdigest(),
              "scope": "nonfirst paragraph: 8 font contexts with/without layout and 45 character setters; no style edits, aliases, complex scripts or rendered geometry"}
    (args.output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"records": len(rows), "setter_errors": len(errors)}))


if __name__ == "__main__":
    main()
