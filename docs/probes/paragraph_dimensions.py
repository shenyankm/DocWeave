"""Owned point-valued paragraph dimensions: layers, quantization and native edits."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROPERTIES = {
    "left_indent": ("ind", "left"), "right_indent": ("ind", "right"),
    "first_line_indent": ("ind", "firstLine"), "space_before": ("spacing", "before"),
    "space_after": ("spacing", "after"),
}
PATTERNS = (
    (None, None, None, None), (240, None, None, None),
    (None, 360, None, None), (None, None, 480, None),
    (None, None, None, 600), (240, 360, 480, 600),
    (240, 360, 0, None), (240, 360, 480, 0),
    (240, 360, None, None),
)
VALUES = (-12.375, 0.0, 12.35, 12.375)


def inputs():
    helper = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))
    for prop, (tag, attribute) in PROPERTIES.items():
        for index, values in enumerate(PATTERNS):
            def fragment(value, sibling=False, tag=tag, attribute=attribute):
                # Unrelated attributes at the child layer must not erase inherited values.
                other = "right" if attribute == "left" else "left" if tag == "ind" else "after" if attribute == "before" else "before"
                attrs = f' w:{other}="80"' if sibling else ""
                if value is not None:
                    attrs += f' w:{attribute}="{value}"'
                return f'<w:{tag}{attrs}/>' if attrs else ""

            default, base, child, direct = values
            styles = '<w:docDefaults><w:pPrDefault><w:pPr>' + fragment(default)
            styles += '</w:pPr></w:pPrDefault></w:docDefaults>'
            for name, value, parent in (("Base", base, ""), ("Derived", child, '<w:basedOn w:val="Base"/>')):
                styles += f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>{parent}<w:pPr>{fragment(value, index == 8 and name == "Derived")}</w:pPr></w:style>'
            raw = helper["document"](styles, '<w:pStyle w:val="Derived"/>' + fragment(direct), "")
            with ZipFile(BytesIO(raw)) as package:
                parts = {name: package.read(name) for name in package.namelist()}
            parts["word/document.xml"] = parts["word/document.xml"].replace(b"<w:body>", b"<w:body><w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>", 1)
            stream = BytesIO()
            with ZipFile(stream, "w") as package:
                for part, data in parts.items():
                    package.writestr(ZipInfo(part, (2020, 1, 1, 0, 0, 0)), data)
            yield f"{prop}/{index}.docx", prop, stream.getvalue()


def paragraph(document, aw):
    return next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                if "IMPORT" in node.get_text())


def snapshot(document, aw):
    return {name: {prop: getattr(node.paragraph_format, prop) for prop in PROPERTIES}
            for name, node in (("base", document.styles.get_by_name("Base")),
                               ("derived", document.styles.get_by_name("Derived")),
                               ("paragraph", paragraph(document, aw)))}


def main(input_generator=inputs, values=VALUES):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    records, errors = [], []
    generated = list(input_generator())
    with ZipFile(args.output / "inputs.zip", "w", compression=ZIP_DEFLATED) as corpus, \
            ZipFile(args.output / "outputs.zip", "w", compression=ZIP_DEFLATED) as outputs:
        for name, prop, raw in generated:
            info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            corpus.writestr(info, raw)
            for target in ("style", "paragraph"):
                for value in values:
                    document = aw.Document(BytesIO(raw))
                    before = snapshot(document, aw)
                    node = document.styles.get_by_name("Derived") if target == "style" else paragraph(document, aw)
                    try:
                        setattr(node.paragraph_format, prop, value)
                        error = None
                    except (TypeError, ValueError, RuntimeError) as exc:
                        error = type(exc).__name__
                    edited = snapshot(document, aw)
                    row = {"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(),
                           "property": prop, "target": target, "value": value, "error": error,
                           "before_edit": before, "after_edit": edited}
                    if error is None:
                        stream = BytesIO()
                        document.save(stream, aw.SaveFormat.DOCX)
                        saved = stream.getvalue()
                        output = f"{len(records):03d}.docx"
                        info = ZipInfo(output, (2020, 1, 1, 0, 0, 0))
                        info.compress_type = ZIP_DEFLATED
                        outputs.writestr(info, saved)
                        row.update(output=output, output_sha256=hashlib.sha256(saved).hexdigest(),
                                   after_save_live=snapshot(document, aw),
                                   after_reopen=snapshot(aw.Document(BytesIO(saved)), aw))
                    else:
                        assert edited == before
                    records.append(row)
                for value in (None, True, "bad"):
                    document = aw.Document(BytesIO(raw))
                    node = document.styles.get_by_name("Derived") if target == "style" else paragraph(document, aw)
                    before = snapshot(document, aw)
                    try:
                        setattr(node.paragraph_format, prop, value)
                        error = None
                    except (TypeError, ValueError, RuntimeError) as exc:
                        error = type(exc).__name__
                    after = snapshot(document, aw)
                    assert after == before
                    errors.append({"input": name, "property": prop, "target": target, "value": value, "error": error})
    assert len(records) == len(generated) * 2 * len(values) and len(errors) == len(generated) * 6
    report = {"version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "records": records, "setter_errors": errors,
              "corpus": "inputs.zip", "corpus_sha256": hashlib.sha256((args.output / "inputs.zip").read_bytes()).hexdigest(),
              "outputs": "outputs.zip", "outputs_sha256": hashlib.sha256((args.output / "outputs.zip").read_bytes()).hexdigest(),
              "scope": "owned point-valued paragraph inputs, style/direct edits and DOCX reopening; not complete paragraph formatting or rendering acceptance"}
    (args.output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"records": len(records), "setter_errors": len(errors)}))


if __name__ == "__main__":
    main()
