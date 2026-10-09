"""Owned paragraph pagination flags: inheritance, style/direct edits and DOCX reopening."""

import argparse
import hashlib
import itertools
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROPERTIES = {"keep_with_next": "keepNext", "keep_together": "keepLines",
              "page_break_before": "pageBreakBefore", "widow_control": "widowControl"}


def inputs(first_paragraph=False):
    document = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))["document"]
    for prop, tag in PROPERTIES.items():
        for values in itertools.product((None, False, True), repeat=4):
            def flag(value, tag=tag):
                return "" if value is None else f'<w:{tag} w:val="{int(value)}"/>'
            default, base, child, direct = values
            styles = '<w:docDefaults><w:pPrDefault><w:pPr>' + flag(default) + '</w:pPr></w:pPrDefault></w:docDefaults>'
            for name, value, parent in (("Base", base, ""), ("Derived", child, '<w:basedOn w:val="Base"/>')):
                styles += f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>{parent}<w:pPr>{flag(value)}</w:pPr></w:style>'
            name = prop + "/" + "-".join(str(value) for value in values) + ".docx"
            raw = document(styles, '<w:pStyle w:val="Derived"/>' + flag(direct), "")
            with ZipFile(BytesIO(raw)) as package:
                parts = {name: package.read(name) for name in package.namelist()}
            parts["word/document.xml"] = parts["word/document.xml"] if first_paragraph else parts["word/document.xml"].replace(b"<w:body>", b"<w:body><w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>", 1)
            stream = BytesIO()
            with ZipFile(stream, "w") as package:
                for part, data in parts.items():
                    package.writestr(ZipInfo(part, (2020, 1, 1, 0, 0, 0)), data)
            yield name, prop, stream.getvalue()


def snapshot(document, prop, aw):
    paragraph = next(node for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                     if "IMPORT" in node.get_text())
    return {"base": getattr(document.styles.get_by_name("Base").paragraph_format, prop),
            "derived": getattr(document.styles.get_by_name("Derived").paragraph_format, prop),
            "paragraph": getattr(paragraph.as_paragraph().paragraph_format, prop)}


def setter_errors(aw, raw):
    records = []
    for prop, target, value in itertools.product(PROPERTIES, ("style", "paragraph"), (None, 1, "bad")):
        document = aw.Document(BytesIO(raw))
        node = document.styles.get_by_name("Derived") if target == "style" else next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True) if "IMPORT" in node.get_text())
        try:
            setattr(node.paragraph_format, prop, value)
            error = None
        except (TypeError, ValueError, RuntimeError) as exc:
            error = type(exc).__name__
        records.append({"property": prop, "target": target, "value": value, "error": error})
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw
    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    records = []
    with ZipFile(args.output / "inputs.zip", "w", compression=ZIP_DEFLATED) as corpus, ZipFile(args.output / "outputs.zip", "w", compression=ZIP_DEFLATED) as outputs:
        for name, prop, raw in inputs():
            info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            corpus.writestr(info, raw)
            for target, value in itertools.product(("style", "paragraph"), (False, True)):
                document = aw.Document(BytesIO(raw))
                before = snapshot(document, prop, aw)
                node = document.styles.get_by_name("Derived") if target == "style" else next(node for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True) if "IMPORT" in node.get_text())
                setattr((node if target == "style" else node.as_paragraph()).paragraph_format, prop, value)
                edited = snapshot(document, prop, aw)
                stream = BytesIO()
                document.save(stream, aw.SaveFormat.DOCX)
                saved = stream.getvalue()
                output = f"{len(records):04d}.docx"
                info = ZipInfo(output, (2020, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                outputs.writestr(info, saved)
                records.append({"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(), "property": prop,
                                "target": target, "value": value, "before": before, "edited": edited,
                                "saved_live": snapshot(document, prop, aw), "reopened": snapshot(aw.Document(BytesIO(saved)), prop, aw),
                                "output": output, "output_sha256": hashlib.sha256(saved).hexdigest()})
    report = {"version": version("aspose-words"), "licensed": False, "python": platform.python_version(),
              "platform": platform.platform(), "rendering_acceptance": False,
              "corpus": "corpus/paragraph-pagination-26.9.zip", "outputs": "corpus/paragraph-pagination-26.9-outputs.zip",
              "setter_errors": setter_errors(aw, next(inputs())[2]),
              "corpus_sha256": hashlib.sha256((args.output / "inputs.zip").read_bytes()).hexdigest(),
              "outputs_sha256": hashlib.sha256((args.output / "outputs.zip").read_bytes()).hexdigest(), "records": records}
    (args.output / "observations.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"observations": len(records)}))


if __name__ == "__main__":
    main()
