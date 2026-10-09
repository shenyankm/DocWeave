"""Observe style alignment inheritance and editing without altering owned fonts."""

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


def inputs():
    helpers = runpy.run_path(str(Path(__file__).with_name("import_style_conflicts.py")))
    for default, base, child in itertools.product((None, "left", "right"), (None, "center", "both"), (None, "left", "right")):
        defaults = '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
        defaults += '<w:pPrDefault><w:pPr>' + (f'<w:jc w:val="{default}"/>' if default else "") + '</w:pPr></w:pPrDefault></w:docDefaults>'
        styles = defaults + helpers["style"]("Base", "paragraph", {"jc": base}) + helpers["style"]("Derived", "paragraph", {"jc": child}, "Base")
        yield f"{default}-{base}-{child}.docx", helpers["document"](styles, '<w:pStyle w:val="Derived"/>', "")


def observe(document, aw):
    from style_save_state import snapshot

    value = snapshot(document, aw)
    value["style_alignment"] = {name: document.styles.get_by_name(name).paragraph_format.alignment.name for name in ("Base", "Derived")}
    return value


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("generate", "commercial"))
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.action == "generate":
        with ZipFile(args.corpus, "w") as archive:
            for name, data in inputs():
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), data)
        return
    if args.output is None:
        parser.error("--output is required")
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    records = []
    with ZipFile(args.corpus) as inputs_archive, ZipFile(args.output / "outputs.zip", "w", compression=ZIP_DEFLATED) as outputs:
        for name in sorted(inputs_archive.namelist()):
            data = inputs_archive.read(name)
            for alignment in (aw.ParagraphAlignment.LEFT, aw.ParagraphAlignment.CENTER, aw.ParagraphAlignment.RIGHT, aw.ParagraphAlignment.JUSTIFY):
                document = aw.Document(BytesIO(data))
                before = observe(document, aw)
                document.styles.get_by_name("Derived").paragraph_format.alignment = alignment
                edited = observe(document, aw)
                stream = BytesIO()
                document.save(stream, aw.SaveFormat.DOCX)
                saved = stream.getvalue()
                output = f"{len(records):03d}.docx"
                info = ZipInfo(output, (2020, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                outputs.writestr(info, saved)
                records.append({"input": name, "input_sha256": hashlib.sha256(data).hexdigest(),
                                "alignment": alignment.name, "before_edit": before, "after_edit": edited,
                                "after_save_live": observe(document, aw), "after_reopen": observe(aw.Document(BytesIO(saved)), aw),
                                "output": output, "output_sha256": hashlib.sha256(saved).hexdigest()})
    report = {"version": version("aspose-words"), "licensed": False, "python": platform.python_version(),
              "platform": platform.platform(), "corpus": "corpus/" + args.corpus.name,
              "corpus_sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest(),
              "outputs": "corpus/style-paragraph-format-26.9-outputs.zip",
              "outputs_sha256": hashlib.sha256((args.output / "outputs.zip").read_bytes()).hexdigest(),
              "records": records, "full_format_acceptance": False, "rendering_acceptance": False}
    (args.output / "observations.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"observations": len(records)}))


if __name__ == "__main__":
    main()
