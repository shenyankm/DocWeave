"""Observe inherited Style.font edits and the live/cold run formats they affect."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from style_save_state import snapshot


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    raw = args.corpus.read_bytes()
    records = []
    with ZipFile(BytesIO(raw)) as corpus, ZipFile(args.output / "outputs.zip", "w", compression=ZIP_DEFLATED) as outputs:
        inputs = sorted(name for name in corpus.namelist() if name.endswith("/source.docx"))
        assert len(inputs) == 247
        for name in inputs:
            data = corpus.read(name)
            for prop, values in (("bold", (False, True)), ("italic", (False, True)), ("size", (10, 17.5))):
                for value in values:
                    document = aw.Document(BytesIO(data))
                    style = document.styles.get_by_name("Derived")
                    assert style is not None
                    before = snapshot(document, aw)
                    setattr(style.font, prop, value)
                    edited = snapshot(document, aw)
                    assert edited["styles"]["Derived"][prop] == value
                    stream = BytesIO()
                    document.save(stream, aw.SaveFormat.DOCX)
                    saved = stream.getvalue()
                    output = f"{len(records):04d}.docx"
                    info = ZipInfo(output, (2020, 1, 1, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    outputs.writestr(info, saved)
                    records.append({
                        "input": name, "input_sha256": hashlib.sha256(data).hexdigest(),
                        "style": "Derived", "property": prop, "value": value,
                        "before_edit": before, "after_edit": edited,
                        "after_save_live": snapshot(document, aw),
                        "after_reopen": snapshot(aw.Document(BytesIO(saved)), aw),
                        "output": output, "output_sha256": hashlib.sha256(saved).hexdigest(),
                    })
    report = {
        "module": "aspose.words", "version": version("aspose-words"), "licensed": False,
        "python": platform.python_version(), "platform": platform.platform(),
        "corpus": args.corpus.name, "corpus_sha256": hashlib.sha256(raw).hexdigest(),
        "outputs_sha256": hashlib.sha256((args.output / "outputs.zip").read_bytes()).hexdigest(),
        "scope": "Derived paragraph/character Style.font bold, italic, size edits; owned live/cold run and style getters",
        "full_format_acceptance": False, "rendering_acceptance": False, "records": records,
    }
    (args.output / "observations.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"observations": len(records)}))


if __name__ == "__main__":
    main()
