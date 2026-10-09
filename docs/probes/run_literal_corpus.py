"""Replay the fixed, independently generated DOCX corpus with either runtime."""

import argparse
import hashlib
import importlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--module", required=True, choices=["aspose.words", "aspose.words_foss"])
    parser.add_argument("--label", required=True, choices=["commercial", "docweave"])
    args = parser.parse_args()
    aw = importlib.import_module(args.module)
    args.output.mkdir(parents=True, exist_ok=True)
    records = []
    with ZipFile(args.corpus) as corpus:
        for name in sorted(corpus.namelist()):
            assert Path(name).name == name and name.endswith(".docx"), name
            source = corpus.read(name)
            output = args.output / (Path(name).stem + "." + args.label + ".md")
            aw.Document(BytesIO(source)).save(str(output), aw.SaveFormat.MARKDOWN)
            records.append({"source": name, "source_sha256": hashlib.sha256(source).hexdigest(),
                            "markdown": output.read_text(), "output_sha256": hashlib.sha256(output.read_bytes()).hexdigest()})
    assert len(records) == 9, "unexpected fixed corpus count"
    (args.output / (args.label + ".json")).write_text(json.dumps(records, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"module": args.module, "samples": len(records)}))


if __name__ == "__main__":
    main()
