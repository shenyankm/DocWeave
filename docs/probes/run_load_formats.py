"""Load actual baseline outputs; missing source formats remain unverified."""

import argparse
import hashlib
import importlib
import json
from io import BytesIO
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("enums", type=Path)
    parser.add_argument("saved", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--module", required=True, choices=["aspose.words", "aspose.words_foss"])
    args = parser.parse_args()
    aw = importlib.import_module(args.module)
    formats = json.loads(args.enums.read_text())["LoadFormat"]
    args.output.mkdir(parents=True, exist_ok=True)
    records = []
    for member in formats:
        source = args.source if member["name"] in {"AUTO", "UNKNOWN"} else args.saved / ("WORD_ML" if member["name"] == "XML" else member["name"]) / "output.bin"
        if not source.is_file():
            records.append({"name": member["name"], "value": member["value"], "status": "missing_legal_sample_not_verified"})
            continue
        data = source.read_bytes()
        record = {"name": member["name"], "value": member["value"], "source_sha256": hashlib.sha256(data).hexdigest(),
                  "source_kind": "original DOCX with UNKNOWN option" if member["name"] == "UNKNOWN" else "original DOCX" if member["name"] == "AUTO" else "commercial save output; XML uses WordML", "loaded": False}
        phase = "options"
        try:
            options = aw.loading.LoadOptions()
            options.load_format = aw.LoadFormat(member["value"]) if args.module == "aspose.words" else member["value"]
            phase = "load"
            document = aw.Document(BytesIO(data), options)
            record["loaded"] = True
            phase = "get_text"
            text = document.get_text()
            record["text"] = text
            record["label_offsets"] = {label: text.find(label) for label in ("ALPHA", "BETA", "CELL")}
            phase = "save_docx"
            output = args.output / (member["name"] + ".docx")
            document.save(str(output), aw.SaveFormat.DOCX)
            record["saved_docx"] = {"name": output.name, "sha256": hashlib.sha256(output.read_bytes()).hexdigest()}
            record["status"] = "returned"
            record["diagnostic_codes"] = [item.code for item in getattr(document, "diagnostics", [])]
        except Exception as error:  # noqa: BLE001 — backend exception types are benchmark observations.
            record.update({"status": "raised", "phase": phase, "exception_type": type(error).__name__, "message": str(error)[:800]})
        records.append(record)
    report = {"schema": 1, "module": args.module, "scope": "explicit load options plus get_text/save DOCX, one small actual input per format; missing formats retained",
              "records": records}
    (args.output / "observations.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"module": args.module, "formats": len(records), "loaded": sum(item.get("loaded", False) for item in records),
                      "missing_samples": [item["name"] for item in records if item["status"] == "missing_legal_sample_not_verified"]}))


if __name__ == "__main__":
    main()
