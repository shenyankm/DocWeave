"""Observe explicit save-format values with a fixed native DOCX input."""

import argparse
import hashlib
import importlib
import json
import time
from io import BytesIO
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("enums", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--module", required=True, choices=["aspose.words", "aspose.words_foss"])
    args = parser.parse_args()
    aw = importlib.import_module(args.module)
    source = args.source.read_bytes()
    formats = json.loads(args.enums.read_text())["SaveFormat"]
    args.output.mkdir(parents=True, exist_ok=True)
    records = []
    for member in formats:
        folder = args.output / member["name"]
        folder.mkdir(exist_ok=True)
        assert not any(folder.iterdir()), "use an empty output folder for each probe run"
        document = aw.Document(BytesIO(source))
        value = aw.SaveFormat(member["value"]) if args.module == "aspose.words" else member["value"]
        start = time.monotonic()
        try:
            result = document.save(str(folder / "output.bin"), value)
            outcome = {"status": "returned", "return_type": type(result).__name__}
        except (RuntimeError, ValueError, TypeError, NotImplementedError) as error:
            outcome = {"status": "raised", "exception_type": type(error).__name__, "message": str(error)[:800]}
        files = []
        for output in sorted(folder.rglob("*")):
            if output.is_file():
                data = output.read_bytes()
                files.append({"name": str(output.relative_to(folder)), "size": len(data),
                              "sha256": hashlib.sha256(data).hexdigest(), "first_16_bytes_hex": data[:16].hex()})
        records.append({"name": member["name"], "value": member["value"], "outcome": outcome,
                        "elapsed_seconds": time.monotonic() - start, "files": files,
                        "diagnostic_codes": [item.code for item in getattr(document, "diagnostics", [])]})
    report = {"schema": 1, "module": args.module, "source_sha256": hashlib.sha256(source).hexdigest(),
              "adapter": "commercial SaveFormat(value); local documented integer argument; explicit value with .bin output filename",
              "scope": "single small native DOCX, default options, file output only; no format fidelity or all-overload claim",
              "records": records}
    (args.output / "observations.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"module": args.module, "formats": len(records),
                      "returned": sum(item["outcome"]["status"] == "returned" for item in records)}))


if __name__ == "__main__":
    main()
