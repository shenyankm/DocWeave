"""Compare exposed enum constants; equal integers are not functional parity."""

import argparse
import importlib
import json
from enum import Enum
from pathlib import Path


def compare(enums, root):
    rows = []
    for baseline in enums:
        path = baseline["id"].removeprefix("aspose.words.").split(".")
        cls = root
        for name in path:
            cls = getattr(cls, name, None) if cls is not None else None
        values = []
        for member in baseline["values"]:
            actual = getattr(cls, member["name"], None) if cls is not None else None
            integer = actual if isinstance(actual, int) and not isinstance(actual, bool) else None
            status = "absent_or_not_integer" if integer is None else "integer_matches" if integer == member["value"] else "integer_differs"
            values.append({"name": member["name"], "baseline": member["value"], "current": integer, "status": status})
        rows.append({"id": baseline["id"], "current_type_present": cls is not None,
                     "current_is_enum": isinstance(cls, type) and issubclass(cls, Enum), "values": values})
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    baseline = json.loads(args.baseline.read_text())
    root = importlib.import_module("aspose.words_foss")
    rows = compare(baseline["enums"], root)
    result = {"schema": 1, "baseline_version": baseline["version"], "current_version": root.__version__,
              "scope": "direct public path and integer constants only; aliases, constructors and behavior not certified",
              "enum_types": rows}
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    counts = {status: sum(value["status"] == status for row in rows for value in row["values"])
              for status in ("integer_matches", "integer_differs", "absent_or_not_integer")}
    print(json.dumps(counts))


if __name__ == "__main__":
    main()
