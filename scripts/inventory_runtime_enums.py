"""Snapshot enum members, including aliases omitted by dir(Enum)."""

import argparse
import importlib
import json
from enum import Enum
from pathlib import Path


def members(cls):
    return [{"name": name, "value": int(value), "canonical_name": value.name}
            for name, value in cls.__members__.items() if not name.startswith("_")]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("runtime", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    runtime = json.loads(args.runtime.read_text())
    enums = []
    for module in runtime["modules"]:
        loaded = importlib.import_module(module["module"])
        for export in module["exports"]:
            if export["kind"] != "type":
                continue
            name = export["id"].rsplit(".", 1)[-1]
            cls = getattr(loaded, name)
            if not issubclass(cls, Enum):
                continue
            support = "published_declaration" if export["declaration_present"] else "official_lowcode_documentation_and_26.9_runtime" if module["module"] == "aspose.words.lowcode" else "runtime_only_unconfirmed"
            enums.append({"id": export["id"], "declaration_module_present": module["declaration_module_present"],
                          "support_evidence": support, "values": members(cls)})
    assert enums and len({item["id"] for item in enums}) == len(enums)
    result = {"schema": 1, "version": runtime["runtime_version"], "source": "runtime Enum.__members__, preserving aliases",
              "scope": "declared enums, documented lowcode exports, and separately marked unconfirmed runtime types",
              "enum_count": len(enums), "value_count": sum(len(item["values"]) for item in enums),
              "declared_enum_count": sum(item["support_evidence"] == "published_declaration" for item in enums),
              "documented_runtime_only_enum_count": sum(item["support_evidence"] == "official_lowcode_documentation_and_26.9_runtime" for item in enums),
              "unconfirmed_runtime_only_enum_count": sum(item["support_evidence"] == "runtime_only_unconfirmed" for item in enums),
              "alias_count": sum(value["name"] != value["canonical_name"] for item in enums for value in item["values"]), "enums": enums}
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({key: result[key] for key in ("enum_count", "value_count")}))


if __name__ == "__main__":
    main()
