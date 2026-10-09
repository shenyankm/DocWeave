"""Observe public runtime names and inheritance without constructing documents."""

import argparse
import importlib
import inspect
import json
import platform
import sys
from pathlib import Path
from types import ModuleType


def public_names(value, special_names):
    return sorted(name for name in dir(value) if not name.startswith("_") or name in special_names)


def observe(root, declarations):
    declared = {item["id"]: item for module in declarations["modules"] for item in module["symbols"]}
    special = {name.rsplit(".", 1)[-1] for name in declared if name.rsplit(".", 1)[-1].startswith("__")}
    pending = [root]
    seen = set()
    modules = []
    while pending:
        module = pending.pop()
        if module.__name__ in seen:
            continue
        seen.add(module.__name__)
        members = []
        for name in public_names(module, set()):
            value = getattr(module, name)
            qualified = module.__name__ + "." + name
            if isinstance(value, ModuleType):
                if value.__name__.startswith(root.__name__ + "."):
                    pending.append(value)
                members.append({"id": qualified, "kind": "module", "target": value.__name__})
            elif inspect.isclass(value):
                mro = inspect.getmro(value)
                visible = []
                for member in public_names(value, special):
                    owner = next((base for base in mro if member in vars(base)), None)
                    declared_at = None if owner is None else owner.__module__ + "." + owner.__name__ + "." + member
                    visible.append({"name": member, "owner": declared_at, "declaration_present": declared_at in declared})
                members.append({"id": qualified, "kind": "type", "declaration_present": qualified in declared, "target": value.__module__ + "." + value.__name__,
                                "mro": [base.__module__ + "." + base.__name__ for base in mro], "members": visible})
            else:
                members.append({"id": qualified, "kind": "callable" if callable(value) else "value", "declaration_present": qualified in declared})
        modules.append({"module": module.__name__, "declaration_module_present": any(item["module"] == module.__name__ for item in declarations["modules"]), "exports": members})
    return sorted(modules, key=lambda item: item["module"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("declarations", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--module", default="aspose.words")
    args = parser.parse_args()
    declarations = json.loads(args.declarations.read_text())
    root = importlib.import_module(args.module)
    modules = observe(root, declarations)
    version = getattr(root, "__version__", None)
    if version is None and hasattr(root, "BuildVersionInfo"):
        version = root.BuildVersionInfo.version
    result = {"schema": 1, "baseline_version": declarations["version"], "runtime_version": version,
              "root_module": args.module, "module_file": getattr(root, "__file__", None),
              "python": sys.version, "architecture": platform.machine(),
              "scope": "observed public names and class MRO; no behavioral parity claim",
              "module_count": len(modules), "modules": modules}
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"modules": len(modules), "modules_without_stubs": [item["module"] for item in modules if not item["declaration_module_present"]]}))


if __name__ == "__main__":
    main()
