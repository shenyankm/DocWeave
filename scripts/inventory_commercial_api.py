"""Inventory published type declarations; never inspect proprietary binaries."""

import argparse
import ast
import hashlib
import importlib.metadata
import json
from pathlib import Path


def declaration_source(source):
    # Published chart docstrings have malformed closing quotes; ignore prose only.
    lines = []
    inside = False
    for line in source.splitlines():
        stripped = line.strip()
        if not inside and stripped.startswith('"""'):
            lines.append(line[:len(line) - len(line.lstrip())] + "pass")
            inside = not (len(stripped) >= 6 and stripped.endswith('"""'))
        elif inside:
            lines.append("")
            inside = not stripped.endswith('"""')
        else:
            lines.append(line)
    assert not inside, "unclosed declaration docstring"
    return "\n".join(lines)


def inventory(root):
    modules = []
    for path in sorted(root.rglob("*.pyi")):
        # Packaging hooks are tooling, not document-processing APIs.
        if any(part.startswith("__") for part in path.relative_to(root).parent.parts):
            continue
        relative = path.relative_to(root)
        parts = list(relative.with_suffix("").parts)
        if parts[-1] == "__init__":
            parts.pop()
        module = ".".join(["aspose", "words", *parts])
        tree = ast.parse(declaration_source(path.read_text(encoding="utf-8-sig")))
        symbols = []

        def collect(body, owner, symbols):
            groups = {}
            for node in body:
                if isinstance(node, ast.ClassDef):
                    if node.name.startswith("_"):
                        continue
                    name = owner + "." + node.name
                    bases = [ast.unparse(base) for base in node.bases]
                    symbols.append({"id": name, "kind": "enum" if "Enum" in bases else "type", "bases": bases})
                    collect(node.body, name, symbols)
                elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    if node.name.startswith("_") and not (node.name.startswith("__") and node.name.endswith("__")):
                        continue
                    name = owner + "." + node.name
                    decorators = [ast.unparse(value) for value in node.decorator_list]
                    signature = ast.unparse(node.args)
                    returns = ast.unparse(node.returns) if node.returns else None
                    entry = groups.setdefault(name, {"id": name, "kind": "method", "declarations": []})
                    if "property" in decorators or any(value.endswith(".setter") for value in decorators):
                        entry["kind"] = "property"
                    entry["declarations"].append({"arguments": signature, "returns": returns, "decorators": decorators})
                elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
                    if not node.target.id.startswith("_"):
                        symbols.append({"id": owner + "." + node.target.id, "kind": "attribute", "annotation": ast.unparse(node.annotation),
                                        "value": ast.unparse(node.value) if node.value is not None else None})
                elif isinstance(node, ast.Assign):
                    for target in node.targets:
                        if isinstance(target, ast.Name) and not target.id.startswith("_"):
                            symbols.append({"id": owner + "." + target.id, "kind": "constant", "value": ast.unparse(node.value)})
            symbols.extend(groups.values())

        collect(tree.body, module, symbols)
        assert len({item["id"] for item in symbols}) == len(symbols), module
        modules.append({"module": module, "declaration_file": str(relative), "sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "symbols": sorted(symbols, key=lambda item: item["id"])})
    assert modules and any(module["module"] == "aspose.words" for module in modules)
    return modules


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    distribution = importlib.metadata.distribution("aspose-words")
    root = Path(distribution.locate_file("aspose/words"))
    modules = inventory(root)
    result = {
        "schema": 1,
        "product": "Aspose.Words for Python via .NET",
        "version": distribution.version,
        "source": "published wheel type declarations, docstrings omitted before parsing; no implementation bodies",
        "coverage": "declared aspose.words public API; runtime inherited members and behavior remain unverified",
        "capability_defaults": {
            "baseline_behavior": "not_measured", "implementation": "not_mapped",
            "gap": "not_assessed", "dependencies": [], "approach": "pending_behavior_probe",
            "samples": [], "validation": [], "exit_criteria": "fixed-version behavior and independent output verified",
            "limits": ["evaluation license; inherited members require runtime expansion"], "delivery": "not_started",
        },
        "module_count": len(modules),
        "symbol_count": sum(len(module["symbols"]) for module in modules),
        "modules": modules,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({key: result[key] for key in ("version", "module_count", "symbol_count")}))


if __name__ == "__main__":
    main()
