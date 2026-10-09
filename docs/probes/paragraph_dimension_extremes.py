"""Unresolved extreme native setter behavior; no saving or rendering acceptance."""

import json
import platform
import runpy
import sys
from importlib.metadata import version
from io import BytesIO
from pathlib import Path

import aspose.words as aw

root = Path(__file__).resolve().parents[2]
assert version("aspose-words") == "26.9.0"
h = runpy.run_path(str(root / "docs/probes/paragraph_dimensions.py"))
raw = next(h["inputs"]())[2]
values = [
    -107374182.4,
    -107374182.35,
    107374182.35,
    107374182.4,
    1e100,
    float("inf"),
    float("nan"),
]
rows = []
for prop in h["PROPERTIES"]:
    for value in values:
        doc = aw.Document(BytesIO(raw))
        p = h["paragraph"](doc, aw)
        row = {"property": prop, "value": repr(value)}
        try:
            setattr(p.paragraph_format, prop, value)
            row["read"] = repr(getattr(p.paragraph_format, prop))
        except (TypeError, ValueError, RuntimeError) as exc:
            row["error"] = type(exc).__name__
            row["message"] = str(exc)
        rows.append(row)
report = {
    "version": version("aspose-words"),
    "licensed": False,
    "python": platform.python_version(),
    "platform": platform.platform(),
    "scope": "live setters only; no save/reopen or rendering acceptance",
    "records": rows,
    "status": "unresolved: indent clamping, signed lower boundary and nonfinite sentinel behavior differ from SDK validation",
}
Path(sys.argv[1]).write_text(json.dumps(report, indent=2) + "\n")
print(len(rows))
