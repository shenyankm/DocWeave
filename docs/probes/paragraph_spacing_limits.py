"""Owned paragraph/style spacing boundary observations, including DOCX reopening."""

import runpy
from pathlib import Path

VALUES = (-0.001, 0.0, 0.025, 0.075, 1583.975, 1584.0, 1584.001, 1e100)


def inputs():
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_dimensions.py")))
    selected = [row for row in probe["inputs"]() if row[0] in {"space_before/5.docx", "space_after/5.docx"}]
    assert len(selected) == 2
    yield from selected


if __name__ == "__main__":
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_dimensions.py")))
    probe["main"](inputs, VALUES)
