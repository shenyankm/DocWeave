"""Frozen theme evidence replays through the shared baseline verifier."""

import runpy
from pathlib import Path


def test_theme_declarations_and_native_cold_getters_are_frozen():
    root = Path(__file__).parents[1]
    check = runpy.run_path(str(root / "scripts/verify_font_themes.py"))["verify_theme_roundtrips"]
    assert check(root / "docs/benchmarks") == 128
