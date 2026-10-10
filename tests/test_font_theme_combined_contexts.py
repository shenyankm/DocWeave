"""Frozen theme evidence replays through the shared baseline verifier."""

import runpy
from pathlib import Path


def test_combined_theme_contexts_and_painted_colors_replay():
    root = Path(__file__).parents[1]
    check = runpy.run_path(str(root / "scripts/verify_font_themes.py"))["verify_theme_contexts"]
    assert check(root / "docs/benchmarks") == 16
