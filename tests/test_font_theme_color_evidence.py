"""Frozen theme evidence replays through the shared baseline verifier."""

import runpy
from pathlib import Path


def test_owned_theme_color_observations_are_reproducible():
    root = Path(__file__).parents[1]
    check = runpy.run_path(str(root / "scripts/verify_font_themes.py"))["verify_theme_colors"]
    assert check(root / "docs/benchmarks") == 32
