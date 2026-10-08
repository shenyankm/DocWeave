"""The benchmark leaves useful machine-readable measurements, not just printed timings."""

import json
from pathlib import Path
import subprocess
import sys


def test_cold_process_benchmark_smoke():
    script = Path(__file__).resolve().parents[1] / "scripts/benchmark.py"
    run = subprocess.run([sys.executable, str(script), "--case", "small", "--repeat", "1"],
                         capture_output=True, text=True, timeout=60)
    assert run.returncode == 0, run.stderr
    report = json.loads(run.stdout)
    result = report["median"]["small"]
    assert result["pages"] == 1 and result["output_bytes"] > 1000
    assert result["process_ms"] >= result["total_ms"]
    assert all(result[name] >= 0 for name in ("import_ms", "parse_ms", "layout_ms", "serialize_ms", "write_ms"))
    assert result["peak_rss_bytes"] is None or result["peak_rss_bytes"] > 0
