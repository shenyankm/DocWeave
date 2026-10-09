"""Bounded subprocess execution with process-group cleanup and capped error messages."""

import math
import os
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path

MAX_PROCESS_LOG_BYTES = 64 * 1024 * 1024


def _group_rss_kib(group):
    if sys.platform.startswith("linux"):
        total = 0
        for path in Path("/proc").iterdir():
            if not path.name.isdecimal():
                continue
            try:
                stat = (path / "stat").read_text()
                if int(stat[stat.rfind(")") + 2 :].split()[2]) != group:
                    continue
                for line in (path / "status").read_text().splitlines():
                    if line.startswith("VmRSS:"):
                        total += int(line.split()[1])
            except (OSError, ValueError, IndexError):
                if path.name == str(group) and path.exists():
                    raise RuntimeError("Cannot monitor worker memory")
        return total
    result = subprocess.run(
        ["ps", "-axo", "pgid=,rss="], check=True, capture_output=True, text=True, timeout=2
    )
    return sum(
        int(fields[1])
        for line in result.stdout.splitlines()
        if len(fields := line.split()) == 2 and int(fields[0]) == group
    )


def run_process(command, timeout, *, new_session=True, memory_mb=0):
    if not math.isfinite(timeout) or timeout <= 0:
        raise ValueError("timeout must be a positive finite number")
    grouped = new_session and os.name == "posix"
    with tempfile.TemporaryFile() as log:
        process = subprocess.Popen(
            command, stdout=log, stderr=subprocess.STDOUT, start_new_session=grouped
        )
        deadline = time.monotonic() + timeout
        group_cleaned = False
        try:
            while True:
                if os.fstat(log.fileno()).st_size > MAX_PROCESS_LOG_BYTES:
                    raise RuntimeError("Conversion worker exceeded the log output limit")
                if process.poll() is not None:
                    break
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise subprocess.TimeoutExpired(command, timeout)
                if memory_mb and _group_rss_kib(process.pid) > memory_mb * 1024:
                    raise RuntimeError(f"Conversion worker exceeded {memory_mb} MiB memory")
                try:
                    process.wait(timeout=min(0.1, remaining))
                except subprocess.TimeoutExpired:
                    pass
        except BaseException:
            if grouped:
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
                group_cleaned = True
            else:
                try:
                    process.kill()
                except ProcessLookupError:
                    pass
            process.wait(timeout=5)
            raise
        finally:
            if grouped and not group_cleaned:
                # Clean up descendants even when the parent exits first.
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
        log.seek(0)
        output = log.read(4096).decode("utf-8", errors="replace")
        if process.returncode:
            raise subprocess.CalledProcessError(process.returncode, command, output=output)
        return output
