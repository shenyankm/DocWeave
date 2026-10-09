"""Real subprocess checks; LibreOffice protocol is mocked when not installed."""

import os
import subprocess
import sys
import time
from pathlib import Path

import pytest
from pypdf import PdfReader

from aspose.words_foss import _process, libreoffice
from aspose.words_foss._process import run_process

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "tests/data/input/chinese_pdf.docx"


def test_libreoffice_receives_original_file_and_private_profile(tmp_path, monkeypatch):
    monkeypatch.setattr(libreoffice.shutil, "which", lambda name: "/fake/soffice")
    profiles = []

    def convert(command, timeout, **kwargs):
        original = Path(command[-1])
        assert original.read_bytes() == SOURCE.read_bytes()
        assert "--headless" in command and "pdf:writer_pdf_Export" in command
        assert kwargs["new_session"]
        profiles.append(next(arg for arg in command if arg.startswith("-env:UserInstallation=")))
        outdir = Path(command[command.index("--outdir") + 1])
        (outdir / (original.stem + ".pdf")).write_bytes(b"%PDF-1.7\nmock\n%%EOF")
        return "converted"

    monkeypatch.setattr(libreoffice, "run_process", convert)
    output = tmp_path / "report.pdf"
    libreoffice.convert_to_pdf(SOURCE, output)
    libreoffice.convert_to_pdf(SOURCE, output)
    assert profiles[0] != profiles[1]
    assert output.read_bytes().startswith(b"%PDF-")
    assert not list(tmp_path.glob(".libreoffice-*"))


def test_libreoffice_relative_output(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(libreoffice.shutil, "which", lambda name: "/fake/soffice")

    def convert(command, timeout, **kwargs):
        assert "file:///" in command[1]
        outdir = Path(command[command.index("--outdir") + 1])
        (outdir / (SOURCE.stem + ".pdf")).write_bytes(b"%PDF-1.7\nmock\n%%EOF")
        return ""

    monkeypatch.setattr(libreoffice, "run_process", convert)
    libreoffice.convert_to_pdf(SOURCE, Path("out") / "report.pdf")
    assert Path("out/report.pdf").is_file()


@pytest.mark.parametrize("failure", ["missing", "invalid", "incomplete", "timeout"])
def test_libreoffice_failure_preserves_output(tmp_path, monkeypatch, failure):
    output = tmp_path / "report.pdf"
    output.write_bytes(b"existing output")
    monkeypatch.setattr(libreoffice.shutil, "which", lambda name: "/fake/soffice")

    def convert(command, timeout, **kwargs):
        if failure == "timeout":
            raise subprocess.TimeoutExpired(command, timeout)
        if failure in ("invalid", "incomplete"):
            outdir = Path(command[command.index("--outdir") + 1])
            data = b"not a PDF" if failure == "invalid" else b"%PDF-1.7\npartial"
            (outdir / (SOURCE.stem + ".pdf")).write_bytes(data)
        return "load failed"

    monkeypatch.setattr(libreoffice, "run_process", convert)
    with pytest.raises((RuntimeError, subprocess.TimeoutExpired)):
        libreoffice.convert_to_pdf(SOURCE, output)
    assert output.read_bytes() == b"existing output"
    assert sorted(p.name for p in tmp_path.iterdir()) == ["report.pdf"]


@pytest.mark.skipif(os.name != "posix", reason="Bounded CLI requires process groups")
def test_bounded_cli_converts_chinese(tmp_path):
    output = tmp_path / "report.pdf"
    result = subprocess.run(
        [sys.executable, "-m", "aspose.words_foss.convert", str(SOURCE), str(output), "--strict"],
        capture_output=True,
        text=True,
        timeout=20,
    )
    assert result.returncode == 0, result.stderr
    assert "中文" in "".join(page.extract_text() for page in PdfReader(output).pages)
    assert not list(tmp_path.glob(".conversion-*"))


@pytest.mark.skipif(os.name != "posix", reason="Bounded CLI requires process groups")
@pytest.mark.parametrize("mode", ["strict", "timeout", "memory"])
def test_bounded_cli_failure_preserves_output(tmp_path, mode):
    source = tmp_path / "source.txt"
    source.write_text("中文 " + chr(0x1FAE0))
    output = tmp_path / "report.pdf"
    output.write_bytes(b"existing output")
    options = {
        "strict": ["--strict"],
        "timeout": ["--timeout", "0.001"],
        "memory": ["--memory-mb", "1"],
    }[mode]
    result = subprocess.run(
        [sys.executable, "-m", "aspose.words_foss.convert", str(source), str(output), *options],
        capture_output=True,
        text=True,
        timeout=20,
    )
    assert result.returncode != 0
    assert output.read_bytes() == b"existing output"
    assert sorted(p.name for p in tmp_path.iterdir()) == ["report.pdf", "source.txt"]


@pytest.mark.skipif(os.name != "posix", reason="Bounded CLI requires process groups")
def test_strict_cli_rejects_known_table_write_loss(tmp_path):
    from aspose.words_foss import light_document_model as ldm
    from aspose.words_foss.docx_writer import LdmDocxWriter
    paragraph = ldm.Paragraph(children=[ldm.Run(text="中文" * 200)])
    grid = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph])],
                                 row_format=ldm.RowFormat(allow_break_across_pages=False))])
    model = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(page_width=240, page_height=150),
                                              body=ldm.Body(children=[grid]))])
    source = tmp_path / "source.docx"
    LdmDocxWriter().write(model, source)
    output = tmp_path / "existing.pdf"
    output.write_bytes(b"old document")
    run = subprocess.run([sys.executable, "-m", "aspose.words_foss.convert", str(source), str(output), "--strict"],
                         capture_output=True, text=True, timeout=20)
    assert run.returncode != 0 and "cantSplit" in run.stderr
    assert output.read_bytes() == b"old document"


@pytest.mark.skipif(os.name != "posix", reason="Process groups require POSIX")
def test_timeout_kills_descendants(tmp_path):
    pid_file = tmp_path / "child.pid"
    script = (
        "import subprocess,sys,time,pathlib; "
        "p=subprocess.Popen([sys.executable,'-c','import time;time.sleep(30)']); "
        f"pathlib.Path({str(pid_file)!r}).write_text(str(p.pid)); time.sleep(30)"
    )
    with pytest.raises(subprocess.TimeoutExpired):
        run_process([sys.executable, "-c", script], 1)
    assert pid_file.exists()
    pid = pid_file.read_text()
    # A dead child may briefly remain a zombie before its new parent reaps it.
    for _ in range(20):
        result = subprocess.run(["ps", "-o", "stat=", "-p", pid], capture_output=True, text=True)
        if not result.stdout.strip() or result.stdout.lstrip().startswith("Z"):
            break
        time.sleep(0.05)
    else:
        pytest.fail("Timed-out child is still running")


@pytest.mark.skipif(os.name != "posix", reason="Process groups require POSIX")
def test_timeout_does_not_signal_reaped_process_group_twice(monkeypatch):
    original = os.killpg
    calls = []

    def signal_group(group, sig):
        calls.append((group, sig))
        if len(calls) > 1:
            raise PermissionError("reaped process group cannot be signalled again")
        return original(group, sig)

    monkeypatch.setattr(_process.os, "killpg", signal_group)
    with pytest.raises(subprocess.TimeoutExpired):
        run_process([sys.executable, "-c", "import time; time.sleep(30)"], 0.2)
    assert len(calls) == 1


@pytest.mark.skipif(os.name != "posix", reason="RSS watchdog requires POSIX")
def test_memory_watchdog_kills_worker():
    script = "import time; payload=bytearray(50*1024*1024); time.sleep(10)"
    with pytest.raises(RuntimeError, match="exceeded.*memory"):
        run_process([sys.executable, "-c", script], 5, memory_mb=12)


@pytest.mark.skipif(os.name != "posix", reason="Process groups require POSIX")
def test_parent_crash_cleans_descendant_and_next_task_runs(tmp_path):
    pid_file = tmp_path / "child.pid"
    script = (
        "import subprocess,sys,os,pathlib; "
        "p=subprocess.Popen([sys.executable,'-c','import time;time.sleep(30)']); "
        f"pathlib.Path({str(pid_file)!r}).write_text(str(p.pid)); os._exit(9)"
    )
    with pytest.raises(subprocess.CalledProcessError) as failure:
        run_process([sys.executable, "-c", script], 5)
    assert failure.value.returncode == 9
    pid = pid_file.read_text()
    for _ in range(20):
        state = subprocess.run(["ps", "-o", "stat=", "-p", pid], capture_output=True, text=True)
        if not state.stdout.strip() or state.stdout.lstrip().startswith("Z"):
            break
        time.sleep(0.05)
    else:
        pytest.fail("Crashed parent's child is still running")
    assert run_process([sys.executable, "-c", "print('healthy next task')"], 5).strip() == "healthy next task"


@pytest.mark.parametrize("idle", [False, True])
@pytest.mark.parametrize("channel", ["stdout", "stderr"])
def test_excessive_logs_fail_even_when_worker_exits_quickly(monkeypatch, idle, channel):
    monkeypatch.setattr(_process, "MAX_PROCESS_LOG_BYTES", 64 * 1024)
    script = f"import sys,time;sys.{channel}.buffer.write(b'x'*262144);sys.{channel}.flush()"
    if idle:
        script += ";time.sleep(20)"
    with pytest.raises(RuntimeError, match="log output"):
        run_process([sys.executable, "-c", script], 2)
    assert run_process([sys.executable, "-c", "print('healthy after log limit')"], 2).strip() == "healthy after log limit"


def test_exact_log_limit_keeps_existing_return_prefix(monkeypatch):
    monkeypatch.setattr(_process, "MAX_PROCESS_LOG_BYTES", 64 * 1024)
    script = "import sys;sys.stdout.buffer.write(b'x'*65536)"
    assert run_process([sys.executable, "-c", script], 2) == "x" * 4096


@pytest.mark.skipif(os.name != "posix", reason="Process groups require POSIX")
def test_native_log_failure_reaps_descendant_and_preserves_existing_output(tmp_path, monkeypatch):
    output = tmp_path / "output.pdf"
    output.write_bytes(b"original output")
    pid_file = tmp_path / "child.pid"
    monkeypatch.setattr(_process, "MAX_PROCESS_LOG_BYTES", 64 * 1024)
    monkeypatch.setattr(libreoffice.shutil, "which", lambda name: "/fake/soffice")

    def noisy(command, timeout, **kwargs):
        script = (
            "import subprocess,sys,time,pathlib; "
            "p=subprocess.Popen([sys.executable,'-c','import time;time.sleep(20)']); "
            f"pathlib.Path({str(pid_file)!r}).write_text(str(p.pid)); "
            "sys.stdout.buffer.write(b'x'*262144);sys.stdout.flush();time.sleep(20)"
        )
        return run_process([sys.executable, "-c", script], timeout, **kwargs)

    monkeypatch.setattr(libreoffice, "run_process", noisy)
    with pytest.raises(RuntimeError, match="log output"):
        libreoffice.convert_to_pdf(SOURCE, output, timeout=2)
    assert output.read_bytes() == b"original output"
    assert not list(tmp_path.glob(".libreoffice-*"))
    pid = pid_file.read_text()
    for _ in range(20):
        state = subprocess.run(["ps", "-o", "stat=", "-p", pid], capture_output=True, text=True)
        if not state.stdout.strip() or state.stdout.lstrip().startswith("Z"):
            break
        time.sleep(0.05)
    else:
        pytest.fail("Log-limited child's descendant is still running")
