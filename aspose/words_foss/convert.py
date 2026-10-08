"""One-job conversion CLI with wall timeout, memory limits, and atomic publication.

Usage: python -m aspose.words_foss.convert input.docx output.pdf --strict
"""

import argparse
import math
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import warnings

from aspose.words_foss._io import atomic_output
from aspose.words_foss._process import run_process


def _set_limits(memory_mb, timeout):
    if os.name != "posix":
        if memory_mb:
            raise NotImplementedError(
                "Hard memory limits require POSIX; --memory-mb 0 explicitly disables them"
            )
        return
    import resource

    if memory_mb and sys.platform.startswith("linux"):
        requested = memory_mb * 1024 * 1024
        _, hard = resource.getrlimit(resource.RLIMIT_AS)
        limit = requested if hard == resource.RLIM_INFINITY else min(requested, hard)
        resource.setrlimit(resource.RLIMIT_AS, (limit, limit))
    for kind, requested in (
        (resource.RLIMIT_CPU, math.ceil(timeout) + 1),
        (resource.RLIMIT_FSIZE, 256 * 1024 * 1024),
    ):
        _, hard = resource.getrlimit(kind)
        limit = requested if hard == resource.RLIM_INFINITY else min(requested, hard)
        resource.setrlimit(kind, (limit, limit))


def _worker(args):
    _set_limits(args.memory_mb, args.timeout)
    if args.backend == "libreoffice":
        from aspose.words_foss.libreoffice import convert_to_pdf

        convert_to_pdf(args.input, args.output, timeout=args.timeout, _new_session=False)
    else:
        import aspose.words_foss as aw
        from aspose.words_foss.diagnostics import ContentLossWarning
        from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning

        if args.strict:
            for category in (PdfMissingGlyphWarning, ContentLossWarning):
                warnings.simplefilter("error", category)
        options = aw.saving.PdfSaveOptions() if args.fallback_font or args.text_shaping else None
        if options:
            options.fallback_fonts = args.fallback_font
            options.text_shaping = args.text_shaping
        aw.Document(args.input).save(args.output, options)


def _convert(args):
    if os.name != "posix":
        raise NotImplementedError(
            "The bounded CLI requires POSIX process groups; use the library API on Windows"
        )
    output = Path(args.output)
    with atomic_output(output) as temporary:
        with tempfile.TemporaryDirectory(prefix=".conversion-", dir=output.parent) as directory:
            result = Path(directory).resolve() / output.name
            command = [
                sys.executable,
                "-m",
                "aspose.words_foss.convert",
                str(Path(args.input).resolve()),
                str(result),
                "--worker",
                "--backend",
                args.backend,
                "--timeout",
                str(args.timeout),
                "--memory-mb",
                str(args.memory_mb),
            ]
            if args.strict:
                command.append("--strict")
            if args.text_shaping:
                command.append("--text-shaping")
            for font in args.fallback_font:
                command.extend(("--fallback-font", str(Path(font).resolve())))
            log = run_process(command, args.timeout, memory_mb=args.memory_mb)
            if log:
                sys.stderr.write(log)
            if not result.is_file():
                raise RuntimeError("Worker produced no output")
            os.replace(result, temporary)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--backend", choices=("builtin", "libreoffice"), default="builtin")
    parser.add_argument("--timeout", type=float, default=60)
    parser.add_argument(
        "--memory-mb",
        type=int,
        default=1024,
        help="RSS watchdog threshold; Linux also uses RLIMIT_AS; 0 disables memory checks",
    )
    parser.add_argument(
        "--strict", action="store_true", help="Fail on missing glyphs and known load/write loss (not a full fidelity validator)"
    )
    parser.add_argument("--fallback-font", action="append", default=[])
    parser.add_argument("--text-shaping", action="store_true", help="Enable fpdf2/HarfBuzz shaping; install [shaping]")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if not math.isfinite(args.timeout) or args.timeout <= 0 or args.memory_mb < 0:
        parser.error("timeout must be positive/finite and memory-mb must be nonnegative")
    if args.backend == "libreoffice" and (args.strict or args.fallback_font or args.text_shaping):
        parser.error("strict checks, fallback-font and text-shaping require the builtin backend")
    if (args.fallback_font or args.text_shaping) and Path(args.output).suffix.lower() != ".pdf":
        parser.error("fallback-font and text-shaping only apply to PDF output")
    try:
        if args.worker:
            _worker(args)
        else:
            _convert(args)
    except (
        OSError,
        ValueError,
        RuntimeError,
        MemoryError,
        Warning,
        subprocess.SubprocessError,
    ) as exc:
        parser.exit(1, "Conversion failed: " + (getattr(exc, "output", None) or str(exc)) + "\n")


if __name__ == "__main__":
    main()
