"""Cold-process DOCX -> PDF benchmark; JSON output, no runtime instrumentation or caches.

Usage: python scripts/benchmark.py --repeat 3 > benchmark.json
Each sample runs in a new process. RSS is unavailable on Windows (reported as null).
"""

import argparse
from io import BytesIO
import json
from pathlib import Path
from statistics import median
import subprocess
import sys
import tempfile
from time import perf_counter

CASES = ("small", "hundred_pages", "image_table", "repeated_images", "mixed_styles")


def fixtures(directory):
    from PIL import Image
    from aspose.words_foss import light_document_model as ldm
    from aspose.words_foss.docx_writer import LdmDocxWriter

    def para(text, **font):
        return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(**font))])

    data = BytesIO()
    Image.new("RGB", (240, 120), "navy").save(data, format="PNG")

    def image():
        return ldm.Shape(has_image=True, is_inline=True, width=100, height=50,
            image_data=ldm.ImageData(image_bytes=data.getvalue(), image_type=ldm.ImageData.from_mime("image/png")))

    pages = []
    for i in range(100):
        if i:
            pages.append(para("\f"))
        pages.extend(para(f"第{i + 1}页：中文 document conversion " * 8) for _ in range(8))
    table = ldm.Table(rows=[ldm.Row(cells=[
        ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(text=f"图{i}"), image()])]),
        ldm.Cell(paragraphs=[para("中文表格内容 " * 10, bold=True)]),
    ]) for i in range(24)])
    bodies = {
        "small": [para("一页中文转换基准 mixed Latin 123。")],
        "hundred_pages": pages,
        "image_table": [table],
        "repeated_images": [ldm.Paragraph(children=[image()]) for _ in range(40)],
        "mixed_styles": [para("中文 Mixed styles " * 8, size=11 + i % 3, bold=bool(i % 2),
                              italic=bool(i % 3), highlight_color="Color [Yellow]") for i in range(100)],
    }
    paths = {}
    for name, children in bodies.items():
        path = directory / f"{name}.docx"
        LdmDocxWriter().write(ldm.Document(sections=[ldm.Section(body=ldm.Body(children=children))]), path)
        paths[name] = path
    return paths


def measure(path):
    start = perf_counter()
    import aspose.words_foss as aw
    from aspose.words_foss._io import atomic_output
    from aspose.words_foss.pdf_writer import LdmPdfWriter

    imported = perf_counter()
    document = aw.Document(path)
    parsed = perf_counter()
    pdf = LdmPdfWriter()._render_pdf(document.light_document_model)
    rendered = perf_counter()
    payload = pdf.output()
    serialized = perf_counter()
    with tempfile.TemporaryDirectory() as directory:
        with atomic_output(Path(directory) / "result.pdf") as output:
            output.write_bytes(payload)
    written = perf_counter()
    try:
        import resource
        rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        rss *= 1 if sys.platform == "darwin" else 1024
    except ImportError:
        rss = None
    result = {
        "import_ms": (imported - start) * 1000,
        "parse_ms": (parsed - imported) * 1000,
        "layout_ms": (rendered - parsed) * 1000,
        "serialize_ms": (serialized - rendered) * 1000,
        "write_ms": (written - serialized) * 1000,
        "total_ms": (written - start) * 1000,
        "peak_rss_bytes": rss, "output_bytes": len(payload), "pages": pdf.pages_count,
    }
    assert result["output_bytes"] > 0 and result["pages"] > 0
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--case", choices=CASES)
    parser.add_argument("--worker", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.repeat < 1:
        parser.error("repeat must be positive")
    if args.worker:
        print(json.dumps(measure(args.worker)))
        return
    results = {}
    with tempfile.TemporaryDirectory() as directory:
        paths = fixtures(Path(directory))
        for name in ([args.case] if args.case else CASES):
            samples = []
            for _ in range(args.repeat):
                start = perf_counter()
                run = subprocess.run([sys.executable, str(Path(__file__).resolve()), "--worker", str(paths[name])],
                                     capture_output=True, text=True, timeout=300)
                if run.returncode:
                    raise RuntimeError(run.stderr or "Benchmark worker failed")
                sample = json.loads(run.stdout)
                sample["process_ms"] = (perf_counter() - start) * 1000
                samples.append(sample)
            results[name] = {key: median(sample[key] for sample in samples)
                             if samples[0][key] is not None else None for key in samples[0]}
    print(json.dumps({"python": sys.version, "platform": sys.platform,
                      "repeat": args.repeat, "median": results}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
