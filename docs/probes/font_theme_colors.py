"""Owned theme inputs: distinguish native color getters from PDF target colors."""

import argparse
import hashlib
import itertools
import json
import platform
import runpy
from collections import Counter
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def inputs():
    make = runpy.run_path(str(Path(__file__).with_name('style_toggles.py')))['document']
    modifiers = ('', ' w:themeTint="80"', ' w:themeShade="80"', ' w:themeTint="00"',
                 ' w:themeTint="FF"', ' w:themeShade="00"', ' w:themeShade="FF"',
                 ' w:themeTint="80" w:themeShade="80"')
    for index, (base, fallback, modifier) in enumerate(itertools.product(
            ('123456', 'C0504D'), ('auto', 'FF0000'), modifiers)):
        raw = make('', '', f'<w:color w:val="{fallback}" w:themeColor="accent1"{modifier}/>')
        with ZipFile(BytesIO(raw)) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
            b'</Relationships>', b'<Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
        parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
            b'</Types>', b'<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
        colors = {'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '112233', 'lt2': 'EEEEEE',
                  **{f'accent{i}': base for i in range(1, 7)}, 'hlink': '0000FF', 'folHlink': '800080'}
        scheme = ''.join(f'<a:{name}><a:srgbClr val="{value}"/></a:{name}>' for name, value in colors.items())
        fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        matrix = ('<a:fillStyleLst>' + fill * 3 + '</a:fillStyleLst><a:lnStyleLst>'
                  + ('<a:ln w="9525">' + fill + '</a:ln>') * 3 + '</a:lnStyleLst>'
                  + '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3
                  + '</a:effectStyleLst><a:bgFillStyleLst>' + fill * 3 + '</a:bgFillStyleLst>')
        parts['word/theme/theme1.xml'] = (
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Owned">'
            '<a:themeElements><a:clrScheme name="Owned">' + scheme + '</a:clrScheme>'
            '<a:fontScheme name="Owned"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '<a:fmtScheme name="Owned">' + matrix + '</a:fmtScheme></a:themeElements></a:theme>').encode()
        output = BytesIO()
        with ZipFile(output, 'w') as archive:
            for name, value in parts.items():
                info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, value)
        yield f'{index:02d}.docx', {'base': base, 'fallback': fallback, 'modifiers': modifier}, output.getvalue()


def pdf_observation(raw):
    """Measure extraction color and modal painted RGB in the owned target box."""
    import pymupdf
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        page = next(p for p in pdf if any(w[4] == 'IMPORT' for w in p.get_text('words')))
        word = next(w for w in page.get_text('words') if w[4] == 'IMPORT')
        colors = [s['color'] for b in page.get_text('dict')['blocks'] for line in b.get('lines', [])
                  for s in line['spans'] if s['text'] == 'IMPORT']
        assert len(colors) == 1
        pixels = page.get_pixmap(matrix=pymupdf.Matrix(4, 4), clip=pymupdf.Rect(word[:4]), alpha=False,
                                colorspace=pymupdf.csRGB)
        counts = Counter(zip(*(iter(pixels.samples),) * 3))
        counts.pop((255, 255, 255), None)
        rgb = counts.most_common(1)[0][0] if counts else (255, 255, 255)
        return {'pdf_target_color': colors[0], 'painted_target_rgb': list(rgb),
                'nonwhite_target_pixels': sum(counts.values())}


def main(input_generator=inputs, expected_count=32):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('generate', 'commercial', 'current'))
    parser.add_argument('corpus', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--outputs', type=Path)
    args = parser.parse_args()
    if args.action == 'generate':
        with ZipFile(args.corpus, 'w', ZIP_DEFLATED) as archive:
            for name, _, raw in input_generator():
                archive.writestr(name, raw)
        return
    if args.outputs is None:
        parser.error('observations require --outputs for raw PDFs')
    native = args.action == 'commercial'
    if native:
        import aspose.words as aw
        assert version('aspose-words') == '26.9.0'
    else:
        from aspose.words_foss import Document, SaveFormat
    observe = runpy.run_path(str(Path(__file__).with_name('font_color_categories.py')))['observed']
    records = []
    with ZipFile(args.corpus) as sources, ZipFile(args.outputs, 'w', ZIP_DEFLATED) as outputs:
        for name, case, _ in input_generator():
            raw = sources.read(name)
            doc = aw.Document(BytesIO(raw)) if native else Document(BytesIO(raw))
            getter = observe(doc, native)
            if native:
                result = BytesIO()
                doc.save(result, aw.SaveFormat.PDF)
                pdf_raw = result.getvalue()
            else:
                pdf_raw = doc.to_bytes(SaveFormat.PDF)
            paint = pdf_observation(pdf_raw)
            output = name.replace('.docx', '.pdf')
            outputs.writestr(output, pdf_raw)
            records.append({'input': name, **case, 'sha256': hashlib.sha256(raw).hexdigest(),
                            'getter': getter, **paint, 'output': output,
                            'output_sha256': hashlib.sha256(pdf_raw).hexdigest()})
    assert len(records) == expected_count
    args.report.write_text(json.dumps({'version': version('aspose-words' if native else 'aspose-words-foss-enhanced'),
                                      'python': platform.python_version(), 'platform': platform.platform(),
                                      'licensed': False if native else None, 'records': records,
                                      'full_Font_acceptance': False, 'rendering_acceptance': False}, indent=2) + '\n')
    print(len(records))


if __name__ == '__main__':
    main()
