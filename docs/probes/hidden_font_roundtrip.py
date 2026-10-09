"""Reuse 113 owned inputs for hidden-font DOCX/Flat OPC saves and native cold reads."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def inputs(root):
    pdf = json.loads((root / 'font-hidden-rendering-26.9.json').read_text())
    visible = {row['input']: row['import_visible'] for row in pdf['records']}
    for filename in ('font-boolean-contexts-26.9.json', 'hidden-style-contexts-26.9.json'):
        report = json.loads((root / filename).read_text())
        with ZipFile(root / report['corpus']) as archive:
            for row in report['records']:
                if row.get('field', 'hidden') != 'hidden':
                    continue
                yield row['input'], archive.read(row['input']), row.get('value', row.get('hidden')), visible.get(row['input'], row.get('visible'))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('current', 'commercial-read'))
    parser.add_argument('root', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--outputs', type=Path, required=True)
    args = parser.parse_args()
    if args.action == 'current':
        from aspose.words_foss import Document, SaveFormat
        from aspose.words_foss import light_document_model as ldm
        from aspose.words_foss.docx_writer import LdmDocxWriter
        from aspose.words_foss.saving import OoxmlSaveOptions

        records = []
        with ZipFile(args.outputs, 'w', ZIP_DEFLATED) as outputs:
            for name, source, hidden, visible in inputs(args.root):
                model = Document(BytesIO(source)).light_document_model
                for phase in ('direct', 'json'):
                    current = model if phase == 'direct' else ldm.Document.model_validate_json(model.model_dump_json())
                    for label, fmt in (('docx', SaveFormat.DOCX), ('flat_opc', SaveFormat.FLAT_OPC)):
                        raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
                        output = phase + '/' + label + '/' + name.removesuffix('.docx') + ('.docx' if label == 'docx' else '.xml')
                        outputs.writestr(ZipInfo(output, (2020, 1, 1, 0, 0, 0)), raw, compress_type=ZIP_DEFLATED)
                        records.append({'input': name, 'source_sha256': hashlib.sha256(source).hexdigest(),
                                        'phase': phase, 'format': label, 'output': output,
                                        'output_sha256': hashlib.sha256(raw).hexdigest(),
                                        'expected_hidden': hidden, 'expected_visible': visible})
        assert len(records) == 452
        report = {'version': version('aspose-words-foss-enhanced'), 'python': platform.python_version(),
                  'platform': platform.platform(), 'outputs': 'corpus/' + args.outputs.name,
                  'outputs_sha256': hashlib.sha256(args.outputs.read_bytes()).hexdigest(), 'records': records,
                  'full_format_acceptance': False, 'rendering_acceptance': False,
                  'scope': '113 reused inputs; hidden getter, target PDF text visibility and direct declaration/reference preservation; native cold PDFs inspected live, not archived; no whole-page visual acceptance'}
        args.report.write_text(json.dumps(report, indent=2) + '\n')
        print(len(records), 'outputs')
        return
    import aspose.words as aw
    import pymupdf

    assert version('aspose-words') == '26.9.0'
    report = json.loads(args.report.read_text())
    observed = []
    with ZipFile(args.outputs) as archive:
        for row in report['records']:
            raw = archive.read(row['output'])
            assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
            doc = aw.Document(BytesIO(raw))
            paragraph = next(p.as_paragraph() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if p.get_text().strip() == 'IMPORT')
            hidden = paragraph.runs[0].as_run().font.hidden
            stream = BytesIO()
            doc.save(stream, aw.SaveFormat.PDF)
            with pymupdf.open(stream=stream.getvalue(), filetype='pdf') as pdf:
                visible = any('IMPORT' in page.get_text() for page in pdf)
                if visible:
                    assert any(any(pixel < 200 for pixel in page.get_pixmap(
                        clip=pymupdf.Rect(word[:4]), matrix=pymupdf.Matrix(2, 2), alpha=False).samples)
                        for page in pdf for word in page.get_text('words') if word[4] == 'IMPORT')
            observed.append({'output': row['output'], 'hidden': hidden, 'visible': visible})
    report['native_reread'] = {'version': version('aspose-words'), 'licensed': False,
                               'python': platform.python_version(), 'platform': platform.platform(), 'records': observed,
                               'visible_target_pixels_checked': True}
    args.report.write_text(json.dumps(report, indent=2) + '\n')
    print(len(observed), 'native cold reads')


if __name__ == '__main__':
    main()
