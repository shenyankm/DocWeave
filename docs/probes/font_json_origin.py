"""Reuse the owned toggle corpus for model JSON and native saved-file rereads."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


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

        source = json.loads((args.root / 'style-toggles-26.9.json').read_text())
        corpus = args.root / 'corpus/style-toggles-26.9.zip'
        assert hashlib.sha256(corpus.read_bytes()).hexdigest() == source['corpus_sha256']
        records = []
        previous = json.loads(args.report.read_text()) if args.report.exists() else {}
        with ZipFile(corpus) as inputs, ZipFile(args.outputs, 'w', compression=ZIP_DEFLATED) as outputs:
            for row in source['records']:
                raw = inputs.read(row['input'])
                assert hashlib.sha256(raw).hexdigest() == row['sha256']
                model = Document(BytesIO(raw)).light_document_model
                model = ldm.Document.model_validate_json(model.model_dump_json())
                for name, format in (('docx', SaveFormat.DOCX), ('flat_opc', SaveFormat.FLAT_OPC)):
                    data = LdmDocxWriter(OoxmlSaveOptions(format)).write_to_bytes(model)
                    filename = row['input'].removesuffix('.docx') + '/' + name + ('.docx' if name == 'docx' else '.xml')
                    outputs.writestr(ZipInfo(filename, (2020, 1, 1, 0, 0, 0)), data, compress_type=ZIP_DEFLATED)
                    records.append({'input': row['input'], 'source_sha256': row['sha256'],
                                    'format': name, 'output': filename,
                                    'output_sha256': hashlib.sha256(data).hexdigest(),
                                    'expected_bold': row['bold'], 'expected_italic': row['italic']})
        assert len(records) == 1490
        report = {'version': version('aspose-words-foss-enhanced'), 'python': platform.python_version(),
                  'platform': platform.platform(), 'source_report': 'style-toggles-26.9.json',
                  'source_corpus_sha256': source['corpus_sha256'], 'outputs': 'corpus/' + args.outputs.name,
                  'outputs_sha256': hashlib.sha256(args.outputs.read_bytes()).hexdigest(),
                  'records': records, 'full_format_acceptance': False, 'rendering_acceptance': False,
                  'scope': '745 reused ordinary b/i inheritance inputs; model JSON then DOCX/Flat OPC; '
                           'native cold getters and direct Run/character b/i declarations; '
                           'not complete font editing or rendering acceptance'}
        if 'before_cascade_fix' in previous:
            report['before_cascade_fix'] = previous['before_cascade_fix']
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        print(json.dumps({'outputs': len(records)}))
        return
    import aspose.words as aw

    report = json.loads(args.report.read_text())
    observed = []
    with ZipFile(args.outputs) as archive:
        for row in report['records']:
            raw = archive.read(row['output'])
            assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
            doc = aw.Document(BytesIO(raw))
            paragraph = next(p.as_paragraph() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if p.get_text().strip() == 'IMPORT')
            font = paragraph.runs[0].as_run().font
            observed.append({'output': row['output'], 'bold': font.bold, 'italic': font.italic})
    report['native_reread'] = {'version': version('aspose-words'), 'licensed': False,
                               'python': platform.python_version(), 'platform': platform.platform(), 'records': observed}
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'native_rereads': len(observed)}))


if __name__ == '__main__':
    main()
