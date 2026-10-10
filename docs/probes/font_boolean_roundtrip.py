"""Reuse 810 owned inputs for Boolean font saves and fixed native cold getters."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def inputs(root):
    report = json.loads((root / 'font-boolean-contexts-26.9.json').read_text())
    with ZipFile(root / report['corpus']) as archive:
        for row in report['records']:
            if row['field'] != 'hidden':
                yield row, archive.read(row['input'])


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
        stored = {}
        with ZipFile(args.outputs, 'w', ZIP_DEFLATED) as archive:
            for row, raw in inputs(args.root):
                model = Document(BytesIO(raw)).light_document_model
                for phase in ('direct', 'json'):
                    current = model if phase == 'direct' else ldm.Document.model_validate_json(model.model_dump_json())
                    for label, fmt in (('docx', SaveFormat.DOCX), ('flat_opc', SaveFormat.FLAT_OPC)):
                        saved = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
                        name = phase + '/' + label + '/' + row['input'].removesuffix('.docx') + ('.docx' if label == 'docx' else '.xml')
                        digest = hashlib.sha256(saved).hexdigest()
                        if digest not in stored:
                            archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), saved, compress_type=ZIP_DEFLATED)
                            stored[digest] = name
                        records.append({'input': row['input'], 'field': row['field'], 'phase': phase, 'format': label,
                                        'expected': row['value'], 'source_sha256': hashlib.sha256(raw).hexdigest(),
                                        'output': name, 'storage': stored[digest], 'output_sha256': digest})
        assert len(records) == 3240
        report = {'version': version('aspose-words-foss-enhanced'), 'python': platform.python_version(),
                  'platform': platform.platform(), 'outputs': 'corpus/' + args.outputs.name,
                  'outputs_sha256': hashlib.sha256(args.outputs.read_bytes()).hexdigest(), 'records': records,
                  'full_format_acceptance': False, 'rendering_acceptance': False,
                  'scope': '10 Boolean properties; 810 owned category inputs; direct/model JSON to DOCX/Flat OPC; cold getters and direct declaration preservation; no full Font API or visual parity acceptance'}
        args.report.write_text(json.dumps(report, indent=2) + '\n')
        print(len(records), 'saved outputs')
        return
    import aspose.words as aw

    assert version('aspose-words') == '26.9.0'
    report = json.loads(args.report.read_text())
    observed = []
    with ZipFile(args.outputs) as archive:
        for row in report['records']:
            raw = archive.read(row.get('storage', row['output']))
            assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
            document = aw.Document(BytesIO(raw))
            paragraph = next(p.as_paragraph() for p in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if p.get_text().strip() == 'IMPORT')
            actual = getattr(paragraph.runs[0].as_run().font, row['field'])
            observed.append({'output': row['output'], 'value': actual})
    assert len(observed) == 3240
    report['native_reread'] = {'version': version('aspose-words'), 'licensed': False,
                               'python': platform.python_version(), 'platform': platform.platform(), 'records': observed}
    args.report.write_text(json.dumps(report, indent=2) + '\n')
    expected = {r['output']: r['expected'] for r in report['records']}
    print(sum(r['value'] is not expected[r['output']] for r in observed), 'native getter differences')


if __name__ == '__main__':
    main()
