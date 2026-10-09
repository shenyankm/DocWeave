"""Reuse owned size baselines to verify default presence through JSON and OOXML."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

SOURCES = ('font-defaults-26.9.json', 'font-default-matrix-26.9.json', 'font-size-loading-26.9.json')


def cases(root):
    for filename in SOURCES:
        report = json.loads((root / filename).read_text())
        with ZipFile(root / report['corpus']) as archive:
            for row in report['records']:
                expected = row.get('loaded', row)
                if 'run_size' not in expected or not 0 < expected['run_size'] < 1000:
                    continue
                raw = archive.read(row['input'])
                assert hashlib.sha256(raw).hexdigest() == row['sha256']
                yield filename, row, raw, expected


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
        with ZipFile(args.outputs, 'w', compression=ZIP_DEFLATED) as archive:
            for filename, row, raw, expected in cases(args.root):
                original = Document(BytesIO(raw)).light_document_model
                for format, enum in (('docx', SaveFormat.DOCX), ('flat_opc', SaveFormat.FLAT_OPC)):
                    for use_json in (False, True):
                        model = ldm.Document.model_validate_json(original.model_dump_json()) if use_json else original
                        output = LdmDocxWriter(OoxmlSaveOptions(enum)).write_to_bytes(model)
                        name = (filename.removesuffix('.json') + '/' + row['input'].removesuffix('.docx') +
                                f'/{format}-{int(use_json)}.' + ('docx' if format == 'docx' else 'xml'))
                        archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), output, compress_type=ZIP_DEFLATED)
                        records.append({'source_report': filename, 'input': row['input'], 'source_sha256': row['sha256'],
                                        'format': format, 'json_roundtrip': use_json, 'output': name,
                                        'output_sha256': hashlib.sha256(output).hexdigest(),
                                        'rpr_default_present': model.doc_defaults_rpr_present,
                                        'expected_run_size': expected['run_size'], 'expected_style_size': expected['style_size']})
        report = {'version': version('aspose-words-foss-enhanced'), 'python': platform.python_version(),
                  'platform': platform.platform(), 'sources': list(SOURCES),
                  'outputs': 'corpus/' + args.outputs.name,
                  'outputs_sha256': hashlib.sha256(args.outputs.read_bytes()).hexdigest(), 'records': records,
                  'full_format_acceptance': False, 'rendering_acceptance': False,
                  'remaining_limits': ['Legacy model JSON without presence/origin fields cannot restore source provenance.',
                                       'Character styles with undeclared bold/italic can acquire explicit false values after full model JSON; this stage accepts ordinary size origin only.',
                                       'Complex-script size, complete font-origin semantics, malformed native recovery and full rendering remain unaccepted.'],
                  'scope': '87 reused ordinary-size inputs; DOCX/Flat OPC with/without model JSON; '
                           'default-group presence and ordinary Run/Derived sizes; complex-script and full layout unaccepted'}
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
            paragraph = next(node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if node.get_text().strip() == 'IMPORT')
            observed.append({'output': row['output'], 'run_size': paragraph.runs[0].as_run().font.size,
                             'style_size': doc.styles.get_by_name('Derived').font.size})
    report['native_reread'] = {'version': version('aspose-words'), 'licensed': False,
                               'python': platform.python_version(), 'platform': platform.platform(), 'records': observed}
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'native_rereads': len(observed)}))


if __name__ == '__main__':
    main()
