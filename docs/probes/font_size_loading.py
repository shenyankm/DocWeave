"""Owned ordinary font-size inputs and official load/save/cold-read observations."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

VALUES = ('24', '24.0', '25.5', '+24', ' 24 ', '1e2', '12pt', '6.25pt', '0.5in',
          '2.54cm', '25.4mm', '1pc', '1pi', '24hp', '24twip', '0', '-2',
          'PRIVATE FONT VALUE', '2_4', '٢٤', '', 'NaN', 'Infinity', '9' * 400,
          '18446744073709551615', '+24.0', '24px', '1.13', '23.99', '6.13pt',
          '6.125pt', '0.01pt', '24PT', '-1pt', '+12pt', ' 12pt ', '24.99', '25.01',
          '22.5', '1e-2', '1.5in', '12 pt')


def inputs():
    helpers = runpy.run_path(str(Path(__file__).with_name('import_style_conflicts.py')))
    for scope in ('run', 'style', 'default'):
        for index, value in enumerate(VALUES):
            encoded = escape(value, {'"': '&quot;'})
            size = f'<w:sz w:val="{encoded}"/>'
            defaults = ('<w:docDefaults><w:rPrDefault><w:rPr>' + size +
                        '</w:rPr></w:rPrDefault></w:docDefaults>') if scope == 'default' else ''
            style = ('<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/>'
                     '<w:rPr>' + (size if scope == 'style' else '') + '</w:rPr></w:style>')
            data = helpers['document'](defaults + style, '<w:pStyle w:val="Derived"/>',
                                       size if scope == 'run' else '')
            yield f'{scope}-{index:02d}.docx', scope, value, data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('generate', 'commercial', 'current', 'commercial-read'))
    parser.add_argument('corpus', type=Path)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--outputs', type=Path)
    args = parser.parse_args()
    if args.action == 'generate':
        with ZipFile(args.corpus, 'w', compression=ZIP_DEFLATED) as archive:
            for name, _, _, data in inputs():
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), data, compress_type=ZIP_DEFLATED)
        return
    if args.output is None or args.outputs is None:
        parser.error('--output and --outputs are required')
    if args.action == 'current':
        from aspose.words_foss import Document

        report = json.loads(args.output.read_text())
        records = []
        with ZipFile(args.corpus) as source, ZipFile(args.outputs, 'w', compression=ZIP_DEFLATED) as outputs:
            for row in report['records']:
                if 'loaded' not in row or not 0 < row['loaded']['run_size'] < 1000:
                    continue
                raw = Document(BytesIO(source.read(row['input']))).to_bytes('docx')
                outputs.writestr(ZipInfo(row['input'], (2020, 1, 1, 0, 0, 0)), raw, compress_type=ZIP_DEFLATED)
                records.append({'input': row['input'], 'output_sha256': hashlib.sha256(raw).hexdigest()})
        report['docweave_roundtrip'] = {'version': version('aspose-words-foss-enhanced'),
                                       'python': platform.python_version(), 'platform': platform.platform(),
                                       'outputs': 'corpus/' + args.outputs.name,
                                       'outputs_sha256': hashlib.sha256(args.outputs.read_bytes()).hexdigest(),
                                       'records': records, 'full_acceptance': False}
        args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        print(json.dumps({'docweave_outputs': len(records)}))
        return
    import aspose.words as aw

    def snapshot(document):
        paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                         if node.get_text().strip() == 'IMPORT')
        return {'run_size': paragraph.runs[0].as_run().font.size,
                'style_size': document.styles.get_by_name('Derived').font.size}

    if args.action == 'commercial-read':
        report = json.loads(args.output.read_text())
        observed = []
        with ZipFile(args.outputs) as archive:
            for row in report['docweave_roundtrip']['records']:
                raw = archive.read(row['input'])
                assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
                observed.append({'input': row['input'], **snapshot(aw.Document(BytesIO(raw)))})
        report['docweave_roundtrip']['native_reread'] = {
            'version': version('aspose-words'), 'python': platform.python_version(),
            'platform': platform.platform(), 'licensed': False, 'records': observed}
        args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        print(json.dumps({'native_rereads': len(observed)}))
        return
    records = []
    with ZipFile(args.corpus) as source, ZipFile(args.outputs, 'w', compression=ZIP_DEFLATED) as outputs:
        for name, scope, value, _ in inputs():
            data = source.read(name)
            row = {'input': name, 'scope': scope, 'value': value, 'sha256': hashlib.sha256(data).hexdigest()}
            try:
                document = aw.Document(BytesIO(data))
                row['loaded'] = snapshot(document)
            except (RuntimeError, ValueError, OverflowError) as error:
                row['load_error'] = type(error).__name__
                records.append(row)
                continue
            stream = BytesIO()
            try:
                document.save(stream, aw.SaveFormat.DOCX)
                raw = stream.getvalue()
                row['after_save'] = snapshot(document)
                row['cold'] = snapshot(aw.Document(BytesIO(raw)))
                row['output'] = name
                row['output_sha256'] = hashlib.sha256(raw).hexdigest()
                outputs.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), raw, compress_type=ZIP_DEFLATED)
            except (RuntimeError, ValueError, OverflowError) as error:
                row['save_error'] = type(error).__name__
            records.append(row)
    report = {'module': 'aspose.words', 'version': version('aspose-words'), 'licensed': False,
              'python': platform.python_version(), 'platform': platform.platform(),
              'corpus': 'corpus/' + args.corpus.name,
              'corpus_sha256': hashlib.sha256(args.corpus.read_bytes()).hexdigest(),
              'outputs': 'corpus/' + args.outputs.name,
              'outputs_sha256': hashlib.sha256(args.outputs.read_bytes()).hexdigest(),
              'records': records, 'rendering_acceptance': False, 'full_format_acceptance': False,
              'scope': '42 ordinary w:sz values across direct, paragraph style and document defaults; '
                       'trial watermarks retained in saved outputs; complex-script size and rendering unaccepted',
              'references': ['https://reference.aspose.com/words/python-net/aspose.words/font/size/',
                             'https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.hpsmeasuretype.val']}
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'inputs': len(records), 'outputs': sum('output' in row for row in records)}))


if __name__ == '__main__':
    main()
