"""Owned font Boolean category inputs and fixed-version native getter observations."""

import argparse
import hashlib
import itertools
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

FIELDS = {'all_caps': 'caps', 'small_caps': 'smallCaps', 'strike_through': 'strike',
          'hidden': 'vanish', 'outline': 'outline', 'shadow': 'shadow', 'emboss': 'emboss',
          'engrave': 'imprint', 'no_proofing': 'noProof', 'bold_bi': 'bCs', 'italic_bi': 'iCs'}


def inputs():
    document = runpy.run_path(str(Path(__file__).with_name('style_toggles.py')))['document']
    for field, tag in FIELDS.items():
        for values in itertools.product((None, False, True), repeat=4):
            default, paragraph, character, direct = values
            def properties(value, tag=tag):
                return '' if value is None else f'<w:{tag} w:val="{int(value)}"/>'
            styles = (f'<w:docDefaults><w:rPrDefault><w:rPr>{properties(default)}</w:rPr></w:rPrDefault></w:docDefaults>' +
                      f'<w:style w:type="paragraph" w:styleId="P"><w:name w:val="P"/><w:rPr>{properties(paragraph)}</w:rPr></w:style>' +
                      f'<w:style w:type="character" w:styleId="C"><w:name w:val="C"/><w:rPr>{properties(character)}</w:rPr></w:style>')
            name = field + '-' + '-'.join('n' if v is None else str(int(v)) for v in values) + '.docx'
            yield name, field, values, document(styles, '<w:pStyle w:val="P"/>',
                                                 '<w:rStyle w:val="C"/>' + properties(direct))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('generate', 'commercial', 'commercial-pdf', 'current'))
    parser.add_argument('corpus', type=Path)
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    generated = {name: (field, values, raw) for name, field, values, raw in inputs()}
    assert len(generated) == 891
    if args.action == 'generate':
        with ZipFile(args.corpus, 'w', compression=ZIP_DEFLATED) as archive:
            for name, (_, _, raw) in generated.items():
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), raw, compress_type=ZIP_DEFLATED)
        print(json.dumps({'inputs': len(generated)}))
        return
    if args.report is None:
        parser.error('--report is required')
    commercial = args.action.startswith('commercial')
    if commercial:
        import aspose.words as aw
    else:
        import aspose.words_foss as aw
    records = []
    if args.action == 'commercial-pdf':
        import pymupdf

        outputs = args.report.parent / 'corpus/font-hidden-rendering-26.9.zip'
        with ZipFile(args.corpus) as corpus, ZipFile(outputs, 'w', ZIP_DEFLATED) as archive:
            for name, (field, values, _) in generated.items():
                if field != 'hidden':
                    continue
                source = corpus.read(name)
                doc = aw.Document(BytesIO(source))
                stream = BytesIO()
                doc.save(stream, aw.SaveFormat.PDF)
                raw = stream.getvalue()
                output = name.replace('.docx', '.pdf')
                archive.writestr(output, raw)
                with pymupdf.open(stream=raw, filetype='pdf') as pdf:
                    found = any('IMPORT' in page.get_text() for page in pdf)
                records.append({'input': name, 'output': output, 'sha256': hashlib.sha256(raw).hexdigest(),
                                'import_visible': found})
        report = {'version': version('aspose-words'), 'python': platform.python_version(),
                  'platform': platform.platform(), 'licensed': False,
                  'outputs': 'corpus/' + outputs.name,
                  'outputs_sha256': hashlib.sha256(outputs.read_bytes()).hexdigest(),
                  'records': records, 'rendering_acceptance': False,
                  'scope': '81 owned hidden-font inputs; native trial PDF target text presence only; watermark retained; no whole-page visual parity acceptance'}
        args.report.write_text(json.dumps(report, indent=2) + '\n')
        print(json.dumps({'native_pdf_observations': len(records)}))
        return
    with ZipFile(args.corpus) as archive:
        for name, (field, values, _) in generated.items():
            raw = archive.read(name)
            doc = aw.Document(BytesIO(raw))
            if commercial:
                paragraph = next(p.as_paragraph() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                                 if p.get_text().strip() == 'IMPORT')
                font = paragraph.runs[0].as_run().font
            else:
                font = doc.light_document_model.sections[0].body.paragraphs[0].runs[0].font
            records.append({'input': name, 'field': field, 'values': values,
                            'sha256': hashlib.sha256(raw).hexdigest(), 'value': getattr(font, field)})
    report = {'version': version('aspose-words' if commercial else 'aspose-words-foss-enhanced'),
              'python': platform.python_version(), 'platform': platform.platform(), 'licensed': False if commercial else None,
              'corpus': 'corpus/' + args.corpus.name, 'corpus_sha256': hashlib.sha256(args.corpus.read_bytes()).hexdigest(),
              'records': records, 'full_format_acceptance': False, 'rendering_acceptance': False,
              'scope': '11 font Boolean properties; 891 docDefaults/paragraph/character/direct inputs; getter observations only'}
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'observations': len(records)}))


if __name__ == '__main__':
    main()
