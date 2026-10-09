"""Measure explicit versus implicit default paragraph styles on owned hidden text."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def inputs():
    original = runpy.run_path(str(Path(__file__).with_name('style_toggles.py')))['inputs']
    for name, raw in original():
        if not name.startswith('default-'):
            continue
        for explicit in (False, True):
            output = BytesIO()
            with ZipFile(BytesIO(raw)) as source, ZipFile(output, 'w', ZIP_DEFLATED) as dest:
                for part in source.namelist():
                    data = source.read(part)
                    if part in ('word/document.xml', 'word/styles.xml'):
                        xml = ET.fromstring(data)
                        for node in xml.iter():
                            if node.tag == W + 'b':
                                node.tag = W + 'vanish'
                            elif node.tag == W + 'i':
                                node.tag = W + 'noProof'
                        if explicit and part == 'word/document.xml':
                            ET.SubElement(xml.find('.//' + W + 'pPr'), W + 'pStyle', {W + 'val': 'P'})
                        data = ET.tostring(xml, encoding='utf-8', xml_declaration=True)
                    dest.writestr(ZipInfo(part, (2020, 1, 1, 0, 0, 0)), data, compress_type=ZIP_DEFLATED)
            yield ('explicit-' if explicit else 'implicit-') + name + '.docx', output.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('generate', 'commercial'))
    parser.add_argument('corpus', type=Path)
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    if args.action == 'generate':
        with ZipFile(args.corpus, 'w', ZIP_DEFLATED) as archive:
            for name, raw in inputs():
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), raw, compress_type=ZIP_DEFLATED)
        return
    if args.report is None:
        parser.error('--report is required')
    import aspose.words as aw
    import pymupdf

    assert version('aspose-words') == '26.9.0'
    outputs = args.report.parent / 'corpus/hidden-style-contexts-26.9-outputs.zip'
    records = []
    with ZipFile(args.corpus) as corpus, ZipFile(outputs, 'w', ZIP_DEFLATED) as archive:
        for name in corpus.namelist():
            source = corpus.read(name)
            doc = aw.Document(BytesIO(source))
            paragraph = next(p.as_paragraph() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if p.get_text().strip() == 'IMPORT')
            stream = BytesIO()
            doc.save(stream, aw.SaveFormat.PDF)
            raw = stream.getvalue()
            output = name.replace('.docx', '.pdf')
            archive.writestr(output, raw)
            with pymupdf.open(stream=raw, filetype='pdf') as pdf:
                visible = any('IMPORT' in page.get_text() for page in pdf)
            records.append({'input': name, 'source_sha256': hashlib.sha256(source).hexdigest(),
                            'hidden': paragraph.runs[0].as_run().font.hidden,
                            'output': output, 'output_sha256': hashlib.sha256(raw).hexdigest(),
                            'visible': visible})
    report = {'version': version('aspose-words'), 'python': platform.python_version(),
              'platform': platform.platform(), 'licensed': False,
              'corpus': 'corpus/' + args.corpus.name, 'corpus_sha256': hashlib.sha256(args.corpus.read_bytes()).hexdigest(),
              'outputs': 'corpus/' + outputs.name, 'outputs_sha256': hashlib.sha256(outputs.read_bytes()).hexdigest(),
              'records': records, 'rendering_acceptance': False,
              'scope': '32 explicit/implicit paragraph style references with default P/C declarations; getters and trial PDF target text only'}
    args.report.write_text(json.dumps(report, indent=2) + '\n')
    print(len(records), 'observations')


if __name__ == '__main__':
    main()
