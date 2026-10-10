"""Observe four-layer RGB/automatic colors and cold-read conversion outputs."""

import argparse
import hashlib
import itertools
import json
import platform
import re
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def inputs():
    make = runpy.run_path(str(Path(__file__).with_name('style_toggles.py')))['document']
    for values in itertools.product((None, 'auto', 'FF0000', '0000FF'), repeat=4):
        def prop(val):
            return '' if val is None else f'<w:color w:val="{val}"/>'
        default, p, c, direct = values
        styles = ('<w:docDefaults><w:rPrDefault><w:rPr>' + prop(default) + '</w:rPr></w:rPrDefault></w:docDefaults>'
                  + '<w:style w:type="paragraph" w:styleId="P"><w:name w:val="P"/><w:rPr>' + prop(p) + '</w:rPr></w:style>'
                  + '<w:style w:type="character" w:styleId="C"><w:name w:val="C"/><w:rPr>' + prop(c) + '</w:rPr></w:style>')
        name = '-'.join('n' if v is None else v for v in values) + '.docx'
        yield name, list(values), make(styles, '<w:pStyle w:val="P"/>', '<w:rStyle w:val="C"/>' + prop(direct))


def observed(doc, native):
    if native:
        import aspose.words as aw
        p = next(p.as_paragraph() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                 if p.get_text().strip() == 'IMPORT')
        color = p.runs[0].as_run().font.color
        return {key: getattr(color, key) for key in ('a', 'r', 'g', 'b', 'is_empty')}
    color = doc.light_document_model.sections[0].body.paragraphs[0].runs[0].font.color
    if color == 'Color [Empty]':
        return {'a': 0, 'r': 0, 'g': 0, 'b': 0, 'is_empty': True}
    return dict(zip(('a', 'r', 'g', 'b'), map(int, re.findall(r'\d+', color)))) | {'is_empty': False}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('generate', 'commercial', 'current', 'save', 'reread'))
    parser.add_argument('corpus', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--outputs', type=Path)
    parser.add_argument('--manifest', type=Path)
    args = parser.parse_args()
    if args.action == 'generate':
        with ZipFile(args.corpus, 'w') as archive:
            for name, _, raw in inputs():
                info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, raw)
        return
    native = args.action in ('commercial', 'reread')
    if native:
        import aspose.words as aw
        assert version('aspose-words') == '26.9.0'
    else:
        from aspose.words_foss import Document
    records = []
    if args.action == 'reread':
        if args.manifest is None or args.outputs is None:
            parser.error('reread requires --manifest and --outputs')
        manifest = json.loads(args.manifest.read_text())
        with ZipFile(args.outputs) as archive:
            for row in manifest['records']:
                raw = archive.read(row['output'])
                assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
                records.append({**row, 'observed': observed(aw.Document(BytesIO(raw)), True)})
        assert len(records) == 1024
    elif args.action == 'save':
        if args.outputs is None:
            parser.error('save requires --outputs')
        from aspose.words_foss import SaveFormat
        from aspose.words_foss import light_document_model as ldm
        from aspose.words_foss.docx_writer import LdmDocxWriter
        from aspose.words_foss.saving import OoxmlSaveOptions
        with ZipFile(args.corpus) as archive, ZipFile(args.outputs, 'w', ZIP_DEFLATED) as out:
            for name, _, _ in inputs():
                model = Document(BytesIO(archive.read(name))).light_document_model
                for json_roundtrip in (False, True):
                    current = ldm.Document.model_validate_json(model.model_dump_json()) if json_roundtrip else model
                    for label, fmt in (('docx', SaveFormat.DOCX), ('flat_opc', SaveFormat.FLAT_OPC)):
                        raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
                        output = f"{len(records):04d}.{'docx' if fmt == SaveFormat.DOCX else 'xml'}"
                        info = ZipInfo(output, (2020, 1, 1, 0, 0, 0))
                        info.compress_type = ZIP_DEFLATED
                        out.writestr(info, raw)
                        records.append({'input': name, 'json_roundtrip': json_roundtrip, 'format': label,
                                        'output': output, 'output_sha256': hashlib.sha256(raw).hexdigest()})
        assert len(records) == 1024
    else:
        with ZipFile(args.corpus) as archive:
            for name, values, _ in inputs():
                raw = archive.read(name)
                doc = aw.Document(BytesIO(raw)) if native else Document(BytesIO(raw))
                records.append({'input': name, 'values': values, 'sha256': hashlib.sha256(raw).hexdigest(),
                                'observed': observed(doc, native)})
        assert len(records) == 256
    report = {'version': version('aspose-words' if native else 'aspose-words-foss-enhanced'),
              'python': platform.python_version(), 'platform': platform.platform(), 'licensed': False if native else None,
              'records': records, 'full_Font_acceptance': False, 'rendering_acceptance': False}
    args.report.write_text(json.dumps(report, indent=2) + '\n')
    print(len(records))


if __name__ == '__main__':
    main()
