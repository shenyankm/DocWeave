"""Observe DOM Boolean reads and run/paragraph/character assignments on owned inputs."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def snapshot(document, field, commercial):
    if commercial:
        import aspose.words as aw
        paragraph = next(p.as_paragraph() for p in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                         if p.get_text().strip() == 'IMPORT')
        fonts = (paragraph.runs[0].as_run().font, document.styles.get_by_name('P').font,
                 document.styles.get_by_name('C').font)
    else:
        fonts = (document.body.paragraphs[0].runs[0].effective_font,
                 document.styles.get_by_id('P').font, document.styles.get_by_id('C').font)
    return {name: getattr(font, field) for name, font in zip(('run', 'paragraph', 'character'), fonts)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('commercial', 'current', 'reread'))
    parser.add_argument('root', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--outputs', type=Path)
    parser.add_argument('--manifest', type=Path)
    args = parser.parse_args()
    if args.action == 'reread' and (args.manifest is None or args.outputs is None):
        parser.error('reread requires --manifest and --outputs')
    if args.action == 'commercial' and args.outputs is not None:
        parser.error('--outputs is only supported for current and reread')
    commercial = args.action != 'current'
    if commercial:
        import aspose.words as aw
        assert version('aspose-words') == '26.9.0'
    else:
        from aspose.words_foss import DocxDocument
    if args.action == 'reread':
        source = json.loads(args.manifest.read_text())
        records = []
        with ZipFile(args.outputs) as archive:
            for row in source['records']:
                for saved in row['outputs']:
                    raw = archive.read(saved['storage'])
                    assert hashlib.sha256(raw).hexdigest() == saved['sha256']
                    records.append({**saved, 'observed': snapshot(aw.Document(BytesIO(raw)), row['field'], True)})
        assert len(records) == 10692
        args.report.write_text(json.dumps({'version': version('aspose-words'), 'python': platform.python_version(),
                                          'platform': platform.platform(), 'licensed': False, 'records': records}, indent=2) + '\n')
        print(len(records), 'cold observations')
        return
    source = json.loads((args.root / 'font-boolean-contexts-26.9.json').read_text())
    records = []
    outputs = ZipFile(args.outputs, 'w', compression=ZIP_DEFLATED) if args.outputs else None
    stored = {}
    with ZipFile(args.root / source['corpus']) as archive:
        for row in source['records']:
            raw = archive.read(row['input'])
            for target in ('run', 'paragraph', 'character'):
                for value in (False, True):
                    document = aw.Document(BytesIO(raw)) if commercial else DocxDocument(BytesIO(raw))
                    if commercial:
                        paragraph = next(p.as_paragraph() for p in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                                         if p.get_text().strip() == 'IMPORT')
                        run = paragraph.runs[0].as_run()
                        p, c = document.styles.get_by_name('P'), document.styles.get_by_name('C')
                        before = snapshot(document, row['field'], commercial)
                        font = run.font if target == 'run' else (p if target == 'paragraph' else c).font
                        setattr(font, row['field'], value)
                        observed = snapshot(document, row['field'], commercial)
                    else:
                        run = document.body.paragraphs[0].runs[0]
                        p, c = document.styles.get_by_id('P'), document.styles.get_by_id('C')
                        before = snapshot(document, row['field'], commercial)
                        font = run.font if target == 'run' else (p if target == 'paragraph' else c).font
                        setattr(font, row['field'], value)
                        observed = snapshot(document, row['field'], commercial)
                    records.append({'input': row['input'], 'source_sha256': hashlib.sha256(raw).hexdigest(),
                                    'field': row['field'], 'target': target, 'value': value,
                                    'before': before, 'after': observed})
                    if outputs:
                        assert not commercial
                        records[-1]['outputs'] = []
                        for fmt in ('docx', 'flat_opc'):
                            data = document.to_bytes() if fmt == 'docx' else document.to_flat_opc()
                            digest = hashlib.sha256(data).hexdigest()
                            storage = stored.get(digest)
                            if storage is None:
                                storage = f'{len(stored):05d}.{fmt}'
                                info = ZipInfo(storage, (2020, 1, 1, 0, 0, 0))
                                info.compress_type = ZIP_DEFLATED
                                outputs.writestr(info, data)
                                stored[digest] = storage
                            records[-1]['outputs'].append({'format': fmt, 'storage': storage, 'sha256': digest,
                                                           'current_cold': snapshot(DocxDocument(BytesIO(data)), row['field'], False)})
    if outputs:
        outputs.close()
    assert len(records) == 5346
    report = {'version': version('aspose-words' if commercial else 'aspose-words-foss-enhanced'),
              'python': platform.python_version(), 'platform': platform.platform(), 'licensed': False if commercial else None,
              'records': records, 'full_Font_acceptance': False, 'rendering_acceptance': False,
              'scope': '891 owned inputs; run effective and style nearest Boolean getters, explicit run/paragraph/character style setters; '
                       + ('DOCX/Flat OPC saved getters; no rendering acceptance' if args.outputs else 'no save or rendering acceptance')}
    args.report.write_text(json.dumps(report, indent=2) + '\n')
    print(len(records), 'observations')


if __name__ == '__main__':
    main()
