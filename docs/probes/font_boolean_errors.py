"""Record Boolean setter errors on owned run and style objects."""

import argparse
import json
import platform
from importlib.metadata import version
from io import BytesIO
from zipfile import ZipFile


def main():
    from pathlib import Path

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('commercial', 'current'))
    parser.add_argument('root', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    source = json.loads((args.root / 'font-boolean-contexts-26.9.json').read_text())
    native = args.action == 'commercial'
    if native:
        import aspose.words as aw
        assert version('aspose-words') == '26.9.0'
    else:
        from aspose.words_foss import DocxDocument
    values = (None, 0, 1, -1, 1.5, 'false', '', [], {})
    records = []
    with ZipFile(args.root / source['corpus']) as corpus:
        for row in source['records']:
            if row['values'] != [True, True, None, None]:
                continue
            for target in ('run', 'paragraph', 'character'):
                for value in values:
                    if native:
                        document = aw.Document(BytesIO(corpus.read(row['input'])))
                        paragraph = next(p.as_paragraph() for p in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                                         if p.get_text().strip() == 'IMPORT')
                        run = paragraph.runs[0].as_run()
                        font = run.font if target == 'run' else document.styles.get_by_name(
                            'P' if target == 'paragraph' else 'C').font
                        before = getattr(run.font, row['field'])
                    else:
                        document = DocxDocument(BytesIO(corpus.read(row['input'])))
                        run = document.body.paragraphs[0].runs[0]
                        font = run.font if target == 'run' else document.styles.get_by_id(
                            'P' if target == 'paragraph' else 'C').font
                        before = getattr(run.effective_font, row['field'])
                    error = None
                    try:
                        setattr(font, row['field'], value)
                    except (TypeError, ValueError) as exc:
                        error = type(exc).__name__
                    after = getattr(run.font if native else run.effective_font, row['field'])
                    records.append({'field': row['field'], 'target': target, 'input': value,
                                    'error': error, 'before': before, 'after': after})
    assert len(records) == 297
    report = {'version': version('aspose-words' if native else 'aspose-words-foss-enhanced'),
              'python': platform.python_version(), 'platform': platform.platform(),
              'licensed': False if native else None, 'records': records}
    args.report.write_text(json.dumps(report, indent=2) + '\n')
    print('observed', len(records))


if __name__ == '__main__':
    main()
