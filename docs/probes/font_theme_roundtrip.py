"""Save owned theme inputs and measure native cold color/theme getters."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('save', 'reread'))
    parser.add_argument('corpus', type=Path)
    parser.add_argument('--outputs', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--manifest', type=Path)
    args = parser.parse_args()
    native = args.action == 'reread'
    rows = []
    if native:
        if args.manifest is None:
            parser.error('reread requires --manifest')
        import aspose.words as aw
        assert version('aspose-words') == '26.9.0'
        observed = runpy.run_path(str(Path(__file__).with_name('font_color_categories.py')))['observed']
        manifest = json.loads(args.manifest.read_text())
        with ZipFile(args.outputs) as archive:
            for row in manifest['records']:
                raw = archive.read(row['output'])
                assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
                doc = aw.Document(BytesIO(raw))
                paragraph = next(p.as_paragraph() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                                 if p.get_text().strip() == 'IMPORT')
                font = paragraph.runs[0].as_run().font
                rows.append({**row, 'getter': observed(doc, True), 'theme_color': int(font.theme_color),
                             'tint_and_shade': font.tint_and_shade})
    else:
        from aspose.words_foss import Document, SaveFormat
        from aspose.words_foss import light_document_model as ldm
        from aspose.words_foss.docx_writer import LdmDocxWriter
        from aspose.words_foss.saving import OoxmlSaveOptions
        inputs = runpy.run_path(str(Path(__file__).with_name('font_theme_colors.py')))['inputs']
        with ZipFile(args.corpus) as sources, ZipFile(args.outputs, 'w', ZIP_DEFLATED) as outputs:
            for name, _, _ in inputs():
                model = Document(BytesIO(sources.read(name))).light_document_model
                for phase in (False, True):
                    current = ldm.Document.model_validate_json(model.model_dump_json()) if phase else model
                    for label, fmt in (('docx', SaveFormat.DOCX), ('flat_opc', SaveFormat.FLAT_OPC)):
                        raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
                        output = f"{len(rows):03d}.{'docx' if label == 'docx' else 'xml'}"
                        outputs.writestr(output, raw)
                        rows.append({'input': name, 'json_roundtrip': phase, 'format': label, 'output': output,
                                     'output_sha256': hashlib.sha256(raw).hexdigest()})
    assert len(rows) == 128
    args.report.write_text(json.dumps({'version': version('aspose-words' if native else 'aspose-words-foss-enhanced'),
                                      'python': platform.python_version(), 'platform': platform.platform(),
                                      'licensed': False if native else None, 'records': rows,
                                      'full_Font_acceptance': False, 'rendering_acceptance': False}, indent=2) + '\n')
    print(len(rows))


if __name__ == '__main__':
    main()
