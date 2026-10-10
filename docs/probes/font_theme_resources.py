"""Generate an owned theme-image input and reproduce current/native cold saves."""

import argparse
import importlib.metadata
import runpy
from base64 import b64decode
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def generate_input():
    inputs = runpy.run_path(str(Path(__file__).with_name('font_theme_colors.py')))['inputs']
    _, _, raw = next(inputs())
    with ZipFile(BytesIO(raw)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    theme = ET.fromstring(parts.pop('word/theme/theme1.xml'))
    fills = theme.find('.//' + a + 'fillStyleLst')
    fills.remove(fills[0])
    fill = ET.Element(a + 'blipFill', dpi='0', rotWithShape='1')
    ET.SubElement(fill, a + 'blip', {r + 'embed': 'resource'})
    ET.SubElement(ET.SubElement(fill, a + 'stretch'), a + 'fillRect')
    fills.insert(0, fill)
    parts['custom/themes/owned.xml'] = ET.tostring(theme)
    # Freeze the owned 2x2 red PNG independently of encoder versions.
    parts['custom/assets/theme.png'] = b64decode('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP8zwACTGCSAQANHQEDgslx/wAAAABJRU5ErkJggg==')
    links = ET.Element(rel + 'Relationships')
    ET.SubElement(links, rel + 'Relationship', Id='resource', Type=r[1:-1] + '/image',
                  Target='../assets/theme.png')
    parts['custom/themes/_rels/owned.xml.rels'] = ET.tostring(links)
    links = ET.fromstring(parts['word/_rels/document.xml.rels'])
    for item in links:
        if item.get('Type', '').endswith('/theme'):
            item.set('Target', '../custom/themes/owned.xml')
    parts['word/_rels/document.xml.rels'] = ET.tostring(links)
    types = ET.fromstring(parts['[Content_Types].xml'])
    for item in types:
        if item.get('PartName') == '/word/theme/theme1.xml':
            item.set('PartName', '/custom/themes/owned.xml')
    ET.SubElement(types, ct + 'Override', PartName='/custom/assets/theme.png', ContentType='image/png')
    parts['[Content_Types].xml'] = ET.tostring(types)
    output = BytesIO()
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    return output.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['generate', 'current', 'commercial'])
    parser.add_argument('directory', type=Path)
    args = parser.parse_args()
    args.directory.mkdir(parents=True, exist_ok=True)
    if args.action == 'generate':
        (args.directory / 'input.docx').write_bytes(generate_input())
    elif args.action == 'current':
        from aspose.words_foss import Document, SaveFormat
        from aspose.words_foss.docx_writer import LdmDocxWriter
        from aspose.words_foss.saving import OoxmlSaveOptions
        model = Document(args.directory / 'input.docx').light_document_model
        for name, fmt in [('current.docx', SaveFormat.DOCX), ('current.xml', SaveFormat.FLAT_OPC)]:
            (args.directory / name).write_bytes(LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(model))
    else:
        assert importlib.metadata.version('aspose-words') == '26.9.0'
        import aspose.words as aw
        for name in ['input.docx', 'current.docx', 'current.xml']:
            aw.Document(str(args.directory / name)).save(str(args.directory / ('native-' + name + '.docx')),
                                                       aw.SaveFormat.DOCX)


if __name__ == '__main__':
    main()
