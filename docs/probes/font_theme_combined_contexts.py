"""Move owned combined tint/shade into defaults and paragraph/character styles."""

import runpy
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

BASE = runpy.run_path(str(Path(__file__).with_name('font_theme_colors.py')))
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def inputs():
    for name, case, raw in BASE['inputs']():
        if 'themeTint' not in case['modifiers'] or 'themeShade' not in case['modifiers']:
            continue
        for layer in ('default', 'paragraph', 'character', 'direct'):
            with ZipFile(BytesIO(raw)) as archive:
                parts = {n: archive.read(n) for n in archive.namelist()}
            document, styles = (ET.fromstring(parts[n]) for n in ('word/document.xml', 'word/styles.xml'))
            paragraph = document.find(W + 'body/' + W + 'p')
            rpr = paragraph.find(W + 'r/' + W + 'rPr')
            color = rpr.find(W + 'color')
            if layer != 'direct':
                rpr.remove(color)
                if layer == 'default':
                    defaults = ET.SubElement(styles, W + 'docDefaults')
                    properties = ET.SubElement(ET.SubElement(defaults, W + 'rPrDefault'), W + 'rPr')
                else:
                    style = ET.SubElement(styles, W + 'style', {W + 'type': layer, W + 'styleId': 'Owned'})
                    ET.SubElement(style, W + 'name', {W + 'val': 'Owned'})
                    properties = ET.SubElement(style, W + 'rPr')
                    target = paragraph.find(W + 'pPr') if layer == 'paragraph' else rpr
                    ET.SubElement(target, W + ('pStyle' if layer == 'paragraph' else 'rStyle'), {W + 'val': 'Owned'})
                properties.append(color)
            parts['word/document.xml'], parts['word/styles.xml'] = ET.tostring(document), ET.tostring(styles)
            output = BytesIO()
            with ZipFile(output, 'w') as archive:
                for part, value in parts.items():
                    info = ZipInfo(part, (2020, 1, 1, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    archive.writestr(info, value)
            yield name.replace('.docx', '-' + layer + '.docx'), {**case, 'layer': layer}, output.getvalue()


if __name__ == '__main__':
    BASE['main'](inputs, 16)
