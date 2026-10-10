"""PDF projection of the calibrated zero-blur rectangular outer-shadow subset."""
from math import cos, sin, radians

from defusedxml.ElementTree import fromstring

from aspose.words_foss._drawing_fill import A, _colour


def resolve_outer_shadow(source, theme_data):
    """Return page-space (dx_mm, dy_mm, RGB[A]); never change source declarations."""
    root = fromstring(source.xml, forbid_dtd=True)
    if root.tag != A + 'effectLst' or len(root) != 1 or root[0].tag != A + 'outerShdw':
        raise NotImplementedError('DrawingML shape effects are preserved but not rendered')
    shadow = root[0]
    if len(shadow) != 1:
        raise NotImplementedError('DrawingML outer shadow colour declarations')
    known = {'blurRad', 'dist', 'dir', 'sx', 'sy', 'kx', 'ky', 'algn', 'rotWithShape'}
    if set(shadow.attrib) - known:
        raise NotImplementedError('DrawingML outer shadow attributes')
    try:
        blur = int(shadow.get('blurRad', '0'))
        distance = int(shadow.get('dist', '0'))
        direction = int(shadow.get('dir', '0'))
        scale = (int(shadow.get('sx', '100000')), int(shadow.get('sy', '100000')))
        skew = (int(shadow.get('kx', '0')), int(shadow.get('ky', '0')))
    except ValueError as error:
        raise NotImplementedError('DrawingML outer shadow numeric attributes') from error
    if blur != 0 or scale != (100000, 100000) or skew != (0, 0):
        raise NotImplementedError('DrawingML outer shadow blur, scale or skew')
    if (not 0 <= distance <= 2147483647 or not 0 <= direction < 21600000 or
            shadow.get('algn', 'b') not in {'b', 'ctr'} or
            shadow.get('rotWithShape', 'true') not in {'true', 'false', '0', '1'}):
        raise NotImplementedError('DrawingML outer shadow alignment or range')
    # Resolve through the same supported colour transforms as shape fills.
    from aspose.words_foss._theme import default_theme_bytes
    theme = fromstring(theme_data if theme_data is not None else default_theme_bytes())
    scheme = {}
    colours = theme.find(A + 'themeElements/' + A + 'clrScheme')
    if colours is not None:
        for child in colours:
            primitive = next(iter(child), None)
            if primitive is not None:
                scheme[child.tag.removeprefix(A)] = primitive.get('lastClr', primitive.get('val'))
    try:
        colour = _colour(shadow, scheme, with_alpha=True)
    except ValueError as error:
        raise NotImplementedError('DrawingML outer shadow colour') from error
    angle = radians(direction / 60000)
    return distance / 36000 * cos(angle), distance / 36000 * sin(angle), colour
