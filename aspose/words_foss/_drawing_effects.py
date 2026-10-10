"""PDF projection of the calibrated zero-blur rectangular outer-shadow subset."""
from math import cos, sin, radians
from contextlib import contextmanager

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


def resolve_drawing_line_join(source):
    """Resolve the captured rectangular outline join, keeping legacy borders separate."""
    if not source.xml:
        return 'round'
    join = fromstring(source.xml, forbid_dtd=True)
    name = join.tag.removeprefix(A)
    if len(join) or set(join.attrib) - ({'lim'} if name == 'miter' else set()):
        raise NotImplementedError('DrawingML line join attributes or children')
    if name == 'miter' and 'lim' in join.attrib:
        try:
            if int(join.get('lim')) < 0:
                raise ValueError
        except ValueError as error:
            raise NotImplementedError('DrawingML miter join limit') from error
    # Native 26.9 rectangular paths retain lim in XML, but use PDF's existing
    # miter limit for both omitted and explicitly low limits. Do not invent an M
    # operator from general non-rectangular DrawingML semantics.
    return name


def paint_drawing_line_rect(pdf, x, y, width, height, style, *, rotation=0, page_space=False):
    """Keep captured EMU geometry precise through PDF point and rotation conversion."""
    if page_space:
        _paint_page_space_line_rect(pdf, x, y, width, height, style, rotation)
        return
    operator = {"F": "f", "D": "S", "DF": "B"}[style]
    rectangle = (f"{x * pdf.k:.6f} {(pdf.h - y) * pdf.k:.6f} "
                 f"{width * pdf.k:.6f} {-height * pdf.k:.6f} re {operator}")
    if not rotation:
        pdf._out(rectangle)
        return
    from fpdf.drawing_primitives import Transform
    user_to_pdf = Transform.scaling(pdf.k, -pdf.k).translate(0, pdf.h_pt)
    outer = Transform.rotation_d(rotation).about(x + width / 2, y + height / 2)
    exact = user_to_pdf.inverse() @ outer @ user_to_pdf
    command, _ = exact.render(None)
    actual = Transform(*(float(value) for value in command.split()[:6]))
    inverse = actual.inverse()
    def matrix(value):
        return ' '.join(f'{v:.12f}' for v in
                        (value.a, value.b, value.c, value.d, value.e, value.f)) + ' cm'
    # Cancel the matrix actually emitted by FPDF, then apply the exact rotation
    # in this path's private graphics state. Following text/shape CTMs stay intact.
    pdf._out('q ' + matrix(inverse) + ' ' + matrix(exact) + ' ' + rectangle + ' Q')


def _page_space_rect_geometry(pdf, x, y, width, height, rotation):
    from struct import pack, unpack
    def single(value):
        return unpack('f', pack('f', value))[0]
    w, h = single(width * pdf.k), single(height * pdf.k)
    cx, cy = single(w / 2), single(h / 2)
    angle = radians((rotation + 180) % 360 - 180)
    c, s = single(cos(angle)), single(sin(angle))
    tx = single(cx - single(single(c * cx) - single(s * cy)))
    ty = single(cy - single(single(s * cx) + single(c * cy)))
    points = [(single(single(single(c * u) - single(s * v)) + tx),
               single(single(single(s * u) + single(c * v)) + ty))
              for u, v in ((0, 0), (w, 0), (w, h), (0, h))]
    # Native layout rounds the page position and rotated figure origin in points.
    # Independent EMU controls on either side of .001pt boundaries establish both
    # steps. Keep the resulting origin and local points separate in the PDF: a
    # flattened path or rounded rotation/inverse pair changes rasterizer precision.
    left = min(u for u, _ in points)
    top = min(v for _, v in points)
    px = single(round(round(x * pdf.k, 3) + left, 3) - left)
    py = single(round(round(y * pdf.k, 3) + top, 3) - top)
    return c, s, points, px, py


@contextmanager
def drawing_rect_text_context(pdf, x, y, width, height, rotation):
    """Share the captured rectangle's precise local rotation with its text."""
    c, s, points, px, py = _page_space_rect_geometry(pdf, x, y, width, height, rotation)
    tx, ty = points[0]
    with pdf.local_context():
        pdf._out(f'1 0 0 -1 0 {pdf.h_pt:.8f} cm '
                 f'1 0 0 1 {px:.8f} {py:.8f} cm '
                 f'{c:.8f} {s:.8f} {-s:.8f} {c:.8f} {tx:.8f} {ty:.8f} cm '
                 f'1 0 0 -1 {-x * pdf.k:.8f} {(pdf.h - y) * pdf.k:.8f} cm')
        yield


def _paint_page_space_line_rect(pdf, x, y, width, height, style, rotation):
    """Project captured solid geometry once, without cancelling a rounded outer CTM."""
    _, _, points, px, py = _page_space_rect_geometry(pdf, x, y, width, height, rotation)
    path = ' '.join(f'{u:.8f} {v:.8f} ' + ('m' if i == 0 else 'l')
                    for i, (u, v) in enumerate(points))
    operator = {'F': 'f*', 'D': 'S', 'DF': 'B*'}[style]
    pdf._out(f'q 1 0 0 -1 0 {pdf.h_pt:.8f} cm '
             f'1 0 0 1 {px:.8f} {py:.8f} cm {path} h {operator} Q')
