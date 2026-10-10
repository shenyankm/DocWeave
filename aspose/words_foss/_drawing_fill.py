"""DrawingML source-fill projection for rectangular multi-stop linear gradients.

Reference indices follow ECMA-376 fillRef (1-based fillStyleLst / 1001-based
bgFillStyleLst); phClr is supplied by the reference's colour child. Source XML
is never changed by resolving a PDF paint.
"""
from colorsys import rgb_to_hls, hls_to_rgb
from math import cos, sin, radians

from defusedxml.ElementTree import fromstring

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def _colour(node, scheme, placeholder=None, *, with_alpha=False):
    primitive = next((child for child in node if child.tag in
                      {A + 'srgbClr', A + 'schemeClr', A + 'sysClr'}), None)
    if primitive is None:
        raise NotImplementedError('DrawingML fill colour primitive')
    token = primitive.get('val', '')
    if primitive.tag == A + 'schemeClr':
        token = placeholder if token == 'phClr' else scheme.get(token)
    elif primitive.tag == A + 'sysClr':
        token = primitive.get('lastClr')
    if not token or len(token) != 6:
        raise NotImplementedError('DrawingML fill colour reference')
    rgb = tuple(int(token[i:i + 2], 16) for i in (0, 2, 4))
    alpha = 255
    for mod in primitive:
        value = int(mod.get('val', '0')) / 100000
        if mod.tag in {A + 'alpha', A + 'alphaMod', A + 'alphaOff'} and with_alpha:
            if ((mod.tag == A + 'alpha' and not 0 <= value <= 1) or
                    (mod.tag == A + 'alphaMod' and value < 0) or
                    (mod.tag == A + 'alphaOff' and not -1 <= value <= 1)):
                raise NotImplementedError('DrawingML alpha outside supported range')
            if mod.tag == A + 'alphaMod':
                alpha = round(alpha * value)
            elif mod.tag == A + 'alphaOff':
                alpha += round(value * 255)
            else:
                alpha = round(value * 255)
            # The fixed 26.9 renderer transforms opacity in byte space, with
            # ties-to-even rounding. Offsets are quantized before addition.
            alpha = min(255, max(0, alpha))
        elif mod.tag in {A + 'shade', A + 'tint'}:
            # DrawingML shade/tint mixes in linear-light RGB, then quantizes
            # to sRGB before the next colour transform.
            linear = [(v / 255 / 12.92 if v / 255 <= .04045 else
                       ((v / 255 + .055) / 1.055) ** 2.4) for v in rgb]
            linear = [v * value if mod.tag == A + 'shade' else
                      v * value + 1 - value for v in linear]
            rgb = tuple(round(255 * (12.92 * v if v <= .0031308 else
                                    1.055 * v ** (1 / 2.4) - .055)) for v in linear)
        elif mod.tag == A + 'satMod':
            hue, light, saturation = rgb_to_hls(*(v / 255 for v in rgb))
            rgb = tuple(int(v * 255) for v in
                        hls_to_rgb(hue, light, max(0, saturation * value)))
        else:
            raise NotImplementedError('DrawingML fill colour transform ' + mod.tag)
    rgb = tuple(min(255, max(0, v)) for v in rgb)
    return rgb + (alpha / 255,) if alpha != 255 else rgb


def resolve_drawing_fill(source, theme_data, *, with_transform=False):
    """Return (solid RGB[A], gradient (angle, positions, RGB[A]s)), preserving source.

    With ``with_transform``, gradient is paired with its rotWithShape flag;
    the PDF caller rotates both geometry and page-space paint consistently.

    An absent theme part uses the independently licensed default save theme.
    Existing incomplete themes are not silently replaced. Unsupported paints
    raise NotImplementedError so the PDF caller can report content loss.
    """
    from aspose.words_foss._theme import default_theme_bytes

    theme = fromstring(theme_data if theme_data is not None else default_theme_bytes())
    scheme = {}
    colours = theme.find(A + 'themeElements/' + A + 'clrScheme')
    if colours is not None:
        for child in colours:
            primitive = next(iter(child), None)
            if primitive is not None:
                scheme[child.tag.removeprefix(A)] = primitive.get('lastClr', primitive.get('val'))
    scheme.update({alias: scheme.get(target) for alias, target in
                   {'bg1': 'lt1', 'tx1': 'dk1', 'bg2': 'lt2', 'tx2': 'dk2'}.items()})
    fill = fromstring(source.direct_xml) if source.direct_xml else None
    placeholder = None
    if fill is None and source.style_xml:
        style = fromstring(source.style_xml)
        ref = style.find(A + 'fillRef')
        if ref is None:
            return None, None
        index = int(ref.get('idx', '0'))
        if index in (0, 1000):
            return None, None
        placeholder = ''.join(f'{v:02X}' for v in _colour(ref, scheme))
        group = 'fillStyleLst' if index < 1000 else 'bgFillStyleLst'
        index -= 1 if index < 1000 else 1001
        fills = theme.find(A + 'themeElements/' + A + 'fmtScheme/' + A + group)
        if fills is None or index < 0 or index >= len(fills):
            raise NotImplementedError('DrawingML fill style index')
        fill = fills[index]
    if fill is None or fill.tag == A + 'noFill':
        return None, None
    if fill.tag == A + 'solidFill':
        return _colour(fill, scheme, placeholder, with_alpha=True), None
    if fill.tag != A + 'gradFill':
        raise NotImplementedError('DrawingML fill type')
    line = fill.find(A + 'lin')
    if line is None:
        raise NotImplementedError('DrawingML path gradient')
    angle = int(line.get('ang', '0')) / 60000 % 360
    # Reflect about the rectangle's centre only when the fill follows its
    # transform. The default rotWithShape=true also applies to reflections.
    if fill.get('rotWithShape', '1') not in {'0', 'false'}:
        if source.flip_horizontal:
            angle = (180 - angle) % 360
        if source.flip_vertical:
            angle = -angle % 360
    stops = fill.find(A + 'gsLst')
    if stops is None or len(stops) < 2:
        raise NotImplementedError('DrawingML gradient stops')
    positions = tuple(int(stop.get('pos', '0')) / 100000 for stop in stops)
    if positions[0] < 0 or positions[-1] > 1 or any(
            b <= a for a, b in zip(positions, positions[1:])):
        raise NotImplementedError('DrawingML non-monotone/out-of-range gradient stops')
    colours = tuple(_colour(stop, scheme, placeholder, with_alpha=True) for stop in stops)
    uniform = all(colour == colours[0] for colour in colours)
    if uniform and len(colours[0]) == 3:
        return colours[0], None
    if not uniform and len(stops) > 2 and colours[0] == colours[-1]:
        raise NotImplementedError('DrawingML repeated-endpoint nonlinear gradient interpolation')
    if not uniform and len(stops) == 2 and positions == (0, 1):
        raise NotImplementedError('DrawingML two-stop nonlinear gradient interpolation')
    # Outside the declared interval, retain the nearest stop colour.
    if positions[0] > 0:
        positions, colours = (0,) + positions, (colours[0],) + colours
    if positions[-1] < 1:
        positions, colours = positions + (1,), colours + (colours[-1],)
    gradient = (angle, positions, colours)
    if with_transform:
        return None, (gradient, fill.get('rotWithShape', '1') not in {'0', 'false'})
    return None, gradient


def paint_gradient(pdf, gradient, x, y, width, height, *, rotation=0, geometry_rotation=0):
    """Fill the rectangle with a vector PDF shading (FPDF page coordinates)."""
    from fpdf.pattern import LinearGradient

    angle, positions, colours = gradient
    alphas = tuple(colour[3] if len(colour) == 4 else 1 for colour in colours)
    # Fresh 26.9 rectangle observations preserve @scaled in OOXML but use
    # the unscaled direction for 0/1/true/omitted at cardinal and oblique angles.
    dx, dy = cos(radians(angle)), sin(radians(angle))
    radius = abs(dx) * width / 2 + abs(dy) * height / 2
    center_x, center_y = x + width / 2, y + height / 2
    # FPDF patterns are defined in default page space, independently of the
    # current graphics transform. Rotate the endpoints only for shape-following
    # fills, preserving the span of the unrotated rectangle.
    if rotation:
        angle = radians(rotation)
        dx, dy = dx * cos(angle) - dy * sin(angle), dx * sin(angle) + dy * cos(angle)
    if any(len(colour) == 4 for colour in colours):
        from fpdf.drawing_primitives import DeviceRGB
        colours = tuple(DeviceRGB(*(v / 255 for v in colour[:3]),
                                  colour[3] if len(colour) == 4 else 1) for colour in colours)
    pattern = LinearGradient((center_x - dx * radius),
                             (center_y - dy * radius),
                             (center_x + dx * radius),
                             (center_y + dy * radius),
                             colours, bounds=positions[1:-1],
                             extend_before=True, extend_after=True)
    if pattern.has_alpha():
        _paint_local_alpha_gradient(pdf, pattern, alphas, x, y, width, height,
                                    geometry_rotation=geometry_rotation)
    else:
        with pdf.use_pattern(pattern):
            pdf.rect(x, y, width, height, 'F')


def _paint_local_alpha_gradient(pdf, pattern, alphas, x, y, width, height, *, geometry_rotation):
    """Keep the colour and opacity shadings in the same local point space."""
    from fpdf.drawing import DrawingContext, GradientPaint, PaintedPath, PaintSoftMask
    from fpdf.drawing_primitives import DeviceGray, Point, Transform
    from fpdf.enums import PathPaintRule
    from fpdf.pattern import LinearGradient

    k = pdf.k
    # The fixed 26.9 exporter quantizes positioned rectangle origins to 0.001
    # point. Extents and gradient stops retain their independent precision.
    origin_x, origin_y = round(x * k, 3), round(y * k, 3)
    page_height = pdf.h * k
    coords = tuple((value - (x if index % 2 == 0 else y)) * k
                   for index, value in enumerate(pattern.coords))
    local_gradient = LinearGradient(*coords, pattern.colors, bounds=pattern.bounds,
                                    extend_before=True, extend_after=True)
    center_x, center_y = width * k / 2, height * k / 2
    theta = radians(geometry_rotation)
    corners = [(0, 0), (width * k, 0), (width * k, height * k), (0, height * k)]
    path = PaintedPath()
    for index, (px, py) in enumerate(corners):
        corner = (center_x + (px - center_x) * cos(theta) - (py - center_y) * sin(theta),
                  center_y + (px - center_x) * sin(theta) + (py - center_y) * cos(theta))
        (path.move_to if index == 0 else path.line_to)(*corner)
    path.close()
    path.style.stroke_color = None
    path.style.paint_rule = PathPaintRule.FILL_NONZERO
    from copy import deepcopy
    mask = deepcopy(path)
    if all(alpha == alphas[0] for alpha in alphas):
        mask.style.fill_color = DeviceGray(alphas[0])
    else:
        alpha_gradient = LinearGradient(*coords, tuple(DeviceGray(alpha) for alpha in alphas),
                                        bounds=pattern.bounds, extend_before=True, extend_after=True)
        mask.style.fill_color = GradientPaint(alpha_gradient, apply_page_ctm=False)
    mask.style.fill_opacity = 1
    path.style.soft_mask = PaintSoftMask(mask, use_luminosity=True)
    path.style.fill_color = GradientPaint(local_gradient, apply_page_ctm=False,
        gradient_transform=Transform(a=1, b=0, c=0, d=-1, e=origin_x,
                                     f=page_height - origin_y))
    path.style.fill_color.skip_alpha = True
    from contextlib import nullcontext
    context = pdf.local_context() if geometry_rotation else nullcontext()
    with context:
        if geometry_rotation:
            # Invert the matrix actually emitted by FPDF, whose four-decimal
            # coefficients are not an orthogonal rotation. Applying another
            # rounded opposite-angle rotation leaves a residual scale in the
            # opacity mask while the page-space colour pattern is unaffected.
            user_to_pdf = Transform.scaling(k, -k).translate(0, pdf.h_pt)
            outer = Transform.rotation_d(geometry_rotation).about(x + width / 2, y + height / 2)
            command, _ = (user_to_pdf.inverse() @ outer @ user_to_pdf).render(None)
            actual = Transform(*(float(value) for value in command.split()[:6]))
            inverse = actual.inverse()
            pdf._out(' '.join(f'{value:.12f}' for value in
                             (inverse.a, inverse.b, inverse.c, inverse.d, inverse.e, inverse.f)) + ' cm')
        drawing = DrawingContext()
        drawing.add_item(path)
        # FPDF's millimetre drawing CTM rounds its scale to four decimals while
        # Pattern matrices retain eight. Render directly in points so the mask
        # does not acquire a different coordinate scale from its colour ramp.
        rendered = drawing.render(pdf._resource_catalog, Point(0, 0), 1, 0,
                                  pdf._current_graphic_style())
        pdf._resource_catalog.index_stream_resources(rendered, pdf.page)
        pdf._out(f'q 1 0 0 1 {origin_x:.8f} {page_height-origin_y:.8f} cm {rendered} Q')
        pdf._set_min_pdf_version('1.4')
