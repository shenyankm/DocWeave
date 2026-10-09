"""Visible-run filter: drop instruction runs inside a field's code section."""


from aspose.words_foss import light_document_model as ldm

# ShapeType.RECTANGLE — Shape.CreateHorizontalRule builds a filled inline
# rectangle, and nothing else in the pipeline assigns this shape type.
HORIZONTAL_RULE_SHAPE_TYPE = 1


def is_horizontal_rule_shape(shape: ldm.Shape) -> bool:
    """Whether *shape* is the rectangle a horizontal rule is drawn with."""
    return (
        shape.shape_type == HORIZONTAL_RULE_SHAPE_TYPE
        and not shape.has_image
        and shape.text_box is None
    )


def visible_children(para: ldm.Paragraph) -> list:
    """Exclude instructions while retaining each nested field's code/result state."""
    code_sections = []
    out: list[ldm.Run] = []
    for child in para._children:
        if isinstance(child, ldm.FieldStart):
            code_sections.append(True)
        elif isinstance(child, ldm.FieldSeparator):
            if code_sections:
                code_sections[-1] = False
        elif isinstance(child, ldm.FieldEnd):
            if code_sections:
                code_sections.pop()
        elif not any(code_sections):
            out.append(child)
    return out


def visible_runs(para: ldm.Paragraph) -> list[ldm.Run]:
    """Return visible runs; hidden-font filtering remains the caller's policy."""
    return [child for child in visible_children(para) if isinstance(child, ldm.Run)]
