"""Read-only style identities and native default selection over source metadata."""


def style_kind(attributes):
    kind = attributes.get('type') or 'paragraph'
    if kind not in {'paragraph', 'character', 'table', 'numbering'}:
        raise ValueError('Invalid style type')
    return kind


def style_flag(value):
    if value not in {'', '0', '1', 'false', 'true', 'off', 'on'}:
        raise ValueError('Invalid style flag')
    return value in {'1', 'true', 'on'}


def builtin_normal(attributes, name):
    return (style_kind(attributes) == 'paragraph'
            and not style_flag(attributes.get('customStyle', ''))
            and (name or attributes.get('styleId', '')).lower() == 'normal')

