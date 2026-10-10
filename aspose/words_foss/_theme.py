"""Licensed default DrawingML theme used for a missing-theme paint projection."""
from functools import lru_cache
from importlib import resources


@lru_cache(maxsize=1)
def default_theme_bytes() -> bytes:
    """Return the MIT-licensed template; source packages are not modified."""
    return (resources.files('aspose.words_foss.docx_writer') / 'resources' /
            'default_theme.xml').read_bytes()
