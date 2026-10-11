"""Document-model enums (bare-int values, no Python ``Enum`` subclassing)."""

from dockiln.model.enums.image import ImageType  # noqa: F401
from dockiln.model.enums.paragraph import (  # noqa: F401
    ParagraphAlignment,
    LineSpacingRule,
)
from dockiln.model.enums.style import (  # noqa: F401
    StyleIdentifier,
    StyleType,
    TabAlignment,
    TabLeader,
    NumberStyle,
)
from dockiln.model.enums.font import Underline  # noqa: F401
from dockiln.model.enums.layout import (  # noqa: F401
    HeightRule,
    SectionStart,
    Orientation,
)
from dockiln.model.enums.border import LineStyle  # noqa: F401
from dockiln.model.enums.table import (  # noqa: F401
    PreferredWidthType,
    CellMerge,
    CellVerticalAlignment,
)
