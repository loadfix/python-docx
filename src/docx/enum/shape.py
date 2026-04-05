"""Enumerations related to DrawingML shapes in WordprocessingML files."""
from __future__ import annotations

import enum


class WD_INLINE_SHAPE_TYPE(enum.Enum):
    """Corresponds to WdInlineShapeType enumeration.

    http://msdn.microsoft.com/en-us/library/office/ff192587.aspx.
    """

    CHART = 12
    LINKED_PICTURE = 4
    PICTURE = 3
    SMART_ART = 15
    NOT_IMPLEMENTED = -6


WD_INLINE_SHAPE = WD_INLINE_SHAPE_TYPE


class WD_DRAWING_TYPE(enum.Enum):
    """Type of content contained in a `<w:drawing>` element."""

    SHAPE = 1
    TEXT_BOX = 2
    GROUP = 3
    CHART = 4
    DIAGRAM = 5
    PICTURE = 6
