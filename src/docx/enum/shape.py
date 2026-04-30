"""Enumerations related to DrawingML shapes in WordprocessingML files."""

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


class WD_ANCHOR_H(enum.Enum):
    """Horizontal anchor frame-of-reference for a floating image.

    Maps to `wp:positionH/@relativeFrom` attribute on a `wp:anchor` element.
    """

    PAGE = "page"
    MARGIN = "margin"
    COLUMN = "column"
    CHARACTER = "character"


class WD_ANCHOR_V(enum.Enum):
    """Vertical anchor frame-of-reference for a floating image.

    Maps to `wp:positionV/@relativeFrom` attribute on a `wp:anchor` element.
    """

    PAGE = "page"
    MARGIN = "margin"
    PARAGRAPH = "paragraph"
    LINE = "line"


class WD_WRAP_TYPE(enum.Enum):
    """Text-wrap style for a floating image.

    SQUARE, TIGHT, THROUGH, and TOP_AND_BOTTOM correspond to `wp:wrapSquare`,
    `wp:wrapTight`, `wp:wrapThrough`, and `wp:wrapTopAndBottom` respectively.
    BEHIND and IN_FRONT are both `wp:wrapNone`, distinguished by the `behindDoc`
    attribute on the parent `wp:anchor` element.
    """

    SQUARE = "square"
    TIGHT = "tight"
    THROUGH = "through"
    TOP_AND_BOTTOM = "topAndBottom"
    BEHIND = "behind"
    IN_FRONT = "inFront"
