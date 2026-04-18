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


class WD_WRAP_TYPE(enum.Enum):
    """Specifies the text wrapping mode for a floating image.

    Maps to the `wp:wrapNone`, `wp:wrapSquare`, `wp:wrapTight`,
    `wp:wrapThrough`, and `wp:wrapTopAndBottom` child elements of `wp:anchor`.

    Note: Both ``IN_FRONT`` and ``BEHIND`` map to the ``wp:wrapNone`` element.
    They are distinguished by the ``behindDoc`` attribute on the ``wp:anchor``
    element — ``IN_FRONT`` sets ``behindDoc="0"`` and ``BEHIND`` sets
    ``behindDoc="1"``.
    """

    SQUARE = 1
    TIGHT = 2
    THROUGH = 3
    TOP_AND_BOTTOM = 4
    IN_FRONT = 5
    BEHIND = 6


class WD_RELATIVE_HORZ_POS(enum.Enum):
    """Specifies the horizontal reference frame for a floating image position.

    Maps to the `relativeFrom` attribute on `wp:positionH`.
    """

    CHARACTER = "character"
    COLUMN = "column"
    INSIDE_MARGIN = "insideMargin"
    LEFT_MARGIN = "leftMargin"
    MARGIN = "margin"
    OUTSIDE_MARGIN = "outsideMargin"
    PAGE = "page"
    RIGHT_MARGIN = "rightMargin"


class WD_RELATIVE_VERT_POS(enum.Enum):
    """Specifies the vertical reference frame for a floating image position.

    Maps to the `relativeFrom` attribute on `wp:positionV`.
    """

    BOTTOM_MARGIN = "bottomMargin"
    INSIDE_MARGIN = "insideMargin"
    LINE = "line"
    MARGIN = "margin"
    OUTSIDE_MARGIN = "outsideMargin"
    PAGE = "page"
    PARAGRAPH = "paragraph"
    TOP_MARGIN = "topMargin"
