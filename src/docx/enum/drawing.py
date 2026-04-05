"""Enumerations related to DrawingML positioning in WordprocessingML files."""

import enum


class WD_WRAP_TYPE(enum.Enum):
    """Specifies the text wrapping mode for a floating image.

    Corresponds to the choice group under `wp:anchor` for wrap elements.
    """

    NONE = 0
    """No text wrapping — image appears in front of or behind text."""

    SQUARE = 1
    """Text wraps around the bounding box of the image."""

    TIGHT = 2
    """Text wraps tightly around the image contour."""

    THROUGH = 3
    """Text wraps through the image."""

    TOP_AND_BOTTOM = 4
    """Text appears above and below the image only."""


WD_WRAP = WD_WRAP_TYPE


class WD_RELATIVE_HORZ_POS(enum.Enum):
    """Specifies the horizontal reference frame for a floating image position.

    Corresponds to `wp:positionH/@relativeFrom` attribute values.
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

    Corresponds to `wp:positionV/@relativeFrom` attribute values.
    """

    INSIDE_MARGIN = "insideMargin"
    LINE = "line"
    MARGIN = "margin"
    OUTSIDE_MARGIN = "outsideMargin"
    PAGE = "page"
    PARAGRAPH = "paragraph"
    TOP_MARGIN = "topMargin"
    BOTTOM_MARGIN = "bottomMargin"
