.. _WdAnchorV:

``WD_ANCHOR_V``
===============

Specifies the vertical anchor used for positioning a floating shape or image.

Example::

    from docx.enum.shape import WD_ANCHOR_V

    floating_image.vertical_anchor = WD_ANCHOR_V.PARAGRAPH

----

PAGE
    Vertical position is measured relative to the page edge.

MARGIN
    Vertical position is measured relative to the page margin.

PARAGRAPH
    Vertical position is measured relative to the anchor paragraph.

LINE
    Vertical position is measured relative to a line of text.
