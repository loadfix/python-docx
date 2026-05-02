.. _WdAnchorH:

``WD_ANCHOR_H``
===============

Specifies the horizontal anchor used for positioning a floating shape or image.

Example::

    from docx.enum.shape import WD_ANCHOR_H

    floating_image.horizontal_anchor = WD_ANCHOR_H.MARGIN

----

PAGE
    Horizontal position is measured relative to the page edge.

MARGIN
    Horizontal position is measured relative to the page margin.

COLUMN
    Horizontal position is measured relative to the column.

CHARACTER
    Horizontal position is measured relative to a character anchor.
