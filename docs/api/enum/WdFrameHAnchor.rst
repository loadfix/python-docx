.. _WdFrameHAnchor:

``WD_FRAME_H_ANCHOR``
=====================

Specifies the horizontal anchor of a text frame.

Example::

    from docx.enum.text import WD_FRAME_H_ANCHOR

    paragraph.paragraph_format.frame.horizontal_anchor = WD_FRAME_H_ANCHOR.MARGIN

----

TEXT
    Horizontal position is relative to the text of the paragraph.

MARGIN
    Horizontal position is relative to the page margin.

PAGE
    Horizontal position is relative to the page edge.
