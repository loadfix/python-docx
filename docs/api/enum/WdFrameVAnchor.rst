.. _WdFrameVAnchor:

``WD_FRAME_V_ANCHOR``
=====================

Specifies the vertical anchor of a text frame.

Example::

    from docx.enum.text import WD_FRAME_V_ANCHOR

    paragraph.paragraph_format.frame.vertical_anchor = WD_FRAME_V_ANCHOR.PAGE

----

TEXT
    Vertical position is relative to the text of the paragraph.

MARGIN
    Vertical position is relative to the page margin.

PAGE
    Vertical position is relative to the page edge.
