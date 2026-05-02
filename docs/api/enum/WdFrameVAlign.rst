.. _WdFrameVAlign:

``WD_FRAME_V_ALIGN``
====================

Specifies the vertical alignment of a text frame.

Example::

    from docx.enum.text import WD_FRAME_V_ALIGN

    paragraph.paragraph_format.frame.vertical_align = WD_FRAME_V_ALIGN.TOP

----

INLINE
    Frame is positioned inline with the surrounding text.

TOP
    Frame is top-aligned.

CENTER
    Frame is center-aligned vertically.

BOTTOM
    Frame is bottom-aligned.

INSIDE
    Frame is aligned to the inside of the page (for facing pages).

OUTSIDE
    Frame is aligned to the outside of the page (for facing pages).
