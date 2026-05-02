.. _WdFrameHAlign:

``WD_FRAME_H_ALIGN``
====================

Specifies the horizontal alignment of a text frame.

Example::

    from docx.enum.text import WD_FRAME_H_ALIGN

    paragraph.paragraph_format.frame.horizontal_align = WD_FRAME_H_ALIGN.CENTER

----

LEFT
    Frame is left-aligned.

CENTER
    Frame is center-aligned.

RIGHT
    Frame is right-aligned.

INSIDE
    Frame is aligned to the inside of the page (for facing pages).

OUTSIDE
    Frame is aligned to the outside of the page (for facing pages).
