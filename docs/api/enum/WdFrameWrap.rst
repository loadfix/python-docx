.. _WdFrameWrap:

``WD_FRAME_WRAP``
=================

Specifies how text wraps around a text frame.

Example::

    from docx.enum.text import WD_FRAME_WRAP

    paragraph.paragraph_format.frame.wrap = WD_FRAME_WRAP.AROUND

----

AUTO
    Text wraps around the frame on all sides.

NOT_BESIDE
    Text does not wrap beside the frame.

AROUND
    Text wraps around the frame.

NONE
    Text does not wrap around the frame.

TIGHT
    Text wraps tightly around the frame.

THROUGH
    Text wraps through the frame.
