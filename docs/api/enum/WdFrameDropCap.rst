.. _WdFrameDropCap:

``WD_FRAME_DROP_CAP``
=====================

Specifies whether a text frame is a drop-cap frame and where it is located.

Example::

    from docx.enum.text import WD_FRAME_DROP_CAP

    paragraph.paragraph_format.frame.drop_cap = WD_FRAME_DROP_CAP.DROP

----

NONE
    Not a drop-cap frame.

DROP
    Drop-cap frame dropped into the paragraph text.

MARGIN
    Drop-cap frame positioned in the margin.
