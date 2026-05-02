.. _WdBorderOffsetFrom:

``WD_BORDER_OFFSET_FROM``
=========================

Specifies the reference point used when measuring the offset of a page border.

Example::

    from docx.enum.section import WD_BORDER_OFFSET_FROM

    section.page_borders.offset_from = WD_BORDER_OFFSET_FROM.PAGE

----

TEXT
    Border is positioned relative to the text extents.

PAGE
    Border is positioned relative to the page edge.
