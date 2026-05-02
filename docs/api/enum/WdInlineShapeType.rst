.. _WdInlineShapeType:

``WD_INLINE_SHAPE_TYPE``
========================

Identifies the kind of content carried by an :class:`.InlineShape`.

Example::

    from docx.enum.shape import WD_INLINE_SHAPE_TYPE

    if inline_shape.type == WD_INLINE_SHAPE_TYPE.PICTURE:
        ...

----

CHART
    The inline shape is a chart.

LINKED_PICTURE
    The inline shape is a linked picture (external reference).

PICTURE
    The inline shape is an embedded picture.

SMART_ART
    The inline shape is a SmartArt diagram.

NOT_IMPLEMENTED
    The inline shape is of a kind not currently recognised by ``python-docx``.
