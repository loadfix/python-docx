.. _WdDrawingType:

``WD_DRAWING_TYPE``
===================

Identifies the kind of DrawingML content contained in a ``w:drawing`` element.

Example::

    from docx.enum.shape import WD_DRAWING_TYPE

    if drawing.type == WD_DRAWING_TYPE.PICTURE:
        ...

----

SHAPE
    A DrawingML shape.

TEXT_BOX
    A text box.

GROUP
    A group of shapes.

CHART
    An embedded chart.

DIAGRAM
    A SmartArt diagram.

PICTURE
    A picture.
