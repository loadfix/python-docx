.. _WdShape:

``WD_SHAPE``
============

Identifies the preset geometry of a DrawingML shape.

Example::

    from docx.enum.shape import WD_SHAPE

    shape.preset_geometry = WD_SHAPE.OVAL

----

RECTANGLE
    Rectangle shape.

ROUNDED_RECTANGLE
    Rounded-rectangle shape.

OVAL
    Oval (ellipse) shape.

ARROW_RIGHT
    Right-arrow shape.

CALLOUT_ROUNDED_RECTANGLE
    Rounded-rectangle callout shape.
