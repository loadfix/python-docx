.. _WdWrapType:

``WD_WRAP_TYPE``
================

Specifies how text wraps around a floating shape or image.

Example::

    from docx.enum.shape import WD_WRAP_TYPE

    floating_image.wrap_type = WD_WRAP_TYPE.SQUARE

----

SQUARE
    Text wraps around the bounding box of the shape.

TIGHT
    Text wraps tightly around the shape contour.

THROUGH
    Text wraps through the shape, filling available concavities.

TOP_AND_BOTTOM
    Text flows above and below the shape only.

BEHIND
    Shape floats behind the text.

IN_FRONT
    Shape floats in front of the text.
