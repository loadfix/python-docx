.. _WdShadingPattern:

``WD_SHADING_PATTERN``
======================

Specifies the background pattern applied to a cell or run.

Example::

    from docx.enum.table import WD_SHADING_PATTERN

    cell.shading.pattern = WD_SHADING_PATTERN.SOLID

----

CLEAR
    No pattern, just background fill color.

SOLID
    Solid pattern (foreground color fills entire area).

HORZ_STRIPE
    Horizontal stripe pattern.

VERT_STRIPE
    Vertical stripe pattern.

REVERSE_DIAG_STRIPE
    Reverse diagonal stripe pattern.

DIAG_STRIPE
    Diagonal stripe pattern.

HORZ_CROSS
    Horizontal cross pattern.

DIAG_CROSS
    Diagonal cross pattern.

THIN_HORZ_STRIPE
    Thin horizontal stripe pattern.

THIN_VERT_STRIPE
    Thin vertical stripe pattern.

THIN_REVERSE_DIAG_STRIPE
    Thin reverse diagonal stripe pattern.

THIN_DIAG_STRIPE
    Thin diagonal stripe pattern.

THIN_HORZ_CROSS
    Thin horizontal cross pattern.

THIN_DIAG_CROSS
    Thin diagonal cross pattern.

NIL
    No shading.
