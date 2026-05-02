.. _WdBorderStyle:

``WD_BORDER_STYLE``
===================

Specifies the line style for a paragraph, run, table, or page border.

Example::

    from docx.enum.text import WD_BORDER_STYLE

    run.font.border_style = WD_BORDER_STYLE.DOUBLE

----

NIL
    No border.

NONE
    No border.

SINGLE
    A single line.

THICK
    A single thick line.

DOUBLE
    A double line.

DOTTED
    A dotted line.

DASHED
    A dashed line.

DOT_DASH
    An alternating dot-dash line.

DOT_DOT_DASH
    An alternating dot-dot-dash line.

TRIPLE
    A triple line.

THIN_THICK_SMALL_GAP
    A thin-thick line with a small gap.

THICK_THIN_SMALL_GAP
    A thick-thin line with a small gap.

THIN_THICK_THIN_SMALL_GAP
    A thin-thick-thin line with a small gap.

THIN_THICK_MEDIUM_GAP
    A thin-thick line with a medium gap.

THICK_THIN_MEDIUM_GAP
    A thick-thin line with a medium gap.

THIN_THICK_THIN_MEDIUM_GAP
    A thin-thick-thin line with a medium gap.

THIN_THICK_LARGE_GAP
    A thin-thick line with a large gap.

THICK_THIN_LARGE_GAP
    A thick-thin line with a large gap.

THIN_THICK_THIN_LARGE_GAP
    A thin-thick-thin line with a large gap.

WAVE
    A wavy line.

DOUBLE_WAVE
    A double wavy line.

DASH_SMALL_GAP
    A dashed line with a small gap.

DASH_DOT_STROKED
    A dash-dot stroked line.

THREE_D_EMBOSS
    A 3D embossed line.

THREE_D_ENGRAVE
    A 3D engraved line.

OUTSET
    An outset line.

INSET
    An inset line.
