.. _WdBreakType:

``WD_BREAK_TYPE``
=================

Specifies the type of break inserted into the text flow.

Example::

    from docx.enum.text import WD_BREAK_TYPE

    run.add_break(WD_BREAK_TYPE.PAGE)

----

LINE
    A line break.

LINE_CLEAR_LEFT
    Line break, clearing text wrap on the left.

LINE_CLEAR_RIGHT
    Line break, clearing text wrap on the right.

LINE_CLEAR_ALL
    Line break, clearing text wrap on both sides.

PAGE
    A page break.

COLUMN
    A column break.

SECTION_CONTINUOUS
    A continuous section break.

SECTION_EVEN_PAGE
    A section break that begins on the next even page.

SECTION_NEXT_PAGE
    A section break that begins on the next page.

SECTION_ODD_PAGE
    A section break that begins on the next odd page.

TEXT_WRAPPING
    A text-wrapping break.
