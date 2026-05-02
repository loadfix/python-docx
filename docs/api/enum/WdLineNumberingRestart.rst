.. _WdLineNumberingRestart:

``WD_LINE_NUMBERING_RESTART``
=============================

Specifies when automatic line numbering restarts in a section.

Example::

    from docx.enum.section import WD_LINE_NUMBERING_RESTART

    section.line_numbering.restart = WD_LINE_NUMBERING_RESTART.NEW_PAGE

----

CONTINUOUS
    Line numbering continues from the previous section.

NEW_SECTION
    Line numbering restarts at the beginning of each section.

NEW_PAGE
    Line numbering restarts at the beginning of each page.
