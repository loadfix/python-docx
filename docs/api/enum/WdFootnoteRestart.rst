.. _WdFootnoteRestart:

``WD_FOOTNOTE_RESTART``
=======================

Specifies when footnote numbering restarts.

Example::

    from docx.enum.text import WD_FOOTNOTE_RESTART

    section.footnote_properties.restart = WD_FOOTNOTE_RESTART.EACH_SECTION

----

CONTINUOUS
    Continuous numbering throughout the document.

EACH_SECTION
    Numbering restarts at the beginning of each section.

EACH_PAGE
    Numbering restarts at the beginning of each page.
