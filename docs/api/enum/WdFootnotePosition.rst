.. _WdFootnotePosition:

``WD_FOOTNOTE_POSITION``
========================

Specifies the position of footnotes on the page.

Example::

    from docx.enum.text import WD_FOOTNOTE_POSITION

    section.footnote_properties.position = WD_FOOTNOTE_POSITION.BOTTOM_OF_PAGE

----

BOTTOM_OF_PAGE
    Footnotes appear at the bottom of the page.

BENEATH_TEXT
    Footnotes appear immediately beneath the body text on the page.
