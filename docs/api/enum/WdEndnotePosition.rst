.. _WdEndnotePosition:

``WD_ENDNOTE_POSITION``
=======================

Specifies the position of endnotes in the document.

Example::

    from docx.enum.text import WD_ENDNOTE_POSITION

    section.endnote_properties.position = WD_ENDNOTE_POSITION.END_OF_SECTION

----

END_OF_DOCUMENT
    Endnotes appear at the end of the document.

END_OF_SECTION
    Endnotes appear at the end of each section.
