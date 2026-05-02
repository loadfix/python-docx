.. _WdBorderDisplay:

``WD_BORDER_DISPLAY``
=====================

Specifies which pages of a section display a page border.

Example::

    from docx.enum.section import WD_BORDER_DISPLAY

    section.page_borders.display = WD_BORDER_DISPLAY.FIRST_PAGE

----

ALL_PAGES
    Border is displayed on every page.

FIRST_PAGE
    Border is displayed only on the first page.

NOT_FIRST_PAGE
    Border is displayed on every page except the first.
