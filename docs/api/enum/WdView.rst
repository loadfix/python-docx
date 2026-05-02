.. _WdView:

``WD_VIEW``
===========

Specifies the initial view mode Word uses when opening the document.

Example::

    from docx.enum.text import WD_VIEW

    settings.view = WD_VIEW.PRINT

----

NONE
    No view mode is specified.

PRINT
    Print layout view (Word's default editing view).

OUTLINE
    Outline view, showing document headings and hierarchy.

MASTER_PAGES
    Master-pages (master document) view.

NORMAL
    Normal (draft) view, emphasizing text flow over layout.

WEB
    Web layout view, showing the document as it would appear in a browser.

READING
    Full-screen reading view optimized for reading.
