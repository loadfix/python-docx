.. _WdDocGridType:

``WD_DOC_GRID_TYPE``
====================

Specifies the type of document grid applied to a section.

Example::

    from docx.enum.section import WD_DOC_GRID_TYPE

    section.document_grid.type = WD_DOC_GRID_TYPE.LINES

----

DEFAULT
    No document grid is applied.

LINES
    Grid specifies lines per page only.

LINES_AND_CHARS
    Grid specifies both lines per page and characters per line.

SNAP_TO_CHARS
    Grid snaps characters to a fixed-width column.
