.. _WdHeaderFooterIndex:

``WD_HEADER_FOOTER_INDEX``
==========================

Identifies a header or footer in a section by its logical role.

Example::

    from docx.enum.section import WD_HEADER_FOOTER_INDEX

    header = section.header_for(WD_HEADER_FOOTER_INDEX.FIRST_PAGE)

----

PRIMARY
    Primary header/footer - used on odd pages and on pages not covered by
    the other indexes.

FIRST_PAGE
    Header/footer used on the first page of the section.

EVEN_PAGE
    Header/footer used on even pages of a recto/verso section.
