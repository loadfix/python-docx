.. _sections-advanced:

Advanced Section Features
=========================

The :ref:`sections` page covers the section properties that appear on every
Word document — page size, margins, orientation, section start type, and the
primary header/footer pair. This page covers the less-common *per-section*
settings that |docx| also exposes:

* page borders,
* line numbering,
* printer paper-source hints,
* the East Asian document grid,
* text direction and right-to-left flow,
* distinct odd/even and first-page header/footer definitions, and
* multi-column page layout.

Every feature documented here lives on the |Section| object. Unless otherwise
noted, all setters are safe to call on a freshly-opened section without first
ensuring the underlying XML element exists — |docx| creates and removes the
element as needed.


Page borders
------------

.. currentmodule:: docx.section

Word can draw a decorative border around the printable area of each page in a
section. The :attr:`Section.page_borders` property returns a |PageBorders|
proxy exposing the four edges (``top``, ``bottom``, ``left``, ``right``) plus
the ``display`` and ``offset_from`` attributes. Each edge is a |PageBorder|
object with ``style``, ``width``, ``color``, and ``space`` attributes::

    >>> from docx import Document
    >>> from docx.enum.section import WD_BORDER_DISPLAY, WD_BORDER_OFFSET_FROM
    >>> from docx.enum.text import WD_BORDER_STYLE
    >>> from docx.shared import Pt, RGBColor

    >>> document = Document()
    >>> section = document.sections[0]
    >>> section.page_borders.top.style      # no border defined
    None

The :meth:`Section.set_page_border` convenience method writes every
attribute of a single edge in one call. Any argument left as ``None`` leaves
the corresponding XML attribute untouched::

    >>> section.set_page_border(
    ...     "top",
    ...     style=WD_BORDER_STYLE.DOUBLE,
    ...     width=Pt(1.5),
    ...     color=RGBColor(0x00, 0x66, 0xCC),
    ...     space=Pt(24),
    ... )
    <docx.section.PageBorder object at 0x...>

Individual attributes can also be assigned directly on the edge proxy::

    >>> section.page_borders.bottom.style = WD_BORDER_STYLE.SINGLE
    >>> section.page_borders.bottom.width = Pt(1)

Assigning ``None`` to any attribute clears it; this is how you "reset" an
edge without deleting the whole |PageBorders| element::

    >>> section.page_borders.top.color = None

The |PageBorders| proxy also carries two whole-element attributes:

* :attr:`PageBorders.display` — a :ref:`WdBorderDisplay` member specifying
  on which pages the border is drawn (``ALL_PAGES``, ``FIRST_PAGE``,
  ``NOT_FIRST_PAGE``);
* :attr:`PageBorders.offset_from` — a :ref:`WdBorderOffsetFrom` member
  specifying whether the ``space`` attribute on each edge is measured from
  the text extents (``TEXT``) or from the page edge (``PAGE``).

::

    >>> section.page_borders.display = WD_BORDER_DISPLAY.ALL_PAGES
    >>> section.page_borders.offset_from = WD_BORDER_OFFSET_FROM.PAGE

To remove every page-border definition from a section, call
:meth:`Section.remove_page_borders`. The call is a no-op when the section
has no borders defined::

    >>> section.remove_page_borders()


Line numbering
--------------

Line numbers are displayed in the margin alongside each numbered line. They
are commonly used in legal documents and screenplays.
:attr:`Section.line_numbering` returns a |LineNumbering| proxy when the
section has a ``<w:lnNumType>`` element, or ``None`` when it does not::

    >>> from docx.enum.section import WD_LINE_NUMBERING_RESTART
    >>> from docx.shared import Pt

    >>> section.line_numbering
    None

Use :meth:`Section.set_line_numbering` to create or update the element.
Arguments left as ``None`` leave any existing attribute unchanged::

    >>> ln = section.set_line_numbering(
    ...     count_by=1,                              # every line
    ...     start=1,
    ...     distance=Pt(20),                         # 20pt from the text
    ...     restart=WD_LINE_NUMBERING_RESTART.NEW_PAGE,
    ... )
    >>> ln.count_by, ln.start, ln.restart
    (1, 1, NEW_PAGE (2))

The four attributes are also individually settable after the fact::

    >>> section.line_numbering.count_by = 5            # only every 5th line
    >>> section.line_numbering.restart = WD_LINE_NUMBERING_RESTART.CONTINUOUS

To turn line numbering off for a section, call
:meth:`Section.remove_line_numbering`. The call is a no-op when the section
has no line numbering defined::

    >>> section.remove_line_numbering()
    >>> section.line_numbering is None
    True


Paper source (printer tray)
---------------------------

Word exposes a printer-tray hint on each section so that, for example, the
first sheet of a multi-page document can be drawn from a letterhead tray
while the remaining sheets come from a standard-paper tray. |docx| surfaces
the hint as two properties on |Section|:

* :attr:`Section.first_page_paper_source` — the tray number used for the
  first page of the section;
* :attr:`Section.other_pages_paper_source` — the tray number used for
  subsequent pages.

Both return ``int`` or ``None``::

    >>> section.first_page_paper_source, section.other_pages_paper_source
    (None, None)
    >>> section.first_page_paper_source = 7
    >>> section.other_pages_paper_source = 15
    >>> section.first_page_paper_source, section.other_pages_paper_source
    (7, 15)

Clearing a value by assigning ``None`` removes the underlying XML attribute.
When both values are cleared the enclosing ``<w:paperSrc>`` element is
removed from the ``<w:sectPr>``::

    >>> section.first_page_paper_source = None
    >>> section.other_pages_paper_source = None

.. note::
   Tray numbers are printer-specific. Word doesn't validate the value against
   any printer's supported bins — the integer is carried through and passed
   to the printer driver at print time.


East Asian document grid
------------------------

The ``<w:docGrid>`` element controls the East Asian character grid for a
section: whether text is laid out against a grid of lines, or a grid of both
lines and characters, and what the pitch of that grid is.

:attr:`Section.document_grid` returns a |DocumentGrid| proxy or ``None``::

    >>> from docx.enum.section import WD_DOC_GRID_TYPE

    >>> dg = section.document_grid
    >>> dg.type, dg.line_pitch, dg.char_space
    (None, 360, None)

:meth:`Section.set_document_grid` creates or updates the element::

    >>> section.set_document_grid(
    ...     type=WD_DOC_GRID_TYPE.LINES_AND_CHARS,
    ...     line_pitch=312,
    ...     char_space=0,
    ... )
    <docx.section.DocumentGrid object at 0x...>

:meth:`Section.remove_document_grid` deletes the element entirely. Typical
Western-language documents do not need a document grid; the default template
created by ``Document()`` already writes a minimal ``<w:docGrid>`` carrying
only ``linePitch``.


Text direction and right-to-left
--------------------------------

Two properties on |Section| together control text-flow direction:

* :attr:`Section.text_direction` — a :ref:`WdTextDirection` member or
  ``None``. Maps to the ``<w:textDirection>`` child of ``<w:sectPr>``. Use
  this to rotate section body text 90° for East Asian vertical layouts.
* :attr:`Section.right_to_left` — ``True`` when this section flows
  right-to-left (e.g. for Arabic or Hebrew body text). Maps to the
  ``<w:bidi>`` child.

::

    >>> from docx.enum.table import WD_TEXT_DIRECTION

    >>> section.text_direction
    None
    >>> section.right_to_left
    False
    >>> section.text_direction = WD_TEXT_DIRECTION.TB_RL
    >>> section.right_to_left = True
    >>> section.text_direction, section.right_to_left
    (TB_RL (1), True)

Assigning ``None`` to :attr:`Section.text_direction` removes the
``<w:textDirection>`` element. Assigning ``False`` or ``None`` to
:attr:`Section.right_to_left` removes the ``<w:bidi>`` element.

.. note::
   :attr:`Section.right_to_left` is orthogonal to the individual paragraph
   or run *bidi* settings. Setting it ``True`` affects the default column
   order, gutter placement, and default paragraph direction for the whole
   section; run-level and paragraph-level RTL settings still apply on top.


Odd, even, and first-page headers & footers
-------------------------------------------

Every section carries three pairs of header/footer slots:

================  ========================================================
``header``        primary — used for every page unless overridden
``first_page_*``  used for the first page of the section when enabled
``even_page_*``   used for even-numbered pages when enabled
================  ========================================================

Each slot is a |_Header| or |_Footer| proxy accessed through the
corresponding :class:`Section` property (:attr:`Section.header`,
:attr:`Section.first_page_header`, :attr:`Section.even_page_header`, plus
the ``_footer`` variants).

Two toggles control whether the non-primary slots are honored by Word:

* :attr:`Section.different_first_page_header_footer` is a **per-section**
  flag mapped to ``<w:titlePg/>``. It enables the *first-page* header and
  footer for only the section it is set on.
* :attr:`Section.different_odd_and_even_pages_header_footer` is a
  **document-level** flag mapped to ``<w:evenAndOddHeaders/>`` in the
  settings part. Setting it affects every section in the document. It is
  surfaced on |Section| purely for discoverability — any section exposes
  the same underlying document-wide value.

::

    >>> section.different_first_page_header_footer = True
    >>> section.first_page_header.paragraphs[0].text = "First-page header"
    >>> section.first_page_footer.paragraphs[0].text = "First-page footer"

    >>> section.different_odd_and_even_pages_header_footer = True
    >>> section.even_page_header.paragraphs[0].text = "Even-page header"
    >>> section.even_page_footer.paragraphs[0].text = "Even-page footer"

    >>> section.header.paragraphs[0].text = "Odd-page header"
    >>> section.footer.paragraphs[0].text = "Odd-page footer"

Like the primary header/footer, each slot's ``is_linked_to_previous``
property controls whether the slot inherits its content from the
corresponding slot in the preceding section. Setting ``is_linked_to_previous
= False`` creates an empty definition in this section that you can then
populate; setting it ``True`` drops the definition (if any) so the slot
inherits again.


Multi-column layout
-------------------

The :attr:`Section.columns` property returns a |SectionColumns| proxy
backed by the ``<w:cols>`` element. It behaves like a sequence of
|Column| objects and also carries three whole-element attributes:

* :attr:`SectionColumns.count` — number of columns (defaults to 1);
* :attr:`SectionColumns.equal_width` — ``True`` when every column has the
  same width (defaults to ``True``);
* :attr:`SectionColumns.space` — the gutter between columns when they are
  equal-width.

A brand-new section has no ``<w:cols>`` element and reports a single
column::

    >>> from docx.shared import Inches, Pt

    >>> cols = section.columns
    >>> cols.count, cols.equal_width, cols.space, len(cols)
    (1, True, None, 0)

To lay out three equal-width columns with an 18-point gutter, just assign
the three attributes::

    >>> cols.count = 3
    >>> cols.equal_width = True
    >>> cols.space = Pt(18)

Unequal columns are expressed as a sequence of explicit ``<w:col>``
children. Set ``equal_width`` to ``False`` and then populate the sequence
individually::

    >>> cols.count = 2
    >>> cols.equal_width = False
    >>> cols[0].width = Inches(2.5)
    >>> cols[0].space = Inches(0.5)
    >>> cols[1].width = Inches(4.0)

Each :class:`Column` exposes two properties — :attr:`Column.width` and
:attr:`Column.space` — both of which accept |Length| values or ``None``.

.. note::
   The ``<w:cols>`` element does not *store* an explicit
   ``<w:col>`` child for equal-width columns; Word computes per-column
   widths from ``count`` and the section's content width. Adding ``<w:col>``
   children only makes sense together with ``equal_width = False``.


API reference
-------------

The classes used on this page are documented in the :doc:`../api/section`
reference. The enumerations are documented in
:doc:`../api/enum/WdBorderDisplay`, :doc:`../api/enum/WdBorderOffsetFrom`,
:doc:`../api/enum/WdBorderStyle`,
:doc:`../api/enum/WdLineNumberingRestart`,
:doc:`../api/enum/WdDocGridType`, and
:doc:`../api/enum/WdTextDirection`.
