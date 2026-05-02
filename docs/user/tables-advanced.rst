.. _tables_advanced:

Advanced Table Formatting
=========================

The :ref:`tables` chapter covers the basics of creating a table, populating
it with text, and iterating over its rows and columns. This chapter covers
the *formatting* capabilities that were added to *python-docx* by the fork's
Phase D work: borders, shading, cell margins, autofit layout, row
properties, table-style conditional flags, cell text direction, and merge
introspection.

All of the proxies documented here are accessed off an existing |Table|,
|_Row|, or |_Cell|. They are created lazily and assigning a value to any
property writes the appropriate ``w:tblPr``, ``w:tcPr``, or ``w:trPr``
child on demand.


Borders
-------

Both tables and cells expose a *borders* proxy. On a |Table|, the proxy
covers six edges — ``top``, ``bottom``, ``left``, ``right``, ``inside_h``
(the horizontal rules between rows), and ``inside_v`` (the vertical rules
between columns). On a |_Cell|, only the four outer edges are available.

Each edge is a :class:`~docx.table.BorderElement` with three read/write
properties: ``style`` (a :class:`~docx.enum.table.WD_BORDER_STYLE` member),
``width`` (a |Length|), and ``color`` (an |RGBColor|). A fourth property,
``space``, controls the gap between the border and the cell content.

Reading a border edge when nothing has been set yields ``None`` on every
property::

    >>> from docx import Document
    >>> document = Document()
    >>> table = document.add_table(rows=3, cols=3)
    >>> table.borders.top.style
    >>> table.borders.top.style is None
    True

Assigning a value to any property creates the underlying ``w:tblBorders``
(or ``w:tcBorders``) element and the specific edge on demand::

    >>> from docx.enum.table import WD_BORDER_STYLE
    >>> from docx.shared import Pt, RGBColor
    >>> table.borders.top.style = WD_BORDER_STYLE.SINGLE
    >>> table.borders.top.width = Pt(0.5)
    >>> table.borders.top.color = RGBColor(0x00, 0x00, 0x00)

To clear an edge, assign ``None``::

    >>> table.borders.top.style = None

The :meth:`.Table.set_borders` convenience method lets you apply a
consistent border treatment across several edges in one call. It is
particularly handy for the APA-7 "horizontal-only" table style::

    >>> table.set_borders(top=True, bottom=True, inside_h=True)

``set_borders`` always writes to all six edges: edges passed as ``True``
are set to the supplied style/width/color (defaulting to
``WD_BORDER_STYLE.SINGLE``, ``Pt(0.5)``, and black), while those left as
``False`` are explicitly set to ``WD_BORDER_STYLE.NONE`` so the table
style's defaults do not show through.

Cell-level borders work the same way and override the table-level values
for that one cell::

    >>> cell = table.cell(0, 0)
    >>> cell.borders.left.style = WD_BORDER_STYLE.THICK
    >>> cell.borders.left.width = Pt(1)
    >>> cell.borders.left.color = RGBColor(0xFF, 0x00, 0x00)


Cell shading
------------

The :attr:`._Cell.shading` property returns a
:class:`~docx.table.CellShading` proxy with two properties: ``fill_color``
(an |RGBColor|) and ``pattern`` (a
:class:`~docx.enum.table.WD_SHADING_PATTERN` member). Setting
``fill_color`` is all that's required for the common "solid background
color" case::

    >>> cell = table.cell(0, 0)
    >>> cell.shading.fill_color = RGBColor(0xCC, 0xFF, 0xAA)
    >>> cell.shading.pattern
    <WD_SHADING_PATTERN.CLEAR: 0>

When ``fill_color`` is assigned without an explicit ``pattern``,
``WD_SHADING_PATTERN.CLEAR`` is written as the pattern value (this is the
Word default and is what tells Word to render the fill color as a solid
background). Assigning ``None`` to ``fill_color`` removes the attribute
without disturbing ``pattern``, and vice versa.


Per-cell margins
----------------

Every cell inherits its padding from the table defaults, but individual
cells can override each edge by assigning to :attr:`._Cell.margins`::

    >>> from docx.shared import Inches
    >>> cell.margins.top = Inches(0.05)
    >>> cell.margins.start = Inches(0.08)

The four edges are ``top``, ``bottom``, ``start`` (leading edge), and
``end`` (trailing edge). Reading an edge that has no explicit override
returns ``None``, not the table default.

Two convenience methods on |_Cell| round out the API:

* :meth:`._Cell.set_margins` writes only the edges you pass; edges you
  omit are left untouched::

      >>> cell.set_margins(top=Inches(0.05), end=Inches(0.08))

* :meth:`._Cell.remove_margins` clears the ``w:tcMar`` element entirely,
  restoring full inheritance from the table defaults.

Assigning ``None`` to an individual edge removes just that edge, and when
the last edge is cleared the empty ``w:tcMar`` is removed automatically
to keep the XML tidy.


Table autofit layout
--------------------

OOXML distinguishes two interacting concepts that together decide how
column widths behave: ``w:tblLayout`` (``fixed`` vs. ``autofit``) and
``w:tblW`` (the preferred total width, which may be ``dxa``, ``pct``, or
``auto``). *python-docx* exposes three complementary properties on |Table|:

* :attr:`.Table.autofit_behavior` — a tri-state
  :class:`~docx.enum.table.WD_TABLE_AUTOFIT` enum that combines both
  concerns into a single, intention-revealing setter.
* :attr:`.Table.allow_autofit` — a narrow boolean view of the
  ``w:tblLayout`` child. Writing ``True`` removes any explicit
  ``w:tblLayout``; writing ``False`` writes ``w:type="fixed"``.
* :attr:`.Table.preferred_width` — the total table width as a |Length|
  (mapping to ``w:tblW`` with ``@w:type="dxa"``), or ``None`` when the
  preferred width is absent or expressed as a percentage.

The three :class:`~docx.enum.table.WD_TABLE_AUTOFIT` members map as
follows::

    FIXED_WIDTH           — w:tblLayout/@w:type="fixed" is written.
    AUTOFIT_TO_CONTENTS   — no w:tblLayout; w:tblW set to "auto".
    AUTOFIT_TO_WINDOW     — no w:tblLayout; w:tblW set to "5000 pct"
                            (i.e. 100% of the window).

Typical usage::

    >>> from docx.enum.table import WD_TABLE_AUTOFIT
    >>> from docx.shared import Inches
    >>> table.autofit_behavior = WD_TABLE_AUTOFIT.FIXED_WIDTH
    >>> table.preferred_width = Inches(4)

If all you care about is flipping the ``w:tblLayout`` bit without
touching the preferred width, use ``allow_autofit`` directly::

    >>> table.allow_autofit = False   # fixed layout, w:tblW untouched


Row properties
--------------

Three row-level properties are most likely to matter when laying out a
table for print:

* :attr:`._Row.height` and :attr:`._Row.height_rule` — the row's
  minimum/exact height in EMU and whether it is a minimum (``AT_LEAST``),
  exact (``EXACT``), or unconstrained (``AUTO``) value. Either property
  reads as ``None`` when no explicit value is set.

* :attr:`._Row.allow_break_across_pages` — when ``False``, the row cannot
  split across a page break; Word will push the entire row to the next
  page instead. Defaults to ``True``.

* :attr:`._Row.is_header` — when ``True``, the row repeats at the top of
  each page the table spans. Only the first N consecutive rows can be
  header rows (a Word limitation).

Example: mark the first row as a repeating header and keep every row
intact across page breaks::

    >>> header_row = table.rows[0]
    >>> header_row.is_header = True
    >>> for row in table.rows:
    ...     row.allow_break_across_pages = False


Table style conditional flags
-----------------------------

Table styles can define different formatting for the first row, last row,
first column, last column, banded rows, and banded columns. Which of those
conditional formats get applied is controlled by six flags on the table's
``w:tblLook`` element, exposed by :attr:`.Table.style_flags`:

* ``first_row``, ``last_row``, ``first_column``, ``last_column`` — enable
  the matching conditional formatting from the table style.
* ``no_horizontal_banding``, ``no_vertical_banding`` — *suppress* banding.
  That is, ``no_horizontal_banding == False`` means banded rows are
  active.

When ``w:tblLook`` is absent, every flag reads as ``False``. Writing any
flag creates ``w:tblLook`` on demand::

    >>> flags = table.style_flags
    >>> flags.first_row = True
    >>> flags.first_column = True
    >>> flags.first_row
    True
    >>> flags.no_horizontal_banding
    False

Banded rows are the Word default, so you typically only need to touch
``no_horizontal_banding`` to *suppress* row banding on a style that
normally provides it.


Cell text direction
-------------------

:attr:`._Cell.text_direction` takes a member of
:class:`~docx.enum.table.WD_TEXT_DIRECTION`. The two most common values
for rotated-heading cells are ``TB_RL`` (text reads top-to-bottom,
rotated 90 degrees clockwise) and ``BT_LR`` (bottom-to-top, rotated 90
degrees counter-clockwise)::

    >>> from docx.enum.table import WD_TEXT_DIRECTION
    >>> heading = table.cell(0, 0)
    >>> heading.text_direction = WD_TEXT_DIRECTION.TB_RL

Reading the property when no explicit direction is set returns ``None``.
Assigning ``None`` removes the ``w:textDirection`` child, restoring
inheritance.


Merged-cell introspection
-------------------------

Two properties on |_Cell| make it possible to work with merged regions
without dropping to the XML layer:

* :attr:`._Cell.is_merge_origin` is a tri-state ``bool | None``:

  * ``None`` — the cell is not part of any merged region.
  * ``True`` — the cell is the *origin* (top-left) of a merged region
    (either ``w:vMerge/@w:val="restart"`` or a horizontal-only span with
    ``w:gridSpan > 1``).
  * ``False`` — the cell is a *continuation* of a vertically merged
    region (``w:vMerge`` without an explicit ``@w:val="restart"``).

* :attr:`._Cell.merge_origin` walks up any ``w:vMerge`` continuations
  and returns the cell containing the actual content of the merge. If
  the cell is already the origin (or not merged), it returns itself.

Example: collect the distinct content cells of a table, ignoring
continuations of vertical spans::

    >>> seen = set()
    >>> content_cells = []
    >>> for row in table.rows:
    ...     for cell in row.cells:
    ...         origin = cell.merge_origin
    ...         key = id(origin._tc)
    ...         if key in seen:
    ...             continue
    ...         seen.add(key)
    ...         content_cells.append(origin)

Accessing cells via :meth:`.Table.cell` already resolves continuations
for you — the returned |_Cell| is always the origin cell. The raw
``w:tc`` elements surface only when you iterate over
``row._tr.tc_lst`` directly.
