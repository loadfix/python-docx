"""Page-layout shorthands — multi-column newspaper-style sections.

Closes #286.

Word's section-column model is a small but fiddly piece of OOXML to
drive directly: a multi-column run of body content is bracketed by two
*continuous* section breaks, the first carrying a ``<w:cols w:num="N"
w:equalWidth="1"/>`` element on its terminating section properties and
the second carrying a single-column ``<w:cols w:num="1"/>`` so the body
returns to a normal one-column layout for whatever follows. Authors
almost always want the simple shape — "make these next few paragraphs
flow newspaper-style across N equal columns" — without learning the
section-break dance.

This module exposes two helpers built entirely on python-docx's public
API (:meth:`Document.add_section`, :meth:`Section.set_columns`)::

    from docx import Document
    from docx.kit import layout

    doc = Document()
    layout.multi_column(doc, columns=2, equal_width=True)
    doc.add_paragraph("This text flows across two columns like a newspaper.")
    doc.add_paragraph("More text continues to flow.")
    layout.end_multi_column(doc)
    doc.save("out.docx")

Both helpers add a *continuous* section break (i.e. the new section
starts on the same page rather than forcing a page break — that is
what newspaper-style columns mean):

* :func:`multi_column` opens a fresh section, applies the requested
  column geometry, and returns the |Section| object so callers can
  tweak it further.
* :func:`end_multi_column` closes the multi-column run by opening a
  second continuous section that resets back to a single column. It
  also returns the new |Section|.

Equal-width is the simple path. Pass ``widths_in=[2.0, 4.0]`` (in
inches) to emit per-column ``<w:col w:w="..."/>`` children for an
unequal layout — a wide editorial column next to a narrow sidebar, for
example. ``spacing_in`` controls the gutter between equal-width
columns; the value is in inches and converts to twentieths of a point
(twips) on serialisation.

Implementation notes:

* The helpers compose the existing public surface — they call
  :meth:`Document.add_section` and :meth:`Section.set_columns` and do
  *no* ``oxml`` / ``etree`` reach-down. If the underlying public API
  changes shape, these helpers move with it.
* :func:`multi_column` validates that ``columns >= 1`` and that, when
  ``widths_in`` is supplied, ``len(widths_in) == columns``. Surfacing
  a clean :class:`ValueError` here saves a confusing failure mode at
  save-time when Word would silently ignore a malformed ``w:cols``.
* ``equal_width=True`` (the default) is mutually exclusive with
  ``widths_in``; passing both raises :class:`ValueError` rather than
  silently picking one. Pass ``widths_in`` *or* ``equal_width=True``,
  not both.
* :func:`end_multi_column` is idempotent in spirit — calling it
  without a preceding :func:`multi_column` still emits a clean
  single-column continuous section, which is harmless. The helper is
  not strict about pairing because Word tolerates extra single-column
  section breaks.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Sequence, Union

from docx.enum.section import WD_SECTION
from docx.shared import Inches

if TYPE_CHECKING:
    from docx.document import Document
    from docx.section import Section


# -- Default gutter between columns. 0.5 inch matches Word's "Two" and
# -- "Three" column presets in the Page Layout ribbon, and is the
# -- conventional newspaper-style choice. Twentieths of a point: 0.5
# -- inch == 720 twips.
_DEFAULT_SPACING_IN = 0.5


def _coerce_widths(widths_in):
    # type: (Optional[Sequence[Union[float, int]]]) -> Optional[List[Inches]]
    """Convert a sequence of inch values to a list of |Inches| Length objects.

    Returns ``None`` when ``widths_in`` is ``None`` so callers can
    forward the result straight through to :meth:`Section.set_columns`,
    whose ``widths`` parameter is also ``Optional``.
    """
    if widths_in is None:
        return None
    return [Inches(float(w)) for w in widths_in]


def multi_column(
    document,
    columns=2,
    equal_width=True,
    spacing_in=_DEFAULT_SPACING_IN,
    widths_in=None,
):
    # type: (Document, int, bool, float, Optional[Sequence[Union[float, int]]]) -> Section
    """Open a continuous multi-column section at the end of ``document``.

    Appends a continuous section break and applies a ``<w:cols>``
    configuration for ``columns`` columns to the new section. All
    paragraphs added after this call (until :func:`end_multi_column`)
    flow newspaper-style across the columns.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    columns
        Number of columns. Must be ``>= 1``. Defaults to ``2``.
    equal_width
        When ``True`` (the default), all columns are equal-width and
        ``<w:cols w:equalWidth="1"/>`` is emitted. Mutually exclusive
        with ``widths_in``.
    spacing_in
        Gutter between equal-width columns, in inches. Maps to
        ``<w:cols w:space="..."/>`` in twentieths of a point on
        serialisation. Defaults to ``0.5`` (Word's preset). Ignored
        when ``widths_in`` is supplied (per-column space is set via
        ``<w:col w:space="..."/>`` in that case, which the kit does
        not surface — use :meth:`SectionColumns.__getitem__` if needed).
    widths_in
        Optional sequence of per-column widths in inches. When
        supplied, ``len(widths_in)`` must equal ``columns``, one
        ``<w:col w:w="..."/>`` child is emitted per width, and
        ``equal_width`` is forced to |False|. Mutually exclusive with
        the default ``equal_width=True``.

    Returns
    -------
    Section
        The newly-appended :class:`Section`. Callers can tweak it
        further (e.g. set a column separator) before adding content.

    Raises
    ------
    ValueError
        If ``columns < 1``, if ``widths_in`` is supplied with a length
        that does not equal ``columns``, or if both ``equal_width=True``
        and ``widths_in`` are supplied.
    """
    if columns < 1:
        raise ValueError(
            "columns must be >= 1; got %r" % (columns,)
        )
    if widths_in is not None and len(widths_in) != columns:
        raise ValueError(
            "widths_in must have exactly `columns` entries; got "
            "%d widths for columns=%d" % (len(widths_in), columns)
        )
    # -- The default for ``equal_width`` is True; callers who supply
    # -- ``widths_in`` must pass ``equal_width=False`` explicitly so the
    # -- intent is unambiguous.
    if widths_in is not None and equal_width:
        raise ValueError(
            "equal_width=True is mutually exclusive with widths_in; pass "
            "equal_width=False (or omit it) when supplying widths_in"
        )

    # -- Open a continuous section break.  ``WD_SECTION.CONTINUOUS``
    # -- means "start the new section on the same page" — Word's
    # -- newspaper-column convention.
    section = document.add_section(WD_SECTION.CONTINUOUS)

    widths = _coerce_widths(widths_in)
    if widths is not None:
        # -- Per-column widths path.  ``set_columns`` will set
        # -- equal_width=False implicitly when widths is supplied; we
        # -- pass equal_width=False explicitly for clarity.  Column
        # -- spacing is per-column in this mode (set via Column.space
        # -- by the caller); the section-level ``w:space`` is not
        # -- meaningful, so we omit it.
        section.set_columns(
            count=columns,
            equal_width=False,
            widths=widths,
        )
    else:
        # -- Equal-width path.  Convert ``spacing_in`` (inches) to a
        # -- |Length| once and let ``set_columns`` write it to
        # -- ``w:cols/@w:space`` (twentieths of a point on the wire).
        section.set_columns(
            count=columns,
            equal_width=equal_width,
            space=Inches(spacing_in),
        )
    return section


def end_multi_column(document):
    # type: (Document) -> Section
    """Close a multi-column section back to a single column.

    Appends a continuous section break and resets the new section to a
    single equal-width column. Subsequent paragraphs flow normally —
    one column across the page width.

    The helper does not require a matching :func:`multi_column` call:
    appending a redundant single-column continuous break is harmless,
    so callers can use the helper defensively without tracking section
    state.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.

    Returns
    -------
    Section
        The newly-appended single-column :class:`Section`.
    """
    section = document.add_section(WD_SECTION.CONTINUOUS)
    # -- Reset to the python-docx default single-column shape.  We do
    # -- *not* pass ``space`` here: a single column has no gutter.
    section.set_columns(count=1, equal_width=True)
    return section


__all__ = [
    "multi_column",
    "end_multi_column",
]
