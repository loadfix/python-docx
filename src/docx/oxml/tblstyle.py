"""Re-export shim for the shared DrawingML table-style element classes.

The DrawingML table-style vocabulary (``a:tblStyleLst`` and
descendants — ECMA-376 Part 1 §20.1.4.2) is format-neutral: the
same element tree appears in every OOXML format that hosts a
``tableStyles.xml`` resource (PowerPoint slide-masters, Word
inline graphic-frame tables, Excel chart-sheet tables).

Note this is the *DrawingML* table-style grammar (``a:`` namespace),
not WordprocessingML's table-style family (``w:tblStyle`` /
``w:tblStylePr`` in ``styles.xml`` — a distinct vocabulary that
python-docx models locally in :mod:`docx.oxml.styles`).

docx did not previously own a Python element model for the
DrawingML subtree — when it appeared on DrawingML inline tables
it was parsed as opaque ``lxml`` elements. The
``python-ooxml-shared-drawingml`` 0.4.0 surface now provides a full
descriptor-driven model, and docx adopts it verbatim here. No
docx-specific subclasses are required: the grammar as shipped by
the shared package matches docx's needs exactly.

The nine ``CT_*`` classes re-exported from this module are:

- :class:`CT_TableStyleList` — the ``a:tblStyleLst`` root.
- :class:`CT_TableStyle` — an individual ``a:tblStyle`` entry.
- :class:`CT_TablePartStyle` — the polymorphic part-style shape
  used for ``a:wholeTbl``, the four bands, the four edges, and
  the four corner cells.
- :class:`CT_TableStyleTextStyle` — ``a:tcTxStyle``.
- :class:`CT_TableStyleCellStyle` — ``a:tcStyle``.
- :class:`CT_TableCellBorderStyle` — ``a:tcBdr`` (eight border
  slots).
- :class:`CT_TableBackgroundStyle` — ``a:tblBg``.
- :class:`CT_ThemeableLineStyle` — the ``ln`` / ``lnRef`` choice
  used by each of the eight border slots.
- :class:`CT_Cell3D` — ``a:cell3D`` (bevel + light-rig + material).

Element-class registrations for the 22 tag names these classes
serve live in :mod:`docx.oxml` alongside the other
``register_element_cls`` calls.

.. versionadded:: adopt-sdml-0.4
   Table-style CT_* grammar lifted to
   ``python-ooxml-shared-drawingml`` 0.4.0.
"""

from __future__ import annotations

from ooxml_shared_drawingml.tblstyle import (
    CT_Cell3D,
    CT_TableBackgroundStyle,
    CT_TableCellBorderStyle,
    CT_TablePartStyle,
    CT_TableStyle,
    CT_TableStyleCellStyle,
    CT_TableStyleList,
    CT_TableStyleTextStyle,
    CT_ThemeableLineStyle,
)

__all__ = [
    "CT_Cell3D",
    "CT_TableBackgroundStyle",
    "CT_TableCellBorderStyle",
    "CT_TablePartStyle",
    "CT_TableStyle",
    "CT_TableStyleCellStyle",
    "CT_TableStyleList",
    "CT_TableStyleTextStyle",
    "CT_ThemeableLineStyle",
]
