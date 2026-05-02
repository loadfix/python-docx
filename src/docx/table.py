"""The |Table| object and related proxy classes."""

from __future__ import annotations

import os
import warnings
from typing import IO, TYPE_CHECKING, cast, overload
from collections.abc import Iterator

from typing_extensions import TypeAlias

from docx.blkcntnr import BlockItemContainer
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import (
    WD_BORDER_STYLE,
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_SHADING_PATTERN,
    WD_TABLE_AUTOFIT,
)
from docx.oxml.simpletypes import ST_Merge
from docx.oxml.table import CT_Tbl, CT_TblGridCol
from docx.shared import Emu, Inches, Parented, Pt, RGBColor, StoryChild, lazyproperty
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.table import (
        WD_ROW_HEIGHT_RULE,
        WD_TABLE_ALIGNMENT,
        WD_TABLE_DIRECTION,
        WD_TEXT_DIRECTION,
    )
    from docx.oxml.table import (
        CT_Border,
        CT_Row,
        CT_Shd,
        CT_TblBorders,
        CT_TblCellMar,
        CT_TblLook,
        CT_TblPr,
        CT_Tc,
        CT_TcBorders,
        CT_TcMar,
    )
    from docx.oxml.text.paragraph import CT_P
    from docx.shared import Length
    from docx.styles.style import (
        ParagraphStyle,
        _TableStyle,  # pyright: ignore[reportPrivateUsage]
    )
    from docx.tracked_changes import FormattingChange

TableParent: TypeAlias = "Table | _Columns | _Rows"


class Table(StoryChild):
    """Proxy class for a WordprocessingML ``<w:tbl>`` element."""

    def __init__(self, tbl: CT_Tbl, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._element = tbl
        self._tbl = tbl

    def delete(self) -> None:
        """Remove this table from the document.

        The table element is removed from its parent. After calling this method,
        this |Table| object is "defunct" and should not be used further.

        .. versionadded:: 1.3.0.dev0
        """
        tbl = self._tbl
        parent = tbl.getparent()
        if parent is None:
            return
        parent.remove(tbl)

    @property
    def spans_page_break(self) -> bool:
        """``True`` when this table contains one or more rendered page-breaks.

        Detects ``w:lastRenderedPageBreak`` elements anywhere inside the
        table's cell content. This marker is written by Word when it renders
        the document; programmatically-created tables typically have none
        until Word opens and re-saves the file.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.ns import qn

        return self._tbl.find(".//" + qn("w:lastRenderedPageBreak")) is not None

    def add_column(self, width: Length):
        """Return a |_Column| object of `width`, newly added rightmost to the table."""
        tblGrid = self._tbl.tblGrid
        gridCol = tblGrid.add_gridCol()
        gridCol.w = width
        for tr in self._tbl.tr_lst:
            tc = tr.add_tc()
            tc.width = width
        return _Column(gridCol, self)

    def add_row(self, source_row: _Row | None = None):
        """Return a |_Row| instance, newly added bottom-most to the table.

        When a preceding row is present, its ``w:trPr/w:cnfStyle`` element is
        copied onto the new row so conditional-formatting (banding,
        first/last-row emphasis) flows to the new row as Word itself would.
        (upstream#306)

        When `source_row` is provided, the new row is cloned from it, preserving
        ``w:tc`` / ``w:trPr`` / run-level formatting. Any ``w:id`` attributes
        nested inside the cloned row are stripped so they do not collide with
        bookmark or comment IDs elsewhere in the document. (upstream#1189, #205)

        .. versionadded:: 1.3.0.dev0
           ``source_row`` parameter.
        """
        from copy import deepcopy

        from docx.oxml.ns import qn

        tbl = self._tbl

        if source_row is not None:
            # -- clone path: deepcopy source tr and append as a sibling --
            new_tr = _clone_tr(source_row._tr)
            tbl.append(new_tr)
            return _Row(new_tr, self)

        # -- capture the prior row (if any) before we append --
        preceding_tr = tbl.tr_lst[-1] if tbl.tr_lst else None
        tr = tbl.add_tr()
        for gridCol in tbl.tblGrid.gridCol_lst:
            tc = tr.add_tc()
            if gridCol.w is not None:
                tc.width = gridCol.w

        if preceding_tr is not None:
            prior_trPr = preceding_tr.trPr
            if prior_trPr is not None:
                prior_cnfStyle = prior_trPr.find(qn("w:cnfStyle"))
                if prior_cnfStyle is not None:
                    new_trPr = tr.get_or_add_trPr()
                    # -- remove any existing cnfStyle to avoid duplicates --
                    existing = new_trPr.find(qn("w:cnfStyle"))
                    if existing is not None:
                        new_trPr.remove(existing)
                    # -- cnfStyle is the first child in w:trPr (see _tag_seq) --
                    new_trPr.insert(0, deepcopy(prior_cnfStyle))

        return _Row(tr, self)

    def insert_row(self, index: int) -> _Row:
        """Return a new |_Row| inserted at `index` in this table's rows.

        Creates a blank row with one ``w:tc`` per column, matching each
        ``w:gridCol`` width. `index` behaves like a Python list insert: an
        `index` equal to the current row count appends the row; a negative
        `index` counts from the end. Existing rows at or after `index` are
        shifted down by one.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.parser import OxmlElement

        tbl = self._tbl
        tr_lst = tbl.tr_lst
        row_count = len(tr_lst)
        if index < 0:
            index += row_count
        if index < 0 or index > row_count:
            raise IndexError(f"row index [{index}] is out of range")

        new_tr = cast("CT_Row", OxmlElement("w:tr"))
        for gridCol in tbl.tblGrid.gridCol_lst:
            tc = new_tr.add_tc()
            if gridCol.w is not None:
                tc.width = gridCol.w

        if index == row_count:
            tbl.append(new_tr)
        else:
            tr_lst[index].addprevious(new_tr)
        return _Row(new_tr, self)

    def delete_column(self, index: int) -> None:
        """Remove the column at `index` from this table.

        Removes the matching ``w:gridCol`` from ``w:tblGrid`` and the
        corresponding ``w:tc`` from each row. For a cell that horizontally
        spans the deleted column (``w:gridSpan > 1``), the gridSpan is
        decremented by 1 rather than orphaning the cell.

        .. versionadded:: 1.3.0.dev0
        """
        tbl = self._tbl
        gridCol_lst = tbl.tblGrid.gridCol_lst
        col_count = len(gridCol_lst)
        if index < 0:
            index += col_count
        if index < 0 or index >= col_count:
            raise IndexError(f"column index [{index}] is out of range")

        # -- remove the gridCol entry --
        gridCol = gridCol_lst[index]
        tbl.tblGrid.remove(gridCol)

        # -- for each row, remove or shrink the cell covering the column --
        for tr in tbl.tr_lst:
            # -- gridBefore/gridAfter shift the tc grid_offset; account for
            # -- them so a delete at `index` that falls inside the
            # -- gridBefore/gridAfter omitted region is a no-op for this row.
            current_offset = tr.grid_before
            if index < current_offset:
                continue
            removed = False
            for tc in list(tr.tc_lst):
                span = tc.grid_span
                if current_offset <= index < current_offset + span:
                    if span == 1:
                        tr.remove(tc)
                    else:
                        tc.grid_span = span - 1
                    removed = True
                    break
                current_offset += span
            # -- if no tc covers the index (gridAfter region), nothing to do --
            if not removed:
                continue

    def insert_paragraph_before(
        self, text: str = "", style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this table.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that paragraph style is assigned to the new paragraph.
        The new paragraph is inserted as a sibling of this table in its parent element.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.parser import OxmlElement

        new_p = cast("CT_P", OxmlElement("w:p"))
        self._tbl.addprevious(new_p)
        paragraph = Paragraph(new_p, self._parent)
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def insert_paragraph_after(
        self, text: str = "", style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly after this table.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that paragraph style is assigned to the new paragraph.
        The new paragraph is inserted as a sibling of this table in its parent element.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.parser import OxmlElement

        new_p = cast("CT_P", OxmlElement("w:p"))
        self._tbl.addnext(new_p)
        paragraph = Paragraph(new_p, self._parent)
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def insert_table_before(
        self,
        rows: int,
        cols: int,
        style: str | _TableStyle | None = None,
        width: Length | None = None,
    ) -> Table:
        """Return a new table with `rows` rows and `cols` cols, inserted directly
        before this table.

        If `style` is supplied, that style is assigned to the new table. The new
        table is inserted as a sibling of this table in its parent element. `width`
        is an optional total table width; if not provided it defaults to 6 inches.

        .. versionadded:: 1.3.0.dev0
        """
        table_width = width if width is not None else Inches(6)
        tbl = CT_Tbl.new_tbl(rows, cols, table_width)
        self._tbl.addprevious(tbl)
        table = Table(tbl, self._parent)
        if style is not None:
            table.style = style
        return table

    def insert_table_after(
        self,
        rows: int,
        cols: int,
        style: str | _TableStyle | None = None,
        width: Length | None = None,
    ) -> Table:
        """Return a new table with `rows` rows and `cols` cols, inserted directly
        after this table.

        If `style` is supplied, that style is assigned to the new table. The new
        table is inserted as a sibling of this table in its parent element. `width`
        is an optional total table width; if not provided it defaults to 6 inches.

        .. versionadded:: 1.3.0.dev0
        """
        table_width = width if width is not None else Inches(6)
        tbl = CT_Tbl.new_tbl(rows, cols, table_width)
        self._tbl.addnext(tbl)
        table = Table(tbl, self._parent)
        if style is not None:
            table.style = style
        return table

    @property
    def alignment(self) -> WD_TABLE_ALIGNMENT | None:
        """Read/write.

        A member of :ref:`WdRowAlignment` or None, specifying the positioning of this
        table between the page margins. |None| if no setting is specified, causing the
        effective value to be inherited from the style hierarchy.
        """
        return self._tblPr.alignment

    @alignment.setter
    def alignment(self, value: WD_TABLE_ALIGNMENT | None):
        self._tblPr.alignment = value

    @property
    def autofit(self) -> bool:
        """|True| if column widths can be automatically adjusted to improve the fit of
        cell contents.

        |False| if table layout is fixed. Column widths are adjusted in either case if
        total column width exceeds page width. Read/write boolean.

        Backward-compatible alias. For richer control over the autofit behavior use
        :attr:`autofit_behavior` (which returns a :class:`WD_TABLE_AUTOFIT` member)
        and :attr:`allow_autofit` (a narrow bool-only view onto `w:tblLayout`).
        """
        return self._tblPr.autofit

    @autofit.setter
    def autofit(self, value: bool):
        self._tblPr.autofit = value

    @property
    def allow_autofit(self) -> bool:
        """|True| when the table layout type is ``"autofit"`` (or absent).

        |False| when the table has an explicit ``<w:tblLayout w:type="fixed"/>``
        child; in that case Word keeps each column at exactly its declared width.

        Read/write. Setting this property only affects the ``w:tblLayout`` child;
        it does not alter ``w:tblW`` (use :attr:`preferred_width` or
        :attr:`autofit_behavior` for that). Assigning |True| removes any explicit
        ``w:tblLayout`` rather than writing ``w:type="autofit"`` so the table falls
        back to the OOXML default.

        .. versionadded:: 1.3.0.dev0
        """
        return self._tblPr.autofit

    @allow_autofit.setter
    def allow_autofit(self, value: bool):
        tblPr = self._tblPr
        if value:
            # -- default is autofit; remove the explicit element rather than write it --
            tblPr._remove_tblLayout()  # pyright: ignore[reportPrivateUsage]
        else:
            tblPr.get_or_add_tblLayout().type = "fixed"

    @property
    def autofit_behavior(self) -> WD_TABLE_AUTOFIT:
        """The autofit behavior of this table as a |WD_TABLE_AUTOFIT| member.

        Combines the semantics of the ``w:tblLayout/@w:type`` and
        ``w:tblW/@w:type`` attributes:

        - ``FIXED_WIDTH`` if ``w:tblLayout/@w:type="fixed"``.
        - ``AUTOFIT_TO_WINDOW`` if layout is autofit and ``w:tblW/@w:type="pct"``.
        - ``AUTOFIT_TO_CONTENTS`` otherwise (the OOXML default).

        Read/write. Assigning a new value rewrites both ``w:tblLayout`` and
        ``w:tblW`` to a consistent state.

        .. versionadded:: 1.3.0.dev0
        """
        tblPr = self._tblPr
        if not tblPr.autofit:
            return WD_TABLE_AUTOFIT.FIXED_WIDTH
        tblW = tblPr.tblW
        if tblW is not None and tblW.type == "pct":
            return WD_TABLE_AUTOFIT.AUTOFIT_TO_WINDOW
        return WD_TABLE_AUTOFIT.AUTOFIT_TO_CONTENTS

    @autofit_behavior.setter
    def autofit_behavior(self, value: WD_TABLE_AUTOFIT):
        tblPr = self._tblPr
        if value == WD_TABLE_AUTOFIT.FIXED_WIDTH:
            tblPr.get_or_add_tblLayout().type = "fixed"
            return
        if value == WD_TABLE_AUTOFIT.AUTOFIT_TO_WINDOW:
            tblPr._remove_tblLayout()  # pyright: ignore[reportPrivateUsage]
            tblPr.set_tblW(5000, "pct")
            return
        if value == WD_TABLE_AUTOFIT.AUTOFIT_TO_CONTENTS:
            tblPr._remove_tblLayout()  # pyright: ignore[reportPrivateUsage]
            tblPr.set_tblW(0, "auto")
            return
        raise ValueError(f"unsupported WD_TABLE_AUTOFIT value: {value!r}")

    @property
    def preferred_width(self) -> Length | None:
        """The preferred total width of this table in EMU, or |None|.

        Maps to ``w:tblPr/w:tblW`` with ``@w:type="dxa"``. Returns |None| when
        ``w:tblW`` is absent or when its ``w:type`` is not ``"dxa"`` (e.g. when
        the table width is declared as a percentage or ``auto``).

        Read/write. Assigning |None| removes ``w:tblW`` entirely.

        .. versionadded:: 1.3.0.dev0
        """
        return self._tblPr.preferred_width

    @preferred_width.setter
    def preferred_width(self, value: Length | None):
        self._tblPr.preferred_width = value

    @property
    def alt_text(self) -> str | None:
        """The table's accessibility "caption" (alt-text title), or |None|.

        Maps to ``w:tblPr/w:tblCaption/@w:val``. This is the short
        accessibility label Word exposes in the **Table Properties -> Alt Text**
        dialog as *Title*. Assigning |None| removes the ``w:tblCaption`` child.

        .. versionadded:: 1.3.0.dev0
        """
        tblCaption = self._tblPr.tblCaption
        return None if tblCaption is None else tblCaption.val

    @alt_text.setter
    def alt_text(self, value: str | None):
        tblPr = self._tblPr
        tblPr._remove_tblCaption()  # pyright: ignore[reportPrivateUsage]
        if value is None:
            return
        tblPr._add_tblCaption().val = value  # pyright: ignore[reportPrivateUsage]

    @property
    def alt_description(self) -> str | None:
        """The table's accessibility description (long alt-text), or |None|.

        Maps to ``w:tblPr/w:tblDescription/@w:val``. This is the longer
        accessibility description Word exposes in the **Table Properties ->
        Alt Text** dialog. Assigning |None| removes the ``w:tblDescription``
        child.

        .. versionadded:: 1.3.0.dev0
        """
        tblDescription = self._tblPr.tblDescription
        return None if tblDescription is None else tblDescription.val

    @alt_description.setter
    def alt_description(self, value: str | None):
        tblPr = self._tblPr
        tblPr._remove_tblDescription()  # pyright: ignore[reportPrivateUsage]
        if value is None:
            return
        tblPr._add_tblDescription().val = value  # pyright: ignore[reportPrivateUsage]

    @property
    def indent(self) -> Length | None:
        """The left-indent of this table as an EMU |Length|, or |None|.

        Maps to ``w:tblPr/w:tblInd``. Returns |None| when the ``w:tblInd``
        child is absent or when its ``w:type`` is not ``"dxa"`` (e.g. when
        the indent is expressed as a percentage). Read/write. Assigning
        |None| removes ``w:tblInd`` entirely.

        .. versionadded:: 1.3.0.dev0
        """
        tblInd = self._tblPr.tblInd
        if tblInd is None:
            return None
        return tblInd.width

    @indent.setter
    def indent(self, value: Length | None):
        tblPr = self._tblPr
        if value is None:
            tblPr._remove_tblInd()  # pyright: ignore[reportPrivateUsage]
            return
        tblPr.get_or_add_tblInd().width = value

    # -- upstream-compat alias (upstream#586) --
    @property
    def left_indent(self) -> Length | None:
        """Alias of :attr:`indent` for upstream-compat (upstream#586).

        .. versionadded:: 1.3.0.dev0
        """
        return self.indent

    @left_indent.setter
    def left_indent(self, value: Length | None):
        self.indent = value

    @property
    def cell_margins(self) -> TableCellMargins:
        """Read-only. |TableCellMargins| proxy for table-level default cell margins.

        Always returns a |TableCellMargins| object. When no ``w:tblCellMar``
        element is present, each edge reads as |None|; assigning to an edge
        creates the ``w:tblPr/w:tblCellMar`` structure on demand. Mirrors
        :attr:`_Cell.margins` but operates at the table level, setting the
        default cell margins inherited by every cell that doesn't override
        them via ``w:tcMar``.

        .. versionadded:: 1.3.0.dev0
        """
        return TableCellMargins(self._tbl)

    def merged_cell_ranges(self) -> list[tuple[int, int, int, int]]:
        """Return a list of rectangular merged-cell regions in this table.

        Each entry is a ``(row_start, row_end_exclusive, col_start,
        col_end_exclusive)`` 4-tuple describing a merged region. A cell is
        considered "merged" when it spans more than one grid column
        (``w:gridSpan > 1``) or when it participates in a vertical span
        (``w:vMerge="restart"``). The result is sorted by ``(row_start,
        col_start)`` and excludes single-cell (1x1) regions.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.simpletypes import ST_Merge

        ranges: list[tuple[int, int, int, int]] = []
        for tc in self._tbl.iter_tcs():
            vMerge = tc.vMerge
            grid_span = tc.grid_span
            # -- only origin cells: vMerge in {None, "restart"} --
            if vMerge == ST_Merge.CONTINUE:
                continue
            # -- skip "plain" cells: grid_span 1 and no vertical span --
            if grid_span <= 1 and vMerge is None:
                continue
            row_start = tc.top
            row_end = tc.bottom
            col_start = tc.left
            col_end = tc.right
            ranges.append((row_start, row_end, col_start, col_end))
        ranges.sort(key=lambda r: (r[0], r[2]))
        return ranges

    def split(self, before_row: int) -> tuple[Paragraph, Table]:
        """Split this table in two at `before_row`.

        A new ``w:p`` and a new ``w:tbl`` are inserted immediately after this
        table. The rows from `before_row` onward are moved from this table
        into the newly-created table; this table retains the rows above
        `before_row`. Returns a ``(paragraph, new_table)`` tuple.

        Raises |ValueError| if `before_row` is out of range, is the first
        row (``before_row == 0``), or is beyond the last row.

        .. versionadded:: 1.3.0.dev0
        """
        from copy import deepcopy

        from docx.oxml.parser import OxmlElement
        from docx.oxml.table import CT_Tbl as _CT_Tbl

        tr_lst = list(self._tbl.tr_lst)
        row_count = len(tr_lst)
        if before_row <= 0 or before_row >= row_count:
            raise ValueError(
                f"before_row must be in range [1, {row_count - 1}]; got {before_row}"
            )
        # -- build new w:tbl preserving the original tblPr and tblGrid --
        new_tbl = cast(_CT_Tbl, OxmlElement("w:tbl"))
        new_tbl.append(deepcopy(self._tbl.tblPr))
        new_tbl.append(deepcopy(self._tbl.tblGrid))
        # -- move rows from before_row..end into new_tbl --
        for tr in tr_lst[before_row:]:
            self._tbl.remove(tr)
            new_tbl.append(tr)
        # -- insert the separator paragraph and the new table immediately
        # -- after this table in its parent (document order). --
        new_p = cast("CT_P", OxmlElement("w:p"))
        self._tbl.addnext(new_tbl)
        self._tbl.addnext(new_p)
        paragraph = Paragraph(new_p, self._parent)
        new_table = Table(new_tbl, self._parent)
        return paragraph, new_table

    @property
    def borders(self) -> TableBorders:
        """Read-only. |TableBorders| object providing access to table border properties.

        Always returns a |TableBorders| object; setting border properties on it will
        create the required XML elements on demand.

        .. versionadded:: 1.3.0.dev0
        """
        return TableBorders(self._tbl)

    @property
    def style_flags(self) -> TableStyleFlags:
        """Read-only. |TableStyleFlags| access to the `w:tblLook` conditional-style flags.

        Always returns a |TableStyleFlags| object. Reading a flag when the
        `w:tblLook` child is absent yields |False|; writing to any flag creates
        the `w:tblLook` element on demand.

        .. versionadded:: 1.3.0.dev0
        """
        return TableStyleFlags(self._tbl)

    def set_borders(
        self,
        top: bool = False,
        bottom: bool = False,
        left: bool = False,
        right: bool = False,
        inside_h: bool = False,
        inside_v: bool = False,
        style: WD_BORDER_STYLE = WD_BORDER_STYLE.SINGLE,
        width: Length | None = None,
        color: RGBColor | None = None,
    ) -> None:
        """Convenience method to set multiple table borders at once.

        Each boolean parameter controls whether that border edge is enabled.
        Enabled borders use the specified `style`, `width`, and `color`.
        Disabled borders are set to ``WD_BORDER_STYLE.NONE``.

        Example for APA 7 tables (horizontal-only borders)::

            table.set_borders(top=True, bottom=True, inside_h=True)

        .. versionadded:: 1.3.0.dev0
        """
        border_width = width if width is not None else Pt(0.5)
        border_color = color if color is not None else RGBColor(0, 0, 0)
        borders = self.borders
        for attr, enabled in [
            ("top", top),
            ("bottom", bottom),
            ("left", left),
            ("right", right),
            ("inside_h", inside_h),
            ("inside_v", inside_v),
        ]:
            border = getattr(borders, attr)
            if enabled:
                border.style = style
                border.width = border_width
                border.color = border_color
            else:
                border.style = WD_BORDER_STYLE.NONE
                border.width = None
                border.color = None

    def cell(self, row_idx: int, col_idx: int) -> _Cell:
        """|_Cell| at `row_idx`, `col_idx` intersection.

        (0, 0) is the top, left-most cell.
        """
        cell_idx = col_idx + (row_idx * self._column_count)
        return self._cells[cell_idx]

    def column_cells(self, column_idx: int) -> list[_Cell]:
        """Sequence of cells in the column at `column_idx` in this table."""
        cells = self._cells
        idxs = range(column_idx, len(cells), self._column_count)
        return [cells[idx] for idx in idxs]

    @lazyproperty
    def columns(self):
        """|_Columns| instance representing the sequence of columns in this table."""
        return _Columns(self._tbl, self)

    def row_cells(self, row_idx: int) -> list[_Cell]:
        """DEPRECATED: Use `table.rows[row_idx].cells` instead.

        Sequence of cells in the row at `row_idx` in this table.
        """
        warnings.warn(
            "Table.row_cells() is deprecated, use table.rows[row_idx].cells instead",
            DeprecationWarning,
            stacklevel=2,
        )
        column_count = self._column_count
        start = row_idx * column_count
        end = start + column_count
        return self._cells[start:end]

    @lazyproperty
    def rows(self) -> _Rows:
        """|_Rows| instance containing the sequence of rows in this table."""
        return _Rows(self._tbl, self)

    @property
    def style(self) -> _TableStyle | None:
        """|_TableStyle| object representing the style applied to this table.

        Read/write. The default table style for the document (often `Normal Table`) is
        returned if the table has no directly-applied style. Assigning |None| to this
        property removes any directly-applied table style causing it to inherit the
        default table style of the document.

        Note that the style name of a table style differs slightly from that displayed
        in the user interface; a hyphen, if it appears, must be removed. For example,
        `Light Shading - Accent 1` becomes `Light Shading Accent 1`.
        """
        style_id = self._tbl.tblStyle_val
        return cast("_TableStyle | None", self.part.get_style(style_id, WD_STYLE_TYPE.TABLE))

    @style.setter
    def style(self, style_or_name: _TableStyle | str | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.TABLE)
        self._tbl.tblStyle_val = style_id

    @property
    def table(self):
        """Provide child objects with reference to the |Table| object they belong to,
        without them having to know their direct parent is a |Table| object.

        This is the terminus of a series of `parent._table` calls from an arbitrary
        child through its ancestors.
        """
        return self

    @property
    def formatting_change(self) -> FormattingChange | None:
        """|FormattingChange| proxy for this table's `w:tblPrChange`, or |None|.

        Returns a read-only :class:`~docx.tracked_changes.FormattingChange`
        exposing the prior table properties (`w:tblPr`) when this table carries
        a `w:tblPr/w:tblPrChange` tracked-revision marker. Returns |None| when
        the table has no `w:tblPrChange` child.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.tracked_changes import FormattingChange

        tblPrChange = self._tblPr.tblPrChange
        if tblPrChange is None:
            return None
        return FormattingChange(tblPrChange)

    @property
    def stable_id(self) -> str:
        """A 16-character hex stable identifier for this table.

        The ID is derived from the table's position within its parent and the
        concatenated text of its cells. It is stable across save/reload *when
        the table keeps the same position with the same cell content*; it
        changes if the table is reordered or its text is edited. The value is
        recomputed on each access and never persisted on the element.

        The ``w:tbl`` element itself has no ``@w:rsidR``, so only the
        structural hash is used. For more robust cross-session tracking,
        compare the table's content directly.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.ids import compute_stable_id

        text = "\n".join(cell.text for cell in self._cells)
        return compute_stable_id(self._tbl, text)

    @property
    def table_direction(self) -> WD_TABLE_DIRECTION | None:
        """Member of :ref:`WdTableDirection` indicating cell-ordering direction.

        For example: `WD_TABLE_DIRECTION.LTR`. |None| indicates the value is inherited
        from the style hierarchy.
        """
        return cast("WD_TABLE_DIRECTION | None", self._tbl.bidiVisual_val)

    @table_direction.setter
    def table_direction(self, value: WD_TABLE_DIRECTION | None):
        self._element.bidiVisual_val = value

    @property
    def direction(self) -> WD_TABLE_DIRECTION | None:
        """Alias of :attr:`table_direction` for upstream-compat (upstream#1227).

        The property was renamed to ``table_direction`` in this fork; this
        alias preserves the original name so callers written against upstream
        python-docx continue to work.
        """
        return self.table_direction

    @direction.setter
    def direction(self, value: WD_TABLE_DIRECTION | None):
        self.table_direction = value

    @property
    def _cells(self) -> list[_Cell]:
        """A sequence of |_Cell| objects, one for each cell of the layout grid.

        If the table contains a span, one or more |_Cell| object references are
        repeated. Per-row ``w:gridBefore`` / ``w:gridAfter`` allowances are
        honoured so the returned sequence is always length
        ``rows * _column_count`` and ``cell(r, c)`` / ``column_cells(c)``
        arithmetic lines up with the declared grid. Omitted-cell slots are
        filled with a placeholder |_Cell| (borrowing the cell directly above
        where possible, else the nearest real ``w:tc`` in the same row) so
        callers never encounter |None|. (upstream#939, #1367, #1334, #1193)
        """
        col_count = self._column_count
        cells: list[_Cell] = []
        for tr in self._tbl.tr_lst:
            row_start = len(cells)
            grid_before = tr.grid_before
            grid_after = tr.grid_after

            # -- build row cells first so we can fill gridBefore placeholders
            # -- by looking at the row's own real cells when no row-above exists.
            row_real_cells: list[_Cell] = []
            for tc in tr.tc_lst:
                for grid_span_idx in range(tc.grid_span):
                    if tc.vMerge == ST_Merge.CONTINUE:
                        # -- continuation cell: delegate to the cell one row above.
                        # -- If there's no preceding row (orphan continuation in
                        # -- the first row, malformed document), fall back to
                        # -- treating this tc as its own cell.
                        if row_start >= col_count:
                            # -- same column in the preceding row --
                            col = grid_before + len(row_real_cells)
                            if col < col_count:
                                row_real_cells.append(cells[row_start - col_count + col])
                                continue
                        row_real_cells.append(_Cell(tc, self))
                    elif grid_span_idx > 0:
                        row_real_cells.append(row_real_cells[-1])
                    else:
                        row_real_cells.append(_Cell(tc, self))

            # -- choose a placeholder for omitted slots: prefer the cell
            # -- directly above, fall back to the row's first/last real cell,
            # -- or synthesise from the first tc if the row has none. --
            def _placeholder(col: int, fallback: _Cell | None) -> _Cell:
                if row_start >= col_count and col < col_count:
                    return cells[row_start - col_count + col]
                if fallback is not None:
                    return fallback
                # -- absolute last resort: wrap the first tc of the row --
                first_tc = tr.tc_lst[0] if tr.tc_lst else None
                if first_tc is not None:
                    return _Cell(first_tc, self)
                # -- row is empty and no row above; return a no-op placeholder
                # -- by wrapping the table element itself would be unsafe, so
                # -- fall back to repeating adjacent placeholder if any. --
                if cells:
                    return cells[-1]
                raise AssertionError("cannot synthesise placeholder for empty row")

            # -- gridBefore slots --
            first_real = row_real_cells[0] if row_real_cells else None
            for i in range(grid_before):
                cells.append(_placeholder(i, first_real))

            # -- real cells --
            cells.extend(row_real_cells)

            # -- gridAfter slots --
            last_real = row_real_cells[-1] if row_real_cells else first_real
            start_after = grid_before + len(row_real_cells)
            for i in range(grid_after):
                cells.append(_placeholder(start_after + i, last_real))

            # -- defensive padding: if a row still under-fills col_count
            # -- (e.g. malformed doc missing cells without gridAfter), pad so
            # -- downstream indexing is safe rather than mis-aligned. --
            row_len = len(cells) - row_start
            while row_len < col_count:
                cells.append(_placeholder(row_len, last_real or first_real))
                row_len += 1
        return cells

    @property
    def _column_count(self):
        """The number of grid columns in this table."""
        return self._tbl.col_count

    @property
    def _tblPr(self) -> CT_TblPr:
        return self._tbl.tblPr


class _Cell(BlockItemContainer):
    """Table cell."""

    def __init__(self, tc: CT_Tc, parent: TableParent):
        super().__init__(tc, cast("t.ProvidesStoryPart", parent))
        self._parent = parent
        self._tc = self._element = tc

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None):
        """Return a paragraph newly added to the end of the content in this cell.

        If present, `text` is added to the paragraph in a single run. If specified, the
        paragraph style `style` is applied. If `style` is not specified or is |None|,
        the result is as though the 'Normal' style was applied. Note that the formatting
        of text in a cell can be influenced by the table style. `text` can contain tab
        (``\\t``) characters, which are converted to the appropriate XML form for a tab.
        `text` can also include newline (``\\n``) or carriage return (``\\r``)
        characters, each of which is converted to a line break.
        """
        return super().add_paragraph(text, style)

    def add_picture(
        self,
        image_path_or_stream: "str | os.PathLike[str] | IO[bytes]",
        width: int | Length | None = None,
        height: int | Length | None = None,
    ):
        """Return |InlineShape| for a picture added to this cell.

        The picture is appended to the cell, in its own paragraph, and scaled based on
        `width` and `height`. If neither is specified, the picture appears at its native
        size. If only one is specified, it is used to compute a scaling factor that
        preserves the image's aspect ratio.

        `image_path_or_stream` may be a ``str`` path, an :class:`os.PathLike` (e.g.
        :class:`pathlib.Path`), or a binary file-like object containing an image.

        .. versionadded:: 1.3.0.dev0
        """
        if isinstance(image_path_or_stream, os.PathLike):
            image_path_or_stream = os.fspath(image_path_or_stream)
        run = self.add_paragraph().add_run()
        return run.add_picture(image_path_or_stream, width, height)

    @property
    def borders(self) -> CellBorders:
        """Read-only. |CellBorders| object providing access to cell border properties.

        Always returns a |CellBorders| object; setting border properties on it will
        create the required XML elements on demand.

        .. versionadded:: 1.3.0.dev0
        """
        return CellBorders(self._tc)

    def add_table(  # pyright: ignore[reportIncompatibleMethodOverride]
        self,
        rows: int,
        cols: int,
        style: "str | _TableStyle | None" = None,
    ) -> Table:
        """Return a table newly added to this cell after any existing cell content.

        The new table will have `rows` rows and `cols` columns. If `style` is
        supplied it is applied to the new table (by style name or
        :class:`_TableStyle` object).

        An empty paragraph is added after the table because Word requires a paragraph
        element as the last element in every cell.

        .. versionadded:: 1.3.0.dev0
           ``style`` parameter. (upstream#1285)
        """
        width = self.width if self.width is not None else Inches(1)
        table = super().add_table(rows, cols, width)
        if style is not None:
            table.style = style
        self.add_paragraph()
        return table

    def split(self) -> list[_Cell]:
        """Split this cell back into individual single-grid cells.

        Clears any ``w:gridSpan`` and ``w:vMerge`` on this cell so the cell is
        no longer merged. When the cell has ``w:gridSpan > 1``, it is replaced
        by N sibling ``w:tc`` elements (one per spanned grid column), each a
        clone of the original without the ``w:gridSpan`` / ``w:vMerge``
        children and with empty paragraph content in all but the first cell
        (the original content is preserved on the first cell). The original
        cell's width is evenly distributed across the new cells when it had
        an explicit width.

        Returns the list of resulting |_Cell| objects (left-to-right). If
        this cell is not merged the list contains only this cell.

        Note: this implementation splits at row level only. A vertically
        merged cell (``vMerge="restart"`` with ``vMerge="continue"`` rows
        below) has the continuation rows' ``w:vMerge`` attribute cleared on
        the matching column only; no new continuation cells are fabricated.
        Callers needing to fully unmerge a vertical span should iterate each
        row and call :meth:`split` on each continuation cell.

        .. versionadded:: 1.3.0.dev0
        """
        from copy import deepcopy

        tc = self._tc
        grid_span = tc.grid_span
        tcPr = tc.tcPr

        # -- clear any vMerge on this cell --
        if tcPr is not None and tcPr.vMerge is not None:
            tcPr._remove_vMerge()  # pyright: ignore[reportPrivateUsage]

        if grid_span <= 1:
            return [self]

        # -- clear gridSpan on this cell --
        if tcPr is not None:
            tcPr._remove_gridSpan()  # pyright: ignore[reportPrivateUsage]

        # -- redistribute width across the N split cells --
        original_width = tc.width
        per_cell_width: Length | None = None
        if original_width is not None:
            per_cell_width = Emu(int(original_width) // grid_span)
            tc.width = per_cell_width

        # -- build sibling cells. Clone tc (content + tcPr) then strip inner
        # -- block-content on the clones so only the first cell keeps the
        # -- merged content. --
        from docx.oxml.ns import qn

        result: list[_Cell] = [self]
        current = tc
        for _ in range(grid_span - 1):
            clone = cast("CT_Tc", deepcopy(tc))
            # -- strip paragraphs / tables / sdt from clone so it becomes empty --
            for child in list(clone):
                if child.tag != qn("w:tcPr"):
                    clone.remove(child)
            # -- ensure clone has the required single empty w:p --
            clone.add_p()
            # -- clone's tcPr already lacks gridSpan/vMerge since we cleared
            # -- them on tc before deepcopy. --
            current.addnext(clone)
            current = clone
            result.append(_Cell(clone, self._parent))
        return result

    @property
    def formatting_change(self) -> FormattingChange | None:
        """|FormattingChange| proxy for this cell's `w:tcPrChange`, or |None|.

        Returns a read-only :class:`~docx.tracked_changes.FormattingChange`
        exposing the prior cell properties (`w:tcPr`) when this cell carries
        a `w:tcPr/w:tcPrChange` tracked-revision marker. Returns |None| when
        the cell has no `w:tcPr` or no `w:tcPrChange` child.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.tracked_changes import FormattingChange

        tcPr = self._tc.tcPr
        if tcPr is None:
            return None
        tcPrChange = tcPr.tcPrChange
        if tcPrChange is None:
            return None
        return FormattingChange(tcPrChange)

    @property
    def grid_span(self) -> int:
        """Number of layout-grid cells this cell spans horizontally.

        A "normal" cell has a grid-span of 1. A horizontally merged cell has a grid-span of 2 or
        more.
        """
        return self._tc.grid_span

    @property
    def stable_id(self) -> str:
        """A 16-character hex stable identifier for this cell.

        The ID is derived from the cell's position within its parent row and
        its text content (paragraphs joined by ``"\\n"``). It is stable across
        save/reload *when the cell keeps the same position with the same
        text*; it changes if the row is reordered, the cell is moved within
        its row, or its text is edited. The value is recomputed on each
        access and never persisted on the element.

        The ``w:tc`` element itself has no ``@w:rsidR``, so only the
        structural hash is used. For more robust cross-session tracking,
        compare the cell's content directly.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.ids import compute_stable_id

        return compute_stable_id(self._tc, self.text)

    @property
    def is_tracked_insertion(self) -> bool:
        """|True| when this cell carries a `w:tcPr/w:cellIns` revision marker.

        A `w:cellIns` element indicates the cell was inserted by a tracked
        change. Returns |False| when the cell has no `w:tcPr` or no `w:cellIns`
        child.

        .. versionadded:: 1.3.0.dev0
        """
        tcPr = self._tc.tcPr
        if tcPr is None:
            return False
        return tcPr.cellIns is not None

    @property
    def is_tracked_deletion(self) -> bool:
        """|True| when this cell carries a `w:tcPr/w:cellDel` revision marker.

        A `w:cellDel` element indicates the cell was deleted by a tracked
        change. Returns |False| when the cell has no `w:tcPr` or no `w:cellDel`
        child.

        .. versionadded:: 1.3.0.dev0
        """
        tcPr = self._tc.tcPr
        if tcPr is None:
            return False
        return tcPr.cellDel is not None

    @property
    def is_merge_origin(self) -> bool | None:
        """Tri-state indicator of this cell's role in a merged region.

        - |None| when this cell is not part of any merged region (``w:gridSpan`` is
          1 or absent *and* ``w:vMerge`` is absent).
        - |True| when this cell is the top-left of a merged region. That is,
          either ``w:vMerge/@w:val="restart"`` or ``w:gridSpan > 1`` without a
          ``w:vMerge`` child indicating it is a later row of a vertical span.
        - |False| when this cell is a continuation of a merged region (has
          ``w:vMerge`` without ``@w:val="restart"``, i.e. the default
          ``"continue"`` value).

        Read-only.

        .. versionadded:: 1.3.0.dev0
        """
        tc = self._tc
        vMerge = tc.vMerge
        # -- treat grid_span <= 1 as "no horizontal span" --
        has_h_span = tc.grid_span > 1
        if vMerge is None and not has_h_span:
            return None
        if vMerge == ST_Merge.CONTINUE:
            return False
        # -- vMerge == "restart" or (vMerge is None and has_h_span) --
        return True

    @property
    def merge_origin(self) -> _Cell:
        """The top-left |_Cell| of the merged region this cell belongs to.

        Returns this cell itself if it is not part of any merged region or if it
        is already the origin (e.g. horizontal-only merge, or vMerge="restart").
        Walks up the vertical span, following ``w:vMerge`` continuations, until
        it reaches the ``w:vMerge/@w:val="restart"`` cell.

        Raises |ValueError| if this cell is an orphan continuation — i.e. it has
        ``w:vMerge`` but no ancestor row contains a corresponding
        ``w:vMerge="restart"`` cell.

        .. versionadded:: 1.3.0.dev0
        """
        tc = self._tc
        if tc.vMerge != ST_Merge.CONTINUE:
            return self
        # -- walk up following vMerge="continue" until we find "restart" --
        current = tc
        visited: set[int] = set()
        while current.vMerge == ST_Merge.CONTINUE:
            if id(current) in visited:
                raise ValueError("cycle detected while locating merge origin")
            visited.add(id(current))
            try:
                above = current._tc_above  # pyright: ignore[reportPrivateUsage]
            except (ValueError, IndexError):
                raise ValueError(
                    "orphan vMerge continuation cell has no restart ancestor"
                )
            current = above
        return _Cell(current, self._parent)

    def merge(self, other_cell: _Cell):
        """Return a merged cell created by spanning the rectangular region having this
        cell and `other_cell` as diagonal corners.

        Raises |InvalidSpanError| if the cells do not define a rectangular region.
        """
        tc, tc_2 = self._tc, other_cell._tc
        merged_tc = tc.merge(tc_2)
        return _Cell(merged_tc, self._parent)

    @property
    def paragraphs(self):
        """List of paragraphs in the cell.

        A table cell is required to contain at least one block-level element and end
        with a paragraph. By default, a new cell contains a single paragraph. Read-only
        """
        return super().paragraphs

    @property
    def tables(self):
        """List of tables in the cell, in the order they appear.

        Read-only.
        """
        return super().tables

    @property
    def text(self) -> str:
        """The entire contents of this cell as a string of text.

        Assigning a string to this property replaces all existing content with a single
        paragraph containing the assigned text in a single run.
        """
        return "\n".join(p.text for p in self.paragraphs)

    @text.setter
    def text(self, text: str):
        """Write-only.

        Set entire contents of cell to the string `text`. Any existing content or
        revisions are replaced.
        """
        tc = self._tc
        tc.clear_content()
        p = tc.add_p()
        r = p.add_r()
        r.text = text

    @property
    def shading(self) -> CellShading:
        """Read-only. |CellShading| object providing access to shading properties.

        Always returns a |CellShading| object; setting shading properties on it will
        create the required XML elements on demand.

        .. versionadded:: 1.3.0.dev0
        """
        return CellShading(self._tc)

    @property
    def margins(self) -> CellMargins:
        """Read-only. |CellMargins| proxy for per-cell margin overrides.

        Always returns a |CellMargins| object. When no ``w:tcMar`` element is
        present, each edge reads as |None|; assigning to an edge creates the
        ``w:tcPr/w:tcMar`` structure on demand.

        .. versionadded:: 1.3.0.dev0
        """
        return CellMargins(self._tc)

    def set_margins(
        self,
        top: "Length | None" = None,
        bottom: "Length | None" = None,
        start: "Length | None" = None,
        end: "Length | None" = None,
    ) -> CellMargins:
        """Set one or more cell-margin edges in a single call.

        Only arguments explicitly provided (i.e. not |None|) are written; existing
        edges not mentioned in the call are left unchanged. To explicitly clear
        an edge, assign |None| directly via the |CellMargins| proxy or call
        :meth:`remove_margins`. Returns the |CellMargins| proxy.

        .. versionadded:: 1.3.0.dev0
        """
        margins = self.margins
        if top is not None:
            margins.top = top
        if bottom is not None:
            margins.bottom = bottom
        if start is not None:
            margins.start = start
        if end is not None:
            margins.end = end
        return margins

    def remove_margins(self) -> None:
        """Remove any ``w:tcMar`` element from this cell, clearing all per-cell
        margin overrides. Leaves the cell inheriting table-level cell margins.

        .. versionadded:: 1.3.0.dev0
        """
        tcPr = self._tc.tcPr
        if tcPr is None:
            return
        tcPr._remove_tcMar()  # pyright: ignore[reportPrivateUsage]

    @property
    def text_direction(self) -> WD_TEXT_DIRECTION | None:
        """Member of :ref:`WdTextDirection` or |None|.

        Controls the flow direction of text within the cell. A value of |None|
        indicates the text direction for this cell is inherited. Assigning |None|
        causes any explicitly defined text direction to be removed, restoring
        inheritance.

        The common cell-rotation cases are ``WD_TEXT_DIRECTION.TB_RL`` (rotate
        90 degrees clockwise) and ``WD_TEXT_DIRECTION.BT_LR`` (rotate 90 degrees
        counter-clockwise).

        .. versionadded:: 1.3.0.dev0
        """
        tcPr = self._element.tcPr
        if tcPr is None:
            return None
        return tcPr.text_direction

    @text_direction.setter
    def text_direction(self, value: WD_TEXT_DIRECTION | None):
        tcPr = self._element.get_or_add_tcPr()
        tcPr.text_direction = value

    @property
    def vertical_alignment(self):
        """Member of :ref:`WdCellVerticalAlignment` or None.

        A value of |None| indicates vertical alignment for this cell is inherited.
        Assigning |None| causes any explicitly defined vertical alignment to be removed,
        restoring inheritance.
        """
        tcPr = self._element.tcPr
        if tcPr is None:
            return None
        return tcPr.vAlign_val

    @vertical_alignment.setter
    def vertical_alignment(self, value: WD_CELL_VERTICAL_ALIGNMENT | None):
        tcPr = self._element.get_or_add_tcPr()
        tcPr.vAlign_val = value

    @property
    def width(self):
        """The width of this cell in EMU, or |None| if no explicit width is set."""
        return self._tc.width

    @width.setter
    def width(self, value: Length):
        self._tc.width = value


class CellShading:
    """Provides access to shading properties for a table cell.

    Accessed via ``_Cell.shading``.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, tc: CT_Tc):
        self._tc = tc

    @property
    def fill_color(self) -> RGBColor | None:
        """The background fill color as an |RGBColor| value, or |None| if not set.

        Note: returns |None| when the fill attribute is ``"auto"`` (foreground-dependent).

        .. versionadded:: 1.3.0.dev0
        """
        shd = self._shd
        if shd is None:
            return None
        fill = shd.fill
        if fill is None or not isinstance(fill, RGBColor):
            return None
        return fill

    @fill_color.setter
    def fill_color(self, value: RGBColor | None):
        if value is None:
            tcPr = self._tc.tcPr
            if tcPr is not None and tcPr.shd is not None:
                tcPr.shd.fill = None
            return
        shd = self._get_or_add_shd()
        shd.fill = value

    @property
    def pattern(self) -> WD_SHADING_PATTERN | None:
        """The shading pattern as a |WD_SHADING_PATTERN| value, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        shd = self._shd
        if shd is None:
            return None
        return shd.val

    @pattern.setter
    def pattern(self, value: WD_SHADING_PATTERN | None):
        if value is None:
            tcPr = self._tc.tcPr
            if tcPr is not None and tcPr.shd is not None:
                tcPr.shd.val = None
            return
        shd = self._get_or_add_shd()
        shd.val = value

    @property
    def _shd(self) -> CT_Shd | None:
        tcPr = self._tc.tcPr
        if tcPr is None:
            return None
        return tcPr.shd

    def _get_or_add_shd(self) -> CT_Shd:
        tcPr = self._tc.get_or_add_tcPr()
        shd = tcPr.get_or_add_shd()
        if shd.val is None:
            shd.val = WD_SHADING_PATTERN.CLEAR
        return shd


class BorderElement:
    """Provides access to properties of a single border edge.

    Wraps a ``CT_Border`` element (e.g. ``<w:top>``, ``<w:bottom>``).

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, border: CT_Border | None, get_or_add: Callable[[], CT_Border]):
        self._border = border
        self._get_or_add = get_or_add

    @property
    def style(self) -> WD_BORDER_STYLE | None:
        """The border style as a |WD_BORDER_STYLE| value, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border
        if border is None:
            return None
        return border.val

    @style.setter
    def style(self, value: WD_BORDER_STYLE | None):
        if value is None:
            border = self._border
            if border is not None:
                border.val = None
            return
        border = self._get_or_add()
        self._border = border
        border.val = value

    @property
    def width(self) -> Length | None:
        """The border width as an EMU |Length| value, or |None| if not set.

        The ``w:sz`` attribute stores the width in eighths of a point; the
        underlying element class already converts that to a |Length| (EMU) on
        read, so it is returned as-is here.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border
        if border is None:
            return None
        return border.sz

    @width.setter
    def width(self, value: Length | None):
        if value is None:
            border = self._border
            if border is not None:
                border.sz = None
            return
        border = self._get_or_add()
        self._border = border
        border.sz = value

    @property
    def color(self) -> RGBColor | None:
        """The border color as an |RGBColor| value, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border
        if border is None:
            return None
        color = border.color
        if color is None or not isinstance(color, RGBColor):
            return None
        return color

    @color.setter
    def color(self, value: RGBColor | None):
        if value is None:
            border = self._border
            if border is not None:
                border.color = None
            return
        border = self._get_or_add()
        self._border = border
        border.color = value

    @property
    def space(self) -> Length | None:
        """The border spacing as a |Length| (EMU) value, or |None| if not set.

        The ``w:space`` attribute stores whole points; the underlying element
        class converts that to a |Length| on read.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border
        if border is None:
            return None
        return border.space

    @space.setter
    def space(self, value: Length | int | None):
        if value is None:
            border = self._border
            if border is not None:
                border.space = None
            return
        border = self._get_or_add()
        self._border = border
        border.space = value


class TableBorders:
    """Provides access to border properties for a table.

    Accessed via ``Table.borders``.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, tbl: CT_Tbl):
        self._tbl = tbl

    @property
    def top(self) -> BorderElement:
        """The top border of the table.

        .. versionadded:: 1.3.0.dev0
        """
        tblBorders = self._tblBorders
        return BorderElement(
            tblBorders.top if tblBorders is not None else None,
            lambda: self._get_or_add_tblBorders().get_or_add_top(),
        )

    @property
    def bottom(self) -> BorderElement:
        """The bottom border of the table.

        .. versionadded:: 1.3.0.dev0
        """
        tblBorders = self._tblBorders
        return BorderElement(
            tblBorders.bottom if tblBorders is not None else None,
            lambda: self._get_or_add_tblBorders().get_or_add_bottom(),
        )

    @property
    def left(self) -> BorderElement:
        """The left border of the table.

        .. versionadded:: 1.3.0.dev0
        """
        tblBorders = self._tblBorders
        return BorderElement(
            tblBorders.left if tblBorders is not None else None,
            lambda: self._get_or_add_tblBorders().get_or_add_left(),
        )

    @property
    def right(self) -> BorderElement:
        """The right border of the table.

        .. versionadded:: 1.3.0.dev0
        """
        tblBorders = self._tblBorders
        return BorderElement(
            tblBorders.right if tblBorders is not None else None,
            lambda: self._get_or_add_tblBorders().get_or_add_right(),
        )

    @property
    def inside_h(self) -> BorderElement:
        """The inside horizontal border of the table.

        .. versionadded:: 1.3.0.dev0
        """
        tblBorders = self._tblBorders
        return BorderElement(
            tblBorders.insideH if tblBorders is not None else None,
            lambda: self._get_or_add_tblBorders().get_or_add_insideH(),
        )

    @property
    def inside_v(self) -> BorderElement:
        """The inside vertical border of the table.

        .. versionadded:: 1.3.0.dev0
        """
        tblBorders = self._tblBorders
        return BorderElement(
            tblBorders.insideV if tblBorders is not None else None,
            lambda: self._get_or_add_tblBorders().get_or_add_insideV(),
        )

    @property
    def _tblBorders(self) -> CT_TblBorders | None:
        return self._tbl.tblPr.tblBorders

    def _get_or_add_tblBorders(self) -> CT_TblBorders:
        return self._tbl.tblPr.get_or_add_tblBorders()


class TableStyleFlags:
    """Provides access to the `w:tblLook` conditional-formatting flags for a table.

    Each flag corresponds to one of the individual ST_OnOff attributes on
    `w:tblLook` and controls which table-style "conditional" features Word will
    render (e.g. banded rows, first-row/column emphasis).

    Accessed via ``Table.style_flags``. When the underlying `w:tblLook` element
    is absent, reading a flag returns |False|; writing any flag creates the
    `w:tblLook` element on demand.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, tbl: CT_Tbl):
        self._tbl = tbl

    @property
    def first_row(self) -> bool:
        """|True| when the table-style formatting for the first row is applied.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_flag("firstRow")

    @first_row.setter
    def first_row(self, value: bool) -> None:
        self._set_flag("firstRow", value)

    @property
    def last_row(self) -> bool:
        """|True| when the table-style formatting for the last row is applied.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_flag("lastRow")

    @last_row.setter
    def last_row(self, value: bool) -> None:
        self._set_flag("lastRow", value)

    @property
    def first_column(self) -> bool:
        """|True| when table-style formatting for the first column is applied.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_flag("firstColumn")

    @first_column.setter
    def first_column(self, value: bool) -> None:
        self._set_flag("firstColumn", value)

    @property
    def last_column(self) -> bool:
        """|True| when table-style formatting for the last column is applied.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_flag("lastColumn")

    @last_column.setter
    def last_column(self, value: bool) -> None:
        self._set_flag("lastColumn", value)

    @property
    def no_horizontal_banding(self) -> bool:
        """|True| when row-banding formatting is suppressed (``@w:noHBand="1"``).

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_flag("noHBand")

    @no_horizontal_banding.setter
    def no_horizontal_banding(self, value: bool) -> None:
        self._set_flag("noHBand", value)

    @property
    def no_vertical_banding(self) -> bool:
        """|True| when column-banding formatting is suppressed (``@w:noVBand="1"``).

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_flag("noVBand")

    @no_vertical_banding.setter
    def no_vertical_banding(self, value: bool) -> None:
        self._set_flag("noVBand", value)

    def _get_flag(self, name: str) -> bool:
        tblLook = self._tbl.tblPr.tblLook
        if tblLook is None:
            return False
        return tblLook.get_flag(name)

    def _set_flag(self, name: str, value: bool) -> None:
        tblLook = self._tbl.tblPr.get_or_add_tblLook()
        tblLook.set_flag(name, value)


class CellBorders:
    """Provides access to border properties for a table cell.

    Accessed via ``_Cell.borders``.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, tc: CT_Tc):
        self._tc = tc

    @property
    def top(self) -> BorderElement:
        """The top border of the cell.

        .. versionadded:: 1.3.0.dev0
        """
        tcBorders = self._tcBorders
        return BorderElement(
            tcBorders.top if tcBorders is not None else None,
            lambda: self._get_or_add_tcBorders().get_or_add_top(),
        )

    @property
    def bottom(self) -> BorderElement:
        """The bottom border of the cell.

        .. versionadded:: 1.3.0.dev0
        """
        tcBorders = self._tcBorders
        return BorderElement(
            tcBorders.bottom if tcBorders is not None else None,
            lambda: self._get_or_add_tcBorders().get_or_add_bottom(),
        )

    @property
    def left(self) -> BorderElement:
        """The left border of the cell.

        .. versionadded:: 1.3.0.dev0
        """
        tcBorders = self._tcBorders
        return BorderElement(
            tcBorders.left if tcBorders is not None else None,
            lambda: self._get_or_add_tcBorders().get_or_add_left(),
        )

    @property
    def right(self) -> BorderElement:
        """The right border of the cell.

        .. versionadded:: 1.3.0.dev0
        """
        tcBorders = self._tcBorders
        return BorderElement(
            tcBorders.right if tcBorders is not None else None,
            lambda: self._get_or_add_tcBorders().get_or_add_right(),
        )

    @property
    def _tcBorders(self) -> CT_TcBorders | None:
        tcPr = self._tc.tcPr
        if tcPr is None:
            return None
        return tcPr.tcBorders

    def _get_or_add_tcBorders(self) -> CT_TcBorders:
        return self._tc.get_or_add_tcPr().get_or_add_tcBorders()


class CellMargins:
    """Proxy for per-cell margin overrides (the ``w:tcMar`` element).

    Accessed via :attr:`_Cell.margins`. Provides read/write access to the four
    margin edges: ``top``, ``bottom``, ``start`` and ``end``. The underlying
    ``w:tcMar`` element (and its parent ``w:tcPr``) are created lazily on first
    write. When no ``w:tcMar`` is present, each edge reads as |None|.

    The edge names ``start`` and ``end`` map to either the modern ``w:start`` /
    ``w:end`` tags or the legacy ``w:left`` / ``w:right`` tags. Reads accept
    either form; writes produce ``w:start`` / ``w:end``.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, tc: CT_Tc):
        self._tc = tc

    @property
    def _tcMar(self) -> "CT_TcMar | None":
        tcPr = self._tc.tcPr
        if tcPr is None:
            return None
        return tcPr.tcMar

    def _get_or_add_tcMar(self) -> "CT_TcMar":
        return self._tc.get_or_add_tcPr().get_or_add_tcMar()

    def _get_edge(self, edge: str) -> "Length | None":
        tcMar = self._tcMar
        if tcMar is None:
            return None
        return tcMar.get_margin(edge)

    def _set_edge(self, edge: str, value: "Length | None") -> None:
        if value is None:
            tcMar = self._tcMar
            if tcMar is None:
                return
            tcMar.remove_margin(edge)
            # -- if the tcMar is now empty, remove it to keep the XML tidy --
            if len(tcMar) == 0:
                tcPr = self._tc.tcPr
                if tcPr is not None:
                    tcPr._remove_tcMar()  # pyright: ignore[reportPrivateUsage]
            return
        self._get_or_add_tcMar().set_margin(edge, value)

    @property
    def top(self) -> "Length | None":
        """Top cell-margin as a |Length|, or |None| when not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("top")

    @top.setter
    def top(self, value: "Length | None") -> None:
        self._set_edge("top", value)

    @property
    def bottom(self) -> "Length | None":
        """Bottom cell-margin as a |Length|, or |None| when not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("bottom")

    @bottom.setter
    def bottom(self, value: "Length | None") -> None:
        self._set_edge("bottom", value)

    @property
    def start(self) -> "Length | None":
        """Start (leading-edge) cell-margin as a |Length|, or |None| when not set.

        Reads ``w:start`` when present, otherwise the legacy ``w:left``. Writes
        always produce ``w:start``.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("start")

    @start.setter
    def start(self, value: "Length | None") -> None:
        self._set_edge("start", value)

    @property
    def end(self) -> "Length | None":
        """End (trailing-edge) cell-margin as a |Length|, or |None| when not set.

        Reads ``w:end`` when present, otherwise the legacy ``w:right``. Writes
        always produce ``w:end``.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("end")

    @end.setter
    def end(self, value: "Length | None") -> None:
        self._set_edge("end", value)


class TableCellMargins:
    """Proxy for table-level default cell margins (the ``w:tblCellMar`` element).

    Accessed via :attr:`Table.cell_margins`. Provides read/write access to the
    four edges: ``top``, ``bottom``, ``start`` and ``end``. The underlying
    ``w:tblCellMar`` element is created lazily on first write. When no
    ``w:tblCellMar`` is present, each edge reads as |None|.

    Mirrors :class:`CellMargins` but operates at the table level — writes go
    to ``w:tblPr/w:tblCellMar`` rather than ``w:tcPr/w:tcMar``. The edge
    names ``start`` and ``end`` read either the modern ``w:start``/``w:end``
    tags or the legacy ``w:left``/``w:right`` pair; writes always produce
    the modern form.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, tbl: CT_Tbl):
        self._tbl = tbl

    @property
    def _tblCellMar(self) -> "CT_TblCellMar | None":
        return self._tbl.tblPr.tblCellMar

    def _get_or_add_tblCellMar(self) -> "CT_TblCellMar":
        return self._tbl.tblPr.get_or_add_tblCellMar()

    def _get_edge(self, edge: str) -> "Length | None":
        tblCellMar = self._tblCellMar
        if tblCellMar is None:
            return None
        return tblCellMar.get_margin(edge)

    def _set_edge(self, edge: str, value: "Length | None") -> None:
        if value is None:
            tblCellMar = self._tblCellMar
            if tblCellMar is None:
                return
            tblCellMar.remove_margin(edge)
            # -- if the tblCellMar is now empty, remove it to keep the XML tidy --
            if len(tblCellMar) == 0:
                self._tbl.tblPr._remove_tblCellMar()  # pyright: ignore[reportPrivateUsage]
            return
        self._get_or_add_tblCellMar().set_margin(edge, value)

    @property
    def top(self) -> "Length | None":
        """Top default cell-margin as a |Length|, or |None| when not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("top")

    @top.setter
    def top(self, value: "Length | None") -> None:
        self._set_edge("top", value)

    @property
    def bottom(self) -> "Length | None":
        """Bottom default cell-margin as a |Length|, or |None| when not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("bottom")

    @bottom.setter
    def bottom(self, value: "Length | None") -> None:
        self._set_edge("bottom", value)

    @property
    def start(self) -> "Length | None":
        """Start-edge default cell-margin as a |Length|, or |None| when not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("start")

    @start.setter
    def start(self, value: "Length | None") -> None:
        self._set_edge("start", value)

    @property
    def end(self) -> "Length | None":
        """End-edge default cell-margin as a |Length|, or |None| when not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._get_edge("end")

    @end.setter
    def end(self, value: "Length | None") -> None:
        self._set_edge("end", value)


class _Column(Parented):
    """Table column."""

    def __init__(self, gridCol: CT_TblGridCol, parent: TableParent):
        super().__init__(parent)
        self._parent = parent
        self._gridCol = gridCol

    @property
    def cells(self) -> tuple[_Cell, ...]:
        """Sequence of |_Cell| instances corresponding to cells in this column."""
        return tuple(self.table.column_cells(self._index))

    @property
    def table(self) -> Table:
        """Reference to the |Table| object this column belongs to."""
        return self._parent.table

    @property
    def width(self) -> Length | None:
        """The width of this column in EMU, or |None| if no explicit width is set."""
        return self._gridCol.w

    @width.setter
    def width(self, value: Length | None):
        """Set the column width and propagate to each row's corresponding cell.

        Writes ``@w:w`` on this ``w:gridCol`` and updates ``w:tcW`` on every
        ``w:tc`` at the matching grid-offset. Cells that horizontally span more
        than one grid column (``w:gridSpan > 1``) are left untouched to avoid
        clobbering a merged cell's existing width. Assigning |None| removes
        ``@w:w`` on the gridCol and removes ``w:tcW`` from each matching cell.
        """
        self._gridCol.w = value
        tblGrid = self._gridCol.getparent()
        if tblGrid is None:
            return
        tbl = cast("CT_Tbl | None", tblGrid.getparent())
        if tbl is None:
            return
        col_idx = self._index
        # -- propagate to each row's single-span cell at this grid offset --
        for tr in tbl.tr_lst:
            for tc in tr.tc_lst:
                if tc.grid_offset != col_idx:
                    continue
                if tc.grid_span != 1:
                    # -- leave merged / spanned cells alone --
                    break
                if value is None:
                    tcPr = tc.tcPr
                    if tcPr is not None:
                        tcW = tcPr.tcW
                        if tcW is not None:
                            tcPr.remove(tcW)
                else:
                    tc.width = value
                break

    def delete(self) -> None:
        """Remove this column from its table.

        Removes the matching ``w:gridCol`` from ``w:tblGrid`` and the
        corresponding ``w:tc`` from each row. Cells horizontally spanning
        this column (``w:gridSpan > 1``) have their gridSpan decremented by
        one rather than being orphaned. After calling this method the
        |_Column| instance is "defunct" and should not be used further.

        .. versionadded:: 1.3.0.dev0
        """
        self.table.delete_column(self._index)

    @property
    def index(self) -> int:
        """Zero-based index of this column within its table.

        The index reflects the position of the underlying ``w:gridCol`` element
        among its siblings in the table's ``w:tblGrid``. Equivalent to
        ``table.columns[n].index == n`` for non-merged tables.

        .. versionadded:: 1.3.0.dev0
        """
        return self._gridCol.gridCol_idx

    # -- legacy private alias retained for backwards compatibility --
    _index = index


class _Columns(Parented):
    """Sequence of |_Column| instances corresponding to the columns in a table.

    Supports ``len()``, iteration, indexed access, and slicing.
    """

    def __init__(self, tbl: CT_Tbl, parent: TableParent):
        super().__init__(parent)
        self._parent = parent
        self._tbl = tbl

    @overload
    def __getitem__(self, idx: int) -> _Column: ...

    @overload
    def __getitem__(self, idx: slice) -> list[_Column]: ...

    def __getitem__(self, idx: int | slice) -> _Column | list[_Column]:
        """Provide indexed and sliced access, e.g. ``columns[0]`` or ``columns[1:]``.

        A slice returns a ``list[_Column]`` rather than a ``_Columns`` view;
        integer access returns a single |_Column|. Raises |IndexError| when an
        integer `idx` is out of range.
        """
        if isinstance(idx, slice):
            return [_Column(gridCol, self) for gridCol in self._gridCol_lst[idx]]
        try:
            gridCol = self._gridCol_lst[idx]
        except IndexError:
            msg = "column index [%d] is out of range" % idx
            raise IndexError(msg)
        return _Column(gridCol, self)

    def __iter__(self):
        for gridCol in self._gridCol_lst:
            yield _Column(gridCol, self)

    def __len__(self):
        return len(self._gridCol_lst)

    @property
    def table(self) -> Table:
        """Reference to the |Table| object this column collection belongs to."""
        return self._parent.table

    @property
    def _gridCol_lst(self):
        """Sequence containing ``<w:gridCol>`` elements for this table, each
        representing a table column."""
        tblGrid = self._tbl.tblGrid
        return tblGrid.gridCol_lst


class _Row(Parented):
    """Table row."""

    def __init__(self, tr: CT_Row, parent: TableParent):
        super().__init__(parent)
        self._parent = parent
        self._tr = self._element = tr

    @property
    def allow_break_across_pages(self) -> bool:
        """True when row can be split across page boundaries.

        When set to |False|, the entire row is moved to the next page rather than
        allowing it to be split across a page break. Defaults to |True|.

        .. versionadded:: 1.3.0.dev0
        """
        return self._tr.allow_break_across_pages

    @allow_break_across_pages.setter
    def allow_break_across_pages(self, value: bool):
        self._tr.allow_break_across_pages = value

    @property
    def cells(self) -> tuple[_Cell, ...]:
        """Sequence of |_Cell| instances corresponding to cells in this row.

        Note that Word allows table rows to start later than the first column and end before the
        last column.

        - Only cells actually present are included in the return value.
        - This implies the length of this cell sequence may differ between rows of the same table.
        - If you are reading the cells from each row to form a rectangular "matrix" data structure
          of the table cell values, you will need to account for empty leading and/or trailing
          layout-grid positions using `.grid_cols_before` and `.grid_cols_after`.

        """

        def iter_tc_cells(tc: CT_Tc) -> Iterator[_Cell]:
            """Generate a cell object for each layout-grid cell in `tc`.

            In particular, a `<w:tc>` element with a horizontal "span" with generate the same cell
            multiple times, one for each grid-cell being spanned. This approximates a row in a
            "uniform" table, where each row has a cell for each column in the table.
            """
            # -- a cell comprising the second or later row of a vertical span is indicated by
            # -- tc.vMerge="continue" (the default value of the `w:vMerge` attribute, when it is
            # -- present in the XML). The `w:tc` element at the same grid-offset in the prior row
            # -- is guaranteed to be the same width (gridSpan). So we can delegate content
            # -- discovery to that prior-row `w:tc` element (recursively) until we arrive at the
            # -- "root" cell -- for the vertical span.
            if tc.vMerge == "continue":
                try:
                    above = tc._tc_above  # pyright: ignore[reportPrivateUsage]
                except (ValueError, IndexError):
                    # -- orphan continuation cell (no preceding row or no tc at
                    # -- same grid offset). Treat this tc as its own cell so
                    # -- iteration doesn't crash on malformed documents.
                    above = None
                if above is not None:
                    yield from iter_tc_cells(above)
                    return

            # -- Otherwise, vMerge is either "restart" or None, meaning this `tc` holds the actual
            # -- content of the cell (whether it is vertically merged or not).
            cell = _Cell(tc, self.table)
            for _ in range(tc.grid_span):
                yield cell

        def _iter_row_cells() -> Iterator[_Cell]:
            """Generate `_Cell` instance for each populated layout-grid cell in this row."""
            for tc in self._tr.tc_lst:
                yield from iter_tc_cells(tc)

        return tuple(_iter_row_cells())

    @property
    def formatting_change(self) -> FormattingChange | None:
        """|FormattingChange| proxy for this row's `w:trPrChange`, or |None|.

        Returns a read-only :class:`~docx.tracked_changes.FormattingChange`
        exposing the prior row properties (`w:trPr`) when this row carries a
        `w:trPr/w:trPrChange` tracked-revision marker. Returns |None| when
        the row has no `w:trPr` or no `w:trPrChange` child.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.tracked_changes import FormattingChange

        trPr = self._tr.trPr
        if trPr is None:
            return None
        trPrChange = trPr.trPrChange
        if trPrChange is None:
            return None
        return FormattingChange(trPrChange)

    @property
    def grid_cols_after(self) -> int:
        """Count of unpopulated grid-columns after the last cell in this row.

        Word allows a row to "end early", meaning that one or more cells are not present at the
        end of that row.

        Note these are not simply "empty" cells. The renderer reads this value and "skips" this
        many columns after drawing the last cell.

        Note this also implies that not all rows are guaranteed to have the same number of cells,
        e.g. `_Row.cells` could have length `n` for one row and `n - m` for the next row in the same
        table. Visually this appears as a column (at the beginning or end, not in the middle) with
        one or more cells missing.
        """
        return self._tr.grid_after

    @property
    def grid_cols_before(self) -> int:
        """Count of unpopulated grid-columns before the first cell in this row.

        Word allows a row to "start late", meaning that one or more cells are not present at the
        beginning of that row.

        Note these are not simply "empty" cells. The renderer reads this value and skips forward to
        the table layout-grid position of the first cell in this row; the renderer "skips" this many
        columns before drawing the first cell.

        Note this also implies that not all rows are guaranteed to have the same number of cells,
        e.g. `_Row.cells` could have length `n` for one row and `n - m` for the next row in the same
        table.
        """
        return self._tr.grid_before

    @property
    def height(self) -> Length | None:
        """Return a |Length| object representing the height of this cell, or |None| if
        no explicit height is set."""
        return self._tr.trHeight_val

    @property
    def is_header(self) -> bool:
        """True when this row is a header row that repeats at the top of each page.

        Read/write. Only the first N consecutive rows can be header rows (Word limitation).

        .. versionadded:: 1.3.0.dev0
        """
        return self._tr.is_header

    @is_header.setter
    def is_header(self, value: bool) -> None:
        self._tr.is_header = value

    @height.setter
    def height(self, value: Length | None):
        self._tr.trHeight_val = value

    @property
    def height_rule(self) -> WD_ROW_HEIGHT_RULE | None:
        """Return the height rule of this cell as a member of the :ref:`WdRowHeightRule`.

        This value is |None| if no explicit height_rule is set.
        """
        return self._tr.trHeight_hRule

    @height_rule.setter
    def height_rule(self, value: WD_ROW_HEIGHT_RULE | None):
        self._tr.trHeight_hRule = value

    @property
    def table(self) -> Table:
        """Reference to the |Table| object this row belongs to."""
        return self._parent.table

    def insert_row_before(self) -> _Row:
        """Insert a new blank row as a sibling directly before this row.

        The new row has one ``w:tc`` per column in the table, sized to each
        ``w:gridCol``'s width. Returns the new |_Row|.

        .. versionadded:: 1.3.0.dev0
        """
        return self.table.insert_row(self._index)

    def insert_row_after(self) -> _Row:
        """Insert a new blank row as a sibling directly after this row.

        The new row has one ``w:tc`` per column in the table, sized to each
        ``w:gridCol``'s width. Returns the new |_Row|.

        .. versionadded:: 1.3.0.dev0
        """
        return self.table.insert_row(self._index + 1)

    def clone(self) -> _Row:
        """Insert a deep-copy of this row directly after it and return the new row.

        The cloned ``w:tr`` preserves ``w:trPr`` / ``w:tc`` / run-level
        formatting. Any ``w:id`` attributes inside the cloned tree are
        stripped so bookmark / comment IDs do not collide with the rest of
        the document. (upstream#1189, #205)

        .. versionadded:: 1.3.0.dev0
        """
        new_tr = _clone_tr(self._tr)
        self._tr.addnext(new_tr)
        return _Row(new_tr, self._parent)

    def add_cell(self) -> _Cell:
        """Append a new blank ``w:tc`` to this row and return it as a |_Cell|.

        The new cell carries no explicit width; row-level ``w:gridAfter``
        omitted-cell allowances (if any) still apply visually.
        (upstream#532)

        .. versionadded:: 1.3.0.dev0
        """
        tc = self._tr.add_tc()
        return _Cell(tc, self.table)

    def insert_cell(self, index: int) -> _Cell:
        """Insert a new blank ``w:tc`` at `index` in this row and return it.

        `index` counts from zero against the row's existing ``w:tc``
        sequence. An `index` equal to the number of cells appends; a
        negative `index` counts from the end. (upstream#532)

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.table import CT_Tc

        tc_lst = self._tr.tc_lst
        count = len(tc_lst)
        if index < 0:
            index += count
        if index < 0 or index > count:
            raise IndexError(f"cell index [{index}] is out of range")
        new_tc = CT_Tc.new()
        if index == count:
            self._tr.append(new_tc)
        else:
            tc_lst[index].addprevious(new_tc)
        return _Cell(new_tc, self.table)

    def apply_shading(self, color: RGBColor | str | None) -> None:
        """Apply a solid fill `color` to every cell in this row.

        `color` may be an |RGBColor|, a 6-character hex string (e.g.
        ``"FF0000"``), or |None| to remove the fill from every cell in this
        row. The shading pattern is set to ``WD_SHADING_PATTERN.CLEAR`` (the
        default solid fill). Cells that are continuation cells of a vertical
        merge still receive shading on the underlying ``w:tc`` element; the
        rendered appearance is controlled by Word's merge resolution.

        .. versionadded:: 1.3.0.dev0
        """
        if color is None:
            rgb: RGBColor | None = None
        elif isinstance(color, RGBColor):
            rgb = color
        else:
            rgb = RGBColor.from_string(color)
        for tc in self._tr.tc_lst:
            cell = _Cell(tc, self.table)
            cell.shading.fill_color = rgb

    @property
    def index(self) -> int:
        """Zero-based index of this row within its table.

        Reflects the position of the underlying ``w:tr`` element among its
        siblings in the table's ``w:tbl`` in document order.

        .. versionadded:: 1.3.0.dev0
        """
        return self._tr.tr_idx

    # -- legacy private alias retained for backwards compatibility --
    _index = index


def _clone_tr(tr: CT_Row) -> CT_Row:
    """Return a deep-copy of `tr` with all nested ``w:id`` attributes stripped.

    Used by :meth:`Table.add_row` and :meth:`_Row.clone` to duplicate a row's
    full XML content while preventing ID collisions on bookmarks, comments,
    footnotes, etc. Any element in the cloned subtree that carries a ``w:id``
    attribute has it removed.
    """
    from copy import deepcopy

    from docx.oxml.ns import qn

    clone = cast("CT_Row", deepcopy(tr))
    id_qn = qn("w:id")
    # -- strip ``w:id`` on the clone and all descendants --
    if clone.get(id_qn) is not None:
        del clone.attrib[id_qn]
    for desc in clone.iter():
        if desc.get(id_qn) is not None:
            del desc.attrib[id_qn]
    return clone


class _Rows(Parented):
    """Sequence of |_Row| objects corresponding to the rows in a table.

    Supports ``len()``, iteration, indexed access, and slicing.
    """

    def __init__(self, tbl: CT_Tbl, parent: TableParent):
        super().__init__(parent)
        self._parent = parent
        self._tbl = tbl

    @overload
    def __getitem__(self, idx: int) -> _Row: ...

    @overload
    def __getitem__(self, idx: slice) -> list[_Row]: ...

    def __getitem__(self, idx: int | slice) -> _Row | list[_Row]:
        """Provide indexed access, (e.g. `rows[0]` or `rows[1:3]`)"""
        return list(self)[idx]

    def __iter__(self):
        return (_Row(tr, self) for tr in self._tbl.tr_lst)

    def __len__(self):
        return len(self._tbl.tr_lst)

    @property
    def table(self) -> Table:
        """Reference to the |Table| object this row collection belongs to."""
        return self._parent.table
