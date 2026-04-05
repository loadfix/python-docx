"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.text import WD_BORDER_STYLE, WD_LINE_SPACING
from docx.shared import ElementProxy, Emu, Length, Pt, RGBColor, Twips, lazyproperty
from docx.text.tabstops import TabStops

if TYPE_CHECKING:
    from docx.oxml.text.parfmt import CT_Border, CT_PBdr, CT_PPr


class ParagraphFormat(ElementProxy):
    """Provides access to paragraph formatting such as justification, indentation, line
    spacing, space before and after, and widow/orphan control."""

    @lazyproperty
    def borders(self) -> ParagraphBorders:
        """|ParagraphBorders| object providing access to the borders defined for this
        paragraph format."""
        pPr = self._element.get_or_add_pPr()
        return ParagraphBorders(pPr)

    def bottom_border(
        self,
        style: WD_BORDER_STYLE = WD_BORDER_STYLE.SINGLE,
        width: Length | None = None,
        color: RGBColor | str | None = None,
        space: Length | None = None,
    ) -> Border:
        """Convenience method to set the bottom border of this paragraph.

        Returns the |Border| object for the bottom border after applying the specified
        settings.
        """
        border = self.borders.bottom
        border.style = style
        if width is not None:
            border.width = width
        if color is not None:
            border.color = color if isinstance(color, RGBColor) else RGBColor.from_string(color)
        if space is not None:
            border.space = space
        return border

    @property
    def alignment(self):
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates paragraph alignment is inherited from the style
        hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value):
        pPr = self._element.get_or_add_pPr()
        pPr.jc_val = value

    @property
    def first_line_indent(self):
        """|Length| value specifying the relative difference in indentation for the
        first line of the paragraph.

        A positive value causes the first line to be indented. A negative value produces
        a hanging indent. |None| indicates first line indentation is inherited from the
        style hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.first_line_indent

    @first_line_indent.setter
    def first_line_indent(self, value):
        pPr = self._element.get_or_add_pPr()
        pPr.first_line_indent = value

    @property
    def keep_together(self):
        """|True| if the paragraph should be kept "in one piece" and not broken across a
        page boundary when the document is rendered.

        |None| indicates its effective value is inherited from the style hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.keepLines_val

    @keep_together.setter
    def keep_together(self, value):
        self._element.get_or_add_pPr().keepLines_val = value

    @property
    def keep_with_next(self):
        """|True| if the paragraph should be kept on the same page as the subsequent
        paragraph when the document is rendered.

        For example, this property could be used to keep a section heading on the same
        page as its first paragraph. |None| indicates its effective value is inherited
        from the style hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.keepNext_val

    @keep_with_next.setter
    def keep_with_next(self, value):
        self._element.get_or_add_pPr().keepNext_val = value

    @property
    def left_indent(self):
        """|Length| value specifying the space between the left margin and the left side
        of the paragraph.

        |None| indicates the left indent value is inherited from the style hierarchy.
        Use an |Inches| value object as a convenient way to apply indentation in units
        of inches.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.ind_left

    @left_indent.setter
    def left_indent(self, value):
        pPr = self._element.get_or_add_pPr()
        pPr.ind_left = value

    @property
    def line_spacing(self):
        """|float| or |Length| value specifying the space between baselines in
        successive lines of the paragraph.

        A value of |None| indicates line spacing is inherited from the style hierarchy.
        A float value, e.g. ``2.0`` or ``1.75``, indicates spacing is applied in
        multiples of line heights. A |Length| value such as ``Pt(12)`` indicates spacing
        is a fixed height. The |Pt| value class is a convenient way to apply line
        spacing in units of points. Assigning |None| resets line spacing to inherit from
        the style hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return self._line_spacing(pPr.spacing_line, pPr.spacing_lineRule)

    @line_spacing.setter
    def line_spacing(self, value):
        pPr = self._element.get_or_add_pPr()
        if value is None:
            pPr.spacing_line = None
            pPr.spacing_lineRule = None
        elif isinstance(value, Length):
            pPr.spacing_line = value
            if pPr.spacing_lineRule != WD_LINE_SPACING.AT_LEAST:
                pPr.spacing_lineRule = WD_LINE_SPACING.EXACTLY
        else:
            pPr.spacing_line = Emu(value * Twips(240))
            pPr.spacing_lineRule = WD_LINE_SPACING.MULTIPLE

    @property
    def line_spacing_rule(self):
        """A member of the :ref:`WdLineSpacing` enumeration indicating how the value of
        :attr:`line_spacing` should be interpreted.

        Assigning any of the :ref:`WdLineSpacing` members :attr:`SINGLE`,
        :attr:`DOUBLE`, or :attr:`ONE_POINT_FIVE` will cause the value of
        :attr:`line_spacing` to be updated to produce the corresponding line spacing.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return self._line_spacing_rule(pPr.spacing_line, pPr.spacing_lineRule)

    @line_spacing_rule.setter
    def line_spacing_rule(self, value):
        pPr = self._element.get_or_add_pPr()
        if value == WD_LINE_SPACING.SINGLE:
            pPr.spacing_line = Twips(240)
            pPr.spacing_lineRule = WD_LINE_SPACING.MULTIPLE
        elif value == WD_LINE_SPACING.ONE_POINT_FIVE:
            pPr.spacing_line = Twips(360)
            pPr.spacing_lineRule = WD_LINE_SPACING.MULTIPLE
        elif value == WD_LINE_SPACING.DOUBLE:
            pPr.spacing_line = Twips(480)
            pPr.spacing_lineRule = WD_LINE_SPACING.MULTIPLE
        else:
            pPr.spacing_lineRule = value

    @property
    def page_break_before(self):
        """|True| if the paragraph should appear at the top of the page following the
        prior paragraph.

        |None| indicates its effective value is inherited from the style hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.pageBreakBefore_val

    @page_break_before.setter
    def page_break_before(self, value):
        self._element.get_or_add_pPr().pageBreakBefore_val = value

    @property
    def right_indent(self):
        """|Length| value specifying the space between the right margin and the right
        side of the paragraph.

        |None| indicates the right indent value is inherited from the style hierarchy.
        Use a |Cm| value object as a convenient way to apply indentation in units of
        centimeters.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.ind_right

    @right_indent.setter
    def right_indent(self, value):
        pPr = self._element.get_or_add_pPr()
        pPr.ind_right = value

    @property
    def space_after(self):
        """|Length| value specifying the spacing to appear between this paragraph and
        the subsequent paragraph.

        |None| indicates this value is inherited from the style hierarchy. |Length|
        objects provide convenience properties, such as :attr:`~.Length.pt` and
        :attr:`~.Length.inches`, that allow easy conversion to various length units.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.spacing_after

    @space_after.setter
    def space_after(self, value):
        self._element.get_or_add_pPr().spacing_after = value

    @property
    def space_before(self):
        """|Length| value specifying the spacing to appear between this paragraph and
        the prior paragraph.

        |None| indicates this value is inherited from the style hierarchy. |Length|
        objects provide convenience properties, such as :attr:`~.Length.pt` and
        :attr:`~.Length.cm`, that allow easy conversion to various length units.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.spacing_before

    @space_before.setter
    def space_before(self, value):
        self._element.get_or_add_pPr().spacing_before = value

    @lazyproperty
    def tab_stops(self):
        """|TabStops| object providing access to the tab stops defined for this
        paragraph format."""
        pPr = self._element.get_or_add_pPr()
        return TabStops(pPr)

    @property
    def widow_control(self):
        """|True| if the first and last lines in the paragraph remain on the same page
        as the rest of the paragraph when Word repaginates the document.

        |None| indicates its effective value is inherited from the style hierarchy.
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.widowControl_val

    @widow_control.setter
    def widow_control(self, value):
        self._element.get_or_add_pPr().widowControl_val = value

    @staticmethod
    def _line_spacing(spacing_line, spacing_lineRule):
        """Return the line spacing value calculated from the combination of
        `spacing_line` and `spacing_lineRule`.

        Returns a |float| number of lines when `spacing_lineRule` is
        ``WD_LINE_SPACING.MULTIPLE``, otherwise a |Length| object of absolute line
        height is returned. Returns |None| when `spacing_line` is |None|.
        """
        if spacing_line is None:
            return None
        if spacing_lineRule == WD_LINE_SPACING.MULTIPLE:
            return spacing_line / Pt(12)
        return spacing_line

    @staticmethod
    def _line_spacing_rule(line, lineRule):
        """Return the line spacing rule value calculated from the combination of `line`
        and `lineRule`.

        Returns special members of the :ref:`WdLineSpacing` enumeration when line
        spacing is single, double, or 1.5 lines.
        """
        if lineRule == WD_LINE_SPACING.MULTIPLE:
            if line == Twips(240):
                return WD_LINE_SPACING.SINGLE
            if line == Twips(360):
                return WD_LINE_SPACING.ONE_POINT_FIVE
            if line == Twips(480):
                return WD_LINE_SPACING.DOUBLE
        return lineRule


class ParagraphBorders(ElementProxy):
    """Provides access to the border settings of a paragraph.

    Accessed using the :attr:`~.ParagraphFormat.borders` property of ParagraphFormat;
    not intended to be constructed directly.
    """

    def __init__(self, element: CT_PPr):
        super().__init__(element, None)
        self._pPr = element

    @property
    def bottom(self) -> Border:
        """The |Border| object for the bottom border."""
        return Border(self._pPr, "bottom")

    @property
    def top(self) -> Border:
        """The |Border| object for the top border."""
        return Border(self._pPr, "top")

    @property
    def left(self) -> Border:
        """The |Border| object for the left border."""
        return Border(self._pPr, "left")

    @property
    def right(self) -> Border:
        """The |Border| object for the right border."""
        return Border(self._pPr, "right")

    @property
    def between(self) -> Border:
        """The |Border| object for the between border."""
        return Border(self._pPr, "between")


class Border:
    """Provides access to a single border's properties (style, width, color, space).

    Lazily creates the underlying XML element on first write.
    """

    _SIDE_GETTERS = {
        "top": lambda pBdr: pBdr.top,
        "bottom": lambda pBdr: pBdr.bottom,
        "left": lambda pBdr: pBdr.left,
        "right": lambda pBdr: pBdr.right,
        "between": lambda pBdr: pBdr.between,
    }

    _SIDE_ADDERS = {
        "top": lambda pBdr: pBdr.get_or_add_top(),
        "bottom": lambda pBdr: pBdr.get_or_add_bottom(),
        "left": lambda pBdr: pBdr.get_or_add_left(),
        "right": lambda pBdr: pBdr.get_or_add_right(),
        "between": lambda pBdr: pBdr.get_or_add_between(),
    }

    def __init__(self, pPr: CT_PPr, side: str):
        self._pPr = pPr
        self._side = side

    def _get_border(self) -> CT_Border | None:
        pBdr = self._pPr.pBdr
        if pBdr is None:
            return None
        return self._SIDE_GETTERS[self._side](pBdr)

    def _get_or_add_border(self) -> CT_Border:
        pBdr = self._pPr.get_or_add_pBdr()
        return self._SIDE_ADDERS[self._side](pBdr)

    @property
    def style(self) -> WD_BORDER_STYLE | None:
        """The border style as a member of :ref:`WdBorderStyle`, or |None| if no
        border is defined."""
        border = self._get_border()
        if border is None:
            return None
        return border.val

    @style.setter
    def style(self, value: WD_BORDER_STYLE | None):
        if value is None:
            border = self._get_border()
            if border is not None:
                border.val = None
            return
        self._get_or_add_border().val = value

    @property
    def width(self) -> Length | None:
        """The border width as a |Length| value, or |None| if not specified.

        The XML ``w:sz`` attribute is stored in eighths of a point.
        """
        border = self._get_border()
        if border is None:
            return None
        return border.sz

    @width.setter
    def width(self, value: Length | None):
        if value is None:
            border = self._get_border()
            if border is not None:
                border.sz = None
            return
        self._get_or_add_border().sz = value

    @property
    def color(self) -> RGBColor | None:
        """The border color as an |RGBColor|, or |None| if not specified."""
        border = self._get_border()
        if border is None:
            return None
        return border.color

    @color.setter
    def color(self, value: RGBColor | None):
        if value is None:
            border = self._get_border()
            if border is not None:
                border.color = None
            return
        self._get_or_add_border().color = value

    @property
    def space(self) -> Length | None:
        """The spacing between the border and paragraph text as a |Length| value, or
        |None| if not specified.

        The XML ``w:space`` attribute is stored in points.
        """
        border = self._get_border()
        if border is None:
            return None
        return border.space

    @space.setter
    def space(self, value: Length | None):
        if value is None:
            border = self._get_border()
            if border is not None:
                border.space = None
            return
        self._get_or_add_border().space = value
