"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.table import WD_SHADING_PATTERN
from docx.enum.text import WD_LINE_SPACING, WD_OUTLINELVL
from docx.shared import ElementProxy, Emu, Length, Pt, RGBColor, Twips, lazyproperty
from docx.text.tabstops import TabStops

if TYPE_CHECKING:
    from docx.enum.text import (
        WD_BORDER_STYLE,
        WD_FRAME_DROP_CAP,
        WD_FRAME_H_ALIGN,
        WD_FRAME_H_ANCHOR,
        WD_FRAME_V_ALIGN,
        WD_FRAME_V_ANCHOR,
        WD_FRAME_WRAP,
    )
    from docx.oxml.text.parfmt import CT_Border, CT_FramePr
    from docx.text.font import Font


class ParagraphFormat(ElementProxy):
    """Provides access to paragraph formatting such as justification, indentation, line
    spacing, space before and after, and widow/orphan control."""

    @property
    def borders(self) -> ParagraphBorders:
        """|ParagraphBorders| object providing access to the border settings for this
        paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return ParagraphBorders(self._element)

    @property
    def frame(self) -> TextFrame | None:
        """|TextFrame| proxy for this paragraph's ``w:framePr`` element, or |None|.

        Returns |None| when the paragraph has no ``w:pPr/w:framePr`` child. A text
        frame is an absolutely-positioned text container, the legacy predecessor
        to text boxes.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        framePr = pPr.framePr
        if framePr is None:
            return None
        return TextFrame(framePr)

    def set_frame(
        self,
        *,
        width: Length | None = None,
        height: Length | None = None,
        horizontal_position: Length | None = None,
        vertical_position: Length | None = None,
        horizontal_anchor: WD_FRAME_H_ANCHOR | None = None,
        vertical_anchor: WD_FRAME_V_ANCHOR | None = None,
        wrap: WD_FRAME_WRAP | None = None,
        drop_cap: WD_FRAME_DROP_CAP | None = None,
        lines: int | None = None,
        horizontal_alignment: WD_FRAME_H_ALIGN | None = None,
        vertical_alignment: WD_FRAME_V_ALIGN | None = None,
    ) -> TextFrame:
        """Create or update the ``w:framePr`` element on this paragraph.

        Any keyword argument left at its default of |None| is left unchanged when the
        frame already exists. To clear an attribute, use the corresponding setter on
        the returned |TextFrame| (e.g. ``frame.width = None``) or call
        :meth:`remove_frame` to drop the frame entirely.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.get_or_add_pPr()
        framePr = pPr.get_or_add_framePr()
        frame = TextFrame(framePr)
        if width is not None:
            frame.width = width
        if height is not None:
            frame.height = height
        if horizontal_position is not None:
            frame.horizontal_position = horizontal_position
        if vertical_position is not None:
            frame.vertical_position = vertical_position
        if horizontal_anchor is not None:
            frame.horizontal_anchor = horizontal_anchor
        if vertical_anchor is not None:
            frame.vertical_anchor = vertical_anchor
        if wrap is not None:
            frame.wrap = wrap
        if drop_cap is not None:
            frame.drop_cap = drop_cap
        if lines is not None:
            frame.lines = lines
        if horizontal_alignment is not None:
            frame.horizontal_alignment = horizontal_alignment
        if vertical_alignment is not None:
            frame.vertical_alignment = vertical_alignment
        return frame

    def remove_frame(self) -> None:
        """Remove the ``w:framePr`` element, if present.

        No-op when no ``w:pPr`` or no ``w:framePr`` child is present.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return
        if pPr.framePr is None:
            return
        pPr._remove_framePr()

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
    def kinsoku(self) -> bool | None:
        """Tri-state value controlling Japanese kinsoku line-break rules.

        Maps to ``w:pPr/w:kinsoku``. Returns |True| when the element is
        present and its ``w:val`` is truthy, |False| when present but
        explicitly turned off, and |None| when the element is absent
        (inherited from the style hierarchy). Kinsoku rules constrain
        punctuation from appearing at the start or end of a line.

        Assigning |True| or |False| inserts the element. Assigning |None|
        removes it.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.kinsoku_val

    @kinsoku.setter
    def kinsoku(self, value: bool | None) -> None:
        pPr = self._element.get_or_add_pPr()
        pPr.kinsoku_val = value

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
    def right_to_left(self) -> bool:
        """|True| if paragraph uses right-to-left (bidirectional) layout.

        Maps to the ``w:pPr/w:bidi`` element. Returns |False| when the element is
        absent. Assigning |True| inserts ``w:bidi``; assigning |False| or |None|
        removes it. When |True|, paragraph-level runs are laid out right-to-left
        (e.g. for Arabic, Hebrew, or Farsi text).

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return False
        return pPr.bidi_val

    @right_to_left.setter
    def right_to_left(self, value: bool | None):
        pPr = self._element.get_or_add_pPr()
        pPr.bidi_val = value

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
    def word_wrap(self) -> bool | None:
        """Tri-state value controlling Latin-text word-wrap behaviour.

        Maps to ``w:pPr/w:wordWrap``. Returns |True| when Latin text wraps
        on word boundaries (the default behaviour), |False| when the
        paragraph uses aggressive Asian word-wrap (allowing breaks within
        words), and |None| when the element is absent (inherited).

        Assigning |True| or |False| inserts the element. Assigning |None|
        removes it.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.wordWrap_val

    @word_wrap.setter
    def word_wrap(self, value: bool | None) -> None:
        pPr = self._element.get_or_add_pPr()
        pPr.wordWrap_val = value

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

    @property
    def outline_level(self) -> WD_OUTLINELVL | None:
        """Outline level (``w:pPr/w:outlineLvl``), or |None| if not set.

        Values are members of :ref:`WdOutlineLvl` — ``LEVEL_1`` through
        ``LEVEL_10`` for heading levels and ``BODY_TEXT`` for body text
        (``10``). Returns |None| when the element is absent (inherited from
        the style hierarchy).

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        val = pPr.outlineLvl_val
        if val is None:
            return None
        return WD_OUTLINELVL(val)

    @outline_level.setter
    def outline_level(self, value: WD_OUTLINELVL | int | None) -> None:
        pPr = self._element.get_or_add_pPr()
        if value is None:
            pPr.outlineLvl_val = None
            return
        val = int(value)
        if not 0 <= val <= 10:
            raise ValueError(
                "outline_level must be 0..10 or a WD_OUTLINELVL member, got %r" % (value,)
            )
        pPr.outlineLvl_val = val

    @property
    def contextual_spacing(self) -> bool | None:
        """Tri-state value controlling ``w:pPr/w:contextualSpacing``.

        When |True|, space above and below this paragraph is ignored when the
        neighbouring paragraph uses the same paragraph style (typical for list
        items). Returns |None| when the element is absent (inherited from the
        style hierarchy).

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.contextualSpacing_val

    @contextual_spacing.setter
    def contextual_spacing(self, value: bool | None) -> None:
        pPr = self._element.get_or_add_pPr()
        pPr.contextualSpacing_val = value

    @property
    def first_line_chars(self) -> int | None:
        """Value of ``w:pPr/w:ind/@w:firstLineChars``, or |None| if not set.

        Specifies the first-line indent in units of 1/100 of a character
        width (the "character unit" used for East-Asian layouts). Returns
        |None| when the ``w:ind`` element or the attribute is absent.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.first_line_chars

    @first_line_chars.setter
    def first_line_chars(self, value: int | None) -> None:
        if value is None and (self._element.pPr is None or self._element.pPr.ind is None):
            return
        pPr = self._element.get_or_add_pPr()
        pPr.first_line_chars = value

    @property
    def auto_space_de(self) -> bool | None:
        """Tri-state value controlling ``w:pPr/w:autoSpaceDE``.

        When |True|, automatically adjusts spacing between East-Asian and
        Latin text in this paragraph. Returns |None| when the element is
        absent (inherited from the style hierarchy).

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.autoSpaceDE_val

    @auto_space_de.setter
    def auto_space_de(self, value: bool | None) -> None:
        pPr = self._element.get_or_add_pPr()
        pPr.autoSpaceDE_val = value

    @property
    def auto_space_dn(self) -> bool | None:
        """Tri-state value controlling ``w:pPr/w:autoSpaceDN``.

        When |True|, automatically adjusts spacing between East-Asian text
        and numerals in this paragraph. Returns |None| when the element is
        absent (inherited from the style hierarchy).

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        return pPr.autoSpaceDN_val

    @auto_space_dn.setter
    def auto_space_dn(self, value: bool | None) -> None:
        pPr = self._element.get_or_add_pPr()
        pPr.autoSpaceDN_val = value

    @property
    def shading_color(self) -> RGBColor | None:
        """Paragraph-level background (shading) color as an |RGBColor|, or |None|.

        Read/write. Reads the ``w:fill`` attribute of ``w:pPr/w:shd``.
        Returns |None| when ``w:shd`` is absent or its ``w:fill`` is missing
        or set to ``"auto"``.

        Assigning an |RGBColor| writes ``w:pPr/w:shd`` with ``w:val="clear"``
        and ``w:fill="RRGGBB"``. Assigning |None| removes the ``w:shd``
        child. Mirrors :attr:`Font.shading_color` but applies at the
        paragraph level.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._element.pPr
        if pPr is None:
            return None
        shd = pPr.shd
        if shd is None:
            return None
        fill = shd.fill
        if fill is None or not isinstance(fill, RGBColor):
            return None
        return fill

    @shading_color.setter
    def shading_color(self, value: RGBColor | None) -> None:
        if value is None:
            pPr = self._element.pPr
            if pPr is None:
                return
            pPr._remove_shd()  # pyright: ignore[reportPrivateUsage]
            return
        pPr = self._element.get_or_add_pPr()
        shd = pPr.get_or_add_shd()
        shd.val = WD_SHADING_PATTERN.CLEAR
        shd.fill = value

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


class ParagraphBorders:
    """Provides access to the border settings for a paragraph.

    Accessed via the :attr:`ParagraphFormat.borders` property.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, element: object):
        self._element = element

    @property
    def top(self) -> Border:
        """The |Border| object for the top edge of the paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return Border(self._element, "top")

    @property
    def bottom(self) -> Border:
        """The |Border| object for the bottom edge of the paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return Border(self._element, "bottom")

    @property
    def left(self) -> Border:
        """The |Border| object for the left edge of the paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return Border(self._element, "left")

    @property
    def right(self) -> Border:
        """The |Border| object for the right edge of the paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return Border(self._element, "right")

    @property
    def between(self) -> Border:
        """The |Border| object for the border between identical paragraphs.

        .. versionadded:: 1.3.0.dev0
        """
        return Border(self._element, "between")

    @property
    def bar(self) -> Border:
        """The |Border| object for the bar border of the paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return Border(self._element, "bar")


class Border:
    """Provides access to a single border edge of a paragraph.

    Accessed via the properties of |ParagraphBorders|, e.g.
    ``paragraph_format.borders.bottom``.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, element: object, side: str):
        self._element = element
        self._side = side

    @property
    def _border_elm(self) -> CT_Border | None:
        pPr = self._element.pPr  # type: ignore[attr-defined]
        if pPr is None:
            return None
        pBdr = pPr.pBdr
        if pBdr is None:
            return None
        return getattr(pBdr, self._side)

    def _get_or_add_border_elm(self) -> CT_Border:
        pPr = self._element.get_or_add_pPr()  # type: ignore[attr-defined]
        pBdr = pPr.get_or_add_pBdr()
        return getattr(pBdr, f"get_or_add_{self._side}")()

    @property
    def style(self) -> WD_BORDER_STYLE | None:
        """The border style as a member of :ref:`WdBorderStyle`, or |None| if no border
        is defined.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border_elm
        if border is None:
            return None
        return border.val

    @style.setter
    def style(self, value: WD_BORDER_STYLE | None) -> None:
        if value is None:
            pPr = self._element.pPr  # type: ignore[attr-defined]
            if pPr is not None:
                pBdr = pPr.pBdr
                if pBdr is not None:
                    remove_fn = getattr(pBdr, f"_remove_{self._side}", None)
                    if remove_fn is not None:
                        remove_fn()
            return
        self._get_or_add_border_elm().val = value

    @property
    def width(self) -> Length | None:
        """The border width as a |Length| value, or |None| if not defined.

        Stored in the XML as eighths of a point in the ``w:sz`` attribute.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border_elm
        if border is None:
            return None
        return border.sz

    @width.setter
    def width(self, value: Length | None) -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.sz = None
            return
        self._get_or_add_border_elm().sz = value

    @property
    def color(self) -> RGBColor | None:
        """|RGBColor| value of the border color, or |None| if not defined.

        An ``"auto"`` value in the XML is returned as |None|.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border_elm
        if border is None:
            return None
        color = border.color
        if isinstance(color, str):
            return None
        return color

    @color.setter
    def color(self, value: RGBColor | None) -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.color = None
            return
        self._get_or_add_border_elm().color = value

    @property
    def space(self) -> Length | None:
        """The spacing between the border and paragraph text as a |Length| value, or
        |None| if not defined.

        .. versionadded:: 1.3.0.dev0
        """
        border = self._border_elm
        if border is None:
            return None
        return border.space

    @space.setter
    def space(self, value: Length | None) -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.space = None
            return
        self._get_or_add_border_elm().space = value


class TextFrame:
    """Proxy object for a paragraph-level text frame (``w:framePr``).

    Provides read/write access to the attributes of the ``w:framePr`` element. A
    text frame is an absolutely-positioned text container, the legacy predecessor
    to text boxes.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, framePr: CT_FramePr):
        self._framePr = framePr

    @property
    def width(self) -> Length | None:
        """Frame width (``w:framePr/@w:w``) as a |Length|, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.w

    @width.setter
    def width(self, value: Length | None) -> None:
        self._framePr.w = value

    @property
    def height(self) -> Length | None:
        """Frame height (``w:framePr/@w:h``) as a |Length|, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.h

    @height.setter
    def height(self, value: Length | None) -> None:
        self._framePr.h = value

    @property
    def horizontal_position(self) -> Length | None:
        """Horizontal position (``w:framePr/@w:x``) as a |Length|, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.x

    @horizontal_position.setter
    def horizontal_position(self, value: Length | None) -> None:
        self._framePr.x = value

    @property
    def vertical_position(self) -> Length | None:
        """Vertical position (``w:framePr/@w:y``) as a |Length|, or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.y

    @vertical_position.setter
    def vertical_position(self, value: Length | None) -> None:
        self._framePr.y = value

    @property
    def horizontal_anchor(self) -> WD_FRAME_H_ANCHOR | None:
        """Horizontal anchor (``w:framePr/@w:hAnchor``), or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.hAnchor

    @horizontal_anchor.setter
    def horizontal_anchor(self, value: WD_FRAME_H_ANCHOR | None) -> None:
        self._framePr.hAnchor = value

    @property
    def vertical_anchor(self) -> WD_FRAME_V_ANCHOR | None:
        """Vertical anchor (``w:framePr/@w:vAnchor``), or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.vAnchor

    @vertical_anchor.setter
    def vertical_anchor(self, value: WD_FRAME_V_ANCHOR | None) -> None:
        self._framePr.vAnchor = value

    @property
    def wrap(self) -> WD_FRAME_WRAP | None:
        """Text-wrap behaviour (``w:framePr/@w:wrap``), or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.wrap

    @wrap.setter
    def wrap(self, value: WD_FRAME_WRAP | None) -> None:
        self._framePr.wrap = value

    @property
    def drop_cap(self) -> WD_FRAME_DROP_CAP | None:
        """Drop-cap positioning (``w:framePr/@w:dropCap``), or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.dropCap

    @drop_cap.setter
    def drop_cap(self, value: WD_FRAME_DROP_CAP | None) -> None:
        self._framePr.dropCap = value

    @property
    def lines(self) -> int | None:
        """Number of lines for a drop-cap frame (``w:framePr/@w:lines``), or |None|.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.lines

    @lines.setter
    def lines(self, value: int | None) -> None:
        self._framePr.lines = value

    @property
    def horizontal_alignment(self) -> WD_FRAME_H_ALIGN | None:
        """Horizontal alignment (``w:framePr/@w:xAlign``), or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.xAlign

    @horizontal_alignment.setter
    def horizontal_alignment(self, value: WD_FRAME_H_ALIGN | None) -> None:
        self._framePr.xAlign = value

    @property
    def vertical_alignment(self) -> WD_FRAME_V_ALIGN | None:
        """Vertical alignment (``w:framePr/@w:yAlign``), or |None| if not set.

        .. versionadded:: 1.3.0.dev0
        """
        return self._framePr.yAlign

    @vertical_alignment.setter
    def vertical_alignment(self, value: WD_FRAME_V_ALIGN | None) -> None:
        self._framePr.yAlign = value
