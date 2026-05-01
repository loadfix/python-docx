"""Custom element classes related to paragraph properties (CT_PPr)."""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Callable

from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_BORDER_STYLE,
    WD_FRAME_DROP_CAP,
    WD_FRAME_H_ALIGN,
    WD_FRAME_H_ANCHOR,
    WD_FRAME_V_ALIGN,
    WD_FRAME_V_ANCHOR,
    WD_FRAME_WRAP,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.oxml.shared import CT_DecimalNumber, CT_OnOff
from docx.oxml.simpletypes import (
    ST_DecimalNumber,
    ST_EighthPointMeasure,
    ST_HexColor,
    ST_PointMeasure,
    ST_SignedTwipsMeasure,
    ST_TwipsMeasure,
)
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneOrMore,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)
from docx.shared import Length

if TYPE_CHECKING:
    from docx.oxml.section import CT_SectPr
    from docx.oxml.shared import CT_String
    from docx.shared import RGBColor


class CT_Border(BaseOxmlElement):
    """``<w:top>``, ``<w:bottom>``, ``<w:left>``, ``<w:right>``, ``<w:between>``,
    ``<w:bar>`` element.

    Defines a single paragraph border edge.
    """

    val: WD_BORDER_STYLE | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", WD_BORDER_STYLE
    )
    sz: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:sz", ST_EighthPointMeasure
    )
    space: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:space", ST_PointMeasure
    )
    color: RGBColor | str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:color", ST_HexColor
    )


class CT_PBdr(BaseOxmlElement):
    """``<w:pBdr>`` element, containing the paragraph border properties."""

    _tag_seq = ("w:top", "w:left", "w:bottom", "w:right", "w:between", "w:bar")
    top: CT_Border | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:top", successors=_tag_seq[1:]
    )
    left: CT_Border | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:left", successors=_tag_seq[2:]
    )
    bottom: CT_Border | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:bottom", successors=_tag_seq[3:]
    )
    right: CT_Border | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:right", successors=_tag_seq[4:]
    )
    between: CT_Border | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:between", successors=_tag_seq[5:]
    )
    bar: CT_Border | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:bar", successors=()
    )
    del _tag_seq


class CT_FramePr(BaseOxmlElement):
    """``<w:framePr>`` element, specifying text-frame properties for a paragraph.

    A text frame is an absolutely-positioned text container, the legacy predecessor
    to text boxes. All attributes are optional.
    """

    w: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:w", ST_TwipsMeasure
    )
    h: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:h", ST_TwipsMeasure
    )
    x: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:x", ST_SignedTwipsMeasure
    )
    y: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:y", ST_SignedTwipsMeasure
    )
    hAnchor: WD_FRAME_H_ANCHOR | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:hAnchor", WD_FRAME_H_ANCHOR
    )
    vAnchor: WD_FRAME_V_ANCHOR | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:vAnchor", WD_FRAME_V_ANCHOR
    )
    wrap: WD_FRAME_WRAP | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:wrap", WD_FRAME_WRAP
    )
    dropCap: WD_FRAME_DROP_CAP | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:dropCap", WD_FRAME_DROP_CAP
    )
    lines: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:lines", ST_DecimalNumber
    )
    xAlign: WD_FRAME_H_ALIGN | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:xAlign", WD_FRAME_H_ALIGN
    )
    yAlign: WD_FRAME_V_ALIGN | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:yAlign", WD_FRAME_V_ALIGN
    )


class CT_Ind(BaseOxmlElement):
    """``<w:ind>`` element, specifying paragraph indentation."""

    left: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:left", ST_SignedTwipsMeasure
    )
    right: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:right", ST_SignedTwipsMeasure
    )
    firstLine: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:firstLine", ST_TwipsMeasure
    )
    hanging: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:hanging", ST_TwipsMeasure
    )


class CT_Jc(BaseOxmlElement):
    """``<w:jc>`` element, specifying paragraph justification."""

    val: WD_ALIGN_PARAGRAPH = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", WD_ALIGN_PARAGRAPH
    )


class CT_PPr(BaseOxmlElement):
    """``<w:pPr>`` element, containing the properties for a paragraph."""

    get_or_add_bidi: Callable[[], CT_OnOff]
    get_or_add_framePr: Callable[[], CT_FramePr]
    get_or_add_ind: Callable[[], CT_Ind]
    get_or_add_pBdr: Callable[[], CT_PBdr]
    get_or_add_pStyle: Callable[[], CT_String]
    get_or_add_sectPr: Callable[[], CT_SectPr]
    _insert_sectPr: Callable[[CT_SectPr], None]
    _remove_bidi: Callable[[], None]
    _remove_framePr: Callable[[], None]
    _remove_pBdr: Callable[[], None]
    _remove_pStyle: Callable[[], None]
    _remove_sectPr: Callable[[], None]

    _tag_seq = (
        "w:pStyle",
        "w:keepNext",
        "w:keepLines",
        "w:pageBreakBefore",
        "w:framePr",
        "w:widowControl",
        "w:numPr",
        "w:suppressLineNumbers",
        "w:pBdr",
        "w:shd",
        "w:tabs",
        "w:suppressAutoHyphens",
        "w:kinsoku",
        "w:wordWrap",
        "w:overflowPunct",
        "w:topLinePunct",
        "w:autoSpaceDE",
        "w:autoSpaceDN",
        "w:bidi",
        "w:adjustRightInd",
        "w:snapToGrid",
        "w:spacing",
        "w:ind",
        "w:contextualSpacing",
        "w:mirrorIndents",
        "w:suppressOverlap",
        "w:jc",
        "w:textDirection",
        "w:textAlignment",
        "w:textboxTightWrap",
        "w:outlineLvl",
        "w:divId",
        "w:cnfStyle",
        "w:rPr",
        "w:sectPr",
        "w:pPrChange",
    )
    pStyle: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:pStyle", successors=_tag_seq[1:]
    )
    keepNext = ZeroOrOne("w:keepNext", successors=_tag_seq[2:])
    keepLines = ZeroOrOne("w:keepLines", successors=_tag_seq[3:])
    pageBreakBefore = ZeroOrOne("w:pageBreakBefore", successors=_tag_seq[4:])
    framePr: CT_FramePr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:framePr", successors=_tag_seq[5:]
    )
    widowControl = ZeroOrOne("w:widowControl", successors=_tag_seq[6:])
    numPr = ZeroOrOne("w:numPr", successors=_tag_seq[7:])
    pBdr: CT_PBdr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:pBdr", successors=_tag_seq[9:]
    )
    tabs = ZeroOrOne("w:tabs", successors=_tag_seq[11:])
    bidi: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:bidi", successors=_tag_seq[19:]
    )
    spacing = ZeroOrOne("w:spacing", successors=_tag_seq[22:])
    ind: CT_Ind | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:ind", successors=_tag_seq[23:]
    )
    jc = ZeroOrOne("w:jc", successors=_tag_seq[27:])
    outlineLvl: CT_DecimalNumber = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:outlineLvl", successors=_tag_seq[31:]
    )
    sectPr = ZeroOrOne("w:sectPr", successors=_tag_seq[35:])
    pPrChange = ZeroOrOne("w:pPrChange", successors=())
    del _tag_seq

    @property
    def bidi_val(self) -> bool:
        """Value of `w:bidi/@val` or |False| if `./w:bidi` is not present."""
        bidi = self.bidi
        if bidi is None:
            return False
        return bidi.val

    @bidi_val.setter
    def bidi_val(self, value: bool | None):
        if value in [None, False]:
            self._remove_bidi()
        else:
            self.get_or_add_bidi().val = True

    @property
    def first_line_indent(self) -> Length | None:
        """A |Length| value calculated from the values of `w:ind/@w:firstLine` and
        `w:ind/@w:hanging`.

        Returns |None| if the `w:ind` child is not present.
        """
        ind = self.ind
        if ind is None:
            return None
        hanging = ind.hanging
        if hanging is not None:
            return Length(-hanging)
        firstLine = ind.firstLine
        if firstLine is None:
            return None
        return firstLine

    @first_line_indent.setter
    def first_line_indent(self, value: Length | None):
        if self.ind is None and value is None:
            return
        ind = self.get_or_add_ind()
        ind.firstLine = ind.hanging = None
        if value is None:
            return
        elif value < 0:
            ind.hanging = -value
        else:
            ind.firstLine = value

    @property
    def ind_left(self) -> Length | None:
        """The value of `w:ind/@w:left` or |None| if not present."""
        ind = self.ind
        if ind is None:
            return None
        return ind.left

    @ind_left.setter
    def ind_left(self, value: Length | None):
        if value is None and self.ind is None:
            return
        ind = self.get_or_add_ind()
        ind.left = value

    @property
    def ind_right(self) -> Length | None:
        """The value of `w:ind/@w:right` or |None| if not present."""
        ind = self.ind
        if ind is None:
            return None
        return ind.right

    @ind_right.setter
    def ind_right(self, value: Length | None):
        if value is None and self.ind is None:
            return
        ind = self.get_or_add_ind()
        ind.right = value

    @property
    def jc_val(self) -> WD_ALIGN_PARAGRAPH | None:
        """Value of the `<w:jc>` child element or |None| if not present."""
        return self.jc.val if self.jc is not None else None

    @jc_val.setter
    def jc_val(self, value):
        if value is None:
            self._remove_jc()
            return
        self.get_or_add_jc().val = value

    @property
    def keepLines_val(self):
        """The value of `keepLines/@val` or |None| if not present."""
        keepLines = self.keepLines
        if keepLines is None:
            return None
        return keepLines.val

    @keepLines_val.setter
    def keepLines_val(self, value):
        if value is None:
            self._remove_keepLines()
        else:
            self.get_or_add_keepLines().val = value

    @property
    def keepNext_val(self):
        """The value of `keepNext/@val` or |None| if not present."""
        keepNext = self.keepNext
        if keepNext is None:
            return None
        return keepNext.val

    @keepNext_val.setter
    def keepNext_val(self, value):
        if value is None:
            self._remove_keepNext()
        else:
            self.get_or_add_keepNext().val = value

    @property
    def pageBreakBefore_val(self):
        """The value of `pageBreakBefore/@val` or |None| if not present."""
        pageBreakBefore = self.pageBreakBefore
        if pageBreakBefore is None:
            return None
        return pageBreakBefore.val

    @pageBreakBefore_val.setter
    def pageBreakBefore_val(self, value):
        if value is None:
            self._remove_pageBreakBefore()
        else:
            self.get_or_add_pageBreakBefore().val = value

    @property
    def spacing_after(self):
        """The value of `w:spacing/@w:after` or |None| if not present."""
        spacing = self.spacing
        if spacing is None:
            return None
        return spacing.after

    @spacing_after.setter
    def spacing_after(self, value):
        if value is None and self.spacing is None:
            return
        self.get_or_add_spacing().after = value

    @property
    def spacing_before(self):
        """The value of `w:spacing/@w:before` or |None| if not present."""
        spacing = self.spacing
        if spacing is None:
            return None
        return spacing.before

    @spacing_before.setter
    def spacing_before(self, value):
        if value is None and self.spacing is None:
            return
        self.get_or_add_spacing().before = value

    @property
    def spacing_line(self):
        """The value of `w:spacing/@w:line` or |None| if not present."""
        spacing = self.spacing
        if spacing is None:
            return None
        return spacing.line

    @spacing_line.setter
    def spacing_line(self, value):
        if value is None and self.spacing is None:
            return
        self.get_or_add_spacing().line = value

    @property
    def spacing_lineRule(self):
        """The value of `w:spacing/@w:lineRule` as a member of the :ref:`WdLineSpacing`
        enumeration.

        Only the `MULTIPLE`, `EXACTLY`, and `AT_LEAST` members are used. It is the
        responsibility of the client to calculate the use of `SINGLE`, `DOUBLE`, and
        `MULTIPLE` based on the value of `w:spacing/@w:line` if that behavior is
        desired.
        """
        spacing = self.spacing
        if spacing is None:
            return None
        lineRule = spacing.lineRule
        if lineRule is None and spacing.line is not None:
            return WD_LINE_SPACING.MULTIPLE
        return lineRule

    @spacing_lineRule.setter
    def spacing_lineRule(self, value):
        if value is None and self.spacing is None:
            return
        self.get_or_add_spacing().lineRule = value

    @property
    def style(self) -> str | None:
        """String contained in `./w:pStyle/@val`, or None if child is not present."""
        pStyle = self.pStyle
        if pStyle is None:
            return None
        return pStyle.val

    @style.setter
    def style(self, style: str | None):
        """Set `./w:pStyle/@val` `style`, adding a new element if necessary.

        If `style` is |None|, remove `./w:pStyle` when present.
        """
        if style is None:
            self._remove_pStyle()
            return
        pStyle = self.get_or_add_pStyle()
        pStyle.val = style

    @property
    def widowControl_val(self):
        """The value of `widowControl/@val` or |None| if not present."""
        widowControl = self.widowControl
        if widowControl is None:
            return None
        return widowControl.val

    @widowControl_val.setter
    def widowControl_val(self, value):
        if value is None:
            self._remove_widowControl()
        else:
            self.get_or_add_widowControl().val = value


class CT_Spacing(BaseOxmlElement):
    """``<w:spacing>`` element, specifying paragraph spacing attributes such as space
    before and line spacing.

    Also used as ``w:rPr/w:spacing`` for character spacing, where only the ``w:val``
    attribute is used (type ``ST_SignedTwipsMeasure``).
    """

    after = OptionalAttribute("w:after", ST_TwipsMeasure)
    before = OptionalAttribute("w:before", ST_TwipsMeasure)
    line = OptionalAttribute("w:line", ST_SignedTwipsMeasure)
    lineRule = OptionalAttribute("w:lineRule", WD_LINE_SPACING)
    val: Length | None = OptionalAttribute("w:val", ST_SignedTwipsMeasure)


class CT_TabStop(BaseOxmlElement):
    """`<w:tab>` element, representing an individual tab stop.

    Overloaded to use for a tab-character in a run, which also uses the w:tab tag but
    only needs a __str__ method.
    """

    val: WD_TAB_ALIGNMENT = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", WD_TAB_ALIGNMENT
    )
    leader: WD_TAB_LEADER | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:leader", WD_TAB_LEADER, default=WD_TAB_LEADER.SPACES
    )
    pos: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:pos", ST_SignedTwipsMeasure
    )

    def __str__(self) -> str:
        """Text equivalent of a `w:tab` element appearing in a run.

        Allows text of run inner-content to be accessed consistently across all text
        inner-content.
        """
        return "\t"


class CT_TabStops(BaseOxmlElement):
    """``<w:tabs>`` element, container for a sorted sequence of tab stops."""

    tab = OneOrMore("w:tab", successors=())

    def insert_tab_in_order(self, pos, align, leader):
        """Insert a newly created `w:tab` child element in `pos` order."""
        new_tab = self._new_tab()
        new_tab.pos, new_tab.val, new_tab.leader = pos, align, leader
        for tab in self.tab_lst:
            if new_tab.pos < tab.pos:
                tab.addprevious(new_tab)
                return new_tab
        self.append(new_tab)
        return new_tab
