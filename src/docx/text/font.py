"""Font-related proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.dml.color import ColorFormat
from docx.enum.table import WD_SHADING_PATTERN
from docx.enum.text import WD_UNDERLINE
from docx.shared import ElementProxy, Emu, RGBColor

if TYPE_CHECKING:
    from docx.enum.text import WD_BORDER_STYLE, WD_COLOR_INDEX
    from docx.oxml.text.run import CT_R
    from docx.shared import Length


class Font(ElementProxy):
    """Proxy object for parent of a `<w:rPr>` element and providing access to
    character properties such as font name, font size, bold, and subscript."""

    def __init__(self, r: CT_R, parent: Any | None = None):
        super().__init__(r, parent)
        self._element = r
        self._r = r

    @property
    def all_caps(self) -> bool | None:
        """Read/write.

        Causes text in this font to appear in capital letters.
        """
        return self._get_bool_prop("caps")

    @all_caps.setter
    def all_caps(self, value: bool | None) -> None:
        self._set_bool_prop("caps", value)

    @property
    def character_spacing(self) -> Length | None:
        """Read/write.

        |Length| value specifying the spacing between characters. Positive values expand
        the spacing, negative values condense it. |None| indicates the value is inherited
        from the style hierarchy.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.spacing_val

    @character_spacing.setter
    def character_spacing(self, value: int | Length | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.spacing_val = None if value is None else Emu(value)

    @property
    def bold(self) -> bool | None:
        """Read/write.

        Causes text in this font to appear in bold.
        """
        return self._get_bool_prop("b")

    @bold.setter
    def bold(self, value: bool | None) -> None:
        self._set_bool_prop("b", value)

    @property
    def border_color(self) -> RGBColor | None:
        """Run-border color as an |RGBColor|, or |None| if not set.

        Read/write. Reads ``w:rPr/w:bdr/@w:color``. Returns |None| when the
        ``w:bdr`` element is absent, when the attribute is missing, or when it
        is set to ``"auto"``. Assigning an |RGBColor| creates the ``w:bdr``
        child if necessary. Assigning |None| clears the attribute but leaves
        any sibling border attributes intact.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        bdr = rPr.bdr
        if bdr is None:
            return None
        color = bdr.color
        if color is None or not isinstance(color, RGBColor):
            return None
        return color

    @border_color.setter
    def border_color(self, value: RGBColor | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            bdr = rPr.bdr
            if bdr is None:
                return
            bdr.color = None
            return
        rPr = self._element.get_or_add_rPr()
        bdr = rPr.get_or_add_bdr()
        bdr.color = value

    @property
    def border_space(self) -> Length | None:
        """Space between the border and the text, in points.

        Read/write. Maps to ``w:rPr/w:bdr/@w:space``. Returns |None| when the
        ``w:bdr`` element or the attribute is absent. Assigning a |Length| or
        |Pt| value creates the ``w:bdr`` child if necessary. Assigning |None|
        clears the attribute.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        bdr = rPr.bdr
        if bdr is None:
            return None
        return bdr.space

    @border_space.setter
    def border_space(self, value: Length | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            bdr = rPr.bdr
            if bdr is None:
                return
            bdr.space = None
            return
        rPr = self._element.get_or_add_rPr()
        bdr = rPr.get_or_add_bdr()
        bdr.space = value

    @property
    def border_style(self) -> WD_BORDER_STYLE | None:
        """Border style as a member of :ref:`WdBorderStyle`, or |None| if not set.

        Read/write. Maps to ``w:rPr/w:bdr/@w:val``. Returns |None| when the
        ``w:bdr`` element or the attribute is absent. Assigning a
        |WD_BORDER_STYLE| member creates the ``w:bdr`` child if necessary.
        Assigning |None| clears the attribute.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        bdr = rPr.bdr
        if bdr is None:
            return None
        return bdr.val

    @border_style.setter
    def border_style(self, value: WD_BORDER_STYLE | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            bdr = rPr.bdr
            if bdr is None:
                return
            bdr.val = None
            return
        rPr = self._element.get_or_add_rPr()
        bdr = rPr.get_or_add_bdr()
        bdr.val = value

    @property
    def border_width(self) -> Length | None:
        """Border width as a |Length| value, or |None| if not set.

        Read/write. Maps to ``w:rPr/w:bdr/@w:sz`` which is measured in
        eighth-points. Returns |None| when the ``w:bdr`` element or the
        attribute is absent. Assigning a |Length| or |Pt| value creates the
        ``w:bdr`` child if necessary. Assigning |None| clears the attribute.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        bdr = rPr.bdr
        if bdr is None:
            return None
        return bdr.sz

    @border_width.setter
    def border_width(self, value: Length | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            bdr = rPr.bdr
            if bdr is None:
                return
            bdr.sz = None
            return
        rPr = self._element.get_or_add_rPr()
        bdr = rPr.get_or_add_bdr()
        bdr.sz = value

    def remove_border(self) -> None:
        """Remove the entire ``w:rPr/w:bdr`` child element, if present.

        Clears all run-border state in a single call. Has no effect when no
        ``w:bdr`` element is present.
        """
        rPr = self._element.rPr
        if rPr is None:
            return
        rPr._remove_bdr()  # pyright: ignore[reportPrivateUsage]

    @property
    def color(self):
        """A |ColorFormat| object providing a way to get and set the text color for this
        font."""
        return ColorFormat(self._element)

    @property
    def complex_script(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the characters in the run to be treated as complex script
        regardless of their Unicode values.
        """
        return self._get_bool_prop("cs")

    @complex_script.setter
    def complex_script(self, value: bool | None) -> None:
        self._set_bool_prop("cs", value)

    @property
    def cs_bold(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the complex script characters in the run to be displayed in
        bold typeface.
        """
        return self._get_bool_prop("bCs")

    @cs_bold.setter
    def cs_bold(self, value: bool | None) -> None:
        self._set_bool_prop("bCs", value)

    @property
    def cs_italic(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the complex script characters in the run to be displayed in
        italic typeface.
        """
        return self._get_bool_prop("iCs")

    @cs_italic.setter
    def cs_italic(self, value: bool | None) -> None:
        self._set_bool_prop("iCs", value)

    @property
    def double_strike(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text in the run to appear with double strikethrough.
        """
        return self._get_bool_prop("dstrike")

    @double_strike.setter
    def double_strike(self, value: bool | None) -> None:
        self._set_bool_prop("dstrike", value)

    @property
    def emboss(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text in the run to appear as if raised off the page in
        relief.
        """
        return self._get_bool_prop("emboss")

    @emboss.setter
    def emboss(self, value: bool | None) -> None:
        self._set_bool_prop("emboss", value)

    @property
    def hidden(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text in the run to be hidden from display, unless
        applications settings force hidden text to be shown.
        """
        return self._get_bool_prop("vanish")

    @hidden.setter
    def hidden(self, value: bool | None) -> None:
        self._set_bool_prop("vanish", value)

    @property
    def highlight_color(self) -> WD_COLOR_INDEX | None:
        """Color of highlighing applied or |None| if not highlighted."""
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.highlight_val

    @highlight_color.setter
    def highlight_color(self, value: WD_COLOR_INDEX | None):
        rPr = self._element.get_or_add_rPr()
        rPr.highlight_val = value

    @property
    def italic(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text of the run to appear in italics. |None| indicates
        the effective value is inherited from the style hierarchy.
        """
        return self._get_bool_prop("i")

    @italic.setter
    def italic(self, value: bool | None) -> None:
        self._set_bool_prop("i", value)

    @property
    def kerning(self) -> Length | None:
        """Read/write.

        |Length| value specifying the minimum font size for which kerning is automatically
        adjusted. |None| indicates kerning is not specified (inherited from style
        hierarchy).
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.kern_val

    @kerning.setter
    def kerning(self, value: int | Length | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.kern_val = None if value is None else Emu(value)

    @property
    def imprint(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text in the run to appear as if pressed into the page.
        """
        return self._get_bool_prop("imprint")

    @imprint.setter
    def imprint(self, value: bool | None) -> None:
        self._set_bool_prop("imprint", value)

    @property
    def math(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, specifies this run contains WML that should be handled as though it
        was Office Open XML Math.
        """
        return self._get_bool_prop("oMath")

    @math.setter
    def math(self, value: bool | None) -> None:
        self._set_bool_prop("oMath", value)

    @property
    def name_cs(self) -> str | None:
        """The Complex Script typeface name for this |Font|.

        Causes Complex Script text it controls to appear in the named font. |None|
        indicates the typeface is inherited from the style hierarchy.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.rFonts_cs

    @name_cs.setter
    def name_cs(self, value: str | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.rFonts_cs = value

    @property
    def name_east_asia(self) -> str | None:
        """The East Asian typeface name for this |Font|.

        Causes East Asian text it controls to appear in the named font. |None| indicates
        the typeface is inherited from the style hierarchy. Alias for `name_far_east`.
        """
        return self.name_far_east

    @name_east_asia.setter
    def name_east_asia(self, value: str | None) -> None:
        self.name_far_east = value

    @property
    def name_far_east(self) -> str | None:
        """The East Asian typeface name for this |Font|.

        Causes East Asian (CJK) text it controls to appear in the named font. |None|
        indicates the typeface is inherited from the style hierarchy.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.rFonts_eastAsia

    @name_far_east.setter
    def name_far_east(self, value: str | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.rFonts_eastAsia = value

    @property
    def name(self) -> str | None:
        """The typeface name for this |Font|.

        Causes the text it controls to appear in the named font, if a matching font is
        found. |None| indicates the typeface is inherited from the style hierarchy.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.rFonts_ascii

    @name.setter
    def name(self, value: str | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.rFonts_ascii = value
        rPr.rFonts_hAnsi = value

    @property
    def no_proof(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, specifies that the contents of this run should not report any
        errors when the document is scanned for spelling and grammar.
        """
        return self._get_bool_prop("noProof")

    @no_proof.setter
    def no_proof(self, value: bool | None) -> None:
        self._set_bool_prop("noProof", value)

    @property
    def outline(self) -> bool | None:
        """Read/write tri-state value.

        When |True| causes the characters in the run to appear as if they have an
        outline, by drawing a one pixel wide border around the inside and outside
        borders of each character glyph.
        """
        return self._get_bool_prop("outline")

    @outline.setter
    def outline(self, value: bool | None) -> None:
        self._set_bool_prop("outline", value)

    @property
    def rtl(self) -> bool | None:
        """Read/write tri-state value.

        When |True| causes the text in the run to have right-to-left characteristics.
        """
        return self._get_bool_prop("rtl")

    @rtl.setter
    def rtl(self, value: bool | None) -> None:
        self._set_bool_prop("rtl", value)

    @property
    def shading_color(self) -> RGBColor | None:
        """Run-level background (shading) color as an |RGBColor|, or |None| if not set.

        Read/write. Reads the ``w:fill`` attribute of ``w:rPr/w:shd``. Returns |None|
        when ``w:shd`` is absent or its ``w:fill`` is missing or set to ``"auto"``.

        Assigning an |RGBColor| writes ``w:rPr/w:shd`` with ``w:val="clear"`` and
        ``w:fill="RRGGBB"``. Assigning |None| removes the ``w:shd`` child. Distinct
        from :attr:`highlight_color`, which is a predefined palette applied as
        ``w:highlight``.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        shd = rPr.shd
        if shd is None:
            return None
        fill = shd.fill
        if fill is None or not isinstance(fill, RGBColor):
            return None
        return fill

    @shading_color.setter
    def shading_color(self, value: RGBColor | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            rPr._remove_shd()  # pyright: ignore[reportPrivateUsage]
            return
        rPr = self._element.get_or_add_rPr()
        shd = rPr.get_or_add_shd()
        shd.val = WD_SHADING_PATTERN.CLEAR
        shd.fill = value

    @property
    def shadow(self) -> bool | None:
        """Read/write tri-state value.

        When |True| causes the text in the run to appear as if each character has a
        shadow.
        """
        return self._get_bool_prop("shadow")

    @shadow.setter
    def shadow(self, value: bool | None) -> None:
        self._set_bool_prop("shadow", value)

    @property
    def size(self) -> Length | None:
        """Font height in English Metric Units (EMU).

        |None| indicates the font size should be inherited from the style hierarchy.
        |Length| is a subclass of |int| having properties for convenient conversion into
        points or other length units. The :class:`docx.shared.Pt` class allows
        convenient specification of point values::

            >>> font.size = Pt(24)
            >>> font.size
            304800
            >>> font.size.pt
            24.0

        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.sz_val

    @size.setter
    def size(self, emu: int | Length | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.sz_val = None if emu is None else Emu(emu)

    @property
    def small_caps(self) -> bool | None:
        """Read/write tri-state value.

        When |True| causes the lowercase characters in the run to appear as capital
        letters two points smaller than the font size specified for the run.
        """
        return self._get_bool_prop("smallCaps")

    @small_caps.setter
    def small_caps(self, value: bool | None) -> None:
        self._set_bool_prop("smallCaps", value)

    @property
    def snap_to_grid(self) -> bool | None:
        """Read/write tri-state value.

        When |True| causes the run to use the document grid characters per line settings
        defined in the docGrid element when laying out the characters in this run.
        """
        return self._get_bool_prop("snapToGrid")

    @snap_to_grid.setter
    def snap_to_grid(self, value: bool | None) -> None:
        self._set_bool_prop("snapToGrid", value)

    @property
    def spec_vanish(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, specifies that the given run shall always behave as if it is
        hidden, even when hidden text is being displayed in the current document. The
        property has a very narrow, specialized use related to the table of contents.
        Consult the spec (§17.3.2.36) for more details.
        """
        return self._get_bool_prop("specVanish")

    @spec_vanish.setter
    def spec_vanish(self, value: bool | None) -> None:
        self._set_bool_prop("specVanish", value)

    @property
    def strike(self) -> bool | None:
        """Read/write tri-state value.

        When |True| causes the text in the run to appear with a single horizontal line
        through the center of the line.
        """
        return self._get_bool_prop("strike")

    @strike.setter
    def strike(self, value: bool | None) -> None:
        self._set_bool_prop("strike", value)

    @property
    def subscript(self) -> bool | None:
        """Boolean indicating whether the characters in this |Font| appear as subscript.

        |None| indicates the subscript/subscript value is inherited from the style
        hierarchy.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.subscript

    @subscript.setter
    def subscript(self, value: bool | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.subscript = value

    @property
    def superscript(self) -> bool | None:
        """Boolean indicating whether the characters in this |Font| appear as
        superscript.

        |None| indicates the subscript/superscript value is inherited from the style
        hierarchy.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.superscript

    @superscript.setter
    def superscript(self, value: bool | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.superscript = value

    @property
    def underline(self) -> bool | WD_UNDERLINE | None:
        """The underline style for this |Font|.

        The value is one of |None|, |True|, |False|, or a member of :ref:`WdUnderline`.

        |None| indicates the font inherits its underline value from the style hierarchy.
        |False| indicates no underline. |True| indicates single underline. The values
        from :ref:`WdUnderline` are used to specify other outline styles such as double,
        wavy, and dotted.
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        val = rPr.u_val
        return (
            None
            if val == WD_UNDERLINE.INHERITED
            else True
            if val == WD_UNDERLINE.SINGLE
            else False
            if val == WD_UNDERLINE.NONE
            else val
        )

    @underline.setter
    def underline(self, value: bool | WD_UNDERLINE | None) -> None:
        rPr = self._element.get_or_add_rPr()
        # -- works fine without these two mappings, but only because True == 1 and
        # -- False == 0, which happen to match the mapping for WD_UNDERLINE.SINGLE
        # -- and .NONE respectively.
        val = (
            WD_UNDERLINE.SINGLE if value is True else WD_UNDERLINE.NONE if value is False else value
        )
        rPr.u_val = val

    @property
    def web_hidden(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, specifies that the contents of this run shall be hidden when the
        document is displayed in web page view.
        """
        return self._get_bool_prop("webHidden")

    @web_hidden.setter
    def web_hidden(self, value: bool | None) -> None:
        self._set_bool_prop("webHidden", value)

    def _get_bool_prop(self, name: str) -> bool | None:
        """Return the value of boolean child of `w:rPr` having `name`."""
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr._get_bool_val(name)  # pyright: ignore[reportPrivateUsage]

    def _set_bool_prop(self, name: str, value: bool | None):
        """Assign `value` to the boolean child `name` of `w:rPr`."""
        rPr = self._element.get_or_add_rPr()
        rPr._set_bool_val(name, value)  # pyright: ignore[reportPrivateUsage]
