"""Font-related proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.dml.color import ColorFormat
from docx.enum.table import WD_SHADING_PATTERN
from docx.enum.text import WD_UNDERLINE
from docx.shared import ElementProxy, Emu, RGBColor

if TYPE_CHECKING:
    from docx.enum.text import WD_BORDER_STYLE, WD_COLOR_INDEX
    from docx.oxml.text.font import CT_EastAsianLayout
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

        .. versionadded:: 2026.05.0
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
        # Mirror to the complex-script bold toggle. Word emits both
        # <w:b/> and <w:bCs/> together; omitting <w:bCs/> silently drops
        # bold on Arabic/Hebrew/Thai runs when Word reopens the file.
        # Callers that need divergent values can still set cs_bold
        # explicitly after this setter.
        self._set_bool_prop("bCs", value)

    @property
    def border_color(self) -> RGBColor | None:
        """Run-border color as an |RGBColor|, or |None| if not set.

        Read/write. Reads ``w:rPr/w:bdr/@w:color``. Returns |None| when the
        ``w:bdr`` element is absent, when the attribute is missing, or when it
        is set to ``"auto"``. Assigning an |RGBColor| creates the ``w:bdr``
        child if necessary. Assigning |None| clears the attribute but leaves
        any sibling border attributes intact.

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return
        rPr._remove_bdr()  # pyright: ignore[reportPrivateUsage]

    @property
    def color(self):
        """A |ColorFormat| object providing a way to get and set the text color for this
        font.

        Read-only property returning a |ColorFormat|; assignments set the run's
        RGB color. Assigning an |RGBColor| is equivalent to
        ``font.color.rgb = value``. Assigning |None| clears any direct color
        (``w:rPr/w:color``).

        .. versionadded:: 2026.05.0
            Assignment shorthand for ``font.color.rgb = <value>``.
        """
        return ColorFormat(self._element)

    @color.setter
    def color(self, value: RGBColor | None) -> None:
        ColorFormat(self._element).rgb = value

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
        # Mirror to the complex-script italic toggle. See bold setter
        # for the rationale — Word drops italic on complex-script runs
        # if only <w:i/> is present.
        self._set_bool_prop("iCs", value)

    @property
    def kerning(self) -> Length | None:
        """Read/write.

        |Length| value specifying the minimum font size for which kerning is automatically
        adjusted. |None| indicates kerning is not specified (inherited from style
        hierarchy).

        .. versionadded:: 2026.05.0
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
    def language(self) -> str | None:
        """Primary (Latin-script) language tag for this run.

        BCP-47 language tag (e.g. ``"en-US"``) or |None| when unset. Maps to
        ``w:rPr/w:lang/@w:val``. Assigning a string creates the ``w:lang``
        child if necessary. Assigning |None| clears only the ``w:val``
        attribute — use :meth:`remove_language` to remove the entire
        ``w:lang`` element.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        lang = rPr.lang
        if lang is None:
            return None
        return lang.val

    @language.setter
    def language(self, value: str | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            lang = rPr.lang
            if lang is None:
                return
            lang.val = None
            return
        rPr = self._element.get_or_add_rPr()
        lang = rPr.get_or_add_lang()
        lang.val = value

    @property
    def east_asian_language(self) -> str | None:
        """East Asian language tag for this run.

        BCP-47 language tag or |None| when unset. Maps to
        ``w:rPr/w:lang/@w:eastAsia``. Assigning a string creates the
        ``w:lang`` child if necessary. Assigning |None| clears only the
        ``w:eastAsia`` attribute — use :meth:`remove_language` to remove the
        entire ``w:lang`` element.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        lang = rPr.lang
        if lang is None:
            return None
        return lang.eastAsia

    @east_asian_language.setter
    def east_asian_language(self, value: str | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            lang = rPr.lang
            if lang is None:
                return
            lang.eastAsia = None
            return
        rPr = self._element.get_or_add_rPr()
        lang = rPr.get_or_add_lang()
        lang.eastAsia = value

    @property
    def bidi_language(self) -> str | None:
        """Complex-script (bidirectional) language tag for this run.

        BCP-47 language tag or |None| when unset. Maps to
        ``w:rPr/w:lang/@w:bidi``. Assigning a string creates the ``w:lang``
        child if necessary. Assigning |None| clears only the ``w:bidi``
        attribute — use :meth:`remove_language` to remove the entire
        ``w:lang`` element.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        lang = rPr.lang
        if lang is None:
            return None
        return lang.bidi

    @bidi_language.setter
    def bidi_language(self, value: str | None) -> None:
        if value is None:
            rPr = self._element.rPr
            if rPr is None:
                return
            lang = rPr.lang
            if lang is None:
                return
            lang.bidi = None
            return
        rPr = self._element.get_or_add_rPr()
        lang = rPr.get_or_add_lang()
        lang.bidi = value

    def remove_language(self) -> None:
        """Remove the entire ``w:rPr/w:lang`` child element, if present.

        Clears all language-tag state in a single call. Has no effect when no
        ``w:lang`` element is present.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return
        rPr._remove_lang()  # pyright: ignore[reportPrivateUsage]

    @property
    def east_asian_layout(self) -> EastAsianLayout | None:
        """|EastAsianLayout| proxy for this run's ``w:eastAsianLayout``, or |None|.

        Returns |None| when the run has no ``w:rPr/w:eastAsianLayout`` child.
        Use :meth:`set_east_asian_layout` to create or update the element and
        :meth:`remove_east_asian_layout` to drop it entirely.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        eal = rPr.eastAsianLayout
        if eal is None:
            return None
        return EastAsianLayout(eal)

    def set_east_asian_layout(
        self,
        *,
        id: int | None = None,
        two_lines_in_one: bool | None = None,
        vertical_alignment: bool | None = None,
        compressed: bool | None = None,
    ) -> EastAsianLayout:
        """Create or update the ``w:eastAsianLayout`` element on this run.

        Any keyword argument left at its default of |None| is left unchanged
        when the element already exists. To clear an attribute, use the
        corresponding setter on the returned |EastAsianLayout| (e.g.
        ``layout.two_lines_in_one = None``) or call
        :meth:`remove_east_asian_layout` to drop the element entirely.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.get_or_add_rPr()
        eal = rPr.get_or_add_eastAsianLayout()
        layout = EastAsianLayout(eal)
        if id is not None:
            layout.id = id
        if two_lines_in_one is not None:
            layout.two_lines_in_one = two_lines_in_one
        if vertical_alignment is not None:
            layout.vertical_alignment = vertical_alignment
        if compressed is not None:
            layout.compressed = compressed
        return layout

    def remove_east_asian_layout(self) -> None:
        """Remove the ``w:rPr/w:eastAsianLayout`` element, if present.

        Has no effect when no ``w:rPr`` or no ``w:eastAsianLayout`` child is
        present.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return
        if rPr.eastAsianLayout is None:
            return
        rPr._remove_eastAsianLayout()  # pyright: ignore[reportPrivateUsage]

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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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
        # -- mirror the ascii / hAnsi name onto the complex-script slot so
        # -- bidi (RTL) runs use the same typeface (upstream #510, #430, #973).
        # -- Callers that want a different CS font can explicitly set
        # -- :attr:`Font.name_cs` afterwards. --
        rPr.rFonts_cs = value

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
    def right_to_left(self) -> bool:
        """|True| when the run is flagged for right-to-left (bidi) rendering.

        Maps to ``w:rPr/w:rtl``. Returns |False| when the element is absent.
        Assigning |True| inserts ``w:rtl``; assigning |False| or |None| removes
        it. When |True|, the run is rendered right-to-left using the
        complex-script (CS) font—appropriate for Arabic, Hebrew, or Farsi text.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return False
        rtl = rPr.rtl
        if rtl is None:
            return False
        return rtl.val

    @right_to_left.setter
    def right_to_left(self, value: bool | None) -> None:
        rPr = self._element.get_or_add_rPr()
        if value in (None, False):
            rPr._remove_rtl()  # pyright: ignore[reportPrivateUsage]
        else:
            rtl = rPr.get_or_add_rtl()
            rtl.val = True

    @property
    def shading_color(self) -> RGBColor | None:
        """Run-level background (shading) color as an |RGBColor|, or |None| if not set.

        Read/write. Reads the ``w:fill`` attribute of ``w:rPr/w:shd``. Returns |None|
        when ``w:shd`` is absent or its ``w:fill`` is missing or set to ``"auto"``.

        Assigning an |RGBColor| writes ``w:rPr/w:shd`` with ``w:val="clear"`` and
        ``w:fill="RRGGBB"``. Assigning |None| removes the ``w:shd`` child. Distinct
        from :attr:`highlight_color`, which is a predefined palette applied as
        ``w:highlight``.

        .. versionadded:: 2026.05.0
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
    def cs_size(self) -> Length | None:
        """Complex-script (RTL / bidi) font height in English Metric Units.

        Maps to ``w:rPr/w:szCs``. Returns |None| when ``w:szCs`` is absent
        (inherited from the style hierarchy). Assigning |None| removes the
        attribute.

        Word uses ``w:szCs`` for Arabic / Hebrew / Farsi glyph sizing and
        leaves them at the default when only ``w:sz`` is set. The main
        :attr:`size` setter also writes ``w:szCs`` for symmetry; use
        ``cs_size`` to override the complex-script size independently.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.szCs_val

    @cs_size.setter
    def cs_size(self, emu: int | Length | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.szCs_val = None if emu is None else Emu(emu)

    @property
    def character_scale(self) -> int | None:
        """Horizontal character-scale percentage (``w:rPr/w:w/@w:val``).

        Integer percent, e.g. ``100`` for normal width, ``200`` for double-
        width, ``50`` for half-width. Returns |None| when ``w:w`` is absent
        (inherited). Assigning |None| removes the element.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.w_val

    @character_scale.setter
    def character_scale(self, value: int | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.w_val = value

    @property
    def ligatures(self) -> str | None:
        """OpenType ligature style (``w:rPr/w14:ligatures/@w14:val``).

        String value such as ``"none"``, ``"standard"``,
        ``"standardContextual"``, ``"historical"``, ``"discretional"``,
        ``"all"``, or combinations like ``"standardContextualHistorical"``.
        Returns |None| when ``w14:ligatures`` is absent. Assigning |None|
        removes the element.

        .. versionadded:: 2026.05.0
        """
        rPr = self._element.rPr
        if rPr is None:
            return None
        return rPr.ligatures_val

    @ligatures.setter
    def ligatures(self, value: str | None) -> None:
        rPr = self._element.get_or_add_rPr()
        rPr.ligatures_val = value

    def copy_to(self, target: "Font") -> None:
        """Replace `target`'s ``w:rPr`` children with a deep copy of this ``w:rPr``.

        When this font has no ``w:rPr``, `target`'s ``w:rPr`` children (and
        attributes) are cleared but the element is preserved. When this font
        does have an ``w:rPr``, the target's ``w:rPr`` is ensured to exist
        and its contents are replaced — the target run's character
        formatting becomes identical to this run's.

        .. versionadded:: 2026.05.0
        """
        from copy import deepcopy

        source_rPr = self._element.rPr
        target_rPr = target._element.get_or_add_rPr()
        # -- clear target's existing children and attributes --
        for child in list(target_rPr):
            target_rPr.remove(child)
        for attr_name in list(target_rPr.attrib):
            del target_rPr.attrib[attr_name]
        if source_rPr is None:
            return
        # -- copy attributes --
        for attr_name, attr_value in source_rPr.attrib.items():
            target_rPr.set(attr_name, attr_value)
        # -- deep-copy children --
        for child in source_rPr:
            target_rPr.append(deepcopy(child))

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
        length = None if emu is None else Emu(emu)
        rPr.sz_val = length
        # -- also write ``w:szCs`` so complex-script / bidi (RTL) runs inherit
        # -- the same size. Word uses ``w:szCs`` for Arabic / Hebrew / Farsi
        # -- glyphs and leaves them at the default when only ``w:sz`` is set
        # -- (upstream #510, #430, #973). --
        rPr.szCs_val = length

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


class EastAsianLayout:
    """Proxy for a run-level ``w:eastAsianLayout`` element.

    Provides read/write access to the East Asian typography attributes
    (``@w:id``, ``@w:combine``, ``@w:vert``, ``@w:vertCompress``). Accessed
    via :attr:`Font.east_asian_layout`.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, eastAsianLayout: CT_EastAsianLayout):
        self._element = eastAsianLayout

    @property
    def id(self) -> int | None:
        """Unique identifier (``w:eastAsianLayout/@w:id``), or |None|.

        .. versionadded:: 2026.05.0
        """
        return self._element.id

    @id.setter
    def id(self, value: int | None) -> None:
        self._element.id = value

    @property
    def two_lines_in_one(self) -> bool | None:
        """|True| when two lines are rendered as one combined glyph.

        Maps to ``w:eastAsianLayout/@w:combine``. Returns |None| when the
        attribute is absent.

        .. versionadded:: 2026.05.0
        """
        return self._element.combine

    @two_lines_in_one.setter
    def two_lines_in_one(self, value: bool | None) -> None:
        self._element.combine = value

    @property
    def vertical_alignment(self) -> bool | None:
        """|True| when the run is laid out vertically.

        Maps to ``w:eastAsianLayout/@w:vert``. Returns |None| when the
        attribute is absent.

        .. versionadded:: 2026.05.0
        """
        return self._element.vert

    @vertical_alignment.setter
    def vertical_alignment(self, value: bool | None) -> None:
        self._element.vert = value

    @property
    def compressed(self) -> bool | None:
        """|True| when vertical text is compressed.

        Maps to ``w:eastAsianLayout/@w:vertCompress``. Returns |None| when
        the attribute is absent.

        .. versionadded:: 2026.05.0
        """
        return self._element.vertCompress

    @compressed.setter
    def compressed(self, value: bool | None) -> None:
        self._element.vertCompress = value
