"""Custom element classes related to the ``word/theme/theme1.xml`` part.

The theme part carries the document-wide DrawingML theme: a color scheme
(``a:clrScheme``), a font scheme (``a:fontScheme``), and a format scheme
(``a:fmtScheme``). This module models just the pieces required to resolve
``w:themeColor``/``w:themeFont`` references on documents opened via
python-docx.

Only the subset exposed through :class:`docx.theme.Theme` is modelled. In
particular the ``a:fmtScheme`` children (line styles, fills, effects) are
not broken out because the read-only API does not surface them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.shared import RGBColor


class CT_SRgbColor(BaseOxmlElement):
    """``<a:srgbClr>`` element — a direct sRGB color value.

    The ``val`` attribute is a six-hex-digit RGB value (e.g. ``"5B9BD5"``).
    """

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "val", ST_String
    )


class CT_SysColor(BaseOxmlElement):
    """``<a:sysClr>`` element — a named system color.

    ``val`` is a system-color token (e.g. ``"windowText"``). ``lastClr``
    records the resolved RGB value at the time the document was saved and
    serves as the fallback when no system-color resolver is available —
    which is always the case for python-docx.
    """

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "val", ST_String
    )
    lastClr: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lastClr", ST_String
    )


class CT_ColorChoice(BaseOxmlElement):
    """Parent for a single color choice (one of ``a:dk1``, ``a:accent1``, etc.).

    Each scheme slot contains at most one color child; the two we model are
    ``a:srgbClr`` and ``a:sysClr``. Other color-choice flavors (``a:scrgbClr``,
    ``a:hslClr``, ``a:prstClr``, ``a:schemeClr``) are rare in practice and not
    modelled.
    """

    srgbClr: CT_SRgbColor | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:srgbClr", successors=()
    )
    sysClr: CT_SysColor | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:sysClr", successors=()
    )

    @property
    def rgb(self) -> RGBColor | None:
        """Resolved |RGBColor| for this color-scheme slot, or |None| if unresolved.

        Prefers the ``a:srgbClr/@val`` direct RGB value. Falls back to
        ``a:sysClr/@lastClr`` — the resolved RGB recorded at save time — when
        only a system color is present. Returns |None| when neither a supported
        color child is present nor a resolved ``lastClr`` is available.
        """
        from docx.shared import RGBColor

        srgb = self.srgbClr
        if srgb is not None:
            return RGBColor.from_string(srgb.val)
        sys = self.sysClr
        if sys is not None and sys.lastClr is not None:
            return RGBColor.from_string(sys.lastClr)
        return None


class CT_ClrScheme(BaseOxmlElement):
    """``<a:clrScheme>`` element — the color scheme of a theme.

    Contains the twelve named color slots — two dark, two light, six
    accents, and two hyperlink colors — in the order prescribed by the
    DrawingML schema.
    """

    _tag_seq = (
        "a:dk1",
        "a:lt1",
        "a:dk2",
        "a:lt2",
        "a:accent1",
        "a:accent2",
        "a:accent3",
        "a:accent4",
        "a:accent5",
        "a:accent6",
        "a:hlink",
        "a:folHlink",
    )

    dk1: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:dk1", successors=_tag_seq[1:]
    )
    lt1: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:lt1", successors=_tag_seq[2:]
    )
    dk2: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:dk2", successors=_tag_seq[3:]
    )
    lt2: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:lt2", successors=_tag_seq[4:]
    )
    accent1: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:accent1", successors=_tag_seq[5:]
    )
    accent2: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:accent2", successors=_tag_seq[6:]
    )
    accent3: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:accent3", successors=_tag_seq[7:]
    )
    accent4: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:accent4", successors=_tag_seq[8:]
    )
    accent5: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:accent5", successors=_tag_seq[9:]
    )
    accent6: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:accent6", successors=_tag_seq[10:]
    )
    hlink: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:hlink", successors=_tag_seq[11:]
    )
    folHlink: CT_ColorChoice | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:folHlink", successors=()
    )
    del _tag_seq

    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", ST_String
    )

    def color_for(self, name: str) -> CT_ColorChoice | None:
        """Return the color-choice child named `name`, or |None|.

        `name` is a theme-color token such as ``"accent1"``, ``"dk1"``,
        ``"lt2"``, ``"hlink"``, or ``"folHlink"`` — the local part of the
        ``a:*`` child tag. Unknown names return |None|.
        """
        return getattr(self, name, None) if name in _VALID_COLOR_NAMES else None


_VALID_COLOR_NAMES = frozenset(
    (
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    )
)


class CT_TextFont(BaseOxmlElement):
    """``<a:latin>`` / ``<a:ea>`` / ``<a:cs>`` element — a theme typeface slot.

    Carries a ``typeface`` attribute naming the font. The attribute is
    required by the schema but commonly carries the empty string to mean
    "inherit" for non-primary scripts.
    """

    typeface: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "typeface", ST_String
    )


class CT_FontCollection(BaseOxmlElement):
    """``<a:majorFont>`` / ``<a:minorFont>`` element within the font scheme.

    Holds the Latin, East-Asian, and complex-script typeface slots used by
    the ``major*`` / ``minor*`` ``w:themeFont`` references. The ``a:font``
    script-specific children are not modelled because the read-only API
    only exposes the three primary slots.
    """

    latin: CT_TextFont | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:latin", successors=("a:ea", "a:cs", "a:font")
    )
    ea: CT_TextFont | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:ea", successors=("a:cs", "a:font")
    )
    cs: CT_TextFont | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:cs", successors=("a:font",)
    )


class CT_FontScheme(BaseOxmlElement):
    """``<a:fontScheme>`` element — the font scheme of a theme."""

    majorFont: CT_FontCollection | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:majorFont", successors=("a:minorFont", "a:extLst")
    )
    minorFont: CT_FontCollection | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:minorFont", successors=("a:extLst",)
    )

    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", ST_String
    )


class CT_ThemeElements(BaseOxmlElement):
    """``<a:themeElements>`` — container for the three scheme children."""

    clrScheme: CT_ClrScheme | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:clrScheme", successors=("a:fontScheme", "a:fmtScheme", "a:extLst")
    )
    fontScheme: CT_FontScheme | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:fontScheme", successors=("a:fmtScheme", "a:extLst")
    )


class CT_Theme(BaseOxmlElement):
    """``<a:theme>`` element, root of the ``word/theme/theme1.xml`` part."""

    themeElements: CT_ThemeElements | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:themeElements",
        successors=(
            "a:objectDefaults",
            "a:extraClrSchemeLst",
            "a:custClrLst",
            "a:extLst",
        ),
    )

    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", ST_String
    )

    @property
    def clrScheme(self) -> CT_ClrScheme | None:
        """The ``a:clrScheme`` element, or |None| if the theme lacks one."""
        elements = self.themeElements
        if elements is None:
            return None
        return elements.clrScheme

    @property
    def fontScheme(self) -> CT_FontScheme | None:
        """The ``a:fontScheme`` element, or |None| if the theme lacks one."""
        elements = self.themeElements
        if elements is None:
            return None
        return elements.fontScheme
