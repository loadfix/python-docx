"""|Theme| proxy for ``word/theme/theme1.xml``.

Provides read-only access to the document theme: color scheme, font
scheme, and theme name. Access via
:attr:`docx.document.Document.theme`, which returns a :class:`Theme`
instance when the document has a ``theme`` relationship, or |None|
otherwise.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.shared import ElementProxy

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.theme import (
        CT_ClrScheme,
        CT_FontCollection,
        CT_FontScheme,
        CT_Theme,
    )
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.shared import RGBColor


# -- OOXML ``w:themeColor`` tokens we accept as ``ThemeColors[...]`` keys. --
# The sequence mirrors ``CT_ClrScheme``'s child order; the ``*Reference``
# aliases exposed by ``ST_ThemeColor`` are not implemented here — callers
# resolve theme-color references by looking up the matching slot.
_COLOR_SLOTS = (
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


class Theme(ElementProxy):
    """Proxy for the ``a:theme`` root element of the theme part.

    Exposes the theme name and lazy accessors for the color scheme and
    font scheme. Read-only — python-docx does not support authoring
    themes.
    """

    def __init__(
        self,
        element: BaseOxmlElement,
        parent: t.ProvidesXmlPart | None = None,
    ):
        super().__init__(element, parent)
        self._theme = cast("CT_Theme", element)

    @property
    def name(self) -> str | None:
        """The value of ``a:theme/@name``, or |None| if the attribute is absent."""
        return self._theme.name

    @property
    def colors(self) -> ThemeColors:
        """A |ThemeColors| proxy for the ``a:clrScheme`` of this theme.

        The returned object exposes a no-op view when the theme has no
        color scheme (every slot returns |None|) — mirroring how Word
        falls back to the default theme for missing references.
        """
        return ThemeColors(self._theme.clrScheme)

    @property
    def fonts(self) -> ThemeFonts:
        """A |ThemeFonts| proxy for the ``a:fontScheme`` of this theme.

        The returned object returns |None| for every slot when the theme
        has no font scheme.
        """
        return ThemeFonts(self._theme.fontScheme)


class ThemeColors:
    """Read-only view over the twelve-slot theme color scheme.

    Slot accessors return an |RGBColor| resolved from either the
    ``a:srgbClr/@val`` attribute or, for ``a:sysClr``, the ``lastClr``
    fallback. A missing slot — or a scheme child with no resolvable color
    — yields |None|. Lookup by OOXML token is available via
    ``colors[name]`` (e.g. ``colors["accent1"]`` or ``colors["hlink"]``).
    """

    def __init__(self, clrScheme: CT_ClrScheme | None):
        self._clrScheme = clrScheme

    def __getitem__(self, name: str) -> RGBColor | None:
        """Return the |RGBColor| for theme-color token `name`, or |None|.

        `name` is an OOXML ``w:themeColor`` token — one of the twelve
        scheme slots: ``"dk1"``, ``"lt1"``, ``"dk2"``, ``"lt2"``,
        ``"accent1"``..``"accent6"``, ``"hlink"``, ``"folHlink"``.
        Returns |None| when the theme does not define the slot or when
        the slot's color cannot be resolved to RGB (e.g. an ``a:sysClr``
        with no ``lastClr`` fallback). Raises |KeyError| for unknown
        tokens so callers can distinguish "undefined slot" from "unknown
        name".
        """
        if name not in _COLOR_SLOTS:
            raise KeyError(name)
        if self._clrScheme is None:
            return None
        choice = self._clrScheme.color_for(name)
        if choice is None:
            return None
        return choice.rgb

    def _get(self, name: str) -> RGBColor | None:
        if self._clrScheme is None:
            return None
        choice = self._clrScheme.color_for(name)
        if choice is None:
            return None
        return choice.rgb

    @property
    def name(self) -> str | None:
        """The value of ``a:clrScheme/@name``, or |None| when absent."""
        if self._clrScheme is None:
            return None
        return self._clrScheme.name

    @property
    def dark_1(self) -> RGBColor | None:
        """The ``a:dk1`` color, or |None| when unresolved."""
        return self._get("dk1")

    @property
    def dark_2(self) -> RGBColor | None:
        """The ``a:dk2`` color, or |None| when unresolved."""
        return self._get("dk2")

    @property
    def light_1(self) -> RGBColor | None:
        """The ``a:lt1`` color, or |None| when unresolved."""
        return self._get("lt1")

    @property
    def light_2(self) -> RGBColor | None:
        """The ``a:lt2`` color, or |None| when unresolved."""
        return self._get("lt2")

    @property
    def accent_1(self) -> RGBColor | None:
        """The ``a:accent1`` color, or |None| when unresolved."""
        return self._get("accent1")

    @property
    def accent_2(self) -> RGBColor | None:
        """The ``a:accent2`` color, or |None| when unresolved."""
        return self._get("accent2")

    @property
    def accent_3(self) -> RGBColor | None:
        """The ``a:accent3`` color, or |None| when unresolved."""
        return self._get("accent3")

    @property
    def accent_4(self) -> RGBColor | None:
        """The ``a:accent4`` color, or |None| when unresolved."""
        return self._get("accent4")

    @property
    def accent_5(self) -> RGBColor | None:
        """The ``a:accent5`` color, or |None| when unresolved."""
        return self._get("accent5")

    @property
    def accent_6(self) -> RGBColor | None:
        """The ``a:accent6`` color, or |None| when unresolved."""
        return self._get("accent6")

    @property
    def hyperlink(self) -> RGBColor | None:
        """The ``a:hlink`` color, or |None| when unresolved."""
        return self._get("hlink")

    @property
    def followed_hyperlink(self) -> RGBColor | None:
        """The ``a:folHlink`` color, or |None| when unresolved."""
        return self._get("folHlink")


class ThemeFonts:
    """Read-only view over the theme's font scheme.

    Each property returns the typeface string from the matching
    ``a:latin``/``a:ea``/``a:cs`` child of ``a:majorFont`` or
    ``a:minorFont``, or |None| when the slot is missing.
    """

    def __init__(self, fontScheme: CT_FontScheme | None):
        self._fontScheme = fontScheme

    @property
    def name(self) -> str | None:
        """The value of ``a:fontScheme/@name``, or |None| when absent."""
        if self._fontScheme is None:
            return None
        return self._fontScheme.name

    @staticmethod
    def _typeface(collection: CT_FontCollection | None, slot: str) -> str | None:
        if collection is None:
            return None
        child = getattr(collection, slot)
        if child is None:
            return None
        return child.typeface

    @property
    def major_latin(self) -> str | None:
        """Typeface at ``a:majorFont/a:latin/@typeface``, or |None|."""
        if self._fontScheme is None:
            return None
        return self._typeface(self._fontScheme.majorFont, "latin")

    @property
    def minor_latin(self) -> str | None:
        """Typeface at ``a:minorFont/a:latin/@typeface``, or |None|."""
        if self._fontScheme is None:
            return None
        return self._typeface(self._fontScheme.minorFont, "latin")

    @property
    def major_east_asian(self) -> str | None:
        """Typeface at ``a:majorFont/a:ea/@typeface``, or |None|."""
        if self._fontScheme is None:
            return None
        return self._typeface(self._fontScheme.majorFont, "ea")

    @property
    def minor_east_asian(self) -> str | None:
        """Typeface at ``a:minorFont/a:ea/@typeface``, or |None|."""
        if self._fontScheme is None:
            return None
        return self._typeface(self._fontScheme.minorFont, "ea")

    @property
    def major_cs(self) -> str | None:
        """Typeface at ``a:majorFont/a:cs/@typeface``, or |None|."""
        if self._fontScheme is None:
            return None
        return self._typeface(self._fontScheme.majorFont, "cs")

    @property
    def minor_cs(self) -> str | None:
        """Typeface at ``a:minorFont/a:cs/@typeface``, or |None|."""
        if self._fontScheme is None:
            return None
        return self._typeface(self._fontScheme.minorFont, "cs")
