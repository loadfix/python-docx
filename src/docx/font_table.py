"""Read-only proxy objects for the document's ``word/fontTable.xml`` part.

The font table lists every font referenced by the document together with
descriptive metadata (family classification, charset, PANOSE, etc.). This is
read-only from the python-docx perspective — document authors don't create or
edit these entries; Word generates them when saving.

Use :attr:`docx.document.Document.font_table` to obtain a :class:`FontTable`
collection (or |None| if the document has no ``fontTable`` part).
"""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Iterator

if TYPE_CHECKING:
    from docx.oxml.font_table import CT_Font, CT_Fonts
    from docx.parts.font_table import FontTablePart


class FontTable:
    """Read-only collection of :class:`FontMetadata` entries for a document.

    Supports iteration, ``len()``, membership testing (``"Arial" in font_table``),
    indexing by font name (``font_table["Arial"]``), and safe lookup
    (``font_table.get("Arial")``). Iteration order matches the XML order of
    the ``w:font`` children.
    """

    def __init__(self, fonts_elm: "CT_Fonts", part: "FontTablePart"):
        self._fonts = fonts_elm
        self._part = part

    def __iter__(self) -> Iterator["FontMetadata"]:
        return (FontMetadata(font_elm) for font_elm in self._fonts.font_lst)

    def __len__(self) -> int:
        return len(self._fonts.font_lst)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        return self._fonts.get_font_by_name(name) is not None

    def __getitem__(self, name: str) -> "FontMetadata":
        font_elm = self._fonts.get_font_by_name(name)
        if font_elm is None:
            raise KeyError(name)
        return FontMetadata(font_elm)

    def get(self, name: str) -> "FontMetadata | None":
        """Return the :class:`FontMetadata` for `name`, or |None| if not present."""
        font_elm = self._fonts.get_font_by_name(name)
        if font_elm is None:
            return None
        return FontMetadata(font_elm)

    @property
    def element(self) -> "CT_Fonts":
        return self._fonts

    @property
    def part(self) -> "FontTablePart":
        return self._part


class FontMetadata:
    """Read-only view of a single ``<w:font>`` entry in the font table."""

    def __init__(self, font_elm: "CT_Font"):
        self._font = font_elm

    @property
    def name(self) -> str:
        """The font name (``w:font/@w:name``), e.g. ``"Arial"``."""
        return self._font.name

    @property
    def family(self) -> str | None:
        """The font-family classification (``w:family/@w:val``) or |None|.

        Common values: ``"swiss"``, ``"roman"``, ``"modern"``, ``"script"``,
        ``"decorative"``, ``"auto"``.
        """
        family = self._font.family
        if family is None:
            return None
        return family.val

    @property
    def charset(self) -> str | None:
        """The charset (``w:charset/@w:val``), typically a two-character hex string."""
        charset = self._font.charset
        if charset is None:
            return None
        return charset.val

    @property
    def pitch(self) -> str | None:
        """The pitch classification (``w:pitch/@w:val``) or |None|.

        Common values: ``"fixed"``, ``"variable"``, ``"default"``.
        """
        pitch = self._font.pitch
        if pitch is None:
            return None
        return pitch.val

    @property
    def panose(self) -> str | None:
        """The 10-byte PANOSE classification (``w:panose1/@w:val``) or |None|.

        Returned as the raw 20-character hex string as stored in XML, with no
        case-normalisation.
        """
        panose1 = self._font.panose1
        if panose1 is None:
            return None
        return panose1.val

    @property
    def alt_name(self) -> str | None:
        """The alternate font name (``w:altName/@w:val``) or |None|.

        Word falls back to this name when the primary font is not available.
        """
        altName = self._font.altName
        if altName is None:
            return None
        return altName.val

    @property
    def embed_regular(self) -> bool:
        """True if a ``<w:embedRegular>`` element is present on this font entry."""
        return self._font.embedRegular is not None

    @property
    def embed_bold(self) -> bool:
        """True if a ``<w:embedBold>`` element is present on this font entry."""
        return self._font.embedBold is not None

    @property
    def embed_italic(self) -> bool:
        """True if a ``<w:embedItalic>`` element is present on this font entry."""
        return self._font.embedItalic is not None

    @property
    def embed_bold_italic(self) -> bool:
        """True if a ``<w:embedBoldItalic>`` element is present on this font entry."""
        return self._font.embedBoldItalic is not None

    @property
    def element(self) -> "CT_Font":
        return self._font
