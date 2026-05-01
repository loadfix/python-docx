"""Proxy object for a `w:sym` (special-character-from-font) element."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.oxml.text.run import CT_Sym


class Symbol:
    """A special character whose glyph is drawn from a named font.

    Wraps a `w:sym` element inside a run. Word uses this element to represent
    characters whose glyph is taken from a font like "Wingdings" where the
    glyph at a given code point isn't the normal Unicode character at that
    code point.
    """

    def __init__(self, sym: CT_Sym):
        self._sym = sym
        self._element = sym

    @property
    def char_code(self) -> int:
        """Integer Unicode code point of the symbol within ``font``.

        The ``w:char`` attribute stores this value as a hex string in the XML
        (e.g. ``"F0E0"``); this property returns it as an ``int``.
        """
        return int(self._sym.char, 16)

    @property
    def char_hex(self) -> str:
        """The 4-character uppercase hex string representation of the code point.

        This is the form used by Word to serialize the ``w:char`` attribute.
        """
        # -- normalise whatever the XML actually stored so the value returned
        # -- is always 4+ uppercase hex digits, padded to at least 4 chars --
        return format(self.char_code, "04X")

    @property
    def font(self) -> str:
        """The font the glyph is rendered from, e.g. ``"Wingdings"``."""
        return self._sym.font

    def delete(self) -> None:
        """Remove this symbol element from its parent run.

        After calling this method, this |Symbol| object is "defunct" and
        should not be used further.
        """
        sym = self._sym
        parent = sym.getparent()
        if parent is None:
            return
        parent.remove(sym)
