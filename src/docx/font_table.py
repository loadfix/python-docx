"""Proxy objects for the document's ``word/fontTable.xml`` part.

The font table lists every font referenced by the document together with
descriptive metadata (family classification, charset, PANOSE, etc.). Read
access is always available via :attr:`docx.document.Document.font_table` (or
|None| if the document has no ``fontTable`` part). Write access includes:

* :meth:`FontTable.add_embedded_font` — embed a font from a file path
  using the unobfuscated ``application/x-fontdata`` content-type (kept
  for back-compat with 2026.05.0).
* :meth:`FontTable.embed_font` — embed a font from raw TrueType bytes
  using Word's obfuscated-font format (ECMA-376 Part 1 §17.8).
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Literal, Union
from collections.abc import Iterator

from docx.font_obfuscation import (
    deobfuscate_font_bytes,
    generate_font_key,
    obfuscate_font_bytes,
)
from docx.opc.constants import RELATIONSHIP_TYPE as RT

if TYPE_CHECKING:
    from docx.oxml.font_table import CT_Font, CT_Fonts
    from docx.parts.font_table import FontTablePart


EmbedVariant = Literal["regular", "bold", "italic", "bold_italic"]

_EMBED_TAG = {
    "regular": "embedRegular",
    "bold": "embedBold",
    "italic": "embedItalic",
    "bold_italic": "embedBoldItalic",
}

# -- mapping from variant name to FontMetadata attribute returning the
# -- deobfuscated font bytes, used by ``embed_font`` for documentation and
# -- elsewhere for uniform dispatch. --
_VARIANT_FIELD = {
    "regular": "embedRegular",
    "bold": "embedBold",
    "italic": "embedItalic",
    "bold_italic": "embedBoldItalic",
}


class FontTable:
    """Read-only collection of :class:`FontMetadata` entries for a document.

    Supports iteration, ``len()``, membership testing (``"Arial" in font_table``),
    indexing by font name (``font_table["Arial"]``), and safe lookup
    (``font_table.get("Arial")``). Iteration order matches the XML order of
    the ``w:font`` children.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, fonts_elm: "CT_Fonts", part: "FontTablePart"):
        self._fonts = fonts_elm
        self._part = part

    def __iter__(self) -> Iterator["FontMetadata"]:
        return (
            FontMetadata(font_elm, self._part) for font_elm in self._fonts.font_lst
        )

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
        return FontMetadata(font_elm, self._part)

    def get(self, name: str) -> "FontMetadata | None":
        """Return the :class:`FontMetadata` for `name`, or |None| if not present.

        .. versionadded:: 2026.05.0
        """
        font_elm = self._fonts.get_font_by_name(name)
        if font_elm is None:
            return None
        return FontMetadata(font_elm, self._part)

    @property
    def element(self) -> "CT_Fonts":
        return self._fonts

    @property
    def part(self) -> "FontTablePart":
        return self._part

    @property
    def fonts(self) -> dict[str, "FontMetadata"]:
        """A ``{name: FontMetadata}`` snapshot of every ``<w:font>`` entry.

        The mapping is a freshly-constructed dict (mutating it does not
        affect the XML); the values are live :class:`FontMetadata`
        wrappers whose setters still update the underlying element.

        .. versionadded:: 2026.05.10
        """
        return {
            f.name: FontMetadata(f, self._part) for f in self._fonts.font_lst
        }

    def embed_font(
        self,
        name: str,
        regular: Union[bytes, bytearray, None] = None,
        bold: Union[bytes, bytearray, None] = None,
        italic: Union[bytes, bytearray, None] = None,
        bold_italic: Union[bytes, bytearray, None] = None,
    ) -> "FontMetadata":
        """Embed obfuscated TrueType bytes for one or more variants of `name`.

        Each ``*_ttf_bytes`` argument, when supplied, is obfuscated per
        ECMA-376 Part 1 §17.8 with a fresh GUID, stored as a package
        part with content-type
        ``application/vnd.openxmlformats-officedocument.obfuscatedFont``,
        and referenced from the matching ``<w:embedRegular>``/
        ``<w:embedBold>``/``<w:embedItalic>``/``<w:embedBoldItalic>``
        child of the ``<w:font>`` entry. If an entry with the given
        `name` already exists it is updated in place; otherwise a new
        entry is appended.

        Returns the :class:`FontMetadata` for the affected entry. At
        least one of the four variants must be supplied — passing only a
        ``name`` is a no-op the caller probably did not intend, so it
        raises :class:`ValueError`.

        .. versionadded:: 2026.05.10
        """
        variants: list[tuple[EmbedVariant, bytes]] = []
        for variant_name, blob in (
            ("regular", regular),
            ("bold", bold),
            ("italic", italic),
            ("bold_italic", bold_italic),
        ):
            if blob is None:
                continue
            variants.append((variant_name, bytes(blob)))

        if not variants:
            raise ValueError(
                "embed_font() requires at least one of regular/bold/italic/"
                "bold_italic to be supplied"
            )

        font_elm = self._fonts.get_font_by_name(name)
        if font_elm is None:
            font_elm = self._fonts.add_font()
            font_elm.name = name

        for variant, blob in variants:
            font_key = generate_font_key()
            obfuscated = obfuscate_font_bytes(blob, font_key)
            font_part = self._part.add_obfuscated_font_part(obfuscated)
            rId = self._part.relate_to(font_part, RT.FONT)

            tag = _EMBED_TAG[variant]
            getattr(font_elm, f"_remove_{tag}")()
            embed = getattr(font_elm, f"_add_{tag}")()
            embed.rId = rId
            embed.fontKey = font_key

        return FontMetadata(font_elm, self._part)

    def add_embedded_font(
        self,
        path: str | Path,
        family: EmbedVariant = "regular",
        name: str | None = None,
    ) -> "FontMetadata":
        """Embed the font binary at `path` into the document's font table.

        A :class:`docx.parts.font_table.FontPart` is created to hold the raw
        binary payload and related to the font-table part via an ``r:font``
        relationship. A matching ``<w:font>`` entry is added (or the existing
        one updated) with a ``<w:embedRegular>``/``<w:embedBold>``/
        ``<w:embedItalic>``/``<w:embedBoldItalic>`` child pointing at the new
        part. `family` selects which weight/style this embedded file represents
        (default ``"regular"``); `name` overrides the displayed font name and
        defaults to the file stem.

        Closes upstream#1231, #1307.

        .. versionadded:: 2026.05.0
        """
        if family not in _EMBED_TAG:
            raise ValueError(
                f"family must be one of {sorted(_EMBED_TAG)}, got {family!r}"
            )

        font_path = Path(path)
        font_name = name if name is not None else font_path.stem

        font_elm = self._fonts.get_font_by_name(font_name)
        if font_elm is None:
            font_elm = self._fonts.add_font()
            font_elm.name = font_name

        font_part = self._part.add_font_part(font_path)
        rId = self._part.relate_to(font_part, RT.FONT)

        # -- set (or replace) the appropriate embed child with the new rId --
        tag = _EMBED_TAG[family]
        getattr(font_elm, f"_remove_{tag}")()
        embed = getattr(font_elm, f"_add_{tag}")()
        embed.rId = rId

        return FontMetadata(font_elm, self._part)


class FontMetadata:
    """Read-only view of a single ``<w:font>`` entry in the font table.

    The optional `part` argument attaches the owning
    :class:`docx.parts.font_table.FontTablePart` so properties like
    :attr:`embedded_regular` can resolve the ``r:id`` on each
    ``<w:embed*>`` child to its target :class:`FontPart` and return the
    deobfuscated TrueType bytes.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        font_elm: "CT_Font",
        part: "FontTablePart | None" = None,
    ):
        self._font = font_elm
        self._part = part

    @property
    def name(self) -> str:
        """The font name (``w:font/@w:name``), e.g. ``"Arial"``.

        .. versionadded:: 2026.05.0
        """
        return self._font.name

    @property
    def family(self) -> str | None:
        """The font-family classification (``w:family/@w:val``) or |None|.

        Common values: ``"swiss"``, ``"roman"``, ``"modern"``, ``"script"``,
        ``"decorative"``, ``"auto"``.

        .. versionadded:: 2026.05.0
        """
        family = self._font.family
        if family is None:
            return None
        return family.val

    @property
    def charset(self) -> str | None:
        """The charset (``w:charset/@w:val``), typically a two-character hex string.

        .. versionadded:: 2026.05.0
        """
        charset = self._font.charset
        if charset is None:
            return None
        return charset.val

    @property
    def pitch(self) -> str | None:
        """The pitch classification (``w:pitch/@w:val``) or |None|.

        Common values: ``"fixed"``, ``"variable"``, ``"default"``.

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
        """
        panose1 = self._font.panose1
        if panose1 is None:
            return None
        return panose1.val

    @property
    def alt_name(self) -> str | None:
        """The alternate font name (``w:altName/@w:val``) or |None|.

        Word falls back to this name when the primary font is not available.

        .. versionadded:: 2026.05.0
        """
        altName = self._font.altName
        if altName is None:
            return None
        return altName.val

    @property
    def embed_regular(self) -> bool:
        """True if a ``<w:embedRegular>`` element is present on this font entry.

        .. versionadded:: 2026.05.0
        """
        return self._font.embedRegular is not None

    @property
    def embed_bold(self) -> bool:
        """True if a ``<w:embedBold>`` element is present on this font entry.

        .. versionadded:: 2026.05.0
        """
        return self._font.embedBold is not None

    @property
    def embed_italic(self) -> bool:
        """True if a ``<w:embedItalic>`` element is present on this font entry.

        .. versionadded:: 2026.05.0
        """
        return self._font.embedItalic is not None

    @property
    def embed_bold_italic(self) -> bool:
        """True if a ``<w:embedBoldItalic>`` element is present on this font entry.

        .. versionadded:: 2026.05.0
        """
        return self._font.embedBoldItalic is not None

    @property
    def embedded_regular(self) -> bytes | None:
        """Deobfuscated ``regular`` TTF bytes, or |None| if not embedded.

        Returns |None| when no ``<w:embedRegular>`` child is present, or
        when the owning :class:`FontTablePart` was not attached (for
        cxml-based unit fixtures). Callers that need the original font
        file should save it with ``Path(...).write_bytes(metadata.embedded_regular)``.

        .. versionadded:: 2026.05.10
        """
        return self._embedded_bytes("embedRegular")

    @property
    def embedded_bold(self) -> bytes | None:
        """Deobfuscated ``bold`` TTF bytes, or |None| if not embedded.

        .. versionadded:: 2026.05.10
        """
        return self._embedded_bytes("embedBold")

    @property
    def embedded_italic(self) -> bytes | None:
        """Deobfuscated ``italic`` TTF bytes, or |None| if not embedded.

        .. versionadded:: 2026.05.10
        """
        return self._embedded_bytes("embedItalic")

    @property
    def embedded_bold_italic(self) -> bytes | None:
        """Deobfuscated ``bold-italic`` TTF bytes, or |None| if not embedded.

        .. versionadded:: 2026.05.10
        """
        return self._embedded_bytes("embedBoldItalic")

    def _embedded_bytes(self, tag: str) -> bytes | None:
        """Return deobfuscated TTF bytes for `tag` or |None| if unavailable.

        `tag` is the local name of the embed element (e.g.
        ``"embedRegular"``). Returns |None| if the child is absent, if
        the part is not attached, or if the referenced part does not
        exist. Unobfuscated (``application/x-fontdata`` /
        ``application/x-font-ttf``) parts are returned as-is; obfuscated
        parts are XOR-deobfuscated using the ``w:fontKey`` GUID on the
        embed element.
        """
        embed = getattr(self._font, tag)
        if embed is None or self._part is None:
            return None
        rId = embed.rId
        if rId is None:
            return None
        try:
            related = self._part.related_parts[rId]
        except KeyError:
            return None
        blob = related.blob
        font_key = embed.fontKey
        if font_key is None:
            # -- un-obfuscated part (e.g. the 2026.05.0 add_embedded_font
            # -- path); no key means the bytes are stored raw. --
            return blob
        return deobfuscate_font_bytes(blob, font_key)

    @property
    def element(self) -> "CT_Font":
        return self._font
