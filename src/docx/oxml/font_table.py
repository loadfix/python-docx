"""Custom element classes related to the font table part.

The font table part (``word/fontTable.xml``) lists every font referenced by the
document together with descriptive metadata such as PANOSE number, charset,
pitch, family classification, and (optionally) embedded font data. The part is
read-only from a document-authoring perspective — Word generates it and refreshes
it when a document is saved.
"""

from __future__ import annotations

from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_Fonts(BaseOxmlElement):
    """``<w:fonts>`` element, the root of the font table part.

    Contains a collection of ``<w:font>`` children, each describing a single
    font referenced by the document.
    """

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    font_lst: list[CT_Font]

    font = ZeroOrMore("w:font")

    def get_font_by_name(self, name: str) -> CT_Font | None:
        """Return the first ``w:font`` child whose ``w:name`` matches `name`, or |None|."""
        matches = self.xpath("./w:font[@w:name=$name]", name=name)
        return matches[0] if matches else None


class CT_Font(BaseOxmlElement):
    """``<w:font>`` element, metadata for a single font referenced by the document."""

    # -- the ``w:name`` attribute is the key identifying the font entry --
    name: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:name", ST_String
    )

    # -- child element declarations, in XSD order --
    _tag_seq = (
        "w:altName",
        "w:panose1",
        "w:charset",
        "w:family",
        "w:notTrueType",
        "w:pitch",
        "w:sig",
        "w:embedRegular",
        "w:embedBold",
        "w:embedItalic",
        "w:embedBoldItalic",
    )

    altName: CT_FontName | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:altName", successors=_tag_seq[1:]
    )
    panose1: CT_Panose | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:panose1", successors=_tag_seq[2:]
    )
    charset: CT_Charset | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:charset", successors=_tag_seq[3:]
    )
    family: CT_FontFamily | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:family", successors=_tag_seq[4:]
    )
    pitch: CT_Pitch | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:pitch", successors=_tag_seq[5:]
    )
    embedRegular: CT_FontRel | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:embedRegular", successors=_tag_seq[8:]
    )
    embedBold: CT_FontRel | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:embedBold", successors=_tag_seq[9:]
    )
    embedItalic: CT_FontRel | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:embedItalic", successors=_tag_seq[10:]
    )
    embedBoldItalic: CT_FontRel | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:embedBoldItalic", successors=()
    )

    del _tag_seq


class CT_FontName(BaseOxmlElement):
    """``<w:altName>`` element — alternate font name for substitution."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_Charset(BaseOxmlElement):
    """``<w:charset>`` element — character set identifier as a hex string."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )
    characterSet: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:characterSet", ST_String
    )


class CT_FontFamily(BaseOxmlElement):
    """``<w:family>`` element — font-family classification (swiss, roman, ...)."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_Pitch(BaseOxmlElement):
    """``<w:pitch>`` element — pitch classification (fixed, variable, default)."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_Panose(BaseOxmlElement):
    """``<w:panose1>`` element — 10-byte PANOSE-1 classification as a hex string."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_FontRel(BaseOxmlElement):
    """``<w:embedRegular>``/``<w:embedBold>``/etc. — reference to an embedded font part."""

    rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:id", ST_String
    )
    fontKey: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:fontKey", ST_String
    )
    subsetted: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:subsetted", ST_String
    )
