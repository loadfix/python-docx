"""Custom element classes related to ruby (phonetic annotation) content.

Ruby is used for Japanese furigana and similar above-the-line pronunciation hints.
The OOXML schema places `w:ruby` inside a run as a pairing of base text with the
ruby (annotation) text.
"""

from __future__ import annotations

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


class CT_RubyAlign(BaseOxmlElement):
    """`<w:rubyAlign>` inside `w:rubyPr`."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_RubyLang(BaseOxmlElement):
    """`<w:lid>` inside `w:rubyPr` (language identifier)."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_RubyHps(BaseOxmlElement):
    """`<w:hps>` / `<w:hpsRaise>` / `<w:hpsBaseText>` (half-point measures)."""

    val: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_DecimalNumber
    )


class CT_RubyPr(BaseOxmlElement):
    """`<w:rubyPr>` — ruby properties (alignment, sizes, language)."""

    rubyAlign: "CT_RubyAlign | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rubyAlign"
    )
    hps: "CT_RubyHps | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:hps"
    )
    hpsRaise: "CT_RubyHps | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:hpsRaise"
    )
    hpsBaseText: "CT_RubyHps | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:hpsBaseText"
    )
    lid: "CT_RubyLang | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:lid"
    )


class CT_RubyContent(BaseOxmlElement):
    """Container for runs inside `w:rt` (annotation) or `w:rubyBase` (base)."""

    @property
    def text(self) -> str:
        """Concatenated `w:t` text of any run children."""
        return "".join(str(t) for t in self.xpath(".//w:t"))


class CT_Ruby(BaseOxmlElement):
    """`<w:ruby>` — container for a ruby annotation.

    Has three children: `w:rubyPr` (properties), `w:rt` (annotation), and
    `w:rubyBase` (base text).
    """

    rubyPr: "CT_RubyPr | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rubyPr"
    )
    rt: "CT_RubyContent | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rt"
    )
    rubyBase: "CT_RubyContent | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rubyBase"
    )

    @property
    def base_text(self) -> str:
        """Concatenated text of the ruby base runs, empty if absent."""
        return self.rubyBase.text if self.rubyBase is not None else ""

    @property
    def ruby_text(self) -> str:
        """Concatenated text of the annotation (`w:rt`) runs, empty if absent."""
        return self.rt.text if self.rt is not None else ""

    def __str__(self) -> str:
        """Base text — used by run.text extraction so paragraph.text stays sensible."""
        return self.base_text
