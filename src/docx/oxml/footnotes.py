"""Custom element classes related to the footnotes part."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Footnotes(BaseOxmlElement):
    """`w:footnotes` element, the root element for the footnotes part."""

    footnote_lst: list[CT_Footnote]

    footnote = ZeroOrMore("w:footnote")


class CT_Footnote(BaseOxmlElement):
    """`w:footnote` element, representing a single footnote.

    A footnote can contain paragraphs and tables, much like a comment or table-cell.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("w:type", ST_String)  # pyright: ignore[reportAssignmentType]

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Return all `w:p` and `w:tbl` elements in this footnote."""
        return self.xpath("./w:p | ./w:tbl")
