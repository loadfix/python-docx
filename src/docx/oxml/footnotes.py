"""Custom element classes related to the footnotes part."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from docx.oxml.simpletypes import ST_DecimalNumber
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Footnotes(BaseOxmlElement):
    """`w:footnotes` element, the root element for the footnotes part."""

    footnote_lst: list[CT_Footnote]

    footnote = ZeroOrMore("w:footnote")

    def _next_available_footnote_id(self) -> int:
        """Return the next available footnote id (>= 2, since 0 and 1 are reserved)."""
        used_ids = [int(x) for x in self.xpath("./w:footnote/@w:id")]
        next_id = max(used_ids, default=1) + 1
        return max(next_id, 2)


class CT_Footnote(BaseOxmlElement):
    """`w:footnote` element, representing a single footnote.

    A footnote can contain paragraphs and tables, much like a comment or table-cell.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("w:type", str)  # pyright: ignore[reportAssignmentType]

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    add_p: Callable[[], CT_P]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this footnote."""
        return self.xpath("./w:p | ./w:tbl")
