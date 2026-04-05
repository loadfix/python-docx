"""Custom element classes related to the endnotes part."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Endnotes(BaseOxmlElement):
    """`w:endnotes` element, the root element for the endnotes part."""

    endnote_lst: list[CT_Endnote]

    endnote = ZeroOrMore("w:endnote")

    def add_endnote(self) -> CT_Endnote:
        """Return newly added `w:endnote` child element.

        The returned `w:endnote` element has a unique `w:id` value and contains a single
        paragraph with an endnote reference run. Content is added by adding runs to this first
        paragraph and by adding additional paragraphs as needed.
        """
        next_id = self._next_available_endnote_id()
        endnote = cast(
            CT_Endnote,
            parse_xml(
                f'<w:endnote {nsdecls("w")} w:id="{next_id}">'
                f"  <w:p>"
                f"    <w:pPr>"
                f'      <w:pStyle w:val="EndnoteText"/>'
                f"    </w:pPr>"
                f"    <w:r>"
                f"      <w:rPr>"
                f'        <w:rStyle w:val="EndnoteReference"/>'
                f"      </w:rPr>"
                f"      <w:endnoteRef/>"
                f"    </w:r>"
                f"  </w:p>"
                f"</w:endnote>"
            ),
        )
        self.append(endnote)
        return endnote

    def _next_available_endnote_id(self) -> int:
        """The next available endnote id.

        IDs 0 and 1 are reserved for the separator and continuation separator. User endnotes
        start at 2.
        """
        used_ids = [int(x) for x in self.xpath("./w:endnote/@w:id")]

        next_id = max(used_ids, default=1) + 1

        if next_id < 2:
            return 2

        if next_id <= 2**31 - 1:
            return next_id

        # -- fall-back to enumerating all used ids to find the first unused one --
        used_id_set = set(used_ids)
        for expected_id in range(2, 2**31):
            if expected_id not in used_id_set:
                return expected_id

        raise ValueError("No available endnote ID: document has reached the maximum endnote count.")


class CT_Endnote(BaseOxmlElement):
    """`w:endnote` element, representing a single endnote.

    An endnote can contain paragraphs and tables, much like a comment or table-cell.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("w:type", ST_String)  # pyright: ignore[reportAssignmentType]

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --
    add_p: Callable[[], CT_P]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    def clear_content(self) -> None:
        """Remove all child elements and add a single empty paragraph.

        The empty paragraph has the "EndnoteText" style applied and contains a
        `w:endnoteRef` run so the auto-numbered reference mark is preserved.
        """
        for child in list(self):
            self.remove(child)
        self.append(
            parse_xml(
                f'<w:p {nsdecls("w")}>'
                f"  <w:pPr>"
                f'    <w:pStyle w:val="EndnoteText"/>'
                f"  </w:pPr>"
                f"  <w:r>"
                f"    <w:rPr>"
                f'      <w:rStyle w:val="EndnoteReference"/>'
                f"    </w:rPr>"
                f"    <w:endnoteRef/>"
                f"  </w:r>"
                f"</w:p>"
            )
        )

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Return all `w:p` and `w:tbl` elements in this endnote."""
        return self.xpath("./w:p | ./w:tbl")
