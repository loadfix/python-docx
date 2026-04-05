"""Custom element classes related to the footnotes part."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Footnotes(BaseOxmlElement):
    """`w:footnotes` element, the root element for the footnotes part."""

    footnote_lst: list[CT_Footnote]

    footnote = ZeroOrMore("w:footnote")

    def add_footnote(self) -> CT_Footnote:
        """Return newly added `w:footnote` child element.

        The returned `w:footnote` element has a unique `w:id` value and contains a single
        paragraph with a footnote reference run. Content is added by adding runs to this first
        paragraph and by adding additional paragraphs as needed.
        """
        next_id = self._next_available_footnote_id()
        footnote = cast(
            CT_Footnote,
            parse_xml(
                f'<w:footnote {nsdecls("w")} w:id="{next_id}">'
                f"  <w:p>"
                f"    <w:pPr>"
                f'      <w:pStyle w:val="FootnoteText"/>'
                f"    </w:pPr>"
                f"    <w:r>"
                f"      <w:rPr>"
                f'        <w:rStyle w:val="FootnoteReference"/>'
                f"      </w:rPr>"
                f"      <w:footnoteRef/>"
                f"    </w:r>"
                f"  </w:p>"
                f"</w:footnote>"
            ),
        )
        self.append(footnote)
        return footnote

    def _next_available_footnote_id(self) -> int:
        """The next available footnote id.

        IDs 0 and 1 are reserved for the separator and continuation separator. User footnotes
        start at 2.
        """
        used_ids = [int(x) for x in self.xpath("./w:footnote/@w:id")]

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

        raise ValueError("No available footnote ID: document has reached the maximum footnote count.")


class CT_Footnote(BaseOxmlElement):
    """`w:footnote` element, representing a single footnote.

    A footnote can contain paragraphs and tables, much like a comment or table-cell.
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

        The empty paragraph has the "FootnoteText" style applied, which is the default
        style for footnote content.
        """
        for child in list(self):
            self.remove(child)
        self.append(
            parse_xml(
                f'<w:p {nsdecls("w")}>'
                f"  <w:pPr>"
                f'    <w:pStyle w:val="FootnoteText"/>'
                f"  </w:pPr>"
                f"</w:p>"
            )
        )

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Return all `w:p` and `w:tbl` elements in this footnote."""
        return self.xpath("./w:p | ./w:tbl")
