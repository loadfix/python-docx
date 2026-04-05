"""Custom element classes related to tracked changes (revisions)."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, List

from docx.oxml.simpletypes import ST_DateTime, ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.text.run import CT_R


class CT_RunTrackChange(BaseOxmlElement):
    """Base for `<w:ins>` and `<w:del>` elements wrapping runs in a paragraph.

    Both share the same attribute set: `w:id`, `w:author`, and `w:date`.
    """

    r_lst: List[CT_R]

    r = ZeroOrMore("w:r", successors=())

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    author: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:author", ST_String
    )
    date: dt.datetime | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:date", ST_DateTime
    )


class CT_Ins(CT_RunTrackChange):
    """`<w:ins>` element, containing runs that were inserted."""

    @property
    def text(self) -> str:
        """The textual content of the inserted runs."""
        return "".join(r.text for r in self.r_lst)


class CT_Del(CT_RunTrackChange):
    """`<w:del>` element, containing runs that were deleted."""

    @property
    def text(self) -> str:
        """The textual content of the deleted runs.

        Deleted runs use `w:delText` elements rather than `w:t`.
        """
        return "".join(
            str(e) for e in self.xpath("w:r/w:delText")
        )


class CT_DelText(BaseOxmlElement):
    """`<w:delText>` element, containing text in a deleted run."""

    def __str__(self) -> str:
        """Text contained in this element, the empty string if it has no content."""
        return self.text or ""
