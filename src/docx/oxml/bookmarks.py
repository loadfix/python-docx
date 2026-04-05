"""Custom element classes related to bookmarks."""

from __future__ import annotations

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute


class CT_BookmarkStart(BaseOxmlElement):
    """`w:bookmarkStart` element, marking the start of a bookmarked range."""

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_BookmarkEnd(BaseOxmlElement):
    """`w:bookmarkEnd` element, marking the end of a bookmarked range."""

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
