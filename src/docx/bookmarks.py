"""Bookmark-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Iterator

from docx.oxml.bookmarks import CT_BookmarkStart

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body

class Bookmarks:
    """Collection of |Bookmark| objects in the document."""

    def __init__(self, body: CT_Body):
        self._body = body

    def __iter__(self) -> Iterator[Bookmark]:
        return (
            Bookmark(bookmarkStart, self._body)
            for bookmarkStart in self._body.xpath(".//w:bookmarkStart")
        )

    def __len__(self) -> int:
        return len(self._body.xpath(".//w:bookmarkStart"))

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        return self.get(name) is not None

    def get(self, name: str) -> Bookmark | None:
        """Return the bookmark with `name`, or |None| if not found."""
        for bs in self._body.xpath(".//w:bookmarkStart"):
            if bs.name == name:
                return Bookmark(bs, self._body)
        return None

class Bookmark:
    """Proxy for a bookmark defined by a w:bookmarkStart/w:bookmarkEnd pair."""

    def __init__(self, bookmarkStart: CT_BookmarkStart, body: CT_Body):
        self._bookmarkStart = bookmarkStart
        self._body = body

    @property
    def name(self) -> str:
        return self._bookmarkStart.name

    @property
    def bookmark_id(self) -> int:
        return self._bookmarkStart.id

    def delete(self) -> None:
        """Remove this bookmark from the document."""
        bookmark_id = str(self._bookmarkStart.id)
        # -- find and remove the matching bookmarkEnd --
        ends = self._body.xpath(f".//w:bookmarkEnd[@w:id='{bookmark_id}']")
        for end in ends:
            end.getparent().remove(end)
        # -- remove the bookmarkStart --
        self._bookmarkStart.getparent().remove(self._bookmarkStart)
