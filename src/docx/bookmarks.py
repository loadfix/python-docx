"""Bookmark-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.oxml.bookmarks import CT_BookmarkStart

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body
    from docx.text.paragraph import Paragraph


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
        return bool(self._body.xpath(f".//w:bookmarkStart[@w:name='{name}']"))

    def get(self, name: str) -> Bookmark | None:
        """Return the bookmark with `name`, or |None| if not found."""
        results = self._body.xpath(f".//w:bookmarkStart[@w:name='{name}']")
        if not results:
            return None
        return Bookmark(results[0], self._body)


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

    @property
    def paragraph(self) -> Paragraph:
        """The paragraph containing the bookmarkStart element."""
        from docx.text.paragraph import Paragraph

        p = self._bookmarkStart.getparent()
        # --- walk up if bookmarkStart is not directly inside a w:p ---
        from docx.oxml.text.paragraph import CT_P

        while p is not None and not isinstance(p, CT_P):
            p = p.getparent()

        if p is None:
            raise ValueError("bookmarkStart is not contained in a paragraph")

        # --- find the story part parent for the Paragraph ---
        # --- walk up from the CT_P to find the body, then use the body's parent ---
        return Paragraph(p, None)  # type: ignore[arg-type]

    def delete(self) -> None:
        """Remove this bookmark from the document."""
        bookmark_id = str(self._bookmarkStart.id)
        # -- find and remove the matching bookmarkEnd --
        ends = self._body.xpath(f".//w:bookmarkEnd[@w:id='{bookmark_id}']")
        for end in ends:
            end.getparent().remove(end)
        # -- remove the bookmarkStart --
        self._bookmarkStart.getparent().remove(self._bookmarkStart)
