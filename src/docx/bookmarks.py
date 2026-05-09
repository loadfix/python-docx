"""Bookmark-related proxy types.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING

from docx.oxml.bookmarks import CT_BookmarkStart
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body
    from docx.oxml.text.paragraph import CT_P
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


__all__ = ["Bookmark", "Bookmarks"]


class Bookmarks:
    """Collection of |Bookmark| objects in the document.

    Supports ``len()``, iteration in document order, ``name in bookmarks``
    containment checks, and dict-like lookup by bookmark name.

    .. versionadded:: 2026.05.0
    """

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

    def __getitem__(self, name: str) -> Bookmark:
        """Return the bookmark with `name`, raising ``KeyError`` if absent.

        .. versionadded:: 2026.05.0
        """
        bm = self.get(name)
        if bm is None:
            raise KeyError(name)
        return bm

    def get(self, name: str) -> Bookmark | None:
        """Return the bookmark with `name`, or |None| if not found.

        .. versionadded:: 2026.05.0
        """
        for bs in self._body.xpath(".//w:bookmarkStart"):
            if bs.name == name:
                return Bookmark(bs, self._body)
        return None

    def next_id(self) -> int:
        """Return the next unused ``@w:id`` value across every range marker.

        ECMA-376 requires every ``w:bookmarkStart`` / ``w:bookmarkEnd`` /
        ``w:moveFromRangeStart`` / ``w:moveFromRangeEnd`` /
        ``w:moveToRangeStart`` / ``w:moveToRangeEnd`` / ``w:commentRangeStart`` /
        ``w:commentRangeEnd`` to carry a document-unique ``@w:id``. This helper
        scans the full range-marker family so a newly-allocated ID will not
        collide with an existing one regardless of its origin.

        .. versionadded:: 2026.05.0
        """
        xpath = " | ".join(
            f".//w:{elt}/@w:id"
            for elt in (
                "bookmarkStart",
                "bookmarkEnd",
                "moveFromRangeStart",
                "moveFromRangeEnd",
                "moveToRangeStart",
                "moveToRangeEnd",
                "commentRangeStart",
                "commentRangeEnd",
            )
        )
        used_ids = [int(x) for x in self._body.xpath(xpath)]
        return max(used_ids, default=-1) + 1

    def add(
        self,
        name: str,
        start: Run | Paragraph,
        end: Run | Paragraph | None = None,
    ) -> Bookmark:
        """Add a bookmark named `name` spanning from `start` to `end`.

        `start` and `end` may be |Run| or |Paragraph| objects. When `end` is
        |None| it defaults to `start`, producing a bookmark that wraps a
        single run or a single paragraph.

        A |Run| anchor inserts the range-marker directly before/after that
        run. A |Paragraph| anchor inserts the marker after the paragraph's
        ``w:pPr`` (or as the first child if no ``pPr``) for ``start`` and as
        the last child for ``end`` — i.e. the bookmark wraps the paragraph's
        runnable content.

        `name` must be unique within the document; this method does not check
        — Word accepts duplicate names silently but treats them as ambiguous
        cross-reference targets.

        .. versionadded:: 2026.05.0
        """
        # -- import lazily to avoid a circular import at module load time --
        from docx.text.paragraph import Paragraph
        from docx.text.run import Run

        if end is None:
            end = start

        bookmark_id = self.next_id()

        if isinstance(start, Run):
            start._r.insert_bookmark_start_before(bookmark_id, name)
        elif isinstance(start, Paragraph):
            _insert_bookmark_start_into_paragraph(start._p, bookmark_id, name)
        else:
            raise TypeError(
                f"start must be a Run or Paragraph, got {type(start).__name__}"
            )

        if isinstance(end, Run):
            end._r.insert_bookmark_end_after(bookmark_id)
        elif isinstance(end, Paragraph):
            _append_bookmark_end_to_paragraph(end._p, bookmark_id)
        else:
            raise TypeError(
                f"end must be a Run or Paragraph, got {type(end).__name__}"
            )

        bookmarkStart = self._body.xpath(f".//w:bookmarkStart[@w:id='{bookmark_id}']")[0]
        return Bookmark(bookmarkStart, self._body)

    def remove(self, name: str) -> None:
        """Remove the bookmark named `name`.

        Raises ``KeyError`` if no bookmark with that name exists.

        .. versionadded:: 2026.05.0
        """
        bm = self.get(name)
        if bm is None:
            raise KeyError(name)
        bm.delete()


class Bookmark:
    """Proxy for a bookmark defined by a w:bookmarkStart/w:bookmarkEnd pair.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, bookmarkStart: CT_BookmarkStart, body: CT_Body):
        self._bookmarkStart = bookmarkStart
        self._body = body

    @property
    def name(self) -> str:
        return self._bookmarkStart.name

    @name.setter
    def name(self, value: str) -> None:
        """Rename this bookmark.

        Writes the new value to the ``w:bookmarkStart/@w:name`` attribute.
        The matching ``w:bookmarkEnd`` is identified by ``@w:id`` and does
        not carry a name, so no further update is needed.

        .. versionadded:: 2026.05.0
        """
        self._bookmarkStart.name = value

    @property
    def bookmark_id(self) -> int:
        return self._bookmarkStart.id

    @property
    def start_paragraph(self) -> Paragraph | None:
        """The |Paragraph| containing this bookmark's ``w:bookmarkStart``.

        Returns |None| if the start marker is not inside a ``w:p`` (e.g. it
        was inserted as a direct child of ``w:body`` between paragraphs).

        .. versionadded:: 2026.05.0
        """
        return self._paragraph_for(self._bookmarkStart)

    @property
    def end_paragraph(self) -> Paragraph | None:
        """The |Paragraph| containing this bookmark's ``w:bookmarkEnd``.

        Returns |None| if no matching ``w:bookmarkEnd`` is found or if the
        end marker is not inside a ``w:p``.

        .. versionadded:: 2026.05.0
        """
        end = self._bookmark_end
        if end is None:
            return None
        return self._paragraph_for(end)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """List of |Paragraph| objects the bookmark overlaps, in document order.

        The list starts with :attr:`start_paragraph`, ends with
        :attr:`end_paragraph`, and includes every intervening ``w:p`` sibling
        inside the same ``w:body``. For a same-paragraph bookmark the list has
        exactly one entry. An orphaned bookmark (no matching end marker)
        returns an empty list.

        .. versionadded:: 2026.05.0
        """
        from docx.text.paragraph import Paragraph

        start_p = self._closest_ancestor_p(self._bookmarkStart)
        end = self._bookmark_end
        end_p = self._closest_ancestor_p(end) if end is not None else None
        if start_p is None or end_p is None:
            return []

        p_list = self._body.xpath(".//w:p")
        try:
            start_idx = p_list.index(start_p)
            end_idx = p_list.index(end_p)
        except ValueError:
            return []
        if end_idx < start_idx:
            # -- malformed document: end precedes start --
            return []
        return [Paragraph(p, None) for p in p_list[start_idx : end_idx + 1]]  # type: ignore[arg-type]

    @property
    def text(self) -> str:
        """Concatenated text of every ``w:t`` descendant between start and end.

        The return value is the plain-text content of the bookmarked range:
        every ``w:t`` that is a descendant of a run following the
        ``w:bookmarkStart`` in document order and preceding the matching
        ``w:bookmarkEnd``. Tabs, line-breaks, and non-text run-level children
        are ignored — this mirrors the behaviour of :attr:`Paragraph.text`
        for the range between the markers.

        Returns an empty string for an orphan bookmark (no matching end) or a
        bookmark with no intervening ``w:t``.

        .. versionadded:: 2026.05.0
        """
        end = self._bookmark_end
        if end is None:
            return ""

        # -- collect every w:t descendant of the body, in document order,
        #    then slice to the ones that fall strictly after the start marker
        #    and strictly before the end marker in that order. --
        start = self._bookmarkStart
        in_range: list[str] = []
        started = False
        w_t = qn("w:t")
        w_bookmark_start = qn("w:bookmarkStart")
        w_bookmark_end = qn("w:bookmarkEnd")

        for node in self._body.iter():
            if node is start:
                started = True
                continue
            if node is end:
                break
            if not started:
                continue
            # -- a nested bookmarkStart/End with the same id would be
            #    malformed; filter range-markers out unconditionally. --
            if node.tag in (w_bookmark_start, w_bookmark_end):
                continue
            if node.tag == w_t and node.text:
                in_range.append(node.text)

        return "".join(in_range)

    def delete(self) -> None:
        """Remove this bookmark from the document.

        .. versionadded:: 2026.05.0
        """
        bookmark_id = str(self._bookmarkStart.id)
        # -- find and remove the matching bookmarkEnd --
        ends = self._body.xpath(f".//w:bookmarkEnd[@w:id='{bookmark_id}']")
        for end in ends:
            end.getparent().remove(end)
        # -- remove the bookmarkStart --
        self._bookmarkStart.getparent().remove(self._bookmarkStart)

    # -- internal helpers -------------------------------------------------

    @property
    def _bookmark_end(self):
        """The matching ``w:bookmarkEnd`` element, or |None| if absent."""
        bookmark_id = str(self._bookmarkStart.id)
        ends = self._body.xpath(f".//w:bookmarkEnd[@w:id='{bookmark_id}']")
        return ends[0] if ends else None

    def _paragraph_for(self, elt) -> Paragraph | None:
        """Return |Paragraph| ancestor of `elt`, or |None| if none."""
        from docx.text.paragraph import Paragraph

        p = self._closest_ancestor_p(elt)
        if p is None:
            return None
        return Paragraph(p, None)  # type: ignore[arg-type]

    @staticmethod
    def _closest_ancestor_p(elt) -> CT_P | None:
        """The closest ``w:p`` ancestor of `elt`, or |None| if none.

        Returns |None| if `elt` itself is outside any ``w:p`` (e.g. a
        range-marker sitting directly inside ``w:body``).
        """
        if elt is None:
            return None
        w_p = qn("w:p")
        anc = elt.getparent()
        while anc is not None:
            if anc.tag == w_p:
                return anc  # type: ignore[return-value]
            anc = anc.getparent()
        return None


def _insert_bookmark_start_into_paragraph(p, bookmark_id: int, name: str) -> None:
    """Insert a ``w:bookmarkStart`` element as the first runnable child of `p`.

    The marker is placed immediately after ``w:pPr`` when present, otherwise
    as the first child — this keeps the paragraph-properties-first ordering
    required by the schema.
    """
    from docx.oxml.parser import OxmlElement

    element = OxmlElement(
        "w:bookmarkStart",
        attrs={qn("w:id"): str(bookmark_id), qn("w:name"): name},
    )
    pPr = p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(element)
    else:
        p.insert(0, element)


def _append_bookmark_end_to_paragraph(p, bookmark_id: int) -> None:
    """Append a ``w:bookmarkEnd`` element as the last child of `p`."""
    from docx.oxml.parser import OxmlElement

    element = OxmlElement("w:bookmarkEnd", attrs={qn("w:id"): str(bookmark_id)})
    p.append(element)
