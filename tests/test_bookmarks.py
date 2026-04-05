# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.bookmarks` module."""

from __future__ import annotations

from typing import cast

from docx.bookmarks import Bookmark, Bookmarks
from docx.oxml.bookmarks import CT_BookmarkStart
from docx.oxml.document import CT_Body

from .unitutil.cxml import element


class DescribeBookmarks:
    """Unit-test suite for `docx.bookmarks.Bookmarks` objects."""

    def it_knows_how_many_bookmarks_it_contains(self):
        body = cast(CT_Body, element("w:body"))
        assert len(Bookmarks(body)) == 0

        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
            ),
        )
        assert len(Bookmarks(body)) == 1

        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
                ",w:p/(w:bookmarkStart{w:id=1,w:name=bm2},w:bookmarkEnd{w:id=1}))"
            ),
        )
        assert len(Bookmarks(body)) == 2

    def it_is_iterable_over_bookmarks(self):
        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
                ",w:p/(w:bookmarkStart{w:id=1,w:name=bm2},w:bookmarkEnd{w:id=1}))"
            ),
        )
        bookmarks = Bookmarks(body)

        bm_iter = iter(bookmarks)
        bm1 = next(bm_iter)
        assert isinstance(bm1, Bookmark)
        assert bm1.name == "bm1"
        bm2 = next(bm_iter)
        assert isinstance(bm2, Bookmark)
        assert bm2.name == "bm2"

    def it_supports_containment_check_by_name(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
            ),
        )
        bookmarks = Bookmarks(body)
        assert "bm1" in bookmarks
        assert "nonexistent" not in bookmarks

    def it_can_get_a_bookmark_by_name(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
            ),
        )
        bookmarks = Bookmarks(body)

        bm = bookmarks.get("bm1")
        assert bm is not None
        assert bm.name == "bm1"

        assert bookmarks.get("nonexistent") is None


class DescribeBookmark:
    """Unit-test suite for `docx.bookmarks.Bookmark`."""

    def it_knows_its_name(self):
        body = cast(CT_Body, element("w:body"))
        bookmarkStart = cast(
            CT_BookmarkStart,
            element("w:bookmarkStart{w:id=5,w:name=test_bookmark}"),
        )
        bm = Bookmark(bookmarkStart, body)
        assert bm.name == "test_bookmark"

    def it_knows_its_bookmark_id(self):
        body = cast(CT_Body, element("w:body"))
        bookmarkStart = cast(
            CT_BookmarkStart,
            element("w:bookmarkStart{w:id=42,w:name=bm1}"),
        )
        bm = Bookmark(bookmarkStart, body)
        assert bm.bookmark_id == 42

    def it_can_delete_itself(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1}"
                ",w:r/w:t\"hello\""
                ",w:bookmarkEnd{w:id=0})"
            ),
        )
        bookmarks = Bookmarks(body)
        assert len(bookmarks) == 1

        bm = next(iter(bookmarks))
        bm.delete()

        assert len(bookmarks) == 0
        # -- bookmarkEnd is also removed --
        assert len(body.xpath(".//w:bookmarkEnd")) == 0

    def it_can_delete_a_cross_paragraph_bookmark(self):
        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:r/w:t\"hello\")"
                ",w:p/(w:r/w:t\"world\",w:bookmarkEnd{w:id=0}))"
            ),
        )
        bookmarks = Bookmarks(body)
        assert len(bookmarks) == 1

        bm = next(iter(bookmarks))
        bm.delete()

        assert len(bookmarks) == 0
        assert len(body.xpath(".//w:bookmarkEnd")) == 0
