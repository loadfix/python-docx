# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.bookmarks` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.bookmarks import CT_BookmarkEnd, CT_BookmarkStart

from ..unitutil.cxml import element


class DescribeCT_BookmarkStart:
    """Unit-test suite for `docx.oxml.bookmarks.CT_BookmarkStart`."""

    def it_knows_its_id(self):
        bookmarkStart = cast(CT_BookmarkStart, element("w:bookmarkStart{w:id=7,w:name=bm1}"))
        assert bookmarkStart.id == 7

    def it_knows_its_name(self):
        bookmarkStart = cast(CT_BookmarkStart, element("w:bookmarkStart{w:id=7,w:name=bm1}"))
        assert bookmarkStart.name == "bm1"


class DescribeCT_BookmarkEnd:
    """Unit-test suite for `docx.oxml.bookmarks.CT_BookmarkEnd`."""

    def it_knows_its_id(self):
        bookmarkEnd = cast(CT_BookmarkEnd, element("w:bookmarkEnd{w:id=7}"))
        assert bookmarkEnd.id == 7
