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

    def it_exposes_its_start_and_end_paragraphs(self):
        from docx.text.paragraph import Paragraph as ParagraphCls

        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:r/w:t\"aaa\")"
                ",w:p/w:r/w:t\"bbb\""
                ",w:p/(w:r/w:t\"ccc\",w:bookmarkEnd{w:id=0}))"
            ),
        )
        bm = next(iter(Bookmarks(body)))

        assert isinstance(bm.start_paragraph, ParagraphCls)
        assert isinstance(bm.end_paragraph, ParagraphCls)
        # -- start_paragraph is the first w:p, end_paragraph is the third --
        p_list = body.xpath(".//w:p")
        assert bm.start_paragraph._p is p_list[0]
        assert bm.end_paragraph._p is p_list[2]

    def it_returns_none_paragraphs_for_an_orphan_bookmark(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:r/w:t\"hello\")"
            ),
        )
        bm = next(iter(Bookmarks(body)))

        assert bm.start_paragraph is not None
        assert bm.end_paragraph is None

    def it_lists_every_paragraph_overlapped_by_the_range(self):
        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:r/w:t\"aaa\")"
                ",w:p/w:r/w:t\"bbb\""
                ",w:p/(w:r/w:t\"ccc\",w:bookmarkEnd{w:id=0}))"
            ),
        )
        bm = next(iter(Bookmarks(body)))

        paras = bm.paragraphs
        assert len(paras) == 3
        assert [p.text for p in paras] == ["aaa", "bbb", "ccc"]

    def it_computes_text_between_start_and_end_markers(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:r/w:t\"before\""
                ",w:bookmarkStart{w:id=0,w:name=bm1}"
                ",w:r/w:t\"inside \""
                ",w:r/w:t\"range\""
                ",w:bookmarkEnd{w:id=0}"
                ",w:r/w:t\"after\")"
            ),
        )
        bm = next(iter(Bookmarks(body)))

        assert bm.text == "inside range"

    def it_computes_text_spanning_paragraphs(self):
        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:r/w:t\"aaa\")"
                ",w:p/(w:r/w:t\"bbb\",w:bookmarkEnd{w:id=0}))"
            ),
        )
        bm = next(iter(Bookmarks(body)))

        assert bm.text == "aaabbb"

    def it_returns_empty_text_for_orphan_bookmark(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:r/w:t\"hello\")"
            ),
        )
        bm = next(iter(Bookmarks(body)))

        assert bm.text == ""


class DescribeBookmarks_dict_like:
    """Unit-test suite for dict-like lookup and mutation."""

    def it_can_subscript_by_name(self):
        import pytest

        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
            ),
        )
        bookmarks = Bookmarks(body)

        assert bookmarks["bm1"].name == "bm1"
        with pytest.raises(KeyError):
            _ = bookmarks["missing"]

    def it_allocates_ids_above_all_range_markers(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1}"
                ",w:bookmarkEnd{w:id=0}"
                ",w:commentRangeStart{w:id=5}"
                ",w:commentRangeEnd{w:id=5}"
                ",w:moveFromRangeStart{w:id=9,w:name=mv,w:author=u,w:date=2026-01-01T00:00:00Z}"
                ",w:moveFromRangeEnd{w:id=9})"
            ),
        )
        bookmarks = Bookmarks(body)

        assert bookmarks.next_id() == 10

    def it_can_add_a_bookmark_spanning_a_single_run(self):
        from docx.text.paragraph import Paragraph

        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        para = Paragraph(body.p_lst[0], None)  # type: ignore[arg-type]
        run = para.runs[0]
        bookmarks = Bookmarks(body)

        bm = bookmarks.add("bm1", run)

        assert bm.name == "bm1"
        assert bm.bookmark_id == 0
        assert len(bookmarks) == 1
        assert bm.text == "hello"

    def it_can_add_a_bookmark_spanning_runs(self):
        from docx.text.paragraph import Paragraph

        body = cast(
            CT_Body, element('w:body/w:p/(w:r/w:t"aaa",w:r/w:t"bbb",w:r/w:t"ccc")')
        )
        para = Paragraph(body.p_lst[0], None)  # type: ignore[arg-type]
        runs = para.runs
        bookmarks = Bookmarks(body)

        bm = bookmarks.add("bm_mid", runs[0], runs[2])

        assert bm.text == "aaabbbccc"

    def it_can_add_a_bookmark_spanning_paragraphs(self):
        from docx.text.paragraph import Paragraph

        body = cast(
            CT_Body, element('w:body/(w:p/w:r/w:t"aaa",w:p/w:r/w:t"bbb")')
        )
        p1 = Paragraph(body.p_lst[0], None)  # type: ignore[arg-type]
        p2 = Paragraph(body.p_lst[1], None)  # type: ignore[arg-type]
        bookmarks = Bookmarks(body)

        bm = bookmarks.add("span", p1, p2)

        assert bm.text == "aaabbb"
        assert len(bm.paragraphs) == 2

    def it_can_remove_a_bookmark_by_name(self):
        import pytest

        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
            ),
        )
        bookmarks = Bookmarks(body)

        bookmarks.remove("bm1")

        assert len(bookmarks) == 0
        with pytest.raises(KeyError):
            bookmarks.remove("bm1")

    def it_parses_nested_bookmarks(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:bookmarkStart{w:id=0,w:name=outer}"
                ",w:r/w:t\"aa\""
                ",w:bookmarkStart{w:id=1,w:name=inner}"
                ",w:r/w:t\"bb\""
                ",w:bookmarkEnd{w:id=1}"
                ",w:r/w:t\"cc\""
                ",w:bookmarkEnd{w:id=0})"
            ),
        )
        bookmarks = Bookmarks(body)

        assert len(bookmarks) == 2
        outer = bookmarks["outer"]
        inner = bookmarks["inner"]
        assert outer.text == "aabbcc"
        assert inner.text == "bb"
