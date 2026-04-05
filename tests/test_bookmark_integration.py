# pyright: reportPrivateUsage=false

"""Integration tests for bookmark feature across paragraph and document."""

from __future__ import annotations

from typing import cast

from docx.bookmarks import Bookmark, Bookmarks
from docx.oxml.document import CT_Body, CT_Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .unitutil.cxml import element


class DescribeParagraph_add_bookmark:
    """Unit-test suite for `Paragraph.add_bookmark()`."""

    def it_can_add_a_bookmark_wrapping_whole_paragraph(self):
        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        bm = para.add_bookmark("test_bm")

        assert isinstance(bm, Bookmark)
        assert bm.name == "test_bm"
        assert bm.bookmark_id == 0
        # -- bookmarkStart is first child (no pPr), bookmarkEnd is last --
        children = list(p_elm)
        assert children[0].tag == qn("w:bookmarkStart")
        assert children[-1].tag == qn("w:bookmarkEnd")

    def it_can_add_a_bookmark_wrapping_whole_paragraph_with_pPr(self):
        body = cast(CT_Body, element('w:body/w:p/(w:pPr,w:r/w:t"hello")'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        bm = para.add_bookmark("test_bm")

        assert bm.name == "test_bm"
        children = list(p_elm)
        # -- pPr is first, then bookmarkStart, then run, then bookmarkEnd --
        assert children[0].tag == qn("w:pPr")
        assert children[1].tag == qn("w:bookmarkStart")
        assert children[-1].tag == qn("w:bookmarkEnd")

    def it_can_add_a_bookmark_around_specific_runs(self):
        body = cast(
            CT_Body,
            element('w:body/w:p/(w:r/w:t"aaa",w:r/w:t"bbb",w:r/w:t"ccc")'),
        )
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]
        runs = para.runs

        bm = para.add_bookmark("mid", start_run=runs[1], end_run=runs[1])

        assert bm.name == "mid"
        # -- bookmarkStart is before the second run, bookmarkEnd is after it --
        children = list(p_elm)
        tags = [c.tag for c in children]
        bs_idx = tags.index(qn("w:bookmarkStart"))
        be_idx = tags.index(qn("w:bookmarkEnd"))
        # bookmarkStart should be right before second w:r
        assert tags[bs_idx + 1] == qn("w:r")
        # bookmarkEnd should be right after that same w:r
        assert be_idx == bs_idx + 2

    def it_allocates_unique_ids(self):
        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        bm1 = para.add_bookmark("bm1")
        bm2 = para.add_bookmark("bm2")

        assert bm1.bookmark_id == 0
        assert bm2.bookmark_id == 1

    def it_can_add_a_bookmark_with_only_start_run(self):
        body = cast(
            CT_Body,
            element('w:body/w:p/(w:r/w:t"aaa",w:r/w:t"bbb")'),
        )
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]
        runs = para.runs

        bm = para.add_bookmark("single", start_run=runs[0])

        assert bm.name == "single"
        assert bm.bookmark_id == 0


class DescribeDocument_bookmarks:
    """Unit-test suite for `Document.bookmarks`."""

    def it_provides_access_to_document_bookmarks(self):
        from docx.document import Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/"
                "(w:bookmarkStart{w:id=0,w:name=bm1},w:bookmarkEnd{w:id=0})"
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        bookmarks = doc.bookmarks

        assert isinstance(bookmarks, Bookmarks)
        assert len(bookmarks) == 1
        bm = next(iter(bookmarks))
        assert bm.name == "bm1"
