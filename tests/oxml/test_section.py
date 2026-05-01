"""Unit-test suite for `docx.oxml.section` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.section import (
    WD_BORDER_DISPLAY,
    WD_BORDER_OFFSET_FROM,
    WD_LINE_NUMBERING_RESTART,
)
from docx.enum.text import WD_BORDER_STYLE
from docx.oxml.section import (
    CT_Col,
    CT_Cols,
    CT_HdrFtr,
    CT_LineNumber,
    CT_PgBorders,
    CT_SectPr,
)
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Emu, Inches, Length, RGBColor, Twips

from ..unitutil.cxml import element, xml


class DescribeCT_Col:
    """Unit-test suite for `docx.oxml.section.CT_Col`."""

    @pytest.mark.parametrize(
        ("col_cxml", "expected_w", "expected_space"),
        [
            ("w:col", None, None),
            ("w:col{w:w=4320,w:space=720}", Twips(4320), Twips(720)),
        ],
    )
    def it_knows_its_width_and_space(self, col_cxml, expected_w, expected_space):
        col = cast(CT_Col, element(col_cxml))
        assert col.w == expected_w
        assert col.space == expected_space


class DescribeCT_Cols:
    """Unit-test suite for `docx.oxml.section.CT_Cols`."""

    @pytest.mark.parametrize(
        ("cols_cxml", "expected_num", "expected_space", "expected_eq"),
        [
            ("w:cols", None, None, None),
            ("w:cols{w:num=2,w:space=720,w:equalWidth=1}", 2, Twips(720), True),
            ("w:cols{w:num=3,w:equalWidth=0}", 3, None, False),
        ],
    )
    def it_knows_its_attributes(self, cols_cxml, expected_num, expected_space, expected_eq):
        cols = cast(CT_Cols, element(cols_cxml))
        assert cols.num == expected_num
        assert cols.space == expected_space
        assert cols.equalWidth == expected_eq

    def it_provides_access_to_its_col_children(self):
        cols = cast(
            CT_Cols,
            element("w:cols/(w:col{w:w=4320,w:space=720},w:col{w:w=4320})"),
        )
        col_lst = cols.col_lst
        assert len(col_lst) == 2
        assert col_lst[0].w == Twips(4320)
        assert col_lst[0].space == Twips(720)
        assert col_lst[1].w == Twips(4320)
        assert col_lst[1].space is None


class DescribeCT_SectPr_cols:
    """Unit-test suite for CT_SectPr column-related features."""

    def it_can_access_its_cols_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr/w:cols{w:num=2}"))
        cols = sectPr.cols
        assert cols is not None
        assert cols.num == 2

    def it_returns_None_when_no_cols_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.cols is None

    def it_can_add_a_cols_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        cols = sectPr.get_or_add_cols()
        assert cols is not None
        assert sectPr.cols is cols

    def it_inserts_cols_in_the_right_position(self):
        sectPr = cast(CT_SectPr, element("w:sectPr/w:pgMar"))
        cols = sectPr.get_or_add_cols()
        assert cols is not None
        expected = xml("w:sectPr/(w:pgMar,w:cols)")
        assert sectPr.xml == expected


class DescribeCT_PgBorders:
    """Unit-test suite for `docx.oxml.section.CT_PgBorders`."""

    @pytest.mark.parametrize(
        ("pgBorders_cxml", "expected_display", "expected_offset"),
        [
            ("w:pgBorders", None, None),
            (
                "w:pgBorders{w:display=allPages,w:offsetFrom=page}",
                WD_BORDER_DISPLAY.ALL_PAGES,
                WD_BORDER_OFFSET_FROM.PAGE,
            ),
            (
                "w:pgBorders{w:display=firstPage,w:offsetFrom=text}",
                WD_BORDER_DISPLAY.FIRST_PAGE,
                WD_BORDER_OFFSET_FROM.TEXT,
            ),
            (
                "w:pgBorders{w:display=notFirstPage}",
                WD_BORDER_DISPLAY.NOT_FIRST_PAGE,
                None,
            ),
        ],
    )
    def it_knows_its_attributes(
        self, pgBorders_cxml, expected_display, expected_offset
    ):
        pgBorders = cast(CT_PgBorders, element(pgBorders_cxml))
        assert pgBorders.display == expected_display
        assert pgBorders.offsetFrom == expected_offset

    def it_can_access_each_edge_child(self):
        from docx.shared import Pt

        pgBorders = cast(
            CT_PgBorders,
            element(
                "w:pgBorders/(w:top{w:val=single,w:sz=24,w:space=24,w:color=FF0000},"
                "w:left{w:val=dashed,w:sz=8,w:space=12,w:color=00FF00},"
                "w:bottom{w:val=double,w:sz=4,w:space=6,w:color=0000FF},"
                "w:right{w:val=dotted,w:sz=16,w:space=18,w:color=AABBCC})"
            ),
        )
        assert pgBorders.top is not None
        assert pgBorders.top.val == WD_BORDER_STYLE.SINGLE
        assert pgBorders.top.sz == Pt(24 / 8.0)
        assert pgBorders.top.space == Pt(24)
        assert pgBorders.top.color == RGBColor(0xFF, 0x00, 0x00)
        assert pgBorders.left is not None
        assert pgBorders.left.val == WD_BORDER_STYLE.DASHED
        assert pgBorders.bottom is not None
        assert pgBorders.bottom.val == WD_BORDER_STYLE.DOUBLE
        assert pgBorders.right is not None
        assert pgBorders.right.val == WD_BORDER_STYLE.DOTTED

    def it_returns_None_for_missing_edge_children(self):
        pgBorders = cast(CT_PgBorders, element("w:pgBorders"))
        assert pgBorders.top is None
        assert pgBorders.bottom is None
        assert pgBorders.left is None
        assert pgBorders.right is None

    def it_can_add_each_edge(self):
        pgBorders = cast(CT_PgBorders, element("w:pgBorders"))
        top = pgBorders.get_or_add_top()
        left = pgBorders.get_or_add_left()
        bottom = pgBorders.get_or_add_bottom()
        right = pgBorders.get_or_add_right()
        assert pgBorders.top is top
        assert pgBorders.left is left
        assert pgBorders.bottom is bottom
        assert pgBorders.right is right
        expected = xml("w:pgBorders/(w:top,w:left,w:bottom,w:right)")
        assert pgBorders.xml == expected


class DescribeCT_SectPr_pgBorders:
    """Unit-test suite for CT_SectPr page-border features."""

    def it_returns_None_when_no_pgBorders_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.pgBorders is None

    def it_can_access_its_pgBorders_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:pgBorders{w:display=allPages}"),
        )
        pgBorders = sectPr.pgBorders
        assert pgBorders is not None
        assert pgBorders.display == WD_BORDER_DISPLAY.ALL_PAGES

    def it_can_add_a_pgBorders_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        pgBorders = sectPr.get_or_add_pgBorders()
        assert pgBorders is not None
        assert sectPr.pgBorders is pgBorders

    def it_inserts_pgBorders_in_the_right_position(self):
        sectPr = cast(CT_SectPr, element("w:sectPr/(w:pgSz,w:pgMar,w:cols)"))
        sectPr.get_or_add_pgBorders()
        expected = xml("w:sectPr/(w:pgSz,w:pgMar,w:pgBorders,w:cols)")
        assert sectPr.xml == expected

    def it_can_remove_its_pgBorders_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:pgBorders/(w:top{w:val=single})"),
        )
        sectPr._remove_pgBorders()  # pyright: ignore[reportPrivateUsage]
        assert sectPr.pgBorders is None


class DescribeCT_LineNumber:
    """Unit-test suite for `docx.oxml.section.CT_LineNumber`."""

    @pytest.mark.parametrize(
        ("lnNumType_cxml", "count_by", "start", "distance", "restart"),
        [
            ("w:lnNumType", None, None, None, None),
            (
                "w:lnNumType{w:countBy=1,w:start=1,w:distance=360,w:restart=continuous}",
                1,
                1,
                Twips(360),
                WD_LINE_NUMBERING_RESTART.CONTINUOUS,
            ),
            (
                "w:lnNumType{w:countBy=5,w:start=10,w:distance=720,w:restart=newSection}",
                5,
                10,
                Twips(720),
                WD_LINE_NUMBERING_RESTART.NEW_SECTION,
            ),
            (
                "w:lnNumType{w:restart=newPage}",
                None,
                None,
                None,
                WD_LINE_NUMBERING_RESTART.NEW_PAGE,
            ),
        ],
    )
    def it_knows_its_attributes(
        self, lnNumType_cxml, count_by, start, distance, restart
    ):
        lnNumType = cast(CT_LineNumber, element(lnNumType_cxml))
        assert lnNumType.countBy == count_by
        assert lnNumType.start == start
        assert lnNumType.distance == distance
        assert lnNumType.restart == restart

    def it_can_set_its_attributes(self):
        lnNumType = cast(CT_LineNumber, element("w:lnNumType"))
        lnNumType.countBy = 3
        lnNumType.start = 2
        lnNumType.distance = Twips(720)
        lnNumType.restart = WD_LINE_NUMBERING_RESTART.NEW_PAGE
        assert lnNumType.xml == xml(
            "w:lnNumType{w:countBy=3,w:start=2,w:distance=720,w:restart=newPage}"
        )


class DescribeCT_SectPr_lnNumType:
    """Unit-test suite for CT_SectPr line-numbering features."""

    def it_returns_None_when_no_lnNumType_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.lnNumType is None

    def it_can_access_its_lnNumType_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:lnNumType{w:countBy=1}"),
        )
        lnNumType = sectPr.lnNumType
        assert lnNumType is not None
        assert lnNumType.countBy == 1

    def it_can_add_a_lnNumType_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        lnNumType = sectPr.get_or_add_lnNumType()
        assert lnNumType is not None
        assert sectPr.lnNumType is lnNumType

    def it_inserts_lnNumType_in_the_right_position(self):
        sectPr = cast(CT_SectPr, element("w:sectPr/(w:pgSz,w:pgMar,w:cols)"))
        sectPr.get_or_add_lnNumType()
        expected = xml("w:sectPr/(w:pgSz,w:pgMar,w:lnNumType,w:cols)")
        assert sectPr.xml == expected

    def it_can_remove_its_lnNumType_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:lnNumType{w:countBy=1}"),
        )
        sectPr._remove_lnNumType()  # pyright: ignore[reportPrivateUsage]
        assert sectPr.lnNumType is None


class DescribeCT_HdrFtr:
    """Unit-test suite for selected units of `docx.oxml.section.CT_HdrFtr`."""

    def it_knows_its_inner_content_block_item_elements(self):
        hdr = cast(CT_HdrFtr, element("w:hdr/(w:tbl,w:tbl,w:p)"))
        assert [type(e) for e in hdr.inner_content_elements] == [CT_Tbl, CT_Tbl, CT_P]
