"""Unit-test suite for `docx.oxml.section` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.section import CT_Col, CT_Cols, CT_HdrFtr, CT_SectPr
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Inches, Twips

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


class DescribeCT_HdrFtr:
    """Unit-test suite for selected units of `docx.oxml.section.CT_HdrFtr`."""

    def it_knows_its_inner_content_block_item_elements(self):
        hdr = cast(CT_HdrFtr, element("w:hdr/(w:tbl,w:tbl,w:p)"))
        assert [type(e) for e in hdr.inner_content_elements] == [CT_Tbl, CT_Tbl, CT_P]
