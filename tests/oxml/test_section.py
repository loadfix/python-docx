"""Unit-test suite for `docx.oxml.section` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.section import (
    WD_BORDER_DISPLAY,
    WD_BORDER_OFFSET_FROM,
    WD_DOC_GRID_TYPE,
    WD_LINE_NUMBERING_RESTART,
    WD_ORIENTATION,
)
from docx.enum.text import WD_BORDER_STYLE
from docx.oxml.section import (
    CT_Col,
    CT_Cols,
    CT_DocGrid,
    CT_HdrFtr,
    CT_LineNumber,
    CT_PaperSource,
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


class DescribeCT_SectPr_orientation_swap:
    """Unit-test suite for CT_SectPr orientation setter w/h swap behavior."""

    def it_swaps_w_and_h_when_changing_portrait_to_landscape(self):
        sectPr = cast(
            CT_SectPr, element("w:sectPr/w:pgSz{w:w=12240,w:h=15840}")
        )

        sectPr.orientation = WD_ORIENTATION.LANDSCAPE

        expected = xml(
            "w:sectPr/w:pgSz{w:w=15840,w:h=12240,w:orient=landscape}"
        )
        assert sectPr.xml == expected

    def it_swaps_w_and_h_when_changing_landscape_to_portrait(self):
        sectPr = cast(
            CT_SectPr,
            element(
                "w:sectPr/w:pgSz{w:w=15840,w:h=12240,w:orient=landscape}"
            ),
        )

        sectPr.orientation = WD_ORIENTATION.PORTRAIT

        # -- orient is dropped (default is portrait), dims are swapped back --
        expected = xml("w:sectPr/w:pgSz{w:w=12240,w:h=15840}")
        assert sectPr.xml == expected

    def it_treats_None_as_portrait_and_swaps_from_landscape(self):
        sectPr = cast(
            CT_SectPr,
            element(
                "w:sectPr/w:pgSz{w:w=15840,w:h=12240,w:orient=landscape}"
            ),
        )

        sectPr.orientation = None

        expected = xml("w:sectPr/w:pgSz{w:w=12240,w:h=15840}")
        assert sectPr.xml == expected

    def it_is_idempotent_when_setting_same_orientation_landscape(self):
        sectPr = cast(
            CT_SectPr,
            element(
                "w:sectPr/w:pgSz{w:w=15840,w:h=12240,w:orient=landscape}"
            ),
        )

        sectPr.orientation = WD_ORIENTATION.LANDSCAPE

        expected = xml(
            "w:sectPr/w:pgSz{w:w=15840,w:h=12240,w:orient=landscape}"
        )
        assert sectPr.xml == expected

    def it_is_idempotent_when_setting_same_orientation_portrait(self):
        sectPr = cast(
            CT_SectPr, element("w:sectPr/w:pgSz{w:w=12240,w:h=15840}")
        )

        sectPr.orientation = WD_ORIENTATION.PORTRAIT

        expected = xml("w:sectPr/w:pgSz{w:w=12240,w:h=15840}")
        assert sectPr.xml == expected

    def it_skips_swap_when_width_is_missing(self):
        sectPr = cast(
            CT_SectPr, element("w:sectPr/w:pgSz{w:h=15840}")
        )

        sectPr.orientation = WD_ORIENTATION.LANDSCAPE

        expected = xml("w:sectPr/w:pgSz{w:h=15840,w:orient=landscape}")
        assert sectPr.xml == expected

    def it_skips_swap_when_height_is_missing(self):
        sectPr = cast(
            CT_SectPr, element("w:sectPr/w:pgSz{w:w=12240}")
        )

        sectPr.orientation = WD_ORIENTATION.LANDSCAPE

        expected = xml("w:sectPr/w:pgSz{w:w=12240,w:orient=landscape}")
        assert sectPr.xml == expected

    def it_skips_swap_when_both_dims_missing(self):
        sectPr = cast(CT_SectPr, element("w:sectPr/w:pgSz"))

        sectPr.orientation = WD_ORIENTATION.LANDSCAPE

        expected = xml("w:sectPr/w:pgSz{w:orient=landscape}")
        assert sectPr.xml == expected

    def it_creates_pgSz_with_no_dims_when_none_present(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))

        sectPr.orientation = WD_ORIENTATION.LANDSCAPE

        expected = xml("w:sectPr/w:pgSz{w:orient=landscape}")
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


class DescribeCT_PaperSource:
    """Unit-test suite for `docx.oxml.section.CT_PaperSource`."""

    @pytest.mark.parametrize(
        ("paperSrc_cxml", "expected_first", "expected_other"),
        [
            ("w:paperSrc", None, None),
            ("w:paperSrc{w:first=1}", 1, None),
            ("w:paperSrc{w:other=2}", None, 2),
            ("w:paperSrc{w:first=3,w:other=4}", 3, 4),
        ],
    )
    def it_knows_its_attributes(self, paperSrc_cxml, expected_first, expected_other):
        paperSrc = cast(CT_PaperSource, element(paperSrc_cxml))
        assert paperSrc.first == expected_first
        assert paperSrc.other == expected_other

    def it_can_set_its_attributes(self):
        paperSrc = cast(CT_PaperSource, element("w:paperSrc"))
        paperSrc.first = 5
        paperSrc.other = 6
        assert paperSrc.xml == xml("w:paperSrc{w:first=5,w:other=6}")


class DescribeCT_SectPr_paperSrc:
    """Unit-test suite for CT_SectPr paper-source features."""

    def it_returns_None_when_no_paperSrc_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.paperSrc is None

    def it_can_access_its_paperSrc_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:paperSrc{w:first=1,w:other=2}"),
        )
        paperSrc = sectPr.paperSrc
        assert paperSrc is not None
        assert paperSrc.first == 1
        assert paperSrc.other == 2

    def it_can_add_a_paperSrc_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        paperSrc = sectPr.get_or_add_paperSrc()
        assert paperSrc is not None
        assert sectPr.paperSrc is paperSrc

    def it_inserts_paperSrc_in_the_right_position(self):
        sectPr = cast(CT_SectPr, element("w:sectPr/(w:pgSz,w:pgMar,w:cols)"))
        sectPr.get_or_add_paperSrc()
        expected = xml("w:sectPr/(w:pgSz,w:pgMar,w:paperSrc,w:cols)")
        assert sectPr.xml == expected

    def it_can_remove_its_paperSrc_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:paperSrc{w:first=1}"),
        )
        sectPr._remove_paperSrc()  # pyright: ignore[reportPrivateUsage]
        assert sectPr.paperSrc is None


class DescribeCT_DocGrid:
    """Unit-test suite for `docx.oxml.section.CT_DocGrid`."""

    @pytest.mark.parametrize(
        ("docGrid_cxml", "grid_type", "line_pitch", "char_space"),
        [
            ("w:docGrid", None, None, None),
            (
                "w:docGrid{w:type=default,w:linePitch=360,w:charSpace=0}",
                WD_DOC_GRID_TYPE.DEFAULT,
                360,
                0,
            ),
            (
                "w:docGrid{w:type=lines,w:linePitch=312}",
                WD_DOC_GRID_TYPE.LINES,
                312,
                None,
            ),
            (
                "w:docGrid{w:type=linesAndChars,w:linePitch=400,w:charSpace=100}",
                WD_DOC_GRID_TYPE.LINES_AND_CHARS,
                400,
                100,
            ),
            (
                "w:docGrid{w:type=snapToChars,w:charSpace=-50}",
                WD_DOC_GRID_TYPE.SNAP_TO_CHARS,
                None,
                -50,
            ),
        ],
    )
    def it_knows_its_attributes(
        self, docGrid_cxml, grid_type, line_pitch, char_space
    ):
        docGrid = cast(CT_DocGrid, element(docGrid_cxml))
        assert docGrid.type == grid_type
        assert docGrid.linePitch == line_pitch
        assert docGrid.charSpace == char_space

    def it_can_set_its_attributes(self):
        docGrid = cast(CT_DocGrid, element("w:docGrid"))
        docGrid.type = WD_DOC_GRID_TYPE.LINES_AND_CHARS
        docGrid.linePitch = 360
        docGrid.charSpace = 100
        assert docGrid.xml == xml(
            "w:docGrid{w:type=linesAndChars,w:linePitch=360,w:charSpace=100}"
        )


class DescribeCT_SectPr_docGrid:
    """Unit-test suite for CT_SectPr document-grid features."""

    def it_returns_None_when_no_docGrid_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.docGrid is None

    def it_can_access_its_docGrid_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:docGrid{w:linePitch=360}"),
        )
        docGrid = sectPr.docGrid
        assert docGrid is not None
        assert docGrid.linePitch == 360

    def it_can_add_a_docGrid_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        docGrid = sectPr.get_or_add_docGrid()
        assert docGrid is not None
        assert sectPr.docGrid is docGrid

    def it_inserts_docGrid_in_the_right_position(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/(w:pgSz,w:pgMar,w:cols,w:titlePg)"),
        )
        sectPr.get_or_add_docGrid()
        expected = xml(
            "w:sectPr/(w:pgSz,w:pgMar,w:cols,w:titlePg,w:docGrid)"
        )
        assert sectPr.xml == expected

    def it_can_remove_its_docGrid_child(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/w:docGrid{w:linePitch=360}"),
        )
        sectPr._remove_docGrid()  # pyright: ignore[reportPrivateUsage]
        assert sectPr.docGrid is None


class DescribeCT_SectPr_text_direction:
    """Unit-test suite for `CT_SectPr.text_direction`."""

    def it_returns_None_when_no_textDirection_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.text_direction is None

    @pytest.mark.parametrize(
        ("sectPr_cxml", "expected_value"),
        [
            ("w:sectPr/w:textDirection{w:val=lrTb}", "LR_TB"),
            ("w:sectPr/w:textDirection{w:val=tbRl}", "TB_RL"),
            ("w:sectPr/w:textDirection{w:val=btLr}", "BT_LR"),
            ("w:sectPr/w:textDirection{w:val=lrTbV}", "LR_TB_V"),
            ("w:sectPr/w:textDirection{w:val=tbRlV}", "TB_RL_V"),
            ("w:sectPr/w:textDirection{w:val=tbLrV}", "TB_LR_V"),
        ],
    )
    def it_knows_its_text_direction(self, sectPr_cxml: str, expected_value: str):
        from docx.enum.table import WD_TEXT_DIRECTION

        sectPr = cast(CT_SectPr, element(sectPr_cxml))
        assert sectPr.text_direction is getattr(WD_TEXT_DIRECTION, expected_value)

    @pytest.mark.parametrize(
        ("enum_member", "xml_val"),
        [
            ("LR_TB", "lrTb"),
            ("TB_RL", "tbRl"),
            ("BT_LR", "btLr"),
            ("LR_TB_V", "lrTbV"),
            ("TB_RL_V", "tbRlV"),
            ("TB_LR_V", "tbLrV"),
        ],
    )
    def it_can_set_its_text_direction_round_trip(
        self, enum_member: str, xml_val: str
    ):
        from docx.enum.table import WD_TEXT_DIRECTION

        sectPr = cast(CT_SectPr, element("w:sectPr"))
        sectPr.text_direction = getattr(WD_TEXT_DIRECTION, enum_member)
        assert sectPr.xml == xml(f"w:sectPr/w:textDirection{{w:val={xml_val}}}")

    def it_can_clear_its_text_direction(self):
        sectPr = cast(
            CT_SectPr, element("w:sectPr/w:textDirection{w:val=tbRl}")
        )
        sectPr.text_direction = None
        assert sectPr.xml == xml("w:sectPr")

    def it_inserts_textDirection_in_the_right_position(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/(w:pgSz,w:pgMar,w:cols,w:titlePg,w:docGrid)"),
        )
        sectPr.get_or_add_textDirection()
        expected = xml(
            "w:sectPr/(w:pgSz,w:pgMar,w:cols,w:titlePg,w:textDirection,w:docGrid)"
        )
        assert sectPr.xml == expected


class DescribeCT_SectPr_bidi:
    """Unit-test suite for `CT_SectPr.bidi_val`."""

    def it_returns_False_when_no_bidi_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.bidi_val is False

    @pytest.mark.parametrize(
        ("sectPr_cxml", "expected_value"),
        [
            ("w:sectPr/w:bidi", True),
            ("w:sectPr/w:bidi{w:val=1}", True),
            ("w:sectPr/w:bidi{w:val=true}", True),
            ("w:sectPr/w:bidi{w:val=on}", True),
            ("w:sectPr/w:bidi{w:val=0}", False),
            ("w:sectPr/w:bidi{w:val=false}", False),
            ("w:sectPr/w:bidi{w:val=off}", False),
        ],
    )
    def it_knows_its_bidi_val(self, sectPr_cxml: str, expected_value: bool):
        sectPr = cast(CT_SectPr, element(sectPr_cxml))
        assert sectPr.bidi_val is expected_value

    @pytest.mark.parametrize(
        ("sectPr_cxml", "value", "expected_cxml"),
        [
            ("w:sectPr", True, "w:sectPr/w:bidi"),
            ("w:sectPr/w:bidi", False, "w:sectPr"),
            ("w:sectPr/w:bidi", None, "w:sectPr"),
            ("w:sectPr/w:bidi{w:val=off}", True, "w:sectPr/w:bidi"),
            ("w:sectPr", False, "w:sectPr"),
        ],
    )
    def it_can_change_its_bidi_val(
        self, sectPr_cxml: str, value: bool | None, expected_cxml: str
    ):
        sectPr = cast(CT_SectPr, element(sectPr_cxml))
        sectPr.bidi_val = value
        assert sectPr.xml == xml(expected_cxml)

    def it_inserts_bidi_in_the_right_position(self):
        sectPr = cast(
            CT_SectPr,
            element(
                "w:sectPr/(w:pgSz,w:pgMar,w:cols,w:titlePg"
                ",w:textDirection{w:val=tbRl},w:docGrid)"
            ),
        )
        sectPr.get_or_add_bidi()
        expected = xml(
            "w:sectPr/(w:pgSz,w:pgMar,w:cols,w:titlePg"
            ",w:textDirection{w:val=tbRl},w:bidi,w:docGrid)"
        )
        assert sectPr.xml == expected


class DescribeCT_SectPr_rtlGutter:
    """Unit-test suite for `CT_SectPr.rtlGutter_val`."""

    def it_returns_False_when_no_rtlGutter_child(self):
        sectPr = cast(CT_SectPr, element("w:sectPr"))
        assert sectPr.rtlGutter_val is False

    @pytest.mark.parametrize(
        ("sectPr_cxml", "expected_value"),
        [
            ("w:sectPr/w:rtlGutter", True),
            ("w:sectPr/w:rtlGutter{w:val=1}", True),
            ("w:sectPr/w:rtlGutter{w:val=true}", True),
            ("w:sectPr/w:rtlGutter{w:val=0}", False),
            ("w:sectPr/w:rtlGutter{w:val=false}", False),
        ],
    )
    def it_knows_its_rtlGutter_val(self, sectPr_cxml: str, expected_value: bool):
        sectPr = cast(CT_SectPr, element(sectPr_cxml))
        assert sectPr.rtlGutter_val is expected_value

    @pytest.mark.parametrize(
        ("sectPr_cxml", "value", "expected_cxml"),
        [
            ("w:sectPr", True, "w:sectPr/w:rtlGutter"),
            ("w:sectPr/w:rtlGutter", False, "w:sectPr"),
            ("w:sectPr/w:rtlGutter", None, "w:sectPr"),
            ("w:sectPr/w:rtlGutter{w:val=off}", True, "w:sectPr/w:rtlGutter"),
            ("w:sectPr", False, "w:sectPr"),
        ],
    )
    def it_can_change_its_rtlGutter_val(
        self, sectPr_cxml: str, value: bool | None, expected_cxml: str
    ):
        sectPr = cast(CT_SectPr, element(sectPr_cxml))
        sectPr.rtlGutter_val = value
        assert sectPr.xml == xml(expected_cxml)

    def it_inserts_rtlGutter_after_bidi_and_before_docGrid(self):
        sectPr = cast(
            CT_SectPr,
            element("w:sectPr/(w:pgSz,w:pgMar,w:cols,w:bidi,w:docGrid)"),
        )
        sectPr.get_or_add_rtlGutter()
        expected = xml(
            "w:sectPr/(w:pgSz,w:pgMar,w:cols,w:bidi,w:rtlGutter,w:docGrid)"
        )
        assert sectPr.xml == expected


class DescribeCT_Cols_sep:
    """Unit-test suite for the `w:sep` attribute on `CT_Cols`."""

    @pytest.mark.parametrize(
        ("cols_cxml", "expected_sep"),
        [
            ("w:cols", None),
            ("w:cols{w:sep=1}", True),
            ("w:cols{w:sep=0}", False),
            ("w:cols{w:sep=true}", True),
            ("w:cols{w:sep=false}", False),
        ],
    )
    def it_reads_the_sep_attribute(self, cols_cxml, expected_sep):
        cols = cast(CT_Cols, element(cols_cxml))
        assert cols.sep is expected_sep

    def it_writes_the_sep_attribute(self):
        cols = cast(CT_Cols, element("w:cols"))
        cols.sep = True
        assert cols.sep is True
        cols.sep = None
        assert cols.sep is None


class DescribeCT_PageSz_code:
    """Unit-test suite for the `w:code` attribute on `CT_PageSz`."""

    def it_returns_None_when_code_is_absent(self):
        from docx.oxml.section import CT_PageSz

        pgSz = cast(CT_PageSz, element("w:pgSz"))
        assert pgSz.code is None

    def it_reads_the_code_attribute(self):
        from docx.oxml.section import CT_PageSz

        pgSz = cast(CT_PageSz, element("w:pgSz{w:code=9}"))
        assert pgSz.code == 9

    def it_writes_the_code_attribute(self):
        from docx.oxml.section import CT_PageSz

        pgSz = cast(CT_PageSz, element("w:pgSz"))
        pgSz.code = 1
        assert pgSz.code == 1
        pgSz.code = None
        assert pgSz.code is None


class DescribeCT_HdrFtr:
    """Unit-test suite for selected units of `docx.oxml.section.CT_HdrFtr`."""

    def it_knows_its_inner_content_block_item_elements(self):
        hdr = cast(CT_HdrFtr, element("w:hdr/(w:tbl,w:tbl,w:p)"))
        assert [type(e) for e in hdr.inner_content_elements] == [CT_Tbl, CT_Tbl, CT_P]
