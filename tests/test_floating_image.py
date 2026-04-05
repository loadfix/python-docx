# pyright: reportPrivateUsage=false

"""Test suite for floating image (wp:anchor) support."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.drawing import (
    WD_RELATIVE_HORZ_POS,
    WD_RELATIVE_VERT_POS,
    WD_WRAP_TYPE,
)
from docx.oxml.shape import CT_Anchor, CT_Picture, CT_PosH, CT_PosV
from docx.shape import FloatingImage
from docx.shared import Emu, Inches

from .unitutil.cxml import element


class DescribeCT_Anchor:
    """Unit-test suite for `docx.oxml.shape.CT_Anchor`."""

    def it_can_construct_a_new_pic_anchor(self):
        shape_id = 42
        rId = "rId7"
        filename = "test.png"
        cx = Inches(2)
        cy = Inches(1)
        pos_h = 914400
        pos_v = 457200

        anchor = CT_Anchor.new_pic_anchor(
            shape_id=shape_id,
            rId=rId,
            filename=filename,
            cx=cx,
            cy=cy,
            pos_h=pos_h,
            pos_v=pos_v,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.NONE,
        )

        assert anchor.extent.cx == cx
        assert anchor.extent.cy == cy
        assert anchor.docPr.id == shape_id
        assert anchor.docPr.name == "Picture 42"
        assert anchor.behindDoc is False

        # -- verify positioning --
        posH = anchor.positionH
        assert posH is not None
        assert posH.relativeFrom == "column"
        assert posH.posOffset == 914400

        posV = anchor.positionV
        assert posV is not None
        assert posV.relativeFrom == "paragraph"
        assert posV.posOffset == 457200

        # -- verify wrap type --
        assert anchor.wrapNone is not None
        assert anchor.wrap_type == WD_WRAP_TYPE.NONE

    def it_can_construct_an_anchor_with_square_wrap(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.PAGE,
            relative_from_v=WD_RELATIVE_VERT_POS.PAGE,
            wrap_type=WD_WRAP_TYPE.SQUARE,
        )

        assert anchor.wrap_type == WD_WRAP_TYPE.SQUARE
        assert anchor.wrapSquare is not None

    def it_can_construct_an_anchor_behind_doc(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.NONE,
            behind_doc=True,
        )

        assert anchor.behindDoc is True

    @pytest.mark.parametrize(
        ("wrap_type", "expected_wrap_type"),
        [
            (WD_WRAP_TYPE.NONE, WD_WRAP_TYPE.NONE),
            (WD_WRAP_TYPE.SQUARE, WD_WRAP_TYPE.SQUARE),
            (WD_WRAP_TYPE.TIGHT, WD_WRAP_TYPE.TIGHT),
            (WD_WRAP_TYPE.THROUGH, WD_WRAP_TYPE.THROUGH),
            (WD_WRAP_TYPE.TOP_AND_BOTTOM, WD_WRAP_TYPE.TOP_AND_BOTTOM),
        ],
    )
    def it_knows_its_wrap_type(
        self, wrap_type: WD_WRAP_TYPE, expected_wrap_type: WD_WRAP_TYPE
    ):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=wrap_type,
        )
        assert anchor.wrap_type == expected_wrap_type


class DescribeCT_PosH:
    """Unit-test suite for `docx.oxml.shape.CT_PosH`."""

    def it_knows_its_relativeFrom(self):
        posH = cast(
            CT_PosH,
            element('wp:positionH{relativeFrom=column}/wp:posOffset"914400"'),
        )
        assert posH.relativeFrom == "column"

    def it_can_get_and_set_posOffset(self):
        posH = cast(
            CT_PosH,
            element('wp:positionH{relativeFrom=column}/wp:posOffset"0"'),
        )
        assert posH.posOffset == 0
        posH.posOffset = 914400
        assert posH.posOffset == 914400


class DescribeCT_PosV:
    """Unit-test suite for `docx.oxml.shape.CT_PosV`."""

    def it_knows_its_relativeFrom(self):
        posV = cast(
            CT_PosV,
            element('wp:positionV{relativeFrom=paragraph}/wp:posOffset"0"'),
        )
        assert posV.relativeFrom == "paragraph"


class DescribeFloatingImage:
    """Unit-test suite for `docx.shape.FloatingImage`."""

    def it_knows_its_dimensions(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(2),
            cy=Inches(3),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.NONE,
        )
        floating = FloatingImage(anchor)

        assert floating.width == Inches(2)
        assert floating.height == Inches(3)

    def it_can_change_its_dimensions(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(2),
            cy=Inches(3),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.NONE,
        )
        floating = FloatingImage(anchor)

        floating.width = Emu(500000)
        floating.height = Emu(600000)

        assert floating.width == 500000
        assert floating.height == 600000

    def it_knows_its_position(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=914400,
            pos_v=457200,
            relative_from_h=WD_RELATIVE_HORZ_POS.PAGE,
            relative_from_v=WD_RELATIVE_VERT_POS.PAGE,
            wrap_type=WD_WRAP_TYPE.SQUARE,
        )
        floating = FloatingImage(anchor)

        assert floating.pos_h == Emu(914400)
        assert floating.pos_v == Emu(457200)
        assert floating.relative_from_h == WD_RELATIVE_HORZ_POS.PAGE
        assert floating.relative_from_v == WD_RELATIVE_VERT_POS.PAGE

    def it_can_change_its_position(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.NONE,
        )
        floating = FloatingImage(anchor)

        floating.pos_h = Emu(914400)
        floating.pos_v = Emu(457200)
        floating.relative_from_h = WD_RELATIVE_HORZ_POS.PAGE
        floating.relative_from_v = WD_RELATIVE_VERT_POS.PAGE

        assert floating.pos_h == Emu(914400)
        assert floating.pos_v == Emu(457200)
        assert floating.relative_from_h == WD_RELATIVE_HORZ_POS.PAGE
        assert floating.relative_from_v == WD_RELATIVE_VERT_POS.PAGE

    def it_knows_its_wrap_type(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.TIGHT,
        )
        floating = FloatingImage(anchor)
        assert floating.wrap_type == WD_WRAP_TYPE.TIGHT

    def it_knows_its_behind_doc_state(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="img.png",
            cx=Inches(1),
            cy=Inches(1),
            pos_h=0,
            pos_v=0,
            relative_from_h=WD_RELATIVE_HORZ_POS.COLUMN,
            relative_from_v=WD_RELATIVE_VERT_POS.PARAGRAPH,
            wrap_type=WD_WRAP_TYPE.NONE,
            behind_doc=True,
        )
        floating = FloatingImage(anchor)
        assert floating.behind_doc is True

        floating.behind_doc = False
        assert floating.behind_doc is False
