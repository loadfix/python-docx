# pyright: reportPrivateUsage=false

"""Test suite for floating image (wp:anchor) support."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.shape import WD_RELATIVE_HORZ_POS, WD_RELATIVE_VERT_POS, WD_WRAP_TYPE
from docx.oxml.shape import CT_Anchor, CT_Inline, CT_Picture
from docx.shape import FloatingImage
from docx.shared import Emu, Length

from .unitutil.cxml import element, xml


class DescribeCT_Anchor:
    """Unit-test suite for `docx.oxml.shape.CT_Anchor` objects."""

    def it_can_construct_a_new_anchor_element(self):
        pic = CT_Picture.new(0, "test.png", "rId1", Emu(914400), Emu(914400))
        anchor = CT_Anchor.new(Emu(914400), Emu(914400), 1, pic)

        assert anchor.extent.cx == 914400
        assert anchor.extent.cy == 914400
        assert anchor.docPr.id == 1
        assert anchor.docPr.name == "Picture 1"

    def it_can_construct_a_new_pic_anchor(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            horz_offset=457200,
            vert_offset=228600,
            horz_relative_from="page",
            vert_relative_from="margin",
            wrap_type="square",
            behind_doc=False,
        )

        assert anchor.extent.cx == 914400
        assert anchor.extent.cy == 457200
        assert anchor.horz_offset == 457200
        assert anchor.vert_offset == 228600
        assert anchor.horz_relative_from == "page"
        assert anchor.vert_relative_from == "margin"
        assert anchor.wrap_type_str == "square"
        assert anchor.behind_doc is False

    def it_can_construct_a_behind_doc_anchor(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            behind_doc=True,
        )

        assert anchor.behind_doc is True

    def it_defaults_to_wrap_none(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
        )

        assert anchor.wrap_type_str == "none"

    @pytest.mark.parametrize(
        "wrap_type",
        ["none", "square", "tight", "through", "topAndBottom"],
    )
    def it_supports_all_wrap_types(self, wrap_type: str):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            wrap_type=wrap_type,
        )

        assert anchor.wrap_type_str == wrap_type


class DescribeFloatingImage:
    """Unit-test suite for `docx.shape.FloatingImage` objects."""

    def it_knows_its_display_dimensions(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
        )
        floating_image = FloatingImage(anchor)

        assert isinstance(floating_image.width, Length)
        assert floating_image.width == 914400
        assert isinstance(floating_image.height, Length)
        assert floating_image.height == 457200

    def it_can_change_its_display_dimensions(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
        )
        floating_image = FloatingImage(anchor)

        floating_image.width = Emu(1828800)
        floating_image.height = Emu(914400)

        assert floating_image.width == 1828800
        assert floating_image.height == 914400

    def it_knows_its_wrap_type(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            wrap_type="square",
        )
        floating_image = FloatingImage(anchor)

        assert floating_image.wrap_type == WD_WRAP_TYPE.SQUARE

    def it_reports_behind_doc_as_behind_wrap_type(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            behind_doc=True,
        )
        floating_image = FloatingImage(anchor)

        assert floating_image.wrap_type == WD_WRAP_TYPE.BEHIND

    def it_reports_not_behind_doc_none_as_in_front(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            wrap_type="none",
            behind_doc=False,
        )
        floating_image = FloatingImage(anchor)

        assert floating_image.wrap_type == WD_WRAP_TYPE.IN_FRONT

    def it_knows_its_position_offsets(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            horz_offset=457200,
            vert_offset=228600,
        )
        floating_image = FloatingImage(anchor)

        assert floating_image.horz_offset == 457200
        assert floating_image.vert_offset == 228600

    def it_knows_its_position_references(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            horz_relative_from="page",
            vert_relative_from="margin",
        )
        floating_image = FloatingImage(anchor)

        assert floating_image.horz_pos_relative == WD_RELATIVE_HORZ_POS.PAGE
        assert floating_image.vert_pos_relative == WD_RELATIVE_VERT_POS.MARGIN

    def it_can_change_its_position_offsets(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            horz_offset=100000,
            vert_offset=200000,
        )
        floating_image = FloatingImage(anchor)

        floating_image.horz_offset = 300000
        floating_image.vert_offset = 400000

        assert floating_image.horz_offset == 300000
        assert floating_image.vert_offset == 400000

    def it_can_change_its_position_references(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            horz_relative_from="column",
            vert_relative_from="paragraph",
        )
        floating_image = FloatingImage(anchor)

        floating_image.horz_pos_relative = WD_RELATIVE_HORZ_POS.PAGE
        floating_image.vert_pos_relative = WD_RELATIVE_VERT_POS.MARGIN

        assert floating_image.horz_pos_relative == WD_RELATIVE_HORZ_POS.PAGE
        assert floating_image.vert_pos_relative == WD_RELATIVE_VERT_POS.MARGIN

    @pytest.mark.parametrize(
        ("wrap_type_str", "behind_doc", "expected"),
        [
            ("none", False, WD_WRAP_TYPE.IN_FRONT),
            ("none", True, WD_WRAP_TYPE.BEHIND),
            ("square", False, WD_WRAP_TYPE.SQUARE),
            ("tight", False, WD_WRAP_TYPE.TIGHT),
            ("through", False, WD_WRAP_TYPE.THROUGH),
            ("topAndBottom", False, WD_WRAP_TYPE.TOP_AND_BOTTOM),
        ],
    )
    def it_maps_all_wrap_types_correctly(
        self, wrap_type_str: str, behind_doc: bool, expected: WD_WRAP_TYPE
    ):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=1,
            rId="rId1",
            filename="test.png",
            cx=Emu(914400),
            cy=Emu(457200),
            wrap_type=wrap_type_str,
            behind_doc=behind_doc,
        )
        floating_image = FloatingImage(anchor)

        assert floating_image.wrap_type == expected
