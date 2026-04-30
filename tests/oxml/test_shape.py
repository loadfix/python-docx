# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.oxml.shape` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.ns import qn
from docx.oxml.shape import CT_Anchor

from ..unitutil.cxml import element


class DescribeCT_Anchor:
    """Unit-test suite for `docx.oxml.shape.CT_Anchor`."""

    def it_can_construct_a_new_pic_anchor(self):
        anchor = CT_Anchor.new_pic_anchor(
            shape_id=42, rId="rId7", filename="foo.png", cx=1000, cy=2000
        )

        # -- required attributes present --
        assert anchor.simplePos is False
        assert anchor.locked is False
        assert anchor.layoutInCell is True
        assert anchor.allowOverlap is True
        assert anchor.behindDoc is False
        assert anchor.relativeHeight == 0

        # -- default positioning --
        assert anchor.positionH is not None
        assert anchor.positionH.relativeFrom == "column"
        assert anchor.positionV is not None
        assert anchor.positionV.relativeFrom == "paragraph"

        # -- default wrap --
        assert anchor.wrap_type == "square"

        # -- extent populated --
        assert anchor.extent.cx == 1000
        assert anchor.extent.cy == 2000

        # -- docPr populated --
        assert anchor.docPr.id == 42
        assert anchor.docPr.name == "Picture 42"

        # -- graphic contains pic:pic with our rId --
        pic = anchor.graphic.graphicData.find(qn("pic:pic"))
        assert pic is not None
        blip = pic.find(".//" + qn("a:blip"))
        assert blip is not None
        assert blip.get(qn("r:embed")) == "rId7"

    def it_can_set_the_horizontal_position(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        anchor.set_horizontal_position("page", 914400)

        assert anchor.positionH.relativeFrom == "page"
        assert anchor.positionH.posOffset is not None
        assert anchor.positionH.posOffset.text == "914400"

    def it_can_set_the_vertical_position(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        anchor.set_vertical_position("margin", 457200)

        assert anchor.positionV.relativeFrom == "margin"
        assert anchor.positionV.posOffset is not None
        assert anchor.positionV.posOffset.text == "457200"

    @pytest.mark.parametrize(
        ("wrap", "tag", "expected_behind_doc"),
        [
            ("square", "wp:wrapSquare", False),
            ("tight", "wp:wrapTight", False),
            ("through", "wp:wrapThrough", False),
            ("topAndBottom", "wp:wrapTopAndBottom", False),
            ("behind", "wp:wrapNone", True),
            ("inFront", "wp:wrapNone", False),
        ],
    )
    def it_can_set_the_wrap_type(
        self, wrap: str, tag: str, expected_behind_doc: bool
    ):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        anchor.set_wrap(wrap)

        assert anchor.find(qn(tag)) is not None
        assert anchor.behindDoc is expected_behind_doc
        assert anchor.wrap_type == wrap

    def it_replaces_existing_wrap_when_set_wrap_is_called(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        anchor.set_wrap("tight")
        anchor.set_wrap("behind")

        # -- only one wrap element present at a time --
        wrap_elms = [
            tag
            for tag in (
                "wp:wrapNone",
                "wp:wrapSquare",
                "wp:wrapTight",
                "wp:wrapThrough",
                "wp:wrapTopAndBottom",
            )
            if anchor.find(qn(tag)) is not None
        ]
        assert wrap_elms == ["wp:wrapNone"]
        assert anchor.behindDoc is True

    def it_reports_square_wrap_type_by_default(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        assert anchor.wrap_type == "square"

    def it_reads_horizontal_and_vertical_relativeFrom_from_existing_xml(self):
        cxml = (
            "wp:anchor{distT=0,distB=0,distL=0,distR=0,simplePos=0,relativeHeight=1,"
            "behindDoc=0,locked=0,layoutInCell=1,allowOverlap=1}/"
            "(wp:positionH{relativeFrom=page},wp:positionV{relativeFrom=margin})"
        )
        anchor = cast(CT_Anchor, element(cxml))

        assert anchor.positionH.relativeFrom == "page"
        assert anchor.positionV.relativeFrom == "margin"
