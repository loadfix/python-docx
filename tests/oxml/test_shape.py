# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.oxml.shape` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.ns import qn
from docx.oxml.shape import (
    CT_Anchor,
    CT_Inline,
    CT_Picture,
    _rot_for_exif_orientation,
)

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

    def it_emits_a_rot_attribute_when_EXIF_orientation_implies_rotation(self):
        """Regression for upstream#540.

        A portrait photo taken with the camera sideways carries an EXIF
        ``Orientation`` of 6, meaning "rotate 90 degrees clockwise to
        view upright". Word honours the value from `a:xfrm/@rot`.
        """
        anchor = CT_Anchor.new_pic_anchor(
            1, "rId1", "f.png", 100, 100, orientation=6
        )

        xfrm = anchor.find(".//" + qn("a:xfrm"))
        assert xfrm is not None
        assert xfrm.get("rot") == str(90 * 60000)

    def it_omits_rot_for_orientation_1_or_None(self):
        for orientation in (None, 1):
            anchor = CT_Anchor.new_pic_anchor(
                2, "rId2", "f.png", 100, 100, orientation=orientation
            )

            xfrm = anchor.find(".//" + qn("a:xfrm"))
            assert xfrm is not None
            assert xfrm.get("rot") is None


class DescribeCT_Picture:
    """Unit-test suite for `docx.oxml.shape.CT_Picture`."""

    def it_always_emits_an_a_xfrm_with_a_ext_on_new(self):
        """Regression for upstream#1164: the non-SVG `_pic_xml()` branch
        used to emit no `<a:xfrm>`, so Word resized the image back to
        default dimensions until the user nudged it."""
        pic = CT_Picture.new(
            pic_id=0, filename="f.png", rId="rId1", cx=1_234_567, cy=7_654_321
        )

        xfrm = pic.find(".//" + qn("a:xfrm"))
        assert xfrm is not None, "expected a:xfrm in pic:spPr"
        ext = xfrm.find(qn("a:ext"))
        assert ext is not None, "expected a:ext as a:xfrm child"
        assert int(ext.get("cx")) == 1_234_567
        assert int(ext.get("cy")) == 7_654_321
        # -- and the off origin child, so Word anchors the xfrm --
        off = xfrm.find(qn("a:off"))
        assert off is not None
        assert off.get("x") == "0"
        assert off.get("y") == "0"

    def it_sets_rot_on_xfrm_for_rotated_EXIF_orientation(self):
        pic = CT_Picture.new(
            pic_id=0,
            filename="f.jpg",
            rId="rId1",
            cx=100,
            cy=100,
            orientation=8,
        )

        xfrm = pic.find(".//" + qn("a:xfrm"))
        assert xfrm is not None
        assert xfrm.get("rot") == str(270 * 60000)

    def it_emits_xfrm_in_svg_branch_with_rot_when_rotated(self):
        pic = CT_Picture.new_svg(
            pic_id=0,
            filename="f.svg",
            fallback_rId="rId1",
            svg_rId="rId2",
            cx=100,
            cy=100,
            orientation=3,
        )

        xfrm = pic.find(".//" + qn("a:xfrm"))
        assert xfrm is not None
        assert xfrm.get("rot") == str(180 * 60000)


class Describe_rot_for_exif_orientation:
    @pytest.mark.parametrize(
        ("orientation", "expected"),
        [
            (None, 0),
            (0, 0),  # out-of-range
            (1, 0),
            (2, 0),
            (3, 180 * 60000),
            (4, 180 * 60000),
            (5, 90 * 60000),
            (6, 90 * 60000),
            (7, 270 * 60000),
            (8, 270 * 60000),
            (9, 0),  # out-of-range
            (42, 0),  # out-of-range
        ],
    )
    def it_maps_EXIF_orientation_to_xfrm_rot_60000ths_of_degree(
        self, orientation: int | None, expected: int
    ):
        assert _rot_for_exif_orientation(orientation) == expected


class DescribeCT_Inline:
    """Unit-test suite for `docx.oxml.shape.CT_Inline`."""

    def it_threads_orientation_into_pic_xfrm_rot(self):
        """End-to-end: calling `new_pic_inline` with orientation=6 should
        produce a drawing whose `a:xfrm/@rot` is 5_400_000 (90deg)."""
        inline = CT_Inline.new_pic_inline(
            1, "rId1", "f.jpg", 100, 100, orientation=6
        )

        xfrm = inline.find(".//" + qn("a:xfrm"))
        assert xfrm is not None
        assert xfrm.get("rot") == str(90 * 60000)
