# pyright: reportPrivateUsage=false

"""Test suite for the docx.shape module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.document import Document
from docx.enum.shape import (
    WD_ANCHOR_H,
    WD_ANCHOR_V,
    WD_INLINE_SHAPE,
    WD_WRAP_TYPE,
)
from docx.oxml.document import CT_Body
from docx.oxml.ns import nsmap
from docx.oxml.shape import CT_Anchor, CT_Inline
from docx.shape import FloatingImage, InlineShape, InlineShapes
from docx.shared import Emu, Length

from .unitutil.cxml import element, xml
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeInlineShapes:
    """Unit-test suite for `docx.shape.InlineShapes` objects."""

    def it_knows_how_many_inline_shapes_it_contains(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        assert len(inline_shapes) == 2

    def it_can_iterate_over_its_InlineShape_instances(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        assert all(isinstance(s, InlineShape) for s in inline_shapes)
        assert len(list(inline_shapes)) == 2

    def it_provides_indexed_access_to_inline_shapes(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        for idx in range(-2, 2):
            assert isinstance(inline_shapes[idx], InlineShape)

    def it_raises_on_indexed_access_out_of_range(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)

        with pytest.raises(IndexError, match=r"inline shape index \[-3\] out of range"):
            inline_shapes[-3]
        with pytest.raises(IndexError, match=r"inline shape index \[2\] out of range"):
            inline_shapes[2]

    def it_knows_the_part_it_belongs_to(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        assert inline_shapes.part is document_.part

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def body(self) -> CT_Body:
        return cast(
            CT_Body, element("w:body/w:p/(w:r/w:drawing/wp:inline, w:r/w:drawing/wp:inline)")
        )

    @pytest.fixture
    def document_(self, request: FixtureRequest):
        return instance_mock(request, Document)


class DescribeInlineShape:
    """Unit-test suite for `docx.shape.InlineShape` objects."""

    @pytest.mark.parametrize(
        ("uri", "content_cxml", "expected_value"),
        [
            # -- embedded picture --
            (nsmap["pic"], "/pic:pic/pic:blipFill/a:blip{r:embed=rId1}", WD_INLINE_SHAPE.PICTURE),
            # -- linked picture --
            (
                nsmap["pic"],
                "/pic:pic/pic:blipFill/a:blip{r:link=rId2}",
                WD_INLINE_SHAPE.LINKED_PICTURE,
            ),
            # -- linked and embedded picture (not expected) --
            (
                nsmap["pic"],
                "/pic:pic/pic:blipFill/a:blip{r:embed=rId1,r:link=rId2}",
                WD_INLINE_SHAPE.LINKED_PICTURE,
            ),
            # -- chart --
            (nsmap["c"], "", WD_INLINE_SHAPE.CHART),
            # -- SmartArt --
            (nsmap["dgm"], "", WD_INLINE_SHAPE.SMART_ART),
            # -- something else we don't know about --
            ("foobar", "", WD_INLINE_SHAPE.NOT_IMPLEMENTED),
        ],
    )
    def it_knows_what_type_of_shape_it_is(
        self, uri: str, content_cxml: str, expected_value: WD_INLINE_SHAPE
    ):
        cxml = "wp:inline/a:graphic/a:graphicData{uri=%s}%s" % (uri, content_cxml)
        inline = cast(CT_Inline, element(cxml))
        inline_shape = InlineShape(inline)
        assert inline_shape.type == expected_value

    def it_knows_its_display_dimensions(self):
        inline = cast(CT_Inline, element("wp:inline/wp:extent{cx=333, cy=666}"))
        inline_shape = InlineShape(inline)

        width, height = inline_shape.width, inline_shape.height

        assert isinstance(width, Length)
        assert width == 333
        assert isinstance(height, Length)
        assert height == 666

    def it_can_change_its_display_dimensions(self):
        inline_shape = InlineShape(
            cast(
                CT_Inline,
                element(
                    "wp:inline/(wp:extent{cx=333,cy=666},a:graphic/a:graphicData/pic:pic/"
                    "pic:spPr/a:xfrm/a:ext{cx=333,cy=666})"
                ),
            )
        )

        inline_shape.width = Emu(444)
        inline_shape.height = Emu(888)

        assert inline_shape._inline.xml == xml(
            "wp:inline/(wp:extent{cx=444,cy=888},a:graphic/a:graphicData/pic:pic/pic:spPr/"
            "a:xfrm/a:ext{cx=444,cy=888})"
        )

    def it_returns_None_for_alt_text_when_absent(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))

        assert InlineShape(inline).alt_text is None

    def it_returns_the_alt_text_when_present(self):
        inline = cast(
            CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1,descr=a cat}")
        )

        assert InlineShape(inline).alt_text == "a cat"

    def it_can_set_the_alt_text(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        inline_shape = InlineShape(inline)

        inline_shape.alt_text = "a cat"

        assert inline_shape._inline.docPr.descr == "a cat"

    def it_removes_the_alt_text_when_set_to_None(self):
        inline = cast(
            CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1,descr=a cat}")
        )
        inline_shape = InlineShape(inline)

        inline_shape.alt_text = None

        assert inline_shape._inline.docPr.descr is None
        assert "descr" not in inline_shape._inline.docPr.attrib

    def it_returns_None_for_title_when_absent(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))

        assert InlineShape(inline).title is None

    def it_returns_the_title_when_present(self):
        inline = cast(
            CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1,title=Kitty}")
        )

        assert InlineShape(inline).title == "Kitty"

    def it_can_set_the_title(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        inline_shape = InlineShape(inline)

        inline_shape.title = "Kitty"

        assert inline_shape._inline.docPr.title == "Kitty"

    def it_removes_the_title_when_set_to_None(self):
        inline = cast(
            CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1,title=Kitty}")
        )
        inline_shape = InlineShape(inline)

        inline_shape.title = None

        assert inline_shape._inline.docPr.title is None
        assert "title" not in inline_shape._inline.docPr.attrib


class DescribeFloatingImage:
    """Unit-test suite for `docx.shape.FloatingImage`."""

    def it_knows_its_display_dimensions(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 4000, 8000)
        floating = FloatingImage(anchor)

        width, height = floating.width, floating.height

        assert isinstance(width, Length)
        assert width == 4000
        assert isinstance(height, Length)
        assert height == 8000

    def it_exposes_default_horizontal_and_vertical_anchors(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        assert floating.horizontal_anchor == WD_ANCHOR_H.COLUMN
        assert floating.vertical_anchor == WD_ANCHOR_V.PARAGRAPH

    def it_exposes_horizontal_and_vertical_offsets_as_Emu(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.set_horizontal_position("page", 914400)
        anchor.set_vertical_position("margin", 457200)
        floating = FloatingImage(anchor)

        assert floating.horizontal_offset == Emu(914400)
        assert floating.vertical_offset == Emu(457200)
        assert floating.offset == (Emu(914400), Emu(457200))

    def it_returns_zero_offset_when_not_specified(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        # -- default XML has posOffset of 0 --
        assert floating.horizontal_offset == 0
        assert floating.vertical_offset == 0

    @pytest.mark.parametrize(
        ("wrap_str", "expected"),
        [
            ("square", WD_WRAP_TYPE.SQUARE),
            ("tight", WD_WRAP_TYPE.TIGHT),
            ("through", WD_WRAP_TYPE.THROUGH),
            ("topAndBottom", WD_WRAP_TYPE.TOP_AND_BOTTOM),
            ("behind", WD_WRAP_TYPE.BEHIND),
            ("inFront", WD_WRAP_TYPE.IN_FRONT),
        ],
    )
    def it_exposes_its_wrap_type(self, wrap_str: str, expected: WD_WRAP_TYPE):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.set_wrap(wrap_str)
        floating = FloatingImage(anchor)

        assert floating.wrap_type == expected

    def it_exposes_a_position_dict(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.set_horizontal_position("page", 100)
        anchor.set_vertical_position("margin", 200)
        floating = FloatingImage(anchor)

        position = floating.position

        assert position["h_anchor"] == WD_ANCHOR_H.PAGE
        assert position["v_anchor"] == WD_ANCHOR_V.MARGIN
        assert position["horizontal"] == 100
        assert position["vertical"] == 200

    def it_reports_picture_type_for_a_picture_anchor(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        assert floating.type == WD_INLINE_SHAPE.PICTURE

    def it_returns_None_for_alt_text_when_absent(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        assert FloatingImage(anchor).alt_text is None

    def it_returns_the_alt_text_when_present(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.docPr.descr = "a cat"

        assert FloatingImage(anchor).alt_text == "a cat"

    def it_can_set_the_alt_text(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        floating.alt_text = "a cat"

        assert anchor.docPr.descr == "a cat"

    def it_removes_the_alt_text_when_set_to_None(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.docPr.descr = "a cat"
        floating = FloatingImage(anchor)

        floating.alt_text = None

        assert anchor.docPr.descr is None
        assert "descr" not in anchor.docPr.attrib

    def it_returns_None_for_title_when_absent(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)

        assert FloatingImage(anchor).title is None

    def it_returns_the_title_when_present(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.docPr.title = "Kitty"

        assert FloatingImage(anchor).title == "Kitty"

    def it_can_set_the_title(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        floating.title = "Kitty"

        assert anchor.docPr.title == "Kitty"

    def it_removes_the_title_when_set_to_None(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        anchor.docPr.title = "Kitty"
        floating = FloatingImage(anchor)

        floating.title = None

        assert anchor.docPr.title is None
        assert "title" not in anchor.docPr.attrib
