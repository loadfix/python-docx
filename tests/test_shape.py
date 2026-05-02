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
from docx.oxml.ns import nsmap, qn
from docx.oxml.shape import CT_Anchor, CT_Inline
from docx.shape import (
    EffectsFormat,
    FloatingImage,
    InlineShape,
    InlineShapes,
    PictureCrop,
    PictureOutline,
    ShadowFormat,
    _percent_to_thousandths,
)
from docx.shared import Emu, Length, RGBColor

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

    def it_finds_inline_shapes_inside_mc_AlternateContent_Choice(self, document_: Mock):
        # -- one direct drawing plus one wrapped in mc:AlternateContent/mc:Choice;
        #    the alternate-content shape must also be discovered --
        body_cxml = (
            "w:body/w:p/(w:r/w:drawing/wp:inline,"
            "w:r/mc:AlternateContent/(mc:Choice{Requires=wps}/w:drawing/wp:inline,"
            "mc:Fallback/w:pict))"
        )
        body_elm = cast(CT_Body, element(body_cxml))

        inline_shapes = InlineShapes(body_elm, document_)

        # -- two inline shapes: the direct one and the mc:Choice one --
        assert len(inline_shapes) == 2
        for shape in inline_shapes:
            assert isinstance(shape, InlineShape)

    def it_prefers_mc_Fallback_when_Choice_has_no_inline(self, document_: Mock):
        # -- only mc:Fallback contains an inline drawing; the Choice is empty
        #    (e.g. uses a non-DrawingML alternative). The Fallback should
        #    still be discoverable as an inline shape. --
        body_cxml = (
            "w:body/w:p/w:r/mc:AlternateContent/"
            "(mc:Choice{Requires=wps},"
            "mc:Fallback/w:drawing/wp:inline)"
        )
        body_elm = cast(CT_Body, element(body_cxml))

        inline_shapes = InlineShapes(body_elm, document_)

        assert len(inline_shapes) == 1

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

    # -- opacity / alphaModFix (upstream#1316) --------------------------------

    def it_returns_None_for_opacity_when_absent(self):
        # -- a bare picture inline has no a:alphaModFix in its a:blip --
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        assert InlineShape(inline).opacity is None

    def it_can_set_and_read_opacity(self):
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        shape = InlineShape(inline)

        shape.opacity = 0.25

        # -- 25% as ST_PositivePercentage (1/1000ths of a percent) --
        blip = inline.graphic.graphicData.pic.blipFill.blip
        assert blip.alphaModFix is not None
        assert blip.alphaModFix.amt == 25000
        assert shape.opacity == pytest.approx(0.25)

    def it_clamps_opacity_into_unit_interval(self):
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        shape = InlineShape(inline)

        shape.opacity = 2.0
        assert shape.opacity == 1.0

        shape.opacity = -0.5
        assert shape.opacity == 0.0

    def it_removes_alphaModFix_when_opacity_set_to_None(self):
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        shape = InlineShape(inline)
        shape.opacity = 0.5

        shape.opacity = None

        blip = inline.graphic.graphicData.pic.blipFill.blip
        assert blip.alphaModFix is None
        assert shape.opacity is None

    # -- lock_aspect_ratio (upstream#1314) ------------------------------------

    def it_defaults_lock_aspect_ratio_to_False_when_absent(self):
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        # -- the `pic:cNvPicPr` is present but empty; no `a:picLocks` child --
        assert InlineShape(inline).lock_aspect_ratio is False

    def it_can_release_and_reapply_the_aspect_ratio_lock(self):
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        shape = InlineShape(inline)

        shape.lock_aspect_ratio = False
        cNvPicPr = inline.graphic.graphicData.pic.nvPicPr.cNvPicPr
        assert cNvPicPr.picLocks is not None
        assert cNvPicPr.picLocks.noChangeAspect is False

        shape.lock_aspect_ratio = True
        assert cNvPicPr.picLocks.noChangeAspect is True
        assert shape.lock_aspect_ratio is True

    # -- `.image` read-only property (upstream#249) ---------------------------

    def it_exposes_the_underlying_Image_via_rId(self):
        from docx import Document as OpenDocument

        document = OpenDocument()
        inline_shape = document.add_picture("tests/test_files/python-icon.png")

        image = inline_shape.image

        assert image.content_type == "image/png"
        assert image.filename == "python-icon.png"

    def it_raises_when_image_is_requested_without_a_part(self):
        inline = cast(
            CT_Inline,
            CT_Inline.new_pic_inline(shape_id=1, rId="rId1", filename="f.png", cx=1, cy=1),
        )
        with pytest.raises(ValueError, match="part reference"):
            InlineShape(inline).image


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

    # -- opacity / alphaModFix (upstream#1316) --------------------------------

    def it_can_set_opacity_on_a_floating_picture(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        floating.opacity = 0.75

        assert floating.opacity == pytest.approx(0.75)

    def it_returns_None_for_opacity_when_not_set(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        assert FloatingImage(anchor).opacity is None

    # -- lock_aspect_ratio (upstream#1314) ------------------------------------

    def it_can_toggle_lock_aspect_ratio_on_a_floating_picture(self):
        anchor = CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 100, 100)
        floating = FloatingImage(anchor)

        floating.lock_aspect_ratio = False
        assert floating.lock_aspect_ratio is False

        floating.lock_aspect_ratio = True
        assert floating.lock_aspect_ratio is True


# -- helpers for image-effect tests ----------------------------------------------------


def _inline_picture() -> CT_Inline:
    """Return a freshly-minted `wp:inline` wrapping a `pic:pic`."""
    return CT_Inline.new_pic_inline(1, "rId1", "f.png", 1000, 2000)


def _floating_picture() -> CT_Anchor:
    return CT_Anchor.new_pic_anchor(1, "rId1", "f.png", 1000, 2000)


class Describe_percent_to_thousandths:
    """Unit-test suite for the ``_percent_to_thousandths`` helper."""

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (0.0, 0),
            (0.5, 50000),
            (1.0, 100000),
            (0.25, 25000),
            # -- whole-number percents treated as percents --
            (25, 25000),
            (100, 100000),
            # -- clamping --
            (-0.1, 0),
            (1.5, 100000),
            (-5, 0),
            (250, 100000),
        ],
    )
    def it_accepts_fractional_and_percent_inputs(
        self, value: float, expected: int
    ):
        assert _percent_to_thousandths(value) == expected

    def it_rejects_booleans(self):
        with pytest.raises(TypeError):
            _percent_to_thousandths(True)


class DescribePictureOutline:
    """Unit-test suite for `docx.shape.PictureOutline`."""

    def it_is_None_when_no_ln_element_present(self):
        inline = _inline_picture()
        outline = InlineShape(inline).outline

        assert outline.width is None
        assert outline.color is None
        assert outline.transparency == 0.0

    def it_can_set_and_read_the_outline_width(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.outline.width = Emu(12700)

        assert shape.outline.width == Emu(12700)
        assert isinstance(shape.outline.width, Length)

    def it_can_set_and_read_the_outline_color(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.outline.color = RGBColor(0xFF, 0x00, 0x00)

        assert shape.outline.color == RGBColor(0xFF, 0x00, 0x00)

    def it_accepts_a_hex_string_for_color(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.outline.color = "00FF00"

        assert shape.outline.color == RGBColor(0x00, 0xFF, 0x00)

    def it_can_clear_the_width_and_color(self):
        inline = _inline_picture()
        shape = InlineShape(inline)
        shape.outline.width = Emu(12700)
        shape.outline.color = RGBColor(0xFF, 0, 0)

        shape.outline.width = None
        shape.outline.color = None

        assert shape.outline.width is None
        assert shape.outline.color is None

    def it_can_set_and_read_transparency(self):
        inline = _inline_picture()
        shape = InlineShape(inline)
        shape.outline.color = RGBColor(0, 0, 0)

        shape.outline.transparency = 0.25

        # -- transparency 0.25 means alpha 75000 --
        assert shape.outline.transparency == pytest.approx(0.25, abs=1e-6)
        # -- a:alpha lives inside a:srgbClr --
        ln_xml = shape._inline.graphic.graphicData.pic.spPr.ln.xml
        assert 'val="75000"' in ln_xml

    def it_emits_ln_in_schema_order(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.outline.width = Emu(9525)

        spPr = shape._inline.graphic.graphicData.pic.spPr
        tags = [c.tag for c in spPr]
        # -- a:ln after a:prstGeom --
        assert tags.index(qn("a:ln")) > tags.index(qn("a:prstGeom"))

    def it_raises_when_shape_is_not_a_picture(self):
        inline = cast(
            CT_Inline,
            element(
                'wp:inline/(wp:extent{cx=1,cy=1},wp:docPr{id=1,name=P},'
                "a:graphic/a:graphicData{uri=foo})"
            ),
        )
        with pytest.raises(ValueError):
            InlineShape(inline).outline


class DescribePictureCrop:
    """Unit-test suite for `docx.shape.PictureCrop`."""

    def it_defaults_to_zero_for_all_sides(self):
        inline = _inline_picture()
        crop = InlineShape(inline).crop

        assert crop.left == 0.0
        assert crop.top == 0.0
        assert crop.right == 0.0
        assert crop.bottom == 0.0

    def it_can_set_each_side_independently(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.crop.left = 0.1
        shape.crop.top = 0.2
        shape.crop.right = 0.05
        shape.crop.bottom = 0.15

        assert shape.crop.left == pytest.approx(0.1, abs=1e-6)
        assert shape.crop.top == pytest.approx(0.2, abs=1e-6)
        assert shape.crop.right == pytest.approx(0.05, abs=1e-6)
        assert shape.crop.bottom == pytest.approx(0.15, abs=1e-6)

    def it_accepts_percent_ints_as_well(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.crop.left = 25  # -- 25% --

        assert shape.crop.left == pytest.approx(0.25, abs=1e-6)

    def it_supports_the_set_method(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.crop.set(left=0.1, top=0.2, right=0.05, bottom=0.15)

        src = shape._inline.graphic.graphicData.pic.blipFill.srcRect
        assert src is not None
        assert src.l == 10000
        assert src.t == 20000
        assert src.r == 5000
        assert src.b == 15000

    def it_removes_the_attribute_when_set_to_None_or_zero(self):
        inline = _inline_picture()
        shape = InlineShape(inline)
        shape.crop.left = 0.5

        shape.crop.left = None

        src = shape._inline.graphic.graphicData.pic.blipFill.srcRect
        assert src is None or "l" not in src.attrib

    def it_raises_when_shape_is_not_a_picture(self):
        inline = cast(
            CT_Inline,
            element(
                'wp:inline/(wp:extent{cx=1,cy=1},wp:docPr{id=1,name=P},'
                "a:graphic/a:graphicData{uri=foo})"
            ),
        )
        with pytest.raises(ValueError):
            InlineShape(inline).crop

    def it_works_on_floating_images(self):
        anchor = _floating_picture()
        floating = FloatingImage(anchor)

        floating.crop.left = 0.1

        assert floating.crop.left == pytest.approx(0.1, abs=1e-6)


class DescribeEffectsFormat:
    """Unit-test suite for `docx.shape.EffectsFormat` and `ShadowFormat`."""

    def it_reports_shadow_absent_by_default(self):
        inline = _inline_picture()
        effects = InlineShape(inline).effects

        assert isinstance(effects, EffectsFormat)
        assert isinstance(effects.shadow, ShadowFormat)
        assert effects.shadow.exists is False

    def it_can_apply_a_shadow_with_attributes(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.effects.shadow.apply(
            blur_radius=Emu(38100),
            distance=Emu(12700),
            angle=45,
            color=RGBColor(0x80, 0x80, 0x80),
        )

        shadow = shape.effects.shadow
        assert shadow.exists is True
        assert shadow.blur_radius == Emu(38100)
        assert shadow.distance == Emu(12700)
        assert shadow.angle == pytest.approx(45.0, abs=1e-6)
        assert shadow.color == RGBColor(0x80, 0x80, 0x80)

    def it_emits_outerShdw_in_schema_order(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.effects.shadow.apply(blur_radius=Emu(10000))

        spPr = shape._inline.graphic.graphicData.pic.spPr
        tags = [c.tag for c in spPr]
        assert tags.index(qn("a:effectLst")) > tags.index(qn("a:prstGeom"))
        outerShdw = spPr.effectLst.outerShdw
        assert outerShdw is not None
        assert outerShdw.blurRad == 10000

    def it_can_clear_the_shadow(self):
        inline = _inline_picture()
        shape = InlineShape(inline)
        shape.effects.shadow.apply(distance=Emu(10000))
        assert shape.effects.shadow.exists is True

        shape.effects.shadow.clear()

        assert shape.effects.shadow.exists is False

    def it_wraps_angle_modulo_360(self):
        inline = _inline_picture()
        shape = InlineShape(inline)

        shape.effects.shadow.angle = 405

        assert shape.effects.shadow.angle == pytest.approx(45.0, abs=1e-6)

    def it_can_clear_individual_attributes(self):
        inline = _inline_picture()
        shape = InlineShape(inline)
        shape.effects.shadow.apply(
            blur_radius=Emu(10000),
            distance=Emu(5000),
            angle=45,
            color=RGBColor(0, 0, 0),
        )

        shape.effects.shadow.blur_radius = None
        shape.effects.shadow.color = None

        shadow = shape.effects.shadow
        assert shadow.blur_radius is None
        assert shadow.color is None
        # -- distance and angle preserved --
        assert shadow.distance == Emu(5000)

    def it_works_on_floating_images(self):
        anchor = _floating_picture()
        floating = FloatingImage(anchor)

        floating.effects.shadow.apply(
            blur_radius=Emu(20000), color=RGBColor(0xFF, 0, 0)
        )

        assert floating.effects.shadow.exists is True
        assert floating.effects.shadow.color == RGBColor(0xFF, 0, 0)

    def it_raises_when_shape_is_not_a_picture(self):
        inline = cast(
            CT_Inline,
            element(
                'wp:inline/(wp:extent{cx=1,cy=1},wp:docPr{id=1,name=P},'
                "a:graphic/a:graphicData{uri=foo})"
            ),
        )
        with pytest.raises(ValueError):
            InlineShape(inline).effects


class DescribeFloatingImageOutline:
    """Sanity checks for outline on `FloatingImage` specifically."""

    def it_exposes_an_outline_proxy(self):
        anchor = _floating_picture()
        floating = FloatingImage(anchor)

        floating.outline.width = Emu(9525)
        floating.outline.color = RGBColor(0, 0, 0xFF)

        assert isinstance(floating.outline, PictureOutline)
        assert floating.outline.width == Emu(9525)
        assert floating.outline.color == RGBColor(0, 0, 0xFF)
