# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.drawing` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.drawing import Drawing, GroupShape, Picture, WordprocessingShape
from docx.enum.shape import WD_DRAWING_TYPE
from docx.image.image import Image
from docx.oxml.drawing import CT_Drawing, CT_GroupShape
from docx.parts.document import DocumentPart
from docx.parts.image import ImagePart
from docx.text.paragraph import Paragraph

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeDrawing:
    """Unit-test suite for `docx.drawing.Drawing` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic", True),
            ("w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic", True),
            ("w:drawing/wp:inline/a:graphic/a:graphicData/a:grpSp", False),
            ("w:drawing/wp:anchor/a:graphic/a:graphicData/a:chart", False),
        ],
    )
    def it_knows_when_it_contains_a_Picture(
        self, cxml: str, expected_value: bool, document_part_: Mock
    ):
        drawing = Drawing(cast(CT_Drawing, element(cxml)), document_part_)
        assert drawing.has_picture == expected_value

    def it_provides_access_to_the_image_in_a_Picture_drawing(
        self, document_part_: Mock, image_part_: Mock, image_: Mock
    ):
        image_part_.image = image_
        document_part_.part.related_parts = {"rId1": image_part_}
        cxml = (
            "w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip{r:embed=rId1}"
        )
        drawing = Drawing(cast(CT_Drawing, element(cxml)), document_part_)

        image = drawing.image

        assert image is image_

    def but_it_raises_when_the_drawing_does_not_contain_a_Picture(self, document_part_: Mock):
        drawing = Drawing(
            cast(CT_Drawing, element("w:drawing/wp:inline/a:graphic/a:graphicData/a:grpSp")),
            document_part_,
        )

        with pytest.raises(ValueError, match="drawing does not contain a picture"):
            drawing.image

    def it_provides_access_to_text_in_a_text_box(self, document_part_: Mock):
        cxml = (
            "w:drawing/wp:anchor/a:graphic/a:graphicData"
            '/wps:wsp/wps:txbx/w:txbxContent/(w:p/w:r/w:t"Hello",w:p/w:r/w:t"World")'
        )
        drawing = Drawing(cast(CT_Drawing, element(cxml)), document_part_)

        assert drawing.text == "Hello\nWorld"

    def it_returns_empty_text_when_no_text_frames(self, document_part_: Mock):
        drawing = Drawing(
            cast(CT_Drawing, element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic")),
            document_part_,
        )

        assert drawing.text == ""

    def it_provides_access_to_paragraphs_in_a_text_box(self, document_part_: Mock):
        cxml = (
            "w:drawing/wp:anchor/a:graphic/a:graphicData"
            '/wps:wsp/wps:txbx/w:txbxContent/(w:p/w:r/w:t"Hello",w:p/w:r/w:t"World")'
        )
        drawing = Drawing(cast(CT_Drawing, element(cxml)), document_part_)

        paragraphs = drawing.paragraphs

        assert len(paragraphs) == 2
        assert all(isinstance(p, Paragraph) for p in paragraphs)
        assert paragraphs[0].text == "Hello"
        assert paragraphs[1].text == "World"

    def it_returns_empty_paragraphs_when_no_text_frames(self, document_part_: Mock):
        drawing = Drawing(
            cast(CT_Drawing, element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic")),
            document_part_,
        )

        assert drawing.paragraphs == []

    @pytest.mark.parametrize(
        ("cxml", "expected_type"),
        [
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic",
                WD_DRAWING_TYPE.PICTURE,
            ),
            (
                "w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic",
                WD_DRAWING_TYPE.PICTURE,
            ),
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData"
                "/wps:wsp/wps:txbx/w:txbxContent/w:p",
                WD_DRAWING_TYPE.TEXT_BOX,
            ),
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData/c:chart",
                WD_DRAWING_TYPE.CHART,
            ),
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData/wps:wsp",
                WD_DRAWING_TYPE.SHAPE,
            ),
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData/wpg:wgp",
                WD_DRAWING_TYPE.GROUP,
            ),
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData/dgm:relIds",
                WD_DRAWING_TYPE.DIAGRAM,
            ),
        ],
    )
    def it_knows_its_type(
        self, cxml: str, expected_type: WD_DRAWING_TYPE, document_part_: Mock
    ):
        drawing = Drawing(cast(CT_Drawing, element(cxml)), document_part_)

        assert drawing.type == expected_type

    def it_knows_it_is_not_a_group_when_it_wraps_a_picture(self, document_part_: Mock):
        drawing = Drawing(
            cast(CT_Drawing, element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic")),
            document_part_,
        )

        assert drawing.is_group is False
        assert drawing.group_shape is None
        assert drawing.group_shapes == []

    def it_knows_it_is_a_group_when_root_is_grpSp(self, document_part_: Mock):
        cxml = (
            "w:drawing/wp:inline/a:graphic/a:graphicData"
            "/wpg:grpSp/(wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=Group 1},wps:wsp,wps:wsp)"
        )
        drawing = Drawing(cast(CT_Drawing, element(cxml)), document_part_)

        assert drawing.is_group is True

        group_shape = drawing.group_shape
        assert group_shape is not None
        assert isinstance(group_shape, GroupShape)
        assert group_shape.name == "Group 1"

        group_shapes = drawing.group_shapes
        assert len(group_shapes) == 1
        assert isinstance(group_shapes[0], GroupShape)

    def it_recognizes_legacy_wgp_as_a_group(self, document_part_: Mock):
        drawing = Drawing(
            cast(
                CT_Drawing,
                element("w:drawing/wp:inline/a:graphic/a:graphicData/wpg:wgp"),
            ),
            document_part_,
        )

        assert drawing.is_group is True
        assert drawing.group_shape is not None

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def document_part_(self, request: FixtureRequest):
        return instance_mock(request, DocumentPart)

    @pytest.fixture
    def image_(self, request: FixtureRequest):
        return instance_mock(request, Image)

    @pytest.fixture
    def image_part_(self, request: FixtureRequest):
        return instance_mock(request, ImagePart)


class DescribeGroupShape:
    """Unit-test suite for `docx.drawing.GroupShape` proxy."""

    def it_reads_name_from_the_grpSp_element(self, document_part_: Mock):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=My Group}"),
        )
        group = GroupShape(grpSp, document_part_)

        assert group.name == "My Group"

    def its_name_is_None_when_absent(self, document_part_: Mock):
        grpSp = cast(CT_GroupShape, element("wpg:grpSp"))
        group = GroupShape(grpSp, document_part_)

        assert group.name is None

    def it_provides_nested_shapes_in_document_order(self, document_part_: Mock):
        cxml = (
            "wpg:grpSp"
            "/(wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=g},wps:wsp,pic:pic,wps:wsp)"
        )
        group = GroupShape(cast(CT_GroupShape, element(cxml)), document_part_)

        shapes = group.shapes

        assert len(shapes) == 3
        assert isinstance(shapes[0], WordprocessingShape)
        assert isinstance(shapes[1], Picture)
        assert isinstance(shapes[2], WordprocessingShape)

    def it_returns_nested_groups_as_GroupShape(self, document_part_: Mock):
        cxml = (
            "wpg:grpSp"
            "/(wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=outer}"
            ",wpg:grpSp"
            "/(wpg:nvGrpSpPr/wpg:cNvPr{id=2,name=inner},wps:wsp)"
            ")"
        )
        group = GroupShape(cast(CT_GroupShape, element(cxml)), document_part_)

        shapes = group.shapes
        assert len(shapes) == 1
        inner = shapes[0]
        assert isinstance(inner, GroupShape)
        assert inner.name == "inner"

        inner_shapes = inner.shapes
        assert len(inner_shapes) == 1
        assert isinstance(inner_shapes[0], WordprocessingShape)

    def it_returns_empty_shapes_when_group_has_no_children(self, document_part_: Mock):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=g}"),
        )
        group = GroupShape(grpSp, document_part_)

        assert group.shapes == []

    def it_exposes_shape_text_for_textbox_shapes(self, document_part_: Mock):
        cxml = (
            "wpg:grpSp"
            "/(wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=g}"
            ',wps:wsp/wps:txbx/w:txbxContent/(w:p/w:r/w:t"Hi",w:p/w:r/w:t"There"))'
        )
        group = GroupShape(cast(CT_GroupShape, element(cxml)), document_part_)

        shape = group.shapes[0]
        assert isinstance(shape, WordprocessingShape)
        assert shape.text == "Hi\nThere"

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def document_part_(self, request: FixtureRequest):
        return instance_mock(request, DocumentPart)
