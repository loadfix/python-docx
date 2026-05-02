# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.oxml.drawing module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.drawing import (
    CT_Drawing,
    CT_GroupShape,
    CT_TxbxContent,
    CT_WordprocessingCanvas,
    CT_WordprocessingShape,
    new_inline_canvas_drawing,
    new_inline_shape_drawing,
)
from docx.oxml.ns import qn
from docx.oxml.shape import CT_Picture

from ..unitutil.cxml import element


class DescribeCT_Drawing:
    """Unit test suite for `docx.oxml.drawing.CT_Drawing` objects."""

    def it_provides_access_to_txbxContent_descendants(self):
        drawing = cast(
            CT_Drawing,
            element(
                "w:drawing/wp:inline/a:graphic/a:graphicData"
                "/wps:wsp/wps:txbx/w:txbxContent/w:p"
            ),
        )

        txbx_contents = drawing.txbxContent_lst

        assert len(txbx_contents) == 1
        assert isinstance(txbx_contents[0], CT_TxbxContent)

    def it_returns_empty_list_when_no_txbxContent(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic"),
        )

        assert drawing.txbxContent_lst == []


class DescribeCT_TxbxContent:
    """Unit test suite for `docx.oxml.drawing.CT_TxbxContent` objects."""

    def it_provides_access_to_its_paragraph_children(self):
        txbxContent = cast(
            CT_TxbxContent,
            element("w:txbxContent/(w:p,w:p)"),
        )

        assert len(txbxContent.p_lst) == 2

    def it_can_get_concatenated_text(self):
        txbxContent = cast(
            CT_TxbxContent,
            element('w:txbxContent/(w:p/w:r/w:t"Hello",w:p/w:r/w:t"World")'),
        )

        assert txbxContent.text == "Hello\nWorld"

    def it_returns_empty_string_when_no_text(self):
        txbxContent = cast(
            CT_TxbxContent,
            element("w:txbxContent/w:p"),
        )

        assert txbxContent.text == ""


class DescribeCT_Drawing_GroupShape:
    """Unit-test suite for `CT_Drawing.grpSp_lst`."""

    def it_finds_an_inline_group_shape(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:inline/a:graphic/a:graphicData/wpg:grpSp"),
        )

        grpSp_lst = drawing.grpSp_lst

        assert len(grpSp_lst) == 1
        assert isinstance(grpSp_lst[0], CT_GroupShape)

    def it_finds_an_anchor_group_shape(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:anchor/a:graphic/a:graphicData/wpg:grpSp"),
        )

        assert len(drawing.grpSp_lst) == 1

    def it_recognizes_legacy_wgp_as_a_group_shape(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:inline/a:graphic/a:graphicData/wpg:wgp"),
        )

        grpSp_lst = drawing.grpSp_lst

        assert len(grpSp_lst) == 1
        assert isinstance(grpSp_lst[0], CT_GroupShape)

    def it_returns_empty_when_drawing_is_not_a_group(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic"),
        )

        assert drawing.grpSp_lst == []


class DescribeCT_GroupShape:
    """Unit-test suite for `docx.oxml.drawing.CT_GroupShape`."""

    def it_reads_name_from_cNvPr(self):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/wpg:nvGrpSpPr/wpg:cNvPr{id=1,name=Group 1}"),
        )

        assert grpSp.name == "Group 1"

    def its_name_is_None_when_nvGrpSpPr_is_missing(self):
        grpSp = cast(CT_GroupShape, element("wpg:grpSp"))

        assert grpSp.name is None

    def its_name_is_None_when_cNvPr_is_missing(self):
        grpSp = cast(CT_GroupShape, element("wpg:grpSp/wpg:nvGrpSpPr"))

        assert grpSp.name is None

    def it_provides_access_to_direct_child_shapes(self):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/(wpg:nvGrpSpPr,wps:wsp,wps:wsp)"),
        )

        wsp_lst = grpSp.wsp_lst

        assert len(wsp_lst) == 2
        assert all(isinstance(w, CT_WordprocessingShape) for w in wsp_lst)

    def it_provides_access_to_nested_groups(self):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/(wpg:nvGrpSpPr,wpg:grpSp,wpg:grpSp)"),
        )

        assert len(grpSp.grpSp_lst) == 2
        assert all(isinstance(g, CT_GroupShape) for g in grpSp.grpSp_lst)

    def it_provides_access_to_child_pictures(self):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/(wpg:nvGrpSpPr,pic:pic)"),
        )

        assert len(grpSp.pic_lst) == 1
        assert isinstance(grpSp.pic_lst[0], CT_Picture)

    def it_iterates_shape_children_in_document_order(self):
        grpSp = cast(
            CT_GroupShape,
            element("wpg:grpSp/(wpg:nvGrpSpPr,wps:wsp,wpg:grpSp,pic:pic,wps:wsp)"),
        )

        children = grpSp.shape_children

        assert [type(c).__name__ for c in children] == [
            "CT_WordprocessingShape",
            "CT_GroupShape",
            "CT_Picture",
            "CT_WordprocessingShape",
        ]


class DescribeCT_WordprocessingShape:
    """Unit-test suite for `docx.oxml.drawing.CT_WordprocessingShape`."""

    def it_reads_name_from_wps_cNvPr(self):
        wsp = cast(
            CT_WordprocessingShape,
            element("wps:wsp/wps:cNvPr{id=1,name=My Shape}"),
        )

        assert wsp.name == "My Shape"

    def its_name_is_None_when_wps_cNvPr_absent(self):
        wsp = cast(CT_WordprocessingShape, element("wps:wsp"))

        assert wsp.name is None

    def it_reads_prst_from_prstGeom(self):
        wsp = cast(
            CT_WordprocessingShape,
            element("wps:wsp/wps:spPr/a:prstGeom{prst=roundRect}"),
        )

        assert wsp.prst == "roundRect"

    def its_prst_is_None_when_absent(self):
        wsp = cast(CT_WordprocessingShape, element("wps:wsp"))

        assert wsp.prst is None

    def it_can_set_text_on_an_empty_shape(self):
        wsp = cast(CT_WordprocessingShape, element("wps:wsp"))

        wsp.set_text("Hello")

        assert wsp.txbx is not None
        assert wsp.txbx.txbxContent is not None
        assert wsp.txbx.txbxContent.text == "Hello"

    def it_replaces_existing_txbx_content_on_set_text(self):
        wsp = cast(
            CT_WordprocessingShape,
            element('wps:wsp/wps:txbx/w:txbxContent/w:p/w:r/w:t"Old"'),
        )

        wsp.set_text("New")

        # -- only one txbx remains --
        assert len(wsp.findall(qn("wps:txbx"))) == 1
        assert wsp.txbx is not None
        assert wsp.txbx.txbxContent is not None
        assert wsp.txbx.txbxContent.text == "New"

    def it_preserves_leading_and_trailing_whitespace_with_xml_space(self):
        wsp = cast(CT_WordprocessingShape, element("wps:wsp"))

        wsp.set_text("  leading and trailing  ")

        assert wsp.txbx is not None
        assert wsp.txbx.txbxContent is not None
        t_elm = wsp.txbx.txbxContent.find(
            f"{qn('w:p')}/{qn('w:r')}/{qn('w:t')}"
        )
        assert t_elm is not None
        assert t_elm.get(qn("xml:space")) == "preserve"


class DescribeNewInlineShapeDrawing:
    """Unit-test suite for `docx.oxml.drawing.new_inline_shape_drawing`."""

    def it_builds_a_drawing_with_the_expected_structure(self):
        drawing = new_inline_shape_drawing(
            prst="rect",
            cx=1828800,
            cy=914400,
            shape_id=1,
            name="Rectangle 1",
        )

        # -- extent populated --
        extent = drawing.find(f"{qn('wp:inline')}/{qn('wp:extent')}")
        assert extent is not None
        assert extent.get("cx") == "1828800"
        assert extent.get("cy") == "914400"

        # -- docPr populated --
        docPr = drawing.find(f"{qn('wp:inline')}/{qn('wp:docPr')}")
        assert docPr is not None
        assert docPr.get("id") == "1"
        assert docPr.get("name") == "Rectangle 1"

        # -- graphicData uri references the wps namespace --
        graphicData = drawing.find(
            f"{qn('wp:inline')}/{qn('a:graphic')}/{qn('a:graphicData')}"
        )
        assert graphicData is not None
        assert (
            graphicData.get("uri")
            == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        )

        # -- wps:wsp is present and carries expected metadata --
        wsp_list = drawing.xpath(".//wps:wsp")
        assert len(wsp_list) == 1
        wsp = wsp_list[0]
        assert isinstance(wsp, CT_WordprocessingShape)
        assert wsp.name == "Rectangle 1"
        assert wsp.prst == "rect"

        # -- xfrm extent populated --
        ext = wsp.find(f"{qn('wps:spPr')}/{qn('a:xfrm')}/{qn('a:ext')}")
        assert ext is not None
        assert ext.get("cx") == "1828800"
        assert ext.get("cy") == "914400"

    @pytest.mark.parametrize(
        "prst",
        ["rect", "roundRect", "ellipse", "rightArrow", "wedgeRoundRectCallout"],
    )
    def it_round_trips_each_supported_preset(self, prst: str):
        drawing = new_inline_shape_drawing(
            prst=prst, cx=100, cy=200, shape_id=5, name="X"
        )

        wsp = drawing.xpath(".//wps:wsp")[0]
        assert wsp.prst == prst

    def it_includes_a_text_frame_when_text_is_provided(self):
        drawing = new_inline_shape_drawing(
            prst="rect", cx=100, cy=200, shape_id=1, name="R", text="Hi"
        )

        wsp = drawing.xpath(".//wps:wsp")[0]
        assert wsp.txbx is not None
        assert wsp.txbx.txbxContent is not None
        assert wsp.txbx.txbxContent.text == "Hi"

    def it_omits_a_text_frame_when_text_is_None(self):
        drawing = new_inline_shape_drawing(
            prst="rect", cx=100, cy=200, shape_id=1, name="R", text=None
        )

        wsp = drawing.xpath(".//wps:wsp")[0]
        assert wsp.txbx is None


class DescribeNewInlineCanvasDrawing:
    """Unit-test suite for `docx.oxml.drawing.new_inline_canvas_drawing`."""

    def it_builds_a_canvas_drawing_with_the_expected_structure(self):
        drawing = new_inline_canvas_drawing(
            cx=5486400, cy=2743200, shape_id=7, name="Canvas 7"
        )

        # -- extent populated --
        extent = drawing.find(f"{qn('wp:inline')}/{qn('wp:extent')}")
        assert extent is not None
        assert extent.get("cx") == "5486400"
        assert extent.get("cy") == "2743200"

        # -- docPr populated --
        docPr = drawing.find(f"{qn('wp:inline')}/{qn('wp:docPr')}")
        assert docPr is not None
        assert docPr.get("id") == "7"
        assert docPr.get("name") == "Canvas 7"

        # -- graphicData uri references the canvas namespace --
        graphicData = drawing.find(
            f"{qn('wp:inline')}/{qn('a:graphic')}/{qn('a:graphicData')}"
        )
        assert graphicData is not None
        assert (
            graphicData.get("uri")
            == "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
        )

        # -- wpc:wpc is present and empty of shapes --
        wpc_list = drawing.xpath(".//wpc:wpc")
        assert len(wpc_list) == 1
        assert isinstance(wpc_list[0], CT_WordprocessingCanvas)
        assert wpc_list[0].wsp_lst == []
