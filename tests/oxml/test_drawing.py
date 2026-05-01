# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.oxml.drawing module."""

from __future__ import annotations

from typing import cast

from docx.oxml.drawing import (
    CT_Drawing,
    CT_GroupShape,
    CT_TxbxContent,
    CT_WordprocessingShape,
)
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
