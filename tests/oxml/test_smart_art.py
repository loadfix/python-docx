# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.oxml.smart_art` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.drawing import CT_Drawing
from docx.oxml.smart_art import (
    CT_Cxn,
    CT_DataModel,
    CT_Pt,
    CT_PtLst,
    CT_RelIds,
    dgm_relIds_from_drawing,
)

from ..unitutil.cxml import element


class DescribeCT_RelIds:
    """Unit test suite for `docx.oxml.smart_art.CT_RelIds`."""

    def it_exposes_its_four_relationship_ids(self):
        relIds = cast(
            CT_RelIds,
            element(
                "dgm:relIds{r:dm=rId4,r:lo=rId5,r:qs=rId6,r:cs=rId7}"
            ),
        )

        assert relIds.dm_rId == "rId4"
        assert relIds.lo_rId == "rId5"
        assert relIds.qs_rId == "rId6"
        assert relIds.cs_rId == "rId7"

    def its_attributes_default_to_None_when_absent(self):
        relIds = cast(CT_RelIds, element("dgm:relIds"))

        assert relIds.dm_rId is None
        assert relIds.lo_rId is None
        assert relIds.qs_rId is None
        assert relIds.cs_rId is None


class DescribeCT_Pt:
    """Unit test suite for `docx.oxml.smart_art.CT_Pt`."""

    def it_knows_its_modelId(self):
        pt = cast(CT_Pt, element("dgm:pt{modelId=abc}"))

        assert pt.modelId == "abc"

    def it_concatenates_run_text_across_a_paragraph(self):
        pt = cast(
            CT_Pt,
            element('dgm:pt/dgm:t/a:p/(a:r/a:t"Hello ",a:r/a:t"World")'),
        )

        assert pt.text == "Hello World"

    def it_joins_multiple_paragraphs_with_newlines(self):
        pt = cast(
            CT_Pt,
            element('dgm:pt/dgm:t/(a:p/a:r/a:t"Line1",a:p/a:r/a:t"Line2")'),
        )

        assert pt.text == "Line1\nLine2"

    def it_returns_empty_string_when_no_dgm_t_child(self):
        pt = cast(CT_Pt, element("dgm:pt{modelId=x}"))

        assert pt.text == ""

    def it_returns_empty_string_when_dgm_t_is_empty(self):
        pt = cast(CT_Pt, element("dgm:pt/dgm:t"))

        assert pt.text == ""


class DescribeCT_DataModel:
    """Unit test suite for `docx.oxml.smart_art.CT_DataModel`."""

    def it_lists_its_pt_children(self):
        dm = cast(
            CT_DataModel,
            element(
                "dgm:dataModel/dgm:ptLst/("
                "dgm:pt{modelId=a},dgm:pt{modelId=b},dgm:pt{modelId=c})"
            ),
        )

        pts = dm.pt_lst

        assert [p.modelId for p in pts] == ["a", "b", "c"]
        assert all(isinstance(p, CT_Pt) for p in pts)

    def it_returns_empty_list_when_no_ptLst(self):
        dm = cast(CT_DataModel, element("dgm:dataModel"))

        assert dm.pt_lst == []

    def it_lists_its_cxn_children(self):
        dm = cast(
            CT_DataModel,
            element(
                "dgm:dataModel/dgm:cxnLst/("
                "dgm:cxn{type=parOf,srcId=a,destId=b},"
                "dgm:cxn{type=parOf,srcId=a,destId=c})"
            ),
        )

        cxns = dm.cxn_lst

        assert len(cxns) == 2
        assert all(isinstance(c, CT_Cxn) for c in cxns)
        assert cxns[0].srcId == "a"
        assert cxns[0].destId == "b"


class DescribeDgmRelIdsFromDrawing:
    """Unit test suite for `docx.oxml.smart_art.dgm_relIds_from_drawing`."""

    def it_finds_an_inline_dgm_relIds(self):
        drawing = cast(
            CT_Drawing,
            element(
                "w:drawing/wp:inline/a:graphic/a:graphicData"
                "/dgm:relIds{r:dm=rId4}"
            ),
        )

        relIds = dgm_relIds_from_drawing(drawing)

        assert isinstance(relIds, CT_RelIds)
        assert relIds.dm_rId == "rId4"

    def it_finds_an_anchor_dgm_relIds(self):
        drawing = cast(
            CT_Drawing,
            element(
                "w:drawing/wp:anchor/a:graphic/a:graphicData"
                "/dgm:relIds{r:dm=rId9}"
            ),
        )

        relIds = dgm_relIds_from_drawing(drawing)

        assert relIds is not None
        assert relIds.dm_rId == "rId9"

    def it_returns_None_when_drawing_is_not_smart_art(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic"),
        )

        assert dgm_relIds_from_drawing(drawing) is None


class DescribeCT_PtLst:
    """Sanity check registration."""

    def it_is_registered(self):
        pt_lst = element("dgm:ptLst")
        assert isinstance(pt_lst, CT_PtLst)
