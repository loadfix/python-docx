"""Test suite for docx.oxml.text.parfmt module."""

from __future__ import annotations

import pytest

from docx.enum.text import WD_BORDER_STYLE
from docx.oxml.text.parfmt import CT_Border, CT_PBdr, CT_PPr
from docx.shared import Pt

from ...unitutil.cxml import element, xml


class DescribeCT_Border:
    def it_can_get_border_attributes(self, border_get_fixture):
        border_elm, attr, expected_value = border_get_fixture
        assert getattr(border_elm, attr) == expected_value

    def it_can_set_border_attributes(self, border_set_fixture):
        border_elm, attr, value, expected_xml_ = border_set_fixture
        setattr(border_elm, attr, value)
        assert border_elm.xml == expected_xml_

    @pytest.fixture(
        params=[
            ("w:top", "val", None),
            ("w:top{w:val=single}", "val", WD_BORDER_STYLE.SINGLE),
            ("w:top{w:val=double}", "val", WD_BORDER_STYLE.DOUBLE),
            ("w:top", "sz", None),
            ("w:top{w:sz=8}", "sz", Pt(1)),
            ("w:top{w:sz=24}", "sz", Pt(3)),
            ("w:top", "space", None),
            ("w:top{w:space=1}", "space", Pt(1)),
            ("w:top{w:space=4}", "space", Pt(4)),
            ("w:top", "color", None),
        ]
    )
    def border_get_fixture(self, request):
        cxml, attr, expected_value = request.param
        border_elm = element(cxml)
        return border_elm, attr, expected_value

    @pytest.fixture(
        params=[
            ("w:top", "val", WD_BORDER_STYLE.SINGLE, "w:top{w:val=single}"),
            ("w:top{w:val=single}", "val", WD_BORDER_STYLE.DOUBLE, "w:top{w:val=double}"),
            ("w:top{w:val=single}", "val", None, "w:top"),
            ("w:top", "sz", Pt(1), "w:top{w:sz=8}"),
            ("w:top{w:sz=8}", "sz", Pt(3), "w:top{w:sz=24}"),
            ("w:top{w:sz=8}", "sz", None, "w:top"),
            ("w:top", "space", Pt(1), "w:top{w:space=1}"),
            ("w:top{w:space=1}", "space", Pt(4), "w:top{w:space=4}"),
            ("w:top{w:space=4}", "space", None, "w:top"),
        ]
    )
    def border_set_fixture(self, request):
        cxml, attr, value, expected_cxml = request.param
        border_elm = element(cxml)
        expected_xml_ = xml(expected_cxml)
        return border_elm, attr, value, expected_xml_


class DescribeCT_PBdr:
    def it_can_get_border_children(self, get_fixture):
        pBdr_elm, side, has_child = get_fixture
        child = getattr(pBdr_elm, side)
        if has_child:
            assert child is not None
        else:
            assert child is None

    def it_can_add_border_children(self, add_fixture):
        pBdr_elm, side, expected_xml_ = add_fixture
        getattr(pBdr_elm, f"get_or_add_{side}")()
        assert pBdr_elm.xml == expected_xml_

    @pytest.fixture(
        params=[
            ("w:pBdr", "top", False),
            ("w:pBdr/w:top", "top", True),
            ("w:pBdr", "bottom", False),
            ("w:pBdr/w:bottom", "bottom", True),
            ("w:pBdr", "left", False),
            ("w:pBdr/w:left", "left", True),
            ("w:pBdr", "right", False),
            ("w:pBdr/w:right", "right", True),
            ("w:pBdr", "between", False),
            ("w:pBdr/w:between", "between", True),
        ]
    )
    def get_fixture(self, request):
        cxml, side, has_child = request.param
        pBdr_elm = element(cxml)
        return pBdr_elm, side, has_child

    @pytest.fixture(
        params=[
            ("w:pBdr", "top", "w:pBdr/w:top"),
            ("w:pBdr", "bottom", "w:pBdr/w:bottom"),
            ("w:pBdr", "left", "w:pBdr/w:left"),
            ("w:pBdr", "right", "w:pBdr/w:right"),
            ("w:pBdr", "between", "w:pBdr/w:between"),
            ("w:pBdr/w:top", "bottom", "w:pBdr/(w:top,w:bottom)"),
            ("w:pBdr/w:bottom", "top", "w:pBdr/(w:top,w:bottom)"),
        ]
    )
    def add_fixture(self, request):
        cxml, side, expected_cxml = request.param
        pBdr_elm = element(cxml)
        expected_xml_ = xml(expected_cxml)
        return pBdr_elm, side, expected_xml_


class DescribeCT_PPr_pBdr:
    def it_can_get_pBdr(self, get_fixture):
        pPr_elm, has_pBdr = get_fixture
        if has_pBdr:
            assert pPr_elm.pBdr is not None
        else:
            assert pPr_elm.pBdr is None

    def it_can_add_pBdr(self, add_fixture):
        pPr_elm, expected_xml_ = add_fixture
        pPr_elm.get_or_add_pBdr()
        assert pPr_elm.xml == expected_xml_

    @pytest.fixture(
        params=[
            ("w:pPr", False),
            ("w:pPr/w:pBdr", True),
        ]
    )
    def get_fixture(self, request):
        cxml, has_pBdr = request.param
        pPr_elm = element(cxml)
        return pPr_elm, has_pBdr

    @pytest.fixture(
        params=[
            ("w:pPr", "w:pPr/w:pBdr"),
            ("w:pPr/w:pBdr", "w:pPr/w:pBdr"),
        ]
    )
    def add_fixture(self, request):
        cxml, expected_cxml = request.param
        pPr_elm = element(cxml)
        expected_xml_ = xml(expected_cxml)
        return pPr_elm, expected_xml_
