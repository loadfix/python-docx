"""Unit-test suite for `docx.oxml.shape` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.shape import CT_ShapeProperties, CT_Transform2D
from docx.shared import Emu

from ..unitutil.cxml import element


class DescribeCT_ShapeProperties:
    """Unit-test suite for `docx.oxml.shape.CT_ShapeProperties`."""

    def it_returns_None_for_cx_when_xfrm_is_absent(self):
        spPr = cast(CT_ShapeProperties, element("pic:spPr"))
        assert spPr.cx is None

    def it_returns_None_for_cy_when_xfrm_is_absent(self):
        spPr = cast(CT_ShapeProperties, element("pic:spPr"))
        assert spPr.cy is None

    def it_provides_cx_when_xfrm_is_present(self):
        spPr = cast(
            CT_ShapeProperties,
            element("pic:spPr/a:xfrm/a:ext{cx=914400,cy=457200}"),
        )
        assert spPr.cx == 914400

    def it_provides_cy_when_xfrm_is_present(self):
        spPr = cast(
            CT_ShapeProperties,
            element("pic:spPr/a:xfrm/a:ext{cx=914400,cy=457200}"),
        )
        assert spPr.cy == 457200

    def it_can_set_cx(self):
        spPr = cast(CT_ShapeProperties, element("pic:spPr"))
        spPr.cx = Emu(914400)
        assert spPr.cx == 914400

    def it_can_set_cy(self):
        spPr = cast(CT_ShapeProperties, element("pic:spPr"))
        spPr.cy = Emu(457200)
        assert spPr.cy == 457200


class DescribeCT_Transform2D:
    """Unit-test suite for `docx.oxml.shape.CT_Transform2D`."""

    def it_returns_None_for_cx_when_ext_is_absent(self):
        xfrm = cast(CT_Transform2D, element("a:xfrm"))
        assert xfrm.cx is None

    def it_returns_None_for_cy_when_ext_is_absent(self):
        xfrm = cast(CT_Transform2D, element("a:xfrm"))
        assert xfrm.cy is None

    def it_provides_cx_when_ext_is_present(self):
        xfrm = cast(CT_Transform2D, element("a:xfrm/a:ext{cx=100,cy=200}"))
        assert xfrm.cx == 100

    def it_provides_cy_when_ext_is_present(self):
        xfrm = cast(CT_Transform2D, element("a:xfrm/a:ext{cx=100,cy=200}"))
        assert xfrm.cy == 200

    def it_can_set_cx(self):
        xfrm = cast(CT_Transform2D, element("a:xfrm"))
        xfrm.cx = Emu(500)
        assert xfrm.cx == 500

    def it_can_set_cy(self):
        xfrm = cast(CT_Transform2D, element("a:xfrm"))
        xfrm.cy = Emu(600)
        assert xfrm.cy == 600
