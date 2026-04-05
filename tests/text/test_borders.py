"""Test suite for docx.text.parfmt module — paragraph borders."""

import pytest

from docx.enum.text import WD_BORDER_STYLE
from docx.shared import Pt, RGBColor
from docx.text.parfmt import Border, ParagraphBorders, ParagraphFormat

from ..unitutil.cxml import element, xml


class DescribeParagraphFormat:
    def it_provides_access_to_its_borders(self):
        p = element("w:p")
        paragraph_format = ParagraphFormat(p)
        borders = paragraph_format.borders
        assert isinstance(borders, ParagraphBorders)

    def it_can_set_a_bottom_border_via_convenience_method(self):
        p = element("w:p")
        paragraph_format = ParagraphFormat(p)
        border = paragraph_format.bottom_border(
            style=WD_BORDER_STYLE.SINGLE,
            width=Pt(1),
            color=RGBColor(0x00, 0x00, 0x00),
            space=Pt(4),
        )
        assert isinstance(border, Border)
        assert border.style == WD_BORDER_STYLE.SINGLE
        assert border.width == Pt(1)
        assert border.color == RGBColor(0x00, 0x00, 0x00)
        assert border.space == Pt(4)

    def it_can_set_a_bottom_border_with_string_color(self):
        p = element("w:p")
        paragraph_format = ParagraphFormat(p)
        border = paragraph_format.bottom_border(
            style=WD_BORDER_STYLE.SINGLE,
            color="FF0000",
        )
        assert border.color == RGBColor(0xFF, 0x00, 0x00)


class DescribeParagraphBorders:
    def it_provides_access_to_each_border_side(self):
        p = element("w:p")
        borders = ParagraphBorders(p)
        for side in ("top", "bottom", "left", "right", "between", "bar"):
            border = getattr(borders, side)
            assert isinstance(border, Border)


class DescribeBorder:
    def it_returns_None_for_style_when_no_border_exists(self):
        p = element("w:p")
        border = Border(p, "bottom")
        assert border.style is None

    def it_returns_None_for_style_when_pPr_exists_but_no_pBdr(self):
        p = element("w:p/w:pPr")
        border = Border(p, "bottom")
        assert border.style is None

    def it_returns_None_for_style_when_pBdr_exists_but_no_side(self):
        p = element("w:p/w:pPr/w:pBdr")
        border = Border(p, "bottom")
        assert border.style is None

    def it_can_get_the_border_style(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single}")
        border = Border(p, "bottom")
        assert border.style == WD_BORDER_STYLE.SINGLE

    def it_can_set_the_border_style(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.style = WD_BORDER_STYLE.DOUBLE
        assert border.style == WD_BORDER_STYLE.DOUBLE

    def it_can_clear_the_border_style(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single}")
        border = Border(p, "bottom")
        border.style = None
        assert border.style is None

    def it_returns_None_for_width_when_no_border_exists(self):
        p = element("w:p")
        border = Border(p, "top")
        assert border.width is None

    def it_can_get_the_border_width(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single,w:sz=8}")
        border = Border(p, "bottom")
        assert border.width == Pt(1)

    def it_can_set_the_border_width(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.width = Pt(2)
        assert border.width == Pt(2)

    def it_returns_None_for_color_when_no_border_exists(self):
        p = element("w:p")
        border = Border(p, "bottom")
        assert border.color is None

    def it_can_get_the_border_color(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single,w:color=FF0000}")
        border = Border(p, "bottom")
        assert border.color == RGBColor(0xFF, 0x00, 0x00)

    def it_can_set_the_border_color(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.color = RGBColor(0x00, 0x00, 0xFF)
        assert border.color == RGBColor(0x00, 0x00, 0xFF)

    def it_returns_None_for_space_when_no_border_exists(self):
        p = element("w:p")
        border = Border(p, "bottom")
        assert border.space is None

    def it_can_get_the_border_space(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single,w:space=4}")
        border = Border(p, "bottom")
        assert border.space == Pt(4)

    def it_can_set_the_border_space(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.space = Pt(8)
        assert border.space == Pt(8)

    def it_does_not_create_an_element_when_setting_width_to_None_on_a_nonexistent_border(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.width = None
        assert p.xml == xml("w:p")

    def it_does_not_create_an_element_when_setting_space_to_None_on_a_nonexistent_border(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.space = None
        assert p.xml == xml("w:p")

    def it_clears_width_on_an_existing_border_when_set_to_None(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single,w:sz=8}")
        border = Border(p, "bottom")
        border.width = None
        assert border.width is None
        assert border.style == WD_BORDER_STYLE.SINGLE

    def it_clears_space_on_an_existing_border_when_set_to_None(self):
        p = element("w:p/w:pPr/w:pBdr/w:bottom{w:val=single,w:space=4}")
        border = Border(p, "bottom")
        border.space = None
        assert border.space is None
        assert border.style == WD_BORDER_STYLE.SINGLE

    def it_works_for_all_sides(self):
        for side in ("top", "bottom", "left", "right", "between", "bar"):
            p = element("w:p")
            border = Border(p, side)
            border.style = WD_BORDER_STYLE.SINGLE
            border.width = Pt(1)
            assert border.style == WD_BORDER_STYLE.SINGLE
            assert border.width == Pt(1)

    def it_can_set_all_border_properties_at_once(self):
        p = element("w:p")
        border = Border(p, "bottom")
        border.style = WD_BORDER_STYLE.SINGLE
        border.width = Pt(1)
        border.color = RGBColor(0x4F, 0x81, 0xBD)
        border.space = Pt(4)
        expected_xml = xml(
            "w:p/w:pPr/w:pBdr/w:bottom{w:val=single,w:sz=8,w:space=4,w:color=4F81BD}"
        )
        assert p.xml == expected_xml


class DescribeCT_PBdr:
    def it_can_add_border_elements(self):
        pBdr = element("w:pBdr")
        bottom = pBdr.get_or_add_bottom()
        assert bottom is not None
        bottom.val = WD_BORDER_STYLE.SINGLE
        assert pBdr.bottom.val == WD_BORDER_STYLE.SINGLE

    def it_preserves_element_order(self):
        pBdr = element("w:pBdr")
        pBdr.get_or_add_bottom()
        pBdr.get_or_add_top()
        # top should come before bottom in XML
        children = list(pBdr)
        assert children[0].tag.endswith("}top")
        assert children[1].tag.endswith("}bottom")
