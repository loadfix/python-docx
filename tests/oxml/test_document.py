"""Unit-test suite for `docx.oxml.document` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.document import CT_Background, CT_Body, CT_Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import RGBColor

from ..unitutil.cxml import element, xml


class DescribeCT_Body:
    """Unit-test suite for selected units of `docx.oxml.document.CT_Body`."""

    def it_knows_its_inner_content_block_item_elements(self):
        body = cast(CT_Body, element("w:body/(w:tbl, w:p,w:p)"))
        assert [type(e) for e in body.inner_content_elements] == [CT_Tbl, CT_P, CT_P]


class DescribeCT_Background:
    """Unit-test suite for `docx.oxml.document.CT_Background`."""

    def it_parses_its_color_attribute_as_an_RGBColor(self):
        background = cast(CT_Background, element("w:background{w:color=FF0000}"))
        assert background.color == RGBColor(0xFF, 0x00, 0x00)

    def it_returns_None_for_color_when_attribute_is_absent(self):
        background = cast(CT_Background, element("w:background"))
        assert background.color is None


class DescribeCT_Document:
    """Unit-test suite for `docx.oxml.document.CT_Document`."""

    def it_has_no_background_element_by_default(self):
        document = cast(CT_Document, element("w:document/w:body"))
        assert document.background is None

    def it_exposes_its_background_child_when_present(self):
        document = cast(
            CT_Document, element("w:document/(w:background{w:color=112233},w:body)")
        )
        assert document.background is not None
        assert document.background.color == RGBColor(0x11, 0x22, 0x33)

    def it_inserts_background_before_body(self):
        document = cast(CT_Document, element("w:document/w:body"))

        document.get_or_add_background()

        assert document.xml == xml("w:document/(w:background,w:body)")
