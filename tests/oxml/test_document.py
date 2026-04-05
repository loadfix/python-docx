"""Unit-test suite for `docx.oxml.document` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.document import CT_Body
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

from ..unitutil.cxml import element


class DescribeCT_Body:
    """Unit-test suite for selected units of `docx.oxml.document.CT_Body`."""

    def it_knows_its_inner_content_block_item_elements(self):
        body = cast(CT_Body, element("w:body/(w:tbl, w:p,w:p)"))
        assert [type(e) for e in body.inner_content_elements] == [CT_Tbl, CT_P, CT_P]

    def it_can_insert_an_element_before_another(self):
        body = cast(CT_Body, element("w:body/(w:p,w:tbl,w:p)"))
        ref_p = body.p_lst[1]  # second paragraph
        new_p = cast(CT_P, element("w:p"))

        result = body.insert_before(new_p, ref_p)

        assert result is new_p
        elements = body.inner_content_elements
        assert len(elements) == 4
        assert [type(e) for e in elements] == [CT_P, CT_Tbl, CT_P, CT_P]
