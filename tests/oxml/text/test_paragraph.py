# pyright: reportPrivateUsage=false

"""Test suite for the docx.oxml.text.paragraph module."""

from __future__ import annotations

from typing import cast

from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

from ...unitutil.cxml import element, xml


class DescribeCT_P:
    """Unit-test suite for `docx.oxml.text.paragraph.CT_P` objects."""

    def it_can_add_a_p_after_itself(self):
        body = element("w:body/w:p")
        p = cast(CT_P, body[0])

        new_p = p.add_p_after()

        assert new_p.tag == qn("w:p")
        assert p.getnext() is new_p

    def it_can_add_a_tbl_after_itself(self):
        body = element("w:body/w:p")
        p = cast(CT_P, body[0])
        tbl = cast(CT_Tbl, element("w:tbl"))

        result = p.add_tbl_after(tbl)

        assert result is tbl
        assert p.getnext() is tbl
