"""Test suite for the docx.oxml.text.paragraph module."""

from typing import cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.text.paragraph import CT_P

from ...unitutil.cxml import element


class DescribeCT_P:
    """Unit-test suite for the CT_P (<w:p>) element."""

    def it_can_add_an_external_hyperlink(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink(rId="rId7", anchor=None, text="Click", rPr=None)

        assert hyperlink.rId == "rId7"
        assert hyperlink.anchor is None
        assert hyperlink.history is True
        rs = hyperlink.r_lst
        assert len(rs) == 1
        assert rs[0].text == "Click"
        assert rs[0].rPr is None

    def it_can_add_an_internal_hyperlink(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink(rId=None, anchor="bookmark1", text="Go", rPr=None)

        assert hyperlink.rId is None
        assert hyperlink.anchor == "bookmark1"
        assert hyperlink.history is True
        assert hyperlink.r_lst[0].text == "Go"

    def it_can_add_a_hyperlink_with_rPr(self):
        p = cast(CT_P, element("w:p"))
        rPr = OxmlElement("w:rPr")
        rStyle = OxmlElement("w:rStyle")
        rStyle.set(qn("w:val"), "Hyperlink")
        rPr.append(rStyle)

        hyperlink = p.add_hyperlink(rId="rId1", anchor=None, text="Link", rPr=rPr)

        r = hyperlink.r_lst[0]
        assert r.rPr is not None
        rStyle_elem = r.rPr.find(qn("w:rStyle"))
        assert rStyle_elem is not None
        assert rStyle_elem.get(qn("w:val")) == "Hyperlink"

    def it_appends_the_hyperlink_as_the_last_child(self):
        p = cast(CT_P, element('w:p/w:r/w:t"existing"'))

        p.add_hyperlink(rId="rId1", anchor=None, text="Link", rPr=None)

        children = list(p)
        assert children[-1].tag == qn("w:hyperlink")
