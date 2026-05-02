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


class DescribeCT_P_SmartTagTransparency:
    """`CT_P` descends transparently through `w:smartTag` and `w:customXml`.

    See upstream issues #932 and #225 — runs wrapped in smart-tag markup used
    to be silently dropped from `Paragraph.runs` and `CT_P.text`.
    """

    def it_includes_smartTag_wrapped_run_text_in_paragraph_text(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:t>Hello </w:t></w:r>'
            b'<w:smartTag><w:r><w:t>smart</w:t></w:r></w:smartTag>'
            b'<w:r><w:t>!</w:t></w:r>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "Hello smart!"

    def it_includes_customXml_wrapped_run_text_in_paragraph_text(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:t>Hello </w:t></w:r>'
            b'<w:customXml><w:r><w:t>custom</w:t></w:r></w:customXml>'
            b'<w:r><w:t>!</w:t></w:r>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "Hello custom!"

    def it_yields_smartTag_wrapped_runs_from_iter_r_elements(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:t>a</w:t></w:r>'
            b'<w:smartTag><w:r><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r></w:smartTag>'
            b'<w:r><w:t>d</w:t></w:r>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        rs = list(p.iter_r_elements())
        assert [r.text for r in rs] == ["a", "b", "c", "d"]

    def it_descends_recursively_through_nested_smartTags(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:smartTag>'
            b'<w:smartTag><w:r><w:t>nested</w:t></w:r></w:smartTag>'
            b'</w:smartTag>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "nested"
        assert len(list(p.iter_r_elements())) == 1
