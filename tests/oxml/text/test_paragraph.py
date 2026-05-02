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


class DescribeCT_P_TransparentWrapperExpansion:
    """Phase A-v2 #1: descend w:sdt/mc:AlternateContent/w:ins/w:moveTo.

    See upstream #1327, #1389, #335, PR#1538, PR#734.
    """

    def it_descends_into_w_ins_for_paragraph_text(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:t>before </w:t></w:r>'
            b'<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            b'<w:r><w:t>inserted</w:t></w:r>'
            b'</w:ins>'
            b'<w:r><w:t> after</w:t></w:r>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "before inserted after"

    def it_descends_into_w_moveTo_for_paragraph_text(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:t>X </w:t></w:r>'
            b'<w:moveTo w:id="2" w:author="A" w:date="2024-01-01T00:00:00Z">'
            b'<w:r><w:t>moved</w:t></w:r>'
            b'</w:moveTo>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "X moved"

    def it_yields_ins_and_moveTo_runs_from_iter_r_elements(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            b'<w:r><w:t>ins1</w:t></w:r>'
            b'</w:ins>'
            b'<w:moveTo w:id="2" w:author="A" w:date="2024-01-01T00:00:00Z">'
            b'<w:r><w:t>mv</w:t></w:r>'
            b'</w:moveTo>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        rs = list(p.iter_r_elements())
        assert [r.text for r in rs] == ["ins1", "mv"]

    def it_descends_mc_AlternateContent_preferring_Choice(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            b' xmlns:mc='
            b'"http://schemas.openxmlformats.org/markup-compatibility/2006">'
            b'<mc:AlternateContent>'
            b'<mc:Choice Requires="w14"><w:r><w:t>choice</w:t></w:r></mc:Choice>'
            b'<mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback>'
            b'</mc:AlternateContent>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "choice"

    def it_falls_back_when_Choice_has_no_run_like_content(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            b' xmlns:mc='
            b'"http://schemas.openxmlformats.org/markup-compatibility/2006">'
            b'<mc:AlternateContent>'
            b'<mc:Choice Requires="w14"/>'
            b'<mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback>'
            b'</mc:AlternateContent>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "fallback"

    def it_descends_sdt_text_for_paragraph_text(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:t>before </w:t></w:r>'
            b'<w:sdt>'
            b'<w:sdtPr/>'
            b'<w:sdtContent><w:r><w:t>sdt-text</w:t></w:r></w:sdtContent>'
            b'</w:sdt>'
            b'<w:r><w:t> after</w:t></w:r>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        assert p.text == "before sdt-text after"


class DescribeCT_P_AllRunsIterator:
    """Phase A-v2 #2: iter_all_r_elements surfaces nested runs.

    See upstream #1370, #1021.
    """

    def it_yields_runs_inside_hyperlink(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            b' xmlns:r='
            b'"http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            b'<w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        rs = list(p.iter_all_r_elements())
        assert [r.text for r in rs] == ["link"]

    def it_yields_runs_inside_fldSimple(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        rs = list(p.iter_all_r_elements())
        assert [r.text for r in rs] == ["7"]

    def it_yields_runs_inside_sdt_sdtContent(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:sdt>'
            b'<w:sdtPr/>'
            b'<w:sdtContent><w:r><w:t>inside</w:t></w:r></w:sdtContent>'
            b'</w:sdt>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        rs = list(p.iter_all_r_elements())
        assert [r.text for r in rs] == ["inside"]

    def it_skips_field_code_only_runs_in_complex_fields(self):
        from docx.oxml.parser import parse_xml

        xml = (
            b'<w:p xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            b'<w:r><w:instrText> PAGE </w:instrText></w:r>'
            b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            b'<w:r><w:t>42</w:t></w:r>'
            b'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            b'</w:p>'
        )
        p = cast(CT_P, parse_xml(xml))
        texts = [r.text for r in p.iter_all_r_elements()]
        # -- fldChar runs carry no visible text (their .text is "") but the
        # -- instrText run (the field *code*) must NOT appear, only "42" does.
        assert "42" in texts
        assert " PAGE " not in texts
