"""Test suite for the docx.oxml.text.run module."""

from typing import cast

import pytest

from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R, CT_Sym

from ...unitutil.cxml import element, xml


class DescribeCT_R:
    """Unit-test suite for the CT_R (run, <w:r>) element."""

    @pytest.mark.parametrize(
        ("initial_cxml", "text", "expected_cxml"),
        [
            ("w:r", "foobar", 'w:r/w:t"foobar"'),
            ("w:r", "foobar ", 'w:r/w:t{xml:space=preserve}"foobar "'),
            (
                "w:r/(w:rPr/w:rStyle{w:val=emphasis}, w:cr)",
                "foobar",
                'w:r/(w:rPr/w:rStyle{w:val=emphasis}, w:cr, w:t"foobar")',
            ),
        ],
    )
    def it_can_add_a_t_preserving_edge_whitespace(
        self, initial_cxml: str, text: str, expected_cxml: str
    ):
        r = cast(CT_R, element(initial_cxml))
        expected_xml = xml(expected_cxml)

        r.add_t(text)

        assert r.xml == expected_xml

    def it_can_assemble_the_text_in_the_run(self):
        cxml = 'w:r/(w:br,w:cr,w:noBreakHyphen,w:ptab,w:t"foobar",w:tab)'
        r = cast(CT_R, element(cxml))

        assert r.text == "\n\n-\tfoobar\t"

    @pytest.mark.parametrize(
        ("p_cxml", "offset", "expected_left_text", "expected_right_text"),
        [
            # split in middle of text
            ('w:p/w:r/w:t"foobar"', 3, "foo", "bar"),
            # split at beginning — left run is empty
            ('w:p/w:r/w:t"foobar"', 0, "", "foobar"),
            # split at end — right run is empty
            ('w:p/w:r/w:t"foobar"', 6, "foobar", ""),
            # split run with formatting — both get rPr
            ('w:p/w:r/(w:rPr/w:b,w:t"foobar")', 3, "foo", "bar"),
        ],
    )
    def it_can_split_at_a_character_offset(
        self,
        p_cxml: str,
        offset: int,
        expected_left_text: str,
        expected_right_text: str,
    ):
        p = cast(CT_P, element(p_cxml))
        r = p.r_lst[0]

        new_r = r.split_run(offset)

        assert r.text == expected_left_text
        assert new_r.text == expected_right_text
        # -- new run is next sibling --
        assert r.getnext() is new_r
        assert len(p.r_lst) == 2

    def it_copies_rPr_to_the_new_run_on_split(self):
        p = cast(CT_P, element('w:p/w:r/(w:rPr/(w:b,w:i),w:t"foobar")'))
        r = p.r_lst[0]

        new_r = r.split_run(3)

        # -- both runs have bold+italic --
        assert r.rPr is not None
        assert new_r.rPr is not None
        assert r.rPr.xml == new_r.rPr.xml
        # -- but they are distinct elements, not the same object --
        assert r.rPr is not new_r.rPr

    def it_splits_a_run_with_no_formatting(self):
        p = cast(CT_P, element('w:p/w:r/w:t"hello"'))
        r = p.r_lst[0]

        new_r = r.split_run(2)

        assert r.text == "he"
        assert new_r.text == "llo"
        assert r.rPr is None
        assert new_r.rPr is None

    def it_raises_on_invalid_offset(self):
        p = cast(CT_P, element('w:p/w:r/w:t"hello"'))
        r = p.r_lst[0]

        with pytest.raises(ValueError, match="offset -1 out of range"):
            r.split_run(-1)
        with pytest.raises(ValueError, match="offset 6 out of range"):
            r.split_run(6)

    def it_can_add_a_w_sym_child(self):
        r = cast(CT_R, element("w:r"))

        sym = r.add_sym("F0E0", "Wingdings")

        assert r.xml == xml("w:r/w:sym{w:font=Wingdings,w:char=F0E0}")
        assert isinstance(sym, CT_Sym)
        assert sym.font == "Wingdings"
        assert sym.char == "F0E0"

    def it_exposes_its_w_sym_children_via_sym_lst(self):
        r = cast(
            CT_R,
            element(
                "w:r/(w:sym{w:font=Wingdings,w:char=F0E0},"
                "w:t\"x\","
                "w:sym{w:font=Symbol,w:char=0041})"
            ),
        )

        syms = r.sym_lst

        assert [s.char for s in syms] == ["F0E0", "0041"]
        assert [s.font for s in syms] == ["Wingdings", "Symbol"]

    def it_renders_w_sym_as_char_in_run_text(self):
        """Closes upstream#1528 — ``w:sym`` contributes ``chr(@w:char)``."""
        r = cast(
            CT_R,
            element(
                "w:r/(w:t\"a\",w:sym{w:font=Wingdings,w:char=F0E0},w:t\"b\")"
            ),
        )

        assert r.text == "a" + chr(0xF0E0) + "b"


class DescribeCT_Sym:
    """Unit-test suite for :class:`docx.oxml.text.run.CT_Sym`."""

    def it_renders_as_the_derived_character(self):
        sym = cast(CT_Sym, element("w:sym{w:font=Wingdings,w:char=F0E0}"))
        assert str(sym) == chr(0xF0E0)

    def it_returns_empty_string_on_invalid_hex(self):
        sym = cast(CT_Sym, element("w:sym{w:font=Wingdings,w:char=notHex}"))
        assert str(sym) == ""


class DescribeCT_R_TextSetterPreservesReferences:
    """Phase A-v2 #3: Run.text setter preserves reference-carrying children.

    See upstream #1519. Reassigning ``text`` must not silently delete
    ``w:commentReference``, ``w:footnoteReference``, ``w:endnoteReference``,
    ``w:fldChar`` marker, or ``w:instrText`` siblings.
    """

    def it_preserves_commentReference_on_text_reassignment(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import parse_xml

        xml_bytes = (
            b'<w:r xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:t>old</w:t>'
            b'<w:commentReference w:id="3"/>'
            b'</w:r>'
        )
        r = cast(CT_R, parse_xml(xml_bytes))
        r.text = "new"
        assert r.find(qn("w:commentReference")) is not None
        assert r.text == "new"

    def it_preserves_footnoteReference_on_text_reassignment(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import parse_xml

        xml_bytes = (
            b'<w:r xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:t>old</w:t>'
            b'<w:footnoteReference w:id="2"/>'
            b'</w:r>'
        )
        r = cast(CT_R, parse_xml(xml_bytes))
        r.text = "new"
        assert r.find(qn("w:footnoteReference")) is not None

    def it_preserves_endnoteReference_on_text_reassignment(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import parse_xml

        xml_bytes = (
            b'<w:r xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:t>old</w:t>'
            b'<w:endnoteReference w:id="4"/>'
            b'</w:r>'
        )
        r = cast(CT_R, parse_xml(xml_bytes))
        r.text = "new"
        assert r.find(qn("w:endnoteReference")) is not None

    def it_preserves_fldChar_on_text_reassignment(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import parse_xml

        xml_bytes = (
            b'<w:r xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:fldChar w:fldCharType="begin"/>'
            b'<w:t>old</w:t>'
            b'</w:r>'
        )
        r = cast(CT_R, parse_xml(xml_bytes))
        r.text = "new"
        assert r.find(qn("w:fldChar")) is not None

    def it_preserves_instrText_on_text_reassignment(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import parse_xml

        xml_bytes = (
            b'<w:r xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:instrText> REF bookmark1 </w:instrText>'
            b'<w:t>old</w:t>'
            b'</w:r>'
        )
        r = cast(CT_R, parse_xml(xml_bytes))
        r.text = "new"
        assert r.find(qn("w:instrText")) is not None
