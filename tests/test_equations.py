# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.equations` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.equations import (
    Equation,
    build_fraction,
    build_identifier,
    build_radical,
    build_subscript,
    build_superscript,
)
from docx.oxml.math import CT_OMath, CT_OMathPara
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .unitutil.cxml import element


class DescribeEquation:
    """Unit-test suite for `docx.equations.Equation`."""

    def it_exposes_the_raw_lxml_element(self):
        el = cast(CT_OMath, element('m:oMath/m:r/m:t"x"'))
        eq = Equation(el)
        assert eq.xml_element is el

    def it_knows_its_text(self):
        el = cast(
            CT_OMath,
            element('m:oMath/(m:r/m:t"a",m:r/m:t"bc")'),
        )
        eq = Equation(el)
        assert eq.text == "abc"

    def it_knows_its_raw_xml_bytes(self):
        el = cast(CT_OMath, element('m:oMath/m:r/m:t"y"'))
        eq = Equation(el)
        raw = eq.raw_xml
        assert isinstance(raw, bytes)
        assert b"<m:oMath" in raw
        assert b"<m:t>y</m:t>" in raw

    def it_is_inline_by_default(self):
        el = cast(CT_OMath, element('m:oMath/m:r/m:t"x"'))
        assert Equation(el).is_display_mode is False

    def it_is_display_mode_when_wrapped_in_m_oMathPara(self):
        el = cast(CT_OMathPara, element('m:oMathPara/m:oMath/m:r/m:t"x"'))
        assert Equation(el).is_display_mode is True

    def it_rejects_a_non_equation_root_element(self):
        el = element("w:p")
        with pytest.raises(ValueError, match="m:oMath or m:oMathPara"):
            Equation(el)  # pyright: ignore[reportArgumentType]

    def it_can_be_parsed_from_an_omml_xml_string(self):
        xml_str = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>x</m:t></m:r>"
            "</m:oMath>"
        )
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "x"
        assert eq.is_display_mode is False

    def it_can_be_parsed_from_display_mode_xml(self):
        xml_str = (
            '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>"
            "</m:oMathPara>"
        )
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "y"
        assert eq.is_display_mode is True

    def it_from_omml_xml_accepts_bytes(self):
        xml_bytes = (
            b'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            b"<m:r><m:t>z</m:t></m:r>"
            b"</m:oMath>"
        )
        eq = Equation.from_omml_xml(xml_bytes)
        assert eq.text == "z"

    def it_rejects_non_omml_root_in_from_omml_xml(self):
        with pytest.raises(ValueError, match="m:oMath"):
            Equation.from_omml_xml("<root xmlns='urn:x'><a/></root>")


class DescribeEquationBuilderFunctions:
    """Each builder must return valid OMML accepted by `from_omml_xml`."""

    def it_builds_an_identifier(self):
        xml_str = build_identifier("x")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "x"

    def it_escapes_xml_special_chars_in_identifier(self):
        xml_str = build_identifier("a<b&c")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "a<b&c"

    def it_builds_a_fraction(self):
        xml_str = build_fraction("1", "2")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "12"
        assert b'<m:f>' in eq.raw_xml

    def it_builds_a_superscript(self):
        xml_str = build_superscript("x", "2")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "x2"
        assert b'<m:sSup>' in eq.raw_xml

    def it_builds_a_subscript(self):
        xml_str = build_subscript("a", "n")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "an"
        assert b'<m:sSub>' in eq.raw_xml

    def it_builds_a_radical_without_degree(self):
        xml_str = build_radical("x")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "x"
        assert b'<m:rad>' in eq.raw_xml

    def it_builds_a_radical_with_degree(self):
        xml_str = build_radical("x", "3")
        eq = Equation.from_omml_xml(xml_str)
        assert eq.text == "3x"


class DescribeParagraphEquationsIntegration:
    """Tests that exercise the Paragraph.equations / add_equation API."""

    def it_finds_inline_equations_in_a_paragraph(self):
        p = cast(CT_P, element('w:p/m:oMath/m:r/m:t"x"'))
        paragraph = Paragraph(p, _FakeStoryParent())
        equations = paragraph.equations
        assert len(equations) == 1
        assert equations[0].is_display_mode is False
        assert equations[0].text == "x"

    def it_finds_display_mode_equations_in_a_paragraph(self):
        p = cast(CT_P, element('w:p/m:oMathPara/m:oMath/m:r/m:t"y"'))
        paragraph = Paragraph(p, _FakeStoryParent())
        equations = paragraph.equations
        assert len(equations) == 1
        assert equations[0].is_display_mode is True
        assert equations[0].text == "y"

    def it_returns_empty_list_when_no_equation_present(self):
        p = cast(CT_P, element('w:p/w:r/w:t"plain"'))
        paragraph = Paragraph(p, _FakeStoryParent())
        assert paragraph.equations == []

    def it_does_not_double_count_oMath_inside_oMathPara(self):
        p = cast(CT_P, element('w:p/m:oMathPara/m:oMath/m:r/m:t"z"'))
        paragraph = Paragraph(p, _FakeStoryParent())
        assert len(paragraph.equations) == 1

    def it_can_add_an_equation_from_an_omml_xml_string(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _FakeStoryParent())
        xml_str = build_superscript("x", "2")

        eq = paragraph.add_equation(xml_str)

        assert isinstance(eq, Equation)
        assert eq.is_display_mode is False
        assert len(paragraph.equations) == 1
        assert paragraph.equations[0].text == "x2"

    def it_can_add_an_equation_in_display_mode(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _FakeStoryParent())

        eq = paragraph.add_equation(build_fraction("1", "2"), display_mode=True)

        assert eq.is_display_mode is True
        # -- there's exactly one oMathPara and it wraps the oMath --
        assert len(p.xpath(".//m:oMathPara")) == 1
        assert len(p.xpath(".//m:oMathPara/m:oMath")) == 1

    def it_does_not_rewrap_an_oMathPara_input(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _FakeStoryParent())
        already_display = (
            '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"
            "</m:oMathPara>"
        )
        paragraph.add_equation(already_display, display_mode=True)
        # -- no nested oMathPara --
        assert len(p.xpath(".//m:oMathPara")) == 1


class DescribeDocumentEquations:
    """Document.equations walks body paragraphs."""

    def it_lists_equations_from_the_document_body(
        self, request: pytest.FixtureRequest
    ):
        from docx.document import Document
        from docx.oxml.document import CT_Document
        from docx.parts.document import DocumentPart

        from .unitutil.mock import instance_mock

        document_part_ = instance_mock(request, DocumentPart)
        doc_elm = cast(
            CT_Document,
            element(
                'w:document/w:body/('
                'w:p/m:oMath/m:r/m:t"x",'
                'w:p/m:oMathPara/m:oMath/m:r/m:t"y",'
                'w:p/w:r/w:t"plain"'
                ')'
            ),
        )
        document = Document(doc_elm, document_part_)

        equations = document.equations

        assert len(equations) == 2
        assert [e.text for e in equations] == ["x", "y"]
        assert [e.is_display_mode for e in equations] == [False, True]


class DescribeRunEquations:
    """Run.equations — empty for typical runs (OMML lives beside runs)."""

    def it_returns_empty_list_for_a_plain_run(self):
        p = cast(CT_P, element('w:p/w:r/w:t"x"'))
        r = p.xpath("./w:r")[0]
        run = Run(r, _FakeStoryParent())
        assert run.equations == []


class _FakeStoryParent:
    """Minimal stand-in for a StoryPart-providing parent."""

    @property
    def part(self):  # pragma: no cover - not exercised in these tests
        return None
