"""Unit-test suite for docx.oxml.fields."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.fields import CT_FldChar, CT_FldSimple, CT_InstrText, ST_FldCharType
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P

from ..unitutil.cxml import element, xml


class DescribeST_FldCharType:
    """Unit-test suite for the ``w:fldCharType`` simple type."""

    @pytest.mark.parametrize("value", ["begin", "separate", "end"])
    def it_accepts_valid_values(self, value: str):
        ST_FldCharType.validate(value)

    def it_rejects_invalid_values(self):
        with pytest.raises(ValueError, match="w:fldCharType must be one of"):
            ST_FldCharType.validate("nope")


class DescribeCT_FldSimple:
    """Unit-test suite for `docx.oxml.fields.CT_FldSimple`."""

    def it_exposes_its_instruction(self):
        fldSimple = cast(
            CT_FldSimple,
            element('w:fldSimple{w:instr=PAGE}'),
        )
        assert fldSimple.instr == "PAGE"

    def it_exposes_its_result_text(self):
        fldSimple = cast(
            CT_FldSimple,
            element('w:fldSimple{w:instr=PAGE}/w:r/w:t"3"'),
        )
        assert fldSimple.text == "3"

    def it_concatenates_multiple_run_text_children(self):
        fldSimple = cast(
            CT_FldSimple,
            element(
                'w:fldSimple{w:instr=PAGE}/(w:r/w:t"Page ",w:r/w:t"3")'
            ),
        )
        assert fldSimple.text == "Page 3"


class DescribeCT_FldChar:
    """Unit-test suite for `docx.oxml.fields.CT_FldChar`."""

    @pytest.mark.parametrize("fld_type", ["begin", "separate", "end"])
    def it_exposes_its_fldCharType(self, fld_type: str):
        fldChar = cast(
            CT_FldChar,
            element(f"w:fldChar{{w:fldCharType={fld_type}}}"),
        )
        assert fldChar.fldCharType == fld_type


class DescribeCT_InstrText:
    """Unit-test suite for `docx.oxml.fields.CT_InstrText`."""

    def it_exposes_its_text_via_str(self):
        instrText = cast(CT_InstrText, element('w:instrText"PAGE"'))
        assert str(instrText) == "PAGE"

    def it_returns_empty_string_when_no_text(self):
        instrText = cast(CT_InstrText, element("w:instrText"))
        assert str(instrText) == ""


class DescribeCT_P_FieldHelpers:
    """Unit-test suite for field-related helpers on CT_P."""

    def it_can_add_a_simple_field(self):
        p = cast(CT_P, element("w:p"))

        fldSimple = p.add_fldSimple("PAGE", "3")

        assert fldSimple.instr == "PAGE"
        assert len(p.fldSimple_lst) == 1
        assert fldSimple.text == "3"

    def it_can_add_a_simple_field_without_result_text(self):
        p = cast(CT_P, element("w:p"))

        fldSimple = p.add_fldSimple("DATE")

        assert fldSimple.instr == "DATE"
        # -- no run children when no text provided --
        assert len(fldSimple.r_lst) == 0

    def it_preserves_space_in_instruction(self):
        p = cast(CT_P, element("w:p"))

        p.add_complex_field("REF bookmark1", "See here")

        # -- instrText should have xml:space="preserve" since instr has spaces --
        instrTexts = p.xpath(".//w:instrText")
        assert instrTexts[0].get(qn("xml:space")) == "preserve"

    def it_emits_the_five_run_sequence_for_a_complex_field(self):
        p = cast(CT_P, element("w:p"))

        p.add_complex_field("PAGE", "3")

        runs = p.r_lst
        assert len(runs) == 5
        assert runs[0][0].tag == qn("w:fldChar")
        assert runs[0][0].get(qn("w:fldCharType")) == "begin"
        assert runs[1][0].tag == qn("w:instrText")
        assert runs[2][0].tag == qn("w:fldChar")
        assert runs[2][0].get(qn("w:fldCharType")) == "separate"
        assert runs[3][0].tag == qn("w:t")
        assert runs[4][0].tag == qn("w:fldChar")
        assert runs[4][0].get(qn("w:fldCharType")) == "end"

    def it_emits_four_runs_when_result_text_is_omitted(self):
        p = cast(CT_P, element("w:p"))

        p.add_complex_field("PAGE")

        runs = p.r_lst
        assert len(runs) == 4
        # -- no result-text run between separate and end --
        assert runs[3][0].tag == qn("w:fldChar")
        assert runs[3][0].get(qn("w:fldCharType")) == "end"

    def it_returns_the_begin_run_from_add_complex_field(self):
        p = cast(CT_P, element("w:p"))

        begin_run = p.add_complex_field("PAGE", "3")

        assert begin_run is p.r_lst[0]

    def it_iterates_field_elements_in_document_order(self):
        p = cast(CT_P, element("w:p"))
        p.add_fldSimple("PAGE", "1")
        p.add_complex_field("NUMPAGES", "10")
        p.add_fldSimple("DATE", "2026-01-01")

        kinds = [kind for kind, _ in p.iter_field_elements()]

        assert kinds == ["simple", "complex", "simple"]


class DescribeCT_P_Text:
    """Verify `paragraph.text` picks up fldSimple content."""

    def it_includes_fldSimple_text(self):
        p = cast(
            CT_P,
            element(
                'w:p/(w:r/w:t"Page ",w:fldSimple{w:instr=PAGE}/w:r/w:t"3")'
            ),
        )
        assert p.text == "Page 3"

    def it_includes_complex_field_result_text(self):
        # -- complex fields use regular w:r children, so result text is already
        #    covered by the existing xpath; just confirm it here.
        p = cast(CT_P, element('w:p/w:r/w:t"Page "'))
        p.add_complex_field("PAGE", "3")
        assert p.text == "Page 3"
