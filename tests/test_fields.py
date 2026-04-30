"""Unit-test suite for the docx.fields module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.fields import Field, WD_FIELD_TYPE
from docx.oxml.fields import CT_FldSimple
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R

from .unitutil.cxml import element


class DescribeWD_FIELD_TYPE:
    """Sanity check for the constant set."""

    @pytest.mark.parametrize(
        ("name", "value"),
        [
            ("PAGE", "PAGE"),
            ("NUMPAGES", "NUMPAGES"),
            ("DATE", "DATE"),
            ("TIME", "TIME"),
            ("AUTHOR", "AUTHOR"),
            ("REF", "REF"),
            ("TOC", "TOC"),
            ("SEQ", "SEQ"),
            ("HYPERLINK", "HYPERLINK"),
            ("PAGEREF", "PAGEREF"),
        ],
    )
    def it_exposes_common_field_types_as_string_constants(self, name: str, value: str):
        assert getattr(WD_FIELD_TYPE, name) == value


class DescribeField_Simple:
    """Unit-test suite for `Field` wrapping a ``w:fldSimple`` element."""

    def it_is_not_complex(self):
        fldSimple = cast(CT_FldSimple, element('w:fldSimple{w:instr=PAGE}'))
        field = Field.for_simple(fldSimple)
        assert field.is_complex is False

    def it_exposes_the_raw_instruction(self):
        # -- backslashes aren't supported by the cxml attribute grammar; build
        #    the element via OxmlElement and set the w:instr attribute directly.
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "REF bookmark1 \\h")
        field = Field.for_simple(fldSimple)
        assert field.instruction == "REF bookmark1 \\h"

    def it_returns_empty_instruction_when_attr_missing(self):
        # -- w:instr is required, but defensively handle absence --
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        field = Field.for_simple(fldSimple)
        assert field.instruction == ""

    @pytest.mark.parametrize(
        ("instr", "expected_type"),
        [
            ("PAGE", "PAGE"),
            ("NUMPAGES", "NUMPAGES"),
            ("REF bookmark1 \\h", "REF"),
            ("TOC \\o \"1-3\"", "TOC"),
            ("SEQ Table", "SEQ"),
            ("DATE", "DATE"),
            ("TIME \\@ \"h:mm AM/PM\"", "TIME"),
            ("AUTHOR", "AUTHOR"),
            ("HYPERLINK \"https://example.com\"", "HYPERLINK"),
            ("PAGEREF _Toc12345", "PAGEREF"),
        ],
    )
    def it_parses_the_type_from_the_instruction(self, instr: str, expected_type: str):
        from docx.oxml.parser import OxmlElement
        from docx.oxml.ns import qn

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), instr)
        field = Field.for_simple(fldSimple)
        assert field.type == expected_type

    def it_returns_empty_type_for_empty_instruction(self):
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        field = Field.for_simple(fldSimple)
        assert field.type == ""

    def it_exposes_the_result_text(self):
        fldSimple = cast(
            CT_FldSimple,
            element('w:fldSimple{w:instr=PAGE}/w:r/w:t"3"'),
        )
        field = Field.for_simple(fldSimple)
        assert field.result_text == "3"

    def it_returns_empty_result_text_when_no_run(self):
        fldSimple = cast(CT_FldSimple, element('w:fldSimple{w:instr=PAGE}'))
        field = Field.for_simple(fldSimple)
        assert field.result_text == ""


class DescribeField_Complex:
    """Unit-test suite for `Field` wrapping a complex field begin-run."""

    def _build_complex_paragraph(
        self, instr: str, result_text: str | None = None
    ) -> tuple[CT_P, CT_R]:
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field(instr, result_text)
        return p, begin_run

    def it_is_complex(self):
        _, begin_run = self._build_complex_paragraph("PAGE", "3")
        field = Field.for_complex(begin_run)
        assert field.is_complex is True

    def it_reads_the_instruction(self):
        _, begin_run = self._build_complex_paragraph("REF bookmark1 \\h", "See here")
        field = Field.for_complex(begin_run)
        assert field.instruction == "REF bookmark1 \\h"

    def it_reads_the_type(self):
        _, begin_run = self._build_complex_paragraph("NUMPAGES", "10")
        field = Field.for_complex(begin_run)
        assert field.type == "NUMPAGES"

    def it_reads_the_result_text(self):
        _, begin_run = self._build_complex_paragraph("PAGE", "42")
        field = Field.for_complex(begin_run)
        assert field.result_text == "42"

    def it_returns_empty_result_text_when_no_separate_marker(self):
        # -- build a field with only begin/instrText/end; no separate marker --
        p = cast(CT_P, element("w:p"))
        p.add_complex_field("PAGE")  # no result_text => 4 runs
        # -- remove the separate marker to leave only begin/instrText/end --
        from docx.oxml.ns import qn

        seps = p.xpath('.//w:fldChar[@w:fldCharType="separate"]')
        sep_run = seps[0].getparent()
        sep_run.getparent().remove(sep_run)
        begin_run = p.r_lst[0]
        field = Field.for_complex(begin_run)
        assert field.result_text == ""

    def it_returns_empty_result_text_when_omitted(self):
        _, begin_run = self._build_complex_paragraph("PAGE")
        field = Field.for_complex(begin_run)
        assert field.result_text == ""

    def it_concatenates_instruction_split_across_runs(self):
        # -- some producers split the instruction across multiple instrText runs --
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        p = cast(CT_P, element("w:p"))

        r_begin = p.add_r()
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        r_begin.append(fld_begin)

        r_i1 = p.add_r()
        i1 = OxmlElement("w:instrText")
        i1.text = "REF "
        r_i1.append(i1)

        r_i2 = p.add_r()
        i2 = OxmlElement("w:instrText")
        i2.text = "bookmark1"
        r_i2.append(i2)

        r_sep = p.add_r()
        fld_sep = OxmlElement("w:fldChar")
        fld_sep.set(qn("w:fldCharType"), "separate")
        r_sep.append(fld_sep)

        r_end = p.add_r()
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end"),
        r_end.append(fld_end)

        field = Field.for_complex(r_begin)
        assert field.instruction == "REF bookmark1"
        assert field.type == "REF"
