"""Unit-test suite for the docx.form_fields module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.form_fields import (
    CheckboxFormField,
    DropdownFormField,
    FormField,
    TextInputFormField,
    WD_FORM_FIELD_TYPE,
)
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph

from .unitutil.cxml import element


# --- helper utilities ----------------------------------------------------


def _text_form_field_p() -> CT_P:
    """Build a CT_P containing a text form field via the public API."""
    p = cast(CT_P, element("w:p"))
    paragraph = Paragraph(p, None)  # type: ignore[arg-type]
    paragraph.add_text_form_field(name="Text1", default="Hello", maxlength=20)
    return p


def _checkbox_form_field_p(checked: bool = True) -> CT_P:
    p = cast(CT_P, element("w:p"))
    paragraph = Paragraph(p, None)  # type: ignore[arg-type]
    paragraph.add_checkbox_form_field(name="Agree", checked=checked)
    return p


def _dropdown_form_field_p(default_index: int = 0) -> CT_P:
    p = cast(CT_P, element("w:p"))
    paragraph = Paragraph(p, None)  # type: ignore[arg-type]
    paragraph.add_dropdown_form_field(
        name="Country", options=["US", "UK", "AU"], default_index=default_index
    )
    return p


def _begin_run(p: CT_P):
    return p.xpath("./w:r[w:fldChar[@w:fldCharType='begin' and w:ffData]]")[0]


# --- enum tests ----------------------------------------------------------


class DescribeWD_FORM_FIELD_TYPE:
    """Sanity check for the enum members."""

    @pytest.mark.parametrize(
        "member,value",
        [
            ("TEXT", "text"),
            ("CHECKBOX", "checkbox"),
            ("DROPDOWN", "dropdown"),
        ],
    )
    def it_exposes_the_three_form_field_types(self, member, value):
        assert getattr(WD_FORM_FIELD_TYPE, member).value == value


# --- FormField read tests -----------------------------------------------


class DescribeFormField_Text:
    """Reading a text form field."""

    def it_reports_its_type(self):
        p = _text_form_field_p()
        ff = FormField(_begin_run(p))
        assert ff.type is WD_FORM_FIELD_TYPE.TEXT

    def it_reports_its_name(self):
        ff = FormField(_begin_run(_text_form_field_p()))
        assert ff.name == "Text1"

    def it_reports_enabled_true_by_default(self):
        ff = FormField(_begin_run(_text_form_field_p()))
        assert ff.enabled is True

    def it_reports_calc_on_exit_false_by_default(self):
        ff = FormField(_begin_run(_text_form_field_p()))
        assert ff.calc_on_exit is False

    def it_exposes_a_text_input_view(self):
        ff = FormField(_begin_run(_text_form_field_p()))
        ti = ff.text_input
        assert isinstance(ti, TextInputFormField)
        assert ti.default == "Hello"
        assert ti.max_length == 20

    def it_returns_None_for_other_typed_views(self):
        ff = FormField(_begin_run(_text_form_field_p()))
        assert ff.checkbox is None
        assert ff.dropdown is None

    def its_value_is_the_rendered_result_text(self):
        ff = FormField(_begin_run(_text_form_field_p()))
        assert ff.value == "Hello"


class DescribeFormField_Checkbox:
    """Reading a checkbox form field."""

    def it_reports_its_type(self):
        ff = FormField(_begin_run(_checkbox_form_field_p(checked=True)))
        assert ff.type is WD_FORM_FIELD_TYPE.CHECKBOX

    def it_exposes_a_checkbox_view(self):
        ff = FormField(_begin_run(_checkbox_form_field_p(checked=True)))
        cb = ff.checkbox
        assert isinstance(cb, CheckboxFormField)
        assert cb.checked is True
        assert cb.default is True

    def its_value_toggles_with_checked(self):
        ff_on = FormField(_begin_run(_checkbox_form_field_p(checked=True)))
        ff_off = FormField(_begin_run(_checkbox_form_field_p(checked=False)))
        assert ff_on.value is True
        assert ff_off.value is False

    def it_returns_None_for_other_typed_views(self):
        ff = FormField(_begin_run(_checkbox_form_field_p()))
        assert ff.text_input is None
        assert ff.dropdown is None


class DescribeFormField_Dropdown:
    """Reading a dropdown form field."""

    def it_reports_its_type(self):
        ff = FormField(_begin_run(_dropdown_form_field_p()))
        assert ff.type is WD_FORM_FIELD_TYPE.DROPDOWN

    def it_exposes_a_dropdown_view(self):
        ff = FormField(_begin_run(_dropdown_form_field_p(default_index=1)))
        dd = ff.dropdown
        assert isinstance(dd, DropdownFormField)
        assert dd.options == ["US", "UK", "AU"]
        assert dd.default_index == 1
        assert dd.result_index == 1

    def its_value_is_the_selected_option(self):
        ff = FormField(_begin_run(_dropdown_form_field_p(default_index=2)))
        assert ff.value == "AU"

    def its_value_is_empty_string_when_index_is_out_of_range(self):
        # -- manually craft ddList with no entries; result_index defaults to 0
        p = cast(CT_P, element("w:p"))
        Paragraph(p, None).add_dropdown_form_field(  # type: ignore[arg-type]
            name="Empty", options=[], default_index=0
        )
        ff = FormField(_begin_run(p))
        assert ff.value == ""


class DescribeFormField_HelpAndStatus:
    """Help and status text."""

    def it_reads_helpText_and_statusText(self):
        # -- build an ffData block with helpText and statusText via cxml
        xml_str = (
            "w:p/w:r/w:fldChar{w:fldCharType=begin}"
            "/w:ffData/("
            "w:name{w:val=T1}"
            ",w:enabled"
            ",w:helpText{w:val=Type your name}"
            ",w:statusText{w:val=Name field}"
            ",w:textInput"
            ")"
        )
        p = cast(CT_P, element(xml_str))
        begin_run = p.xpath("./w:r")[0]
        ff = FormField(begin_run)
        assert ff.help_text == "Type your name"
        assert ff.status_text == "Name field"


# --- FormField builders ----------------------------------------------------


class DescribeParagraph_add_text_form_field:
    """`paragraph.add_text_form_field()` structure."""

    def it_emits_the_begin_separate_end_run_sequence_with_ffData(self):
        p = _text_form_field_p()

        runs = p.r_lst
        # -- five runs: begin, instrText, separate, result, end --
        assert len(runs) == 5
        fld_begin = runs[0].xpath("./w:fldChar")[0]
        assert fld_begin.get(qn("w:fldCharType")) == "begin"
        assert fld_begin.xpath("./w:ffData") != []
        assert runs[1].xpath("./w:instrText")[0].text == " FORMTEXT "
        assert runs[2].xpath("./w:fldChar")[0].get(qn("w:fldCharType")) == "separate"
        assert runs[3].xpath("./w:t")[0].text == "Hello"
        assert runs[4].xpath("./w:fldChar")[0].get(qn("w:fldCharType")) == "end"

    def it_returns_a_form_field_proxy(self):
        p = cast(CT_P, element("w:p"))
        ff = Paragraph(p, None).add_text_form_field(  # type: ignore[arg-type]
            name="Text1", default="hi"
        )
        assert isinstance(ff, FormField)
        assert ff.type is WD_FORM_FIELD_TYPE.TEXT
        assert ff.name == "Text1"

    def it_writes_maxlength_zero_when_no_limit(self):
        p = cast(CT_P, element("w:p"))
        Paragraph(p, None).add_text_form_field(  # type: ignore[arg-type]
            name="Text1"
        )
        maxLength = p.xpath(".//w:textInput/w:maxLength")[0]
        assert maxLength.get(qn("w:val")) == "0"


class DescribeParagraph_add_checkbox_form_field:
    """`paragraph.add_checkbox_form_field()` structure."""

    def it_writes_the_expected_ffData_shape(self):
        p = _checkbox_form_field_p(checked=True)

        assert p.xpath(".//w:instrText")[0].text == " FORMCHECKBOX "
        cb = p.xpath(".//w:ffData/w:checkBox")[0]
        assert cb.xpath("./w:default")[0].get(qn("w:val")) == "1"
        assert cb.xpath("./w:checked")[0].get(qn("w:val")) == "1"

    def it_defaults_checked_to_false(self):
        p = cast(CT_P, element("w:p"))
        Paragraph(p, None).add_checkbox_form_field(  # type: ignore[arg-type]
            name="Agree"
        )
        cb = p.xpath(".//w:ffData/w:checkBox")[0]
        assert cb.xpath("./w:default")[0].get(qn("w:val")) == "0"
        assert cb.xpath("./w:checked")[0].get(qn("w:val")) == "0"


class DescribeParagraph_add_dropdown_form_field:
    """`paragraph.add_dropdown_form_field()` structure."""

    def it_writes_the_expected_ffData_shape(self):
        p = _dropdown_form_field_p(default_index=1)

        assert p.xpath(".//w:instrText")[0].text == " FORMDROPDOWN "
        dd = p.xpath(".//w:ffData/w:ddList")[0]
        assert dd.xpath("./w:result")[0].get(qn("w:val")) == "1"
        assert dd.xpath("./w:default")[0].get(qn("w:val")) == "1"
        entries = [le.get(qn("w:val")) for le in dd.xpath("./w:listEntry")]
        assert entries == ["US", "UK", "AU"]

    def its_initial_result_text_is_the_selected_option(self):
        p = _dropdown_form_field_p(default_index=2)
        result_run = p.r_lst[3]
        assert result_run.xpath("./w:t")[0].text == "AU"


# --- iteration --------------------------------------------------------------


class DescribeParagraph_form_fields:
    """`paragraph.form_fields` iteration."""

    def it_yields_one_form_field_per_ffData_begin(self):
        p = cast(CT_P, element("w:p"))
        para = Paragraph(p, None)  # type: ignore[arg-type]
        para.add_text_form_field(name="T1", default="a")
        para.add_checkbox_form_field(name="C1", checked=True)
        para.add_dropdown_form_field(
            name="D1", options=["x", "y"], default_index=0
        )
        fields = para.form_fields
        assert len(fields) == 3
        assert [ff.type for ff in fields] == [
            WD_FORM_FIELD_TYPE.TEXT,
            WD_FORM_FIELD_TYPE.CHECKBOX,
            WD_FORM_FIELD_TYPE.DROPDOWN,
        ]

    def it_ignores_complex_fields_without_ffData(self):
        p = cast(CT_P, element("w:p"))
        para = Paragraph(p, None)  # type: ignore[arg-type]
        # -- a plain complex field (no ffData) should not appear in form_fields
        para.add_complex_field("PAGE", "1")
        para.add_text_form_field(name="T1", default="a")
        fields = para.form_fields
        assert len(fields) == 1
        assert fields[0].name == "T1"


class DescribeDocument_form_fields:
    """`document.form_fields` iteration over body paragraphs."""

    def it_walks_body_paragraphs(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element("w:document/w:body/(w:p,w:p)"),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        # -- add a form field to each body paragraph
        for paragraph in doc.paragraphs:
            paragraph.add_text_form_field(name=f"FF_{id(paragraph)}")

        fields = doc.form_fields
        assert len(fields) == 2
        assert all(ff.type is WD_FORM_FIELD_TYPE.TEXT for ff in fields)
