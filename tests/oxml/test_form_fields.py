"""Unit-test suite for docx.oxml.form_fields."""

from __future__ import annotations

from typing import cast

from docx.oxml.fields import CT_FldChar
from docx.oxml.form_fields import (
    CT_FFCheckBox,
    CT_FFData,
    CT_FFDDList,
    CT_FFTextInput,
)
from docx.oxml.ns import qn

from ..unitutil.cxml import element


class DescribeCT_FFData:
    """Unit-test suite for `docx.oxml.form_fields.CT_FFData`."""

    def it_exposes_a_text_input_child(self):
        ffData = cast(
            CT_FFData,
            element("w:ffData/(w:name{w:val=T1},w:enabled,w:textInput)"),
        )
        assert isinstance(ffData.textInput, CT_FFTextInput)
        assert ffData.checkBox is None
        assert ffData.ddList is None

    def it_exposes_a_check_box_child(self):
        ffData = cast(
            CT_FFData,
            element("w:ffData/(w:name{w:val=C1},w:checkBox)"),
        )
        assert isinstance(ffData.checkBox, CT_FFCheckBox)
        assert ffData.textInput is None
        assert ffData.ddList is None

    def it_exposes_a_dd_list_child(self):
        ffData = cast(
            CT_FFData,
            element("w:ffData/(w:name{w:val=D1},w:ddList)"),
        )
        assert isinstance(ffData.ddList, CT_FFDDList)
        assert ffData.textInput is None
        assert ffData.checkBox is None

    def it_exposes_the_name_help_and_status_children(self):
        ffData = cast(
            CT_FFData,
            element(
                "w:ffData/("
                "w:name{w:val=FF1}"
                ",w:enabled"
                ",w:calcOnExit"
                ",w:helpText{w:val=Help}"
                ",w:statusText{w:val=Status}"
                ",w:textInput"
                ")"
            ),
        )
        assert ffData.name.get(qn("w:val")) == "FF1"
        assert ffData.enabled is not None
        assert ffData.calcOnExit is not None
        assert ffData.helpText.get(qn("w:val")) == "Help"
        assert ffData.statusText.get(qn("w:val")) == "Status"


class DescribeCT_FFTextInput:
    """Unit-test suite for `CT_FFTextInput`."""

    def it_exposes_its_default_max_length_and_format(self):
        ti = cast(
            CT_FFTextInput,
            element(
                "w:textInput/("
                "w:default{w:val=hello}"
                ",w:maxLength{w:val=10}"
                ",w:format{w:val=UPPERCASE}"
                ")"
            ),
        )
        assert ti.default.get(qn("w:val")) == "hello"
        assert ti.maxLength.get(qn("w:val")) == "10"
        assert ti.format.get(qn("w:val")) == "UPPERCASE"

    def it_returns_None_for_missing_children(self):
        ti = cast(CT_FFTextInput, element("w:textInput"))
        assert ti.default is None
        assert ti.maxLength is None
        assert ti.format is None


class DescribeCT_FFCheckBox:
    """Unit-test suite for `CT_FFCheckBox`."""

    def it_exposes_default_and_checked(self):
        cb = cast(
            CT_FFCheckBox,
            element("w:checkBox/(w:default{w:val=1},w:checked{w:val=0})"),
        )
        assert cb.default.get(qn("w:val")) == "1"
        assert cb.checked.get(qn("w:val")) == "0"


class DescribeCT_FFDDList:
    """Unit-test suite for `CT_FFDDList`."""

    def it_exposes_result_default_and_entries(self):
        dd = cast(
            CT_FFDDList,
            element(
                "w:ddList/("
                "w:result{w:val=1}"
                ",w:default{w:val=0}"
                ",w:listEntry{w:val=US}"
                ",w:listEntry{w:val=UK}"
                ",w:listEntry{w:val=AU}"
                ")"
            ),
        )
        assert dd.result.get(qn("w:val")) == "1"
        assert dd.default.get(qn("w:val")) == "0"
        entries = [le.get(qn("w:val")) for le in dd.xpath("./w:listEntry")]
        assert entries == ["US", "UK", "AU"]


class DescribeCT_FldChar_ffData:
    """Verify CT_FldChar exposes its ffData child."""

    def it_exposes_its_ffData_child(self):
        fldChar = cast(
            CT_FldChar,
            element(
                "w:fldChar{w:fldCharType=begin}"
                "/w:ffData/(w:name{w:val=T1},w:textInput)"
            ),
        )
        assert fldChar.ffData is not None
        assert fldChar.ffData.textInput is not None

    def it_returns_None_when_ffData_is_absent(self):
        fldChar = cast(
            CT_FldChar,
            element("w:fldChar{w:fldCharType=begin}"),
        )
        assert fldChar.ffData is None
