# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.sdt` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.sdt import CT_Sdt, CT_SdtContent, CT_SdtPr

from ..unitutil.cxml import element


class DescribeCT_SdtPr:
    """Unit-test suite for `docx.oxml.sdt.CT_SdtPr`."""

    def it_can_get_the_tag_val(self):
        sdtPr = cast(CT_SdtPr, element('w:sdtPr/w:tag{w:val=myTag}'))
        assert sdtPr.tag_val == "myTag"

    def it_returns_None_when_no_tag_child(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        assert sdtPr.tag_val is None

    def it_can_set_the_tag_val(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        sdtPr.tag_val = "newTag"
        assert sdtPr.tag_val == "newTag"

    def it_can_clear_the_tag_val(self):
        sdtPr = cast(CT_SdtPr, element('w:sdtPr/w:tag{w:val=myTag}'))
        sdtPr.tag_val = None
        assert sdtPr.tag_val is None

    def it_can_get_the_alias_val(self):
        sdtPr = cast(CT_SdtPr, element('w:sdtPr/w:alias{w:val=myTitle}'))
        assert sdtPr.alias_val == "myTitle"

    def it_returns_None_when_no_alias_child(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        assert sdtPr.alias_val is None

    def it_can_set_the_alias_val(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        sdtPr.alias_val = "newTitle"
        assert sdtPr.alias_val == "newTitle"

    @pytest.mark.parametrize(
        ("cxml", "expected_type"),
        [
            ("w:sdtPr", "richText"),
            ("w:sdtPr/w:text", "plainText"),
            ("w:sdtPr/w:comboBox", "comboBox"),
            ("w:sdtPr/w:dropDownList", "dropDown"),
            ("w:sdtPr/w:date", "date"),
            ("w:sdtPr/w:picture", "picture"),
        ],
    )
    def it_can_determine_the_sdt_type(self, cxml: str, expected_type: str):
        sdtPr = cast(CT_SdtPr, element(cxml))
        assert sdtPr.sdt_type == expected_type


class DescribeCT_SdtContent:
    """Unit-test suite for `docx.oxml.sdt.CT_SdtContent`."""

    def it_can_get_text_from_paragraphs(self):
        sdtContent = cast(
            CT_SdtContent,
            element('w:sdtContent/w:p/w:r/w:t"Hello"'),
        )
        assert sdtContent.text == "Hello"

    def it_can_get_text_from_runs(self):
        sdtContent = cast(
            CT_SdtContent,
            element('w:sdtContent/w:r/w:t"World"'),
        )
        assert sdtContent.text == "World"

    def it_returns_empty_string_for_empty_content(self):
        sdtContent = cast(CT_SdtContent, element("w:sdtContent"))
        assert sdtContent.text == ""


class DescribeCT_Sdt:
    """Unit-test suite for `docx.oxml.sdt.CT_Sdt`."""

    def it_can_create_a_block_level_sdt(self):
        sdt = CT_Sdt.new_block("richText", tag="myTag", title="myTitle")

        assert sdt.sdtPr is not None
        assert sdt.sdtPr.tag_val == "myTag"
        assert sdt.sdtPr.alias_val == "myTitle"
        assert sdt.sdtContent is not None
        # -- should have a paragraph in its content --
        assert len(sdt.sdtContent.p_lst) == 1

    def it_can_create_an_inline_sdt(self):
        sdt = CT_Sdt.new_inline("plainText", tag="field1", title="Field 1")

        assert sdt.sdtPr is not None
        assert sdt.sdtPr.tag_val == "field1"
        assert sdt.sdtPr.alias_val == "Field 1"
        assert sdt.sdtPr.sdt_type == "plainText"
        assert sdt.sdtContent is not None
        # -- should have a run in its content --
        assert len(sdt.sdtContent.r_lst) == 1

    def it_can_create_a_checkbox_sdt(self):
        sdt = CT_Sdt.new_inline("checkbox", tag="cb1")

        assert sdt.sdtPr is not None
        assert sdt.sdtPr.sdt_type == "checkbox"
        checkbox = sdt.sdtPr.checkbox
        assert checkbox is not None
        assert checkbox.checked is False

    def it_can_create_sdt_types(self):
        for sdt_type in ("richText", "plainText", "comboBox", "dropDown", "date", "picture"):
            sdt = CT_Sdt.new_block(sdt_type)
            assert sdt.sdtPr.sdt_type == sdt_type
