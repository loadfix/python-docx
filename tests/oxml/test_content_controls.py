# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.content_controls` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.content_controls import CT_Sdt, CT_SdtContent
from docx.oxml.ns import qn

from ..unitutil.cxml import element


class DescribeCT_Sdt:
    """Unit-test suite for `docx.oxml.content_controls.CT_Sdt`."""

    def it_reads_its_tag_val_from_sdtPr(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:tag{w:val=MyTag}"),
        )
        assert sdt.tag_val == "MyTag"

    def it_returns_None_for_tag_val_when_not_present(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.tag_val is None

    def it_can_set_tag_val(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.tag_val = "MyTag"
        assert sdt.tag_val == "MyTag"
        assert sdt.sdtPr is not None
        tag_elm = sdt.sdtPr.find(qn("w:tag"))
        assert tag_elm is not None
        assert tag_elm.get(qn("w:val")) == "MyTag"

    def it_reads_its_alias_val_from_sdtPr(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:alias{w:val=MyTitle}"),
        )
        assert sdt.alias_val == "MyTitle"

    def it_returns_None_for_alias_val_when_not_present(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.alias_val is None

    def it_can_set_alias_val(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.alias_val = "Hello"
        assert sdt.alias_val == "Hello"

    def it_can_remove_tag_val_by_assigning_None(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:tag{w:val=x}"),
        )
        sdt.tag_val = None
        assert sdt.tag_val is None

    def it_reads_its_id(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:id{w:val=42}"),
        )
        assert sdt.sdt_id == 42

    def it_can_set_its_id(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.sdt_id = 99
        assert sdt.sdt_id == 99

    def it_detects_no_type_marker_as_rich_text_default(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.type_marker_tag() is None

    def it_detects_w_text_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:text"))
        assert sdt.type_marker_tag() == "w:text"

    def it_detects_w_date_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:date"))
        assert sdt.type_marker_tag() == "w:date"

    def it_reads_checked_value_when_present(self):
        xml = (
            '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            '<w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>'
            "</w:sdt>"
        )
        from docx.oxml.parser import parse_xml

        sdt = cast(CT_Sdt, parse_xml(xml))
        assert sdt.checked is True

    def it_returns_None_for_checked_when_no_checkbox(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.checked is None

    def it_can_set_checked(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.checked = True
        assert sdt.checked is True
        sdt.checked = False
        assert sdt.checked is False

    def it_can_set_a_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.set_type_marker("w:text")
        assert sdt.type_marker_tag() == "w:text"

    def it_replaces_existing_type_marker_on_set(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:text"))
        sdt.set_type_marker("w:date")
        assert sdt.type_marker_tag() == "w:date"
        assert sdt.sdtPr is not None
        assert sdt.sdtPr.find(qn("w:text")) is None


class DescribeCT_SdtContent:
    """Unit-test suite for `docx.oxml.content_controls.CT_SdtContent`."""

    def it_concatenates_paragraph_text(self):
        sdtContent = cast(
            CT_SdtContent,
            element('w:sdtContent/(w:p/w:r/w:t"Hello")'),
        )
        assert sdtContent.text == "Hello"

    def it_concatenates_run_text(self):
        sdtContent = cast(
            CT_SdtContent,
            element('w:sdtContent/(w:r/w:t"Hi")'),
        )
        assert sdtContent.text == "Hi"

    def it_returns_empty_string_when_no_children(self):
        sdtContent = cast(CT_SdtContent, element("w:sdtContent"))
        assert sdtContent.text == ""
