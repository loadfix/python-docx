# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.content_controls` module."""

from __future__ import annotations

from typing import cast

from docx.content_controls import ContentControl, ContentControlType, new_sdt
from docx.oxml.content_controls import CT_Sdt
from docx.oxml.ns import qn

from .unitutil.cxml import element


class DescribeContentControl:
    """Unit-test suite for `docx.content_controls.ContentControl`."""

    def it_knows_its_tag(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:tag{w:val=ABC}"),
        )
        cc = ContentControl(sdt)
        assert cc.tag == "ABC"

    def it_can_set_its_tag(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        cc = ContentControl(sdt)
        cc.tag = "New"
        assert cc.tag == "New"

    def it_knows_its_title(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:alias{w:val=Title}"),
        )
        cc = ContentControl(sdt)
        assert cc.title == "Title"

    def it_can_set_its_title(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        cc = ContentControl(sdt)
        cc.title = "Hello"
        assert cc.title == "Hello"

    def it_reports_RICH_TEXT_when_no_marker_is_present(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.RICH_TEXT

    def it_reports_PLAIN_TEXT_for_w_text_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:text"))
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.PLAIN_TEXT

    def it_reports_COMBO_BOX_for_w_comboBox_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:comboBox"))
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.COMBO_BOX

    def it_reports_DROPDOWN_for_w_dropDownList_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:dropDownList"))
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.DROPDOWN

    def it_reports_DATE_for_w_date_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:date"))
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.DATE

    def it_reports_PICTURE_for_w_picture_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:picture"))
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.PICTURE

    def it_reports_CHECKBOX_for_w14_checkbox_marker(self):
        from docx.oxml.parser import parse_xml

        sdt = cast(
            CT_Sdt,
            parse_xml(
                '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:sdtPr><w14:checkbox/></w:sdtPr>"
                "</w:sdt>"
            ),
        )
        cc = ContentControl(sdt)
        assert cc.type is ContentControlType.CHECKBOX

    def it_reads_checkbox_checked_value(self):
        from docx.oxml.parser import parse_xml

        sdt = cast(
            CT_Sdt,
            parse_xml(
                '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:sdtPr><w14:checkbox>"
                '<w14:checked w14:val="1"/>'
                "</w14:checkbox></w:sdtPr>"
                "</w:sdt>"
            ),
        )
        cc = ContentControl(sdt)
        assert cc.checked is True

    def it_can_round_trip_checkbox_checked(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        cc = ContentControl(sdt)
        cc.checked = True
        assert cc.checked is True
        cc.checked = False
        assert cc.checked is False

    def it_concatenates_text_from_sdtContent(self):
        sdt = cast(
            CT_Sdt,
            element('w:sdt/w:sdtContent/w:p/w:r/w:t"hello"'),
        )
        cc = ContentControl(sdt)
        assert cc.text == "hello"

    def it_can_set_text_on_inline_sdt(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtContent/w:r"))
        cc = ContentControl(sdt)
        cc.text = "value"
        assert cc.text == "value"
        # -- should have a w:r child in sdtContent --
        sdtContent = sdt.sdtContent
        assert sdtContent is not None
        assert sdtContent.find(qn("w:r")) is not None

    def it_can_set_text_on_block_sdt(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtContent/w:p"))
        cc = ContentControl(sdt)
        cc.text = "hi"
        assert cc.text == "hi"
        sdtContent = sdt.sdtContent
        assert sdtContent is not None
        assert sdtContent.find(qn("w:p")) is not None


class DescribeContentControlType:
    """Unit-test suite for `docx.content_controls.ContentControlType`."""

    def it_has_expected_members(self):
        assert ContentControlType.PLAIN_TEXT
        assert ContentControlType.RICH_TEXT
        assert ContentControlType.CHECKBOX
        assert ContentControlType.COMBO_BOX
        assert ContentControlType.DROPDOWN
        assert ContentControlType.DATE
        assert ContentControlType.PICTURE


class DescribeNewSdt:
    """Unit-test suite for the `new_sdt()` factory."""

    def it_creates_a_block_level_sdt_with_a_paragraph_child(self):
        sdt = new_sdt(ContentControlType.RICH_TEXT, tag="X", title="T", inline=False)
        assert sdt.sdtContent is not None
        assert sdt.sdtContent.find(qn("w:p")) is not None

    def it_creates_an_inline_sdt_with_a_run_child(self):
        sdt = new_sdt(ContentControlType.PLAIN_TEXT, tag="X", inline=True)
        assert sdt.sdtContent is not None
        assert sdt.sdtContent.find(qn("w:r")) is not None

    def it_sets_a_text_marker_for_PLAIN_TEXT(self):
        sdt = new_sdt(ContentControlType.PLAIN_TEXT, inline=True)
        assert sdt.type_marker_tag() == "w:text"

    def it_does_not_set_a_marker_for_RICH_TEXT(self):
        sdt = new_sdt(ContentControlType.RICH_TEXT, inline=False)
        assert sdt.type_marker_tag() is None

    def it_sets_the_tag_val_and_alias_val_when_provided(self):
        sdt = new_sdt(
            ContentControlType.RICH_TEXT, tag="TagVal", title="TitleVal", inline=False
        )
        assert sdt.tag_val == "TagVal"
        assert sdt.alias_val == "TitleVal"

    def it_always_sets_an_id(self):
        sdt = new_sdt(ContentControlType.RICH_TEXT, inline=False)
        assert isinstance(sdt.sdt_id, int)
        assert sdt.sdt_id is not None and sdt.sdt_id > 0
