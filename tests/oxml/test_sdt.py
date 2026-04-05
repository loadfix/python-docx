# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.oxml.sdt` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.contentcontrol import WD_CONTENT_CONTROL_TYPE
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.sdt import CT_Sdt, CT_SdtCheckbox, CT_SdtContent, CT_SdtPr


class DescribeCT_SdtPr:
    """Unit-test suite for `docx.oxml.sdt.CT_SdtPr` objects."""

    def it_can_report_tag_val(self):
        sdtPr = cast(
            CT_SdtPr,
            parse_xml(f'<w:sdtPr {nsdecls("w")}><w:tag w:val="my_tag"/></w:sdtPr>'),
        )
        assert sdtPr.tag_val == "my_tag"

    def it_returns_None_for_tag_val_when_not_present(self):
        sdtPr = cast(CT_SdtPr, parse_xml(f'<w:sdtPr {nsdecls("w")}/>'))
        assert sdtPr.tag_val is None

    def it_can_set_tag_val(self):
        sdtPr = cast(CT_SdtPr, parse_xml(f'<w:sdtPr {nsdecls("w")}/>'))
        sdtPr.tag_val = "new_tag"
        assert sdtPr.tag_val == "new_tag"

    def it_can_remove_tag_val(self):
        sdtPr = cast(
            CT_SdtPr,
            parse_xml(f'<w:sdtPr {nsdecls("w")}><w:tag w:val="old"/></w:sdtPr>'),
        )
        sdtPr.tag_val = None
        assert sdtPr.tag_val is None

    def it_can_report_title(self):
        sdtPr = cast(
            CT_SdtPr,
            parse_xml(f'<w:sdtPr {nsdecls("w")}><w:alias w:val="My Title"/></w:sdtPr>'),
        )
        assert sdtPr.title == "My Title"

    def it_can_set_title(self):
        sdtPr = cast(CT_SdtPr, parse_xml(f'<w:sdtPr {nsdecls("w")}/>'))
        sdtPr.title = "New Title"
        assert sdtPr.title == "New Title"

    @pytest.mark.parametrize(
        ("type_xml", "expected_type"),
        [
            ("", WD_CONTENT_CONTROL_TYPE.RICH_TEXT),
            ("<w:text/>", WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT),
            (f"<w14:checkbox {nsdecls('w14')}/>", WD_CONTENT_CONTROL_TYPE.CHECKBOX),
            ("<w:comboBox/>", WD_CONTENT_CONTROL_TYPE.COMBO_BOX),
            ("<w:dropDownList/>", WD_CONTENT_CONTROL_TYPE.DROP_DOWN),
            ("<w:date/>", WD_CONTENT_CONTROL_TYPE.DATE),
            ("<w:picture/>", WD_CONTENT_CONTROL_TYPE.PICTURE),
        ],
    )
    def it_can_identify_control_type(
        self, type_xml: str, expected_type: WD_CONTENT_CONTROL_TYPE
    ):
        sdtPr = cast(
            CT_SdtPr,
            parse_xml(f'<w:sdtPr {nsdecls("w", "w14")}>{type_xml}</w:sdtPr>'),
        )
        assert sdtPr.control_type == expected_type


class DescribeCT_SdtCheckbox:
    """Unit-test suite for `docx.oxml.sdt.CT_SdtCheckbox` objects."""

    def it_can_report_checked_state(self):
        checkbox = cast(
            CT_SdtCheckbox,
            parse_xml(
                f'<w14:checkbox {nsdecls("w14")}>'
                f'  <w14:checked w14:val="1"/>'
                f"</w14:checkbox>"
            ),
        )
        assert checkbox.checked is True

    def it_reports_unchecked_when_val_is_0(self):
        checkbox = cast(
            CT_SdtCheckbox,
            parse_xml(
                f'<w14:checkbox {nsdecls("w14")}>'
                f'  <w14:checked w14:val="0"/>'
                f"</w14:checkbox>"
            ),
        )
        assert checkbox.checked is False

    def it_reports_unchecked_when_no_checked_element(self):
        checkbox = cast(
            CT_SdtCheckbox,
            parse_xml(f'<w14:checkbox {nsdecls("w14")}/>'),
        )
        assert checkbox.checked is False

    def it_can_set_checked_state(self):
        checkbox = cast(
            CT_SdtCheckbox,
            parse_xml(
                f'<w14:checkbox {nsdecls("w14")}>'
                f'  <w14:checked w14:val="0"/>'
                f"</w14:checkbox>"
            ),
        )
        checkbox.checked = True
        assert checkbox.checked is True

    def it_can_create_checked_element_when_setting(self):
        checkbox = cast(
            CT_SdtCheckbox,
            parse_xml(f'<w14:checkbox {nsdecls("w14")}/>'),
        )
        checkbox.checked = True
        assert checkbox.checked is True


class DescribeCT_Sdt:
    """Unit-test suite for `docx.oxml.sdt.CT_Sdt` objects."""

    def it_can_create_a_block_level_sdt(self):
        sdt = CT_Sdt.new_block(
            WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT,
            tag="test_tag",
            title="Test Title",
        )
        assert sdt.sdtPr is not None
        assert sdt.sdtPr.tag_val == "test_tag"
        assert sdt.sdtPr.title == "Test Title"
        assert sdt.sdtPr.control_type == WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT
        assert sdt.sdtContent is not None
        assert len(sdt.sdtContent.p_lst) == 1

    def it_can_create_an_inline_sdt(self):
        sdt = CT_Sdt.new_inline(
            WD_CONTENT_CONTROL_TYPE.RICH_TEXT,
            tag="inline_tag",
            title="Inline Title",
        )
        assert sdt.sdtPr is not None
        assert sdt.sdtPr.tag_val == "inline_tag"
        assert sdt.sdtPr.title == "Inline Title"
        assert sdt.sdtPr.control_type == WD_CONTENT_CONTROL_TYPE.RICH_TEXT
        assert sdt.sdtContent is not None
        assert len(sdt.sdtContent.r_lst) == 1

    def it_can_create_a_checkbox_sdt(self):
        sdt = CT_Sdt.new_block(WD_CONTENT_CONTROL_TYPE.CHECKBOX)
        assert sdt.sdtPr is not None
        assert sdt.sdtPr.control_type == WD_CONTENT_CONTROL_TYPE.CHECKBOX
        assert sdt.sdtPr.checkbox is not None
        assert sdt.sdtPr.checkbox.checked is False


class DescribeCT_SdtContent:
    """Unit-test suite for `docx.oxml.sdt.CT_SdtContent` objects."""

    def it_can_report_text_from_runs(self):
        content = cast(
            CT_SdtContent,
            parse_xml(
                f'<w:sdtContent {nsdecls("w")}>'
                f"  <w:r><w:t>Hello </w:t></w:r>"
                f"  <w:r><w:t>World</w:t></w:r>"
                f"</w:sdtContent>"
            ),
        )
        assert content.text == "Hello World"

    def it_can_report_text_from_paragraphs(self):
        content = cast(
            CT_SdtContent,
            parse_xml(
                f'<w:sdtContent {nsdecls("w")}>'
                f"  <w:p><w:r><w:t>Line 1</w:t></w:r></w:p>"
                f"  <w:p><w:r><w:t>Line 2</w:t></w:r></w:p>"
                f"</w:sdtContent>"
            ),
        )
        assert content.text == "Line 1\nLine 2"

    def it_returns_empty_string_when_no_content(self):
        content = cast(
            CT_SdtContent,
            parse_xml(f'<w:sdtContent {nsdecls("w")}/>'),
        )
        assert content.text == ""
