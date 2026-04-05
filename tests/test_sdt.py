# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.sdt` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.sdt import WD_CONTENT_CONTROL_TYPE
from docx.oxml.sdt import CT_Sdt
from docx.sdt import ContentControl


class DescribeContentControl:
    """Unit-test suite for `docx.sdt.ContentControl`."""

    def it_can_get_the_tag(self):
        sdt = CT_Sdt.new_block("richText", tag="myTag")
        cc = ContentControl(sdt)
        assert cc.tag == "myTag"

    def it_returns_None_when_no_tag(self):
        sdt = CT_Sdt.new_block("richText")
        cc = ContentControl(sdt)
        assert cc.tag is None

    def it_can_set_the_tag(self):
        sdt = CT_Sdt.new_block("richText")
        cc = ContentControl(sdt)
        cc.tag = "newTag"
        assert cc.tag == "newTag"

    def it_can_get_the_title(self):
        sdt = CT_Sdt.new_block("richText", title="myTitle")
        cc = ContentControl(sdt)
        assert cc.title == "myTitle"

    def it_returns_None_when_no_title(self):
        sdt = CT_Sdt.new_block("richText")
        cc = ContentControl(sdt)
        assert cc.title is None

    def it_can_set_the_title(self):
        sdt = CT_Sdt.new_block("richText")
        cc = ContentControl(sdt)
        cc.title = "New Title"
        assert cc.title == "New Title"

    @pytest.mark.parametrize(
        ("sdt_type_str", "expected_enum"),
        [
            ("richText", WD_CONTENT_CONTROL_TYPE.RICH_TEXT),
            ("plainText", WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT),
            ("checkbox", WD_CONTENT_CONTROL_TYPE.CHECKBOX),
            ("comboBox", WD_CONTENT_CONTROL_TYPE.COMBO_BOX),
            ("dropDown", WD_CONTENT_CONTROL_TYPE.DROP_DOWN),
            ("date", WD_CONTENT_CONTROL_TYPE.DATE),
            ("picture", WD_CONTENT_CONTROL_TYPE.PICTURE),
        ],
    )
    def it_knows_its_type(self, sdt_type_str: str, expected_enum: WD_CONTENT_CONTROL_TYPE):
        sdt = CT_Sdt.new_block(sdt_type_str)
        cc = ContentControl(sdt)
        assert cc.type == expected_enum

    def it_can_get_text_from_a_block_sdt(self):
        sdt = CT_Sdt.new_block("richText")
        cc = ContentControl(sdt)
        cc.text = "Hello, World!"
        assert cc.text == "Hello, World!"

    def it_can_get_text_from_an_inline_sdt(self):
        sdt = CT_Sdt.new_inline("plainText")
        cc = ContentControl(sdt)
        cc.text = "Inline text"
        assert cc.text == "Inline text"

    def it_returns_empty_string_for_empty_content(self):
        sdt = CT_Sdt.new_block("richText")
        cc = ContentControl(sdt)
        # -- new block SDT has an empty paragraph --
        assert cc.text == ""

    def it_can_get_checked_for_checkbox(self):
        sdt = CT_Sdt.new_inline("checkbox")
        cc = ContentControl(sdt)
        assert cc.checked is False

    def it_can_set_checked_for_checkbox(self):
        sdt = CT_Sdt.new_inline("checkbox")
        cc = ContentControl(sdt)
        cc.checked = True
        assert cc.checked is True

    def it_returns_None_for_checked_on_non_checkbox(self):
        sdt = CT_Sdt.new_inline("plainText")
        cc = ContentControl(sdt)
        assert cc.checked is None

    def it_raises_on_setting_checked_on_non_checkbox(self):
        sdt = CT_Sdt.new_inline("plainText")
        cc = ContentControl(sdt)
        with pytest.raises(ValueError, match="can only set checked"):
            cc.checked = True
