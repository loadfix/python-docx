# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.contentcontrol` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.contentcontrol import BlockContentControl, InlineContentControl
from docx.enum.contentcontrol import WD_CONTENT_CONTROL_TYPE
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.sdt import CT_Sdt

from .unitutil.mock import Mock


class DescribeBlockContentControl:
    """Unit-test suite for `docx.contentcontrol.BlockContentControl` objects."""

    def it_can_report_its_tag(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f'  <w:sdtPr><w:tag w:val="my_tag"/></w:sdtPr>'
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.tag == "my_tag"

    def it_returns_None_for_tag_when_not_set(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.tag is None

    def it_can_report_its_title(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f'  <w:sdtPr><w:alias w:val="My Title"/></w:sdtPr>'
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.title == "My Title"

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
    def it_can_report_its_type(self, type_xml: str, expected_type: WD_CONTENT_CONTROL_TYPE):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w', 'w14')}>"
                f"  <w:sdtPr>{type_xml}</w:sdtPr>"
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.type == expected_type

    def it_can_report_its_text(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent>"
                f"    <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"
                f"  </w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.text == "Hello World"

    def it_can_report_checkbox_state(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w', 'w14')}>"
                f'  <w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>'
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.checked is True

    def it_returns_None_for_checked_when_not_checkbox(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr><w:text/></w:sdtPr>"
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        assert cc.checked is None

    def it_can_set_tag(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        cc.tag = "new_tag"

        assert cc.tag == "new_tag"

    def it_can_set_title(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent><w:p/></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = BlockContentControl(sdt, Mock())

        cc.title = "New Title"

        assert cc.title == "New Title"


class DescribeInlineContentControl:
    """Unit-test suite for `docx.contentcontrol.InlineContentControl` objects."""

    def it_can_report_its_tag(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f'  <w:sdtPr><w:tag w:val="inline_tag"/></w:sdtPr>'
                f"  <w:sdtContent><w:r><w:t>text</w:t></w:r></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        assert cc.tag == "inline_tag"

    def it_can_report_its_text(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent><w:r><w:t>Hello</w:t></w:r></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        assert cc.text == "Hello"

    def it_can_set_text(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent><w:r><w:t>old</w:t></w:r></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        cc.text = "new text"

        assert cc.text == "new text"

    def it_clears_all_children_when_setting_text(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr/>"
                f"  <w:sdtContent>"
                f"    <w:bookmarkStart w:id=\"0\" w:name=\"bm1\"/>"
                f"    <w:r><w:t>old</w:t></w:r>"
                f"    <w:bookmarkEnd w:id=\"0\"/>"
                f"  </w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        cc.text = "new text"

        assert cc.text == "new text"
        sdtContent = sdt.sdtContent
        assert len(list(sdtContent)) == 1
        assert sdtContent[0].tag == qn("w:r")

    @pytest.mark.parametrize(
        ("type_xml", "expected_type"),
        [
            ("", WD_CONTENT_CONTROL_TYPE.RICH_TEXT),
            ("<w:text/>", WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT),
        ],
    )
    def it_can_report_its_type(self, type_xml: str, expected_type: WD_CONTENT_CONTROL_TYPE):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w')}>"
                f"  <w:sdtPr>{type_xml}</w:sdtPr>"
                f"  <w:sdtContent><w:r><w:t/></w:r></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        assert cc.type == expected_type

    def it_can_report_checkbox_state(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w', 'w14')}>"
                f'  <w:sdtPr><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>'
                f"  <w:sdtContent><w:r><w:t/></w:r></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        assert cc.checked is False

    def it_can_set_checkbox_state(self):
        sdt = cast(
            CT_Sdt,
            parse_xml(
                f"<w:sdt {nsdecls('w', 'w14')}>"
                f'  <w:sdtPr><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>'
                f"  <w:sdtContent><w:r><w:t/></w:r></w:sdtContent>"
                f"</w:sdt>"
            ),
        )
        cc = InlineContentControl(sdt, Mock())

        cc.checked = True

        assert cc.checked is True
