"""Test suite for the docx.oxml.text.parfmt module (text-frame focus)."""

from typing import cast

import pytest

from docx.enum.text import (
    WD_FRAME_DROP_CAP,
    WD_FRAME_H_ALIGN,
    WD_FRAME_H_ANCHOR,
    WD_FRAME_V_ALIGN,
    WD_FRAME_V_ANCHOR,
    WD_FRAME_WRAP,
)
from docx.oxml.text.parfmt import CT_FramePr, CT_PPr
from docx.shared import Twips

from ...unitutil.cxml import element, xml


class DescribeCT_PPr:
    """Unit-test suite for CT_PPr's ``w:framePr`` child."""

    def it_exposes_framePr_when_present(self):
        pPr = cast(CT_PPr, element("w:pPr/w:framePr"))
        assert pPr.framePr is not None

    def it_returns_None_for_framePr_when_absent(self):
        pPr = cast(CT_PPr, element("w:pPr"))
        assert pPr.framePr is None

    def it_can_add_framePr_when_absent(self):
        pPr = cast(CT_PPr, element("w:pPr"))
        framePr = pPr.get_or_add_framePr()
        assert framePr is not None
        assert pPr.framePr is framePr

    def it_can_remove_framePr(self):
        pPr = cast(CT_PPr, element("w:pPr/w:framePr"))
        pPr._remove_framePr()
        assert pPr.framePr is None

    def it_inserts_framePr_in_correct_schema_position(self):
        # framePr must come before widowControl per ST_PPr schema order.
        pPr = cast(CT_PPr, element("w:pPr/(w:pageBreakBefore,w:widowControl)"))
        pPr.get_or_add_framePr()
        tags = [child.tag for child in pPr.iterchildren()]
        assert tags == [
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore",
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}framePr",
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}widowControl",
        ]


class DescribeCT_PPr_bidi:
    """Unit-test suite for `CT_PPr.bidi_val`."""

    def it_returns_False_when_no_bidi_child(self):
        pPr = cast(CT_PPr, element("w:pPr"))
        assert pPr.bidi_val is False

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("w:pPr/w:bidi", True),
            ("w:pPr/w:bidi{w:val=1}", True),
            ("w:pPr/w:bidi{w:val=true}", True),
            ("w:pPr/w:bidi{w:val=on}", True),
            ("w:pPr/w:bidi{w:val=0}", False),
            ("w:pPr/w:bidi{w:val=false}", False),
            ("w:pPr/w:bidi{w:val=off}", False),
        ],
    )
    def it_knows_its_bidi_val(self, pPr_cxml: str, expected_value: bool):
        pPr = cast(CT_PPr, element(pPr_cxml))
        assert pPr.bidi_val is expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("w:pPr", True, "w:pPr/w:bidi"),
            ("w:pPr/w:bidi", False, "w:pPr"),
            ("w:pPr/w:bidi", None, "w:pPr"),
            ("w:pPr/w:bidi{w:val=off}", True, "w:pPr/w:bidi"),
            ("w:pPr", False, "w:pPr"),
        ],
    )
    def it_can_change_its_bidi_val(
        self, pPr_cxml: str, value: bool | None, expected_cxml: str
    ):
        pPr = cast(CT_PPr, element(pPr_cxml))
        pPr.bidi_val = value
        assert pPr.xml == xml(expected_cxml)

    def it_inserts_bidi_in_the_right_position(self):
        # w:bidi must come after w:autoSpaceDN and before w:spacing/w:ind/w:jc
        pPr = cast(
            CT_PPr,
            element("w:pPr/(w:pStyle{w:val=Foo},w:spacing,w:ind,w:jc{w:val=center})"),
        )
        pPr.get_or_add_bidi()
        tags = [child.tag.rsplit("}", 1)[-1] for child in pPr.iterchildren()]
        assert tags == ["pStyle", "bidi", "spacing", "ind", "jc"]


class DescribeCT_FramePr:
    """Unit-test suite for the CT_FramePr (<w:framePr>) element."""

    def it_reads_twips_width_and_height(self):
        framePr = cast(
            CT_FramePr, element("w:framePr{w:w=1440,w:h=2880}")
        )
        assert framePr.w == Twips(1440)
        assert framePr.h == Twips(2880)

    def it_reads_signed_position_attributes(self):
        framePr = cast(
            CT_FramePr, element("w:framePr{w:x=720,w:y=-360}")
        )
        assert framePr.x == Twips(720)
        assert framePr.y == Twips(-360)

    @pytest.mark.parametrize(
        ("attr", "xml_val", "enum_val"),
        [
            ("hAnchor", "text", WD_FRAME_H_ANCHOR.TEXT),
            ("hAnchor", "margin", WD_FRAME_H_ANCHOR.MARGIN),
            ("hAnchor", "page", WD_FRAME_H_ANCHOR.PAGE),
            ("vAnchor", "text", WD_FRAME_V_ANCHOR.TEXT),
            ("vAnchor", "margin", WD_FRAME_V_ANCHOR.MARGIN),
            ("vAnchor", "page", WD_FRAME_V_ANCHOR.PAGE),
            ("wrap", "auto", WD_FRAME_WRAP.AUTO),
            ("wrap", "notBeside", WD_FRAME_WRAP.NOT_BESIDE),
            ("wrap", "around", WD_FRAME_WRAP.AROUND),
            ("wrap", "none", WD_FRAME_WRAP.NONE),
            ("wrap", "tight", WD_FRAME_WRAP.TIGHT),
            ("wrap", "through", WD_FRAME_WRAP.THROUGH),
            ("dropCap", "none", WD_FRAME_DROP_CAP.NONE),
            ("dropCap", "drop", WD_FRAME_DROP_CAP.DROP),
            ("dropCap", "margin", WD_FRAME_DROP_CAP.MARGIN),
            ("xAlign", "left", WD_FRAME_H_ALIGN.LEFT),
            ("xAlign", "center", WD_FRAME_H_ALIGN.CENTER),
            ("xAlign", "right", WD_FRAME_H_ALIGN.RIGHT),
            ("xAlign", "inside", WD_FRAME_H_ALIGN.INSIDE),
            ("xAlign", "outside", WD_FRAME_H_ALIGN.OUTSIDE),
            ("yAlign", "inline", WD_FRAME_V_ALIGN.INLINE),
            ("yAlign", "top", WD_FRAME_V_ALIGN.TOP),
            ("yAlign", "center", WD_FRAME_V_ALIGN.CENTER),
            ("yAlign", "bottom", WD_FRAME_V_ALIGN.BOTTOM),
            ("yAlign", "inside", WD_FRAME_V_ALIGN.INSIDE),
            ("yAlign", "outside", WD_FRAME_V_ALIGN.OUTSIDE),
        ],
    )
    def it_reads_enum_attributes(self, attr, xml_val, enum_val):
        framePr = cast(
            CT_FramePr, element(f"w:framePr{{w:{attr}={xml_val}}}")
        )
        assert getattr(framePr, attr) == enum_val

    def it_reads_lines_as_int(self):
        framePr = cast(CT_FramePr, element("w:framePr{w:lines=3}"))
        assert framePr.lines == 3

    def it_returns_None_for_absent_attrs(self):
        framePr = cast(CT_FramePr, element("w:framePr"))
        assert framePr.w is None
        assert framePr.h is None
        assert framePr.x is None
        assert framePr.y is None
        assert framePr.hAnchor is None
        assert framePr.vAnchor is None
        assert framePr.wrap is None
        assert framePr.dropCap is None
        assert framePr.lines is None
        assert framePr.xAlign is None
        assert framePr.yAlign is None
