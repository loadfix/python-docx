# pyright: reportPrivateUsage=false

"""Test suite for the docx.text.run module."""

from __future__ import annotations

from typing import cast

import pytest
from _pytest.fixtures import FixtureRequest

from docx.dml.color import ColorFormat
from docx.enum.text import WD_BORDER_STYLE, WD_COLOR, WD_UNDERLINE
from docx.oxml.text.run import CT_R
from docx.shared import Length, Pt, RGBColor
from docx.text.font import Font

from ..unitutil.cxml import element, xml
from ..unitutil.mock import Mock, class_mock, instance_mock


class DescribeFont:
    """Unit-test suite for `docx.text.font.Font`."""

    def it_provides_access_to_its_color_object(self, ColorFormat_: Mock, color_: Mock):
        r = cast(CT_R, element("w:r"))
        font = Font(r)

        color = font.color

        ColorFormat_.assert_called_once_with(font.element)
        assert color is color_

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:spacing{w:val=40}", Pt(2)),
        ],
    )
    def it_knows_its_character_spacing(self, r_cxml: str, expected_value: Length | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.character_spacing == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", Pt(2), "w:r/w:rPr/w:spacing{w:val=40}"),
            ("w:r/w:rPr", Pt(1), "w:r/w:rPr/w:spacing{w:val=20}"),
            ("w:r/w:rPr/w:spacing{w:val=40}", Pt(3), "w:r/w:rPr/w:spacing{w:val=60}"),
            ("w:r/w:rPr/w:spacing{w:val=40}", None, "w:r/w:rPr"),
        ],
    )
    def it_can_change_its_character_spacing(
        self, r_cxml: str, value: Length | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.character_spacing = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:kern{w:val=28}", Pt(14)),
        ],
    )
    def it_knows_its_kerning(self, r_cxml: str, expected_value: Length | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.kerning == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", Pt(14), "w:r/w:rPr/w:kern{w:val=28}"),
            ("w:r/w:rPr", Pt(16), "w:r/w:rPr/w:kern{w:val=32}"),
            ("w:r/w:rPr/w:kern{w:val=28}", Pt(16), "w:r/w:rPr/w:kern{w:val=32}"),
            ("w:r/w:rPr/w:kern{w:val=28}", None, "w:r/w:rPr"),
        ],
    )
    def it_can_change_its_kerning(
        self, r_cxml: str, value: Length | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.kerning = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:rFonts", None),
            ("w:r/w:rPr/w:rFonts{w:ascii=Arial}", "Arial"),
        ],
    )
    def it_knows_its_typeface_name(self, r_cxml: str, expected_value: str | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.name == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", "Foo", "w:r/w:rPr/w:rFonts{w:ascii=Foo,w:hAnsi=Foo}"),
            ("w:r/w:rPr", "Foo", "w:r/w:rPr/w:rFonts{w:ascii=Foo,w:hAnsi=Foo}"),
            (
                "w:r/w:rPr/w:rFonts{w:hAnsi=Foo}",
                "Bar",
                "w:r/w:rPr/w:rFonts{w:ascii=Bar,w:hAnsi=Bar}",
            ),
            (
                "w:r/w:rPr/w:rFonts{w:ascii=Foo,w:hAnsi=Foo}",
                "Bar",
                "w:r/w:rPr/w:rFonts{w:ascii=Bar,w:hAnsi=Bar}",
            ),
        ],
    )
    def it_can_change_its_typeface_name(self, r_cxml: str, value: str, expected_r_cxml: str):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.name = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:rFonts", None),
            ("w:r/w:rPr/w:rFonts{w:cs=Courier New}", "Courier New"),
        ],
    )
    def it_knows_its_complex_script_typeface_name(
        self, r_cxml: str, expected_value: str | None
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.name_cs == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", "Foo", "w:r/w:rPr/w:rFonts{w:cs=Foo}"),
            ("w:r/w:rPr", "Foo", "w:r/w:rPr/w:rFonts{w:cs=Foo}"),
            (
                "w:r/w:rPr/w:rFonts{w:cs=Foo}",
                "Bar",
                "w:r/w:rPr/w:rFonts{w:cs=Bar}",
            ),
            (
                "w:r/w:rPr/w:rFonts{w:ascii=Arial,w:cs=Foo}",
                "Bar",
                "w:r/w:rPr/w:rFonts{w:ascii=Arial,w:cs=Bar}",
            ),
        ],
    )
    def it_can_change_its_complex_script_typeface_name(
        self, r_cxml: str, value: str, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.name_cs = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:rFonts", None),
            ("w:r/w:rPr/w:rFonts{w:eastAsia=SimSun}", "SimSun"),
        ],
    )
    def it_knows_its_far_east_typeface_name(self, r_cxml: str, expected_value: str | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.name_far_east == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", "SimSun", "w:r/w:rPr/w:rFonts{w:eastAsia=SimSun}"),
            ("w:r/w:rPr", "SimSun", "w:r/w:rPr/w:rFonts{w:eastAsia=SimSun}"),
            (
                "w:r/w:rPr/w:rFonts{w:eastAsia=SimSun}",
                "MS Mincho",
                "w:r/w:rPr/w:rFonts{w:eastAsia=MS Mincho}",
            ),
            (
                "w:r/w:rPr/w:rFonts{w:ascii=Arial,w:eastAsia=SimSun}",
                "MS Mincho",
                "w:r/w:rPr/w:rFonts{w:ascii=Arial,w:eastAsia=MS Mincho}",
            ),
        ],
    )
    def it_can_change_its_far_east_typeface_name(
        self, r_cxml: str, value: str, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.name_far_east = value

        assert font._element.xml == expected_xml

    def it_provides_name_east_asia_as_alias_for_name_far_east(self):
        r = cast(CT_R, element("w:r/w:rPr/w:rFonts{w:eastAsia=SimSun}"))
        font = Font(r)
        assert font.name_east_asia == "SimSun"

        font.name_east_asia = "MS Mincho"
        assert font.name_far_east == "MS Mincho"

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:sz{w:val=28}", Pt(14)),
        ],
    )
    def it_knows_its_size(self, r_cxml: str, expected_value: Length | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.size == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", Pt(12), "w:r/w:rPr/w:sz{w:val=24}"),
            ("w:r/w:rPr", Pt(12), "w:r/w:rPr/w:sz{w:val=24}"),
            ("w:r/w:rPr/w:sz{w:val=24}", Pt(18), "w:r/w:rPr/w:sz{w:val=36}"),
            ("w:r/w:rPr/w:sz{w:val=36}", None, "w:r/w:rPr"),
        ],
    )
    def it_can_change_its_size(self, r_cxml: str, value: Length | None, expected_r_cxml: str):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.size = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "bool_prop_name", "expected_value"),
        [
            ("w:r/w:rPr", "all_caps", None),
            ("w:r/w:rPr/w:caps", "all_caps", True),
            ("w:r/w:rPr/w:caps{w:val=on}", "all_caps", True),
            ("w:r/w:rPr/w:caps{w:val=off}", "all_caps", False),
            ("w:r/w:rPr/w:b{w:val=1}", "bold", True),
            ("w:r/w:rPr/w:i{w:val=0}", "italic", False),
            ("w:r/w:rPr/w:cs{w:val=true}", "complex_script", True),
            ("w:r/w:rPr/w:bCs{w:val=false}", "cs_bold", False),
            ("w:r/w:rPr/w:iCs{w:val=on}", "cs_italic", True),
            ("w:r/w:rPr/w:dstrike{w:val=off}", "double_strike", False),
            ("w:r/w:rPr/w:emboss{w:val=1}", "emboss", True),
            ("w:r/w:rPr/w:vanish{w:val=0}", "hidden", False),
            ("w:r/w:rPr/w:i{w:val=true}", "italic", True),
            ("w:r/w:rPr/w:imprint{w:val=false}", "imprint", False),
            ("w:r/w:rPr/w:oMath{w:val=on}", "math", True),
            ("w:r/w:rPr/w:noProof{w:val=off}", "no_proof", False),
            ("w:r/w:rPr/w:outline{w:val=1}", "outline", True),
            ("w:r/w:rPr/w:rtl{w:val=0}", "rtl", False),
            ("w:r/w:rPr/w:shadow{w:val=true}", "shadow", True),
            ("w:r/w:rPr/w:smallCaps{w:val=false}", "small_caps", False),
            ("w:r/w:rPr/w:snapToGrid{w:val=on}", "snap_to_grid", True),
            ("w:r/w:rPr/w:specVanish{w:val=off}", "spec_vanish", False),
            ("w:r/w:rPr/w:strike{w:val=1}", "strike", True),
            ("w:r/w:rPr/w:webHidden{w:val=0}", "web_hidden", False),
        ],
    )
    def it_knows_its_bool_prop_states(
        self, r_cxml: str, bool_prop_name: str, expected_value: bool | None
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert getattr(font, bool_prop_name) == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "prop_name", "value", "expected_cxml"),
        [
            # nothing to True, False, and None ---------------------------
            ("w:r", "all_caps", True, "w:r/w:rPr/w:caps"),
            ("w:r", "bold", False, "w:r/w:rPr/w:b{w:val=0}"),
            ("w:r", "italic", None, "w:r/w:rPr"),
            # default to True, False, and None ---------------------------
            ("w:r/w:rPr/w:cs", "complex_script", True, "w:r/w:rPr/w:cs"),
            ("w:r/w:rPr/w:bCs", "cs_bold", False, "w:r/w:rPr/w:bCs{w:val=0}"),
            ("w:r/w:rPr/w:iCs", "cs_italic", None, "w:r/w:rPr"),
            # True to True, False, and None ------------------------------
            (
                "w:r/w:rPr/w:dstrike{w:val=1}",
                "double_strike",
                True,
                "w:r/w:rPr/w:dstrike",
            ),
            (
                "w:r/w:rPr/w:emboss{w:val=on}",
                "emboss",
                False,
                "w:r/w:rPr/w:emboss{w:val=0}",
            ),
            ("w:r/w:rPr/w:vanish{w:val=1}", "hidden", None, "w:r/w:rPr"),
            # False to True, False, and None -----------------------------
            ("w:r/w:rPr/w:i{w:val=false}", "italic", True, "w:r/w:rPr/w:i"),
            (
                "w:r/w:rPr/w:imprint{w:val=0}",
                "imprint",
                False,
                "w:r/w:rPr/w:imprint{w:val=0}",
            ),
            ("w:r/w:rPr/w:oMath{w:val=off}", "math", None, "w:r/w:rPr"),
            # random mix -------------------------------------------------
            (
                "w:r/w:rPr/w:noProof{w:val=1}",
                "no_proof",
                False,
                "w:r/w:rPr/w:noProof{w:val=0}",
            ),
            ("w:r/w:rPr", "outline", True, "w:r/w:rPr/w:outline"),
            ("w:r/w:rPr/w:rtl{w:val=true}", "rtl", False, "w:r/w:rPr/w:rtl{w:val=0}"),
            ("w:r/w:rPr/w:shadow{w:val=on}", "shadow", True, "w:r/w:rPr/w:shadow"),
            (
                "w:r/w:rPr/w:smallCaps",
                "small_caps",
                False,
                "w:r/w:rPr/w:smallCaps{w:val=0}",
            ),
            ("w:r/w:rPr/w:snapToGrid", "snap_to_grid", True, "w:r/w:rPr/w:snapToGrid"),
            ("w:r/w:rPr/w:specVanish", "spec_vanish", None, "w:r/w:rPr"),
            ("w:r/w:rPr/w:strike{w:val=foo}", "strike", True, "w:r/w:rPr/w:strike"),
            (
                "w:r/w:rPr/w:webHidden",
                "web_hidden",
                False,
                "w:r/w:rPr/w:webHidden{w:val=0}",
            ),
        ],
    )
    def it_can_change_its_bool_prop_settings(
        self, r_cxml: str, prop_name: str, value: bool | None, expected_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_cxml)

        setattr(font, prop_name, value)

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", False),
            ("w:r/w:rPr", False),
            ("w:r/w:rPr/w:rtl", True),
            ("w:r/w:rPr/w:rtl{w:val=1}", True),
            ("w:r/w:rPr/w:rtl{w:val=true}", True),
            ("w:r/w:rPr/w:rtl{w:val=on}", True),
            ("w:r/w:rPr/w:rtl{w:val=0}", False),
            ("w:r/w:rPr/w:rtl{w:val=false}", False),
            ("w:r/w:rPr/w:rtl{w:val=off}", False),
        ],
    )
    def it_knows_whether_it_is_right_to_left(
        self, r_cxml: str, expected_value: bool
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.right_to_left is expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_cxml"),
        [
            ("w:r", True, "w:r/w:rPr/w:rtl"),
            ("w:r/w:rPr", True, "w:r/w:rPr/w:rtl"),
            ("w:r/w:rPr/w:rtl", False, "w:r/w:rPr"),
            ("w:r/w:rPr/w:rtl", None, "w:r/w:rPr"),
            ("w:r/w:rPr/w:rtl{w:val=off}", True, "w:r/w:rPr/w:rtl"),
            ("w:r", False, "w:r/w:rPr"),
        ],
    )
    def it_can_change_whether_it_is_right_to_left(
        self, r_cxml: str, value: bool | None, expected_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        font.right_to_left = value
        assert font._element.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:vertAlign{w:val=baseline}", False),
            ("w:r/w:rPr/w:vertAlign{w:val=subscript}", True),
            ("w:r/w:rPr/w:vertAlign{w:val=superscript}", False),
        ],
    )
    def it_knows_whether_it_is_subscript(self, r_cxml: str, expected_value: bool | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.subscript == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", True, "w:r/w:rPr/w:vertAlign{w:val=subscript}"),
            ("w:r", False, "w:r/w:rPr"),
            ("w:r", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
                True,
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
            ),
            ("w:r/w:rPr/w:vertAlign{w:val=subscript}", False, "w:r/w:rPr"),
            ("w:r/w:rPr/w:vertAlign{w:val=subscript}", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
                True,
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
            ),
            (
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
                False,
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
            ),
            ("w:r/w:rPr/w:vertAlign{w:val=superscript}", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:vertAlign{w:val=baseline}",
                True,
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
            ),
        ],
    )
    def it_can_change_whether_it_is_subscript(
        self, r_cxml: str, value: bool | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.subscript = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:vertAlign{w:val=baseline}", False),
            ("w:r/w:rPr/w:vertAlign{w:val=subscript}", False),
            ("w:r/w:rPr/w:vertAlign{w:val=superscript}", True),
        ],
    )
    def it_knows_whether_it_is_superscript(self, r_cxml: str, expected_value: bool | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.superscript == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", True, "w:r/w:rPr/w:vertAlign{w:val=superscript}"),
            ("w:r", False, "w:r/w:rPr"),
            ("w:r", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
                True,
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
            ),
            ("w:r/w:rPr/w:vertAlign{w:val=superscript}", False, "w:r/w:rPr"),
            ("w:r/w:rPr/w:vertAlign{w:val=superscript}", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
                True,
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
            ),
            (
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
                False,
                "w:r/w:rPr/w:vertAlign{w:val=subscript}",
            ),
            ("w:r/w:rPr/w:vertAlign{w:val=subscript}", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:vertAlign{w:val=baseline}",
                True,
                "w:r/w:rPr/w:vertAlign{w:val=superscript}",
            ),
        ],
    )
    def it_can_change_whether_it_is_superscript(
        self, r_cxml: str, value: bool | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.superscript = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr/w:u", None),
            ("w:r/w:rPr/w:u{w:val=single}", True),
            ("w:r/w:rPr/w:u{w:val=none}", False),
            ("w:r/w:rPr/w:u{w:val=double}", WD_UNDERLINE.DOUBLE),
            ("w:r/w:rPr/w:u{w:val=wave}", WD_UNDERLINE.WAVY),
        ],
    )
    def it_knows_its_underline_type(self, r_cxml: str, expected_value: WD_UNDERLINE | bool | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.underline is expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", True, "w:r/w:rPr/w:u{w:val=single}"),
            ("w:r", False, "w:r/w:rPr/w:u{w:val=none}"),
            ("w:r", None, "w:r/w:rPr"),
            ("w:r", WD_UNDERLINE.SINGLE, "w:r/w:rPr/w:u{w:val=single}"),
            ("w:r", WD_UNDERLINE.THICK, "w:r/w:rPr/w:u{w:val=thick}"),
            ("w:r/w:rPr/w:u{w:val=single}", True, "w:r/w:rPr/w:u{w:val=single}"),
            ("w:r/w:rPr/w:u{w:val=single}", False, "w:r/w:rPr/w:u{w:val=none}"),
            ("w:r/w:rPr/w:u{w:val=single}", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:u{w:val=single}",
                WD_UNDERLINE.SINGLE,
                "w:r/w:rPr/w:u{w:val=single}",
            ),
            (
                "w:r/w:rPr/w:u{w:val=single}",
                WD_UNDERLINE.DOTTED,
                "w:r/w:rPr/w:u{w:val=dotted}",
            ),
        ],
    )
    def it_can_change_its_underline_type(
        self, r_cxml: str, value: bool | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.underline = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:highlight{w:val=default}", WD_COLOR.AUTO),
            ("w:r/w:rPr/w:highlight{w:val=blue}", WD_COLOR.BLUE),
        ],
    )
    def it_knows_its_highlight_color(self, r_cxml: str, expected_value: WD_COLOR | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.highlight_color is expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", WD_COLOR.AUTO, "w:r/w:rPr/w:highlight{w:val=default}"),
            ("w:r/w:rPr", WD_COLOR.BRIGHT_GREEN, "w:r/w:rPr/w:highlight{w:val=green}"),
            (
                "w:r/w:rPr/w:highlight{w:val=green}",
                WD_COLOR.YELLOW,
                "w:r/w:rPr/w:highlight{w:val=yellow}",
            ),
            ("w:r/w:rPr/w:highlight{w:val=yellow}", None, "w:r/w:rPr"),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r/w:rPr"),
        ],
    )
    def it_can_change_its_highlight_color(
        self, r_cxml: str, value: WD_COLOR | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.highlight_color = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:shd{w:val=clear,w:fill=D9E2F3}", RGBColor(0xD9, 0xE2, 0xF3)),
            ("w:r/w:rPr/w:shd{w:fill=FF0000}", RGBColor(0xFF, 0x00, 0x00)),
            ("w:r/w:rPr/w:shd{w:val=clear,w:fill=auto}", None),
            ("w:r/w:rPr/w:shd{w:val=clear}", None),
        ],
    )
    def it_knows_its_shading_color(
        self, r_cxml: str, expected_value: RGBColor | None
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.shading_color == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            (
                "w:r",
                RGBColor(0xD9, 0xE2, 0xF3),
                "w:r/w:rPr/w:shd{w:val=clear,w:fill=D9E2F3}",
            ),
            (
                "w:r/w:rPr",
                RGBColor(0x00, 0x00, 0xFF),
                "w:r/w:rPr/w:shd{w:val=clear,w:fill=0000FF}",
            ),
            (
                "w:r/w:rPr/w:shd{w:fill=FF0000}",
                RGBColor(0x00, 0xFF, 0x00),
                "w:r/w:rPr/w:shd{w:val=clear,w:fill=00FF00}",
            ),
            (
                "w:r/w:rPr/w:shd{w:val=clear,w:fill=D9E2F3}",
                RGBColor(0xAB, 0xCD, 0xEF),
                "w:r/w:rPr/w:shd{w:val=clear,w:fill=ABCDEF}",
            ),
            (
                "w:r/w:rPr/w:shd{w:val=clear,w:fill=D9E2F3}",
                None,
                "w:r/w:rPr",
            ),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_shading_color(
        self, r_cxml: str, value: RGBColor | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.shading_color = value

        assert font._element.xml == expected_xml

    def it_preserves_sibling_rPr_children_when_setting_shading_color(self):
        r = cast(
            CT_R,
            element("w:r/w:rPr/(w:b,w:color{w:val=FF0000},w:u{w:val=single})"),
        )
        font = Font(r)

        font.shading_color = RGBColor(0xAA, 0xBB, 0xCC)

        expected_xml = xml(
            "w:r/w:rPr/(w:b,w:color{w:val=FF0000},w:u{w:val=single},"
            "w:shd{w:val=clear,w:fill=AABBCC})"
        )
        assert font._element.xml == expected_xml

    # -- run border (w:bdr) ------------------------------------------

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:bdr", None),
            ("w:r/w:rPr/w:bdr{w:val=single}", WD_BORDER_STYLE.SINGLE),
            ("w:r/w:rPr/w:bdr{w:val=double}", WD_BORDER_STYLE.DOUBLE),
        ],
    )
    def it_knows_its_border_style(
        self, r_cxml: str, expected_value: WD_BORDER_STYLE | None
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.border_style == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", WD_BORDER_STYLE.SINGLE, "w:r/w:rPr/w:bdr{w:val=single}"),
            ("w:r/w:rPr", WD_BORDER_STYLE.DOUBLE, "w:r/w:rPr/w:bdr{w:val=double}"),
            (
                "w:r/w:rPr/w:bdr{w:val=single}",
                WD_BORDER_STYLE.DOUBLE,
                "w:r/w:rPr/w:bdr{w:val=double}",
            ),
            ("w:r/w:rPr/w:bdr{w:val=single}", None, "w:r/w:rPr/w:bdr"),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:sz=8}",
                None,
                "w:r/w:rPr/w:bdr{w:sz=8}",
            ),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_border_style(
        self,
        r_cxml: str,
        value: WD_BORDER_STYLE | None,
        expected_r_cxml: str,
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.border_style = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:bdr", None),
            ("w:r/w:rPr/w:bdr{w:val=single,w:sz=8}", Pt(1)),
            ("w:r/w:rPr/w:bdr{w:val=single,w:sz=24}", Pt(3)),
        ],
    )
    def it_knows_its_border_width(self, r_cxml: str, expected_value: Length | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.border_width == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", Pt(1), "w:r/w:rPr/w:bdr{w:sz=8}"),
            ("w:r/w:rPr", Pt(2), "w:r/w:rPr/w:bdr{w:sz=16}"),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:sz=8}",
                Pt(3),
                "w:r/w:rPr/w:bdr{w:val=single,w:sz=24}",
            ),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:sz=8}",
                None,
                "w:r/w:rPr/w:bdr{w:val=single}",
            ),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_border_width(
        self, r_cxml: str, value: Length | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.border_width = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:bdr", None),
            ("w:r/w:rPr/w:bdr{w:val=single,w:color=FF0000}", RGBColor(0xFF, 0, 0)),
            ("w:r/w:rPr/w:bdr{w:val=single,w:color=auto}", None),
        ],
    )
    def it_knows_its_border_color(
        self, r_cxml: str, expected_value: RGBColor | None
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.border_color == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", RGBColor(0xFF, 0, 0), "w:r/w:rPr/w:bdr{w:color=FF0000}"),
            (
                "w:r/w:rPr",
                RGBColor(0, 0, 0xFF),
                "w:r/w:rPr/w:bdr{w:color=0000FF}",
            ),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:color=FF0000}",
                RGBColor(0, 0xFF, 0),
                "w:r/w:rPr/w:bdr{w:val=single,w:color=00FF00}",
            ),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:color=FF0000}",
                None,
                "w:r/w:rPr/w:bdr{w:val=single}",
            ),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_border_color(
        self, r_cxml: str, value: RGBColor | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.border_color = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:bdr", None),
            ("w:r/w:rPr/w:bdr{w:val=single,w:space=4}", Pt(4)),
        ],
    )
    def it_knows_its_border_space(self, r_cxml: str, expected_value: Length | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.border_space == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", Pt(4), "w:r/w:rPr/w:bdr{w:space=4}"),
            ("w:r/w:rPr", Pt(8), "w:r/w:rPr/w:bdr{w:space=8}"),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:space=4}",
                Pt(8),
                "w:r/w:rPr/w:bdr{w:val=single,w:space=8}",
            ),
            (
                "w:r/w:rPr/w:bdr{w:val=single,w:space=4}",
                None,
                "w:r/w:rPr/w:bdr{w:val=single}",
            ),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_border_space(
        self, r_cxml: str, value: Length | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.border_space = value

        assert font._element.xml == expected_xml

    def it_can_set_all_run_border_properties_at_once(self):
        r = cast(CT_R, element("w:r"))
        font = Font(r)

        font.border_style = WD_BORDER_STYLE.SINGLE
        font.border_width = Pt(1)
        font.border_space = Pt(4)
        font.border_color = RGBColor(0x4F, 0x81, 0xBD)

        expected_xml = xml(
            "w:r/w:rPr/w:bdr{w:val=single,w:sz=8,w:space=4,w:color=4F81BD}"
        )
        assert font._element.xml == expected_xml

    def it_can_remove_the_run_border(self):
        r = cast(
            CT_R,
            element(
                "w:r/w:rPr/w:bdr{w:val=single,w:sz=8,w:space=4,w:color=4F81BD}"
            ),
        )
        font = Font(r)

        font.remove_border()

        assert font._element.xml == xml("w:r/w:rPr")
        assert font.border_style is None
        assert font.border_width is None
        assert font.border_color is None
        assert font.border_space is None

    def it_does_not_raise_when_removing_border_that_is_absent(self):
        r = cast(CT_R, element("w:r"))
        font = Font(r)

        font.remove_border()  # should be a no-op

        assert font._element.xml == xml("w:r")

    def it_does_not_raise_when_removing_border_when_rPr_has_no_bdr(self):
        r = cast(CT_R, element("w:r/w:rPr/w:b"))
        font = Font(r)

        font.remove_border()

        assert font._element.xml == xml("w:r/w:rPr/w:b")

    def it_preserves_sibling_rPr_children_when_setting_border(self):
        r = cast(
            CT_R,
            element("w:r/w:rPr/(w:b,w:color{w:val=FF0000},w:u{w:val=single})"),
        )
        font = Font(r)

        font.border_style = WD_BORDER_STYLE.SINGLE
        font.border_width = Pt(1)
        font.border_color = RGBColor(0xAA, 0xBB, 0xCC)

        expected_xml = xml(
            "w:r/w:rPr/(w:b,w:color{w:val=FF0000},w:u{w:val=single},"
            "w:bdr{w:val=single,w:sz=8,w:color=AABBCC})"
        )
        assert font._element.xml == expected_xml

    def it_preserves_correct_schema_order_when_bdr_has_later_sibling(self):
        """w:bdr must precede w:shd in CT_RPr (schema order)."""
        r = cast(
            CT_R,
            element("w:r/w:rPr/w:shd{w:val=clear,w:fill=AABBCC}"),
        )
        font = Font(r)

        font.border_style = WD_BORDER_STYLE.SINGLE

        expected_xml = xml(
            "w:r/w:rPr/(w:bdr{w:val=single},w:shd{w:val=clear,w:fill=AABBCC})"
        )
        assert font._element.xml == expected_xml

    # -- language (w:lang) -------------------------------------------

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:lang", None),
            ("w:r/w:rPr/w:lang{w:val=en-US}", "en-US"),
            ("w:r/w:rPr/w:lang{w:val=fr-FR,w:eastAsia=ja-JP}", "fr-FR"),
        ],
    )
    def it_knows_its_language(self, r_cxml: str, expected_value: str | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.language == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", "en-US", "w:r/w:rPr/w:lang{w:val=en-US}"),
            ("w:r/w:rPr", "fr-FR", "w:r/w:rPr/w:lang{w:val=fr-FR}"),
            (
                "w:r/w:rPr/w:lang{w:val=en-US}",
                "fr-FR",
                "w:r/w:rPr/w:lang{w:val=fr-FR}",
            ),
            (
                "w:r/w:rPr/w:lang{w:val=en-US,w:eastAsia=ja-JP}",
                None,
                "w:r/w:rPr/w:lang{w:eastAsia=ja-JP}",
            ),
            ("w:r/w:rPr/w:lang{w:val=en-US}", None, "w:r/w:rPr/w:lang"),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_language(
        self, r_cxml: str, value: str | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.language = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:lang", None),
            ("w:r/w:rPr/w:lang{w:eastAsia=ja-JP}", "ja-JP"),
            ("w:r/w:rPr/w:lang{w:val=en-US,w:eastAsia=zh-CN}", "zh-CN"),
        ],
    )
    def it_knows_its_east_asian_language(
        self, r_cxml: str, expected_value: str | None
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.east_asian_language == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", "ja-JP", "w:r/w:rPr/w:lang{w:eastAsia=ja-JP}"),
            ("w:r/w:rPr", "zh-CN", "w:r/w:rPr/w:lang{w:eastAsia=zh-CN}"),
            (
                "w:r/w:rPr/w:lang{w:eastAsia=ja-JP}",
                "zh-CN",
                "w:r/w:rPr/w:lang{w:eastAsia=zh-CN}",
            ),
            (
                "w:r/w:rPr/w:lang{w:val=en-US,w:eastAsia=ja-JP}",
                None,
                "w:r/w:rPr/w:lang{w:val=en-US}",
            ),
            ("w:r/w:rPr/w:lang{w:eastAsia=ja-JP}", None, "w:r/w:rPr/w:lang"),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_east_asian_language(
        self, r_cxml: str, value: str | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.east_asian_language = value

        assert font._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:lang", None),
            ("w:r/w:rPr/w:lang{w:bidi=ar-SA}", "ar-SA"),
            ("w:r/w:rPr/w:lang{w:val=en-US,w:bidi=he-IL}", "he-IL"),
        ],
    )
    def it_knows_its_bidi_language(self, r_cxml: str, expected_value: str | None):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        assert font.bidi_language == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "value", "expected_r_cxml"),
        [
            ("w:r", "ar-SA", "w:r/w:rPr/w:lang{w:bidi=ar-SA}"),
            ("w:r/w:rPr", "he-IL", "w:r/w:rPr/w:lang{w:bidi=he-IL}"),
            (
                "w:r/w:rPr/w:lang{w:bidi=ar-SA}",
                "he-IL",
                "w:r/w:rPr/w:lang{w:bidi=he-IL}",
            ),
            (
                "w:r/w:rPr/w:lang{w:val=en-US,w:bidi=ar-SA}",
                None,
                "w:r/w:rPr/w:lang{w:val=en-US}",
            ),
            ("w:r/w:rPr/w:lang{w:bidi=ar-SA}", None, "w:r/w:rPr/w:lang"),
            ("w:r/w:rPr", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_bidi_language(
        self, r_cxml: str, value: str | None, expected_r_cxml: str
    ):
        r = cast(CT_R, element(r_cxml))
        font = Font(r)
        expected_xml = xml(expected_r_cxml)

        font.bidi_language = value

        assert font._element.xml == expected_xml

    def it_can_set_all_language_tags_at_once(self):
        r = cast(CT_R, element("w:r"))
        font = Font(r)

        font.language = "en-US"
        font.east_asian_language = "ja-JP"
        font.bidi_language = "ar-SA"

        expected_xml = xml(
            "w:r/w:rPr/w:lang{w:val=en-US,w:eastAsia=ja-JP,w:bidi=ar-SA}"
        )
        assert font._element.xml == expected_xml

    def it_can_remove_the_language_element(self):
        r = cast(
            CT_R,
            element("w:r/w:rPr/w:lang{w:val=en-US,w:eastAsia=ja-JP,w:bidi=ar-SA}"),
        )
        font = Font(r)

        font.remove_language()

        assert font._element.xml == xml("w:r/w:rPr")
        assert font.language is None
        assert font.east_asian_language is None
        assert font.bidi_language is None

    def it_does_not_raise_when_removing_language_that_is_absent(self):
        r = cast(CT_R, element("w:r"))
        font = Font(r)

        font.remove_language()  # no-op

        assert font._element.xml == xml("w:r")

    def it_does_not_raise_when_removing_language_when_rPr_has_no_lang(self):
        r = cast(CT_R, element("w:r/w:rPr/w:b"))
        font = Font(r)

        font.remove_language()

        assert font._element.xml == xml("w:r/w:rPr/w:b")

    def it_preserves_sibling_rPr_children_when_setting_language(self):
        r = cast(
            CT_R,
            element("w:r/w:rPr/(w:b,w:color{w:val=FF0000},w:u{w:val=single})"),
        )
        font = Font(r)

        font.language = "en-US"

        expected_xml = xml(
            "w:r/w:rPr/(w:b,w:color{w:val=FF0000},w:u{w:val=single},"
            "w:lang{w:val=en-US})"
        )
        assert font._element.xml == expected_xml

    def it_preserves_correct_schema_order_when_lang_has_later_sibling(self):
        """w:lang must precede w:oMath in CT_RPr (schema order)."""
        r = cast(CT_R, element("w:r/w:rPr/w:oMath"))
        font = Font(r)

        font.language = "en-US"

        expected_xml = xml("w:r/w:rPr/(w:lang{w:val=en-US},w:oMath)")
        assert font._element.xml == expected_xml

    # -- fixtures ----------------------------------------------------

    @pytest.fixture
    def color_(self, request: FixtureRequest):
        return instance_mock(request, ColorFormat)

    @pytest.fixture
    def ColorFormat_(self, request: FixtureRequest, color_: Mock):
        return class_mock(request, "docx.text.font.ColorFormat", return_value=color_)
