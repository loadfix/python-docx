# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.dml.color` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.dml.color import ColorFormat
from docx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from docx.oxml.text.run import CT_R
from docx.shared import RGBColor

from ..unitutil.cxml import element, xml


class DescribeColorFormat:
    """Unit-test suite for `docx.dml.color.ColorFormat` objects."""

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:color{w:val=auto}", MSO_COLOR_TYPE.AUTO),
            ("w:r/w:rPr/w:color{w:val=4224FF}", MSO_COLOR_TYPE.RGB),
            ("w:r/w:rPr/w:color{w:themeColor=dark1}", MSO_COLOR_TYPE.THEME),
            (
                "w:r/w:rPr/w:color{w:val=F00BA9,w:themeColor=accent1}",
                MSO_COLOR_TYPE.THEME,
            ),
        ],
    )
    def it_knows_its_color_type(self, r_cxml: str, expected_value: MSO_COLOR_TYPE | None):
        assert ColorFormat(cast(CT_R, element(r_cxml))).type == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "rgb"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:color{w:val=auto}", None),
            ("w:r/w:rPr/w:color{w:val=4224FF}", "4224ff"),
            ("w:r/w:rPr/w:color{w:val=auto,w:themeColor=accent1}", None),
            ("w:r/w:rPr/w:color{w:val=F00BA9,w:themeColor=accent1}", "f00ba9"),
        ],
    )
    def it_knows_its_RGB_value(self, r_cxml: str, rgb: str | None):
        expected_value = RGBColor.from_string(rgb) if rgb else None
        assert ColorFormat(cast(CT_R, element(r_cxml))).rgb == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "new_value", "expected_cxml"),
        [
            ("w:r", RGBColor(10, 20, 30), "w:r/w:rPr/w:color{w:val=0A141E}"),
            ("w:r/w:rPr", RGBColor(1, 2, 3), "w:r/w:rPr/w:color{w:val=010203}"),
            (
                "w:r/w:rPr/w:color{w:val=123abc}",
                RGBColor(42, 24, 99),
                "w:r/w:rPr/w:color{w:val=2A1863}",
            ),
            (
                "w:r/w:rPr/w:color{w:val=auto}",
                RGBColor(16, 17, 18),
                "w:r/w:rPr/w:color{w:val=101112}",
            ),
            (
                "w:r/w:rPr/w:color{w:val=234bcd,w:themeColor=dark1}",
                RGBColor(24, 42, 99),
                "w:r/w:rPr/w:color{w:val=182A63}",
            ),
            ("w:r/w:rPr/w:color{w:val=234bcd,w:themeColor=dark1}", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_RGB_value(
        self, r_cxml: str, new_value: RGBColor | None, expected_cxml: str
    ):
        color_format = ColorFormat(cast(CT_R, element(r_cxml)))
        color_format.rgb = new_value
        assert color_format._element.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr", None),
            ("w:r/w:rPr/w:color{w:val=auto}", None),
            ("w:r/w:rPr/w:color{w:val=4224FF}", None),
            ("w:r/w:rPr/w:color{w:themeColor=accent1}", MSO_THEME_COLOR.ACCENT_1),
            ("w:r/w:rPr/w:color{w:val=F00BA9,w:themeColor=dark1}", MSO_THEME_COLOR.DARK_1),
        ],
    )
    def it_knows_its_theme_color(self, r_cxml: str, expected_value: MSO_THEME_COLOR | None):
        color_format = ColorFormat(cast(CT_R, element(r_cxml)))
        assert color_format.theme_color == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "new_value", "expected_cxml"),
        [
            (
                "w:r",
                MSO_THEME_COLOR.ACCENT_1,
                "w:r/w:rPr/w:color{w:val=000000,w:themeColor=accent1}",
            ),
            (
                "w:r/w:rPr",
                MSO_THEME_COLOR.ACCENT_2,
                "w:r/w:rPr/w:color{w:val=000000,w:themeColor=accent2}",
            ),
            (
                "w:r/w:rPr/w:color{w:val=101112}",
                MSO_THEME_COLOR.ACCENT_3,
                "w:r/w:rPr/w:color{w:val=101112,w:themeColor=accent3}",
            ),
            (
                "w:r/w:rPr/w:color{w:val=234bcd,w:themeColor=dark1}",
                MSO_THEME_COLOR.LIGHT_2,
                "w:r/w:rPr/w:color{w:val=234bcd,w:themeColor=light2}",
            ),
            ("w:r/w:rPr/w:color{w:val=234bcd,w:themeColor=dark1}", None, "w:r/w:rPr"),
            ("w:r", None, "w:r"),
        ],
    )
    def it_can_change_its_theme_color(
        self, r_cxml: str, new_value: MSO_THEME_COLOR | None, expected_cxml: str
    ):
        color_format = ColorFormat(cast(CT_R, element(r_cxml)))
        color_format.theme_color = new_value
        assert color_format._element.xml == xml(expected_cxml)


class DescribeColorFormat_Brightness:
    """Phase A-v2 #7: ColorFormat.brightness implementation.

    See upstream #665 — the documented property used to raise AttributeError.
    """

    def it_returns_zero_when_no_color_is_set(self):
        cf = ColorFormat(cast(CT_R, element("w:r")))
        assert cf.brightness == 0.0

    def it_returns_zero_when_no_tint_or_shade(self):
        cf = ColorFormat(
            cast(CT_R, element("w:r/w:rPr/w:color{w:themeColor=accent1}"))
        )
        assert cf.brightness == 0.0

    def it_reads_a_positive_brightness_from_themeTint(self):
        # -- themeTint=7F (≈127/255) → brightness ≈ 1 - 127/255 ≈ 0.502 --
        cf = ColorFormat(
            cast(
                CT_R,
                element(
                    "w:r/w:rPr/w:color{w:themeColor=accent1,w:themeTint=7F}"
                ),
            )
        )
        assert 0.49 < cf.brightness < 0.51

    def it_reads_a_negative_brightness_from_themeShade(self):
        # -- themeShade=BF (≈191/255) → brightness ≈ 191/255 - 1 ≈ -0.251 --
        cf = ColorFormat(
            cast(
                CT_R,
                element(
                    "w:r/w:rPr/w:color{w:themeColor=accent1,w:themeShade=BF}"
                ),
            )
        )
        assert -0.26 < cf.brightness < -0.24

    def it_writes_themeTint_for_positive_brightness(self):
        from docx.oxml.ns import qn

        r = cast(
            CT_R, element("w:r/w:rPr/w:color{w:val=000000,w:themeColor=accent1}")
        )
        cf = ColorFormat(r)
        cf.brightness = 0.5
        color_elm = r.find(qn("w:rPr")).find(qn("w:color"))
        tint = color_elm.get(qn("w:themeTint"))
        assert tint is not None
        assert 0x7E <= int(tint, 16) <= 0x80
        assert color_elm.get(qn("w:themeShade")) is None

    def it_writes_themeShade_for_negative_brightness(self):
        from docx.oxml.ns import qn

        r = cast(
            CT_R, element("w:r/w:rPr/w:color{w:val=000000,w:themeColor=accent1}")
        )
        cf = ColorFormat(r)
        cf.brightness = -0.25
        color_elm = r.find(qn("w:rPr")).find(qn("w:color"))
        shade = color_elm.get(qn("w:themeShade"))
        assert shade is not None
        assert 0xBE <= int(shade, 16) <= 0xC0
        assert color_elm.get(qn("w:themeTint")) is None

    def it_clears_tint_and_shade_when_brightness_zero(self):
        from docx.oxml.ns import qn

        r = cast(
            CT_R,
            element(
                "w:r/w:rPr/w:color{w:val=000000,w:themeColor=accent1,"
                "w:themeTint=7F}"
            ),
        )
        cf = ColorFormat(r)
        cf.brightness = 0.0
        color_elm = r.find(qn("w:rPr")).find(qn("w:color"))
        assert color_elm.get(qn("w:themeTint")) is None
        assert color_elm.get(qn("w:themeShade")) is None

    def it_rejects_out_of_range_brightness(self):
        cf = ColorFormat(
            cast(
                CT_R,
                element("w:r/w:rPr/w:color{w:val=000000,w:themeColor=accent1}"),
            )
        )
        with pytest.raises(ValueError, match="-1.0 .. \\+1.0"):
            cf.brightness = 1.5

    def it_rejects_brightness_assignment_without_theme_color(self):
        cf = ColorFormat(cast(CT_R, element("w:r/w:rPr/w:color{w:val=000000}")))
        with pytest.raises(ValueError, match="theme color"):
            cf.brightness = 0.5
