# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.theme` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.theme import (
    CT_ClrScheme,
    CT_ColorChoice,
    CT_FontScheme,
    CT_Theme,
)
from docx.shared import RGBColor

from ..unitutil.cxml import element


class DescribeCT_Theme:
    """Unit-test suite for `docx.oxml.theme.CT_Theme`."""

    def it_exposes_its_name_attribute(self):
        theme = cast(CT_Theme, element("a:theme{name=Office Theme}"))
        assert theme.name == "Office Theme"

    def it_returns_None_for_missing_name_attribute(self):
        theme = cast(CT_Theme, element("a:theme"))
        assert theme.name is None

    def it_exposes_the_nested_clrScheme(self):
        theme = cast(
            CT_Theme, element("a:theme/a:themeElements/a:clrScheme")
        )
        assert theme.clrScheme is not None
        assert isinstance(theme.clrScheme, CT_ClrScheme)

    def it_returns_None_when_themeElements_is_absent(self):
        theme = cast(CT_Theme, element("a:theme"))
        assert theme.clrScheme is None
        assert theme.fontScheme is None

    def it_returns_None_when_clrScheme_is_absent(self):
        theme = cast(CT_Theme, element("a:theme/a:themeElements"))
        assert theme.clrScheme is None
        assert theme.fontScheme is None

    def it_exposes_the_nested_fontScheme(self):
        theme = cast(
            CT_Theme, element("a:theme/a:themeElements/a:fontScheme")
        )
        assert theme.fontScheme is not None
        assert isinstance(theme.fontScheme, CT_FontScheme)


class DescribeCT_ClrScheme:
    """Unit-test suite for `docx.oxml.theme.CT_ClrScheme`."""

    @pytest.mark.parametrize(
        ("slot", "cxml"),
        [
            ("dk1", "a:clrScheme/a:dk1/a:srgbClr{val=010203}"),
            ("lt1", "a:clrScheme/a:lt1/a:srgbClr{val=F0E0D0}"),
            ("accent1", "a:clrScheme/a:accent1/a:srgbClr{val=5B9BD5}"),
            ("hlink", "a:clrScheme/a:hlink/a:srgbClr{val=0563C1}"),
            ("folHlink", "a:clrScheme/a:folHlink/a:srgbClr{val=954F72}"),
        ],
    )
    def it_exposes_each_slot_as_a_color_choice(self, slot: str, cxml: str):
        scheme = cast(CT_ClrScheme, element(cxml))
        choice = getattr(scheme, slot)
        assert isinstance(choice, CT_ColorChoice)

    def it_returns_None_for_absent_slots(self):
        scheme = cast(CT_ClrScheme, element("a:clrScheme"))
        for slot in (
            "dk1",
            "lt1",
            "dk2",
            "lt2",
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hlink",
            "folHlink",
        ):
            assert getattr(scheme, slot) is None

    def it_looks_up_slots_by_name(self):
        scheme = cast(
            CT_ClrScheme,
            element("a:clrScheme/a:accent1/a:srgbClr{val=5B9BD5}"),
        )
        choice = scheme.color_for("accent1")
        assert choice is not None
        assert choice.rgb == RGBColor.from_string("5B9BD5")

    def it_returns_None_for_unknown_color_names(self):
        scheme = cast(CT_ClrScheme, element("a:clrScheme"))
        assert scheme.color_for("bogus") is None


class DescribeCT_ColorChoice:
    """Unit-test suite for `docx.oxml.theme.CT_ColorChoice`."""

    def it_resolves_an_srgbClr_directly(self):
        dk2 = cast(CT_ColorChoice, element("a:dk2/a:srgbClr{val=44546A}"))
        assert dk2.rgb == RGBColor.from_string("44546A")

    def it_resolves_a_sysClr_via_its_lastClr_fallback(self):
        dk1 = cast(
            CT_ColorChoice,
            element("a:dk1/a:sysClr{val=windowText,lastClr=000000}"),
        )
        assert dk1.rgb == RGBColor.from_string("000000")

    def it_prefers_srgbClr_over_sysClr_when_both_are_present(self):
        # Schema-wise this is unusual, but defensively ensure the fast path
        # wins so we don't silently return a stale lastClr.
        dk1 = cast(
            CT_ColorChoice,
            element(
                "a:dk1/("
                "a:srgbClr{val=123456},"
                "a:sysClr{val=windowText,lastClr=FFFFFF}"
                ")"
            ),
        )
        assert dk1.rgb == RGBColor.from_string("123456")

    def it_returns_None_when_sysClr_has_no_lastClr(self):
        dk1 = cast(CT_ColorChoice, element("a:dk1/a:sysClr{val=windowText}"))
        assert dk1.rgb is None

    def it_returns_None_when_no_supported_color_child_is_present(self):
        dk1 = cast(CT_ColorChoice, element("a:dk1"))
        assert dk1.rgb is None


class DescribeCT_FontScheme:
    """Unit-test suite for `docx.oxml.theme.CT_FontScheme`."""

    def it_exposes_majorFont_and_minorFont_and_name(self):
        scheme = cast(
            CT_FontScheme,
            element(
                "a:fontScheme{name=Office}/("
                "a:majorFont/a:latin{typeface=Calibri Light},"
                "a:minorFont/a:latin{typeface=Calibri}"
                ")"
            ),
        )
        assert scheme.name == "Office"
        assert scheme.majorFont is not None
        assert scheme.minorFont is not None
        assert scheme.majorFont.latin is not None
        assert scheme.majorFont.latin.typeface == "Calibri Light"
        assert scheme.minorFont.latin is not None
        assert scheme.minorFont.latin.typeface == "Calibri"

    def it_returns_None_for_missing_children(self):
        scheme = cast(CT_FontScheme, element("a:fontScheme"))
        assert scheme.majorFont is None
        assert scheme.minorFont is None
        assert scheme.name is None
