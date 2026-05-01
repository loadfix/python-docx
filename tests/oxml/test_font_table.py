# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.font_table` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.font_table import CT_Font, CT_Fonts

from ..unitutil.cxml import element


class DescribeCT_Fonts:
    """Unit-test suite for `docx.oxml.font_table.CT_Fonts`."""

    def it_exposes_its_fonts_as_a_list(self):
        fonts = cast(CT_Fonts, element("w:fonts"))
        assert fonts.font_lst == []

    def it_enumerates_font_children_in_xml_order(self):
        fonts = cast(
            CT_Fonts,
            element(
                "w:fonts/("
                "w:font{w:name=Arial},"
                "w:font{w:name=Calibri},"
                "w:font{w:name=Times New Roman}"
                ")"
            ),
        )
        assert [f.name for f in fonts.font_lst] == [
            "Arial",
            "Calibri",
            "Times New Roman",
        ]

    def it_can_find_a_font_by_name(self):
        fonts = cast(
            CT_Fonts,
            element("w:fonts/(w:font{w:name=Arial},w:font{w:name=Calibri})"),
        )

        font = fonts.get_font_by_name("Calibri")

        assert font is not None
        assert font.name == "Calibri"

    def but_it_returns_None_when_the_named_font_is_not_present(self):
        fonts = cast(CT_Fonts, element("w:fonts/(w:font{w:name=Arial})"))
        assert fonts.get_font_by_name("Helvetica") is None


class DescribeCT_Font:
    """Unit-test suite for `docx.oxml.font_table.CT_Font`."""

    def it_exposes_its_name_attribute(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert font.name == "Arial"

    def it_exposes_altName_charset_family_pitch_and_panose_children(self):
        font = cast(
            CT_Font,
            element(
                "w:font{w:name=Arial}/("
                "w:altName{w:val=Helvetica},"
                "w:panose1{w:val=020B0604020202020204},"
                "w:charset{w:val=00},"
                "w:family{w:val=swiss},"
                "w:pitch{w:val=variable}"
                ")"
            ),
        )
        assert font.altName is not None
        assert font.altName.val == "Helvetica"
        assert font.panose1 is not None
        assert font.panose1.val == "020B0604020202020204"
        assert font.charset is not None
        assert font.charset.val == "00"
        assert font.family is not None
        assert font.family.val == "swiss"
        assert font.pitch is not None
        assert font.pitch.val == "variable"

    def its_optional_child_elements_are_None_when_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert font.altName is None
        assert font.panose1 is None
        assert font.charset is None
        assert font.family is None
        assert font.pitch is None
        assert font.embedRegular is None
        assert font.embedBold is None
        assert font.embedItalic is None
        assert font.embedBoldItalic is None

    def it_exposes_embed_elements_for_embedded_fonts(self):
        font = cast(
            CT_Font,
            element(
                "w:font{w:name=Arial}/("
                "w:embedRegular{r:id=rId1},"
                "w:embedBold{r:id=rId2},"
                "w:embedItalic{r:id=rId3},"
                "w:embedBoldItalic{r:id=rId4}"
                ")"
            ),
        )
        assert font.embedRegular is not None
        assert font.embedRegular.rId == "rId1"
        assert font.embedBold is not None
        assert font.embedBold.rId == "rId2"
        assert font.embedItalic is not None
        assert font.embedItalic.rId == "rId3"
        assert font.embedBoldItalic is not None
        assert font.embedBoldItalic.rId == "rId4"
