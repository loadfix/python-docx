# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.font_table` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.font_table import FontMetadata, FontTable
from docx.oxml.font_table import CT_Font, CT_Fonts
from docx.parts.font_table import FontTablePart

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeFontTable:
    """Unit-test suite for `docx.font_table.FontTable`."""

    def it_knows_its_length(self, font_table_part_: Mock):
        empty_fonts = cast(CT_Fonts, element("w:fonts"))
        assert len(FontTable(empty_fonts, font_table_part_)) == 0

        fonts = cast(
            CT_Fonts,
            element("w:fonts/(w:font{w:name=Arial},w:font{w:name=Calibri})"),
        )
        assert len(FontTable(fonts, font_table_part_)) == 2

    def it_iterates_FontMetadata_in_xml_order(self, font_table_part_: Mock):
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

        names = [f.name for f in FontTable(fonts, font_table_part_)]

        assert names == ["Arial", "Calibri", "Times New Roman"]

    def it_supports_membership_testing_by_font_name(self, font_table_part_: Mock):
        fonts = cast(CT_Fonts, element("w:fonts/(w:font{w:name=Arial})"))
        font_table = FontTable(fonts, font_table_part_)

        assert "Arial" in font_table
        assert "Helvetica" not in font_table

    def it_returns_False_for_non_string_membership_tests(self, font_table_part_: Mock):
        fonts = cast(CT_Fonts, element("w:fonts/(w:font{w:name=Arial})"))
        font_table = FontTable(fonts, font_table_part_)

        assert 42 not in font_table

    def it_supports_indexing_by_font_name(self, font_table_part_: Mock):
        fonts = cast(
            CT_Fonts,
            element("w:fonts/(w:font{w:name=Arial},w:font{w:name=Calibri})"),
        )
        font_table = FontTable(fonts, font_table_part_)

        font = font_table["Calibri"]

        assert isinstance(font, FontMetadata)
        assert font.name == "Calibri"

    def but_indexing_raises_KeyError_when_missing(self, font_table_part_: Mock):
        fonts = cast(CT_Fonts, element("w:fonts/(w:font{w:name=Arial})"))
        font_table = FontTable(fonts, font_table_part_)

        with pytest.raises(KeyError):
            font_table["Helvetica"]

    def it_can_look_up_a_font_with_get(self, font_table_part_: Mock):
        fonts = cast(
            CT_Fonts,
            element("w:fonts/(w:font{w:name=Arial},w:font{w:name=Calibri})"),
        )
        font_table = FontTable(fonts, font_table_part_)

        font = font_table.get("Calibri")

        assert isinstance(font, FontMetadata)
        assert font.name == "Calibri"

    def and_get_returns_None_when_the_named_font_is_absent(self, font_table_part_: Mock):
        fonts = cast(CT_Fonts, element("w:fonts/(w:font{w:name=Arial})"))
        font_table = FontTable(fonts, font_table_part_)

        assert font_table.get("Helvetica") is None

    def it_exposes_its_underlying_element_and_part(self, font_table_part_: Mock):
        fonts = cast(CT_Fonts, element("w:fonts"))
        font_table = FontTable(fonts, font_table_part_)

        assert font_table.element is fonts
        assert font_table.part is font_table_part_

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def font_table_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, FontTablePart)


class DescribeFontMetadata:
    """Unit-test suite for `docx.font_table.FontMetadata`."""

    def it_exposes_its_name(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).name == "Arial"

    def it_exposes_its_family(self):
        font = cast(
            CT_Font,
            element("w:font{w:name=Arial}/(w:family{w:val=swiss})"),
        )
        assert FontMetadata(font).family == "swiss"

    def and_family_is_None_when_the_family_child_is_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).family is None

    def it_exposes_its_charset(self):
        font = cast(
            CT_Font,
            element("w:font{w:name=Arial}/(w:charset{w:val=00})"),
        )
        assert FontMetadata(font).charset == "00"

    def and_charset_is_None_when_the_charset_child_is_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).charset is None

    def it_exposes_its_pitch(self):
        font = cast(
            CT_Font,
            element("w:font{w:name=Arial}/(w:pitch{w:val=variable})"),
        )
        assert FontMetadata(font).pitch == "variable"

    def and_pitch_is_None_when_the_pitch_child_is_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).pitch is None

    def it_exposes_its_panose_as_a_20_character_hex_string(self):
        font = cast(
            CT_Font,
            element("w:font{w:name=Arial}/(w:panose1{w:val=020B0604020202020204})"),
        )
        assert FontMetadata(font).panose == "020B0604020202020204"

    def and_panose_is_None_when_the_panose1_child_is_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).panose is None

    def it_exposes_its_alt_name(self):
        font = cast(
            CT_Font,
            element("w:font{w:name=Arial}/(w:altName{w:val=Helvetica})"),
        )
        assert FontMetadata(font).alt_name == "Helvetica"

    def and_alt_name_is_None_when_the_altName_child_is_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).alt_name is None

    def it_reports_embed_flags_as_False_when_embed_children_are_absent(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        metadata = FontMetadata(font)

        assert metadata.embed_regular is False
        assert metadata.embed_bold is False
        assert metadata.embed_italic is False
        assert metadata.embed_bold_italic is False

    def it_reports_embed_flags_as_True_when_embed_children_are_present(self):
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
        metadata = FontMetadata(font)

        assert metadata.embed_regular is True
        assert metadata.embed_bold is True
        assert metadata.embed_italic is True
        assert metadata.embed_bold_italic is True

    def it_exposes_its_underlying_element(self):
        font = cast(CT_Font, element("w:font{w:name=Arial}"))
        assert FontMetadata(font).element is font
