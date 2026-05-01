"""Unit test suite for the `docx.parts.font_table` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.font_table import FontTable
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.font_table import CT_Fonts
from docx.package import Package
from docx.parts.font_table import FontTablePart

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeFontTablePart:
    """Unit test suite for `docx.parts.font_table.FontTablePart` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_font_table_part(
        self,
        package_: Mock,
        FontTablePart_load_: Mock,
        font_table_part_: Mock,
    ):
        partname = PackURI("/word/fontTable.xml")
        content_type = CT.WML_FONT_TABLE
        reltype = RT.FONT_TABLE
        blob = b"<w:fonts xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"
        FontTablePart_load_.return_value = font_table_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        FontTablePart_load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is font_table_part_

    def it_provides_access_to_its_font_table_collection(
        self, FontTable_: Mock, font_table_: Mock, package_: Mock
    ):
        FontTable_.return_value = font_table_
        fonts_elm = cast(CT_Fonts, element("w:fonts"))
        font_table_part = FontTablePart(
            PackURI("/word/fontTable.xml"), CT.WML_FONT_TABLE, fonts_elm, package_
        )

        font_table = font_table_part.font_table

        FontTable_.assert_called_once_with(fonts_elm, font_table_part)
        assert font_table is font_table_

    def it_exposes_its_font_table_element(self, package_: Mock):
        fonts_elm = cast(CT_Fonts, element("w:fonts"))
        font_table_part = FontTablePart(
            PackURI("/word/fontTable.xml"), CT.WML_FONT_TABLE, fonts_elm, package_
        )

        assert font_table_part.font_table_element is fonts_elm

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def FontTable_(self, request: FixtureRequest) -> Mock:
        return class_mock(request, "docx.parts.font_table.FontTable")

    @pytest.fixture
    def font_table_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, FontTable)

    @pytest.fixture
    def font_table_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, FontTablePart)

    @pytest.fixture
    def FontTablePart_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, FontTablePart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
