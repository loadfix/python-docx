"""Unit test suite for the `docx.parts.theme` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.theme import CT_Theme
from docx.package import Package
from docx.parts.theme import ThemePart
from docx.theme import Theme

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeThemePart:
    """Unit test suite for `docx.parts.theme.ThemePart` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_theme_part(
        self,
        package_: Mock,
        ThemePart_load_: Mock,
        theme_part_: Mock,
    ):
        partname = PackURI("/word/theme/theme1.xml")
        content_type = CT.OFC_THEME
        reltype = RT.THEME
        blob = (
            b"<a:theme xmlns:a="
            b"'http://schemas.openxmlformats.org/drawingml/2006/main'/>"
        )
        ThemePart_load_.return_value = theme_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        ThemePart_load_.assert_called_once_with(
            partname, content_type, blob, package_
        )
        assert part is theme_part_

    def it_provides_access_to_its_theme_proxy(
        self, Theme_: Mock, theme_: Mock, package_: Mock
    ):
        Theme_.return_value = theme_
        theme_elm = cast(CT_Theme, element("a:theme"))
        theme_part = ThemePart(
            PackURI("/word/theme/theme1.xml"), CT.OFC_THEME, theme_elm, package_
        )

        theme = theme_part.theme

        Theme_.assert_called_once_with(theme_elm, theme_part)
        assert theme is theme_

    def it_exposes_its_theme_element(self, package_: Mock):
        theme_elm = cast(CT_Theme, element("a:theme"))
        theme_part = ThemePart(
            PackURI("/word/theme/theme1.xml"), CT.OFC_THEME, theme_elm, package_
        )

        assert theme_part.theme_element is theme_elm

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def Theme_(self, request: FixtureRequest) -> Mock:
        return class_mock(request, "docx.parts.theme.Theme")

    @pytest.fixture
    def theme_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Theme)

    @pytest.fixture
    def theme_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, ThemePart)

    @pytest.fixture
    def ThemePart_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, ThemePart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
