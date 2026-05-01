"""Unit test suite for the `docx.parts.web_settings` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.web_settings import CT_WebSettings
from docx.package import Package
from docx.parts.web_settings import WebSettingsPart
from docx.web_settings import WebSettings

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeWebSettingsPart:
    """Unit test suite for `docx.parts.web_settings.WebSettingsPart` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_web_settings_part(
        self,
        package_: Mock,
        WebSettingsPart_load_: Mock,
        web_settings_part_: Mock,
    ):
        partname = PackURI("/word/webSettings.xml")
        content_type = CT.WML_WEB_SETTINGS
        reltype = RT.WEB_SETTINGS
        blob = (
            b"<w:webSettings xmlns:w="
            b"'http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"
        )
        WebSettingsPart_load_.return_value = web_settings_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        WebSettingsPart_load_.assert_called_once_with(
            partname, content_type, blob, package_
        )
        assert part is web_settings_part_

    def it_provides_access_to_its_web_settings_proxy(
        self, WebSettings_: Mock, web_settings_: Mock, package_: Mock
    ):
        WebSettings_.return_value = web_settings_
        ws_elm = cast(CT_WebSettings, element("w:webSettings"))
        web_settings_part = WebSettingsPart(
            PackURI("/word/webSettings.xml"), CT.WML_WEB_SETTINGS, ws_elm, package_
        )

        web_settings = web_settings_part.web_settings

        WebSettings_.assert_called_once_with(ws_elm, web_settings_part)
        assert web_settings is web_settings_

    def it_exposes_its_web_settings_element(self, package_: Mock):
        ws_elm = cast(CT_WebSettings, element("w:webSettings"))
        web_settings_part = WebSettingsPart(
            PackURI("/word/webSettings.xml"), CT.WML_WEB_SETTINGS, ws_elm, package_
        )

        assert web_settings_part.web_settings_element is ws_elm

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def WebSettings_(self, request: FixtureRequest) -> Mock:
        return class_mock(request, "docx.parts.web_settings.WebSettings")

    @pytest.fixture
    def web_settings_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, WebSettings)

    @pytest.fixture
    def web_settings_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, WebSettingsPart)

    @pytest.fixture
    def WebSettingsPart_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, WebSettingsPart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
