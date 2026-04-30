"""Unit test suite for the `docx.parts.custom_properties` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.custom_properties import CustomProperties
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.custom_properties import CT_CustomProperties
from docx.oxml.parser import parse_xml
from docx.package import Package
from docx.parts.custom_properties import CustomPropertiesPart

from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock


_EMPTY_PROPERTIES_XML = (
    b'<Properties '
    b'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
)


class DescribeCustomPropertiesPart:
    """Unit test suite for `docx.parts.custom_properties.CustomPropertiesPart`."""

    def it_provides_access_to_its_custom_properties_collection(
        self, CustomProperties_: Mock, custom_properties_: Mock, package_: Mock
    ):
        CustomProperties_.return_value = custom_properties_
        elm = cast(CT_CustomProperties, parse_xml(_EMPTY_PROPERTIES_XML))
        part = CustomPropertiesPart(
            PackURI("/docProps/custom.xml"), CT.OFC_CUSTOM_PROPERTIES, elm, package_
        )

        custom_properties = part.custom_properties

        CustomProperties_.assert_called_once_with(part.element, part)
        assert custom_properties is custom_properties_

    def it_constructs_a_default_custom_properties_part_to_help(self):
        package = Package()

        part = CustomPropertiesPart.default(package)

        assert isinstance(part, CustomPropertiesPart)
        assert part.partname == "/docProps/custom.xml"
        assert part.content_type == CT.OFC_CUSTOM_PROPERTIES
        assert part.package is package
        assert part.element.tag == (
            "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}"
            "Properties"
        )
        assert len(part.element) == 0

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def CustomProperties_(self, request: FixtureRequest) -> Mock:
        return class_mock(request, "docx.parts.custom_properties.CustomProperties")

    @pytest.fixture
    def custom_properties_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, CustomProperties)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
