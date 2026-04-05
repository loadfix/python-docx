# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.opc.parts.custom_properties` module."""

from __future__ import annotations

from docx.custom_properties import CustomProperties
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.package import Package
from docx.opc.parts.custom_properties import CustomPropertiesPart

from ..unitutil.mock import FixtureRequest, instance_mock


class DescribeCustomPropertiesPart:
    """Unit-test suite for `docx.opc.parts.custom_properties.CustomPropertiesPart`."""

    def it_can_create_a_default_part(self, request: FixtureRequest):
        package_ = instance_mock(request, Package)

        part = CustomPropertiesPart.default(package_)

        assert isinstance(part, CustomPropertiesPart)
        assert part.partname == PackURI("/docProps/custom.xml")
        assert part.content_type == CT.OFC_CUSTOM_PROPERTIES

    def it_provides_access_to_custom_properties(self, request: FixtureRequest):
        package_ = instance_mock(request, Package)
        part = CustomPropertiesPart.default(package_)

        custom_properties = part.custom_properties

        assert isinstance(custom_properties, CustomProperties)
