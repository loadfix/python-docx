"""Custom properties part, corresponds to ``/docProps/custom.xml`` part in package."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.custom_properties import CustomProperties
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.custom_properties import CT_CustomProperties

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named ``/docProps/custom.xml``.

    Contains custom document properties as name/value pairs.
    """

    @classmethod
    def default(cls, package: OpcPackage) -> CustomPropertiesPart:
        return cls._new(package)

    @property
    def custom_properties(self) -> CustomProperties:
        return CustomProperties(cast(CT_CustomProperties, self.element))

    @classmethod
    def _new(cls, package: OpcPackage) -> CustomPropertiesPart:
        partname = PackURI("/docProps/custom.xml")
        content_type = CT.OFC_CUSTOM_PROPERTIES
        element = CT_CustomProperties.new()
        return CustomPropertiesPart(partname, content_type, element, package)
