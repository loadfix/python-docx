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
        """Return a new |CustomPropertiesPart| with no custom properties."""
        partname = PackURI("/docProps/custom.xml")
        content_type = CT.OFC_CUSTOM_PROPERTIES
        element = CT_CustomProperties.new()
        return cls(partname, content_type, element, package)

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| object providing read/write access to the custom
        properties contained in this part."""
        return CustomProperties(cast(CT_CustomProperties, self.element))
