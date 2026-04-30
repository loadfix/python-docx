"""Custom document properties part.

The custom properties part (``docProps/custom.xml``) stores user-defined document
metadata as typed name/value pairs. It is distinct from the core properties part
(``docProps/core.xml``) which stores a fixed set of Dublin-Core metadata fields.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.custom_properties import CustomProperties
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.custom_properties import CT_CustomProperties
from docx.oxml.ns import nsmap
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.package import Package


class CustomPropertiesPart(XmlPart):
    """Container part for custom document properties.

    Corresponds to ``/docProps/custom.xml`` in the package.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_CustomProperties,
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._custom_properties_elm = element

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| proxy for the `<Properties>` root element of this part."""
        return CustomProperties(self._custom_properties_elm, self)

    @classmethod
    def default(cls, package: "Package") -> Self:
        """Return a newly created custom properties part with an empty `<Properties>` root."""
        partname = PackURI("/docProps/custom.xml")
        content_type = CT.OFC_CUSTOM_PROPERTIES
        element = cast("CT_CustomProperties", parse_xml(cls._default_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_xml(cls) -> bytes:
        """A byte-string containing the default XML for a custom properties part."""
        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Properties xmlns="{nsmap["custprops"]}" '
            f'xmlns:vt="{nsmap["vt"]}"/>'
        ).encode("utf-8")
