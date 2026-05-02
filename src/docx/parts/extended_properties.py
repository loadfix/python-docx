"""Extended document properties part.

The extended-properties part (``docProps/app.xml``) stores application-written
metadata (``Company``, ``Manager``, ``Application``, ``AppVersion``,
``TotalTime``, cached statistics, etc.). It is distinct from both the core
properties part (``docProps/core.xml``, Dublin-Core) and the custom properties
part (``docProps/custom.xml``, typed user-defined name/value pairs).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.extended_properties import ExtendedProperties
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.extended_properties import CT_ExtendedProperties
from docx.oxml.ns import nsmap
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.package import Package


class ExtendedPropertiesPart(XmlPart):
    """Container part for extended document properties.

    Corresponds to ``/docProps/app.xml`` in the package.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_ExtendedProperties,
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._element = element

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| proxy for the `<Properties>` root element."""
        return ExtendedProperties(self._element)

    @classmethod
    def default(cls, package: "Package") -> Self:
        """Return a newly created extended properties part with an empty root."""
        partname = PackURI("/docProps/app.xml")
        content_type = CT.OFC_EXTENDED_PROPERTIES
        element = cast("CT_ExtendedProperties", parse_xml(cls._default_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_xml(cls) -> bytes:
        """A byte-string containing the default XML for an extended properties part."""
        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Properties xmlns="{nsmap["extprops"]}" '
            f'xmlns:vt="{nsmap["vt"]}"/>'
        ).encode("utf-8")
