"""|EndnotesPart| and closely related objects."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.endnotes import Endnotes
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.endnotes import CT_Endnotes
from docx.oxml.parser import parse_xml
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.package import Package


class EndnotesPart(StoryPart):
    """Proxy for the endnotes.xml part containing endnotes for a document."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_Endnotes, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._endnotes = element

    @property
    def endnotes(self) -> Endnotes:
        """A |Endnotes| proxy object for the `w:endnotes` root element of this part."""
        return Endnotes(self._endnotes, self)

    @property
    def endnotes_element(self) -> CT_Endnotes:
        """The `w:endnotes` root element of this part."""
        return self._endnotes

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created endnotes part, containing separator and continuation separator."""
        partname = PackURI("/word/endnotes.xml")
        content_type = CT.WML_ENDNOTES
        element = cast("CT_Endnotes", parse_xml(cls._default_endnotes_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_endnotes_xml(cls) -> bytes:
        """A byte-string containing XML for a default endnotes part."""
        path = os.path.join(os.path.split(__file__)[0], "..", "templates", "default-endnotes.xml")
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes
