"""|FootnotesPart| and closely related objects."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.footnotes import Footnotes
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.footnotes import CT_Footnotes
from docx.oxml.parser import parse_xml
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.package import Package


class FootnotesPart(StoryPart):
    """Proxy for the footnotes.xml part containing footnotes for a document."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_Footnotes, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._footnotes = element

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| proxy object for the `w:footnotes` root element of this part."""
        return Footnotes(self._footnotes, self)

    @property
    def footnotes_element(self) -> CT_Footnotes:
        """The `w:footnotes` root element of this part."""
        return self._footnotes

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created footnotes part, containing separator and continuation separator."""
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = cast("CT_Footnotes", parse_xml(cls._default_footnotes_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_footnotes_xml(cls) -> bytes:
        """A byte-string containing XML for a default footnotes part."""
        path = os.path.join(os.path.split(__file__)[0], "..", "templates", "default-footnotes.xml")
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes
