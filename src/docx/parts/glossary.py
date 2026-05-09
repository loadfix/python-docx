"""|GlossaryPart| providing access to the ``word/glossary/document.xml`` part.

The glossary-document part carries the AutoText / Quick Parts / cover-page
building blocks that ship with a document. As of 2026.05.10 python-docx
can also create a fresh, empty glossary part on demand — see
:meth:`GlossaryPart.default` and :attr:`docx.document.Document.glossary`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.glossary import Glossary
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.oxml.glossary import CT_GlossaryDocument
    from docx.package import Package


_DEFAULT_GLOSSARY_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:glossaryDocument '
    b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b'<w:docParts/>'
    b'</w:glossaryDocument>\n'
)


class GlossaryPart(XmlPart):
    """Proxy for the ``word/glossary/document.xml`` part.

    :attr:`docx.document.Document.glossary` returns a |Glossary| proxy for
    this part; when the document has no ``glossaryDocument`` relationship,
    the document lazily creates a fresh, empty part on first access via
    :meth:`default`.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_GlossaryDocument,
        package: Package,
    ):
        super().__init__(partname, content_type, element, package)
        self._glossary_elm = element

    @property
    def glossary(self) -> Glossary:
        """A |Glossary| proxy for the ``w:glossaryDocument`` root of this part."""
        return Glossary(self._glossary_elm, self)

    @property
    def glossary_element(self) -> CT_GlossaryDocument:
        """The ``w:glossaryDocument`` root element for this part."""
        return cast("CT_GlossaryDocument", self._element)

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created, empty glossary document part.

        Used by :attr:`docx.document.Document.glossary` when the document
        has no existing ``glossaryDocument`` relationship and the caller
        requests write access. The part starts with an empty
        ``w:docParts`` container so new building blocks can be appended
        without further bookkeeping.

        .. versionadded:: 2026.05.10
        """
        partname = PackURI("/word/glossary/document.xml")
        content_type = CT.WML_DOCUMENT_GLOSSARY
        element = cast("CT_GlossaryDocument", parse_xml(_DEFAULT_GLOSSARY_XML))
        return cls(partname, content_type, element, package)
