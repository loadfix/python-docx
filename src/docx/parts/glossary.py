"""|GlossaryPart| providing access to the ``word/glossary/document.xml`` part.

The glossary-document part carries the AutoText / Quick Parts / cover-page
building blocks that ship with a document. It is Word-authored and
python-docx does not create one on demand; the proxy exposed via
:attr:`docx.document.Document.glossary` is intentionally read-only at this
pass.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.glossary import Glossary
from docx.opc.part import XmlPart

if TYPE_CHECKING:
    from docx.opc.packuri import PackURI
    from docx.oxml.glossary import CT_GlossaryDocument
    from docx.package import Package


class GlossaryPart(XmlPart):
    """Read-only proxy for the ``word/glossary/document.xml`` part.

    Creation of a default (empty) glossary part is out of scope for this
    pass: :attr:`docx.document.Document.glossary` returns |None| for
    documents that do not already relate a glossary document part.
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
