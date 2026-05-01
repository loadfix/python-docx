"""Custom element classes related to the ``word/glossary/document.xml`` part.

The glossary document part carries the "building blocks" (AutoText / Quick
Parts / cover pages, etc.) that ship with a document. Each building block is
a ``w:docPart`` inside ``w:docParts``, containing a ``w:docPartPr`` metadata
block and a ``w:docPartBody`` holding the block's content paragraphs and
tables.

Only the pieces exposed through :class:`docx.glossary.Glossary` are modelled
at this pass — the building-block content is treated as a generic story
(paragraphs plus tables). Creation of new building blocks is intentionally
out of scope; the proxy layer is read-only.

``w:name`` is already registered globally as :class:`CT_String` (it's used as
a simple ``@w:val`` carrier under a handful of WML parents — style names,
footer references, etc.), so the glossary-specific classes below read the
``w:val`` attribute of their ``w:name`` / ``w:description`` / ``w:guid``
children directly via ``xpath`` rather than via a typed child accessor.
That keeps the global ``w:name`` registration intact.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


def _val_of(element: BaseOxmlElement | None) -> str | None:
    """Return the ``w:val`` attribute of `element`, or |None|.

    Returns |None| when `element` is |None| or the attribute is not present.
    Avoids taking a typed ``CT_String`` route because several of the glossary
    children share a local-name (``w:name``) with other WML elements whose
    registered class is ``CT_String`` — this helper keeps the reader tolerant
    of both.
    """
    if element is None:
        return None
    return element.get(qn("w:val"))


class CT_DocPartGallery(BaseOxmlElement):
    """``<w:gallery>`` child of ``w:category`` — the gallery slot for this block.

    The ``w:val`` attribute names the gallery (e.g. ``"quickParts"``,
    ``"coverPg"``). Modelled as a free-form string because the
    ``ST_DocPartGallery`` enumeration is broad and we only surface the value
    read-only.
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_DocPartCategory(BaseOxmlElement):
    """``<w:category>`` child of ``w:docPartPr`` — the block's classification.

    Contains a name (``w:name``) and a gallery (``w:gallery``) — both are
    optional at the XML level but Word always writes them for first-class
    building blocks. ``w:name`` is read via ``xpath`` to avoid colliding
    with the global ``w:name`` registration elsewhere in the schema.
    """

    gallery: CT_DocPartGallery | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:gallery", successors=()
    )

    @property
    def name_val(self) -> str | None:
        """The value of ``w:category/w:name/@w:val``, or |None| when absent."""
        found = self.xpath("./w:name[1]")
        return _val_of(found[0] if found else None)


class CT_DocPartPr(BaseOxmlElement):
    """``<w:docPartPr>`` element — metadata for a building block.

    Holds the block name (``w:name``), its category (``w:category``), a GUID
    (``w:guid``), a description (``w:description``), behaviors, types, and a
    ``w:style`` reference. Only the accessors the read-only proxy surfaces
    are exposed here.
    """

    category: CT_DocPartCategory | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:category",
        successors=("w:types", "w:behaviors", "w:description", "w:guid"),
    )

    @property
    def name_val(self) -> str | None:
        """Value of ``w:docPartPr/w:name/@w:val``, or |None| when absent."""
        found = self.xpath("./w:name[1]")
        return _val_of(found[0] if found else None)

    @property
    def description_val(self) -> str | None:
        """Value of ``w:docPartPr/w:description/@w:val``, or |None| when absent."""
        found = self.xpath("./w:description[1]")
        return _val_of(found[0] if found else None)

    @property
    def guid_val(self) -> str | None:
        """Value of ``w:docPartPr/w:guid/@w:val``, or |None| when absent."""
        found = self.xpath("./w:guid[1]")
        return _val_of(found[0] if found else None)


class CT_DocPartBody(BaseOxmlElement):
    """``<w:docPartBody>`` element — container for a building block's content.

    A story-like container: holds paragraphs (``w:p``) and tables (``w:tbl``)
    in document order.
    """

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """All ``w:p`` and ``w:tbl`` elements in this doc-part body, in order."""
        return self.xpath("./w:p | ./w:tbl")


class CT_DocPart(BaseOxmlElement):
    """``<w:docPart>`` element — a single building block.

    Contains the metadata block (``w:docPartPr``) and the body
    (``w:docPartBody``) holding the block's content.
    """

    _tag_seq = ("w:docPartPr", "w:docPartBody")
    docPartPr: CT_DocPartPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartPr", successors=("w:docPartBody",)
    )
    docPartBody: CT_DocPartBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartBody", successors=()
    )
    del _tag_seq


class CT_DocParts(BaseOxmlElement):
    """``<w:docParts>`` element — container for the building blocks.

    Direct child of ``w:glossaryDocument``. Holds a flat list of
    ``w:docPart`` children.
    """

    docPart = ZeroOrMore("w:docPart", successors=())

    docPart_lst: list[CT_DocPart]


class CT_GlossaryDocument(BaseOxmlElement):
    """``<w:glossaryDocument>`` element — root of the glossary document part.

    Holds a single ``w:docParts`` container. Other children defined in the
    schema are not surfaced by this proxy.
    """

    docParts: CT_DocParts | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docParts", successors=()
    )

    @property
    def docPart_lst(self) -> list[CT_DocPart]:
        """All ``w:docPart`` elements under this glossary document, in order.

        Returns an empty list when no ``w:docParts`` container is present or
        when it is empty.
        """
        docParts = self.docParts
        if docParts is None:
            return []
        return docParts.docPart_lst
