"""Custom element classes for VML watermark-related elements.

Watermarks in Word documents are represented using legacy VML shapes
inside a `w:pict` element. These element classes provide minimal access
to the relevant VML shapes used for text and image watermarks.
"""

from __future__ import annotations

from docx.oxml.simpletypes import ST_RelationshipId, XsdString
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


class CT_Pict(BaseOxmlElement):
    """``<w:pict>`` element, container for VML shapes including watermarks."""

    shape = ZeroOrOne("v:shape", successors=())


class CT_VmlShape(BaseOxmlElement):
    """``<v:shape>`` element, a VML shape used for watermarks."""

    fill = ZeroOrOne("v:fill", successors=("v:stroke", "v:imagedata", "v:textpath"))
    imagedata = ZeroOrOne("v:imagedata", successors=("v:textpath",))
    textpath = ZeroOrOne("v:textpath", successors=())

    id: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdString
    )
    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", XsdString
    )
    style: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "style", XsdString
    )


class CT_VmlFill(BaseOxmlElement):
    """``<v:fill>`` element inside a VML shape."""

    color: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "color", XsdString
    )


class CT_VmlImageData(BaseOxmlElement):
    """``<v:imagedata>`` element referencing an image relationship."""

    rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:id", ST_RelationshipId
    )
    title: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "o:title", XsdString
    )


class CT_VmlTextpath(BaseOxmlElement):
    """``<v:textpath>`` element containing the text string of a text watermark."""

    style: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "style", XsdString
    )
    string: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "string", XsdString
    )
