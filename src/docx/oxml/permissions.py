"""Custom element classes related to permission ranges (`w:permStart`/`w:permEnd`)."""

from __future__ import annotations

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute


class CT_PermStart(BaseOxmlElement):
    """`w:permStart` element, marking the start of a rich-text permission range."""

    id: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:id", ST_DecimalNumber
    )
    edit_group: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:edGrp", ST_String
    )
    user: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:ed", ST_String
    )
    displaced_by_custom_xml: str | None = (
        OptionalAttribute(  # pyright: ignore[reportAssignmentType]
            "w:displacedByCustomXml", ST_String
        )
    )
    col_first: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:colFirst", ST_DecimalNumber
    )
    col_last: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:colLast", ST_DecimalNumber
    )


class CT_PermEnd(BaseOxmlElement):
    """`w:permEnd` element, marking the end of a rich-text permission range."""

    id: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:id", ST_DecimalNumber
    )
    displaced_by_custom_xml: str | None = (
        OptionalAttribute(  # pyright: ignore[reportAssignmentType]
            "w:displacedByCustomXml", ST_String
        )
    )
