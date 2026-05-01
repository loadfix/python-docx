"""|CustomXmlPart| — container part for a custom XML data item.

Custom XML data parts hold arbitrary XML payloads referenced from a document via
``customXml`` relationships. A pair of parts is typical::

    /customXml/item1.xml          application/xml
    /customXml/itemProps1.xml     application/vnd.openxmlformats-officedocument
                                     .customXmlProperties+xml

The data part carries the actual XML tree; the sibling properties part declares
a ``{GUID}`` store-item id and any schema references. python-docx exposes
custom XML data parts as read-only blobs.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.part import Part

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.opc.packuri import PackURI


class CustomXmlPart(Part):
    """Part containing the XML payload of a custom XML data item.

    Corresponds to a part whose content-type is ``application/xml`` (or any
    other XML media type) referenced via a ``customXml`` relationship. The
    sibling properties part (content-type
    ``...customXmlProperties+xml``) is surfaced via :attr:`properties_part`.
    """

    def __init__(self, partname: "PackURI", content_type: str, blob: bytes):
        super().__init__(partname, content_type, blob)

    @classmethod
    def load(
        cls,
        partname: "PackURI",
        content_type: str,
        blob: bytes,
        package: "OpcPackage",
    ) -> "CustomXmlPart":
        """Called by ``PartFactory`` when loading a custom XML data part."""
        return cls(partname, content_type, blob)
