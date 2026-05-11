"""|BibliographyPart| — container part holding ``b:Sources`` citation data.

The bibliography is stored in the package at ``/customXml/item{N}.xml`` with
an ``application/xml`` content type. A sibling ``/customXml/itemProps{N}.xml``
properties part declares the well-known bibliography ``schemaRef``
(``http://schemas.openxmlformats.org/officeDocument/2006/bibliography``) and
carries a unique ``{GUID}`` store-item id that citation SDTs refer to via
``<w:dataBinding/@w:storeItemID>``.

This part subclasses |CustomXmlPart| so it lives in the same rels slot as any
other custom-XML payload; the distinguishing feature is the ``b:Sources``
root element.
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.bibliography import CT_Sources, new_sources_root
from docx.oxml.ns import nsmap
from docx.oxml.parser import parse_xml
from docx.parts.custom_xml import CustomXmlPart

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage


# -- URI stored in the sibling properties part to identify this datastore
# -- item as bibliography data. --
BIBLIOGRAPHY_SCHEMA_URI = nsmap["b"]


class BibliographyPart(CustomXmlPart):
    """Custom-XML data part whose root element is ``<b:Sources>``.

    Structurally identical to a plain |CustomXmlPart| — the difference is
    that python-docx keeps the parsed ``<b:Sources>`` lxml element in
    ``self._sources`` so callers can mutate it in place and then call
    :meth:`flush` (or rely on the automatic flush at save time) to serialize
    it back to ``self._blob``.
    """

    def __init__(self, partname: "PackURI", content_type: str, blob: bytes):
        super().__init__(partname, content_type, blob)
        # -- lazy-parsed CT_Sources; materialized on first access --
        self._sources_elm: "CT_Sources | None" = None

    @property
    def sources_element(self) -> "CT_Sources":
        """The ``<b:Sources>`` root element of this part (parsed lazily)."""
        if self._sources_elm is None:
            self._sources_elm = cast(CT_Sources, parse_xml(self.blob))
        return self._sources_elm

    def flush(self) -> None:
        """Serialize the in-memory ``<b:Sources>`` element back to the blob."""
        if self._sources_elm is None:
            return
        from lxml import etree

        self._blob = etree.tostring(
            self._sources_elm, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @property
    def blob(self) -> bytes:
        """Raw bytes of the data part (flushes pending in-memory mutations first)."""
        self.flush()
        return self._blob

    @blob.setter
    def blob(self, value: bytes) -> None:
        """Reset the raw bytes and clear any cached parsed element."""
        self._blob = value
        self._sources_elm = None

    @classmethod
    def default(cls, package: "OpcPackage") -> Self:
        """Return a newly created bibliography part with an empty ``<b:Sources>`` root.

        Assigns the first free ``/customXml/item{N}.xml`` partname in the
        package and produces a sibling ``itemProps{N}.xml`` (created by
        :meth:`attach_itemProps`) once the part is related to a document.
        """
        partname_str = cls._next_item_partname(package)
        partname = PackURI(partname_str)
        content_type = CT.XML
        sources = new_sources_root()
        from lxml import etree

        blob = etree.tostring(
            sources, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        part = cls(partname, content_type, blob)
        part._sources_elm = sources
        return part

    @staticmethod
    def _next_item_partname(package: "OpcPackage") -> str:
        """Return ``/customXml/itemN.xml`` with the first free ``N``."""
        used_numbers: set[int] = set()
        for part in package.iter_parts():
            partname = str(part.partname)
            if not partname.startswith("/customXml/item"):
                continue
            # -- split off the numeric suffix before `.xml` --
            tail = partname[len("/customXml/item") :]
            if not tail.endswith(".xml"):
                continue
            stem = tail[: -len(".xml")]
            # -- distinguish `item1.xml` from `itemProps1.xml` --
            if stem.startswith("Props"):
                continue
            try:
                used_numbers.add(int(stem))
            except ValueError:
                continue
        n = 1
        while n in used_numbers:
            n += 1
        return f"/customXml/item{n}.xml"

    def attach_itemProps(self, package: "OpcPackage") -> "ItemPropsPart":
        """Create (and relate) a sibling ``itemProps{N}.xml`` datastore part.

        Idempotent — if a ``customXmlProps`` relationship already exists on
        this part the existing props part is returned unchanged.
        """
        # -- bail out when we already have one --
        try:
            existing = self.part_related_by(RT.CUSTOM_XML_PROPS)
            return cast("ItemPropsPart", existing)
        except KeyError:
            pass

        # -- derive the sibling partname: /customXml/item{N}.xml -> itemProps{N}.xml --
        partname = str(self.partname)
        stem = partname[len("/customXml/item") : -len(".xml")]
        props_partname = PackURI(f"/customXml/itemProps{stem}.xml")
        props = ItemPropsPart.default(props_partname)
        self.relate_to(props, RT.CUSTOM_XML_PROPS)
        return props

    @property
    def store_item_id(self) -> "str | None":
        """The ``{GUID}`` store-item id from the sibling itemProps part, or |None|."""
        try:
            props = self.part_related_by(RT.CUSTOM_XML_PROPS)
        except KeyError:
            return None
        return getattr(props, "store_item_id", None)


class ItemPropsPart(Part):
    """Sibling properties part for a |BibliographyPart| (``itemProps{N}.xml``).

    Carries a ``<ds:datastoreItem>`` root that declares the ``{GUID}``
    store-item id and the ``ds:schemaRef`` for the bibliography namespace.
    """

    @classmethod
    def default(cls, partname: "PackURI") -> "ItemPropsPart":
        """Build a fresh ``itemProps{N}.xml`` part with a unique store-item id."""
        store_item_id = "{" + str(uuid.uuid4()).upper() + "}"
        blob = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<ds:datastoreItem '
            f'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" '
            f'ds:itemID="{store_item_id}">\n'
            f'  <ds:schemaRefs>\n'
            f'    <ds:schemaRef ds:uri="{BIBLIOGRAPHY_SCHEMA_URI}"/>\n'
            f'  </ds:schemaRefs>\n'
            f'</ds:datastoreItem>'
        ).encode("utf-8")
        return cls(partname, CT.OFC_CUSTOM_XML_PROPERTIES, blob)

    @property
    def store_item_id(self) -> "str | None":
        """The ``{GUID}`` value from ``ds:datastoreItem/@ds:itemID``, or |None|."""
        from lxml import etree

        # -- hardened parser (resolve_entities=False, no_network=True) guards
        # -- against XXE / SSRF via attacker-supplied bibliography props. --
        try:
            root = parse_xml(self.blob)
        except etree.XMLSyntaxError:
            return None
        ds_ns = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
        return root.get(f"{{{ds_ns}}}itemID")

    @classmethod
    def load(cls, partname, content_type, blob, package):
        return cls(partname, content_type, blob)
