"""Custom XML data parts and their proxy wrappers.

Custom XML data parts hold arbitrary XML payloads that Word uses to back data-bound
content controls (see :class:`docx.content_controls.DataBinding`). A pair of package
parts is typical:

* ``/customXml/item{N}.xml`` — the XML payload (content type ``application/xml``)
* ``/customXml/itemProps{N}.xml`` — a ``ds:datastoreItem`` declaring a ``{GUID}``
  store-item id and any referenced schemas

The :class:`CustomXmlPart` proxy exposes read-only metadata: the store-item id,
the schema-ref URIs, the parsed XML root element, and the raw blob.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from lxml import etree

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT

if TYPE_CHECKING:
    from docx.opc.part import Part
    from docx.parts.custom_xml import CustomXmlPart as _CustomXmlDataPart


# -- namespace for custom-XML datastore props --
_DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"


class CustomXmlPart:
    """Read-only proxy for a custom XML data part plus its sibling properties part.

    Instances are produced by :attr:`docx.document.Document.custom_xml_parts` and
    describe one ``/customXml/item{N}.xml`` data part. :attr:`item_id` and
    :attr:`schema_refs` come from the related ``itemProps{N}.xml`` part when
    present; :attr:`root_element` and :attr:`blob` come from the data part.
    """

    def __init__(self, data_part: "_CustomXmlDataPart"):
        self._data_part = data_part

    @property
    def part(self) -> "_CustomXmlDataPart":
        """The underlying custom XML data |Part|."""
        return self._data_part

    @property
    def partname(self) -> str:
        """The pack URI of the data part, e.g. ``"/customXml/item1.xml"``."""
        return str(self._data_part.partname)

    @property
    def blob(self) -> bytes:
        """Raw bytes of the data part."""
        return self._data_part.blob

    @property
    def root_element(self):
        """Parsed lxml root element of the data part, or |None| on parse failure.

        The data part is parsed lazily. Returns |None| when the payload is not
        well-formed XML — python-docx does not raise for broken custom-XML
        parts.
        """
        try:
            return etree.fromstring(self.blob)
        except etree.XMLSyntaxError:
            return None

    @property
    def item_id(self) -> str | None:
        """The ``{GUID}`` store-item id from the sibling ``itemProps`` part, or |None|.

        Looked up via the ``customXmlProps`` relationship and returned verbatim
        from the ``ds:datastoreItem/@ds:itemID`` attribute. Returns |None| when
        the data part has no properties part, the properties part cannot be
        parsed, or the attribute is missing.
        """
        props_elm = self._itemProps_root()
        if props_elm is None:
            return None
        return props_elm.get(f"{{{_DS_NS}}}itemID")

    @property
    def schema_refs(self) -> list[str]:
        """List of schema URIs declared in the sibling ``itemProps`` part.

        Each entry is the value of a ``ds:schemaRef/@ds:uri`` attribute under
        ``ds:schemaRefs``. Returns an empty list when there is no properties
        part, when it cannot be parsed, or when no schema references are
        declared.
        """
        props_elm = self._itemProps_root()
        if props_elm is None:
            return []
        refs: list[str] = []
        for ref in props_elm.iter(f"{{{_DS_NS}}}schemaRef"):
            uri = ref.get(f"{{{_DS_NS}}}uri")
            if uri is not None:
                refs.append(uri)
        return refs

    # -- internal helpers ----------------------------------------------------

    def _itemProps_part(self):
        """Return the sibling custom XML properties part, or |None|."""
        try:
            return self._data_part.part_related_by(RT.CUSTOM_XML_PROPS)
        except (KeyError, ValueError):
            return None

    def _itemProps_root(self):
        """Parsed root of the sibling properties part, or |None|."""
        props_part = self._itemProps_part()
        if props_part is None:
            return None
        blob = getattr(props_part, "blob", None)
        if not blob:
            return None
        try:
            return etree.fromstring(blob)
        except etree.XMLSyntaxError:
            return None


def iter_custom_xml_parts(document_part: "Part") -> list[CustomXmlPart]:
    """Return a |CustomXmlPart| for each ``customXml`` relationship on ``document_part``.

    Custom XML data parts are located via the ``customXml`` relationship type on
    the main document part. Properties parts (``customXmlProps``) are not
    surfaced as separate entries — they are accessed via
    :attr:`CustomXmlPart.item_id` and :attr:`CustomXmlPart.schema_refs`.
    """
    from docx.parts.custom_xml import CustomXmlPart as _DataPart

    result: list[CustomXmlPart] = []
    seen: set[str] = set()
    for rel in document_part.rels.values():
        if rel.is_external:
            continue
        if rel.reltype != RT.CUSTOM_XML:
            continue
        try:
            target = rel.target_part
        except ValueError:
            continue
        partname = str(target.partname)
        if partname in seen:
            continue
        seen.add(partname)
        # -- only surface parts whose content type matches the XML convention --
        ct = getattr(target, "content_type", None)
        if ct is not None and ct != CT.XML:
            # -- non-xml payloads (rare) are still wrapped; skip only obvious
            #    mis-registrations of itemProps parts --
            if ct == CT.OFC_CUSTOM_XML_PROPERTIES:
                continue
        # -- the part is expected to be a `CustomXmlPart` (per the
        #    content-type registration in `docx/__init__.py`) but we cast
        #    so the type checker is happy when a custom PartFactory has
        #    produced a different Part subclass.  Duck-typing (`.blob` +
        #    `.partname`) is sufficient at runtime. --
        result.append(CustomXmlPart(cast("_DataPart", target)))
    return result
