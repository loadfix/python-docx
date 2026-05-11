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
from ooxml_customxml import CT_DatastoreItem
from ooxml_customxml import DS_NS as _DS_NS
from ooxml_customxml.oxml import parse_xml as _parse_datastore_xml

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.opc.part import Part
    from docx.parts.custom_xml import CustomXmlPart as _CustomXmlDataPart


class CustomXmlPart:
    """Read-only proxy for a custom XML data part plus its sibling properties part.

    Instances are produced by :attr:`docx.document.Document.custom_xml_parts` and
    describe one ``/customXml/item{N}.xml`` data part. :attr:`item_id` and
    :attr:`schema_refs` come from the related ``itemProps{N}.xml`` part when
    present; :attr:`root_element` and :attr:`blob` come from the data part.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, data_part: "_CustomXmlDataPart"):
        self._data_part = data_part

    @property
    def part(self) -> "_CustomXmlDataPart":
        """The underlying custom XML data |Part|.

        .. versionadded:: 2026.05.0
        """
        return self._data_part

    @property
    def partname(self) -> str:
        """The pack URI of the data part, e.g. ``"/customXml/item1.xml"``.

        .. versionadded:: 2026.05.0
        """
        return str(self._data_part.partname)

    @property
    def blob(self) -> bytes:
        """Raw bytes of the data part.

        .. versionadded:: 2026.05.0
        """
        return self._data_part.blob

    @property
    def root_element(self):
        """Parsed lxml root element of the data part, or |None| on parse failure.

        The data part is parsed lazily. Returns |None| when the payload is not
        well-formed XML — python-docx does not raise for broken custom-XML
        parts.

        .. versionadded:: 2026.05.0
        """
        try:
            # -- use hardened parser (resolve_entities=False, no_network=True) to
            # -- prevent XXE / SSRF via attacker-controlled customXml data parts.
            return parse_xml(self.blob)
        except etree.XMLSyntaxError:
            return None

    @property
    def item_id(self) -> str | None:
        """The ``{GUID}`` store-item id from the sibling ``itemProps`` part, or |None|.

        Looked up via the ``customXmlProps`` relationship and returned verbatim
        from the ``ds:datastoreItem/@ds:itemID`` attribute. Returns |None| when
        the data part has no properties part, the properties part cannot be
        parsed, or the attribute is missing.

        .. versionadded:: 2026.05.0
        """
        item = self._datastore_item()
        if item is None:
            return None
        # -- read via the low-level lxml `.get()` rather than CT_DatastoreItem.itemID
        # -- so a malformed part missing the required attribute returns None instead
        # -- of raising InvalidXmlError. Matches the pre-adoption tolerance. --
        return item.get(f"{{{_DS_NS}}}itemID")

    @property
    def schema_refs(self) -> list[str]:
        """List of schema URIs declared in the sibling ``itemProps`` part.

        Each entry is the value of a ``ds:schemaRef/@ds:uri`` attribute under
        ``ds:schemaRefs``. Returns an empty list when there is no properties
        part, when it cannot be parsed, or when no schema references are
        declared.

        .. versionadded:: 2026.05.0
        """
        item = self._datastore_item()
        if item is None:
            return []
        return item.schema_refs

    # -- internal helpers ----------------------------------------------------

    def _itemProps_part(self):
        """Return the sibling custom XML properties part, or |None|."""
        try:
            return self._data_part.part_related_by(RT.CUSTOM_XML_PROPS)
        except (KeyError, ValueError):
            return None

    def _datastore_item(self) -> "CT_DatastoreItem | None":
        """Parsed :class:`CT_DatastoreItem` root of the properties part, or |None|.

        Delegates to the hardened :mod:`ooxml_customxml` parser so the
        returned element has the full descriptor-backed surface (the
        ``schemaRefs`` child proxy, the ``schema_refs`` URI list, etc.)
        rather than a bare lxml element. Returns |None| when there is
        no sibling properties part, when its blob is empty, when
        parsing fails, or when the root is not a ``<ds:datastoreItem>``.
        """
        props_part = self._itemProps_part()
        if props_part is None:
            return None
        blob = getattr(props_part, "blob", None)
        if not blob:
            return None
        try:
            root = _parse_datastore_xml(blob)
        except etree.XMLSyntaxError:
            return None
        if not isinstance(root, CT_DatastoreItem):
            return None
        return root


def iter_custom_xml_parts(document_part: "Part") -> list[CustomXmlPart]:
    """Return a |CustomXmlPart| for each ``customXml`` relationship on ``document_part``.

    Custom XML data parts are located via the ``customXml`` relationship type on
    the main document part. Properties parts (``customXmlProps``) are not
    surfaced as separate entries — they are accessed via
    :attr:`CustomXmlPart.item_id` and :attr:`CustomXmlPart.schema_refs`.

    .. versionadded:: 2026.05.0
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
