# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.custom_xml` module and related wiring."""

from __future__ import annotations

from typing import cast
from unittest.mock import MagicMock

from docx.custom_xml import CustomXmlPart, iter_custom_xml_parts
from docx.document import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.document import CT_Document
from docx.parts.custom_xml import CustomXmlPart as CustomXmlDataPart

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, Mock, instance_mock


# ---------------------------------------------------------------------------
# helpers

_NS_DECL = (
    b'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"'
)


def _make_data_part(blob: bytes = b"<data/>") -> CustomXmlDataPart:
    return CustomXmlDataPart(PackURI("/customXml/item1.xml"), CT.XML, blob)


def _make_props_part_blob(item_id: str, schema_uris: list[str] | None = None) -> bytes:
    schemas = b""
    if schema_uris:
        inner = b"".join(
            b'<ds:schemaRef ds:uri="%s"/>' % uri.encode("utf-8") for uri in schema_uris
        )
        schemas = b"<ds:schemaRefs>%s</ds:schemaRefs>" % inner
    return (
        b'<ds:datastoreItem ' + _NS_DECL
        + b' ds:itemID="' + item_id.encode("utf-8") + b'">'
        + schemas
        + b"</ds:datastoreItem>"
    )


class _FakePropsPart:
    """Stand-in for the itemProps part — only needs `.blob`."""

    def __init__(self, blob: bytes):
        self.blob = blob


def _attach_props(
    data_part: CustomXmlDataPart, props_blob: bytes | None
) -> None:
    """Monkey-patch `data_part.part_related_by` to resolve the props part."""
    if props_blob is None:

        def raise_keyerror(reltype: str):
            raise KeyError(reltype)

        data_part.part_related_by = raise_keyerror  # type: ignore[method-assign]
        return

    props = _FakePropsPart(props_blob)

    def resolve(reltype: str):
        if reltype == RT.CUSTOM_XML_PROPS:
            return props
        raise KeyError(reltype)

    data_part.part_related_by = resolve  # type: ignore[method-assign]


# ---------------------------------------------------------------------------


class DescribeCustomXmlPart:
    """Unit-test suite for `docx.custom_xml.CustomXmlPart`."""

    def it_exposes_the_data_part_blob(self):
        data_part = _make_data_part(b"<root><x/></root>")
        _attach_props(data_part, None)

        proxy = CustomXmlPart(data_part)

        assert proxy.blob == b"<root><x/></root>"
        assert proxy.partname == "/customXml/item1.xml"
        assert proxy.part is data_part

    def it_parses_the_root_element_of_the_data_part(self):
        data_part = _make_data_part(b"<root><child/></root>")
        _attach_props(data_part, None)

        proxy = CustomXmlPart(data_part)

        root = proxy.root_element
        assert root is not None
        assert root.tag == "root"
        assert len(list(root)) == 1

    def it_returns_None_for_root_element_when_blob_is_malformed(self):
        data_part = _make_data_part(b"<not-xml")
        _attach_props(data_part, None)

        proxy = CustomXmlPart(data_part)

        assert proxy.root_element is None

    def it_reads_item_id_from_sibling_props_part(self):
        data_part = _make_data_part()
        _attach_props(
            data_part, _make_props_part_blob("{12345678-1234-1234-1234-1234567890AB}")
        )

        proxy = CustomXmlPart(data_part)

        assert proxy.item_id == "{12345678-1234-1234-1234-1234567890AB}"

    def it_returns_None_for_item_id_when_no_props_part_is_related(self):
        data_part = _make_data_part()
        _attach_props(data_part, None)

        proxy = CustomXmlPart(data_part)

        assert proxy.item_id is None

    def it_returns_None_for_item_id_when_props_blob_is_empty(self):
        data_part = _make_data_part()
        _attach_props(data_part, b"")

        proxy = CustomXmlPart(data_part)

        assert proxy.item_id is None

    def it_returns_None_for_item_id_when_props_blob_is_malformed(self):
        data_part = _make_data_part()
        _attach_props(data_part, b"<not-xml")

        proxy = CustomXmlPart(data_part)

        assert proxy.item_id is None

    def it_reads_schema_refs_from_sibling_props_part(self):
        data_part = _make_data_part()
        _attach_props(
            data_part,
            _make_props_part_blob(
                "{GUID}",
                schema_uris=["http://example.com/a", "http://example.com/b"],
            ),
        )

        proxy = CustomXmlPart(data_part)

        assert proxy.schema_refs == [
            "http://example.com/a",
            "http://example.com/b",
        ]

    def it_returns_empty_list_when_no_schemas_declared(self):
        data_part = _make_data_part()
        _attach_props(data_part, _make_props_part_blob("{GUID}"))

        proxy = CustomXmlPart(data_part)

        assert proxy.schema_refs == []

    def it_returns_empty_list_for_schema_refs_when_no_props_part_is_related(self):
        data_part = _make_data_part()
        _attach_props(data_part, None)

        proxy = CustomXmlPart(data_part)

        assert proxy.schema_refs == []


class DescribeIterCustomXmlParts:
    """Unit-test suite for `docx.custom_xml.iter_custom_xml_parts`."""

    def it_returns_an_empty_list_when_no_customXml_rels_are_present(self):
        document_part = MagicMock()
        document_part.rels.values.return_value = []

        assert iter_custom_xml_parts(document_part) == []

    def it_ignores_rels_of_unrelated_types(self):
        document_part = MagicMock()
        other_rel = MagicMock(is_external=False, reltype=RT.COMMENTS)
        document_part.rels.values.return_value = [other_rel]

        assert iter_custom_xml_parts(document_part) == []

    def it_ignores_external_customXml_rels(self):
        document_part = MagicMock()
        external_rel = MagicMock(is_external=True, reltype=RT.CUSTOM_XML)
        document_part.rels.values.return_value = [external_rel]

        assert iter_custom_xml_parts(document_part) == []

    def it_wraps_each_customXml_data_part(self):
        document_part = MagicMock()
        data_part = _make_data_part(b"<a/>")
        _attach_props(data_part, None)
        rel = MagicMock(is_external=False, reltype=RT.CUSTOM_XML)
        rel.target_part = data_part
        document_part.rels.values.return_value = [rel]

        result = iter_custom_xml_parts(document_part)

        assert len(result) == 1
        assert isinstance(result[0], CustomXmlPart)
        assert result[0].blob == b"<a/>"

    def it_deduplicates_parts_referenced_multiple_times(self):
        document_part = MagicMock()
        data_part = _make_data_part(b"<a/>")
        _attach_props(data_part, None)
        rel1 = MagicMock(is_external=False, reltype=RT.CUSTOM_XML)
        rel1.target_part = data_part
        rel2 = MagicMock(is_external=False, reltype=RT.CUSTOM_XML)
        rel2.target_part = data_part
        document_part.rels.values.return_value = [rel1, rel2]

        result = iter_custom_xml_parts(document_part)

        assert len(result) == 1


class DescribeDocument_custom_xml_parts:
    """Unit-test suite for `Document.custom_xml_parts`."""

    def it_returns_an_empty_list_when_the_document_part_has_no_custom_xml(
        self, document_part_: Mock
    ):
        document_part_.custom_xml_parts = []
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.custom_xml_parts == []

    def it_delegates_to_the_document_part(self, document_part_: Mock):
        part_a = Mock(name="CustomXmlPart_A", spec=CustomXmlPart)
        part_b = Mock(name="CustomXmlPart_B", spec=CustomXmlPart)
        document_part_.custom_xml_parts = [part_a, part_b]
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.custom_xml_parts == [part_a, part_b]

    def it_returns_CustomXmlPart_instances_from_the_default_docx_template(self):
        """Integration-ish check: the default template ships with one custom XML
        part (the Bibliography "Sources" store).  We assert the proxy can read
        its item_id, schema_refs, root_element, and blob."""
        from docx import Document as open_document

        document = open_document()

        parts = document.custom_xml_parts

        assert len(parts) == 1
        part = parts[0]
        assert isinstance(part, CustomXmlPart)
        assert part.partname == "/customXml/item1.xml"
        assert part.item_id is not None and part.item_id.startswith("{")
        assert part.schema_refs  # non-empty
        assert part.root_element is not None
        assert isinstance(part.blob, bytes)

    # -- fixtures ------------------------------------------------------------

    import pytest

    @pytest.fixture
    def document_part_(self, request: FixtureRequest):
        from docx.parts.document import DocumentPart

        return instance_mock(request, DocumentPart)
