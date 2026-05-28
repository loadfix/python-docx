"""Custom XML data-source binding for content controls (issue #80).

A *data source* is a named ``/customXml/item{N}.xml`` part attached to the
document. Content controls (SDTs) declare a ``<w:dataBinding>`` whose
``@w:xpath`` is evaluated against the bound source's payload at save time;
the resolved string is also inlined into the SDT's ``<w:sdtContent>``.

See :meth:`docx.document.Document.bind_data_source` for the public surface.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import os
import uuid
from typing import IO, TYPE_CHECKING, Any, Optional, Union, cast

from lxml import etree

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.parts.custom_xml import CustomXmlPart as CustomXmlDataPart

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.parts.document import DocumentPart


__all__ = [
    "DataSource",
    "DataSourceValidationError",
    "bind_data_source",
    "iter_bound_sources",
    "resolve_bindings_in_document",
]


_DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
_LFXBIND_NS = "https://loadfix.dev/python-docx/xmlns/data-binding"
_LFXBIND_TAG = f"{{{_LFXBIND_NS}}}name"


class DataSourceValidationError(ValueError):
    """Raised when an XSD validation pass over a bound payload fails.

    Carries the list of :class:`ooxml_customxml.ValidationIssue` records
    on the ``.issues`` attribute. The first failing issue's ``message`` is
    used as the exception text for readable tracebacks.

    .. versionadded:: 2026.05.13
    """

    def __init__(self, issues: list[Any]):
        self.issues = issues
        first = issues[0] if issues else None
        msg = "custom XML payload failed XSD validation"
        if first is not None:
            line = getattr(first, "line", None)
            text = getattr(first, "message", str(first))
            msg = f"{msg}: {text}" + (f" (line {line})" if line else "")
        super().__init__(msg)


# ---------------------------------------------------------------------------
# DataSource — proxy over a single bound /customXml/item{N}.xml part


class DataSource:
    """Read-only proxy describing one bound custom-XML data source.

    Carries the logical ``name`` (as supplied to :meth:`bind_data_source`),
    the ``store_item_id`` Word uses to wire SDTs to this part, and the
    parsed payload root.

    .. versionadded:: 2026.05.13
    """

    def __init__(
        self,
        name: str,
        data_part: CustomXmlDataPart,
        store_item_id: str,
    ):
        self._name = name
        self._data_part = data_part
        self._store_item_id = store_item_id

    @property
    def name(self) -> str:
        """Logical id supplied at :meth:`bind_data_source` time."""
        return self._name

    @property
    def store_item_id(self) -> str:
        """``{GUID}`` value Word uses to anchor SDT bindings to this part."""
        return self._store_item_id

    @property
    def part(self) -> CustomXmlDataPart:
        """Underlying ``/customXml/item{N}.xml`` part."""
        return self._data_part

    @property
    def partname(self) -> str:
        """Pack URI of the data part."""
        return str(self._data_part.partname)

    @property
    def root_element(self) -> "etree._Element | None":
        """Parsed payload root, or |None| on parse failure."""
        try:
            return parse_xml(self._data_part.blob)
        except etree.XMLSyntaxError:
            return None


# ---------------------------------------------------------------------------
# helpers


def _read_payload(path_or_blob: "Union[str, bytes, os.PathLike[str], IO[bytes]]") -> bytes:
    """Coerce ``path_or_blob`` to a payload byte-string.

    Accepts a filesystem path, a binary blob, or an open binary file-like.
    """
    if isinstance(path_or_blob, (bytes, bytearray)):
        return bytes(path_or_blob)
    if hasattr(path_or_blob, "read"):
        return cast(IO[bytes], path_or_blob).read()
    with open(os.fspath(path_or_blob), "rb") as fh:
        return fh.read()


def _read_schema(path_or_blob: "Union[str, bytes, os.PathLike[str], IO[bytes], None]") -> "Optional[bytes]":
    """Coerce ``path_or_blob`` to schema bytes, or |None| when absent."""
    if path_or_blob is None:
        return None
    return _read_payload(path_or_blob)


def _validate_payload(payload: bytes, schema: bytes) -> None:
    """Validate ``payload`` against ``schema``; raise on hard failures."""
    try:
        from ooxml_customxml import validate as _validate
    except ImportError:  # pragma: no cover - optional dep
        return
    issues = list(_validate(payload, schemas=[schema]))
    failing = [
        issue
        for issue in issues
        if getattr(issue, "severity", "") in {"error", "fatal"}
    ]
    if failing:
        raise DataSourceValidationError(failing)


def _build_props_blob(store_item_id: str, schema_uri: "Optional[str]") -> bytes:
    """Return a minimal ``ds:datastoreItem`` XML blob for the props part."""
    schema_block = b""
    if schema_uri:
        schema_block = (
            b"<ds:schemaRefs>"
            + (
                b'<ds:schemaRef ds:uri="'
                + schema_uri.encode("utf-8")
                + b'"/>'
            )
            + b"</ds:schemaRefs>"
        )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + b"<ds:datastoreItem "
        + b'xmlns:ds="' + _DS_NS.encode("ascii") + b'" '
        + b'ds:itemID="' + store_item_id.encode("ascii") + b'">'
        + schema_block
        + b"</ds:datastoreItem>"
    )


def _props_partname_for(data_partname: str) -> PackURI:
    """Return the sibling ``itemProps{N}.xml`` partname for a data part."""
    if not data_partname.startswith("/customXml/item") or not data_partname.endswith(".xml"):
        raise ValueError(f"unexpected data partname {data_partname!r}")
    stem = data_partname[len("/customXml/item") : -len(".xml")]
    return PackURI(f"/customXml/itemProps{stem}.xml")


def _next_data_partname(package: "OpcPackage") -> PackURI:
    """Return the first free ``/customXml/item{N}.xml`` partname."""
    used: set[int] = set()
    for part in package.iter_parts():
        partname = str(part.partname)
        if not partname.startswith("/customXml/item"):
            continue
        tail = partname[len("/customXml/item") :]
        if not tail.endswith(".xml"):
            continue
        stem = tail[: -len(".xml")]
        if stem.startswith("Props"):
            continue
        try:
            used.add(int(stem))
        except ValueError:
            continue
    n = 1
    while n in used:
        n += 1
    return PackURI(f"/customXml/item{n}.xml")


def _stamp_name_marker(part: CustomXmlDataPart, name: str) -> None:
    """Stash the logical ``name`` on the data part as a Python attribute."""
    part._lfx_data_source_name = name  # type: ignore[attr-defined]


def _name_marker(part: CustomXmlDataPart) -> "Optional[str]":
    """Return the logical name stamped on ``part`` (in-memory marker only)."""
    return getattr(part, "_lfx_data_source_name", None)


def _document_namespace_of(payload: bytes) -> "Optional[str]":
    """Return the default-namespace URI of the payload's document element."""
    try:
        root = parse_xml(payload)
    except etree.XMLSyntaxError:
        return None
    tag = root.tag
    if isinstance(tag, str) and tag.startswith("{"):
        return tag[1 : tag.index("}")]
    return None


def _root_local_name(payload: bytes) -> "Optional[str]":
    """Return the local name of the payload's document element."""
    try:
        root = parse_xml(payload)
    except etree.XMLSyntaxError:
        return None
    tag = root.tag
    if isinstance(tag, str):
        if tag.startswith("{"):
            return tag.split("}", 1)[1]
        return tag
    return None


# ---------------------------------------------------------------------------
# Core authoring path


def bind_data_source(
    document_part: "DocumentPart",
    path_or_blob: "Union[str, bytes, os.PathLike[str], IO[bytes]]",
    name: str,
    schema: "Union[str, bytes, os.PathLike[str], IO[bytes], None]" = None,
) -> DataSource:
    """Attach (or replace) a data source under logical id ``name``.

    See :meth:`docx.document.Document.bind_data_source` for the user-facing
    contract.

    .. versionadded:: 2026.05.13
    """
    if not name:
        raise ValueError("name must be a non-empty string")

    payload = _read_payload(path_or_blob)
    schema_blob = _read_schema(schema)
    if schema_blob is not None:
        _validate_payload(payload, schema_blob)

    package = document_part.package
    assert package is not None

    # -- look for an existing source with this name ----------------------
    existing = _find_data_source_part(document_part, name)
    if existing is not None:
        existing._blob = payload  # type: ignore[attr-defined]
        store_item_id = _existing_store_item_id(existing) or _new_store_item_id()
        # -- rewrite the props part to keep the schemaRef list current --
        _ensure_props_part(
            existing,
            store_item_id,
            schema_uri=_document_namespace_of(payload),
            package=package,
        )
        # -- ensure the in-memory marker is current so subsequent saves
        #    re-stamp the new payload with the same logical name --
        _stamp_name_marker(existing, name)
        return DataSource(name, existing, store_item_id)

    # -- create a fresh data part ----------------------------------------
    data_partname = _next_data_partname(package)
    data_part = CustomXmlDataPart(data_partname, CT.XML, payload)
    _stamp_name_marker(data_part, name)
    document_part.relate_to(data_part, RT.CUSTOM_XML)

    store_item_id = _new_store_item_id()
    _ensure_props_part(
        data_part,
        store_item_id,
        schema_uri=_document_namespace_of(payload),
        package=package,
    )
    return DataSource(name, data_part, store_item_id)


def _new_store_item_id() -> str:
    """Mint a fresh ``{GUID}``-formatted store-item id."""
    return "{" + str(uuid.uuid4()).upper() + "}"


def _existing_store_item_id(data_part: CustomXmlDataPart) -> "Optional[str]":
    """Read the ``{GUID}`` from the sibling props part, when present."""
    try:
        props = data_part.part_related_by(RT.CUSTOM_XML_PROPS)
    except KeyError:
        return None
    blob = getattr(props, "blob", b"")
    if not blob:
        return None
    try:
        root = parse_xml(blob)
    except etree.XMLSyntaxError:
        return None
    return root.get(f"{{{_DS_NS}}}itemID")


def _ensure_props_part(
    data_part: CustomXmlDataPart,
    store_item_id: str,
    schema_uri: "Optional[str]",
    package: "OpcPackage",
) -> CustomXmlDataPart:
    """Create or update the sibling ``itemProps{N}.xml`` part for ``data_part``."""
    blob = _build_props_blob(store_item_id, schema_uri)
    try:
        existing = data_part.part_related_by(RT.CUSTOM_XML_PROPS)
        if hasattr(existing, "_blob"):
            existing._blob = blob  # type: ignore[attr-defined]
        else:
            # -- fall back to attribute write; Part stores its bytes on _blob --
            existing._blob = blob  # type: ignore[attr-defined]
        return cast(CustomXmlDataPart, existing)
    except KeyError:
        pass
    props_partname = _props_partname_for(str(data_part.partname))
    props_part = CustomXmlDataPart(
        props_partname, CT.OFC_CUSTOM_XML_PROPERTIES, blob
    )
    data_part.relate_to(props_part, RT.CUSTOM_XML_PROPS)
    return props_part


def _find_data_source_part(
    document_part: "DocumentPart", name: str
) -> "Optional[CustomXmlDataPart]":
    """Locate the existing data-source part registered under ``name``, if any."""
    for rel in document_part.rels.values():
        if rel.is_external or rel.reltype != RT.CUSTOM_XML:
            continue
        try:
            target = rel.target_part
        except ValueError:
            continue
        if not isinstance(target, CustomXmlDataPart):
            continue
        in_mem = _name_marker(target)
        if in_mem == name:
            return target
        if in_mem is None and recover_name_from_payload(target) == name:
            _stamp_name_marker(target, name)
            return target
    return None


def iter_bound_sources(document_part: "DocumentPart") -> "list[DataSource]":
    """Return a |DataSource| for every named source on ``document_part``.

    Sources without a logical-name marker (plain ``customXml`` parts loaded
    from a package without a ``lfxbind:name`` attribute on the payload root)
    are skipped. Falls back to scanning each part's payload to recover the
    marker on the first call after a reload.

    .. versionadded:: 2026.05.13
    """
    result: list[DataSource] = []
    for rel in document_part.rels.values():
        if rel.is_external or rel.reltype != RT.CUSTOM_XML:
            continue
        try:
            target = rel.target_part
        except ValueError:
            continue
        if not isinstance(target, CustomXmlDataPart):
            continue
        name = _name_marker(target)
        if name is None:
            name = recover_name_from_payload(target)
            if name is None:
                continue
            _stamp_name_marker(target, name)
        store_item_id = _existing_store_item_id(target)
        if store_item_id is None:
            continue
        result.append(DataSource(name, target, store_item_id))
    return result


# ---------------------------------------------------------------------------
# Save-time binding resolution


def resolve_bindings_in_document(document_part: "DocumentPart") -> int:
    """Walk every SDT in the body, resolve its data-binding XPath, inline the value.

    Returns the count of SDTs whose displayed text was replaced. Misses
    (unbound source, xpath returns |None|, payload won't parse) leave the
    prior content untouched.

    .. versionadded:: 2026.05.13
    """
    try:
        from ooxml_customxml import (
            CustomXmlMapping,
            resolve_binding,
        )
    except ImportError:  # pragma: no cover - optional dep
        return 0

    sources_by_id: dict[str, DataSource] = {
        src.store_item_id: src for src in iter_bound_sources(document_part)
    }
    if not sources_by_id:
        return 0

    body = document_part.element.body
    replaced = 0
    for sdt in body.iter(qn("w:sdt")):
        sdtPr = sdt.find(qn("w:sdtPr"))
        if sdtPr is None:
            continue
        data_binding = sdtPr.find(qn("w:dataBinding"))
        if data_binding is None:
            continue
        store_item_id = data_binding.get(qn("w:storeItemID"))
        if not store_item_id:
            continue
        source = sources_by_id.get(store_item_id)
        if source is None:
            continue
        root = source.root_element
        if root is None:
            continue
        mapping = CustomXmlMapping(data_binding)
        try:
            value = resolve_binding(mapping, root)
        except Exception:
            value = None
        if value is None:
            continue
        if _replace_sdt_content(sdt, value):
            replaced += 1
    return replaced


def _replace_sdt_content(sdt: "etree._Element", value: str) -> bool:
    """Replace the ``<w:sdtContent>`` of ``sdt`` with a single run of ``value``."""
    sdtContent = sdt.find(qn("w:sdtContent"))
    if sdtContent is None:
        sdtContent = OxmlElement("w:sdtContent")
        sdt.append(sdtContent)

    parent = sdt.getparent()
    is_inline = parent is not None and parent.tag == qn("w:p")
    if not is_inline:
        # -- inspect existing content as a fallback when not yet attached --
        for child in sdtContent:
            if child.tag == qn("w:r"):
                is_inline = True
                break
            if child.tag == qn("w:p"):
                is_inline = False
                break

    for child in list(sdtContent):
        sdtContent.remove(child)

    text_value = "" if value is None else str(value)
    if is_inline:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        if text_value != text_value.strip():
            t.set(qn("xml:space"), "preserve")
        t.text = text_value
        r.append(t)
        sdtContent.append(r)
    else:
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        if text_value != text_value.strip():
            t.set(qn("xml:space"), "preserve")
        t.text = text_value
        r.append(t)
        p.append(r)
        sdtContent.append(p)
    return True


# ---------------------------------------------------------------------------
# Round-trip support — persist the logical-name marker into the payload so
# the next reload re-discovers the source under the same name.


def stamp_name_into_payload(data_part: CustomXmlDataPart, name: str) -> None:
    """Persist the logical name into the payload as a fork-scoped attribute."""
    blob = data_part.blob
    if not blob:
        return
    try:
        root = parse_xml(blob)
    except etree.XMLSyntaxError:
        return
    root.set(_LFXBIND_TAG, name)
    nsmap = root.nsmap or {}
    if "lfxbind" not in nsmap:
        # -- lxml doesn't allow rebinding nsmap on an existing tree, so emit
        #    a fresh declaration manually via a parser round-trip. The new
        #    nsmap is set via re-creating the root with the merged map. --
        new_nsmap = dict(nsmap)
        new_nsmap["lfxbind"] = _LFXBIND_NS
        new_root = etree.Element(root.tag, attrib=root.attrib, nsmap=new_nsmap)
        for child in root:
            new_root.append(child)
        if root.text is not None:
            new_root.text = root.text
        root = new_root
    data_part._blob = etree.tostring(  # type: ignore[attr-defined]
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def recover_name_from_payload(data_part: CustomXmlDataPart) -> "Optional[str]":
    """Return the logical-name marker stamped into the payload, if any."""
    blob = getattr(data_part, "blob", b"") or b""
    if not blob:
        return None
    try:
        root = parse_xml(blob)
    except etree.XMLSyntaxError:
        return None
    return root.get(_LFXBIND_TAG)
