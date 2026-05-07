"""Re-export of :mod:`ooxml_opc.flat_opc` with docx-local helpers.

The Flat-OPC sniff (:func:`looks_like_flat_opc`), zip-expander
(:func:`expand_flat_opc_to_zip_stream`) and :class:`FlatOpcWriter`
live in :mod:`ooxml_opc.flat_opc` and are re-exported from here so
external callers and docx-internal code can continue to import from
``docx.opc.flat_opc``.

docx-local function retained:

* :func:`write_flat_opc` — docx historically exposed a function that takes
  ``(pkg_file, zip_blob)`` where ``zip_blob`` is the bytes of an already-
  built zip package. Internally it unpacks the zip, emits the corresponding
  Flat-OPC XML, and writes to ``pkg_file``. The shared library's
  :class:`FlatOpcWriter` takes ``(pkg_file, pkg_rels, parts)`` instead,
  which requires a live ``OpcPackage`` — not always what docx callers have
  on hand. The function form is preserved for backward-compat.

.. versionchanged:: 2026.05.11
   Re-exported from :mod:`ooxml_opc.flat_opc`; :func:`write_flat_opc`
   remains docx-local.
"""

from __future__ import annotations

import base64
import io
from typing import IO, BinaryIO, Union
from zipfile import ZipFile

from lxml import etree
from ooxml_opc.flat_opc import (  # noqa: F401 -- re-exports
    PKG_BINDATA,
    PKG_NS,
    PKG_PACKAGE,
    PKG_PART,
    PKG_XMLDATA,
    FlatOpcWriter,
    expand_flat_opc_to_zip_stream,
    looks_like_flat_opc,
    read_flat_opc,
)

__all__ = [
    "PKG_BINDATA",
    "PKG_NS",
    "PKG_PACKAGE",
    "PKG_PART",
    "PKG_XMLDATA",
    "FlatOpcWriter",
    "expand_flat_opc_to_zip_stream",
    "looks_like_flat_opc",
    "read_flat_opc",
    "write_flat_opc",
]


#: Flat-OPC processing-instruction that Microsoft Word emits at the top of
#: every Flat-OPC file. We emit it too so round-tripped files remain
#: byte-identical in the preamble.
_PROGID_PI_TARGET = "mso-application"
_PROGID_PI_DATA_WORD = 'progid="Word.Document"'

#: Map of zip member path prefix → content type for non-XML binary parts.
#: This is deliberately tiny; Word tolerates a missing/wrong content-type
#: attribute on binary parts because ``[Content_Types].xml`` remains the
#: source of truth.
_BINARY_CONTENT_TYPES: dict[str, str] = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
    ".bin": "application/vnd.ms-office.vbaProject",
}


def write_flat_opc(pkg_file: Union[str, IO[bytes]], zip_blob: bytes) -> None:
    """Serialise a zip-format package `zip_blob` as Flat-OPC to `pkg_file`.

    Binary members are base64-encoded inside ``<pkg:binaryData>``; XML members
    are embedded inline inside ``<pkg:xmlData>``. `pkg_file` may be a path or
    a writable binary stream.

    This is docx-local — the shared :class:`FlatOpcWriter` takes a live
    ``(pkg_rels, parts)`` pair whereas many docx call-sites already have a
    serialised zip blob in memory.
    """
    nsmap = {"pkg": PKG_NS}
    root = etree.Element(PKG_PACKAGE, nsmap=nsmap)
    tree = etree.ElementTree(root)
    pi = etree.ProcessingInstruction(_PROGID_PI_TARGET, _PROGID_PI_DATA_WORD)
    root.addprevious(pi)
    with ZipFile(io.BytesIO(zip_blob), "r") as zf:
        for info in zf.infolist():
            member = info.filename
            data = zf.read(member)
            part_elm = etree.SubElement(root, PKG_PART, nsmap=nsmap)
            part_elm.set(f"{{{PKG_NS}}}name", "/" + member)
            content_type = _guess_content_type(member)
            if content_type:
                part_elm.set(f"{{{PKG_NS}}}contentType", content_type)
            if _is_xml_member(member, data):
                part_elm.set(
                    f"{{{PKG_NS}}}contentType", content_type or _xml_content_type(member)
                )
                xmldata = etree.SubElement(part_elm, PKG_XMLDATA, nsmap=nsmap)
                try:
                    inner = etree.fromstring(data)
                    xmldata.append(inner)
                except etree.XMLSyntaxError:
                    # -- Fall back to binary embedding for junk we can't parse.
                    part_elm.remove(xmldata)
                    part_elm.set(f"{{{PKG_NS}}}compression", "store")
                    bindata = etree.SubElement(part_elm, PKG_BINDATA, nsmap=nsmap)
                    bindata.text = base64.b64encode(data).decode("ascii")
            else:
                bindata = etree.SubElement(part_elm, PKG_BINDATA, nsmap=nsmap)
                bindata.text = base64.b64encode(data).decode("ascii")
    # -- Serialise to a byte buffer so both path and stream targets share one code path. --
    buf = io.BytesIO()
    tree.write(
        buf,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    data = buf.getvalue()
    if isinstance(pkg_file, str):
        with open(pkg_file, "wb") as f:
            f.write(data)
    else:
        pkg_file.write(data)


def _is_xml_member(member: str, data: bytes) -> bool:
    """Return True if zip member `member`/`data` should be embedded as XML."""
    lower = member.lower()
    if lower.endswith((".xml", ".rels")):
        return True
    # -- Heuristic fallback: anything beginning with an XML prolog. --
    return data[:5] == b"<?xml"


def _xml_content_type(member: str) -> str:
    """Return a sensible Flat-OPC content-type attribute for XML `member`."""
    lower = member.lower()
    if lower.endswith(".rels"):
        return "application/vnd.openxmlformats-package.relationships+xml"
    return "application/xml"


def _guess_content_type(member: str) -> str:
    """Best-effort content type for a zip member. Empty string if unknown."""
    lower = member.lower()
    for ext, ct in _BINARY_CONTENT_TYPES.items():
        if lower.endswith(ext):
            return ct
    return ""


# -- `BinaryIO` kept imported for backward-compat type-hint imports; still used by
# -- some callers that typed the public `write_flat_opc` signature.
_ = BinaryIO  # noqa: F841 -- reserved re-export
