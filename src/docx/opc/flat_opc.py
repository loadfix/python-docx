"""Flat-OPC (``<pkg:package>``) reader / writer.

Flat-OPC is the single-XML-file representation of an OPC package defined in
ECMA-376 Part 2. Every zipped part becomes a ``<pkg:part>`` child of a
``<pkg:package>`` root element; XML parts carry their content inline as
``<pkg:xmlData>`` while binary parts are base64-encoded under
``<pkg:binaryData>``.

python-docx's normal reader path expects a zip-format package, so the
Flat-OPC bridge in this module does the minimum necessary work: on open,
expand the flat XML into an in-memory zip blob that the existing
:class:`PhysPkgReader` accepts; on save, re-pack a zip package as Flat-OPC.

Closes upstream#892.

.. versionadded:: 1.3.0.dev0
"""

from __future__ import annotations

import base64
import io
from typing import IO, BinaryIO, Union
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

#: The ``pkg:`` namespace used by Flat-OPC. Defined in ECMA-376 Part 2, the
#: same URI is used for the root ``package`` element and every ``part``.
PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"

#: Convenience Clark-notation names for the Flat-OPC elements.
PKG_PACKAGE = f"{{{PKG_NS}}}package"
PKG_PART = f"{{{PKG_NS}}}part"
PKG_XMLDATA = f"{{{PKG_NS}}}xmlData"
PKG_BINDATA = f"{{{PKG_NS}}}binaryData"

#: Flat-OPC processing-instruction that Microsoft Word emits at the top of
#: every Flat-OPC file. We emit it too so round-tripped files remain
#: byte-identical in the preamble.
_PROGID_PI_TARGET = "mso-application"
_PROGID_PI_DATA_WORD = 'progid="Word.Document"'


def looks_like_flat_opc(pkg_file: Union[str, IO[bytes]]) -> bool:
    """Return True if `pkg_file` appears to be a Flat-OPC ``<pkg:package>`` file.

    Accepts a path or a file-like stream. Detection is a cheap byte-range
    sniff: read the first 4 KB, strip a UTF-8 BOM / leading whitespace, and
    check for ``<?xml`` followed by a ``pkg:`` namespace declaration. The
    stream position is restored so callers can reuse `pkg_file` directly.
    """
    try:
        if isinstance(pkg_file, str):
            try:
                with open(pkg_file, "rb") as f:
                    head = f.read(4096)
            except OSError:
                return False
        elif isinstance(pkg_file, (bytes, bytearray)):
            head = bytes(pkg_file[:4096])
        elif hasattr(pkg_file, "read") and hasattr(pkg_file, "seek") and hasattr(
            pkg_file, "tell"
        ):
            try:
                pos = pkg_file.tell()
            except (OSError, AttributeError, TypeError):
                return False
            try:
                head = pkg_file.read(4096)
            finally:
                try:
                    pkg_file.seek(pos)
                except (OSError, AttributeError, TypeError):
                    pass
        else:
            return False
    except Exception:
        return False
    if not head or not isinstance(head, (bytes, bytearray)):
        return False
    # -- Strip BOM and leading whitespace so comparison is predictable. --
    stripped = head.lstrip(b"\xef\xbb\xbf").lstrip()
    if not stripped.startswith(b"<?xml") and not stripped.startswith(b"<pkg:"):
        return False
    return PKG_NS.encode("ascii") in head


def expand_flat_opc_to_zip_stream(
    pkg_file: Union[str, IO[bytes]],
) -> io.BytesIO:
    """Return a :class:`io.BytesIO` zip stream built from Flat-OPC `pkg_file`.

    Every ``<pkg:part>`` in the source is materialised as a zip member.
    Binary parts are base64-decoded; XML parts are re-serialised with the
    OPC-standard XML declaration. The returned stream is positioned at
    offset 0 so callers can hand it directly to :class:`zipfile.ZipFile`.
    """
    if isinstance(pkg_file, str):
        with open(pkg_file, "rb") as f:
            blob = f.read()
    else:
        blob = pkg_file.read()
    root = etree.fromstring(blob)
    if root.tag != PKG_PACKAGE:
        raise ValueError(
            "Flat-OPC input does not have a <pkg:package> root element "
            "(got %r)" % root.tag
        )
    out = io.BytesIO()
    with ZipFile(out, "w", compression=ZIP_DEFLATED) as zf:
        for part_elm in root.findall(PKG_PART):
            member_name = _member_name_for_part(part_elm)
            payload = _payload_for_part(part_elm)
            zf.writestr(member_name, payload)
    out.seek(0)
    return out


def write_flat_opc(pkg_file: Union[str, IO[bytes]], zip_blob: bytes) -> None:
    """Serialise a zip-format package `zip_blob` as Flat-OPC to `pkg_file`.

    Binary members are base64-encoded inside ``<pkg:binaryData>``; XML members
    are embedded inline inside ``<pkg:xmlData>``. `pkg_file` may be a path
    or a writable binary stream.
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
                part_elm.set(f"{{{PKG_NS}}}contentType", content_type or _xml_content_type(member))
                xmldata = etree.SubElement(part_elm, PKG_XMLDATA, nsmap=nsmap)
                try:
                    inner = etree.fromstring(data)
                    xmldata.append(inner)
                except etree.XMLSyntaxError:
                    # -- Fall back to binary embedding for junk we can't parse.
                    part_elm.remove(xmldata)
                    part_elm.set(
                        f"{{{PKG_NS}}}compression", "store"
                    )
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


def _member_name_for_part(part_elm: "etree._Element") -> str:  # pyright: ignore[reportPrivateUsage]
    """Return the zip member name for `<pkg:part>` element `part_elm`."""
    name_attr = part_elm.get(f"{{{PKG_NS}}}name") or ""
    # -- OPC part-names start with "/"; zip member names don't. --
    return name_attr.lstrip("/")


def _payload_for_part(part_elm: "etree._Element") -> bytes:  # pyright: ignore[reportPrivateUsage]
    """Return the byte payload for `<pkg:part>` element `part_elm`."""
    xmldata = part_elm.find(PKG_XMLDATA)
    if xmldata is not None:
        # -- A <pkg:xmlData> container holds exactly one element child which is
        # -- the actual part contents; serialise that child as a standalone XML document. --
        children = list(xmldata)
        if not children:
            return b""
        return etree.tostring(
            children[0], xml_declaration=True, encoding="UTF-8", standalone=True
        )
    bindata = part_elm.find(PKG_BINDATA)
    if bindata is not None and bindata.text is not None:
        return base64.b64decode(bindata.text)
    return b""


def _is_xml_member(member: str, data: bytes) -> bool:
    """Return True if zip member `member`/`data` should be embedded as XML."""
    lower = member.lower()
    if lower.endswith((".xml", ".rels")):
        return True
    # -- Heuristic fallback: anything beginning with an XML prolog. --
    return data[:5] == b"<?xml"


def _xml_content_type(member: str) -> str:
    """Return a sensible Flat-OPC content-type attribute for XML `member`.

    Flat-OPC parts carry an explicit content-type attribute so readers don't
    have to cross-reference ``[Content_Types].xml`` on the fly. We emit a
    generic ``application/xml`` default for members whose content type we
    can't infer from the filename alone.
    """
    lower = member.lower()
    if lower.endswith(".rels"):
        return "application/vnd.openxmlformats-package.relationships+xml"
    return "application/xml"


#: Map of zip member path prefix → content type for non-XML binary parts. This
#: is deliberately tiny; Word tolerates a missing/wrong content-type attribute
#: on binary parts because ``[Content_Types].xml`` remains the source of truth.
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


def _guess_content_type(member: str) -> str:
    """Best-effort content type for a zip member. Empty string if unknown."""
    lower = member.lower()
    for ext, ct in _BINARY_CONTENT_TYPES.items():
        if lower.endswith(ext):
            return ct
    return ""


def read_flat_opc(pkg_file: Union[str, BinaryIO]) -> io.BytesIO:
    """Public alias for :func:`expand_flat_opc_to_zip_stream`.

    Returns an in-memory zip stream that the standard reader can consume.
    """
    return expand_flat_opc_to_zip_stream(pkg_file)
