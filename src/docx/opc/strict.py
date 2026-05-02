"""Strict-OOXML → Transitional namespace translation.

The ISO/IEC 29500 Strict conformance class uses the ``purl.oclc.org`` namespace
family; the Transitional conformance class (what Microsoft Word actually reads
and writes) uses the ``schemas.openxmlformats.org`` family. A small number of
third-party producers emit Strict packages, which python-docx's element-class
lookup — keyed on Transitional URIs — cannot parse.

This module provides a conservative rewrite: detect Strict by sniffing the
main-document part's root namespace, then stream-replace every Strict URI with
its Transitional counterpart as blobs flow through ``PhysPkgReader``. Content
types and relationship URIs are rewritten too so the rest of the reader never
sees a Strict artefact.

Closes upstream#1520, upstream#693.

.. versionadded:: 1.3.0.dev0
"""

from __future__ import annotations

from typing import Iterable

#: Mapping of Strict OOXML namespace URIs → Transitional equivalents. Only the
#: prefixes we actually consume in a .docx package are listed — the Strict
#: transform spec contains roughly 40 entries but most are SpreadsheetML /
#: PresentationML. The WordprocessingML, Drawing, and OPC entries are the
#: ones that matter for a Word document.
STRICT_TO_TRANSITIONAL: dict[bytes, bytes] = {
    # -- WordprocessingML --
    b"http://purl.oclc.org/ooxml/wordprocessingml/main": (
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ),
    # -- DrawingML --
    b"http://purl.oclc.org/ooxml/drawingml/main": (
        b"http://schemas.openxmlformats.org/drawingml/2006/main"
    ),
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": (
        b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    ),
    b"http://purl.oclc.org/ooxml/drawingml/picture": (
        b"http://schemas.openxmlformats.org/drawingml/2006/picture"
    ),
    b"http://purl.oclc.org/ooxml/drawingml/chart": (
        b"http://schemas.openxmlformats.org/drawingml/2006/chart"
    ),
    b"http://purl.oclc.org/ooxml/drawingml/chartDrawing": (
        b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
    ),
    b"http://purl.oclc.org/ooxml/drawingml/diagram": (
        b"http://schemas.openxmlformats.org/drawingml/2006/diagram"
    ),
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing": (
        b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    ),
    # -- OPC / officeDocument --
    b"http://purl.oclc.org/ooxml/officeDocument/relationships": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ),
    b"http://purl.oclc.org/ooxml/officeDocument/math": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/math"
    ),
    b"http://purl.oclc.org/ooxml/officeDocument/sharedTypes": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes"
    ),
    b"http://purl.oclc.org/ooxml/officeDocument/bibliography": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
    ),
    b"http://purl.oclc.org/ooxml/officeDocument/customXml": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/customXml"
    ),
    b"http://purl.oclc.org/ooxml/officeDocument/extendedProperties": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    ),
    b"http://purl.oclc.org/ooxml/officeDocument/customProperties": (
        b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    ),
    # -- Content types --
    b"application/vnd.ms-package.obfuscated-font-uri": (
        b"application/vnd.openxmlformats-officedocument.obfuscatedFont"
    ),
}

#: The portion of each Strict URI that identifies it as Strict. A cheap
#: substring test against this sentinel lets us decide up-front whether any
#: rewriting is needed at all — the common (Transitional) path remains
#: allocation-free.
STRICT_SENTINEL = b"purl.oclc.org/ooxml"


def is_strict_document_xml(blob: bytes | None) -> bool:
    """Return True if `blob` is a ``word/document.xml`` in Strict OOXML form.

    Detection is a cheap substring scan for the Strict WordprocessingML
    namespace — the handful of producers that emit Strict packages always
    declare this URI on the ``<w:document>`` root.
    """
    if not blob:
        return False
    return b"purl.oclc.org/ooxml/wordprocessingml/main" in blob


def translate_strict_blob(blob: bytes | None) -> bytes | None:
    """Return `blob` with every Strict namespace URI rewritten to Transitional.

    Non-XML parts (images, binary blobs) are detected by the sentinel-miss
    shortcut and passed through untouched. ``None`` is preserved so callers
    can forward rels-for-missing-part lookups unchanged.
    """
    if blob is None:
        return None
    if STRICT_SENTINEL not in blob:
        return blob
    out = blob
    for strict_uri, transitional_uri in STRICT_TO_TRANSITIONAL.items():
        if strict_uri in out:
            out = out.replace(strict_uri, transitional_uri)
    return out


def translate_many(blobs: Iterable[bytes]) -> list[bytes]:
    """Apply :func:`translate_strict_blob` to every blob in `blobs`.

    Convenience helper for tests and batch processing. Production code paths
    call :func:`translate_strict_blob` one blob at a time through
    ``PhysPkgReader``.
    """
    return [translate_strict_blob(b) or b"" for b in blobs]
