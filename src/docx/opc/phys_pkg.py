"""Provides a general interface to a `physical` OPC package, such as a zip file."""

from __future__ import annotations

import os
from typing import IO
from zipfile import ZIP_DEFLATED, ZipFile, is_zipfile

from docx.exceptions import EncryptedDocumentError
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.packuri import CONTENT_TYPES_URI
from docx.opc.strict import STRICT_SENTINEL, translate_strict_blob

#: OLE compound file (CFBF) binary signature. Encrypted Office documents are wrapped
#: in this container rather than the usual ZIP package.
_OLE_COMPOUND_FILE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

_ENCRYPTED_DOCUMENT_MESSAGE = (
    "Document is password-protected (encrypted .docx detected). "
    "Install msoffcrypto-tool to decrypt it first: "
    "https://github.com/nolze/msoffcrypto-tool"
)


def _raise_if_encrypted_path(path: str) -> None:
    """Raise |EncryptedDocumentError| if file at `path` has the OLE signature."""
    try:
        with open(path, "rb") as f:
            header = f.read(len(_OLE_COMPOUND_FILE_SIGNATURE))
    except OSError:
        return
    if header == _OLE_COMPOUND_FILE_SIGNATURE:
        raise EncryptedDocumentError(_ENCRYPTED_DOCUMENT_MESSAGE)


def _raise_if_encrypted_stream(stream: IO[bytes]) -> None:
    """Raise |EncryptedDocumentError| if `stream` begins with the OLE signature.

    The stream position is restored after the peek.
    """
    if not hasattr(stream, "read"):
        return
    try:
        pos = stream.tell()
    except (OSError, AttributeError):
        # Not seekable — we can't safely peek without consuming bytes.
        return
    try:
        header = stream.read(len(_OLE_COMPOUND_FILE_SIGNATURE))
    finally:
        try:
            stream.seek(pos)
        except (OSError, AttributeError):
            pass
    if header == _OLE_COMPOUND_FILE_SIGNATURE:
        raise EncryptedDocumentError(_ENCRYPTED_DOCUMENT_MESSAGE)


class PhysPkgReader:
    """Factory for physical package reader objects.

    Chooses the concrete reader matching `pkg_file` (directory vs zip). When
    the package turns out to be Strict OOXML (detected by sniffing
    ``[Content_Types].xml`` or ``/word/document.xml`` for the Strict
    namespace sentinel), the concrete reader is wrapped in
    :class:`_StrictTranslatingPkgReader` so blobs are rewritten to
    Transitional as they're read. Closes upstream#1520, upstream#693.

    .. versionchanged:: 1.3.0.dev0
       Transparent Strict → Transitional translation on open.
    """

    def __new__(cls, pkg_file):
        # if `pkg_file` is a string, treat it as a path
        if isinstance(pkg_file, str):
            if os.path.isdir(pkg_file):
                reader_cls = _DirPkgReader
            elif is_zipfile(pkg_file):
                reader_cls = _ZipPkgReader
            else:
                # -- check for password-encrypted .docx (OLE compound file) before
                # -- reporting "not found", so users get an actionable error message.
                _raise_if_encrypted_path(pkg_file)
                raise PackageNotFoundError("Package not found at '%s'" % pkg_file)
        else:  # assume it's a stream and pass it to Zip reader to sort out
            _raise_if_encrypted_stream(pkg_file)
            reader_cls = _ZipPkgReader

        return super(PhysPkgReader, cls).__new__(reader_cls)


def _looks_like_strict_package(reader) -> bool:
    """Return True if `reader` exposes a Strict-OOXML package.

    Sniffs ``[Content_Types].xml`` first (cheap, usually decisive); if that
    is Transitional but the main document part is Strict — produced by some
    conversion tools — the content-types check misses, so we fall back to
    peeking at ``/word/document.xml``. A substring match against the Strict
    sentinel ``purl.oclc.org/ooxml`` is false-negative-free: Transitional
    packages never contain it.
    """
    try:
        ct_blob = reader.content_types_xml
    except (KeyError, IOError, ValueError):
        return False
    if ct_blob is not None and STRICT_SENTINEL in ct_blob:
        return True
    try:
        from docx.opc.packuri import PackURI

        doc_blob = reader.blob_for(PackURI("/word/document.xml"))
    except (KeyError, IOError, ValueError):
        return False
    return doc_blob is not None and STRICT_SENTINEL in doc_blob


def open_phys_pkg_reader(pkg_file):
    """Return a physical package reader for `pkg_file` with Strict translation.

    Wraps the concrete :class:`PhysPkgReader` subclass in
    :class:`_StrictTranslatingPkgReader` when the package is Strict OOXML.
    Called from :class:`docx.opc.pkgreader.PackageReader.from_file` in place
    of direct construction.

    .. versionadded:: 1.3.0.dev0
    """
    reader = PhysPkgReader(pkg_file)
    if _looks_like_strict_package(reader):
        return _StrictTranslatingPkgReader(reader)
    return reader


class _StrictTranslatingPkgReader:
    """Wraps a physical reader and rewrites Strict URIs as blobs are read.

    Forwards every PhysPkgReader method to the wrapped reader, applying
    :func:`docx.opc.strict.translate_strict_blob` to each returned blob. The
    wrapped reader retains sole ownership of the underlying zip handle /
    directory, so ``close()`` still delegates.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, inner):
        self._inner = inner

    def blob_for(self, pack_uri):
        blob = self._inner.blob_for(pack_uri)
        return translate_strict_blob(blob)

    def close(self):
        self._inner.close()

    @property
    def content_types_xml(self):
        return translate_strict_blob(self._inner.content_types_xml)

    def rels_xml_for(self, source_uri):
        return translate_strict_blob(self._inner.rels_xml_for(source_uri))


class PhysPkgWriter:
    """Factory for physical package writer objects."""

    def __new__(cls, pkg_file):
        return super(PhysPkgWriter, cls).__new__(_ZipPkgWriter)


class _DirPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for an OPC package extracted into a
    directory."""

    def __init__(self, path):
        """`path` is the path to a directory containing an expanded package."""
        super().__init__()
        self._path = os.path.abspath(path)

    def blob_for(self, pack_uri):
        """Return contents of file corresponding to `pack_uri` in package directory."""
        path = os.path.join(self._path, pack_uri.membername)
        # Guard against path traversal — resolved path must remain within package dir
        real_path = os.path.realpath(path)
        real_root = os.path.realpath(self._path)
        if not real_path.startswith(real_root + os.sep) and real_path != real_root:
            raise ValueError(
                "Pack URI '%s' resolves outside package directory" % pack_uri
            )
        with open(path, "rb") as f:
            blob = f.read()
        return blob

    def close(self):
        """Provides interface consistency with |ZipFileSystem|, but does nothing, a
        directory file system doesn't need closing."""
        pass

    @property
    def content_types_xml(self):
        """Return the `[Content_Types].xml` blob from the package."""
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """Return rels item XML for source with `source_uri`, or None if the item has no
        rels item."""
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except IOError:
            rels_xml = None
        return rels_xml


class _ZipPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for a zip file OPC package."""

    def __init__(self, pkg_file):
        super().__init__()
        self._zipf = ZipFile(pkg_file, "r")

    def blob_for(self, pack_uri):
        """Return blob corresponding to `pack_uri`.

        Raises |ValueError| if no matching member is present in zip archive.
        """
        return self._zipf.read(pack_uri.membername)

    def close(self):
        """Close the zip archive, releasing any resources it is using."""
        self._zipf.close()

    @property
    def content_types_xml(self):
        """Return the `[Content_Types].xml` blob from the zip package."""
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """Return rels item XML for source with `source_uri` or None if no rels item is
        present."""
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except KeyError:
            rels_xml = None
        return rels_xml


class _ZipPkgWriter(PhysPkgWriter):
    """Implements |PhysPkgWriter| interface for a zip file OPC package."""

    def __init__(self, pkg_file):
        super().__init__()
        self._zipf = ZipFile(pkg_file, "w", compression=ZIP_DEFLATED)

    def close(self):
        """Close the zip archive, flushing any pending physical writes and releasing any
        resources it's using."""
        self._zipf.close()

    def write(self, pack_uri, blob):
        """Write `blob` to this zip package with the membername corresponding to
        `pack_uri`."""
        self._zipf.writestr(pack_uri.membername, blob)
