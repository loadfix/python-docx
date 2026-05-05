"""Provides a general interface to a `physical` OPC package, such as a zip file."""

from __future__ import annotations

import io
import os
from typing import IO
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo, is_zipfile

from docx.exceptions import EncryptedDocumentError, RmsProtectedDocumentError
from docx.opc.exceptions import (  # noqa: F401 -- PackageNotFoundError re-export
    MissingDocxFileError,
    NotADocxError,
    PackageNotFoundError,
)
from docx.opc.packuri import CONTENT_TYPES_URI
from docx.opc.strict import STRICT_SENTINEL, translate_strict_blob

#: Fixed timestamp used when writing a reproducible package. The 1980-01-01 epoch
#: is the minimum the ZIP format can express; any ``ZipInfo.date_time`` earlier
#: than that is silently clamped by :mod:`zipfile`. Matches the convention used by
#: reproducible-build tooling (e.g. `SOURCE_DATE_EPOCH`-aware writers).
REPRODUCIBLE_TIMESTAMP = (1980, 1, 1, 0, 0, 0)

#: OLE compound file (CFBF) binary signature. Encrypted Office documents are wrapped
#: in this container rather than the usual ZIP package.
_OLE_COMPOUND_FILE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

_ENCRYPTED_DOCUMENT_MESSAGE = (
    "Document is password-protected (encrypted .docx detected). "
    "Pass `password=...` to `Document(...)` to decrypt it, or install the "
    "optional 'python-ooxml-crypto' package "
    "(https://github.com/loadfix/python-ooxml-crypto)."
)

_RMS_PROTECTED_DOCUMENT_MESSAGE = (
    "package is wrapped in Azure RMS / AIP / IRM protection; python-docx "
    "cannot decrypt RMS-protected files — the payload is encrypted to the "
    "user's Microsoft 365 identity, not a password. Delegate decryption to "
    "Microsoft Office automation or the Microsoft Information Protection SDK "
    "before opening the file with python-docx."
)


def _decrypt_to_stream(stream: IO[bytes], password: str) -> io.BytesIO:
    """Return a BytesIO of plaintext zip bytes decrypted from encrypted `stream`.

    Also raises :class:`RmsProtectedDocumentError` when the stream is wrapped in
    Azure RMS / AIP / IRM protection — python-ooxml-crypto cannot decrypt those
    since the payload is keyed to an Azure AD identity rather than a password.
    """
    from docx.opc._crypto import (
        decrypt_stream,
        is_rms_protected_stream,
    )

    if is_rms_protected_stream(stream):
        raise RmsProtectedDocumentError(_RMS_PROTECTED_DOCUMENT_MESSAGE)

    plain = decrypt_stream(stream, password)
    return io.BytesIO(plain)


def _maybe_decrypt_path(path: str, password: str | None) -> io.BytesIO | None:
    """Return decrypted-payload BytesIO if file at `path` is encrypted, else None.

    Raises :class:`EncryptedDocumentError` when the file is encrypted and
    either `password` is None or ``python-ooxml-crypto`` is not installed.
    Returns None when the file is not encrypted (or does not exist / is
    unreadable — the caller handles those outcomes).
    """
    try:
        with open(path, "rb") as f:
            header = f.read(len(_OLE_COMPOUND_FILE_SIGNATURE))
    except OSError:
        return None
    if header != _OLE_COMPOUND_FILE_SIGNATURE:
        return None

    if password is None:
        raise EncryptedDocumentError(_ENCRYPTED_DOCUMENT_MESSAGE)

    with open(path, "rb") as f:
        return _decrypt_to_stream(f, password)


def _maybe_decrypt_stream(stream: IO[bytes], password: str | None) -> io.BytesIO | None:
    """Return decrypted-payload BytesIO if `stream` is encrypted, else None.

    Raises :class:`EncryptedDocumentError` when the stream is encrypted and
    either `password` is None or ``python-ooxml-crypto`` is not installed.
    The stream's position is restored before decryption proceeds so the
    decryptor sees the full container bytes.
    """
    if not hasattr(stream, "read"):
        return None
    try:
        pos = stream.tell()
    except (OSError, AttributeError):
        # Not seekable — we can't safely peek without consuming bytes.
        return None
    try:
        header = stream.read(len(_OLE_COMPOUND_FILE_SIGNATURE))
    finally:
        try:
            stream.seek(pos)
        except (OSError, AttributeError):
            pass
    if header != _OLE_COMPOUND_FILE_SIGNATURE:
        return None

    if password is None:
        raise EncryptedDocumentError(_ENCRYPTED_DOCUMENT_MESSAGE)

    return _decrypt_to_stream(stream, password)


class PhysPkgReader:
    """Factory for physical package reader objects.

    Chooses the concrete reader matching `pkg_file` (directory vs zip). When
    the package turns out to be Strict OOXML (detected by sniffing
    ``[Content_Types].xml`` or ``/word/document.xml`` for the Strict
    namespace sentinel), the concrete reader is wrapped in
    :class:`_StrictTranslatingPkgReader` so blobs are rewritten to
    Transitional as they're read. Closes upstream#1520, upstream#693.

    When `password` is provided and the input is an encrypted OOXML package
    (CFBF / OLE2 container), the file is decrypted with `password` via the
    optional ``python-ooxml-crypto`` dependency and the resulting plaintext
    bytes are read as a zip package.

    .. versionchanged:: 2026.05.0
       Transparent Strict → Transitional translation on open.
    .. versionchanged:: 2026.05.10
       Added ``password`` parameter for opening ECMA-376 Agile-Encryption
       password-protected packages.
    """

    def __new__(cls, pkg_file, password: str | None = None):
        # -- if `pkg_file` is a string, treat it as a path --
        if isinstance(pkg_file, str):
            if os.path.isdir(pkg_file):
                reader_cls = _DirPkgReader
                return super(PhysPkgReader, cls).__new__(reader_cls)
            if is_zipfile(pkg_file):
                reader_cls = _ZipPkgReader
                return super(PhysPkgReader, cls).__new__(reader_cls)

            # -- detect password-encrypted .docx (OLE compound file). When a
            # -- password was supplied we decrypt here and stash the plain-zip
            # -- BytesIO on the class for `_ZipPkgReader.__init__` to pick up
            # -- (Python auto-calls `__init__` with the original `pkg_file`
            # -- arg after `__new__`, so we can't just pass the decrypted
            # -- stream through — the override lets us swap it in cleanly).
            decrypted = _maybe_decrypt_path(pkg_file, password)
            if decrypted is not None:
                PhysPkgReader._decrypted_stream_override = decrypted
                return super(PhysPkgReader, cls).__new__(_ZipPkgReader)

            # -- distinguish missing file from wrong-format file
            # -- (closes upstream#1410). Both subclass PackageNotFoundError for
            # -- backward-compatibility. --
            if not os.path.exists(pkg_file):
                raise MissingDocxFileError(
                    "Package not found at '%s'" % pkg_file
                )
            raise NotADocxError(
                "File '%s' is not a valid Word (.docx) package "
                "(not a ZIP archive)." % pkg_file
            )

        # -- assume it's a stream and pass it to Zip reader to sort out --
        decrypted = _maybe_decrypt_stream(pkg_file, password)
        if decrypted is not None:
            cls._decrypted_stream_override = decrypted
            return super(PhysPkgReader, cls).__new__(_ZipPkgReader)

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


def open_phys_pkg_reader(pkg_file, password: str | None = None):
    """Return a physical package reader for `pkg_file` with Strict translation.

    Wraps the concrete :class:`PhysPkgReader` subclass in
    :class:`_StrictTranslatingPkgReader` when the package is Strict OOXML.
    Called from :class:`docx.opc.pkgreader.PackageReader.from_file` in place
    of direct construction.

    When `password` is provided and `pkg_file` is an ECMA-376 Agile-Encryption
    container (CFBF / OLE2), it is transparently decrypted via the optional
    ``python-ooxml-crypto`` dependency before being read.

    .. versionadded:: 2026.05.0
    .. versionchanged:: 2026.05.10
       Added ``password`` parameter.
    """
    reader = PhysPkgReader(pkg_file, password=password)
    if _looks_like_strict_package(reader):
        return _StrictTranslatingPkgReader(reader)
    return reader


class _StrictTranslatingPkgReader:
    """Wraps a physical reader and rewrites Strict URIs as blobs are read.

    Forwards every PhysPkgReader method to the wrapped reader, applying
    :func:`docx.opc.strict.translate_strict_blob` to each returned blob. The
    wrapped reader retains sole ownership of the underlying zip handle /
    directory, so ``close()`` still delegates.

    .. versionadded:: 2026.05.0
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
    """Factory for physical package writer objects.

    When `reproducible` is True, the returned writer emits a deterministic zip
    archive — fixed timestamps, sorted member names, and normalized external
    attributes — so repeated saves of the same content produce byte-identical
    output. See upstream#1042 / upstream-PR#810.

    .. versionadded:: 2026.05.0
       The `reproducible` parameter.
    """

    def __new__(cls, pkg_file, reproducible: bool = False):
        writer_cls = _ReproducibleZipPkgWriter if reproducible else _ZipPkgWriter
        return super(PhysPkgWriter, cls).__new__(writer_cls)


class _DirPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for an OPC package extracted into a
    directory."""

    def __init__(self, path, password: str | None = None):
        """`path` is the path to a directory containing an expanded package.

        `password` is accepted for uniformity with :class:`_ZipPkgReader` but
        is ignored: a directory-extracted package is already plaintext.
        """
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

    def __init__(self, pkg_file, password: str | None = None):
        # -- `password` is consumed by `PhysPkgReader.__new__` to dispatch
        # -- decryption; it is accepted here only so Python's default
        # -- ``type(...).__init__`` handling works with the selector kwarg. --
        super().__init__()
        # -- When `PhysPkgReader.__new__` decrypted an encrypted package, it
        # -- stashed the plain-zip BytesIO on the `PhysPkgReader` class.
        # -- Honour that override so we open the decrypted payload rather
        # -- than the original (encrypted) CFBF bytes. The override is
        # -- single-use; consume it here whether or not we were routed by
        # -- the decryption path to keep the flag from leaking across
        # -- unrelated reader instantiations. --
        override = getattr(PhysPkgReader, "_decrypted_stream_override", None)
        if override is not None:
            try:
                del PhysPkgReader._decrypted_stream_override
            except AttributeError:  # pragma: no cover -- defensive
                pass
            pkg_file = override
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

    def __init__(self, pkg_file, reproducible: bool = False):
        # -- `reproducible` is consumed by the `PhysPkgWriter.__new__` dispatch to
        # -- select `_ReproducibleZipPkgWriter`; accepted here only so Python's
        # -- default ``type(...).__init__`` handling works with the selector kwarg. --
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


class _ReproducibleZipPkgWriter(_ZipPkgWriter):
    """Zip writer that produces byte-identical output for identical inputs.

    Buffers every ``(pack_uri, blob)`` pair and flushes them to the underlying
    archive in sorted membername order at ``close()``. Each member is written
    with a fixed timestamp (:data:`REPRODUCIBLE_TIMESTAMP`) and normalized
    external attributes (0o644, regular file). Closes upstream#1042,
    upstream-PR#810.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, pkg_file, reproducible: bool = True):
        # -- ``reproducible`` is how the factory chose this subclass; accept it
        # -- for signature uniformity. --
        super().__init__(pkg_file)
        self._pending: list[tuple[str, bytes]] = []

    def write(self, pack_uri, blob):
        self._pending.append((pack_uri.membername, blob))

    def close(self):
        for membername, blob in sorted(self._pending, key=lambda item: item[0]):
            # -- ZipInfo normalises header fields; use it to fix timestamps and
            # -- external_attr so subsequent saves produce identical bytes. --
            info = ZipInfo(filename=membername, date_time=REPRODUCIBLE_TIMESTAMP)
            info.compress_type = ZIP_DEFLATED
            info.external_attr = (0o644 & 0xFFFF) << 16  # -- rw-r--r--, regular file --
            info.create_system = 3  # -- Unix; stable across platforms --
            self._zipf.writestr(info, blob)
        self._pending = []
        super().close()
