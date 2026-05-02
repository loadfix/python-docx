.. _document_safety:

Document safety: corruption, encryption, macros, signatures
===========================================================

Beyond "happy-path" reads, |docx| provides a handful of APIs that surface
the *safety* attributes of a document: whether its XML parts survived
parsing, whether it is password-encrypted, whether it carries VBA macros,
and whether it bears a digital signature. These matter when a tool has to
decide whether to load, process, forward, or reject a document it received
from somewhere else.

None of these features require additional dependencies. |docx| does not
decrypt, execute VBA, or cryptographically verify signatures — it only
inspects what the package contains.


Recover mode for malformed documents
------------------------------------

When a ``.docx`` has been truncated, had an editor partially rewrite its
XML, or otherwise lost well-formedness, the default
:func:`docx.Document` loader raises :class:`lxml.etree.XMLSyntaxError`.
Passing ``recover=True`` switches lxml into its recovering parser, which
salvages whatever is well-formed and records the parse errors on
:attr:`.Document.recovery_warnings`::

    >>> from docx import Document
    >>> from lxml import etree

    >>> try:
    ...     document = Document("corrupt.docx")
    ... except etree.XMLSyntaxError as e:
    ...     print(f"default open failed: {e}")

    >>> document = Document("corrupt.docx", recover=True)
    >>> len(document.recovery_warnings)
    1
    >>> document.recovery_warnings[0]
    '<string>:10:24:FATAL:PARSER:ERR_TAG_NOT_FINISHED: ...'

The readable prefix of the document is available through the normal API.
Content after the corruption boundary is dropped; in extreme cases where
lxml cannot recover *any* elements from a part, |docx| substitutes an
empty stub for that part so the rest of the package still loads.

Recover mode never masks unrelated failures. If the physical package is
not a zip file, :class:`docx.opc.exceptions.PackageNotFoundError` still
propagates; if the file is an encrypted OLE compound file,
:class:`docx.exceptions.EncryptedDocumentError` still propagates. The
``recover=True`` flag only relaxes XML parsing.


Detecting password-encrypted documents
--------------------------------------

Word stores password-protected documents as OLE compound files (CFBF), not
as regular ZIP packages. The ZIP-based OPC reader cannot process them; the
naive error would be a confusing ``BadZipFile`` from the standard library.

|docx| short-circuits that by peeking at the first eight bytes of the file
and raising :class:`.EncryptedDocumentError` if they match the OLE signature
``D0 CF 11 E0 A1 B1 1A E1``::

    >>> from docx import Document
    >>> from docx.exceptions import EncryptedDocumentError
    >>> try:
    ...     Document("secret.docx")
    ... except EncryptedDocumentError as e:
    ...     print(e)
    Document is password-protected (encrypted .docx detected). Install
    msoffcrypto-tool to decrypt it first: https://github.com/nolze/msoffcrypto-tool

Recover mode does **not** bypass this check — the file is not just
malformed XML, it is an entirely different format.

To decrypt the document before loading it with |docx|, use
`msoffcrypto-tool <https://github.com/nolze/msoffcrypto-tool>`_::

    import io
    import msoffcrypto

    with open("secret.docx", "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password="s3cret")
        plain = io.BytesIO()
        office.decrypt(plain)
        plain.seek(0)
        document = Document(plain)


Macro-enabled documents (.docm)
-------------------------------

``.docm`` documents are OOXML packages whose main document part uses the
macro-enabled content type
(``application/vnd.ms-word.document.macroEnabled.main+xml``) and carry a
``vbaProject`` relationship pointing at a binary ``vbaProject.bin`` part.

|docx| loads them seamlessly — no special flag is required — and surfaces
the VBA relationship through :attr:`.Document.has_macros`::

    >>> document = Document("form.docm")
    >>> document.has_macros
    True
    >>> Document("plain.docx").has_macros
    False

|docx| does not read or author VBA. The ``vbaProject.bin`` part is left
untouched on save; if you inspect or swap VBA code, use a dedicated tool
and then pass the resulting bytes back to |docx|.

.. note::

   VBA projects are an execution vector. Treat a positive
   :attr:`has_macros` result as a security signal unless the document
   came from a trusted source.


Digital signatures
------------------

A signed OOXML package includes:

- A package-level relationship of type
  ``.../digital-signature/origin`` targeting
  ``/_xmlsignatures/origin.sigs``;
- One or more ``digital-signature/signature`` relationships from the origin
  part, each targeting a ``/_xmlsignatures/sigN.xml`` part holding an
  XML-DSig document (optionally with XAdES extensions carrying the signing
  time and signer identity).

|docx| surfaces both the presence and the minimal metadata::

    >>> document = Document("contract.docx")
    >>> document.is_signed
    True
    >>> for sig in document.signatures:
    ...     print(sig.partname, sig.signer, sig.signed_at)
    /_xmlsignatures/sig1.xml CN=Alice Example 2024-04-01 12:34:56+00:00

Each :class:`.SignatureInfo` exposes :attr:`partname`, :attr:`blob`
(the raw XML bytes), :attr:`signer` (the ``X509SubjectName``), and
:attr:`signed_at` (the XAdES ``SigningTime``, or |None| when absent). The
full signature XML is available through :attr:`blob` for callers that want
to perform their own cryptographic verification.

|docx| does not verify signatures — signature validation is a
cryptographic operation outside |docx|'s scope. Consumers that rely on
signed documents should pass the :attr:`blob` to a library such as
``signxml`` and check the result before proceeding.
