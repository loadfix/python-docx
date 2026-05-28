"""Exceptions used with python-docx.

The base exception class is PythonDocxError.
"""


class PythonDocxError(Exception):
    """Generic error class."""


class InvalidSpanError(PythonDocxError):
    """Raised when an invalid merge region is specified in a request to merge table
    cells."""


class InvalidXmlError(PythonDocxError):
    """Raised when invalid XML is encountered, such as on attempt to access a missing
    required child element."""


class EncryptedDocumentError(PythonDocxError):
    """Raised when attempting to open a password-encrypted .docx file.

    Word stores encrypted documents as OLE compound files (CFBF) containing the
    encrypted package, which cannot be opened by the standard zipfile reader.
    Detection is performed by checking the file's magic bytes against the OLE
    compound file signature ``D0 CF 11 E0 A1 B1 1A E1``.

    Also raised when the optional ``python-ooxml-crypto`` dependency is required
    to decrypt or encrypt a package but is not installed, when the supplied
    password does not match the one used to encrypt the package, or when the
    underlying encryption container is malformed.
    """


class NestedSectionError(PythonDocxError):
    """Raised when entering a section context inside another active one.

    The OOXML model encodes sections by attaching a ``w:sectPr`` to the
    last paragraph of a region. Sections cannot nest — every paragraph
    belongs to exactly one section. :meth:`docx.Document.section`
    surfaces this constraint at the API layer.

    .. versionadded:: 2026.05.13
    """


class RmsProtectedDocumentError(EncryptedDocumentError):
    """Raised when opening a .docx wrapped in Azure RMS / AIP / IRM protection.

    "Rights Management Services" (also marketed as Azure Information Protection /
    Microsoft Purview Information Protection / "Information Rights Management")
    wraps the regular OOXML zip inside a CFBF (OLE2 compound file) container
    that stores the encrypted payload under a ``DRMContent`` stream and a
    ``DRMEncryptedTransform`` descriptor. Unlike an ECMA-376 Agile-Encryption
    package, an RMS package cannot be decrypted with a password alone — the
    user's Azure AD / Microsoft 365 identity must be presented to the RMS
    service to retrieve the content key.

    python-docx does not bundle an RMS client (the Microsoft Information
    Protection SDK is C#/.NET-only and requires an interactive Azure AD login
    flow). Callers that need RMS decryption should delegate to Microsoft Office
    automation, the MIP SDK, or a pre-processing step before opening the file
    with python-docx.

    .. versionadded:: 2026.05.10
    """
