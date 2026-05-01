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
    """
