"""Exceptions specific to python-opc.

The base exception class is OpcError.
"""

from __future__ import annotations


class OpcError(Exception):
    """Base error class for python-opc."""


class PackageNotFoundError(OpcError):
    """Raised when a package cannot be found at the specified path."""


class NotADocxError(PackageNotFoundError):
    """Raised when the target path exists but is not a valid OPC package (.docx).

    Subclass of :class:`PackageNotFoundError` for backward compatibility — existing
    callers that catch ``PackageNotFoundError`` will continue to receive the same
    behavior, while new callers can distinguish a missing file
    (:class:`MissingDocxFileError`) from a wrongly-formatted file
    (:class:`NotADocxError`). Closes upstream#1410.

    .. versionadded:: 2026.05.0
    """


class MissingDocxFileError(PackageNotFoundError, FileNotFoundError):
    """Raised when no file exists at the requested path.

    Inherits from both :class:`PackageNotFoundError` (backward compatibility —
    existing ``except PackageNotFoundError`` blocks keep working) and
    :class:`FileNotFoundError` (so the error also behaves like a standard
    ``FileNotFoundError`` for callers that want to handle it uniformly with
    other filesystem missing-file errors). Closes upstream#1410.

    .. versionadded:: 2026.05.0
    """
