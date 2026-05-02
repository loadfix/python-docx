"""Directly exposed API functions and classes, :func:`Document` for now.

Provides a syntactically more convenient API for interacting with the OpcPackage graph.
"""

from __future__ import annotations

from pathlib import Path
from typing import IO, TYPE_CHECKING, cast

from docx.opc.constants import CONTENT_TYPE as CT
from docx.package import Package

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.parts.document import DocumentPart


def Document(
    docx: str | IO[bytes] | None = None,
    recover: bool = False,
    huge_tree: bool = False,
) -> DocumentObject:
    """Return a |Document| object loaded from `docx`, where `docx` can be either a path
    to a ``.docx`` file (a string) or a file-like object.

    If `docx` is missing or ``None``, the built-in default document "template" is
    loaded.

    When `recover` is True, XML parsing falls back to lxml's recovering parser for
    malformed parts (truncated, mismatched tags, invalid characters). Any parse
    warnings are collected on :attr:`Document.recovery_warnings`. Content that
    cannot be recovered is treated as empty. Irrecoverable failures unrelated to
    XML — for example, an invalid zip file or a password-protected document —
    continue to raise (:class:`PackageNotFoundError`, :class:`EncryptedDocumentError`).
    Default behaviour (``recover=False``) is unchanged.

    When `huge_tree` is True, lxml's ``huge_tree=True`` parser variant is used,
    lifting libxml2's default 10 MB-per-AttValue and 256-deep-nesting safety
    limits so extremely large documents can be parsed. Only enable for trusted
    input — the default parser's XML-bomb protections no longer apply. Closes
    upstream#1086.

    .. versionadded:: 1.3.0.dev0
       The `huge_tree` parameter.
    """
    docx = _default_docx_path() if docx is None else docx
    package = Package.open(docx, recover=recover, huge_tree=huge_tree)
    document_part = cast("DocumentPart", package.main_document_part)
    if document_part.content_type not in (CT.WML_DOCUMENT_MAIN, CT.WML_DOCUMENT_MACRO):
        raise ValueError(
            f"file '{docx}' is not a Word file, content type is '{document_part.content_type}'"
        )
    return document_part.document


def _default_docx_path() -> str:
    """Return the path to the built-in default .docx package."""
    return str(Path(__file__).parent / "templates" / "default.docx")
