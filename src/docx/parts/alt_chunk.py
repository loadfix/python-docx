"""|AltChunkPart| — container for an alternate-format chunk payload.

An "alternate format import part" carries a payload in a non-WordprocessingML
format (HTML, RTF, plain text, XHTML, Microsoft Word 97-2003 .doc, etc.) that
Word substitutes for the referencing ``<w:altChunk>`` element at render time.
The relationship that ties the main document to the payload uses reltype
``aFChunk`` (see ``RELATIONSHIP_TYPE.A_F_CHUNK``). The target part's
content-type declares the payload format — for example ``text/html`` for an
HTML chunk.

python-docx can write these parts from arbitrary bytes and the part reads
transparently from an on-disk package.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.part import Part

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.opc.packuri import PackURI


class AltChunkPart(Part):
    """An alternate-format import chunk part.

    Corresponds to the target part of an ``aFChunk`` relationship. The
    contents are raw bytes whose interpretation is governed by the part's
    content-type (``text/html``, ``application/rtf``, ``text/plain``,
    ``application/xhtml+xml``, ``application/msword``, etc.).

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        blob: bytes,
        package: OpcPackage | None = None,
    ):
        super().__init__(partname, content_type, blob, package)

    @classmethod
    def load(
        cls,
        partname: PackURI,
        content_type: str,
        blob: bytes,
        package: OpcPackage,
    ):
        """Called by :class:`docx.opc.package.PartFactory` during package load."""
        return cls(partname, content_type, blob, package)

    @classmethod
    def new(
        cls,
        package: OpcPackage,
        content: bytes,
        content_type: str,
        partname: PackURI | None = None,
    ):
        """Return a newly-constructed |AltChunkPart| containing `content`.

        When `partname` is |None| a fresh partname under
        ``/word/afchunk%d.EXT`` is chosen by picking the lowest positive
        integer not already in use in the package. The extension is
        inferred from `content_type` (``text/html`` → ``.html``,
        ``application/rtf`` → ``.rtf``, ``text/plain`` → ``.txt``,
        ``application/xhtml+xml`` → ``.xhtml``, ``application/msword`` →
        ``.doc``, ``message/rfc822`` / ``multipart/related`` → ``.mht``,
        the wordprocessingml document main+xml type → ``.docx``); anything
        else falls back to ``.bin``.
        """
        from docx.opc.packuri import PackURI

        if partname is None:
            ext = _ext_for_content_type(content_type)
            template = "/word/afchunk%%d%s" % ext
            partname = package.next_partname(template)
        elif not isinstance(partname, PackURI):
            partname = PackURI(str(partname))
        return cls(partname, content_type, content, package)


# Recognised altChunk payload content-types. Word's import filters dispatch
# on the target part's content-type, so we accept any string but prefer the
# types below for extension inference. See ECMA-376 §17.17 for the list.
CT_XHTML = "application/xhtml+xml"
CT_HTML = "text/html"
CT_RTF = "application/rtf"
CT_RTF_ALT = "text/rtf"
CT_PLAIN = "text/plain"
CT_MSWORD = "application/msword"
CT_MHTML = "message/rfc822"
CT_MHTML_ALT = "multipart/related"
CT_DOCX_FRAGMENT = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.document.main+xml"
)


def _ext_for_content_type(content_type: str) -> str:
    """Return the conventional file extension for `content_type`.

    Covers every content-type Word recognises for an ``altChunk`` payload:
    HTML, XHTML, RTF (both MIME spellings), plain text, legacy ``.doc``,
    MHTML, and a wordprocessingml-document fragment part. Unknown types
    fall back to ``.bin``.
    """
    mapping = {
        CT_HTML: ".html",
        CT_XHTML: ".xhtml",
        CT_RTF: ".rtf",
        CT_RTF_ALT: ".rtf",
        CT_PLAIN: ".txt",
        CT_MSWORD: ".doc",
        CT_MHTML: ".mht",
        CT_MHTML_ALT: ".mht",
        CT_DOCX_FRAGMENT: ".docx",
    }
    return mapping.get(content_type.lower(), ".bin")
