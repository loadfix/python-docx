"""Alternate-format chunk (``altChunk``) proxy objects.

Word's *alternate-format chunk* mechanism lets a document pull content from an
external-format payload (HTML, RTF, plain text, XHTML, legacy ``.doc``, ...)
at render time. The referencing marker is a ``<w:altChunk r:id="..."/>``
element embedded in the document body; the payload lives in a separate part
wired via an ``aFChunk`` relationship. Word substitutes the payload for the
``altChunk`` element when it opens the document.

python-docx does not *evaluate* altChunks (no HTML/RTF rendering engine); it
only writes and reads the marker + part structure so callers can round-trip
documents that use them. Closes upstream#1317, upstream#1103, PR#649.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.constants import RELATIONSHIP_TYPE as RT

if TYPE_CHECKING:
    from docx.oxml.document import CT_AltChunk
    from docx.parts.alt_chunk import AltChunkPart
    from docx.parts.document import DocumentPart


class AltChunk:
    """Proxy for a single ``<w:altChunk>`` import reference.

    Wraps the ``w:altChunk`` element and its related |AltChunkPart|. The
    underlying payload is exposed read-only via :attr:`content` and
    :attr:`content_type`.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, element: CT_AltChunk, document_part: DocumentPart):
        self._element = element
        self._document_part = document_part

    @property
    def rId(self) -> str | None:
        """The ``r:id`` attribute value pointing at the payload part.

        Returns |None| for a malformed ``w:altChunk`` that carries no
        ``r:id`` attribute (Word will refuse to open such a document).

        .. versionadded:: 2026.05.0
        """
        return self._element.rId

    @property
    def part(self) -> AltChunkPart | None:
        """The related |AltChunkPart| or |None| if it cannot be resolved.

        .. versionadded:: 2026.05.0
        """
        rId = self.rId
        if rId is None:
            return None
        related = self._document_part.related_parts
        try:
            target = related[rId]
        except KeyError:
            return None
        from docx.parts.alt_chunk import AltChunkPart

        if not isinstance(target, AltChunkPart):
            return None
        return target

    @property
    def content_type(self) -> str | None:
        """Content-type of the related payload part, or |None|.

        .. versionadded:: 2026.05.0
        """
        part = self.part
        if part is None:
            return None
        return part.content_type

    @property
    def content(self) -> bytes:
        """Raw bytes of the related payload part, or empty bytes.

        .. versionadded:: 2026.05.0
        """
        part = self.part
        if part is None:
            return b""
        return part.blob

    @property
    def match_src(self) -> bool | None:
        """Value of the ``w:altChunkPr/w:matchSrc`` child, or |None|.

        Word interprets this flag as "try to match the source formatting of
        the imported payload" (see ECMA-376 §17.17.2.3). Returns |None| when
        the ``w:altChunk`` carries no ``w:altChunkPr/w:matchSrc`` child.

        .. versionadded:: 2026.05.0
        """
        return self._element.match_src

    @match_src.setter
    def match_src(self, value: bool | None) -> None:
        self._element.match_src = value


def iter_alt_chunks(document_part: DocumentPart) -> list[AltChunk]:
    """Return a list of |AltChunk| proxies for each ``w:altChunk`` in the body.

    Elements appear in document order. Only direct children of ``w:body``
    are inspected — nested altChunks (within ``w:sdt`` wrappers, etc.) are
    not supported by Word in the body-level context and are ignored.

    .. versionadded:: 2026.05.0
    """
    body = document_part.element.body  # type: ignore[attr-defined]
    return [AltChunk(el, document_part) for el in body.altChunk_lst]


def add_alt_chunk_to_document(
    document_part: DocumentPart,
    content: bytes | str,
    content_type: str = "text/html",
    match_src: bool | None = None,
) -> AltChunk:
    """Append a new ``w:altChunk`` to the document body and return a proxy.

    Creates a new |AltChunkPart| carrying `content` and relates the
    document part to it via an ``aFChunk`` relationship, then appends a
    ``w:altChunk`` element at the body level that references the new
    relationship by id.

    `content` may be :class:`bytes` or :class:`str` (strings are encoded
    as UTF-8). `content_type` is the MIME type Word uses to dispatch the
    payload through the right import filter (``text/html`` by default).
    Pass `match_src=True` to write a ``w:altChunkPr/w:matchSrc`` child
    requesting Word match the source formatting of the imported payload.

    .. versionadded:: 2026.05.0
    """
    from docx.parts.alt_chunk import AltChunkPart

    if isinstance(content, str):
        content_bytes = content.encode("utf-8")
    else:
        content_bytes = content

    package = document_part.package
    assert package is not None
    part = AltChunkPart.new(package, content_bytes, content_type)
    rId = document_part.relate_to(part, RT.A_F_CHUNK)
    body = document_part.element.body  # type: ignore[attr-defined]
    element = body.add_altChunk(rId)
    if match_src is not None:
        element.match_src = match_src
    return AltChunk(element, document_part)
