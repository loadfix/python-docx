"""Read-only proxy objects for ``w:altChunk`` attachments.

An ``altChunk`` is a Word-specific mechanism for embedding a foreign payload
(HTML, RTF, another docx, plain text, etc.) inline in a document. Word merges
the payload into the rendered output on open but the raw payload remains in
the package as a separate part referenced by an ``r:id`` relationship.

python-docx exposes altChunks read-only via :attr:`Document.attachments`.

.. versionadded:: 1.3.0.dev0
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.opc.part import Part
    from docx.oxml.xmlchemy import BaseOxmlElement


class Attachment:
    """Proxy for a ``w:altChunk`` element and its related part.

    Read-only. Creation and modification are intentionally not supported; use
    the higher-level document-insert APIs (or write the altChunk XML by hand)
    when authoring.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(
        self,
        alt_chunk_elm: BaseOxmlElement,
        target_part: Part | None,
    ):
        self._alt_chunk_elm = alt_chunk_elm
        self._target_part = target_part

    @property
    def r_id(self) -> str | None:
        """Relationship id referenced by ``w:altChunk/@r:id``, or |None|.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.ns import qn

        value = self._alt_chunk_elm.get(qn("r:id"))
        return value if value else None

    @property
    def content_type(self) -> str | None:
        """Content-type of the related part, or |None| when unresolved.

        .. versionadded:: 1.3.0.dev0
        """
        if self._target_part is None:
            return None
        return getattr(self._target_part, "content_type", None)

    @property
    def blob(self) -> bytes:
        """Raw bytes of the altChunk payload (empty when unresolved).

        .. versionadded:: 1.3.0.dev0
        """
        if self._target_part is None:
            return b""
        blob = getattr(self._target_part, "blob", None)
        if isinstance(blob, bytes):
            return blob
        if isinstance(blob, bytearray):
            return bytes(blob)
        return b""

    @property
    def partname(self) -> str | None:
        """OPC partname of the related part, or |None|.

        .. versionadded:: 1.3.0.dev0
        """
        if self._target_part is None:
            return None
        return str(getattr(self._target_part, "partname", "")) or None
