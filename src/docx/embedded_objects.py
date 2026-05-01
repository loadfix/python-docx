"""Read-only proxy objects for embedded OLE objects.

Word supports embedding OLE objects (such as Excel workbooks, PDF files, or
equations) inside a document. They are stored as separate parts, usually under
``word/embeddings/``, with content type
``application/vnd.openxmlformats-officedocument.oleObject``. The relationship
is carried by an ``<o:OLEObject>`` element nested inside a ``<w:object>``
element inside a run.

python-docx exposes these embedded objects read-only. Callers can enumerate
them at the |Document| or |Paragraph| level, inspect the ProgID and
link/embed type, discover the related part, or retrieve the raw OLE bytes
for further processing. Creation and modification are intentionally not
supported.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.embedded_object import EmbeddedObjectPart
    from docx.text.paragraph import Paragraph


class EmbeddedObject:
    """Proxy for an OLE object referenced from a paragraph.

    Wraps an ``<o:OLEObject>`` element (via the paragraph that contains it)
    and optionally its related embedded-object part. Read-only.
    """

    def __init__(
        self,
        paragraph: Paragraph,
        ole_object_elm: BaseOxmlElement,
        embedded_part: EmbeddedObjectPart | None,
    ):
        self._paragraph = paragraph
        self._ole_object_elm = ole_object_elm
        self._embedded_part = embedded_part

    @property
    def blob(self) -> bytes:
        """Raw bytes of the embedded OLE object.

        Returns an empty ``bytes`` object when the relationship cannot be
        resolved to an embedded-object part (for example when the target is
        missing from the package or is of an unexpected part type).
        """
        if self._embedded_part is None:
            return b""
        return self._embedded_part.blob

    @property
    def embedded_partname(self) -> str | None:
        """The OPC partname of the related embedded-object part, or |None|.

        Returns |None| when the relationship cannot be resolved.
        """
        if self._embedded_part is None:
            return None
        return str(self._embedded_part.partname)

    @property
    def paragraph(self) -> Paragraph:
        """The |Paragraph| that contains the ``w:object`` for this embedded object."""
        return self._paragraph

    @property
    def prog_id(self) -> str | None:
        """The ProgID identifying the object's type, e.g. ``"Excel.Sheet.12"``.

        Returns |None| when the ``o:OLEObject`` element has no ``ProgID``
        attribute.
        """
        value = self._ole_object_elm.get("ProgID")
        return value if value else None

    @property
    def r_id(self) -> str | None:
        """The relationship id of the embedded OLE binary, e.g. ``"rId5"``.

        Returns |None| when the ``o:OLEObject`` element has no ``r:id``
        attribute.
        """
        from docx.oxml.ns import qn

        value = self._ole_object_elm.get(qn("r:id"))
        return value if value else None

    @property
    def type(self) -> str | None:
        """The link/embed type of the object.

        Usually ``"Embed"`` for embedded objects and ``"Link"`` for linked
        objects. Returns |None| when the ``Type`` attribute is absent.
        """
        value = self._ole_object_elm.get("Type")
        return value if value else None
