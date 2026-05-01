"""|EmbeddedObjectPart| — container for an embedded OLE object payload.

Embedded objects (OLE objects such as Excel workbooks, PDF files, equations,
etc.) are stored as separate parts (typically under ``word/embeddings/``) with
content type ``application/vnd.openxmlformats-officedocument.oleObject``.
They are referenced from an ``o:OLEObject`` element inside a ``w:object``
element inside a run. python-docx exposes them read-only.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.part import Part

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.opc.packuri import PackURI


class EmbeddedObjectPart(Part):
    """An embedded OLE object part.

    Corresponds to the target part of a relationship whose ``content-type`` is
    ``application/vnd.openxmlformats-officedocument.oleObject``. The contents
    are the raw binary OLE bytes; python-docx does not parse or interpret
    them beyond exposing them as a blob.
    """

    def __init__(self, partname: PackURI, content_type: str, blob: bytes):
        super().__init__(partname, content_type, blob)

    @classmethod
    def load(cls, partname: PackURI, content_type: str, blob: bytes, package: OpcPackage):
        """Called by ``docx.opc.package.PartFactory`` when loading an embedded object part."""
        return cls(partname, content_type, blob)
