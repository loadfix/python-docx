"""|InkPart| — container for an ink annotation XML payload.

Ink annotations are stored as separate parts (``word/ink/ink*.xml``) with
content type ``application/inkml+xml``. They are referenced from a
``w:contentPart`` element inside a run. python-docx exposes them read-only.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.part import Part

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.opc.packuri import PackURI


class InkPart(Part):
    """An ink annotation (InkML) part.

    Corresponds to the target part of a relationship whose ``content-type`` is
    ``application/inkml+xml``. The contents are the raw InkML XML bytes;
    python-docx does not parse or interpret strokes beyond counting them.
    """

    def __init__(self, partname: PackURI, content_type: str, blob: bytes):
        super().__init__(partname, content_type, blob)

    @classmethod
    def load(cls, partname: PackURI, content_type: str, blob: bytes, package: OpcPackage):
        """Called by ``docx.opc.package.PartFactory`` when loading an ink part."""
        return cls(partname, content_type, blob)
