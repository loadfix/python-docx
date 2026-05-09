"""|VmlDrawingPart| — container for a legacy VML drawing payload.

VML (Vector Markup Language) is Microsoft's pre-DrawingML shape format.
Word still emits VML in three niches:

- **Header / footer watermarks** — ``"Draft"`` / ``"Confidential"``
  diagonal banners are drawn with ``<v:shape>`` embellishments.
- **mc:AlternateContent fallback** — DrawingML-only features carry a
  VML down-level rendering in ``<mc:Fallback>``.
- **Legacy form-control anchors** — any content that predates the
  DrawingML object model.

python-docx 2026.05.11+ delegates VML handling to the byte-preserve
slice shipped by ``python-ooxml-vml`` (``VmlDrawingPart`` /
``LegacyDrawingPart``). Loading a VML part stashes the raw bytes on
an :class:`ooxml_vml.VmlDrawingPart` facade; saving re-emits them
byte-identical. No parsing, no authoring — just round-trip fidelity
for the parts Word writes but docx doesn't author.

.. versionadded:: 2026.05.11
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from ooxml_vml import VmlDrawingPart as _OoxmlVmlDrawingPart

from docx.opc.part import Part

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.opc.packuri import PackURI


class VmlDrawingPart(Part):
    """A legacy VML drawing part.

    Corresponds to the target part of a relationship whose
    ``content-type`` is
    ``application/vnd.openxmlformats-officedocument.vmlDrawing``. The
    payload is held as raw bytes and re-emitted verbatim on save:

        ``VmlDrawingPart.load(...).blob == original_blob``

    Byte-preservation is delegated to
    :class:`ooxml_vml.VmlDrawingPart`; the :attr:`vml_part` property
    exposes that facade so downstream code can query the shared type.
    """

    def __init__(self, partname: PackURI, content_type: str, blob: bytes):
        super().__init__(partname, content_type, blob)
        # -- Lazy until first access; see :attr:`vml_part`.
        self._vml_part: _OoxmlVmlDrawingPart | None = None

    @classmethod
    def load(
        cls,
        partname: PackURI,
        content_type: str,
        blob: bytes,
        package: OpcPackage,
    ):
        """Called by ``docx.opc.package.PartFactory`` at load time."""
        return cls(partname, content_type, blob)

    @property
    def vml_part(self) -> _OoxmlVmlDrawingPart:
        """Return the ``ooxml_vml.VmlDrawingPart`` facade for this payload.

        Constructed lazily on first access so a part that is never
        inspected by downstream code doesn't pay the (tiny) facade
        cost. The facade's ``.blob`` is byte-identical to
        :attr:`Part.blob` by construction of
        :meth:`ooxml_vml.VmlDrawingPart.from_bytes`.
        """
        if self._vml_part is None:
            self._vml_part = _OoxmlVmlDrawingPart.from_bytes(
                self.blob, partname=str(self._partname)
            )
        return self._vml_part


# -- Alias preserving Word's "legacy drawing" nomenclature. Word's
# -- header/footer watermark relationship type is ``legacyDrawing``;
# -- callers that prefer that name get a drop-in subclass that shares
# -- behaviour with :class:`VmlDrawingPart` one-for-one. --
class LegacyDrawingPart(VmlDrawingPart):
    """Alias of :class:`VmlDrawingPart` for Word legacy-drawing callers.

    The Word header/footer watermark relationship is historically
    called ``legacyDrawing``; this alias mirrors the nomenclature
    used by :mod:`ooxml_vml` without forcing a rename on callers that
    already import the VML name.
    """
