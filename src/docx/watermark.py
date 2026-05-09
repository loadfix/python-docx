"""The |Watermark| proxy class.

Provides a small read-side API for a VML watermark stored in a section's
default page header.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.oxml.watermark import CT_VmlShape


class Watermark:
    """Proxy for a VML watermark shape (``v:shape``) residing in a header.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, shape: "CT_VmlShape"):
        self._shape = shape

    @property
    def type(self) -> str:
        """``"text"`` or ``"image"``.

        Returns ``"image"`` when an ``<v:imagedata>`` child is present, otherwise
        ``"text"``.

        .. versionadded:: 2026.05.0
        """
        if self._shape.find(qn("v:imagedata")) is not None:
            return "image"
        return "text"

    @property
    def text(self) -> str | None:
        """Text string of a text watermark, or ``None`` for an image watermark.

        .. versionadded:: 2026.05.0
        """
        textpath = self._shape.find(qn("v:textpath"))
        if textpath is None:
            return None
        return textpath.get("string")

    def remove(self) -> None:
        """Remove this watermark's containing paragraph from its header.

        Walks up from the ``v:shape`` through ``w:pict`` / ``w:r`` to the
        enclosing ``w:p`` and detaches that paragraph. Safe to call when
        the shape has already been detached — becomes a no-op.

        .. versionadded:: 2026.05.0
        """
        # -- walk up to the ancestor w:p (paragraph) --
        node = self._shape.getparent()
        target_p = None
        p_tag = qn("w:p")
        while node is not None:
            if node.tag == p_tag:
                target_p = node
                break
            node = node.getparent()
        if target_p is None:
            # -- no enclosing paragraph; detach the shape itself --
            parent = self._shape.getparent()
            if parent is not None:
                parent.remove(self._shape)
            return
        p_parent = target_p.getparent()
        if p_parent is not None:
            p_parent.remove(target_p)
