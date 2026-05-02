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
