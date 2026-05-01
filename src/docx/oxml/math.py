"""Custom element classes for Office Math Markup Language (OMML).

OMML is Word's native equation format, stored inline using the ``m:`` namespace
(``http://schemas.openxmlformats.org/officeDocument/2006/math``), either as
``m:oMath`` (inline) or ``m:oMathPara`` (display-mode paragraph).

Only the minimal set of elements needed for read + simple programmatic create is
modelled here. Full OMML is large; callers that need richer math can construct
OMML XML externally and pass it to :meth:`Equation.from_omml_xml`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    pass


class CT_OMath(BaseOxmlElement):
    """``<m:oMath>`` element, an inline Office Math expression."""

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Concatenated text of all descendant ``<m:t>`` elements.

        Best-effort plain-text rendering — structural elements (fractions,
        superscripts, radicals, …) are flattened to the concatenation of their
        text leaves, which is usually good enough for search/display.
        """
        return "".join(t.text or "" for t in self.xpath(".//m:t"))


class CT_OMathPara(BaseOxmlElement):
    """``<m:oMathPara>`` element, wrapping one or more ``m:oMath`` in display mode."""

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Concatenated text of all descendant ``<m:t>`` elements."""
        return "".join(t.text or "" for t in self.xpath(".//m:t"))


class CT_MathR(BaseOxmlElement):
    """``<m:r>`` element — a run inside an OMML expression."""


class CT_MathT(BaseOxmlElement):
    """``<m:t>`` element — the text leaf of an OMML run."""
