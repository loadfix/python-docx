"""Custom element classes for Office Math Markup Language (OMML).

OMML is Word's native equation format, stored inline using the ``m:`` namespace
(``http://schemas.openxmlformats.org/officeDocument/2006/math``), either as
``m:oMath`` (inline) or ``m:oMathPara`` (display-mode paragraph).

The element classes themselves live in the shared ``ooxml_math`` package —
this module is a thin re-export kept in place so that legacy imports
(``from docx.oxml.math import CT_OMath``) continue to work unchanged. The
shared package was extracted to give ``python-pptx`` equal access to OMML
authoring without copy-pasting the classes, and to make future OMML growth
(the operator tree, accents, fractions, radicals, ...) a single code change
rather than a per-parent fork.

Only the minimal set of elements needed for read + simple programmatic
create is modelled. Full OMML is large; callers that need richer math can
construct OMML XML externally and pass it to :meth:`Equation.from_omml_xml`.
"""

from __future__ import annotations

from ooxml_math.oxml.elements import (
    CT_MathR,
    CT_MathT,
    CT_OMath,
    CT_OMathPara,
)

__all__ = [
    "CT_MathR",
    "CT_MathT",
    "CT_OMath",
    "CT_OMathPara",
]
