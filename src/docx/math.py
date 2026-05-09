"""Pythonic proxy layer for Office Math (OMML) expressions.

Re-exports the :mod:`ooxml_math.proxies` surface under ``docx.math`` so
docx users can construct and inspect equations without touching the raw
``CT_*`` layer::

    from docx import Document
    from docx.math import Fraction, Lit, Var, oMath

    document = Document()
    p = document.add_paragraph("Pythagoras: ")

    expr = oMath(Fraction(Var("x"), Lit(2)))  # x / 2
    p.add_equation(expr)

The proxy layer is a superset of the legacy :func:`docx.equations.build_*`
helpers — the :mod:`docx.equations` functions still work and are
retained for back-compatibility.

Public names (all re-exported verbatim from :mod:`ooxml_math`):

- Abstract base: :class:`MathExpr`, :data:`MathExprLike`,
  :class:`ElementProxy`.
- Leaves: :class:`Var`, :class:`Lit`, :class:`Text`, :class:`Raw`.
- Operator tree: :class:`Fraction`, :class:`Radical`, :class:`Sub`,
  :class:`Sup`, :class:`SubSup`, :class:`Pre`, :class:`Sum`,
  :class:`Product`, :class:`Integral`, :class:`Nary`, :class:`Limit`,
  :class:`FuncApply`, :class:`Delimiter`, :class:`Matrix`,
  :class:`Accent`, :class:`Bar`, :class:`Box`, :class:`BorderBox`,
  :class:`Phantom`, :class:`GroupChar`, :class:`EqArray`.
- Root container: :class:`oMath`.
- Parse dispatch: :func:`from_element`.

.. versionadded:: 2026.05.12
.. versionchanged:: 2026.05.10
   Added the six deferred 0.4.0 proxies :class:`Bar`, :class:`Box`,
   :class:`BorderBox`, :class:`Phantom`, :class:`GroupChar` and
   :class:`EqArray` (tracked by ``python-ooxml-math`` 0.4.0).
"""

from __future__ import annotations

from ooxml_math import (
    Accent,
    Bar,
    BorderBox,
    Box,
    Delimiter,
    ElementProxy,
    EqArray,
    Fraction,
    FuncApply,
    GroupChar,
    Integral,
    Limit,
    Lit,
    MathExpr,
    MathExprLike,
    Matrix,
    Nary,
    Phantom,
    Pre,
    Product,
    Radical,
    Raw,
    Sub,
    SubSup,
    Sum,
    Sup,
    Text,
    Var,
    from_element,
    oMath,
)

__all__ = [
    # -- abstract base --------------------------------------------------
    "ElementProxy",
    "MathExpr",
    "MathExprLike",
    # -- leaves --------------------------------------------------------
    "Lit",
    "Raw",
    "Text",
    "Var",
    # -- operator-tree proxies -----------------------------------------
    "Accent",
    "Bar",
    "BorderBox",
    "Box",
    "Delimiter",
    "EqArray",
    "Fraction",
    "FuncApply",
    "GroupChar",
    "Integral",
    "Limit",
    "Matrix",
    "Nary",
    "Phantom",
    "Pre",
    "Product",
    "Radical",
    "Sub",
    "SubSup",
    "Sum",
    "Sup",
    # -- root + parse dispatch -----------------------------------------
    "from_element",
    "oMath",
]
