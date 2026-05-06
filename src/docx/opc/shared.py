"""Re-export of :mod:`ooxml_opc.shared`.

:class:`CaseInsensitiveDict` and :func:`cls_method_fn` now live in the
shared :mod:`ooxml_opc` package.
"""

from __future__ import annotations

from ooxml_opc.shared import CaseInsensitiveDict, cls_method_fn

__all__ = ["CaseInsensitiveDict", "cls_method_fn"]
