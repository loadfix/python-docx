"""Re-export of :class:`ooxml_docprops.CoreProperties`.

Historically ``docx.opc.coreprops`` defined the ``CoreProperties`` proxy
inline. As of 2026.05 the canonical implementation lives in the shared
:mod:`ooxml_docprops` package; this module keeps the historical import
path working for downstream consumers.

.. versionchanged:: 2026.05.0
    Implementation relocated to ``python-ooxml-docprops``.
"""

from __future__ import annotations

from ooxml_docprops import CoreProperties

__all__ = ["CoreProperties"]
