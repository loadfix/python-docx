"""Re-export of :mod:`ooxml_opc.strict`.

The Strict→Transitional rewriter now lives in the shared
:mod:`ooxml_opc` package. Keeps the ``docx.opc.strict.*`` import paths
working for every existing caller.
"""

from __future__ import annotations

from ooxml_opc.strict import (
    STRICT_SENTINEL,
    STRICT_TO_TRANSITIONAL,
    is_strict_document_xml,
    translate_many,
    translate_strict_blob,
)

__all__ = [
    "STRICT_SENTINEL",
    "STRICT_TO_TRANSITIONAL",
    "is_strict_document_xml",
    "translate_many",
    "translate_strict_blob",
]
