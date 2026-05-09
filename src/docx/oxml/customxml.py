"""Re-exports of the custom XML DataStore / Schema Library element classes.

The canonical ``CT_*`` implementations live in the shared
``python-ooxml-customxml`` package (part of the loadfix OOXML family) so
they can be reused by ``python-pptx`` and ``python-xlsx`` without
copying. This module is the stable ``docx.oxml.customxml`` import path
for the same classes — downstream callers that `import from
docx.oxml.customxml import CT_DatastoreItem` continue to work
unchanged.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

from ooxml_customxml import (
    CT_DatastoreItem,
    CT_DatastoreSchemaRef,
    CT_DatastoreSchemaRefs,
    CT_Schema,
    CT_SchemaLibrary,
)

__all__ = [
    "CT_DatastoreItem",
    "CT_DatastoreSchemaRef",
    "CT_DatastoreSchemaRefs",
    "CT_Schema",
    "CT_SchemaLibrary",
]
