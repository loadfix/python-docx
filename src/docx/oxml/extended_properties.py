"""Re-export of :class:`ooxml_docprops.oxml.CT_ExtendedProperties`.

Historically ``docx.oxml.extended_properties`` defined
``CT_ExtendedProperties`` inline. As of 2026.05 the canonical
implementation lives in the shared :mod:`ooxml_docprops.oxml` package;
this module keeps the historical import path working for downstream
consumers.

.. versionchanged:: 2026.05.0
    Implementation relocated to ``python-ooxml-docprops``.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Namespace-registry safety: importing ``ooxml_docprops.oxml`` reconfigures
# the process-global ``ooxml_xmlchemy`` namespace registry to the shared
# docprops one. Restore docx's registry before returning so subsequent
# CT_* imports in ``docx.oxml.__init__`` resolve their descriptors against
# the docx registry (which knows ``w:``, ``wp:``, ``m:``, ... prefixes).
# ---------------------------------------------------------------------------
from ooxml_docprops.oxml import CT_ExtendedProperties  # noqa: F401
from ooxml_xmlchemy import configure_namespace_registry as _configure

from docx.oxml.parser import _DocxNamespaceRegistry as _DocxRegistry

_configure(_DocxRegistry())

__all__ = ["CT_ExtendedProperties"]
