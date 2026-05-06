"""Re-export of custom-properties element classes from :mod:`ooxml_docprops`.

Historically ``docx.oxml.custom_properties`` defined
``CT_CustomProperties`` and ``CT_CustomProperty`` inline. As of 2026.05
the canonical implementations live in the shared
:mod:`ooxml_docprops.oxml` package; this module keeps the historical
import path working for downstream consumers and surfaces the docx
friendly-name alias :data:`CUSTOM_PROPERTIES_FMTID`.

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
from ooxml_docprops import CUSTOM_PROPERTIES_FMTID
from ooxml_docprops.oxml import (
    CT_CustomProperties as _CT_CustomProperties_Base,
    CT_CustomProperty,
)
from ooxml_xmlchemy import configure_namespace_registry as _configure

from docx.oxml.parser import _DocxNamespaceRegistry as _DocxRegistry


class CT_CustomProperties(_CT_CustomProperties_Base):
    """Back-compat wrapper that preserves docx's pre-2026.05 API shape.

    Two behavioural differences from the shared base:

    1. :meth:`add_property` accepts a ``name`` argument (route to
       :meth:`add_new_property`). The shared base's descriptor-generated
       ``add_property()`` takes no args; keep the friendlier overload
       here for existing callers.

    2. :meth:`_next_available_pid` returns the lowest unused ``pid``
       instead of ``max(pid) + 1``. docx's pre-2026.05 allocator reused
       freed pids; preserve that for deterministic fixture output and
       existing test expectations.
    """

    # -- type declaration so pyright / IDE users see the real signature --
    def add_property(  # type: ignore[override]
        self,
        name: str,
        fmtid: str = CUSTOM_PROPERTIES_FMTID,
    ) -> CT_CustomProperty:
        """Append a new ``<property>`` child with *name* and return it.

        Delegates to :meth:`add_new_property` on the shared base; kept as a
        method to preserve the historical docx name.
        """
        return self.add_new_property(name, fmtid)

    def _next_available_pid(self) -> int:
        """Return the lowest unused ``pid`` (>=2), reusing freed slots.

        Overrides the shared base's ``max(pid) + 1`` strategy so docx keeps
        its historical behaviour of collapsing ``pid`` gaps. The two
        strategies are both spec-legal; the docx one keeps fixtures stable
        under add/remove/add patterns.
        """
        used = {int(p.get("pid", "0")) for p in self.property_lst}
        candidate = 2
        while candidate in used:
            candidate += 1
        return candidate


# -- Restore docx's registry after shared-package imports have completed. --
_configure(_DocxRegistry())

__all__ = [
    "CUSTOM_PROPERTIES_FMTID",
    "CT_CustomProperties",
    "CT_CustomProperty",
]
