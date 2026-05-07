"""Re-export of :mod:`ooxml_opc.rel` with docx-shape adapter on
:class:`Relationships`.

The :class:`_Relationship` value-object lives in :mod:`ooxml_opc.rel` and is
re-exported verbatim (it already matches the docx-shape constructor signature
``(rId, reltype, target, baseURI, external)``).

:class:`Relationships` is wrapped in a thin docx-local subclass that preserves
two pre-adoption behaviours:

1. ``dict`` semantics — docx callers (and a handful of tests) assign into the
   collection directly via ``rels[rId] = rel``. The shared
   :class:`_Relationships` is a :class:`~collections.abc.Mapping` and does not
   expose ``__setitem__`` publicly.
2. :meth:`~Relationships.get_or_add` returns the :class:`_Relationship`
   (docx shape). The shared collection's method returns the rId string
   (pptx shape); docx's returns the value-object. The shim delegates to the
   shared ``get_or_add_rel`` for matching semantics.
3. :meth:`~Relationships.add_relationship` calls through the module-local
   :class:`_Relationship` reference so test fixtures that patch
   ``docx.opc.rel._Relationship`` via :func:`class_mock` continue to work.

.. versionchanged:: 2026.05.11
   Re-exported from :mod:`ooxml_opc.rel`; docx-shape API preserved via a thin
   :class:`Relationships` subclass.
"""

from __future__ import annotations

import contextlib
from typing import TYPE_CHECKING, Any

from ooxml_opc.rel import _Relationship as _SharedRelationship
from ooxml_opc.rel import _Relationships as _SharedRelationships

if TYPE_CHECKING:
    from docx.opc.part import Part

__all__ = ["Relationships", "_Relationship"]


class _Relationship(_SharedRelationship):
    """docx-shape :class:`~ooxml_opc.rel._Relationship`.

    Identical to the shared value-object except for the error message raised
    when :attr:`target_part` is accessed on an external rel — kept to preserve
    the exact regex that pre-adoption callers (and ``pytest.raises(match=...)``
    fixtures) anchor on.
    """

    @property
    def target_part(self):
        """The target :class:`~ooxml_opc.part.Part` (internal rels only)."""
        if self._is_external:
            raise ValueError(
                "target_part property on _Relationship is undefined when "
                'target mode is "External"'
            )
        from typing import cast as _cast

        from docx.opc.part import Part as _Part

        return _cast(_Part, self._target)


class Relationships(_SharedRelationships):
    """docx-shape :class:`~ooxml_opc.rel._Relationships`.

    Preserves two pre-adoption behaviours that docx-local callers and test
    fixtures depend on:

    * ``rels[rId] = rel`` — :meth:`__setitem__` support.
    * :meth:`get_or_add` returns the :class:`_Relationship` value-object
      rather than the rId string.
    """

    def __setitem__(self, rId: Any, rel: _Relationship) -> None:
        """Permit direct dict-style assignment of a :class:`_Relationship`."""
        self._rels[rId] = rel
        if not getattr(rel, "is_external", False):
            # -- Mock objects used in tests may not implement target_part;
            # -- ignore and let the caller inspect via the rels mapping. --
            with contextlib.suppress(ValueError, AttributeError):
                self._target_parts_by_rId[rId] = rel.target_part

    def add_relationship(
        self,
        reltype: str,
        target: Part | str,
        rId: str,
        is_external: bool = False,
    ) -> _Relationship:
        """Return a newly added :class:`_Relationship` with caller-supplied `rId`.

        Duplicates the shared :meth:`_Relationships.add_relationship` logic
        but calls the module-local :class:`_Relationship` name so
        ``class_mock(request, "docx.opc.rel._Relationship")`` in tests patches
        the constructor this method invokes.
        """
        rel = _Relationship(rId, reltype, target, self._base_uri, is_external)
        self._rels[rId] = rel
        if not is_external:
            self._target_parts_by_rId[rId] = target  # type: ignore[assignment]
        return rel

    def get_or_add(self, reltype: str, target_part: Part) -> _Relationship:  # type: ignore[override]
        """Return the :class:`_Relationship` of `reltype` to `target_part`.

        docx-shape override — the shared base returns the rId string; docx
        callers expect the :class:`_Relationship` value-object. Delegates to
        the shared ``get_or_add_rel`` method.
        """
        return self.get_or_add_rel(reltype, target_part)
