"""Collection providing access to custom document properties.

Custom properties are user-defined, typed name/value pairs stored in the
``docProps/custom.xml`` part of the package. They are distinct from the fixed
Dublin-Core "core" properties available via ``document.core_properties``.

Supported value types:

* ``str``     — serialised as ``vt:lpwstr``
* ``int``     — serialised as ``vt:i4``
* ``float``   — serialised as ``vt:r8``
* ``bool``    — serialised as ``vt:bool``
* ``datetime.datetime`` — serialised as ``vt:filetime`` (ISO-8601 with ``Z``)
* ``datetime.date`` (not ``datetime``) — serialised as ``vt:date``
  (ISO-8601 ``YYYY-MM-DD``, no time component)

.. versionchanged:: 2026.05.0
    Delegates to :class:`ooxml_docprops.CustomProperties`. The docx-local
    class remains as a thin adapter preserving the
    ``(element, part)`` constructor, the ``ValueError``-on-duplicate
    contract of :meth:`add`, and the list-returning :meth:`names`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ooxml_docprops import CustomProperties as _SharedCustomProperties

if TYPE_CHECKING:
    from docx.oxml.custom_properties import CT_CustomProperties
    from docx.parts.custom_properties import CustomPropertiesPart


class CustomProperties(_SharedCustomProperties):
    """Mapping-like collection of custom document properties.

    Thin adapter over :class:`ooxml_docprops.CustomProperties` that
    preserves three pre-2026.05 docx API contracts:

    1. ``CustomProperties(element, part)`` two-argument constructor
       (the shared base only takes ``element``).
    2. :meth:`add` raises :class:`ValueError` on a duplicate name; the
       shared base raises :class:`KeyError`.
    3. :meth:`names` returns a concrete ``list[str]``; the shared base
       returns an iterator.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        element: "CT_CustomProperties",
        part: "CustomPropertiesPart" | None = None,
    ):
        super().__init__(element)
        self._part = part

    def add(self, name: str, value: Any) -> None:
        """Add a new property named *name* with *value*.

        Raises :class:`ValueError` if a property with that name already
        exists. Use ``custom_properties[name] = value`` for last-write-wins
        semantics.
        """
        # -- the shared base raises KeyError on duplicate; docx's historical
        # -- contract was ValueError. Translate. Any other exception from the
        # -- shared impl (e.g. InvalidCustomPropertyTypeError / TypeError on
        # -- unsupported value types) is propagated unchanged.
        try:
            super().add(name, value)
        except KeyError as exc:
            raise ValueError(*exc.args) from None

    def names(self) -> list[str]:  # type: ignore[override]
        """Return a list of property names in document order.

        Overrides the shared base's iterator-returning ``names()`` to
        preserve docx's list contract.
        """
        return list(super().names())
