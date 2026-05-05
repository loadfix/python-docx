"""Collection providing access to custom document properties.

Custom properties are user-defined, typed name/value pairs stored in the
``docProps/custom.xml`` part of the package. They are distinct from the fixed
Dublin-Core "core" properties available via `document.core_properties`.

Supported value types:

* ``str``  -- serialised as ``vt:lpwstr``
* ``int``  -- serialised as ``vt:i4``
* ``float`` -- serialised as ``vt:r8``
* ``bool`` -- serialised as ``vt:bool``
* ``datetime.datetime`` -- serialised as ``vt:filetime`` (ISO-8601 with trailing ``Z``)
* ``datetime.date`` (but *not* ``datetime``) -- serialised as ``vt:date``
  (ISO-8601 ``YYYY-MM-DD``, no time component)
"""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Iterator

if TYPE_CHECKING:
    from docx.oxml.custom_properties import CT_CustomProperties
    from docx.parts.custom_properties import CustomPropertiesPart


_MISSING = object()


class CustomProperties:
    """Collection of custom document properties.

    Behaves like a mapping keyed by property name. Iteration yields property names
    (matching ``dict``-style iteration convention).

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        element: "CT_CustomProperties",
        part: "CustomPropertiesPart",
    ):
        self._element = element
        self._part = part

    # -- mapping protocol --------------------------------------------------------------

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        return self._element.get_property_by_name(name) is not None

    def __delitem__(self, name: str) -> None:
        prop = self._element.get_property_by_name(name)
        if prop is None:
            raise KeyError(name)
        self._element.remove(prop)

    def __getitem__(self, name: str) -> object:
        prop = self._element.get_property_by_name(name)
        if prop is None:
            raise KeyError(name)
        return prop.value

    def __iter__(self) -> Iterator[str]:
        """Yield the name of each custom property, in document order."""
        return (prop.name for prop in self._element.property_lst)

    def __len__(self) -> int:
        return len(self._element.property_lst)

    def __setitem__(self, name: str, value: object) -> None:
        """Set `name` to `value`, overwriting any existing property with that name."""
        self._validate_value_type(value)
        existing = self._element.get_property_by_name(name)
        if existing is not None:
            existing.value = value
            return
        prop = self._element.add_property(name)
        prop.value = value

    # -- convenience methods -----------------------------------------------------------

    def add(self, name: str, value: object) -> None:
        """Add a new custom property named `name` with `value`.

        Raises |ValueError| if a property with that name already exists. Use
        ``custom_properties[name] = value`` to overwrite.

        .. versionadded:: 2026.05.0
        """
        self._validate_value_type(value)
        if self._element.get_property_by_name(name) is not None:
            raise ValueError(f"a custom property named {name!r} already exists")
        prop = self._element.add_property(name)
        prop.value = value

    def get(self, name: str, default: object = None) -> object:
        """Return the value of property `name`, or `default` if not present.

        .. versionadded:: 2026.05.0
        """
        prop = self._element.get_property_by_name(name)
        if prop is None:
            return default
        return prop.value

    def names(self) -> list[str]:
        """Return a list of the names of each property, in document order.

        .. versionadded:: 2026.05.0
        """
        return [prop.name for prop in self._element.property_lst]

    def items(self) -> list[tuple[str, object]]:
        """Return a list of ``(name, value)`` pairs, in document order.

        .. versionadded:: 2026.05.0
        """
        return [(prop.name, prop.value) for prop in self._element.property_lst]

    @staticmethod
    def _validate_value_type(value: object) -> None:
        """Raise |TypeError| if `value` is not a supported custom-property value type.

        The accepted ``isinstance`` check matches the set of types supported by
        ``CT_CustomProperty.value`` setter.
        """
        # -- ``bool`` must be accepted explicitly; it subclasses ``int``. ----
        # -- ``date`` covers both ``date`` and ``datetime`` (the setter dispatch
        # -- below is responsible for picking `vt:filetime` vs `vt:date`).
        import datetime as _dt

        if isinstance(value, (bool, int, float, str, _dt.date)):
            return
        raise TypeError(
            f"unsupported custom-property value type: {type(value).__name__}"
        )
