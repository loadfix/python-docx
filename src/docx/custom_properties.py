"""CustomProperties and closely related objects."""

from __future__ import annotations

import datetime as dt
from typing import Iterator

from docx.oxml.custom_properties import CT_CustomProperties, CT_CustomProperty


class CustomProperties:
    """Provides read/write access to custom document properties stored in
    ``/docProps/custom.xml``."""

    def __init__(self, element: CT_CustomProperties):
        self._element = element

    def __getitem__(self, name: str) -> str | int | float | bool | dt.datetime:
        """Return the value of the custom property with `name`.

        Raises |KeyError| if no property with `name` exists.
        """
        prop = self._element.get_property_by_name(name)
        if prop is None:
            raise KeyError("no custom property with name '%s'" % name)
        return prop.value

    def __iter__(self) -> Iterator[CustomProperty]:
        """Generate a |CustomProperty| object for each custom property."""
        for prop_elm in self._element.property_lst:
            yield CustomProperty(prop_elm)

    def __len__(self) -> int:
        """The number of custom properties."""
        return len(self._element.property_lst)

    def add(
        self, name: str, value: str | int | float | bool | dt.datetime
    ) -> CustomProperty:
        """Add a custom property with `name` and `value`.

        `value` can be a string, int, float, bool, or datetime object.
        Raises |ValueError| if a property with `name` already exists.
        """
        if self._element.get_property_by_name(name) is not None:
            raise ValueError("a custom property with name '%s' already exists" % name)
        prop_elm = self._element.add_property(name)
        prop_elm.value = value
        return CustomProperty(prop_elm)

    def __contains__(self, name: str) -> bool:
        """Return True if a custom property with `name` exists."""
        return self._element.get_property_by_name(name) is not None


class CustomProperty:
    """A single custom document property, providing access to its name and value."""

    def __init__(self, element: CT_CustomProperty):
        self._element = element

    @property
    def name(self) -> str:
        """The name of this custom property."""
        return self._element.name

    @property
    def value(self) -> str | int | float | bool | dt.datetime:
        """The value of this custom property."""
        return self._element.value

    @value.setter
    def value(self, val: str | int | float | bool | dt.datetime) -> None:
        """Set the value of this custom property."""
        self._element.value = val
