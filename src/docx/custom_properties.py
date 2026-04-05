"""CustomProperties and related proxy objects."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from docx.oxml.custom_properties import CT_CustomProperties, CT_CustomProperty


class CustomProperties:
    """Provides read/write access to custom document properties stored in
    ``docProps/custom.xml``."""

    def __init__(self, element: CT_CustomProperties):
        self._element = element

    def __getitem__(self, name: str) -> str | int | float | bool | dt.datetime:
        prop = self._element.property_by_name(name)
        if prop is None:
            raise KeyError("no custom property with name '%s'" % name)
        return prop.value

    def __iter__(self) -> Iterator[CustomProperty]:
        for prop_elm in self._element.property_lst:
            yield CustomProperty(prop_elm)

    def __len__(self) -> int:
        return len(self._element.property_lst)

    def __contains__(self, name: str) -> bool:
        return self._element.property_by_name(name) is not None

    def add(self, name: str, value: str | int | float | bool | dt.datetime) -> CustomProperty:
        if self._element.property_by_name(name) is not None:
            raise ValueError("a custom property with name '%s' already exists" % name)
        prop_elm = self._element.add_property(name, value)
        return CustomProperty(prop_elm)


class CustomProperty:
    """Proxy for a single custom document property."""

    def __init__(self, element: CT_CustomProperty):
        self._element = element

    @property
    def name(self) -> str:
        return self._element.property_name

    @property
    def value(self) -> str | int | float | bool | dt.datetime:
        return self._element.value

    @value.setter
    def value(self, val: str | int | float | bool | dt.datetime) -> None:
        self._element.value = val
