"""Custom element classes for custom document properties XML elements.

Custom properties are stored in ``docProps/custom.xml`` and allow arbitrary
name/value metadata to be attached to a document.
"""

from __future__ import annotations

import datetime as dt
from typing import cast

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement

# -- The FMTID is always this GUID for custom properties --
_CUSTOM_PROPS_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


class CT_CustomProperties(BaseOxmlElement):
    """`<Properties>` element, root of the custom properties part.

    Stored as ``/docProps/custom.xml``.
    """

    _customProperties_tmpl = "<cust:Properties %s/>\n" % nsdecls("cust", "vt")

    @classmethod
    def new(cls) -> CT_CustomProperties:
        """Return a new ``<Properties>`` element."""
        xml = cls._customProperties_tmpl
        return cast(CT_CustomProperties, parse_xml(xml))

    @property
    def property_lst(self) -> list[CT_CustomProperty]:
        """All ``<property>`` child elements."""
        return list(self.iterchildren(qn("cust:property")))

    @property
    def _next_pid(self) -> int:
        """The next available property id. PIDs start at 2."""
        pids = [prop.pid for prop in self.property_lst]
        return max(pids, default=1) + 1

    def add_property(self, name: str) -> CT_CustomProperty:
        """Add a new ``<property>`` child element with `name` and return it."""
        from lxml.etree import SubElement

        prop = cast(
            CT_CustomProperty,
            SubElement(
                self,
                qn("cust:property"),
                attrib={
                    "fmtid": _CUSTOM_PROPS_FMTID,
                    "pid": str(self._next_pid),
                    "name": name,
                },
            ),
        )
        return prop

    def get_property_by_name(self, name: str) -> CT_CustomProperty | None:
        """Return the ``<property>`` element with `name`, or None."""
        for prop in self.property_lst:
            if prop.name == name:
                return prop
        return None


class CT_CustomProperty(BaseOxmlElement):
    """`<property>` element, a single custom property."""

    @property
    def name(self) -> str:
        """The `name` attribute of this property."""
        return self.get("name", "")

    @property
    def pid(self) -> int:
        """The `pid` attribute of this property."""
        return int(self.get("pid", "0"))

    @property
    def value(self) -> str | int | float | bool | dt.datetime:
        """The typed value of this property, derived from its child element."""
        for child in self:
            tag = child.tag
            text = child.text or ""
            if tag == qn("vt:lpwstr"):
                return text
            if tag == qn("vt:i4"):
                return int(text)
            if tag == qn("vt:r8"):
                return float(text)
            if tag == qn("vt:bool"):
                return text.lower() in ("true", "1")
            if tag == qn("vt:filetime"):
                return dt.datetime.strptime(text, "%Y-%m-%dT%H:%M:%SZ").replace(
                    tzinfo=dt.timezone.utc
                )
        return ""

    @value.setter
    def value(self, val: str | int | float | bool | dt.datetime) -> None:
        """Set the value, replacing any existing value child element."""
        # -- remove existing value children --
        for child in list(self):
            self.remove(child)

        if isinstance(val, bool):
            child_tag = qn("vt:bool")
            text = "true" if val else "false"
        elif isinstance(val, int):
            child_tag = qn("vt:i4")
            text = str(val)
        elif isinstance(val, float):
            child_tag = qn("vt:r8")
            text = str(val)
        elif isinstance(val, dt.datetime):
            child_tag = qn("vt:filetime")
            text = val.strftime("%Y-%m-%dT%H:%M:%SZ")
        else:
            child_tag = qn("vt:lpwstr")
            text = str(val)

        from lxml.etree import SubElement

        child_elm = SubElement(self, child_tag)
        child_elm.text = text
