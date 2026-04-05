"""Custom element classes for custom document properties XML elements.

Custom properties are stored in ``docProps/custom.xml`` and use the
``http://schemas.openxmlformats.org/officeDocument/2006/custom-properties`` namespace.
Each property has a typed value drawn from the ``vt:`` (docPropsVTypes) namespace.
"""

from __future__ import annotations

import datetime as dt
from typing import List, cast

from lxml import etree

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement


class CT_CustomProperties(BaseOxmlElement):
    """`<Properties>` element, root of the custom properties part."""

    _customProperties_tmpl = (
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
        'custom-properties" %s/>\n' % nsdecls("vt")
    )

    @classmethod
    def new(cls) -> CT_CustomProperties:
        xml = cls._customProperties_tmpl
        return cast(CT_CustomProperties, parse_xml(xml))

    @property
    def property_lst(self) -> List[CT_CustomProperty]:
        return self.findall(qn("cust-p:property"))  # pyright: ignore[reportReturnType]

    @property
    def _next_pid(self) -> int:
        pids = [int(p.get("pid", "1")) for p in self.property_lst]
        return max(pids, default=1) + 1

    def add_property(self, name: str, value: str | int | float | bool | dt.datetime) -> CT_CustomProperty:
        prop = CT_CustomProperty.new(self._next_pid, name, value)
        self.append(prop)
        return prop

    def property_by_name(self, name: str) -> CT_CustomProperty | None:
        for prop in self.property_lst:
            if prop.property_name == name:
                return prop
        return None


class CT_CustomProperty(BaseOxmlElement):
    """`<property>` element within custom properties."""

    @staticmethod
    def new(pid: int, name: str, value: str | int | float | bool | dt.datetime) -> CT_CustomProperty:
        vt_tag, vt_text = CT_CustomProperty._value_to_vt(value)
        xml = (
            '<property xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'custom-properties" %s fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            ' pid="%d" name="%s"><%s>%s</%s></property>'
            % (nsdecls("vt"), pid, name, vt_tag, vt_text, vt_tag)
        )
        return cast(CT_CustomProperty, parse_xml(xml))

    @property
    def property_name(self) -> str:
        return self.get("name", "")

    @property
    def value(self) -> str | int | float | bool | dt.datetime:
        for child in self:
            tag = etree.QName(child.tag).localname
            text = child.text or ""
            if tag == "lpwstr":
                return text
            elif tag == "i4":
                return int(text)
            elif tag == "r8":
                return float(text)
            elif tag == "bool":
                return text.lower() == "true"
            elif tag == "filetime":
                return dt.datetime.strptime(text, "%Y-%m-%dT%H:%M:%SZ").replace(
                    tzinfo=dt.timezone.utc
                )
        return ""

    @value.setter
    def value(self, val: str | int | float | bool | dt.datetime) -> None:
        for child in list(self):
            self.remove(child)
        vt_tag, vt_text = self._value_to_vt(val)
        child = parse_xml("<%s %s>%s</%s>" % (vt_tag, nsdecls("vt"), vt_text, vt_tag))
        self.append(child)

    @staticmethod
    def _value_to_vt(value: str | int | float | bool | dt.datetime) -> tuple[str, str]:
        if isinstance(value, bool):
            return "vt:bool", str(value).lower()
        elif isinstance(value, int):
            return "vt:i4", str(value)
        elif isinstance(value, float):
            return "vt:r8", str(value)
        elif isinstance(value, dt.datetime):
            return "vt:filetime", value.strftime("%Y-%m-%dT%H:%M:%SZ")
        else:
            return "vt:lpwstr", str(value)
