# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.custom_properties` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.custom_properties import CustomProperties, CustomProperty
from docx.oxml.custom_properties import CT_CustomProperties
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml


class DescribeCustomProperties:
    """Unit-test suite for `docx.custom_properties.CustomProperties` objects."""

    def it_can_get_a_property_by_name(self):
        props_elm = self._properties_elm(("Author", "lpwstr", "John"))
        custom_props = CustomProperties(props_elm)

        assert custom_props["Author"] == "John"

    def it_raises_on_missing_property_name(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        with pytest.raises(KeyError):
            custom_props["Nonexistent"]

    def it_is_iterable(self):
        props_elm = self._properties_elm(
            ("Prop1", "lpwstr", "A"),
            ("Prop2", "i4", "42"),
        )
        custom_props = CustomProperties(props_elm)

        items = list(custom_props)
        assert len(items) == 2
        assert all(isinstance(item, CustomProperty) for item in items)
        assert items[0].name == "Prop1"
        assert items[1].name == "Prop2"

    def it_knows_how_many_properties_it_contains(self):
        props_elm = self._properties_elm(
            ("A", "lpwstr", "x"),
            ("B", "i4", "1"),
            ("C", "bool", "true"),
        )
        custom_props = CustomProperties(props_elm)

        assert len(custom_props) == 3

    def it_can_add_a_string_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        prop = custom_props.add("MyProp", "hello")

        assert prop.name == "MyProp"
        assert prop.value == "hello"
        assert len(custom_props) == 1

    def it_can_add_an_int_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        prop = custom_props.add("Count", 42)

        assert prop.value == 42

    def it_can_add_a_float_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        prop = custom_props.add("Rating", 3.14)

        assert prop.value == 3.14

    def it_can_add_a_bool_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        prop = custom_props.add("Reviewed", True)

        assert prop.value is True

    def it_can_add_a_datetime_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        value = dt.datetime(2024, 1, 15, 10, 30, 0, tzinfo=dt.timezone.utc)

        prop = custom_props.add("DueDate", value)

        assert prop.value == value

    def it_raises_on_duplicate_name(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        custom_props.add("Dup", "first")

        with pytest.raises(ValueError, match="already exists"):
            custom_props.add("Dup", "second")

    def it_supports_contains(self):
        props_elm = self._properties_elm(("Exists", "lpwstr", "val"))
        custom_props = CustomProperties(props_elm)

        assert "Exists" in custom_props
        assert "Missing" not in custom_props

    @staticmethod
    def _properties_elm(*props: tuple[str, str, str]) -> CT_CustomProperties:
        fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        xml = "<cust:Properties %s>" % nsdecls("cust", "vt")
        for i, (name, vtype, val) in enumerate(props, start=2):
            xml += (
                '<cust:property fmtid="%s" pid="%d" name="%s">'
                "<vt:%s>%s</vt:%s>"
                "</cust:property>" % (fmtid, i, name, vtype, val, vtype)
            )
        xml += "</cust:Properties>"
        return cast(CT_CustomProperties, parse_xml(xml))


class DescribeCustomProperty:
    """Unit-test suite for `docx.custom_properties.CustomProperty` objects."""

    def it_knows_its_name(self):
        props_elm = DescribeCustomProperties._properties_elm(("Title", "lpwstr", "Test"))
        prop = CustomProperty(props_elm.property_lst[0])

        assert prop.name == "Title"

    def it_can_get_and_set_its_value(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        prop = custom_props.add("Mutable", "original")

        assert prop.value == "original"
        prop.value = "updated"
        assert prop.value == "updated"

    def it_can_change_value_type(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        prop = custom_props.add("Flexible", "text")

        prop.value = 99
        assert prop.value == 99
        assert isinstance(prop.value, int)
