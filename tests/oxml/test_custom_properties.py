# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.oxml.custom_properties module."""

from __future__ import annotations

import datetime as dt

from docx.oxml.custom_properties import CT_CustomProperties, CT_CustomProperty


class DescribeCT_CustomProperties:
    """Unit-test suite for `docx.oxml.custom_properties.CT_CustomProperties`."""

    def it_can_construct_a_new_empty_element(self):
        props = CT_CustomProperties.new()
        assert props.property_lst == []

    def it_can_add_a_string_property(self):
        props = CT_CustomProperties.new()

        prop = props.add_property("Author", "Jane Doe")

        assert isinstance(prop, CT_CustomProperty)
        assert prop.property_name == "Author"
        assert prop.value == "Jane Doe"

    def it_can_add_an_int_property(self):
        props = CT_CustomProperties.new()

        prop = props.add_property("Count", 42)

        assert prop.value == 42

    def it_can_add_a_float_property(self):
        props = CT_CustomProperties.new()

        prop = props.add_property("Rate", 2.5)

        assert prop.value == 2.5

    def it_can_add_a_bool_property(self):
        props = CT_CustomProperties.new()

        prop = props.add_property("Approved", False)

        assert prop.value is False

    def it_can_add_a_datetime_property(self):
        props = CT_CustomProperties.new()
        when = dt.datetime(2024, 1, 15, 8, 0, 0, tzinfo=dt.timezone.utc)

        prop = props.add_property("DueDate", when)

        assert prop.value == when

    def it_assigns_incrementing_pids(self):
        props = CT_CustomProperties.new()
        p1 = props.add_property("A", "a")
        p2 = props.add_property("B", "b")

        assert int(p1.get("pid")) == 2
        assert int(p2.get("pid")) == 3

    def it_can_find_a_property_by_name(self):
        props = CT_CustomProperties.new()
        props.add_property("Target", "found")

        result = props.property_by_name("Target")

        assert result is not None
        assert result.value == "found"

    def it_returns_None_for_missing_property_name(self):
        props = CT_CustomProperties.new()

        assert props.property_by_name("Ghost") is None


class DescribeCT_CustomProperty:
    """Unit-test suite for `docx.oxml.custom_properties.CT_CustomProperty`."""

    def it_can_set_its_value(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Editable", "old")

        prop.value = "new"

        assert prop.value == "new"
