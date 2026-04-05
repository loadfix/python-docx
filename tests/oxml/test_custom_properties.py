# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.oxml.custom_properties` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.custom_properties import CT_CustomProperties, CT_CustomProperty
from docx.oxml.ns import qn


class DescribeCT_CustomProperties:
    """Unit-test suite for `docx.oxml.custom_properties.CT_CustomProperties`."""

    def it_can_construct_a_new_element(self):
        props = CT_CustomProperties.new()

        assert props.tag == qn("cust:Properties")
        assert props.property_lst == []

    def it_can_add_a_property(self):
        props = CT_CustomProperties.new()

        prop = props.add_property("TestProp")

        assert isinstance(prop, CT_CustomProperty)
        assert prop.name == "TestProp"
        assert prop.pid == 2

    def it_assigns_incrementing_pids(self):
        props = CT_CustomProperties.new()
        p1 = props.add_property("First")
        p2 = props.add_property("Second")
        p3 = props.add_property("Third")

        assert p1.pid == 2
        assert p2.pid == 3
        assert p3.pid == 4

    def it_can_find_a_property_by_name(self):
        props = CT_CustomProperties.new()
        props.add_property("Target")
        props.add_property("Other")

        result = props.get_property_by_name("Target")

        assert result is not None
        assert result.name == "Target"

    def it_returns_None_for_missing_name(self):
        props = CT_CustomProperties.new()

        assert props.get_property_by_name("Missing") is None

    def it_provides_a_property_list(self):
        props = CT_CustomProperties.new()
        props.add_property("A")
        props.add_property("B")

        lst = props.property_lst
        assert len(lst) == 2


class DescribeCT_CustomProperty:
    """Unit-test suite for `docx.oxml.custom_properties.CT_CustomProperty`."""

    def it_can_get_and_set_a_string_value(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Str")
        prop.value = "hello"

        assert prop.value == "hello"

    def it_can_get_and_set_an_int_value(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Int")
        prop.value = 42

        assert prop.value == 42
        assert isinstance(prop.value, int)

    def it_can_get_and_set_a_float_value(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Float")
        prop.value = 2.718

        assert prop.value == 2.718

    def it_can_get_and_set_a_bool_value(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Bool")

        prop.value = True
        assert prop.value is True

        prop.value = False
        assert prop.value is False

    def it_can_get_and_set_a_datetime_value(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Date")
        value = dt.datetime(2024, 6, 15, 8, 30, 0, tzinfo=dt.timezone.utc)

        prop.value = value

        assert prop.value == value

    def it_raises_on_naive_datetime(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Date")
        naive = dt.datetime(2024, 6, 15, 8, 30, 0)

        with pytest.raises(ValueError, match="timezone-aware"):
            prop.value = naive

    def it_returns_empty_string_for_no_value_child(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("Empty")
        # no value set yet
        assert prop.value == ""
