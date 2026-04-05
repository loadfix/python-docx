# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.custom_properties module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.custom_properties import CustomProperties, CustomProperty
from docx.oxml.custom_properties import CT_CustomProperties


class DescribeCustomProperties:
    """Unit-test suite for `docx.custom_properties.CustomProperties`."""

    def it_can_iterate_over_custom_properties(self):
        props_elm = CT_CustomProperties.new()
        props_elm.add_property("Author", "Jane")
        props_elm.add_property("Version", 3)
        custom_props = CustomProperties(props_elm)

        props = list(custom_props)

        assert len(props) == 2
        assert all(isinstance(p, CustomProperty) for p in props)
        assert props[0].name == "Author"
        assert props[1].name == "Version"

    def it_knows_how_many_properties_it_contains(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        assert len(custom_props) == 0

        props_elm.add_property("Foo", "bar")
        assert len(custom_props) == 1

    def it_can_get_a_property_by_name(self):
        props_elm = CT_CustomProperties.new()
        props_elm.add_property("Project", "Alpha")
        custom_props = CustomProperties(props_elm)

        assert custom_props["Project"] == "Alpha"

    def it_raises_on_missing_property_name(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        with pytest.raises(KeyError, match="no custom property"):
            custom_props["Missing"]

    def it_can_check_if_a_property_exists(self):
        props_elm = CT_CustomProperties.new()
        props_elm.add_property("Exists", "yes")
        custom_props = CustomProperties(props_elm)

        assert "Exists" in custom_props
        assert "NotThere" not in custom_props

    def it_can_add_a_string_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        prop = custom_props.add("Department", "Engineering")

        assert isinstance(prop, CustomProperty)
        assert prop.name == "Department"
        assert prop.value == "Engineering"
        assert custom_props["Department"] == "Engineering"

    def it_can_add_an_int_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        custom_props.add("Count", 42)

        assert custom_props["Count"] == 42

    def it_can_add_a_float_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        custom_props.add("Rate", 3.14)

        assert custom_props["Rate"] == pytest.approx(3.14)

    def it_can_add_a_bool_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)

        custom_props.add("Approved", True)

        assert custom_props["Approved"] is True

    def it_can_add_a_datetime_property(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        when = dt.datetime(2024, 6, 15, 10, 30, 0, tzinfo=dt.timezone.utc)

        custom_props.add("ReviewDate", when)

        assert custom_props["ReviewDate"] == when

    def it_raises_when_adding_a_duplicate_name(self):
        props_elm = CT_CustomProperties.new()
        custom_props = CustomProperties(props_elm)
        custom_props.add("Unique", "first")

        with pytest.raises(ValueError, match="already exists"):
            custom_props.add("Unique", "second")


class DescribeCustomProperty:
    """Unit-test suite for `docx.custom_properties.CustomProperty`."""

    def it_provides_access_to_name_and_value(self):
        props_elm = CT_CustomProperties.new()
        prop_elm = props_elm.add_property("Status", "Draft")
        prop = CustomProperty(prop_elm)

        assert prop.name == "Status"
        assert prop.value == "Draft"

    def it_can_update_its_value(self):
        props_elm = CT_CustomProperties.new()
        prop_elm = props_elm.add_property("Status", "Draft")
        prop = CustomProperty(prop_elm)

        prop.value = "Final"

        assert prop.value == "Final"

    def it_can_change_value_type(self):
        props_elm = CT_CustomProperties.new()
        prop_elm = props_elm.add_property("Count", "not-a-number")
        prop = CustomProperty(prop_elm)

        prop.value = 99

        assert prop.value == 99
