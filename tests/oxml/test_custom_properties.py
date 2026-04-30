# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.custom_properties` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.custom_properties import (
    CUSTOM_PROPERTIES_FMTID,
    CT_CustomProperties,
    CT_CustomProperty,
)
from docx.oxml.parser import parse_xml


_EMPTY_PROPERTIES_XML = (
    b'<Properties '
    b'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
)


def _empty_properties() -> CT_CustomProperties:
    return cast(CT_CustomProperties, parse_xml(_EMPTY_PROPERTIES_XML))


class DescribeCT_CustomProperties:
    """Unit-test suite for `docx.oxml.custom_properties.CT_CustomProperties`."""

    def it_exposes_an_empty_property_lst_initially(self):
        props = _empty_properties()

        assert props.property_lst == []

    def it_can_add_a_property_with_a_unique_pid(self):
        props = _empty_properties()

        p1 = props.add_property("Project")
        p2 = props.add_property("Year")

        assert p1.pid == 2
        assert p2.pid == 3
        assert p1.name == "Project"
        assert p2.name == "Year"
        assert p1.fmtid == CUSTOM_PROPERTIES_FMTID
        assert p2.fmtid == CUSTOM_PROPERTIES_FMTID

    def it_picks_the_lowest_unused_pid(self):
        props = _empty_properties()
        p2 = props.add_property("A")
        p3 = props.add_property("B")
        props.remove(p2)  # pid 2 is now free

        p_new = props.add_property("C")

        # -- pid 3 is still in use, pid 2 is free --
        assert p_new.pid == 2
        del p3

    def it_can_find_a_property_by_name(self):
        props = _empty_properties()
        props.add_property("Alpha")
        target = props.add_property("Beta")
        props.add_property("Gamma")

        found = props.get_property_by_name("Beta")

        assert found is target

    def but_it_returns_None_for_an_unknown_name(self):
        props = _empty_properties()
        props.add_property("Alpha")

        assert props.get_property_by_name("Missing") is None


class DescribeCT_CustomProperty:
    """Unit-test suite for `docx.oxml.custom_properties.CT_CustomProperty`."""

    @pytest.mark.parametrize(
        ("value", "expected_localname", "expected_text"),
        [
            ("hello", "lpwstr", "hello"),
            (42, "i4", "42"),
            (3.14, "r8", "3.14"),
            (True, "bool", "true"),
            (False, "bool", "false"),
        ],
    )
    def it_writes_the_appropriate_vt_child_for_each_type(
        self, value: object, expected_localname: str, expected_text: str
    ):
        props = _empty_properties()
        prop = props.add_property("P")

        prop.value = value

        child = prop._vt_child
        assert child is not None
        assert child.tag.split("}", 1)[-1] == expected_localname
        assert child.text == expected_text

    def it_writes_a_filetime_for_a_datetime(self):
        props = _empty_properties()
        prop = props.add_property("P")

        prop.value = dt.datetime(2024, 1, 15, 10, 30, 0)

        child = prop._vt_child
        assert child is not None
        assert child.tag.endswith("}filetime")
        assert child.text == "2024-01-15T10:30:00Z"

    def it_converts_aware_datetimes_to_utc(self):
        props = _empty_properties()
        prop = props.add_property("P")
        tz = dt.timezone(dt.timedelta(hours=5))

        prop.value = dt.datetime(2024, 1, 15, 15, 30, 0, tzinfo=tz)

        child = prop._vt_child
        assert child is not None
        # -- 15:30 +05:00 == 10:30 UTC --
        assert child.text == "2024-01-15T10:30:00Z"

    @pytest.mark.parametrize(
        ("assigned", "expected"),
        [
            ("hello", "hello"),
            (42, 42),
            (3.14, 3.14),
            (True, True),
            (False, False),
        ],
    )
    def it_round_trips_each_supported_scalar_type(self, assigned: object, expected: object):
        props = _empty_properties()
        prop = props.add_property("P")

        prop.value = assigned

        assert prop.value == expected
        assert type(prop.value) is type(expected)

    def it_round_trips_a_datetime(self):
        props = _empty_properties()
        prop = props.add_property("P")
        original = dt.datetime(2024, 1, 15, 10, 30, 45)

        prop.value = original

        assert prop.value == original

    def it_raises_on_unsupported_value_type(self):
        props = _empty_properties()
        prop = props.add_property("P")

        with pytest.raises(TypeError):
            prop.value = object()

    def it_replaces_the_existing_value_on_reassignment(self):
        props = _empty_properties()
        prop = props.add_property("P")
        prop.value = "first"

        prop.value = 99

        assert prop.value == 99
        assert type(prop.value) is int

    def it_exposes_pid_as_int(self):
        prop = cast(
            CT_CustomProperty,
            parse_xml(
                b'<property '
                b'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                b'fmtid="{X}" pid="7" name="Foo"/>'
            ),
        )

        assert prop.pid == 7
        assert prop.name == "Foo"
        assert prop.fmtid == "{X}"
