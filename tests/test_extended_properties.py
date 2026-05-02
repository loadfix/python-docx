# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.extended_properties` module."""

from __future__ import annotations

from typing import cast

from docx.extended_properties import ExtendedProperties
from docx.oxml.extended_properties import CT_ExtendedProperties
from docx.oxml.parser import parse_xml


_EMPTY_XML = (
    b'<Properties '
    b'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
)


_POPULATED_XML = (
    b'<Properties '
    b'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    b"<Template>Normal.dotm</Template>"
    b"<TotalTime>42</TotalTime>"
    b"<Pages>7</Pages>"
    b"<Words>1234</Words>"
    b"<Characters>5678</Characters>"
    b"<Application>Microsoft Word</Application>"
    b"<AppVersion>16.0</AppVersion>"
    b"<Company>Acme</Company>"
    b"<Manager>Alice</Manager>"
    b"</Properties>"
)


def _empty_props() -> ExtendedProperties:
    elm = cast(CT_ExtendedProperties, parse_xml(_EMPTY_XML))
    return ExtendedProperties(elm)


def _populated_props() -> ExtendedProperties:
    elm = cast(CT_ExtendedProperties, parse_xml(_POPULATED_XML))
    return ExtendedProperties(elm)


class DescribeExtendedProperties:
    """Unit-test suite for `docx.extended_properties.ExtendedProperties`."""

    def it_returns_None_for_missing_scalar_property(self):
        ep = _empty_props()

        assert ep.company is None
        assert ep.manager is None
        assert ep.application is None
        assert ep.pages is None
        assert ep.total_time is None

    def it_reads_common_scalar_properties(self):
        ep = _populated_props()

        assert ep.template == "Normal.dotm"
        assert ep.application == "Microsoft Word"
        assert ep.app_version == "16.0"
        assert ep.company == "Acme"
        assert ep.manager == "Alice"
        assert ep.total_time == 42
        assert ep.pages == 7
        assert ep.words == 1234
        assert ep.characters == 5678

    def it_can_round_trip_string_properties(self):
        ep = _empty_props()

        ep.company = "Acme Corp"
        ep.manager = "Bob"
        ep.application = "python-docx"

        assert ep.company == "Acme Corp"
        assert ep.manager == "Bob"
        assert ep.application == "python-docx"

    def it_can_round_trip_int_properties(self):
        ep = _empty_props()

        ep.pages = 11
        ep.total_time = 250

        assert ep.pages == 11
        assert ep.total_time == 250

    def it_removes_child_when_value_is_None(self):
        ep = _populated_props()

        ep.company = None

        assert ep.company is None
        # -- the other siblings stay put --
        assert ep.manager == "Alice"

    def it_returns_None_for_non_integer_numeric_property(self):
        # -- AppVersion is a string in the schema, integer helper should be resilient --
        ep = _populated_props()

        # pages is int, but if the XML were "foo" we still want None back;
        # simulate by using the generic setter.
        ep.set("Pages", "not-a-number")

        assert ep.pages is None

    def it_supports_generic_get_and_set(self):
        ep = _empty_props()

        ep.set("HyperlinkBase", "https://example.com")
        assert ep.get("HyperlinkBase") == "https://example.com"

    def it_can_clear_all_scalar_children(self):
        ep = _populated_props()

        ep.clear_all()

        assert ep.company is None
        assert ep.application is None
        assert ep.pages is None
        assert ep.template is None
