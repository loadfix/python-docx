# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.custom_xml` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.custom_xml import (
    CT_Attr,
    CT_CustomXmlBlock,
    CT_CustomXmlCell,
    CT_CustomXmlPr,
    CT_CustomXmlRow,
    CT_CustomXmlRun,
)

from ..unitutil.cxml import element


# ---------------------------------------------------------------------------
# CT_Attr (w:attr)
# ---------------------------------------------------------------------------


class DescribeCT_Attr:
    """Unit-test suite for `docx.oxml.custom_xml.CT_Attr`."""

    def it_reads_its_required_name_and_val(self):
        attr = cast(CT_Attr, element("w:attr{w:name=author,w:val=Alice}"))
        assert attr.name == "author"
        assert attr.val == "Alice"

    def it_reads_its_optional_uri(self):
        attr = cast(
            CT_Attr,
            element(
                "w:attr{w:uri=http://example.org/ns,w:name=author,w:val=Alice}"
            ),
        )
        assert attr.uri == "http://example.org/ns"

    def it_returns_None_uri_when_absent(self):
        attr = cast(CT_Attr, element("w:attr{w:name=author,w:val=Alice}"))
        assert attr.uri is None

    def it_can_set_its_values(self):
        attr = cast(CT_Attr, element("w:attr{w:name=n,w:val=v}"))
        attr.val = "Bob"
        attr.uri = "http://example.org/x"
        assert attr.val == "Bob"
        assert attr.uri == "http://example.org/x"

    def it_registers_on_w_attr(self):
        attr = element("w:attr{w:name=n,w:val=v}")
        assert isinstance(attr, CT_Attr)


# ---------------------------------------------------------------------------
# CT_CustomXmlPr (w:customXmlPr)
# ---------------------------------------------------------------------------


class DescribeCT_CustomXmlPr:
    """Unit-test suite for `docx.oxml.custom_xml.CT_CustomXmlPr`."""

    def it_registers_on_w_customXmlPr(self):
        pr = element("w:customXmlPr")
        assert isinstance(pr, CT_CustomXmlPr)

    def it_has_optional_placeholder_child(self):
        pr = cast(
            CT_CustomXmlPr,
            element("w:customXmlPr/w:placeholder{w:val=DefaultPlaceholderStyle}"),
        )
        assert pr.placeholder is not None

    def it_returns_None_placeholder_when_absent(self):
        pr = cast(CT_CustomXmlPr, element("w:customXmlPr"))
        assert pr.placeholder is None

    def it_holds_an_unbounded_list_of_attr(self):
        pr = cast(
            CT_CustomXmlPr,
            element(
                "w:customXmlPr/(w:attr{w:name=a,w:val=1},"
                "w:attr{w:name=b,w:val=2},w:attr{w:name=c,w:val=3})"
            ),
        )
        assert len(pr.attr_lst) == 3
        assert [a.name for a in pr.attr_lst] == ["a", "b", "c"]

    def it_returns_empty_list_when_no_attrs(self):
        pr = cast(CT_CustomXmlPr, element("w:customXmlPr"))
        assert pr.attr_lst == []


# ---------------------------------------------------------------------------
# CT_CustomXmlBlock (w:customXml — block flavor; default registration)
# ---------------------------------------------------------------------------


class DescribeCT_CustomXmlBlock:
    """Unit-test suite for `docx.oxml.custom_xml.CT_CustomXmlBlock`."""

    def it_registers_on_w_customXml_as_the_default_flavor(self):
        cxb = element("w:customXml{w:element=root}")
        assert isinstance(cxb, CT_CustomXmlBlock)

    def it_reads_its_required_element_attribute(self):
        cxb = cast(
            CT_CustomXmlBlock, element("w:customXml{w:element=MyTag}")
        )
        assert cxb.element == "MyTag"

    def it_reads_its_optional_uri_attribute(self):
        cxb = cast(
            CT_CustomXmlBlock,
            element(
                "w:customXml{w:uri=http://example.org/ns,w:element=MyTag}"
            ),
        )
        assert cxb.uri == "http://example.org/ns"

    def it_returns_None_uri_when_absent(self):
        cxb = cast(
            CT_CustomXmlBlock, element("w:customXml{w:element=MyTag}")
        )
        assert cxb.uri is None

    def it_can_set_its_element_and_uri(self):
        cxb = cast(
            CT_CustomXmlBlock, element("w:customXml{w:element=MyTag}")
        )
        cxb.element = "Renamed"
        cxb.uri = "http://example.org/x"
        assert cxb.element == "Renamed"
        assert cxb.uri == "http://example.org/x"

    def it_reads_its_customXmlPr_child(self):
        cxb = cast(
            CT_CustomXmlBlock,
            element(
                "w:customXml{w:element=MyTag}/w:customXmlPr/"
                "w:attr{w:name=a,w:val=1}"
            ),
        )
        assert cxb.customXmlPr is not None
        assert isinstance(cxb.customXmlPr, CT_CustomXmlPr)
        assert len(cxb.customXmlPr.attr_lst) == 1

    def it_returns_None_customXmlPr_when_absent(self):
        cxb = cast(
            CT_CustomXmlBlock, element("w:customXml{w:element=MyTag}")
        )
        assert cxb.customXmlPr is None


# ---------------------------------------------------------------------------
# Flavor classes — constructible, share outer shape
# ---------------------------------------------------------------------------


class DescribeCT_CustomXmlRun:
    """Unit-test suite for `docx.oxml.custom_xml.CT_CustomXmlRun`.

    ``w:customXml`` is a single polymorphic QName; the parser registers it
    as :class:`CT_CustomXmlBlock`. :class:`CT_CustomXmlRun` is available
    for programmatic construction and is fully descriptor-equipped.
    """

    def it_exposes_element_attribute_descriptor(self):
        assert "element" in CT_CustomXmlRun.__dict__

    def it_exposes_uri_attribute_descriptor(self):
        assert "uri" in CT_CustomXmlRun.__dict__

    def it_exposes_customXmlPr_descriptor(self):
        assert "customXmlPr" in CT_CustomXmlRun.__dict__


class DescribeCT_CustomXmlRow:
    """Unit-test suite for `docx.oxml.custom_xml.CT_CustomXmlRow`."""

    def it_exposes_element_attribute_descriptor(self):
        assert "element" in CT_CustomXmlRow.__dict__

    def it_exposes_uri_attribute_descriptor(self):
        assert "uri" in CT_CustomXmlRow.__dict__

    def it_exposes_customXmlPr_descriptor(self):
        assert "customXmlPr" in CT_CustomXmlRow.__dict__


class DescribeCT_CustomXmlCell:
    """Unit-test suite for `docx.oxml.custom_xml.CT_CustomXmlCell`."""

    def it_exposes_element_attribute_descriptor(self):
        assert "element" in CT_CustomXmlCell.__dict__

    def it_exposes_uri_attribute_descriptor(self):
        assert "uri" in CT_CustomXmlCell.__dict__

    def it_exposes_customXmlPr_descriptor(self):
        assert "customXmlPr" in CT_CustomXmlCell.__dict__
