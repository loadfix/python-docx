"""Test suite for pptx.oxml.__init__.py module, primarily XML parser-related."""

import pytest
from lxml import etree

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement, oxml_parser, parse_xml, register_element_cls
from docx.oxml.shared import BaseOxmlElement


class DescribeOxmlElement:
    def it_returns_an_lxml_element_with_matching_tag_name(self):
        element = OxmlElement("a:foo")
        assert isinstance(element, etree._Element)
        assert element.tag == ("{http://schemas.openxmlformats.org/drawingml/2006/main}foo")

    def it_adds_supplied_attributes(self):
        element = OxmlElement("a:foo", {"a": "b", "c": "d"})
        assert etree.tostring(element) == (
            '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" a="b" c="d"/>'
        ).encode("utf-8")

    def it_adds_additional_namespace_declarations_when_supplied(self):
        ns1 = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ns2 = "other"
        element = OxmlElement("a:foo", nsdecls={"a": ns1, "x": ns2})
        assert len(element.nsmap.items()) == 2
        assert element.nsmap["a"] == ns1
        assert element.nsmap["x"] == ns2


class DescribeOxmlParser:
    def it_strips_whitespace_between_elements(self, whitespace_fixture):
        pretty_xml_text, stripped_xml_text = whitespace_fixture
        element = etree.fromstring(pretty_xml_text, oxml_parser)
        xml_text = etree.tostring(element, encoding="unicode")
        assert xml_text == stripped_xml_text

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def whitespace_fixture(self):
        pretty_xml_text = "<foø>\n  <bår>text</bår>\n</foø>\n"
        stripped_xml_text = "<foø><bår>text</bår></foø>"
        return pretty_xml_text, stripped_xml_text


class DescribeParseXml:
    def it_accepts_bytes_and_assumes_utf8_encoding(self, xml_bytes):
        parse_xml(xml_bytes)

    def it_accepts_unicode_providing_there_is_no_encoding_declaration(self):
        non_enc_decl = '<?xml version="1.0" standalone="yes"?>'
        enc_decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        xml_body = "<foo><bar>føøbår</bar></foo>"
        # unicode body by itself doesn't raise
        parse_xml(xml_body)
        # adding XML decl without encoding attr doesn't raise either
        xml_text = "%s\n%s" % (non_enc_decl, xml_body)
        parse_xml(xml_text)
        # but adding encoding in the declaration raises ValueError
        xml_text = "%s\n%s" % (enc_decl, xml_body)
        with pytest.raises(ValueError, match="Unicode strings with encoding declara"):
            parse_xml(xml_text)

    def it_uses_registered_element_classes(self, xml_bytes):
        register_element_cls("a:foo", CustElmCls)
        element = parse_xml(xml_bytes)
        assert isinstance(element, CustElmCls)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def xml_bytes(self):
        return (
            '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/200'
            '6/main">\n'
            "  <a:bar>foøbår</a:bar>\n"
            "</a:foo>\n"
        ).encode("utf-8")


class DescribeParseXmlRecovery:
    def it_raises_XMLSyntaxError_by_default_on_malformed_xml(self):
        with pytest.raises(etree.XMLSyntaxError):
            parse_xml(b"<w:p>unclosed")

    def it_returns_partial_tree_when_recover_is_True(self):
        element = parse_xml(b"<w:p>unclosed", recover=True)

        assert element is not None
        assert element.tag.endswith("}p") or element.tag == "p" or element.tag.endswith("p")

    def it_returns_None_when_input_is_unrecoverable(self):
        from docx.oxml.parser import recovery_mode

        with recovery_mode() as warnings:
            element = parse_xml(b"", recover=True)

        assert element is None
        assert len(warnings) >= 1

    def it_accumulates_warnings_under_recovery_mode(self):
        from docx.oxml.parser import recovery_mode

        with recovery_mode() as warnings:
            parse_xml(b"<root><child></root>")

        assert len(warnings) >= 1
        assert all(isinstance(w, str) for w in warnings)

    def it_deactivates_recovery_after_context_exits(self):
        from docx.oxml.parser import _recovery_state, recovery_mode

        with recovery_mode():
            assert _recovery_state.active is True
        assert _recovery_state.active is False
        # -- and subsequent calls without recover=True raise as usual --
        with pytest.raises(etree.XMLSyntaxError):
            parse_xml(b"<root>unclosed")


class DescribeParseXmlHugeTree:
    def it_rejects_attvalue_over_10mb_by_default(self):
        # -- libxml2's default AttValue cap is 10 MB. Building a 12 MB attvalue
        # -- triggers the "AttValue length too long" XMLSyntaxError — that's
        # -- the exact failure mode upstream#1086 describes.
        from lxml import etree as _etree

        big = "x" * (12 * 1024 * 1024)
        xml = '<root a="%s"/>' % big
        with pytest.raises(_etree.XMLSyntaxError):
            parse_xml(xml.encode("utf-8"))

    def it_parses_attvalue_over_10mb_when_huge_tree_active(self):
        from docx.oxml.parser import huge_tree_mode

        big = "x" * (12 * 1024 * 1024)
        xml = '<root a="%s"/>' % big
        with huge_tree_mode():
            element = parse_xml(xml.encode("utf-8"))
        assert element is not None
        assert element.get("a") == big

    def it_deactivates_huge_tree_after_context_exits(self):
        from docx.oxml.parser import _huge_tree_state, huge_tree_mode

        with huge_tree_mode():
            assert _huge_tree_state.active is True
        assert _huge_tree_state.active is False

    def it_composes_with_recovery_mode(self):
        from docx.oxml.parser import huge_tree_mode, recovery_mode

        with huge_tree_mode(), recovery_mode() as warnings:
            element = parse_xml(b"<w:p>unclosed")
        # -- malformed + huge-tree must still recover, not crash --
        assert element is not None
        assert isinstance(warnings, list)


class DescribeRegisterElementCls:
    def it_determines_class_used_for_elements_with_matching_tagname(self, xml_text):
        register_element_cls("a:foo", CustElmCls)
        foo = parse_xml(xml_text)
        assert type(foo) is CustElmCls
        assert type(foo.find(qn("a:bar"))) is etree._Element

    # fixture components ---------------------------------------------

    @pytest.fixture
    def xml_text(self):
        return (
            '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/200'
            '6/main">\n'
            "  <a:bar>foøbår</a:bar>\n"
            "</a:foo>\n"
        )


# ===========================================================================
# static fixture
# ===========================================================================


class CustElmCls(BaseOxmlElement):
    pass
