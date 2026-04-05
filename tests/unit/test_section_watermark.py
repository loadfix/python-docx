# pyright: reportPrivateUsage=false

"""Unit test suite for Section watermark methods."""

from __future__ import annotations

from unittest.mock import MagicMock, PropertyMock, patch

import pytest

from docx.oxml.watermark import has_watermark
from docx.section import Section
from docx.shared import Pt, RGBColor


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_VML_NS = "urn:schemas-microsoft-com:vml"


class DescribeSectionWatermark:
    def it_can_add_a_text_watermark(self, request: pytest.FixtureRequest):
        from lxml import etree

        from docx.oxml.section import CT_SectPr

        hdr_xml = (
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:v="urn:schemas-microsoft-com:vml">'
            "<w:p/>"
            "</w:hdr>"
        )
        hdr_element = etree.fromstring(hdr_xml.encode("utf-8"))

        header_mock = MagicMock()
        type(header_mock)._element = PropertyMock(return_value=hdr_element)

        sectPr = MagicMock(spec=CT_SectPr)
        document_part = MagicMock()
        section = Section(sectPr, document_part)

        with patch.object(type(section), "header", new_callable=PropertyMock) as mock_header:
            mock_header.return_value = header_mock
            section.add_text_watermark("DRAFT")

        assert has_watermark(hdr_element) is True

    def it_can_add_a_text_watermark_with_custom_params(self, request: pytest.FixtureRequest):
        from lxml import etree

        from docx.oxml.section import CT_SectPr

        hdr_xml = (
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:v="urn:schemas-microsoft-com:vml">'
            "<w:p/>"
            "</w:hdr>"
        )
        hdr_element = etree.fromstring(hdr_xml.encode("utf-8"))

        header_mock = MagicMock()
        type(header_mock)._element = PropertyMock(return_value=hdr_element)

        sectPr = MagicMock(spec=CT_SectPr)
        document_part = MagicMock()
        section = Section(sectPr, document_part)

        with patch.object(type(section), "header", new_callable=PropertyMock) as mock_header:
            mock_header.return_value = header_mock
            section.add_text_watermark(
                "SECRET",
                font="Arial",
                size=Pt(48),
                color=RGBColor(0xFF, 0x00, 0x00),
                layout="horizontal",
            )

        assert has_watermark(hdr_element) is True
        shapes = hdr_element.xpath(".//v:shape", namespaces={"v": _VML_NS})
        assert shapes[0].get("fillcolor") == "#FF0000"
        textpaths = hdr_element.xpath(".//v:textpath", namespaces={"v": _VML_NS})
        assert textpaths[0].get("string") == "SECRET"

    def it_can_remove_a_watermark(self, request: pytest.FixtureRequest):
        from lxml import etree

        from docx.oxml.section import CT_SectPr

        hdr_xml = (
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:v="urn:schemas-microsoft-com:vml">'
            "<w:p/>"
            '<w:p><w:r><w:pict>'
            '<v:shape id="PowerPlusWaterMarkObject">'
            '<v:textpath string="DRAFT"/>'
            "</v:shape>"
            "</w:pict></w:r></w:p>"
            "</w:hdr>"
        )
        hdr_element = etree.fromstring(hdr_xml.encode("utf-8"))
        assert has_watermark(hdr_element) is True

        header_mock = MagicMock()
        type(header_mock)._element = PropertyMock(return_value=hdr_element)
        type(header_mock)._has_definition = PropertyMock(return_value=True)

        sectPr = MagicMock(spec=CT_SectPr)
        document_part = MagicMock()
        section = Section(sectPr, document_part)

        with patch.object(type(section), "header", new_callable=PropertyMock) as mock_header:
            mock_header.return_value = header_mock
            section.remove_watermark()

        assert has_watermark(hdr_element) is False

    def it_reports_has_watermark_False_when_no_header(self, request: pytest.FixtureRequest):
        from docx.oxml.section import CT_SectPr

        header_mock = MagicMock()
        type(header_mock)._has_definition = PropertyMock(return_value=False)

        sectPr = MagicMock(spec=CT_SectPr)
        document_part = MagicMock()
        section = Section(sectPr, document_part)

        with patch.object(type(section), "header", new_callable=PropertyMock) as mock_header:
            mock_header.return_value = header_mock
            assert section.has_watermark is False

    def it_reports_has_watermark_True_when_present(self, request: pytest.FixtureRequest):
        from lxml import etree

        from docx.oxml.section import CT_SectPr

        hdr_xml = (
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:v="urn:schemas-microsoft-com:vml">'
            '<w:p><w:r><w:pict>'
            '<v:shape id="PowerPlusWaterMarkObject">'
            '<v:textpath string="DRAFT"/>'
            "</v:shape>"
            "</w:pict></w:r></w:p>"
            "</w:hdr>"
        )
        hdr_element = etree.fromstring(hdr_xml.encode("utf-8"))

        header_mock = MagicMock()
        type(header_mock)._element = PropertyMock(return_value=hdr_element)
        type(header_mock)._has_definition = PropertyMock(return_value=True)

        sectPr = MagicMock(spec=CT_SectPr)
        document_part = MagicMock()
        section = Section(sectPr, document_part)

        with patch.object(type(section), "header", new_callable=PropertyMock) as mock_header:
            mock_header.return_value = header_mock
            assert section.has_watermark is True

    def it_replaces_existing_watermark_when_adding(self, request: pytest.FixtureRequest):
        from lxml import etree

        from docx.oxml.section import CT_SectPr

        hdr_xml = (
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:v="urn:schemas-microsoft-com:vml">'
            "<w:p/>"
            '<w:p><w:r><w:pict>'
            '<v:shape id="PowerPlusWaterMarkObject">'
            '<v:textpath string="OLD"/>'
            "</v:shape>"
            "</w:pict></w:r></w:p>"
            "</w:hdr>"
        )
        hdr_element = etree.fromstring(hdr_xml.encode("utf-8"))

        header_mock = MagicMock()
        type(header_mock)._element = PropertyMock(return_value=hdr_element)

        sectPr = MagicMock(spec=CT_SectPr)
        document_part = MagicMock()
        section = Section(sectPr, document_part)

        with patch.object(type(section), "header", new_callable=PropertyMock) as mock_header:
            mock_header.return_value = header_mock
            section.add_text_watermark("NEW")

        # -- should have exactly one watermark, with the new text --
        shapes = hdr_element.xpath(
            ".//v:shape[@id='PowerPlusWaterMarkObject']",
            namespaces={"v": _VML_NS},
        )
        assert len(shapes) == 1
        textpaths = hdr_element.xpath(".//v:textpath", namespaces={"v": _VML_NS})
        assert textpaths[0].get("string") == "NEW"
